import json,re,datetime,math
from collections import defaultdict

KEY_VARS={"EarnedPrem","ReinsPrem","IncClaims","ReinsClaims","TabRes",
          "CededALRstat","CLRes","TS","Comm","PremTax","LivesIssued","AEGAdminPolCount"}

def parse_rbc(rows,surplus_rows=None):
    yrc={}
    for row in rows:
        dts=[c for c in row if isinstance(c,datetime.datetime) and 2020<=c.year<=2045]
        yis=[c for c in row if isinstance(c,(int,float)) and 2020<=c<=2045]
        if dts:
            for j,c in enumerate(row):
                if isinstance(c,datetime.datetime) and 2020<=c.year<=2045:yrc[j]=c.year
            break
        if len(yis)>=3:
            for j,c in enumerate(row):
                if isinstance(c,(int,float)) and 2020<=c<=2045:yrc[j]=int(c)
            break
    if not yrc:return {}
    m={}
    for row in rows:
        if not row or all(c is None for c in row):continue
        lbl=next((str(c).strip() for c in row if isinstance(c,str) and c.strip()),"")
        if not lbl:continue
        vals={yr:float(row[j]) for j,yr in yrc.items() if j<len(row) and isinstance(row[j],(int,float))}
        if vals:m[lbl]=vals
    # Propagate AVR (single value) across all years
    if 'AVR' in m:
        avr_vals=m['AVR']
        avr_base=list(avr_vals.values())[0] if avr_vals else 0
        all_yrs=set()
        for v in m.values():
            if isinstance(v,dict):all_yrs.update(v.keys())
        m['AVR']={yr:avr_base for yr in all_yrs}
    if surplus_rows:
        yrc2={}
        for row in surplus_rows:
            dts=[c for c in row if isinstance(c,datetime.datetime) and 2018<=c.year<=2045]
            yis=[c for c in row if isinstance(c,(int,float)) and 2018<=c<=2045]
            if dts:
                for j,c in enumerate(row):
                    if isinstance(c,datetime.datetime) and 2018<=c.year<=2045:yrc2[j]=c.year
                break
            if len(yis)>=3:
                for j,c in enumerate(row):
                    if isinstance(c,(int,float)) and 2018<=c<=2045:yrc2[j]=int(c)
                break
        if yrc2:
            for row in surplus_rows:
                if not row:continue
                lbl=next((str(c).strip() for c in row if isinstance(c,str) and c.strip()),"")
                if lbl in ("Original MS C2","Conservatism factor"):
                    vals={yr:float(row[j]) for j,yr in yrc2.items() if j<len(row) and isinstance(row[j],(int,float))}
                    if vals:m[lbl]=vals
    return m

def parse_bp(rows):
    yrc={}
    for row in rows:
        yv=[(j,c) for j,c in enumerate(row) if isinstance(c,(int,float)) and 2020<=c<=2060]
        if len(yv)>=3:yrc={j:int(c) for j,c in yv};break
    if not yrc:return {}
    bp={}
    for row in rows:
        if not row or all(c is None for c in row):continue
        lbl=next((str(c).strip() for c in row if isinstance(c,str) and c.strip()),"")
        if not lbl:continue
        vals={yr:float(row[j]) for j,yr in yrc.items() if j<len(row) and isinstance(row[j],(int,float))}
        if vals:bp[lbl]=vals
    return {k:{str(yr):v for yr,v in vd.items()} for k,vd in bp.items()}

def g(agg,vn,p):return agg.get(vn,{}).get(p,0.0)

def cy(by,p):
    # p=0 is baseline (Dec of base year); p=1..12 = year by+1; p=13..24 = year by+2 etc.
    return by+1+((p-1)//12) if p>0 else by

def eff_rate_for_p(rp,iy,by,p):
    """Get ceding rate for issue year iy at projection period p.
    For p=0 (baseline), use the same rate as p=1 (first projection year).
    The last populated calendar year's rate extends to all future years."""
    if p==0:return 0.0
    iy_key=2019 if iy<=2019 else iy
    iy_rates=rp.get(iy_key,{})
    if not iy_rates:return 0.0
    cal=cy(by,p)
    if iy>cal:return 0.0
    max_c=max(iy_rates.keys())
    return float(iy_rates.get(min(cal,max_c),0) or 0)

def apply_scalars(agg_d,agg_iy_d,assum):
    cs=float(assum.get("claim_scalar",1.0));ls=float(assum.get("lapse_scalar",1.0))
    ss=assum.get("sales_scalar",{})  # {iy: scalar} e.g. {2026:1.1, 2027:1.0}
    ss_active={int(k):float(v) for k,v in ss.items() if float(v)!=1.0} if ss else {}
    if cs==1.0 and ls==1.0 and not ss_active:return agg_d,agg_iy_d
    a={k:dict(v) for k,v in agg_d.items()}
    ai={iy:{k:dict(v) for k,v in vm.items()} for iy,vm in agg_iy_d.items()}
    # Apply sales scalar: scale ALL vars for specified IssYrs
    if ss_active:
        for iy_s,mult in ss_active.items():
            iy_key=str(iy_s)
            if iy_key not in ai:continue
            for vn in list(ai[iy_key].keys()):
                ai[iy_key][vn]={p:v*mult for p,v in ai[iy_key][vn].items()}
            # Also update the aggregated agg_d by subtracting old and adding new
            for vn in ai[iy_key]:
                orig_iy={iy_s_orig:{k:dict(v) for k,v in vm.items()} for iy_s_orig,vm in {iy_s:{k:dict(v) for k,v in agg_iy_d.get(iy_key,{}).items()}}.items()}
                for p,old_v in agg_iy_d.get(iy_key,{}).get(vn,{}).items():
                    a.setdefault(vn,{})[p]=(a.get(vn,{}).get(p,0) - old_v) + ai[iy_key][vn].get(p,0)
    if cs!=1.0:
        for vn in {"IncClaims","ReinsClaims"}:
            if vn in a:a[vn]={p:v*cs for p,v in a[vn].items()}
            for iy in ai:
                if vn in ai[iy]:ai[iy][vn]={p:v*cs for p,v in ai[iy][vn].items()}
    if ls!=1.0:
        ep=a.get("EarnedPrem",{});rp2=a.get("ReinsPrem",{})
        ps=sorted(set(list(ep.keys())+list(rp2.keys())))
        prem={p:ep.get(p,0)-rp2.get(p,0) for p in ps}
        surv={0:1.0}
        for p in sorted(p2 for p2 in ps if p2>0):
            prev=p-1;pp=prem.get(prev,0);cp=prem.get(p,0)
            if pp!=0 and cp!=0:
                ol=1.0-cp/pp;nl=min(max(ol*ls,0.0),0.9999)
                surv[p]=surv.get(prev,1.0)*((1-nl)/(1-ol)) if abs(1-ol)>1e-9 else surv.get(prev,1.0)
            else:surv[p]=surv.get(prev,1.0)
        for vn in list(a.keys()):a[vn]={p:v*surv.get(p,1.0) for p,v in a[vn].items()}
        for iy in ai:
            for vn in list(ai[iy].keys()):ai[iy][vn]={p:v*surv.get(p,1.0) for p,v in ai[iy][vn].items()}
    return a,ai

def apply_rp_agg(agg_iy,rp,by,periods):
    """Apply reinsurance % to agg_iy to produce ceded agg."""
    out={}
    for iy_s,vn_map in agg_iy.items():
        iy=int(iy_s)
        for vn,pv in vn_map.items():
            if vn not in out:out[vn]={}
            for p,v in pv.items():
                rate=eff_rate_for_p(rp,iy,by,p)
                if rate:out[vn][p]=out[vn].get(p,0.0)+v*rate
    return out

def bld_stmt(agg_d,assum,by,periods,allowance=False,zero_bop_ts=False):
    nier={int(k):float(v) for k,v in assum.get("nier",{}).items()}
    acqk="acq_exp_allowance" if allowance else "acq_exp"
    mntk="maint_exp_allowance" if allowance else "maint_exp"
    acqy={int(k):float(v) for k,v in assum.get(acqk,assum.get("acq_exp",{})).items()}
    mnty={int(k):float(v) for k,v in assum.get(mntk,assum.get("maint_exp",{})).items()}
    # Premium = EarnedPrem + ReinsPrem (ReinsPrem is already negative in EV or handle sign)
    prem={p:g(agg_d,"EarnedPrem",p)-g(agg_d,"ReinsPrem",p) for p in periods}
    # NII: monthly_rate = (1+annual)^(1/12)-1; reserve base = NET (TabRes+CededALRstat+CLRes+TS)
    nii={}
    for p in periods:
        annual_ni=nier.get(cy(by,p),nier.get(max(nier.keys()),0.04) if nier else 0.04)
        monthly_r=(1+annual_ni)**(1.0/12)-1
        prev=p-1
        total_assets=(g(agg_d,"TabRes",prev)+g(agg_d,"CededALRstat",prev)+
                      g(agg_d,"CLRes",prev)+g(agg_d,"TS",prev))
        nii[p]=total_assets*monthly_r
    revenue={p:prem[p]+nii[p] for p in periods}
    claims={p:-(g(agg_d,"IncClaims",p)-g(agg_d,"ReinsClaims",p)) for p in periods}
    # Reserve change: NET reserve = TabRes + CededALRstat
    def ns(p):return g(agg_d,"TabRes",p)+g(agg_d,"CededALRstat",p)
    delta_res={p:-(ns(p)-ns(p-1)) for p in periods}
    benefits={p:claims[p]+delta_res[p] for p in periods}
    comm={p:-abs(g(agg_d,"Comm",p)) for p in periods}
    pt={p:-abs(g(agg_d,"PremTax",p)) for p in periods}
    selling={p:comm[p]+pt[p] for p in periods}
    acq={p:-abs(g(agg_d,"LivesIssued",p)*acqy.get(cy(by,p),acqy.get(max(acqy.keys()),243.46) if acqy else 243.46)) for p in periods}
    mnt={p:-abs(g(agg_d,"AEGAdminPolCount",p)*mnty.get(cy(by,p),mnty.get(max(mnty.keys()),94.45) if mnty else 94.45)/12) for p in periods}
    op={p:acq[p]+mnt[p] for p in periods}
    pretax={p:revenue[p]+benefits[p]+selling[p]+op[p] for p in periods}
    ts_vals={p:g(agg_d,"TS",p) for p in periods}
    _bop_ts=0.0 if zero_bop_ts else ts_vals.get(min(periods),0)
    def _ts_prev(p):
        if p==min(periods):return _bop_ts
        return ts_vals.get(p-1,_bop_ts)
    de={p:pretax[p]*0.79-(ts_vals[p]-_ts_prev(p)) for p in periods}
    # Balance sheet: NET reserves (year-end = last period of each year)
    pol_res_v={p:g(agg_d,"TabRes",p)+g(agg_d,"CededALRstat",p) for p in periods}
    cl_res_v={p:g(agg_d,"CLRes",p) for p in periods}
    total_assets_v={p:pol_res_v[p]+cl_res_v[p]+ts_vals[p] for p in periods}
    return {"premium":prem,"nii":nii,"revenue":revenue,"claims":claims,
            "delta_reserves":delta_res,"benefits":benefits,"commissions":comm,
            "prem_tax":pt,"selling_expense":selling,"acq_expense":acq,
            "maint_expense":mnt,"op_expense":op,"pretax_income":pretax,
            "distributable_earnings":de,"target_surplus":ts_vals,
            "admin_pol_count":{p:g(agg_d,"AEGAdminPolCount",p) for p in periods},
            "lives_issued":{p:g(agg_d,"LivesIssued",p) for p in periods},
            "policy_reserve":pol_res_v,"claim_reserve":cl_res_v,
            "total_assets":total_assets_v}

def ann_eoy(stmt,by,periods):
    """Annual aggregation: for flow items sum months; for stock items use Dec 31 (last period of year)."""
    FLOW_KEYS={"premium","nii","revenue","claims","delta_reserves","benefits","commissions",
               "prem_tax","selling_expense","acq_expense","maint_expense","op_expense",
               "pretax_income","distributable_earnings","comm1","comm2","admin_pol_count","lives_issued"}
    STOCK_KEYS={"target_surplus","policy_reserve","claim_reserve","total_assets"}
    out={}
    for k,v in stmt.items():
        if not(isinstance(v,dict) and v and all(isinstance(kk,int) for kk in v)):
            out[k]=v;continue
        if k in FLOW_KEYS:
            a=defaultdict(float)
            for p,vv in v.items():
                if p>0:a[cy(by,p)]+=vv
            out[k]=dict(a)
        elif k in STOCK_KEYS:
            # Year-end value: last period of each year (p=12,24,36...)
            a={}
            proj_yrs=sorted(set(cy(by,p) for p in periods if p>0))
            for yr in proj_yrs:
                # Find last period mapping to this year
                last_p=max((p for p in periods if p>0 and cy(by,p)==yr),default=None)
                if last_p is not None:a[yr]=v.get(last_p,0)
            out[k]=a
        else:
            a=defaultdict(float)
            for p,vv in v.items():
                if p>0:a[cy(by,p)]+=vv
            out[k]=dict(a)
    out["periods"]=sorted(set(cy(by,p) for p in periods if p>0))
    return out

def mtr(stmt,disc,periods,by=2025):
    de=stmt.get("distributable_earnings",stmt.get("pretax_income",{}))
    from collections import defaultdict as _dd
    _ann=_dd(float)
    for p,v in de.items():
        if p>0:_ann[by+1+((p-1)//12)]+=v
    _yrs=sorted(_ann.keys())
    pvde=sum(_ann[yr]/(1+disc)**i for i,yr in enumerate(_yrs,1))
    try:
        def npv(r):return sum(v/((1+r)**i) for i,v in enumerate(_dl,1))
        if npv(-0.5)*npv(5.0)<0:
            lo,hi=-0.5,5.0
            for _ in range(80):
                mid=(lo+hi)/2
                if npv(mid)>0:lo=mid
                else:hi=mid
            irr=(1+(lo+hi)/2)**12-1
    except:pass
    _dl=[_ann[yr] for yr in _yrs];irr=None
    cum=mn=0
    for v in _dl:
        cum+=v
        if cum<mn:mn=cum
    return {"pvde":pvde,"irr":irr,"max_neg_cum_de":mn}

def ser(d):
    if isinstance(d,dict):return {str(k):ser(v) for k,v in d.items()}
    if isinstance(d,list):return [ser(i) for i in d]
    return d

def compute_rbc_net(rbc_orig,ap,ac):
    import math as _m
    if not rbc_orig:return {}
    net={k:dict(v) if isinstance(v,dict) else v for k,v in rbc_orig.items()}
    _all_yrs=sorted(rbc_orig.get("TSC2",{}).keys())
    years=[y for y in _all_yrs if y>min(_all_yrs)]
    # Get conservatism factor schedule from original (keyed by year)
    # Updated section uses same factors but indexed from 0 (not year-1)
    orig_cons=rbc_orig.get("Conservatism factor",{})
    sorted_yrs=sorted(orig_cons.keys())
    sorted_cons_vals=[orig_cons[y] for y in sorted_yrs]
    avr_list=list(rbc_orig.get("AVR",{}).values())
    avr=avr_list[0] if avr_list else 0
    # Portion in non-ins cos (used for TAC)
    non_ins=rbc_orig.get("Portion in non-ins cos",{})
    for i,yr in enumerate(years):
        ys=str(yr)
        d_clm=abs((ap.get("claims",{}).get(ys) or ap.get("claims",{}).get(yr,0)) or 0)
        c_clm=abs((ac.get("claims",{}).get(ys) or ac.get("claims",{}).get(yr,0)) or 0)
        cede_r=c_clm/d_clm if d_clm>0 else 0
        # Original MS C2 from the workbook's own 'Original MS C2' row
        orig_ms=rbc_orig.get("Original MS C2",{}).get(yr)
        if orig_ms is None:orig_ms=rbc_orig.get("TSC2",{}).get(yr,0)
        orig_tsc2=rbc_orig.get("TSC2",{}).get(yr,0)
        upd_ms=orig_ms*(1-cede_r)
        delta_ms=orig_ms-upd_ms
        upd_tsc2=orig_tsc2-delta_ms
        net.setdefault("Original MS C2",{})[yr]=orig_ms
        net.setdefault("Updated MS C2",{})[yr]=upd_ms
        net["TSC2"][yr]=upd_tsc2
        # Pre-covariance total = sum of all components
        pre=sum((net.get(r,{}).get(yr,0) or 0) for r in ["TSC0","TSC1","TSLR016","TSC1CS","TSC2","TSC3","TSC4a","TSC4b"])
        net.setdefault("Total (Pre Covariance)",{})[yr]=pre
        # Post-covariance: NAIC formula
        # PostCov = TSC0 + TSC4a + sqrt((TSC1+TSLR+TSC3)^2 + TSC1CS^2 + UpdTSC2^2 + TSC4b^2)
        t0=(net.get("TSC0",{}).get(yr,0) or 0)
        t1=(net.get("TSC1",{}).get(yr,0) or 0)
        tslr=(net.get("TSLR016",{}).get(yr,0) or 0)
        t1cs=(net.get("TSC1CS",{}).get(yr,0) or 0)
        t2=upd_tsc2
        t3=(net.get("TSC3",{}).get(yr,0) or 0)
        t4a=(net.get("TSC4a",{}).get(yr,0) or 0)
        t4b=(net.get("TSC4b",{}).get(yr,0) or 0)
        post=t0+t4a+_m.sqrt((t1+tslr+t3)**2+t1cs**2+t2**2+t4b**2)
        net.setdefault("Post Covariance",{})[yr]=post
        # Conservatism factor: indexed from 0 (year i uses sorted_cons_vals[i])
        # This produces the same values as the 'Updated' section in the RBC workbook
        cons=float(sorted_cons_vals[i]) if i<len(sorted_cons_vals) else float(sorted_cons_vals[-1])
        add1=post*cons;add3=post*0.03
        net.setdefault("Adding 1% conservatism",{})[yr]=add1
        net.setdefault("Adding 3% loss of covariance",{})[yr]=add3
        total_post=post+add1+add3
        net.setdefault("Total (Post Covariance)",{})[yr]=total_post
        # Update covariance ratio display
        net.setdefault("Covariance",{})[yr]=post/pre if pre else 0
        # Surplus update
        c_pretax=(ac.get("pretax_income",{}).get(ys) or ac.get("pretax_income",{}).get(yr,0)) or 0
        after_tax=-c_pretax*0.79
        net.setdefault("diff in ms income after tax",{})[yr]=after_tax/1e6
        orig_surp=rbc_orig.get("Total Surplus",{}).get(yr,0) or 0
        cum_ms=sum(net.get("Change in MS Contribution",{}).get(y,0) for y in years if y<=yr)
        net_surp=orig_surp+cum_ms+after_tax/1e6
        net.setdefault("Change in MS Contribution",{})[yr]=after_tax/1e6
        net.setdefault("Total Surplus",{})[yr]=net_surp
        # TAC = Total Surplus - Portion in non-ins cos + AVR
        non_ins_yr=float(non_ins.get(yr,0) or 0)
        tac=net_surp-non_ins_yr+avr
        net.setdefault("TAC (C&S + AVR)",{})[yr]=tac
        net.setdefault("RBC Ratio pre-covariance",{})[yr]=tac/pre if pre else 0
        net.setdefault("RBC Ratio post-covariance w margin",{})[yr]=tac/total_post if total_post else 0
        net.setdefault("RBC Ratio post-covariance w/o margin",{})[yr]=tac/(post+add3) if (post+add3) else 0
    return net

DIAG_LINES=[("Total Revenue","revenue"),("Benefits","benefits"),
    ("Selling Expense","selling_expense"),("Op Expense","op_expense"),
    ("Pretax Income","pretax_income")]

def compute_ann_de(ann_pti,sp_ts_monthly,sc_ts_monthly,by,periods,mode="predeal"):
    """Compute annual DE = annual_PTI*0.79 - (EOP_TS - BOP_TS)
    mode: predeal, ceded, net
    For ceded: BOP_TS = 0 in first year, sc_ts thereafter
    For net: BOP_TS = predeal BOP, EOP_TS = predeal_EOP - ceded_EOP
    sp_ts_monthly = predeal monthly TS dict {period: value}
    sc_ts_monthly = ceded monthly TS dict {period: value} (None for predeal)
    """
    def last_p_of(yr): return (yr-by)*12
    proj_yrs=sorted(set(by+1+((p-1)//12) for p in periods if p>0))
    result={}
    for i,yr in enumerate(proj_yrs):
        pti=ann_pti.get(str(yr)) or ann_pti.get(yr,0)
        eop_p=last_p_of(yr)
        bop_p=last_p_of(yr-1) if i>0 else 0
        sp_eop=sp_ts_monthly.get(eop_p,0)
        sp_bop=sp_ts_monthly.get(bop_p,0)
        if mode=="predeal":
            de=pti*0.79-(sp_eop-sp_bop)
        elif mode=="ceded":
            sc_ts=sc_ts_monthly or {}
            sc_eop=sc_ts.get(eop_p,0)
            sc_bop=0.0 if i==0 else sc_ts.get(bop_p,0)
            de=pti*0.79-(sc_eop-sc_bop)
        else:  # net
            sc_ts=sc_ts_monthly or {}
            sc_eop=sc_ts.get(eop_p,0)
            net_eop=sp_eop-sc_eop
            net_bop=sp_bop  # ceded BOP=0, so net BOP = predeal BOP
            de=pti*0.79-(net_eop-net_bop)
        result[str(yr)]=de
    return result


def parse_rbc_full(rows):
    """
    Parse the ORIGINAL section of an uploaded RBC file into {years, data}.
    MS section rows stored as 'TSC0','TSC1', etc.
    All-product TSC rows stored as 'All_TSC0','All_TSC1', etc.
    Stops at 'Predeal' or 'Net' marker to avoid overwriting Original values.
    Resets in_allprod=False after 'Total (Pre Covariance)' row so subsequent
    Post Covariance/Surplus/Ratio rows don't get the All_ prefix.
    """
    if not rows: return {'years':[],'data':{}}
    # Detect year row from first 5 rows
    years=[]
    for r in rows[:5]:
        cands=[]
        for c in (r[1:] if r else []):
            if c is None: continue
            try:
                yr=c.year if hasattr(c,'year') else int(float(str(c)))
                if 2000<=yr<=2100: cands.append(yr)
            except: pass
        if len(cands)>=3: years=cands; break
    if not years: years=list(range(2025,2036))

    data={}; in_allprod=False; started_original=False
    for r in rows:
        if not r or r[0] is None: continue
        lbl=str(r[0]).strip()
        if not lbl: continue
        lbl_low=lbl.lower()
        # STOP: end of Original section
        if started_original and lbl_low in ('predeal','net','post-deal','postdeal'):
            break
        # Section markers - reset all-product flag
        if lbl_low in ('original',):
            started_original=True; in_allprod=False; continue
        if lbl_low in ('all product','all products'):
            in_allprod=True; continue
        if lbl_low in ('med sup','med supplement','rbc charges','combined'):
            in_allprod=False; continue
        # Build values
        vals={}
        for i,yr in enumerate(years):
            v=r[i+1] if i+1<len(r) else None
            if v is not None:
                try: vals[yr]=float(v)
                except: pass
        if vals:
            key=('All_'+lbl) if in_allprod else lbl
            data[key]=vals
            # The Total (Pre Covariance) row is the LAST line of the All-product
            # block. After this, Post Cov / Surplus / Ratios are NOT all-product.
            if in_allprod and lbl_low.startswith('total (pre cov'):
                in_allprod=False
    return {'years':years,'data':data}




def compute_original_rbc(orig_full):
    """
    Compute the Original RBC values formulaically from raw uploaded MS + All-product TSC data.
    All values in $M (same units as input file).
    Returns same structure as compute_predeal_rbc for consistent rendering.
    """
    data=orig_full.get('data',{})
    years=orig_full.get('years',[])
    if not years or not data:
        return {'years':[],'original_values':{}}

    def gO(key,yr):
        d=data.get(key,{})
        v=d.get(yr, d.get(str(yr), d.get(int(yr) if isinstance(yr,str) else yr, None)))
        return float(v) if v is not None else 0.0

    base_yr=min(years)
    avr_const=gO('AVR',base_yr)  # AVR held constant
    result={}
    prem_rel=['TSC0','TSC1','TSLR016','TSC1CS','TSC3','TSC4a','TSC4b']

    for yr in sorted(years):
        # Pull MS Income and TSC values directly from data
        ms_prem=gO('Premium',yr)
        ms_claims_raw=gO('Claims',yr)  # already negative
        ms_income=gO('Income',yr)
        ms_tsc={}
        for tsc in prem_rel: ms_tsc[tsc]=gO(tsc,yr)
        ms_tsc['TSC2']=gO('TSC2',yr)

        # All-product TSC values from data
        allprod_tsc={}
        for tsc in prem_rel+['TSC2']:
            allprod_tsc[tsc]=gO('All_'+tsc,yr)

        # Pre-covariance = sum of all-product TSC values (FORMULAIC)
        pre_cov=sum(allprod_tsc.values())

        # Post-covariance via NAIC Health RBC covariance formula:
        # PostCov = TSC0 + TSC4a + sqrt((TSC1+TSLR016+TSC3)^2 + TSC1CS^2 + TSC2^2 + TSC4b^2)
        t0=allprod_tsc.get('TSC0',0.0); t1=allprod_tsc.get('TSC1',0.0)
        tlr=allprod_tsc.get('TSLR016',0.0); t1cs=allprod_tsc.get('TSC1CS',0.0)
        t2=allprod_tsc.get('TSC2',0.0); t3=allprod_tsc.get('TSC3',0.0)
        t4a=allprod_tsc.get('TSC4a',0.0); t4b=allprod_tsc.get('TSC4b',0.0)
        post_cov=t0+t4a+math.sqrt((t1+tlr+t3)**2+t1cs**2+t2**2+t4b**2)

        # Conservatism additions (FORMULAIC)
        consv=gO('Conservatism factor',yr)
        # Base year: no margin additions per workbook convention
        is_base=(yr==min(years))
        add_1pct=post_cov*consv if not is_base else 0.0
        add_3pct=post_cov*0.03 if not is_base else 0.0
        total_post_cov=post_cov+add_1pct+add_3pct

        # Surplus and TAC (FORMULAIC)
        total_surplus=gO('Total Surplus',yr)
        in_lobs=gO('In LOBs',yr)
        portion_non_ins=gO('Portion in non-ins cos',yr)
        tac=total_surplus-portion_non_ins+avr_const

        # RBC Ratios (FORMULAIC)
        ratio_pre_cov=tac/pre_cov if pre_cov else 0.0
        ratio_w_margin=tac/total_post_cov if total_post_cov else 0.0
        ratio_wo_margin=tac/(post_cov+add_3pct) if (post_cov+add_3pct) else 0.0

        result[yr]={
            'ms_prem':ms_prem,'ms_claims':abs(ms_claims_raw),'ms_income':ms_income,
            'ms_tsc':ms_tsc,'allprod_tsc':allprod_tsc,
            'pre_cov':pre_cov,'post_cov':post_cov,
            'add_1pct':add_1pct,'add_3pct':add_3pct,'total_post_cov':total_post_cov,
            'total_surplus':total_surplus,'in_lobs':in_lobs,
            'portion_non_ins':portion_non_ins,'avr':avr_const,'tac':tac,
            'ratio_pre_cov':ratio_pre_cov,
            'ratio_w_margin':ratio_w_margin,'ratio_wo_margin':ratio_wo_margin,
        }
    return {'years':years,'original_values':result}


def compute_predeal_rbc(orig_rbc, predeal_annual, orig_annual):
    """
    Predeal RBC following SurplusRecalc.xlsx rows 47-89.
    All monetary values are in $M (same units as uploaded RBC file).
    predeal_annual values come from engine in $ → divide by 1e6.
    """
    data=orig_rbc.get('data',{}); years=orig_rbc.get('years',[])
    if not years or not data: return {'years':[],'predeal_adjustments':{}}

    def gO(key,yr):
        d=data.get(key,{})
        v=d.get(yr, d.get(str(yr), d.get(int(yr) if isinstance(yr,str) else yr, None)))
        return float(v) if v is not None else 0.0
    def gA(annual,key,yr):
        d=annual.get(key,{})
        v=d.get(str(yr), d.get(yr, None))
        return float(v)/1e6 if v is not None else 0.0

    base_yr=min(years)
    # AVR: use first year value for all years
    avr_const=gO('AVR', base_yr)
    # Conservatism factor per year from original
    consv_by_yr={yr:gO('Conservatism factor',yr) for yr in years}

    cum_surplus_adj=0.0
    result={}

    for yr in sorted(years):
        pred_prem=gA(predeal_annual,'premium',yr)
        pred_claims=abs(gA(predeal_annual,'claims',yr))
        pred_income=gA(predeal_annual,'pretax_income',yr)

        orig_ms_prem=gO('Premium',yr); orig_ms_claims=abs(gO('Claims',yr))
        orig_ms_income=gO('Income',yr)

        if yr==base_yr:
            prem_ratio=1.0; claim_ratio=1.0
        else:
            prem_ratio=pred_prem/orig_ms_prem if orig_ms_prem else 1.0
            claim_ratio=pred_claims/orig_ms_claims if orig_ms_claims else 1.0

        # MS TSC (rows 52-59): updated by premium or claims ratio
        prem_rel=['TSC0','TSC1','TSLR016','TSC1CS','TSC3','TSC4a','TSC4b']
        ms_tsc={}
        for tsc in prem_rel: ms_tsc[tsc]=gO(tsc,yr)*prem_ratio
        ms_tsc['TSC2']=gO('TSC2',yr)*claim_ratio

        # All-product TSC (rows 61-68): Orig_All + (PredealMS - OrigMS)
        allprod_tsc={}
        for tsc in prem_rel+['TSC2']:
            orig_ms=gO(tsc,yr); orig_all=gO('All_'+tsc,yr)
            allprod_tsc[tsc]=orig_all+(ms_tsc[tsc]-orig_ms)

        # Pre-covariance total (sum of updated all-product TSCs)
        pre_cov=sum(allprod_tsc.values())
        # Post-covariance via NAIC Health RBC covariance formula on UPDATED TSCs:
        # PostCov = TSC0 + TSC4a + sqrt((TSC1+TSLR016+TSC3)^2 + TSC1CS^2 + TSC2^2 + TSC4b^2)
        t0=allprod_tsc.get('TSC0',0.0); t1=allprod_tsc.get('TSC1',0.0)
        tlr=allprod_tsc.get('TSLR016',0.0); t1cs=allprod_tsc.get('TSC1CS',0.0)
        t2=allprod_tsc.get('TSC2',0.0); t3=allprod_tsc.get('TSC3',0.0)
        t4a=allprod_tsc.get('TSC4a',0.0); t4b=allprod_tsc.get('TSC4b',0.0)
        post_cov=t0+t4a+math.sqrt((t1+tlr+t3)**2+t1cs**2+t2**2+t4b**2)

        # Conservatism additions (base year: none, per workbook)
        consv=consv_by_yr.get(yr,0.0)
        is_base=(yr==base_yr)
        add_1pct=post_cov*consv if not is_base else 0.0   # R73
        add_3pct=post_cov*0.03 if not is_base else 0.0    # R74
        total_post_cov=post_cov+add_1pct+add_3pct          # R75

        # Income diff (R78): (pred_income - orig_ms_income) * 0.79
        inc_diff=(pred_income-orig_ms_income)*0.79
        if yr>base_yr:
            cum_surplus_adj+=inc_diff

        # Total Surplus (R79) = orig_surplus + cumulative income diffs
        adj_surplus=gO('Total Surplus',yr)+cum_surplus_adj

        # TAC = surplus - portion_non_ins + AVR
        portion_non_ins=gO('Portion in non-ins cos',yr)
        tac=adj_surplus-portion_non_ins+avr_const

        # RBC Ratios
        ratio_pre_cov=tac/pre_cov if pre_cov else 0.0          # TAC/PreCov
        ratio_w_margin=tac/total_post_cov if total_post_cov else 0.0  # TAC/(post+1%+3%)
        ratio_wo_margin=tac/(post_cov+add_3pct) if (post_cov+add_3pct) else 0.0  # TAC/(post+3%)

        in_lobs_val=gO('In LOBs',yr)
        result[yr]={
            'prem_ratio':prem_ratio,'claim_ratio':claim_ratio,
            'ms_prem':pred_prem,'ms_claims':pred_claims,'ms_income':pred_income,
            'ms_tsc':ms_tsc,'allprod_tsc':allprod_tsc,
            'pre_cov':pre_cov,'post_cov':post_cov,
            'add_1pct':add_1pct,'add_3pct':add_3pct,'total_post_cov':total_post_cov,
            'inc_diff':inc_diff,'cum_surplus_adj':cum_surplus_adj,
            'total_surplus':adj_surplus,'in_lobs':in_lobs_val,'portion_non_ins':portion_non_ins,'avr':avr_const,'tac':tac,
            'ratio_pre_cov':ratio_pre_cov,
            'ratio_w_margin':ratio_w_margin,'ratio_wo_margin':ratio_wo_margin,
        }
    return {'years':years,'predeal_adjustments':result}


def compute_net_rbc(predeal_rbc_result, net_annual, predeal_annual, orig_rbc):
    """
    Net RBC following SurplusRecalc.xlsx rows 91-134.
    Net MS TSC = Predeal_MS_TSC × (net/predeal ratio).
    TSC4a, TSC4b: same as predeal.
    All-product = Predeal_All + (Net_MS - Predeal_MS).
    Surplus = Predeal_Surplus + cumsum((net_income-pred_income)*0.79).
    """
    years=predeal_rbc_result.get('years',[]); pred_adj=predeal_rbc_result.get('predeal_adjustments',{})
    data=orig_rbc.get('data',{})
    if not years: return {'years':[],'net_adjustments':{}}

    def gO(key,yr):
        d=data.get(key,{})
        v=d.get(yr, d.get(str(yr), d.get(int(yr) if isinstance(yr,str) else yr, None)))
        return float(v) if v is not None else 0.0
    def gA(annual,key,yr):
        d=annual.get(key,{})
        v=d.get(str(yr), d.get(yr, None))
        return float(v)/1e6 if v is not None else 0.0

    base_yr=min(years)
    avr_const=gO('AVR',base_yr)
    consv_by_yr={yr:gO('Conservatism factor',yr) for yr in years}
    prem_rel=['TSC0','TSC1','TSLR016','TSC1CS','TSC3']  # TSC4a,TSC4b stay as predeal

    cum_surplus_adj=0.0
    result={}

    for yr in sorted(years):
        pi=pred_adj.get(yr, pred_adj.get(str(yr), {}))
        pred_ms_tsc=pi.get('ms_tsc',{}); pred_allprod_tsc=pi.get('allprod_tsc',{})

        pred_prem=gA(predeal_annual,'premium',yr); net_prem=gA(net_annual,'premium',yr)
        pred_claims=abs(gA(predeal_annual,'claims',yr)); net_claims=abs(gA(net_annual,'claims',yr))
        pred_income=gA(predeal_annual,'pretax_income',yr); net_income=gA(net_annual,'pretax_income',yr)

        if yr==base_yr:
            prem_ratio=1.0; claim_ratio=1.0
        else:
            prem_ratio=net_prem/pred_prem if pred_prem else 1.0
            claim_ratio=net_claims/pred_claims if pred_claims else 1.0

        # Net MS TSC = Predeal_MS × (net/predeal ratio)
        ms_tsc={}
        for tsc in prem_rel: ms_tsc[tsc]=pred_ms_tsc.get(tsc,0.0)*prem_ratio
        ms_tsc['TSC4a']=pred_ms_tsc.get('TSC4a',0.0)   # same as predeal
        ms_tsc['TSC4b']=pred_ms_tsc.get('TSC4b',0.0)   # same as predeal
        ms_tsc['TSC2']=pred_ms_tsc.get('TSC2',0.0)*claim_ratio

        # All-product TSC = Predeal_All + (Net_MS - Predeal_MS)
        allprod_tsc={}
        for tsc in prem_rel+['TSC4a','TSC4b','TSC2']:
            pred_ms=pred_ms_tsc.get(tsc,0.0); pred_all=pred_allprod_tsc.get(tsc,0.0)
            allprod_tsc[tsc]=pred_all+(ms_tsc[tsc]-pred_ms)

        # Pre-covariance and post-covariance via NAIC formula on UPDATED net TSCs
        pre_cov=sum(allprod_tsc.values())
        t0=allprod_tsc.get('TSC0',0.0); t1=allprod_tsc.get('TSC1',0.0)
        tlr=allprod_tsc.get('TSLR016',0.0); t1cs=allprod_tsc.get('TSC1CS',0.0)
        t2=allprod_tsc.get('TSC2',0.0); t3=allprod_tsc.get('TSC3',0.0)
        t4a=allprod_tsc.get('TSC4a',0.0); t4b=allprod_tsc.get('TSC4b',0.0)
        post_cov=t0+t4a+math.sqrt((t1+tlr+t3)**2+t1cs**2+t2**2+t4b**2)

        consv=consv_by_yr.get(yr,0.0)
        is_base=(yr==base_yr)
        add_1pct=post_cov*consv if not is_base else 0.0
        add_3pct=post_cov*0.03 if not is_base else 0.0
        total_post_cov=post_cov+add_1pct+add_3pct

        inc_diff=(net_income-pred_income)*0.79
        if yr>base_yr: cum_surplus_adj+=inc_diff

        pred_surplus=pi.get('total_surplus',0.0)
        adj_surplus=pred_surplus+cum_surplus_adj

        portion_non_ins=pi.get('portion_non_ins', gO('Portion in non-ins cos',yr))
        tac=adj_surplus-portion_non_ins+avr_const

        ratio_pre_cov=tac/pre_cov if pre_cov else 0.0
        ratio_w_margin=tac/total_post_cov if total_post_cov else 0.0
        ratio_wo_margin=tac/(post_cov+add_3pct) if (post_cov+add_3pct) else 0.0

        in_lobs_val=gO('In LOBs',yr)
        result[yr]={
            'prem_ratio':prem_ratio,'claim_ratio':claim_ratio,
            'ms_prem':net_prem,'ms_claims':net_claims,'ms_income':net_income,
            'ms_tsc':ms_tsc,'allprod_tsc':allprod_tsc,
            'pre_cov':pre_cov,'post_cov':post_cov,
            'add_1pct':add_1pct,'add_3pct':add_3pct,'total_post_cov':total_post_cov,
            'inc_diff':inc_diff,'cum_surplus_adj':cum_surplus_adj,
            'total_surplus':adj_surplus,'in_lobs':in_lobs_val,'portion_non_ins':portion_non_ins,'avr':avr_const,'tac':tac,
            'ratio_pre_cov':ratio_pre_cov,
            'ratio_w_margin':ratio_w_margin,'ratio_wo_margin':ratio_wo_margin,
        }
    return {'years':years,'net_adjustments':result}



def lookup_cco(lr, cco_table):
    """Given a loss ratio and a cco_table list of [lr_min, lr_max, comm],
    return the commission per policy per year. Table is matched as:
    lr >= lr_min AND lr < lr_max.  Last row uses lr_max=infinity."""
    for row in cco_table:
        lo, hi, comm = float(row[0]), float(row[1]), float(row[2])
        if lr >= lo and (hi == float('inf') or lr < hi):
            return comm
    # Fallback: last row
    if cco_table:
        return float(cco_table[-1][2])
    return 200.0

def compute_lr_by_iy(agg_iy, rp, by, periods):
    """Compute cumulative claims/premium loss ratio for each issue year
    through December of the first calendar year the issue year is ceded.
    Returns {iy: lr} dict."""
    lr_by_iy = {}
    for iy_s, vn_map in agg_iy.items():
        iy = int(iy_s)
        iy_key = 2019 if iy <= 2019 else iy
        iy_rates = rp.get(iy_key, {})
        if not iy_rates:
            continue
        # First calendar year with non-zero cede rate
        first_cede_cy = None
        for cal, rate in sorted(iy_rates.items()):
            if float(rate or 0) > 0:
                first_cede_cy = cal
                break
        if first_cede_cy is None:
            continue
        # Sum claims and premium through December of first_cede_cy
        # period 0 = Dec base year (by), period 1 = Jan by+1, etc.
        # December of first_cede_cy = period 12*(first_cede_cy - by)
        max_p = 12 * (first_cede_cy - by)
        total_prem = 0.0; total_claims = 0.0
        ep_data = vn_map.get('EarnedPrem', {})
        cl_data = vn_map.get('IncClaims', {})
        for p in periods:
            if p > 0 and p <= max_p:
                total_prem += float(ep_data.get(p, 0))
                total_claims += float(cl_data.get(p, 0))
        if total_prem > 0:
            lr_by_iy[iy] = total_claims / total_prem
    return lr_by_iy


def batch_run(ev_agg, base_assum, by, rbc_rows=None, bp_rows=None, surplus_rows=None,
              n_runs=100, claim_std=0.02, lapse_std=0.10, seed=42):
    """Run N stochastic scenarios sampling claim_scalar and lapse_scalar
    from normal distributions. Returns list of result dicts with key metrics."""
    import random
    random.seed(seed)
    results = []
    disc_r = float(base_assum.get('discount_rate', 0.08))
    def gauss(mu, sigma): return max(0.1, random.gauss(mu, sigma))
    for i in range(n_runs):
        a = dict(base_assum)
        a['claim_scalar'] = gauss(1.0, claim_std)
        a['lapse_scalar'] = gauss(1.0, lapse_std)
        try:
            r = run_model(ev_agg, a, by, rbc_rows, bp_rows, surplus_rows)
            mp = r.get('metrics_predeal', {}); mc = r.get('metrics_ceded', {}); mn = r.get('metrics_net', {})
            oc = r.get('rbc_orig_computed', {}); nr = r.get('rbc_net_result', {})
            orig_vals = oc.get('original_values', {})
            net_adjs = nr.get('net_adjustments', {})
            results.append({
                'run': i+1,
                'claim_scalar': a['claim_scalar'],
                'lapse_scalar': a['lapse_scalar'],
                'predeal_pvde': mp.get('pvde', 0),
                'predeal_irr':  mp.get('irr', 0),
                'ceded_pvde':   mc.get('pvde', 0),
                'net_pvde':     mn.get('pvde', 0),
                'net_irr':      mn.get('irr', 0),
                'net_rbc_2026': (net_adjs.get(2026) or {}).get('ratio_w_margin', 0),
                'net_rbc_2031': (net_adjs.get(2031) or {}).get('ratio_w_margin', 0),
                'orig_rbc_2026':(orig_vals.get(2026) or {}).get('ratio_w_margin', 0),
                'lift_2026': ((net_adjs.get(2026) or {}).get('ratio_w_margin', 0) -
                              (orig_vals.get(2026) or {}).get('ratio_w_margin', 0)),
                'lift_2031': ((net_adjs.get(2031) or {}).get('ratio_w_margin', 0) -
                              (orig_vals.get(2031) or {}).get('ratio_w_margin', 0)),
            })
        except Exception:
            pass
    return results



def _frontier_hz_end(ev_agg, by):
    # End of projection horizon (life-of-cohort cession window).
    _per = sorted(int(p) for p in ev_agg.get('periods', []))
    return (int(by) + (_per[-1] // 12) + 1) if _per else 2036


def _frontier_fmt_sched(sched):
    return '-'.join(str(int(round(float(x)))) for x in sched)


def _frontier_make_assum(base_assum, hz_end, cede_pct, sched, cc_mult, iy_scope, nb_years,
                         claim_scalar, lapse_scalar):
    a = {k: v for k, v in base_assum.items()}
    a['claim_scalar'] = float(claim_scalar)
    a['lapse_scalar'] = float(lapse_scalar)
    # Ceded cohorts: in-force (<=2025) per iy_scope, plus new business 2026..2025+N.
    ceded_iys = sorted(set([iy for iy in iy_scope if iy <= 2025]
                           + list(range(2026, 2026 + int(nb_years)))))
    rp = {}
    for iy in ceded_iys:
        start = max(2026, iy + 1)
        rp[iy] = {cy: float(cede_pct) for cy in range(start, hz_end + 1)}
    a['reins_pct'] = rp
    ccf = {2026 + i: float(sched[i]) * 1e6 for i in range(len(sched)) if float(sched[i]) > 0}
    a['ceding_comm_front'] = ccf
    base_tiers = [[0, 0.75, 250], [0.75, 0.85, 200], [0.85, 0.95, 150],
                  [0.95, float('inf'), 100]]
    a['ceding_comm_table'] = [[r[0], r[1], r[2] * float(cc_mult)] for r in base_tiers]
    return a


def frontier_nodeal_baselines(ev_agg, base_assum, by, surplus_rows, claim_scalars, lapse_scalars):
    """No-deal stressed RBC 2029 ratio per (claim, lapse) environment, for rbc_lift."""
    nodeal_rbc29 = {}
    for cs in claim_scalars:
        for ls in lapse_scalars:
            try:
                nd = {k: v for k, v in base_assum.items()}
                nd['claim_scalar'] = float(cs); nd['lapse_scalar'] = float(ls)
                nd['reins_pct'] = {}
                nd['ceding_comm_front'] = {}
                nd['ceding_comm_table'] = [[0, float('inf'), 0.0]]
                nd_r = run_model(ev_agg, nd, by, rbc_rows=None, bp_rows=None,
                                 surplus_rows=surplus_rows, lite=True)
                nd_na = nd_r['rbc_net_result']['net_adjustments']
                nodeal_rbc29[(float(cs), float(ls))] = (nd_na.get(2029) or {}).get('ratio_w_margin', 0)
            except Exception:
                nodeal_rbc29[(float(cs), float(ls))] = 0.0
    return nodeal_rbc29


def run_frontier_one(ev_agg, base_assum, by, surplus_rows,
                     cede, sched, ccm, iy_scope, nby, cs, ls, nodeal29, n_run=0):
    """Run one (structure x environment) point and return its frontier record.
    Uses run_model lite mode (skips the unused per-IssYear diagnostic / cohort
    metrics) for speed. The viewer calls this per scenario in a JS-driven loop so
    it can show a live counter and yield to the event loop between points."""
    hz_end = _frontier_hz_end(ev_agg, by)
    a = _frontier_make_assum(base_assum, hz_end, cede, sched, ccm, iy_scope, nby, cs, ls)
    r = run_model(ev_agg, a, by, rbc_rows=None, bp_rows=None, surplus_rows=surplus_rows, lite=True)
    mp = r['metrics_predeal']; mn = r['metrics_net']; mc = r['metrics_ceded']
    ap = r['annual_predeal']; an = r['annual_net']
    na = r['rbc_net_result']['net_adjustments']
    ca = r.get('cedant_analytics', {})

    def gy(d, yr):
        return (d.get(yr) or d.get(str(yr)) or 0) / 1e6

    cost = mc['pvde'] / 1e6  # X axis: ceded PVDE handed over ($M)
    net_rbc29 = (na.get(2029) or {}).get('ratio_w_margin', 0)
    rbc_lift = net_rbc29 - (nodeal29 or 0)

    # Early strain relief: net PTI vs predeal PTI, 2026-28 ($M)
    strain = sum(gy(an['pretax_income'], yr) - gy(ap['pretax_income'], yr)
                 for yr in [2026, 2027, 2028])

    pred_ptis = [gy(ap['pretax_income'], yr) for yr in range(2026, 2036)]
    net_ptis = [gy(an['pretax_income'], yr) for yr in range(2026, 2036)]
    def _std(arr):
        m2 = sum(arr) / len(arr)
        return math.sqrt(sum((x - m2) ** 2 for x in arr) / len(arr))
    ps = _std(pred_ptis)
    pti_stability = (ps - _std(net_ptis)) / ps if ps else 0

    cap_relief = ca.get('cap_relief_value')
    cap_relief = cap_relief if cap_relief is not None else 0.0
    benefit = cap_relief + strain  # Y axis: capital + strain relief ($M)

    iy_lo = min(iy_scope); iy_hi = max(iy_scope)
    iy_excluded = [iy for iy in range(iy_lo, iy_hi + 1) if iy not in iy_scope]

    # Cohort economics (commissions allocated: front->back, ongoing->new).
    mb = ca.get('metrics_back') or {}
    mnew = ca.get('metrics_new') or {}
    front_pv = ca.get('comm_front_pv') or 0.0
    back_ceded = mb.get('ceded_pvde') or 0.0

    return {
        'n_run': n_run,
        'cost': cost,
        'cap_relief': cap_relief,
        'strain_relief': strain,
        'benefit': benefit,
        'rbc_lift': rbc_lift,
        'pti_stability': pti_stability,
        'net_pvde': mn['pvde'] / 1e6,
        'pred_pvde': mp['pvde'] / 1e6,
        'net_rbc29': net_rbc29,
        'nodeal_rbc29': nodeal29 or 0,
        'net_irr': mn['irr'],
        'nb_net_irr': mnew.get('net_irr'),
        'nb_net_ev': mnew.get('net_pvde'),
        'nb_predeal_ev': mnew.get('predeal_pvde'),
        'back_net_ev': mb.get('net_pvde'),
        'back_predeal_ev': mb.get('predeal_pvde'),
        'back_ceded_pvde': back_ceded,
        'front_comm_pv': front_pv,
        'back_comp_pct': (front_pv / back_ceded) if back_ceded else None,
        'cede_pct': cede,
        'front_label': _frontier_fmt_sched(sched),
        'cc_mult': ccm,
        'iy_min': iy_lo,
        'iy_max': iy_hi,
        'iy_excluded': iy_excluded,
        'nb_years': nby,
        'claim_scalar': cs,
        'lapse_scalar': ls,
        'is_base': (float(cs) == 1.0 and float(ls) == 1.0),
    }


def run_frontier_grid(ev_agg, base_assum, by, surplus_rows,
                      cede_pcts, front_schedules, cc_mults, iy_scopes, nb_years_list,
                      claim_scalars, lapse_scalars=(1.0,)):
    """Headless grid sweep: no-deal baselines + run_frontier_one over the full grid.
    The viewer drives this same loop from JS (frontier_nodeal_baselines +
    run_frontier_one) for live progress; this wrapper is kept for headless use.

    front_schedules: list of (y2026, y2027, y2028) upfront-commission triples ($M).
    nb_years_list:   list of N = years of new business; each N cedes new-business
                     issue cohorts 2026..2025+N alongside the in-force (<=2025) book
                     selected by iy_scope, for the life of each cohort."""
    nodeal_rbc29 = frontier_nodeal_baselines(ev_agg, base_assum, by, surplus_rows,
                                             claim_scalars, lapse_scalars)
    results = []
    for cs in claim_scalars:
      for ls in lapse_scalars:
        for cede in cede_pcts:
            for sched in front_schedules:
                for ccm in cc_mults:
                    for iy_scope in iy_scopes:
                        for nby in nb_years_list:
                            try:
                                rec = run_frontier_one(
                                    ev_agg, base_assum, by, surplus_rows,
                                    cede, sched, ccm, iy_scope, nby, cs, ls,
                                    nodeal_rbc29.get((float(cs), float(ls)), 0),
                                    n_run=len(results) + 1)
                                results.append(rec)
                            except Exception:
                                pass
    return results


# ============================ DEAL ANALYSIS v2 ============================
# Scenario explorer over a 5-bucket reinsurance triangle with drawn upfront +
# ongoing commission and a claims/lapse stress overlay (for risk transfer).

def _deal_bucket_rate(iy, buckets):
    if iy <= 2019: return buckets[0]
    if 2020 <= iy <= 2024: return buckets[1]
    if iy == 2025: return buckets[2]
    if 2026 <= iy <= 2030: return buckets[3]
    return buckets[4]  # 2031+


def _deal_build_reins(buckets, iss_years, hz_end):
    """5 bucket cede %s (decimals) -> reins_pct triangle.
    buckets = (pre2019, 2020-24, 2025, 2026-30, 2031+)."""
    rp = {}
    for iy in iss_years:
        rate = float(_deal_bucket_rate(iy, buckets))
        if rate <= 0: continue
        iy_key = 2019 if iy <= 2019 else iy
        start = max(2026, iy + 1)
        d = rp.setdefault(iy_key, {})
        for cal in range(start, hz_end + 1):
            d[cal] = rate
    return rp


def deal_draw_baselines(ev_agg, base_assum, by, surplus_rows, stress_envs, predeal_cache=None):
    """Per (claim, lapse) env: the no-deal predeal PVDE ($M) and no-deal RBC 2029
    ratio. Predeal-under-stress is structure-independent, so computed once. Also
    primes predeal_cache (its no-deal predeal == every scenario's predeal at that env)."""
    base = {}
    for (cs, ls) in stress_envs:
        try:
            nd = {k: v for k, v in base_assum.items()}
            nd['claim_scalar'] = float(cs); nd['lapse_scalar'] = float(ls)
            nd['reins_pct'] = {}; nd['ceding_comm_front'] = {}
            nd['ceding_comm_table'] = [[0, float('inf'), 0.0]]
            r = run_model(ev_agg, nd, by, rbc_rows=None, bp_rows=None,
                          surplus_rows=surplus_rows, lite=True, predeal_cache=predeal_cache)
            na = r['rbc_net_result']['net_adjustments']
            base[(float(cs), float(ls))] = {
                'pred_pvde': r['metrics_predeal']['pvde'] / 1e6,
                'nodeal_rbc29': (na.get(2029) or {}).get('ratio_w_margin', 0)}
        except Exception:
            base[(float(cs), float(ls))] = {'pred_pvde': 0.0, 'nodeal_rbc29': 0.0}
    return base


# ===================== VECTORIZED FAST DEAL SWEEP =====================
# Cession + flat commissions are LINEAR in the bucket cede %s, the upfront and the
# ongoing rate, so the whole cashflow side can be precomputed as per-bucket bases
# once per stress env and combined per draw with no re-projection (EF-style). Only
# RBC (non-linear covariance) stays in the engine and is computed lazily.
DA_BACK = (0, 1, 2)   # buckets <=2019 / 2020-24 / 2025
DA_NEW = (3, 4)       # buckets 2026-30 / 2031+


def _ev_filter_iy(ev_agg, keep):
    aiy = {iy: vm for iy, vm in ev_agg.get('agg_iy', {}).items() if keep(int(iy))}
    agg = {}
    for vm in aiy.values():
        for vn, pv in vm.items():
            d = agg.setdefault(vn, {})
            for p, v in pv.items():
                d[p] = d.get(p, 0) + v
    return {'agg': agg, 'agg_iy': aiy, 'periods': ev_agg.get('periods', []),
            'iss_years': sorted(int(k) for k in aiy)}


def _vec(annual, key, years):
    d = annual.get(key, {}) or {}
    return [float(d.get(str(y), d.get(y, 0)) or 0) for y in years]


def _pvde_list(de, disc):
    pvde = sum(v / (1 + disc) ** i for i, v in enumerate(de, 1))
    irr = None
    try:
        def npv(r):
            return sum(v / (1 + r) ** i for i, v in enumerate(de, 1))
        if npv(-0.5) * npv(5.0) < 0:
            lo, hi = -0.5, 5.0
            for _ in range(80):
                mid = (lo + hi) / 2
                if npv(mid) > 0:
                    lo = mid
                else:
                    hi = mid
            irr = (lo + hi) / 2
    except Exception:
        pass
    return pvde, irr


def _deal_assum(base_assum, iss, hz_end, cs, ls, buckets, upfront, ongoing):
    a = {k: v for k, v in base_assum.items()}
    a['claim_scalar'] = float(cs); a['lapse_scalar'] = float(ls)
    a['reins_pct'] = _deal_build_reins(buckets, iss, hz_end)
    ut = float(upfront)
    a['ceding_comm_front'] = {y: amt * 1e6 for y, amt in
                              ((2026, ut * 0.5), (2027, ut * 0.25), (2028, ut * 0.25)) if amt > 0}
    a['ceding_comm_table'] = [[0, float('inf'), float(ongoing)]]
    return a


def deal_precompute_bases(ev_agg, base_assum, by, surplus_rows, stress_envs):
    """Per stress env: predeal DE, and per-bucket ceded-DE/PTI bases G_b (gross,
    100% cede), Cann_b/CP_b (ongoing-commission per $1), F/FP (front per $1 upfront).
    Base env also gets cohort predeal + the PTI bases for strain."""
    hz_end = _frontier_hz_end(ev_agg, by)
    years = list(range(2026, hz_end + 1))
    disc = float(base_assum.get('discount_rate', 0.08))
    iss = sorted(int(k) for k in ev_agg.get('agg_iy', {}).keys())
    ev_back = _ev_filter_iy(ev_agg, lambda iy: iy <= 2025)
    ev_new = _ev_filter_iy(ev_agg, lambda iy: iy >= 2026)

    def ced(cs, ls, buckets, upfront, ongoing):
        r = run_model(ev_agg, _deal_assum(base_assum, iss, hz_end, cs, ls, buckets, upfront, ongoing),
                      by, surplus_rows=surplus_rows, metrics_only=True)
        return (_vec(r['annual_ceded'], 'distributable_earnings', years),
                _vec(r['annual_ceded'], 'pretax_income', years))

    def pred(cs, ls, ev):
        r = run_model(ev, _deal_assum(base_assum, iss, hz_end, cs, ls, [0] * 5, 0, 0),
                      by, surplus_rows=surplus_rows, metrics_only=True)
        return _vec(r['annual_predeal'], 'distributable_earnings', years)

    one = lambda i: [1.0 if j == i else 0.0 for j in range(5)]
    ny = len(years)
    bases = []
    for (cs, ls) in stress_envs:
        is_base = float(cs) == 1.0 and float(ls) == 1.0
        Dpre = pred(cs, ls, ev_agg)
        G = []; GP = []; Cann = []; CP = []
        for b in range(5):
            gde, gpti = ced(cs, ls, one(b), 0, 0)
            cde, cpti = ced(cs, ls, one(b), 0, 1)            # ongoing = $1
            G.append(gde); GP.append(gpti)
            Cann.append([gde[y] - cde[y] for y in range(ny)])  # ongoing-comm DE per $1, per 100% cede
            CP.append([gpti[y] - cpti[y] for y in range(ny)])
        fde, fpti = ced(cs, ls, [0] * 5, 1, 0)               # upfront = $1
        bz = {'env': (cs, ls), 'is_base': is_base, 'Dpre': Dpre,
              'pred_pvde': _pvde_list(Dpre, disc)[0] / 1e6,
              'G': G, 'F': [-x for x in fde], 'Cann': Cann}
        if is_base:
            bz.update({'GP': GP, 'FP': [-x for x in fpti], 'CP': CP,
                       'Dpre_back': pred(cs, ls, ev_back), 'Dpre_new': pred(cs, ls, ev_new)})
        bases.append(bz)
    return {'years': years, 'disc': disc, 'bases': bases}


def deal_fast_sweep(packed, draws):
    """Combine the precomputed bases for each draw (pure arithmetic, no re-run).
    Returns the cashflow-side record per draw (RBC filled in lazily later)."""
    years = packed['years']; disc = packed['disc']; bases = packed['bases']; ny = len(years)

    def _std(a):
        if len(a) < 2:
            return 0.0
        m = sum(a) / len(a)
        return math.sqrt(sum((x - m) ** 2 for x in a) / len(a))

    out = []
    for d in draws:
        bk = [float(x) for x in d['buckets']]; up = float(d['upfront']); on = float(d['ongoing'])
        net_pvdes = []; pred_pvdes = []; bm = {}
        for bz in bases:
            G = bz['G']; F = bz['F']; Cann = bz['Cann']
            ced_DE = [sum(bk[b] * G[b][y] for b in range(5)) - up * F[y]
                      - on * sum(bk[b] * Cann[b][y] for b in range(5)) for y in range(ny)]
            net_DE = [bz['Dpre'][y] - ced_DE[y] for y in range(ny)]
            net_pvde = _pvde_list(net_DE, disc)[0]
            net_pvdes.append(net_pvde / 1e6); pred_pvdes.append(bz['pred_pvde'])
            if bz['is_base']:
                cost = _pvde_list(ced_DE, disc)[0] / 1e6
                _, net_irr = _pvde_list(net_DE, disc)
                back_ced = [sum(bk[b] * G[b][y] for b in DA_BACK) for y in range(ny)]
                new_ced = [sum(bk[b] * G[b][y] for b in DA_NEW) for y in range(ny)]
                back_net = [bz['Dpre_back'][y] - back_ced[y] + up * F[y] for y in range(ny)]
                new_net = [bz['Dpre_new'][y] - new_ced[y]
                           + on * sum(bk[b] * Cann[b][y] for b in DA_NEW) for y in range(ny)]
                back_pre = _pvde_list(bz['Dpre_back'], disc)[0] / 1e6
                new_pre, nb_pre_irr = _pvde_list(bz['Dpre_new'], disc)
                back_net_pvde = _pvde_list(back_net, disc)[0] / 1e6
                new_net_pvde, nb_net_irr = _pvde_list(new_net, disc)
                back_ceded_pvde = _pvde_list(back_ced, disc)[0] / 1e6
                front_pv = up * _pvde_list(F, disc)[0] / 1e6
                GP = bz['GP']; FP = bz['FP']; CP = bz['CP']
                strain = 0.0
                for yi, y in enumerate(years):
                    if y in (2026, 2027, 2028):
                        strain += (-sum(bk[b] * GP[b][yi] for b in range(5)) + up * FP[yi]
                                   + on * sum(bk[b] * CP[b][yi] for b in range(5))) / 1e6
                bm = dict(cost=cost, net_pvde=net_pvde / 1e6, net_irr=net_irr,
                          nb_net_irr=nb_net_irr, nb_dEV=new_pre / 1e6 - new_net_pvde / 1e6,
                          back_dEV=back_pre - back_net_pvde, strain_relief=strain,
                          nb_dIRR=(None if nb_net_irr is None or nb_pre_irr is None else nb_net_irr - nb_pre_irr),
                          back_comp_pct=(front_pv / back_ceded_pvde) if back_ceded_pvde else None)
        sps = _std(pred_pvdes)
        rt = (1 - _std(net_pvdes) / sps) if sps else 0.0
        rec = dict(n_run=d['n'], buckets=bk, upfront=up, ongoing=on, risk_transfer=rt,
                   rbc_lift=None, cap_relief_value=None)
        rec.update(bm)
        out.append(rec)
    return out


def deal_rbc_for(ev_agg, base_assum, by, surplus_rows, buckets, upfront, ongoing, nodeal29):
    """Base-env full run for the non-linear RBC metrics (lazy pass)."""
    hz_end = _frontier_hz_end(ev_agg, by)
    iss = sorted(int(k) for k in ev_agg.get('agg_iy', {}).keys())
    r = run_model(ev_agg, _deal_assum(base_assum, iss, hz_end, 1.0, 1.0, buckets, upfront, ongoing),
                  by, surplus_rows=surplus_rows, lite=True)
    na = r['rbc_net_result']['net_adjustments']
    return {'rbc_lift': (na.get(2029) or {}).get('ratio_w_margin', 0) - (nodeal29 or 0),
            'cap_relief_value': r.get('cedant_analytics', {}).get('cap_relief_value')}


def run_deal_scenario(ev_agg, base_assum, by, surplus_rows,
                      buckets, upfront_total, ongoing, stress_envs, baselines, n_run=0,
                      predeal_cache=None):
    """Score one drawn structure (5-bucket triangle + upfront + ongoing) across the
    claims/lapse stress overlay. Returns one record for the constraint finder.
    Base env runs full (RBC + cohort); stress envs run metrics_only (net PVDE only)."""
    hz_end = _frontier_hz_end(ev_agg, by)
    iss_years = sorted(int(k) for k in ev_agg.get('agg_iy', {}).keys())
    reins = _deal_build_reins(buckets, iss_years, hz_end)
    # Upfront total split 50/25/25 over 2026/27/28 (the 10-5-5 shape).
    ut = float(upfront_total)
    front = {y: a * 1e6 for y, a in ((2026, ut * 0.5), (2027, ut * 0.25), (2028, ut * 0.25)) if a > 0}
    table = [[0, float('inf'), float(ongoing)]]

    def assum_for(cs, ls):
        a = {k: v for k, v in base_assum.items()}
        a['claim_scalar'] = float(cs); a['lapse_scalar'] = float(ls)
        a['reins_pct'] = reins; a['ceding_comm_front'] = front
        a['ceding_comm_table'] = table
        return a

    net_pvdes = []; pred_pvdes = []; base_rec = None
    for (cs, ls) in stress_envs:
        is_base = float(cs) == 1.0 and float(ls) == 1.0
        try:
            r = run_model(ev_agg, assum_for(cs, ls), by, rbc_rows=None,
                          bp_rows=None, surplus_rows=surplus_rows, lite=True,
                          metrics_only=not is_base, predeal_cache=predeal_cache)
        except Exception:
            continue
        net_pvdes.append(r['metrics_net']['pvde'] / 1e6)
        pred_pvdes.append((baselines.get((float(cs), float(ls))) or {}).get('pred_pvde', 0.0))
        if is_base:
            base_rec = r

    def _std(arr):
        if len(arr) < 2: return 0.0
        m = sum(arr) / len(arr)
        return math.sqrt(sum((x - m) ** 2 for x in arr) / len(arr))
    sp = _std(pred_pvdes)
    risk_transfer = (1 - _std(net_pvdes) / sp) if sp else 0.0

    if base_rec is None:
        return None
    mp = base_rec['metrics_predeal']; mn = base_rec['metrics_net']; mc = base_rec['metrics_ceded']
    ap = base_rec['annual_predeal']; an = base_rec['annual_net']
    na = base_rec['rbc_net_result']['net_adjustments']
    ca = base_rec.get('cedant_analytics', {})
    mb = ca.get('metrics_back') or {}; mnew = ca.get('metrics_new') or {}

    def gy(d, yr):
        return (d.get(yr) or d.get(str(yr)) or 0) / 1e6
    strain = sum(gy(an['pretax_income'], yr) - gy(ap['pretax_income'], yr) for yr in (2026, 2027, 2028))
    net_rbc29 = (na.get(2029) or {}).get('ratio_w_margin', 0)
    nodeal29 = (baselines.get((1.0, 1.0)) or {}).get('nodeal_rbc29', 0)
    front_pv = ca.get('comm_front_pv') or 0.0
    back_ceded = mb.get('ceded_pvde') or 0.0
    pred_pvde = mp['pvde'] / 1e6; net_pvde = mn['pvde'] / 1e6

    return {
        'n_run': n_run,
        'buckets': [float(b) for b in buckets],
        'upfront': ut, 'ongoing': float(ongoing),
        'cost': mc['pvde'] / 1e6,
        'rbc_lift': net_rbc29 - nodeal29,
        'risk_transfer': risk_transfer,
        'nb_net_irr': mnew.get('net_irr'),
        'nb_dEV': (mnew.get('predeal_pvde') or 0) - (mnew.get('net_pvde') or 0),
        'back_dEV': (mb.get('predeal_pvde') or 0) - (mb.get('net_pvde') or 0),
        'nb_dIRR': (None if mnew.get('net_irr') is None or mnew.get('predeal_irr') is None
                    else mnew['net_irr'] - mnew['predeal_irr']),
        'strain_relief': strain,
        'back_comp_pct': (front_pv / back_ceded) if back_ceded else None,
        'net_deal_value': ca.get('net_deal_value'),
        'cap_relief_value': ca.get('cap_relief_value'),
        'ev_given_up': pred_pvde - net_pvde,
        'net_pvde': net_pvde,
        'net_irr': mn['irr'],
    }


def run_model(ev_agg,assum,by,rbc_rows=None,bp_rows=None,surplus_rows=None,lite=False,
              metrics_only=False,predeal_cache=None):
    # lite=True skips the per-issue-year diagnostic. metrics_only=True returns just the
    # predeal/ceded/net metrics (skips the RBC pipeline, cedant analytics and back/new
    # cohort metrics) for the Deal Analysis stress envs. predeal_cache (a dict keyed by
    # (claim_scalar, lapse_scalar, by)) reuses the structure-independent predeal side
    # (scaled aggregates, sp, ap, mp) across scenario draws. Full path output unchanged.
    rp={int(iy):{int(c2):v for c2,v in yd.items() if v is not None}
        for iy,yd in assum.get("reins_pct",{}).items()}
    periods=sorted(ev_agg.get("periods",[]))
    by=int(by)
    _ck=(float(assum.get("claim_scalar",1.0)),float(assum.get("lapse_scalar",1.0)),by)
    if predeal_cache is not None and _ck in predeal_cache:
        agg_d,agg_iy,sp,_ap_cache,_mp_cache=predeal_cache[_ck]
    else:
        agg_d,agg_iy=apply_scalars(ev_agg.get("agg",{}),ev_agg.get("agg_iy",{}),assum)
        sp=bld_stmt(agg_d,assum,by,periods,False)
        _ap_cache=None;_mp_cache=None
    agg_c=apply_rp_agg(agg_iy,rp,by,periods)
    sc=bld_stmt(agg_c,assum,by,periods,True,zero_bop_ts=True)
    # Build net: predeal - ceded + ceding commissions
    sn={}
    for k in sp:
        if isinstance(sp[k],dict):sn[k]={p:sp[k].get(p,0)-sc[k].get(p,0) for p in periods}
        else:sn[k]=sp[k]
    # Ceding commissions
    ccf={int(k):float(v) for k,v in assum.get("ceding_comm_front",{}).items()}
    # Sliding-scale ceding commission table: [[lr_min, lr_max, comm], ...]
    cco_table=assum.get("ceding_comm_table",None)
    if not cco_table:
        flat=float(assum.get("ceding_comm_ongoing",200))
        cco_table=[[0,0.75,flat],[0.75,0.85,flat],[0.85,0.95,flat],[0.95,float("inf"),flat]]
    # Compute LR by IssYear to determine per-IY ongoing ceding commission
    lr_by_iy=compute_lr_by_iy(agg_iy,rp,by,periods)
    iy_cco={iy:lookup_cco(lr,cco_table) for iy,lr in lr_by_iy.items()}
    # Compute lives issued per calendar year for ongoing cc
    iss_by_cy=defaultdict(float)
    for p in periods:
        if p>0:iss_by_cy[cy(by,p)]+=sp["lives_issued"].get(p,0)
    # For each period, compute ceding comms
    # Net DE: BOP_TS = predeal BOP, EOP_TS = predeal_EOP - ceded_EOP
    sp_ts=sp.get("target_surplus",{});sc_ts=sc.get("target_surplus",{})
    net_pti=sn.get("pretax_income",{});sn_ts=sn.get("target_surplus",{})
    min_p=min(periods) if periods else 0
    def _net_bop(p):
        return sp_ts.get(p-1,sp_ts.get(min_p,0))  # ceded BOP=0
    sn["distributable_earnings"]={p:net_pti.get(p,0)*0.79-(sn_ts.get(p,0)-_net_bop(p)) for p in periods}
    for p in periods:
        cal=cy(by,p)
        # Front-end: paid once at Dec of each calendar year (p%12==0 for p=12,24,...)
        c1=float(ccf.get(cal,0)) if p>0 and p%12==0 else 0.0
        # Ongoing: prior calendar year's issued policies * avg_cede_rate * $200 / 12
        # "policies issued in prior year that are being reinsured for the first time"
        prior_cal=cal-1
        prior_issued=iss_by_cy.get(prior_cal,0)
        # Get the rate that applies to IssYr=prior_cal in CalYr=cal
        iy_key_prior=2019 if prior_cal<=2019 else prior_cal
        prior_rates=rp.get(iy_key_prior,{})
        prior_max_c=max(prior_rates.keys()) if prior_rates else 0
        prior_rate=float(prior_rates.get(min(cal,prior_max_c),0) or 0) if prior_rates else 0
        iy_comm=iy_cco.get(prior_cal,lookup_cco(lr_by_iy.get(prior_cal,0.80),cco_table))
        c2=prior_issued*prior_rate*iy_comm/12 if p>0 else 0.0
        sn.setdefault("comm1",{})[p]=c1
        sn.setdefault("comm2",{})[p]=c2
        # Ceded tab: ceding comms are a COST (negative) in ceded selling expense
        # Front-end comm received by direct carrier = negative in ceded IS (reduces ceded income)
        sc.setdefault("comm1",{})[p]=-c1  # negative on ceded tab
        sc.setdefault("comm2",{})[p]=-c2
        sc["selling_expense"][p]=sc["selling_expense"].get(p,0)+(-c1)+(-c2)
        sc["pretax_income"][p]=sc["revenue"].get(p,0)+sc["benefits"].get(p,0)+sc["selling_expense"][p]+sc["op_expense"].get(p,0)
        # Net: add ceding comms as income
        sn["revenue"][p]=sn["revenue"].get(p,0)+c1+c2
        sn["pretax_income"][p]=sn["revenue"][p]+sn["benefits"].get(p,0)+sn["selling_expense"].get(p,0)+sn["op_expense"].get(p,0)
        ts=sn.get("target_surplus",{})
        sn["distributable_earnings"][p]=sn["pretax_income"][p]-(ts.get(p,0)-ts.get(p-1,0))
    disc=float(assum.get("discount_rate",0.08))
    def _ann_pvde(ann_de,disc_r):
        yrs=sorted(int(k) for k in ann_de.keys())
        pvde=sum(ann_de[str(yr)]/(1+disc_r)**i for i,yr in enumerate(yrs,1))
        cum=mn2=0
        for yr in yrs:
            cum+=ann_de[str(yr)]
            if cum<mn2:mn2=cum
        dl=[ann_de[str(yr)] for yr in yrs]
        irr=None
        try:
            def npv(r):return sum(v/(1+r)**i for i,v in enumerate(dl,1))
            if npv(-0.5)*npv(5.0)<0:
                lo,hi=-0.5,5.0
                for _ in range(80):
                    mid=(lo+hi)/2
                    if npv(mid)>0:lo=mid
                    else:hi=mid
                irr=(1+(lo+hi)/2)-1
        except:pass
        return {"pvde":pvde,"irr":irr,"max_neg_cum_de":mn2}
    _sp_ts=sp.get("target_surplus",{}); _sc_ts=sc.get("target_surplus",{})
    # Predeal side is structure-independent (depends only on claim/lapse) -> cache it.
    if _ap_cache is not None:
        ap=_ap_cache; mp=_mp_cache
    else:
        ap=ann_eoy(sp,by,periods)
        ap["distributable_earnings"]=compute_ann_de(ap["pretax_income"],_sp_ts,_sc_ts,by,periods,"predeal")
        mp=_ann_pvde(ap["distributable_earnings"],disc)
        if predeal_cache is not None:
            predeal_cache[_ck]=(agg_d,agg_iy,sp,ap,mp)
    ac=ann_eoy(sc,by,periods)
    ac["distributable_earnings"]=compute_ann_de(ac["pretax_income"],_sp_ts,_sc_ts,by,periods,"ceded")
    aN=ann_eoy(sn,by,periods)
    # Net DE = Predeal DE - Ceded DE (ensures internal consistency)
    _de_p=ap["distributable_earnings"];_de_c=ac["distributable_earnings"]
    _all_yrs=sorted(set(list(_de_p.keys())+list(_de_c.keys())))
    aN["distributable_earnings"]={yr:(_de_p.get(yr,0)-_de_c.get(yr,0)) for yr in _all_yrs}
    mc=_ann_pvde(ac["distributable_earnings"],disc)
    mn=_ann_pvde(aN["distributable_earnings"],disc)
    if metrics_only:
        return {"metrics_predeal":mp,"metrics_ceded":mc,"metrics_net":mn,
                "annual_predeal":ser(ap),"annual_ceded":ser(ac),"annual_net":ser(aN)}
    iys=sorted(ev_agg.get("iss_years",[]))
    iyd={}
    for iy in ([] if lite else iys):
        iy_d=agg_iy.get(str(iy),{})
        if not iy_d:continue
        single_agg_iy={str(iy):iy_d}
        iy_c=apply_rp_agg(single_agg_iy,rp,by,periods)
        _sp=bld_stmt(iy_d,assum,by,periods,False)
        _sc=bld_stmt(iy_c,assum,by,periods,True)
        sd=ann_eoy(_sp,by,periods);sc2=ann_eoy(_sc,by,periods)
        rows=[]
        for cy3 in ap.get("periods",[]):
            iy_key=2019 if iy<=2019 else iy
            iy_rates=rp.get(iy_key,{})
            exp=None
            if iy_rates and iy<=cy3:
                max_c=max(iy_rates.keys())
                r=iy_rates.get(min(cy3,max_c))
                if r:exp=r
            lr=[]
            for lbl,key in DIAG_LINES:
                dp=sd.get(key,{}).get(cy3,0);cp=sc2.get(key,{}).get(cy3,0)
                # Skip if both zero (e.g. acq expense = 0 for in-force cohorts)
                if dp==0 and cp==0:
                    lr.append({"line":lbl,"direct":dp,"ceded":cp,
                               "expected_rate":exp,"actual_rate":None,"ok":True,"zero_skip":True})
                    continue
                # Yr1 adjustment for lines where ceded BOP reserve = 0 but direct BOP > 0
                # Affects delta_reserves (hence Benefits) and NII (hence PTI)
                dp_adj=dp
                adj_exp=exp
                if exp and cy3==by+1:
                    iy_pr_bop=iy_d.get("TabRes",{}).get(0,0)+iy_d.get("CededALRstat",{}).get(0,0)
                    if lbl=="Benefits" and iy_pr_bop!=0 and dp!=0:
                        # Ceded delta_res uses BOP=0 so is off by rate*pr_p0
                        # expected_ceded_benefits = rate*(dp - pr_p0)
                        dp_adj=dp-iy_pr_bop
                        adj_exp=exp*dp_adj/dp
                    elif lbl=="Pretax Income" and dp!=0:
                        # PTI includes both the benefits adjustment and the NII adjustment
                        # expected_ceded_PTI = rate*(d_pti - d_nii) + c_nii - rate*pr_p0
                        d_nii=sd.get("nii",{}).get(cy3,0)
                        c_nii=sc2.get("nii",{}).get(cy3,0)
                        exp_cp=exp*(dp-d_nii)+c_nii-exp*iy_pr_bop
                        adj_exp=exp_cp/dp if dp!=0 else exp
                ar_raw=cp/dp if dp!=0 else None
                # For ok check: actual adj_rate = adj_exp (both derived from same formula)
                ar_adj=cp/dp_adj if dp_adj!=0 else ar_raw
                ok=True
                if exp and exp>0:
                    if lbl=="Pretax Income" and dp!=0:
                        # Compare actual vs formula-expected directly
                        ok=(ar_raw is not None and abs(ar_raw-adj_exp)<0.005)
                    else:
                        ok=(ar_adj is not None and abs(ar_adj-exp)<0.005)
                lr.append({"line":lbl,"direct":dp,"ceded":cp,
                           "expected_rate":adj_exp,"actual_rate":ar_raw,"ok":ok,"zero_skip":False})
            rows.append({"cal_year":cy3,"expected_rate":exp,"lines":lr})
        iyd[iy]=rows
    rbc_orig={};bp={}
    if rbc_rows:rbc_orig=parse_rbc(rbc_rows,surplus_rows)
    if bp_rows:bp=parse_bp(bp_rows)
    rbc_net=compute_rbc_net(rbc_orig,ser(ap),ser(ac))
    # Compute new-style RBC tabs
    _src_rows=rbc_rows if rbc_rows else surplus_rows
    _rbc_full=parse_rbc_full(_src_rows) if _src_rows else {"years":[],"data":{}}
    _rbc_orig_computed=compute_original_rbc(_rbc_full) if _rbc_full["data"] else {"years":[],"original_values":{}}
    _rbc_predeal_result=compute_predeal_rbc(_rbc_full,ser(ap),ser(rbc_orig)) if _rbc_full["data"] else {}
    _rbc_net_result=compute_net_rbc(_rbc_predeal_result,ser(aN),ser(ap),_rbc_full) if _rbc_predeal_result else {}
    # ===================== CEDANT DECISION ANALYTICS =====================
    # All decision figures expressed in $M for readability.
    # (1) Reinsurer economics. The model is internally zero-sum: mp = mn + mc,
    #     and the ceded statement already nets out ceding commissions, so
    #     metrics_ceded.pvde IS the reinsurer's PV of distributable earnings.
    _comm1_ann=defaultdict(float);_comm2_ann=defaultdict(float)
    for _p in periods:
        if _p>0:
            _comm1_ann[cy(by,_p)]+=sn.get("comm1",{}).get(_p,0)
            _comm2_ann[cy(by,_p)]+=sn.get("comm2",{}).get(_p,0)
    def _pv_ann(d):
        _ys=sorted(d.keys());return sum(d[_y]/(1+disc)**_i for _i,_y in enumerate(_ys,1))
    comm_front_pv=_pv_ann({_y:_comm1_ann[_y]*0.79 for _y in _comm1_ann})/1e6
    comm_ong_pv=_pv_ann({_y:_comm2_ann[_y]*0.79 for _y in _comm2_ann})/1e6
    comm_pv=comm_front_pv+comm_ong_pv                 # after-tax PV of commissions received ($M)
    reinsurer_pvde=mc["pvde"]/1e6                      # reinsurer value = ceded DE net of commissions ($M)
    ceded_gross_pvde=reinsurer_pvde+comm_pv           # lifetime value cedant gives up, pre-commission ($M)
    value_recovery_pct=(comm_pv/ceded_gross_pvde) if ceded_gross_pvde else None
    cedant_giveup_pvde=(mp["pvde"]-mn["pvde"])/1e6    # = reinsurer_pvde by construction ($M)
    # (2) Capital-relief value: required capital freed (total_post_cov, $M) * cost of capital,
    #     discounted at the PVDE rate. Requires RBC/surplus data; None if unavailable.
    coc=float(assum.get("cost_of_capital",0.10))
    cap_relief_value=None;cap_relief_by_yr={}
    _pa=_rbc_predeal_result.get("predeal_adjustments",{}) if _rbc_predeal_result else {}
    _na=_rbc_net_result.get("net_adjustments",{}) if _rbc_net_result else {}
    if _pa and _na:
        _ys=sorted(set(int(_y) for _y in _pa.keys()));_b=min(_ys) if _ys else None
        _relief=[]
        for _y in _ys:
            if _y==_b:continue
            _pre=(_pa.get(_y) or _pa.get(str(_y)) or {}).get("total_post_cov",0)
            _net=(_na.get(_y) or _na.get(str(_y)) or {}).get("total_post_cov",0)
            cap_relief_by_yr[_y]=_pre-_net
            _relief.append((_pre-_net)*coc)
        cap_relief_value=sum(_v/(1+disc)**_i for _i,_v in enumerate(_relief,1))
    # (3) Net Deal Value to the cedant ($M): capital relief gained minus PVDE handed to reinsurer.
    #     Deal is value-accretive when relief outweighs the profit ceded.
    net_deal_value=(cap_relief_value-reinsurer_pvde) if cap_relief_value is not None else None
    # (4) Back-book (IssYr<=2025) vs New-issue (IssYr>=2026) cohort economics.
    #     Predeal = inherent profitability; net = retained after cession (ex portfolio commissions).
    def _cohort_metrics(is_back):
        _sub={_iy:_vm for _iy,_vm in agg_iy.items() if (int(_iy)<=2025)==is_back}
        if not _sub:return None
        _ag=defaultdict(lambda:defaultdict(float))
        for _iy,_vm in _sub.items():
            for _vn,_pv in _vm.items():
                for _p,_v in _pv.items():_ag[_vn][_p]+=_v
        _ag={_k:dict(_v) for _k,_v in _ag.items()}
        _sp=bld_stmt(_ag,assum,by,periods,False)
        _sc=bld_stmt(apply_rp_agg(_sub,rp,by,periods),assum,by,periods,True,zero_bop_ts=True)
        _ap=ann_eoy(_sp,by,periods);_ac=ann_eoy(_sc,by,periods)
        _spt=_sp.get("target_surplus",{});_sct=_sc.get("target_surplus",{})
        _dep=compute_ann_de(_ap["pretax_income"],_spt,_sct,by,periods,"predeal")
        _dec=compute_ann_de(_ac["pretax_income"],_spt,_sct,by,periods,"ceded")
        _yrs=sorted(set(list(_dep.keys())+list(_dec.keys())))
        # Allocate ceding commissions into the cohort net DE so cohort nets reconcile
        # to the portfolio: the front-end (10-5-5) compensates for ceding the in-force
        # book -> back-book; the ongoing per-policy commission is computed only on
        # newly-issued business -> new-issue. After-tax (x0.79), mirroring net revenue.
        _comm_ann=_comm1_ann if is_back else _comm2_ann
        _den={_yr:_dep.get(_yr,0)-_dec.get(_yr,0)+_comm_ann.get(int(_yr),0)*0.79 for _yr in _yrs}
        _mp=_ann_pvde(_dep,disc);_mn=_ann_pvde(_den,disc);_mc=_ann_pvde(_dec,disc)
        return {"predeal_pvde":_mp["pvde"]/1e6,"predeal_irr":_mp["irr"],
                "net_pvde":_mn["pvde"]/1e6,"net_irr":_mn["irr"],
                "ceded_pvde":_mc["pvde"]/1e6}
    metrics_back=_cohort_metrics(True)
    metrics_new=_cohort_metrics(False)
    cedant_analytics={"reinsurer_pvde":reinsurer_pvde,"ceded_gross_pvde":ceded_gross_pvde,
        "comm_pv":comm_pv,"comm_front_pv":comm_front_pv,"comm_ong_pv":comm_ong_pv,
        "value_recovery_pct":value_recovery_pct,"cedant_giveup_pvde":cedant_giveup_pvde,
        "cost_of_capital":coc,"cap_relief_value":cap_relief_value,
        "cap_relief_by_yr":{str(_k):_v for _k,_v in cap_relief_by_yr.items()},
        "net_deal_value":net_deal_value,
        "pred_pvde":mp["pvde"]/1e6,"net_pvde":mn["pvde"]/1e6,
        "metrics_back":metrics_back,"metrics_new":metrics_new}
    # ====================================================================
    return {"stmt_predeal":ser(sp),"stmt_ceded":ser(sc),"stmt_net":ser(sn),
            "annual_predeal":ser(ap),"annual_ceded":ser(ac),"annual_net":ser(aN),
            "metrics_predeal":mp,"metrics_ceded":mc,"metrics_net":mn,
            "iss_years":iys,"iy_diagnostic":{str(k):v for k,v in iyd.items()},
            "rbc_data":ser(rbc_orig),"rbc_net":ser(rbc_net),
            "rbc_orig_full":_rbc_full,"rbc_orig_computed":_rbc_orig_computed,"rbc_predeal_result":_rbc_predeal_result,"rbc_net_result":_rbc_net_result,
            "bp_data":ser(bp),
            "cedant_analytics":ser(cedant_analytics),
            "max_period":max(periods) if periods else 0,
            "ev_records_count":ev_agg.get("row_count",0)}

def run_scenario_matrix(ev_agg, base_assum, by, rbc_rows=None, bp_rows=None, surplus_rows=None,
                        cc2026s=(10,8,5), cc2027s=(5,2,0), cc2028s=(5,2,0),
                        lr_splits=(1.0,0.9,0.8),
                        r_pre2019s=(10,12,15), r_2021_24s=(10,8,12),
                        r_2025s=(10,5,0), r_2026ps=(10,8,12),
                        stress_envs=((1.0,1.0),(1.025,1.0),(1.0,1.10),(1.025,1.10)),
                        horizon_end=2035, limit=None):
    """Full combinatoric cedant scenario sweep with claims/lapse stress as an overlay.
    Each of 3^8=6561 structures is scored under 4 environments. 2020 issue year is
    grouped with the 2021-2024 cohort (per spec). Returns list of scenario dicts."""
    iss_years=sorted(int(k) for k in ev_agg.get("agg_iy",{}).keys())
    base_tiers=[[0,0.75,250],[0.75,0.85,200],[0.85,0.95,150],[0.95,float('inf'),100]]
    def cohort_rate(iy,r_pre,r_mid,r_25,r_26p):
        if iy<=2019:return r_pre
        if 2020<=iy<=2024:return r_mid
        if iy==2025:return r_25
        return r_26p
    def build_reins(r_pre,r_mid,r_25,r_26p):
        rp={}
        for iy in iss_years:
            pct=cohort_rate(iy,r_pre,r_mid,r_25,r_26p)/100.0
            if pct<=0:continue
            start=max(2026,iy+1);iy_key=2019 if iy<=2019 else iy
            d=rp.setdefault(iy_key,{})
            for cyr in range(start,horizon_end+1):d[cyr]=pct
        return rp
    def gy(d,yr):return (d.get(str(yr)) or d.get(yr) or 0)
    results=[];sid=0
    for cc26 in cc2026s:
     for cc27 in cc2027s:
      for cc28 in cc2028s:
       for lrs in lr_splits:
        for rp19 in r_pre2019s:
         for rmid in r_2021_24s:
          for r25 in r_2025s:
           for r26 in r_2026ps:
            sid+=1
            if limit and sid>limit:return results
            reins=build_reins(rp19,rmid,r25,r26)
            front={}
            for yr,amt in ((2026,cc26),(2027,cc27),(2028,cc28)):
                if amt>0:front[yr]=amt*1e6
            tiers=[[t[0],t[1],t[2]*lrs] for t in base_tiers]
            env_recs=[];base_rec=None
            for (cs,ls) in stress_envs:
                a=dict(base_assum)
                a["reins_pct"]=reins;a["ceding_comm_front"]=front
                a["ceding_comm_table"]=tiers;a["claim_scalar"]=cs;a["lapse_scalar"]=ls
                try:
                    r=run_model(ev_agg,a,by,rbc_rows,bp_rows,surplus_rows)
                    ca=r.get("cedant_analytics",{})
                    an=r.get("annual_net",{});ap=r.get("annual_predeal",{})
                    cum_net_pti=sum(gy(an.get("pretax_income",{}),yr) for yr in (2026,2027,2028))/1e6
                    cum_pred_pti=sum(gy(ap.get("pretax_income",{}),yr) for yr in (2026,2027,2028))/1e6
                    rec={"claim":cs,"lapse":ls,
                         "net_deal_value":ca.get("net_deal_value"),
                         "cap_relief_value":ca.get("cap_relief_value"),
                         "reinsurer_pvde":ca.get("reinsurer_pvde"),
                         "net_pvde":ca.get("net_pvde"),"pred_pvde":ca.get("pred_pvde"),
                         "value_recovery_pct":ca.get("value_recovery_pct"),
                         "cum_net_pti_3yr":cum_net_pti,"cum_pred_pti_3yr":cum_pred_pti,
                         "back":ca.get("metrics_back"),"new":ca.get("metrics_new")}
                    env_recs.append(rec)
                    if cs==1.0 and ls==1.0:base_rec=rec
                except Exception:
                    env_recs.append({"claim":cs,"lapse":ls,"error":True})
            prot=None
            try:
                comb=next((e for e in env_recs if e.get("claim",1)>1.0 and e.get("lapse",1)>1.0 and not e.get("error")),None)
                if base_rec and comb:
                    pred_drop=base_rec["cum_pred_pti_3yr"]-comb["cum_pred_pti_3yr"]
                    net_drop=base_rec["cum_net_pti_3yr"]-comb["cum_net_pti_3yr"]
                    prot=pred_drop-net_drop   # $M of 3yr earnings deterioration absorbed by reinsurer
            except Exception:pass
            results.append({
                "scenario_id":"S%05d"%sid,
                "cc_2026":cc26,"cc_2027":cc27,"cc_2028":cc28,"lr_split":lrs,
                "reins_pre2019":rp19,"reins_2021_24":rmid,"reins_2025":r25,"reins_2026p":r26,
                "base":base_rec,"stress":env_recs,"downside_protection_3yr":prot})
    return results

print("Engine v7 ready")
