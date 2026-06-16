// Blockbuster Deals — app layer. Engine and data are loaded from src/ and data/ at startup.
let PY_ENGINE, DEFAULT_EV_CSV, DOC_TEXT, DEFAULT_SURPLUS_ROWS, DEFAULT_ANN_ASSUM, DEFAULT_EV_AGG;










function b64ToBytes(b){const s=atob(b),a=new Uint8Array(s.length);for(let i=0;i<s.length;i++)a[i]=s.charCodeAt(i);return a;}
const KEY_VARS=new Set(["EarnedPrem","ReinsPrem","IncClaims","ReinsClaims","TabRes","CededALRstat","CLRes","TS","Comm","PremTax","LivesIssued","AEGAdminPolCount"]);
const S={py:null,files:{ev:null,rbc:null,bp:null},evData:null,out:null,assumptions:null,
  scenarios:JSON.parse(localStorage.getItem('bb_sc')||'[]'),
  auditLog:JSON.parse(localStorage.getItem('bb_al')||'[]'),
  testCases:JSON.parse(localStorage.getItem('bb_tc')||'null'),
  testResults:[],runId:null};

async function initPy(){
  const st=document.getElementById('pyst'),bar=document.getElementById('ldb'),msg=document.getElementById('ldm');
  try{
    // --- load engine + default data (decomposed sources) ---
    const _f=async(p,t)=>{const r=await fetch(p);if(!r.ok)throw new Error('Failed to load '+p+' ('+r.status+')');return t==='json'?r.json():t==='buf'?new Uint8Array(await r.arrayBuffer()):r.text();};
    PY_ENGINE=await _f('../src/engine.py?v='+Date.now(),'text');
    DEFAULT_EV_CSV=await _f('../data/EV_Data_Final.csv','text');
    DOC_TEXT=await _f('../data/doc.txt','text');
    DEFAULT_SURPLUS_ROWS=await _f('../data/surplus.json','json');
    DEFAULT_ANN_ASSUM=await _f('../data/ann_assum.json','json');
    DEFAULT_EV_AGG=await _f('../data/ev_agg_empty.json','json');
    S.files.bp=await _f('../data/balanceplan.xlsx','buf');

    msg.textContent='Loading Python runtime...';bar.style.width='10%';
    S.py=await loadPyodide({indexURL:"https://cdn.jsdelivr.net/pyodide/v0.25.1/full/"});
    msg.textContent='Installing openpyxl...';bar.style.width='22%';
    await S.py.loadPackage('micropip');
    const mp=S.py.pyimport('micropip');
    await mp.install('openpyxl');
    msg.textContent='Loading actuarial engine...';bar.style.width='45%';
    await S.py.runPythonAsync(PY_ENGINE);
    msg.textContent='Loading default files...';bar.style.width='65%';
    // balance plan bytes loaded via fetch in the loader block above
    // Default: use DEFAULT_SURPLUS_ROWS (correct MS+All-product structure)
    // _rbc_rows is intentionally NOT set so run_model falls back to _surplus_rows
    S.py.globals.set('_sjson',JSON.stringify(DEFAULT_SURPLUS_ROWS));
    await S.py.runPythonAsync(
      "import json as _j\n"+
      "_rbc_rows=None\n"+
      "_surplus_rows=_j.loads(_sjson)"
    );
    S.py.globals.set('_tmpb',S.files.bp);
    await S.py.runPythonAsync(
      "import io,openpyxl as _x\n"+
      "_wb=_x.load_workbook(io.BytesIO(bytes(_tmpb)),data_only=True,read_only=True)\n"+
      "_bp_rows=[list(r) for r in _wb.worksheets[0].iter_rows(values_only=True)]\n"+
      "_wb.close()"
    );
    msg.textContent='Ready — upload EV_Data_template.csv to run model';bar.style.width='82%';
    try{
      S.evData=await parseEVwithCSV(new TextEncoder().encode(DEFAULT_EV_CSV));
      const evSt=document.getElementById('st-ev');
      evSt.style.display='inline-block';evSt.className='upstat ok';
      evSt.textContent='Default EV loaded: '+S.evData.row_count.toLocaleString()+' rows, '+S.evData.iss_years.length+' issue years (upload to override)';
    }catch(_e){
      S.evData=prepEVAgg(DEFAULT_EV_AGG);
      const evSt=document.getElementById('st-ev');
      evSt.style.display='inline-block';evSt.className='upstat err';
      evSt.textContent='Default EV parse error: '+_e.message;
    }
    bar.style.width='100%';
    await new Promise(r=>setTimeout(r,300));
    st.textContent='Ready';st.className='ready';
    document.getElementById('runbtn').disabled=false;
    document.getElementById('runbtn2').disabled=false;
    document.getElementById('loading').classList.add('hidden');
    initAssumptions();renderScenarios();
    // Always start with fresh defaults, append any custom (non-default) cases
    const _defIds=new Set(getDefaultTestCases().map(function(t){return t.id;}));
    const _custom=(S.testCases||[]).filter(function(t){return !_defIds.has(t.id);});
    // Normalize expense tests: expected_change should be negative (cost increases = more negative)
    var _all=getDefaultTestCases().concat(_custom);
    _all.forEach(function(t){
      if(t.type==='expense'&&t.expected_change>0)t.expected_change=-t.expected_change;
    });
    S.testCases=_all;
    renderTestCases();
  }catch(e){st.textContent='Error';st.className='error';msg.textContent='Error: '+e.message;console.error(e);}
}

function prepEVAgg(raw){
  const r=Object.assign({},raw);
  r.agg={};Object.entries(raw.agg).forEach(([vn,pv])=>{r.agg[vn]={};Object.entries(pv).forEach(([p,v])=>r.agg[vn][parseInt(p)]=v);});
  r.agg_iy={};Object.entries(raw.agg_iy).forEach(([iy,vm])=>{r.agg_iy[iy]={};Object.entries(vm).forEach(([vn,pv])=>{r.agg_iy[iy][vn]={};Object.entries(pv).forEach(([p,v])=>r.agg_iy[iy][vn][parseInt(p)]=v);});});
  r.periods=raw.periods.map(Number);
  return r;
}

function showTab(n){
  document.querySelectorAll('.panel').forEach(p=>p.classList.remove('active'));
  document.querySelectorAll('.nt').forEach(t=>t.classList.remove('active'));
  document.getElementById('panel-'+n).classList.add('active');
  document.getElementById('tab-'+n).classList.add('active');
  if(n==='docs')renderDocs();
}
function showSub(n){
  ['summary','stmt','rbc','data','validation','frontier'].forEach(x=>{
    const sp=document.getElementById('sec-'+x);if(sp)sp.style.display=x===n?'':'none';
  });
  document.querySelectorAll('#main-stabs .stab').forEach(s=>s.classList.remove('active'));
  const btn=document.getElementById('stab-'+n);if(btn)btn.classList.add('active');
  if(n==='summary')renderSummary();
  else if(n==='rbc'&&S.out)renderRBC(S.out.rbc_data,S.out.rbc_net,S.out);
  else if(n==='data')renderEV('d');
  else if(n==='frontier')renderFrontierUI();
}
// Secondary view toggle within a Results section
function subView(section,name){
  const groups={stmt:['predeal','ceded','net'],data:['ev','evc'],
    validation:['balance','issyr','testlog','checklist']};
  (groups[section]||[]).forEach(x=>{const el=document.getElementById('sp-'+x);if(el)el.style.display=x===name?'':'none';});
  const bar=document.getElementById('vtog-'+section);
  if(bar)bar.querySelectorAll('button').forEach(b=>b.classList.remove('active'));
  const vb=document.getElementById('vb-'+section+'-'+name);if(vb)vb.classList.add('active');
  if(name==='ev')renderEV('d');
  else if(name==='evc')renderEV('c');
  else if(name==='issyr')renderIYDiag();
}

async function handleUpload(evt,key){
  const file=evt.target.files[0];if(!file)return;
  const st=document.getElementById('st-'+key);
  st.style.display='inline-block';st.className='upstat';st.textContent='Parsing...';
  try{
    const bytes=new Uint8Array(await file.arrayBuffer());S.files[key]=bytes;
    if(key==='ev'){
      const isCSV=file.name.toLowerCase().endsWith('.csv');
      const evData=isCSV?await parseEVwithCSV(bytes):await parseEVwithSheetJS(bytes);
      S.evData=evData;
      document.getElementById('runbtn').disabled=false;document.getElementById('runbtn2').disabled=false;
      st.className='upstat ok';st.textContent=file.name+' ('+evData.row_count.toLocaleString()+' rows)';
    }else{
      S.py.globals.set('_tmpb',bytes);
      if(key==='rbc'){
        S.py.globals.set('_sjson',JSON.stringify(DEFAULT_SURPLUS_ROWS));
        await S.py.runPythonAsync(
          "import io,json as _j,csv as _csv\n"+
          "_data=bytes(_tmpb)\n"+
          "if _data[:2]==b'PK':\n"+
          "    import openpyxl as _x\n"+
          "    _wb=_x.load_workbook(io.BytesIO(_data),data_only=True,read_only=True)\n"+
          "    _rbc_rows=[list(r) for r in _wb.worksheets[0].iter_rows(values_only=True)]\n"+
          "    _sr=([list(r) for r in _wb['Surplus'].iter_rows(values_only=True)] if 'Surplus' in _wb.sheetnames else None)\n"+
          "    _wb.close()\n"+
          "else:\n"+
          "    _txt=_data.decode('utf-8',errors='ignore')\n"+
          "    _rbc_rows=[]\n"+
          "    for _ln in _csv.reader(io.StringIO(_txt)):\n"+
          "        if not _ln: continue\n"+
          "        if _ln[0].startswith('//'): continue\n"+
          "        _row=[]\n"+
          "        for _i,_c in enumerate(_ln):\n"+
          "            _c=_c.strip()\n"+
          "            if _i==0 or not _c: _row.append(_c if _c else None)\n"+
          "            else:\n"+
          "                try: _row.append(float(_c))\n"+
          "                except: _row.append(_c)\n"+
          "        _rbc_rows.append(_row)\n"+
          "    _sr=None\n"+
          "_surplus_rows=(_sr if _sr is not None else _j.loads(_sjson))"
        );
      }else{
        await S.py.runPythonAsync(
          "import io,openpyxl as _x\n"+
          "_wb=_x.load_workbook(io.BytesIO(bytes(_tmpb)),data_only=True,read_only=True)\n"+
          "_bp_rows=[list(r) for r in _wb.worksheets[0].iter_rows(values_only=True)]\n"+
          "_wb.close()"
        );
      }
      st.className='upstat ok';st.textContent=file.name+' ('+Math.round(file.size/1024)+'KB)';
    }
  }catch(e){st.className='upstat err';st.textContent='Error: '+e.message;console.error(e);}
}

async function parseEVwithCSV(bytes){
  return new Promise((resolve,reject)=>{
    try{
      const text=new TextDecoder('utf-8').decode(bytes);
      const lines=text.split(/\r?\n/);
      if(lines.length<2){reject(new Error('CSV empty'));return;}
      // Parse header
      const hdr=lines[0].split(',').map(h=>h.trim().replace(/^"|"$/g,''));
      const vnI=hdr.indexOf('VarName'),iyI=hdr.indexOf('ck.IssYear');
      const imI=hdr.indexOf('ck.IssMon'),nbI=hdr.indexOf('ck.NewBus'),prI=hdr.indexOf('Product');
      if(vnI<0||iyI<0){reject(new Error('Missing VarName/ck.IssYear columns'));return;}
      const valCols=[];
      hdr.forEach((h,i)=>{if(h&&/^Value\d+$/.test(h))valCols.push([i,parseInt(h.slice(5))]);});
      valCols.sort((a,b)=>a[1]-b[1]);
      const agg={},aggIY={},iySet=new Set(),vnSet=new Set(),imSet=new Set(),nbSet=new Set(),prSet=new Set();
      const browseRows=[];let rowCount=0;
      for(let r=1;r<lines.length;r++){
        const line=lines[r].trim();if(!line)continue;
        // Fast CSV split (no quoted commas support needed for numeric data)
        const row=line.split(',');
        let vn=row[vnI]?row[vnI].trim().replace(/^"|"$/g,''):'';if(!vn)continue;
        if(vn==='LivesInForce1')vn='AEGAdminPolCount'; // slimmed EV uses in-force count as the maintenance/admin-policy-count basis
        const _iym=String(row[iyI]).match(/\d{4}/);const iy=_iym?parseInt(_iym[0]):(parseInt(row[iyI])||0); // '<=2019' -> 2019
        const im=parseInt(row[imI])||0;
        const nb=row[nbI]?row[nbI].trim():'';
        const pr=row[prI]?row[prI].trim():'';
        vnSet.add(vn);iySet.add(iy);imSet.add(im);nbSet.add(nb);prSet.add(pr);
        rowCount++;
        if(browseRows.length<2000){
          const vals={};
          valCols.forEach(([ci,p])=>{const v=parseFloat(row[ci]);if(!isNaN(v)&&v!==0)vals[p]=v;});
          browseRows.push({iy,im,nb,pr,vn,vals});
        }
        if(!KEY_VARS.has(vn))continue;
        if(!agg[vn])agg[vn]={};
        const iys=String(iy);
        if(!aggIY[iys])aggIY[iys]={};if(!aggIY[iys][vn])aggIY[iys][vn]={};
        valCols.forEach(([ci,p])=>{
          const fv=parseFloat(row[ci]);if(!fv||isNaN(fv))return;
          agg[vn][p]=(agg[vn][p]||0)+fv;
          aggIY[iys][vn][p]=(aggIY[iys][vn][p]||0)+fv;
        });
      }
      resolve({agg,agg_iy:aggIY,periods:valCols.map(x=>x[1]),
        max_period:valCols.length?valCols[valCols.length-1][1]:0,
        row_count:rowCount,iss_years:Array.from(iySet).sort((a,b)=>a-b),
        all_varnames:Array.from(vnSet).sort(),all_iys:Array.from(iySet).sort((a,b)=>a-b),
        all_ims:Array.from(imSet).sort((a,b)=>a-b),all_nbs:Array.from(nbSet).sort(),
        all_prs:Array.from(prSet).sort(),browse_rows:browseRows});
    }catch(e){reject(e);}
  });
}

async function parseEVwithSheetJS(bytes){
  return new Promise((resolve,reject)=>{
    try{
      // Try parsing - handle both normal and very large files
      let wb;
      try{
        wb=XLSX.read(bytes,{type:'array',cellDates:false,cellHTML:false,cellFormula:false,cellNF:false,sheetRows:0});
      }catch(e2){
        wb=XLSX.read(bytes,{type:'array'});
      }
      if(!wb||!wb.SheetNames||!wb.SheetNames.length)throw new Error('Could not read workbook');
      const ws=wb.Sheets[wb.SheetNames[0]];
      if(!ws)throw new Error('Sheet not found in workbook');
      // Get range
      const ref=ws['!ref'];
      if(!ref)throw new Error('Empty sheet - no cell reference found');
      const range=XLSX.utils.decode_range(ref);
      const nCols=range.e.c-range.s.c+1;
      // Build raw data: header + data rows using cell-by-cell access
      const raw=[];
      const hdrRow=[];
      for(let c=range.s.c;c<=range.e.c;c++){
        const cell=ws[XLSX.utils.encode_cell({r:range.s.r,c})];
        hdrRow.push(cell&&cell.v!=null?cell.v:null);
      }
      raw.push(hdrRow);
      for(let r=range.s.r+1;r<=range.e.r;r++){
        const row=new Array(nCols).fill(null);
        let hasData=false;
        for(let c=range.s.c;c<=range.e.c;c++){
          const cell=ws[XLSX.utils.encode_cell({r,c})];
          if(cell!=null&&cell.v!=null){row[c-range.s.c]=cell.v;hasData=true;}
        }
        if(hasData)raw.push(row);
      }
      if(!raw||raw.length<2){reject(new Error('Empty sheet'));return;}
      const hdr=raw[0].map(h=>h!=null?String(h):null);
      const vnI=hdr.indexOf('VarName'),iyI=hdr.indexOf('ck.IssYear'),imI=hdr.indexOf('ck.IssMon'),nbI=hdr.indexOf('ck.NewBus'),prI=hdr.indexOf('Product');
      if(vnI<0||iyI<0){reject(new Error('Missing VarName/ck.IssYear'));return;}
      const valCols=[];hdr.forEach((h,i)=>{if(h&&/^Value\d+$/.test(h))valCols.push([i,parseInt(h.slice(5))]);});
      valCols.sort((a,b)=>a[1]-b[1]);
      const agg={},aggIY={},iySet=new Set(),vnSet=new Set(),imSet=new Set(),nbSet=new Set(),prSet=new Set();
      const browseRows=[];let rowCount=0;
      for(let r=1;r<raw.length;r++){
        const row=raw[r];if(!row||!row[vnI])continue;
        let vn=String(row[vnI]);if(vn==='LivesInForce1')vn='AEGAdminPolCount';vnSet.add(vn);
        const _m=String(row[iyI]).match(/\d{4}/);const iy=_m?parseInt(_m[0]):(parseInt(row[iyI])||0),im=parseInt(row[imI])||0;
        const nb=row[nbI]!=null?String(row[nbI]):'',pr=row[prI]!=null?String(row[prI]):'';
        iySet.add(iy);imSet.add(im);nbSet.add(nb);prSet.add(pr);rowCount++;
        if(browseRows.length<2000){const vals={};valCols.forEach(([ci,p])=>{const v=row[ci];if(v!=null&&v!==0)vals[p]=v;});browseRows.push({iy,im,nb,pr,vn,vals});}
        if(!KEY_VARS.has(vn))continue;
        if(!agg[vn])agg[vn]={};const iys=String(iy);
        if(!aggIY[iys])aggIY[iys]={};if(!aggIY[iys][vn])aggIY[iys][vn]={};
        valCols.forEach(([ci,p])=>{const fv=parseFloat(row[ci]);if(!fv)return;agg[vn][p]=(agg[vn][p]||0)+fv;aggIY[iys][vn][p]=(aggIY[iys][vn][p]||0)+fv;});
      }
      resolve({agg,agg_iy:aggIY,periods:valCols.map(x=>x[1]),max_period:valCols.length?valCols[valCols.length-1][1]:0,row_count:rowCount,iss_years:Array.from(iySet).sort((a,b)=>a-b),all_varnames:Array.from(vnSet).sort(),all_iys:Array.from(iySet).sort((a,b)=>a-b),all_ims:Array.from(imSet).sort((a,b)=>a-b),all_nbs:Array.from(nbSet).sort(),all_prs:Array.from(prSet).sort(),browse_rows:browseRows});
    }catch(e){reject(e);}
  });
}

// -- ASSUMPTIONS --
function defAssumptions(){
  const rp={};
  for(let iy=2019;iy<=2030;iy++){rp[iy]={};for(let cal=Math.max(2026,iy+1);cal<=2031;cal++)rp[iy][cal]=10;}
  return{premium_tax:0.021,discount_rate:0.08,cost_of_capital:0.10,ceding_comm_ongoing:200,base_year:2025,
    claim_scalar:1.0,lapse_scalar:1.0,
    ceding_comm_front:{2026:10,2027:5,2028:5},
    annual:JSON.parse(JSON.stringify(DEFAULT_ANN_ASSUM)),reins_pct:rp};
}
function initAssumptions(){if(!S.assumptions)S.assumptions=defAssumptions();renderAssumUI();}
function resetAssumptions(){if(!confirm('Reset all assumptions to defaults?'))return;S.assumptions=defAssumptions();renderAssumUI();}

function renderAssumUI(){
  const a=S.assumptions;
  document.getElementById('a-ptax').value=a.premium_tax||0.021;
  document.getElementById('a-disc').value=a.discount_rate||0.08;
  {const _coc=document.getElementById('a-coc');if(_coc)_coc.value=a.cost_of_capital!=null?a.cost_of_capital:0.10;}
  initCCOTable(a.ceding_comm_table||[[0,0.75,250],[0.75,0.85,200],[0.85,0.95,150],[0.95,'Inf',100]]);
  document.getElementById('a-base').value=a.base_year||2025;
  document.getElementById('a-clm').value=a.claim_scalar||1.0;
  document.getElementById('a-lps').value=a.lapse_scalar||1.0;
  renderCCFront(a.ceding_comm_front||{});renderAnnual(a.annual||{});renderReinsMatrix(a.reins_pct||{});
  // Populate sales scalar inputs
  const ss=a.sales_scalar||{};
  [2026,2027,2028,2029,2030].forEach(function(yr){
    const el=document.getElementById('a-ss-'+yr);
    if(el)el.value=(ss[yr]||1.0).toFixed(2);
  });
}

function renderCCFront(cf){
  document.getElementById('ccfront-tbody').innerHTML=Object.entries(cf).sort(([a],[b])=>a-b).map(([y,v])=>
    '<tr><td class="rh">'+y+'</td><td><input type="number" value="'+v+'" step="1" style="width:110px" onchange="S.assumptions.ceding_comm_front['+y+']=parseFloat(this.value)||0"></td><td><button style="color:var(--err);background:none;border:none;cursor:pointer" onclick="delete S.assumptions.ceding_comm_front['+y+'];renderCCFront(S.assumptions.ceding_comm_front)">x</button></td></tr>'
  ).join('');
}
function addCCFrontRow(){const a=S.assumptions;const yrs=Object.keys(a.ceding_comm_front||{}).map(Number);const ny=yrs.length?Math.max(...yrs)+1:2026;if(!a.ceding_comm_front)a.ceding_comm_front={};a.ceding_comm_front[ny]=0;renderCCFront(a.ceding_comm_front);}

const ANN_ROWS=[{k:'nier',l:'NIER (rate)',step:'0.000001',dec:6},{k:'acq_exp',l:'Acq Exp ($/policy)',step:'0.01',dec:2},{k:'maint_exp',l:'Maint Exp ($/pol/yr)',step:'0.01',dec:2},{k:'acq_exp_allowance',l:'Acq Allowance ($/policy)',step:'0.01',dec:2},{k:'maint_exp_allowance',l:'Maint Allowance ($/pol/yr)',step:'0.01',dec:2}];
const ANN_YEARS=[2026,2027,2028,2029,2030,2031,2032,2033,2034,2035];

function renderAnnual(ann){
  document.getElementById('ann-hd').innerHTML='<tr><th style="text-align:left">Assumption</th>'+ANN_YEARS.map(y=>'<th>'+y+'</th>').join('')+'</tr>';
  document.getElementById('ann-bd').innerHTML=ANN_ROWS.map(r=>'<tr><td class="rh">'+r.l+'</td>'+ANN_YEARS.map(y=>{const v=(ann[y]||{})[r.k];return '<td><input type="number" value="'+(v!=null?Number(v).toFixed(r.dec):'')+'" step="'+r.step+'" style="width:78px" onchange="updAnn('+y+',\''+r.k+'\',this.value)"></td>';}).join('')+'</tr>').join('');
}
function updAnn(y,k,v){if(!S.assumptions.annual)S.assumptions.annual={};if(!S.assumptions.annual[y])S.assumptions.annual[y]={};S.assumptions.annual[y][k]=parseFloat(v)||0;}

function handleAnnPaste(e){
  e.preventDefault();
  const text=(e.clipboardData||window.clipboardData).getData('text');
  const rows=text.trim().split(/\r?\n/).map(r=>r.split(/\t/));
  if(!rows.length)return;
  const activeCell=document.activeElement;if(!activeCell||!activeCell.closest('#ann-tbl'))return;
  const tdEl=activeCell.closest('td'),trEl=activeCell.closest('tr');if(!tdEl||!trEl)return;
  const tBody=document.getElementById('ann-bd');
  const allTrs=Array.from(tBody.querySelectorAll('tr'));const startRow=allTrs.indexOf(trEl);
  const allTds=Array.from(trEl.querySelectorAll('td'));const startCol=allTds.indexOf(tdEl)-1;
  if(startRow<0||startCol<0)return;
  rows.forEach((row,ri)=>{
    const tr=allTrs[startRow+ri];if(!tr)return;
    const tds=Array.from(tr.querySelectorAll('td'));const rowKey=ANN_ROWS[startRow+ri]?.k;
    row.forEach((val,ci)=>{
      const tdIdx=startCol+ci+1;if(!tds[tdIdx])return;
      const inp=tds[tdIdx].querySelector('input');if(!inp)return;
      const num=parseFloat(val.replace(/,/g,'').replace(/\$/g,''))||0;
      inp.value=num;const yr=ANN_YEARS[startCol+ci];if(yr&&rowKey)updAnn(yr,rowKey,num);
    });
  });
}

function renderReinsMatrix(rp){
  const iys=Object.keys(rp).map(Number).sort((a,b)=>a-b);
  let cys=[...new Set(Object.values(rp).flatMap(o=>Object.keys(o).map(Number)))].sort((a,b)=>a-b);
  if(!iys.length)for(let y=2019;y<=2030;y++)iys.push(y);
  if(!cys.length)for(let y=2026;y<=2031;y++)cys.push(y);
  const rowLabels=iys.map(iy=>iy<=2019?'2019 & Prior':String(iy));
  document.getElementById('reins-hd').innerHTML='<tr><th>IssYr</th>'+cys.map(cy=>'<th>'+cy+'</th>').join('')+'<th style="font-size:.55rem;color:rgba(255,215,0,.5)">→ extends</th></tr>';
  document.getElementById('reins-bd').innerHTML=iys.map((iy,ri)=>
    '<tr><td class="rh">'+rowLabels[ri]+'</td>'+cys.map(cy=>{
      if(iy>cy&&iy>2019)return '<td class="nc"></td>';
      if(iy===2019&&cy<2026)return '<td class="nc"></td>';
      const v=(rp[iy]||{})[cy];
      return '<td><input type="number" value="'+(v!=null?v:'')+'" step="0.1" min="0" max="100" placeholder="0" style="width:55px" onchange="updReins('+iy+','+cy+',this.value)"></td>';
    }).join('')+'<td style="font-size:.6rem;color:var(--mu);padding-left:4px;background:var(--off)">'+(()=>{const lv=cys[cys.length-1];const r=(rp[iy]||{})[lv];return r?r+'%→∞':'';})()+'</td></tr>'
  ).join('');
}
function updReins(iy,cy,v){if(!S.assumptions.reins_pct)S.assumptions.reins_pct={};if(!S.assumptions.reins_pct[iy])S.assumptions.reins_pct[iy]={};S.assumptions.reins_pct[iy][cy]=v===''?null:parseFloat(v);}
function addReinsRow(){const rp=S.assumptions.reins_pct||{};const iys=Object.keys(rp).map(Number).filter(y=>y<=2030);const ny=iys.length?Math.max(...iys)+1:2020;if(ny>2030){alert('No rows for 2031+ issue years');return;}rp[ny]={};S.assumptions.reins_pct=rp;renderReinsMatrix(rp);}
function addReinsCol(){const rp=S.assumptions.reins_pct||{};const cys=[...new Set(Object.values(rp).flatMap(o=>Object.keys(o).map(Number)))].sort((a,b)=>a-b);const nc=cys.length?Math.max(...cys)+1:2032;Object.keys(rp).forEach(iy=>{if(Number(iy)<nc){const lv=cys[cys.length-1];rp[iy][nc]=rp[iy][lv]||10;}});renderReinsMatrix(rp);}

function collectAssumptions(){
  const a=S.assumptions||{};
  a.premium_tax=parseFloat(document.getElementById('a-ptax').value)||0.021;
  a.discount_rate=parseFloat(document.getElementById('a-disc').value)||0.08;
  {const _coc=document.getElementById('a-coc');a.cost_of_capital=_coc?(parseFloat(_coc.value)||0.10):0.10;}
  a.ceding_comm_table=readCCOTable();
  a.base_year=parseInt(document.getElementById('a-base').value)||2025;
  a.claim_scalar=parseFloat(document.getElementById('a-clm').value)||1.0;
  a.lapse_scalar=parseFloat(document.getElementById('a-lps').value)||1.0;
  const ann=a.annual||{};
  ['nier','acq_exp','maint_exp','acq_exp_allowance','maint_exp_allowance'].forEach(k=>{a[k]={};Object.entries(ann).forEach(([y,obj])=>{if(obj&&obj[k]!=null)a[k][parseInt(y)]=obj[k];});});
  a.reins_pct_decimal={};Object.entries(a.reins_pct||{}).forEach(([iy,yrs])=>{a.reins_pct_decimal[parseInt(iy)]={};Object.entries(yrs).forEach(([cy,v])=>{a.reins_pct_decimal[parseInt(iy)][parseInt(cy)]=v!=null?v/100:null;});});
  a.ceding_comm_front_dollars={};Object.entries(a.ceding_comm_front||{}).forEach(([y,v])=>{a.ceding_comm_front_dollars[parseInt(y)]=(v||0)*1_000_000;});
  // Sales scalar by issue year
  a.sales_scalar={};
  [2026,2027,2028,2029,2030].forEach(function(yr){
    const el=document.getElementById('a-ss-'+yr);
    if(el){const v=parseFloat(el.value)||1.0;if(v!==1.0)a.sales_scalar[yr]=v;}
  });
  S.assumptions=a;return a;
}

// -- MODEL RUN --
async function runModel(){
  if(!S.evData||!S.evData.agg||!Object.keys(S.evData.agg).length){
    alert('Please upload EV_Data_template.csv on the Upload tab before running the model.');
    return;
  }
  if(!S.py||!S.evData){alert('Load EV data first.');return;}
  const prog=document.getElementById('runprog'),bar=document.getElementById('runbar'),msg=document.getElementById('runmsg'),stat=document.getElementById('runstat');
  prog.style.display='';
  try{
    collectAssumptions();
    const a=S.assumptions,by=a.base_year||2025;
    const apython={...a,reins_pct:a.reins_pct_decimal||{},ceding_comm_front:a.ceding_comm_front_dollars||{}};
    const ev_for_py={agg:S.evData.agg,agg_iy:S.evData.agg_iy,periods:S.evData.periods,iss_years:S.evData.iss_years,row_count:S.evData.row_count};
    msg.textContent='Running engine...';bar.style.width='25%';
    S.py.globals.set('_evj',JSON.stringify(ev_for_py));S.py.globals.set('_asj',JSON.stringify(apython));S.py.globals.set('_by',by);
    const runCode=["import json","ev_agg=json.loads(_evj)","assum=json.loads(_asj)","for k in ['acq_exp','maint_exp','nier','ceding_comm_front','acq_exp_allowance','maint_exp_allowance']:","    if k in assum and isinstance(assum[k],dict):","        assum[k]={int(kk):vv for kk,vv in assum[k].items()}","if 'reins_pct' in assum:","    assum['reins_pct']={int(iy):{int(cy):v for cy,v in yd.items()} for iy,yd in assum['reins_pct'].items()}","ev_agg['agg']={k:{int(p):v for p,v in pv.items()} for k,pv in ev_agg['agg'].items()}","ev_agg['agg_iy']={str(iy):{k:{int(p):v for p,v in pv.items()} for k,pv in vm.items()} for iy,vm in ev_agg['agg_iy'].items()}","ev_agg['periods']=[int(p) for p in ev_agg['periods']]","for _cco in assum.get('ceding_comm_table',[]):\n    _cco[1]=float('inf') if str(_cco[1]) in ('Inf','inf','INF') else float(_cco[1])","result=run_model(ev_agg,assum,int(_by),_rbc_rows if '_rbc_rows' in dir() else None,_bp_rows if '_bp_rows' in dir() else None,_surplus_rows if '_surplus_rows' in dir() else None)","json.dumps(result)"].join("\n");
    const rj=await S.py.runPythonAsync(runCode);
    msg.textContent='Rendering...';bar.style.width='85%';
    const out=JSON.parse(rj);if(out.error)throw new Error(out.error);
    S.out=out;S.runId='run_'+Date.now();
    populateFilters(S.evData);renderResults(out);updateReview(out);
    bar.style.width='100%';stat.textContent='Done - '+new Date().toLocaleTimeString();
    const iysel=document.getElementById('iy-select');
    if(iysel)iysel.innerHTML='<option value="">select</option>'+(out.iss_years||[]).map(y=>'<option>'+y+'</option>').join('');
    showTab('results');
    setTimeout(()=>{prog.style.display='none';bar.style.width='0%';},1500);
  }catch(e){bar.style.width='0%';msg.textContent='';stat.textContent='Error: '+e.message;console.error(e);alert('Error: '+e.message);}
}

// -- FORMATTERS --
function fmtM(v){if(v==null||isNaN(v))return'-';const n=Number(v)/1e6;return n<0?'('+Math.abs(n).toFixed(2)+')':n.toFixed(2);}
function fmtPct(v){if(v==null||isNaN(v))return'-';return(Number(v)*100).toFixed(2)+'%';}
function fmtNum(v){if(v==null)return'-';const n=Number(v)/1e6;if(isNaN(n))return'-';const cls=n<-0.0001?'neg':n>0.0001?'pos':'';const s=n<0?'('+Math.abs(n).toFixed(2)+')':n.toFixed(2);return cls?'<span class="'+cls+'">'+s+'</span>':s;}
function fmtRaw(v){if(v==null)return'-';const n=Number(v);if(isNaN(n))return'-';if(n>=1e6)return(n/1e6).toFixed(1)+'M';if(n>=1e3)return(n/1e3).toFixed(0)+'K';return n.toFixed(0);}
function fmtDol(v){if(v==null)return'-';const n=Number(v);if(isNaN(n))return'-';const cls=n<-0.01?'neg':n>0.01?'pos':'';const fmt=Math.abs(n)>=1e6?(Math.abs(n)/1e6).toFixed(2)+'M':Math.abs(n)>=1e3?(Math.abs(n)/1e3).toFixed(0)+'K':Math.abs(n).toFixed(0);const s=n<0?'('+fmt+')':fmt;return cls?'<span class="'+cls+'">'+s+'</span>':s;}
function gv(obj,key){if(!obj)return null;const v=obj[key];if(v!==undefined)return v;const sv=obj[String(key)];return sv!==undefined?sv:null;}

// -- INCOME STATEMENT LINES --
const LINES=[
  {k:'premium',l:'Premium'},
  {k:'nii',l:'Net Investment Income'},
  {k:'comm1',l:'Ceding Comm (Front-End)',netRevOnly:true},
  {k:'comm2',l:'Ceding Comm (Ongoing)',netRevOnly:true},
  {k:'revenue',l:'Total Revenue',cls:'sub'},
  {k:'claims',l:'Claims'},
  {k:'delta_reserves',l:'Change in Reserves'},
  {k:'benefits',l:'Total Benefits',cls:'sub'},
  {k:'comm1',l:'Ceding Comm (Front-End)',cededOnly:true},
  {k:'comm2',l:'Ceding Comm (Ongoing)',cededOnly:true},
  {k:'commissions',l:'Commissions'},
  {k:'prem_tax',l:'Premium Tax'},
  {k:'selling_expense',l:'Selling Expense',cls:'sub'},
  {k:'acq_expense',l:'Acquisition Expense'},
  {k:'maint_expense',l:'Maintenance Expense'},
  {k:'op_expense',l:'Operating Expense',cls:'sub'},
  {k:'pretax_income',l:'Pre-Tax Income',cls:'tot'},
  {k:'distributable_earnings',l:'Distributable Earnings',cls:'de'},
  {k:'rbc_ratio_display',l:'RBC Ratio (w/ Margin)',cls:'ratio',noFmt:true},
  {k:'admin_pol_count',l:'Policies In Force',cls:'hdr',raw:true},
  {k:'lives_issued',l:'Policies Issued',cls:'hdr',raw:true},
  {k:'policy_reserve',l:'Policy Reserve (Net)',cls:'hdr'},
  {k:'claim_reserve',l:'Claim Reserve',cls:'hdr'},
  {k:'target_surplus',l:'Target Surplus',cls:'hdr'},
  {k:'total_assets',l:'Total Assets',cls:'sub'},
];

// ===== CEDANT SCENARIO MATRIX (combinatoric sweep) =====
function mtxFmt(v,d){return (v==null||isNaN(v))?'-':Number(v).toFixed(d==null?1:d);}
function exportMatrixCSV(){
  if(!S.matrix||!S.matrix.length){alert('Run a sweep first.');return;}
  const envName=e=>'c'+e.claim+'_l'+e.lapse;
  const head=['scenario_id','cc_2026','cc_2027','cc_2028','lr_split','reins_pre2019','reins_2021_24','reins_2025','reins_2026p',
    'env_claim','env_lapse','net_deal_value','cap_relief_value','reinsurer_pvde','pred_pvde','net_pvde','value_recovery_pct',
    'cum_net_pti_3yr','cum_pred_pti_3yr','back_pred_irr','back_net_irr','new_pred_irr','new_net_irr','downside_protection_3yr'];
  const lines=[head.join(',')];
  S.matrix.forEach(s=>{(s.stress||[]).forEach(e=>{
    const bk=e.back||{},nw=e.new||{};
    lines.push([s.scenario_id,s.cc_2026,s.cc_2027,s.cc_2028,s.lr_split,s.reins_pre2019,s.reins_2021_24,s.reins_2025,s.reins_2026p,
      e.claim,e.lapse,e.net_deal_value,e.cap_relief_value,e.reinsurer_pvde,e.pred_pvde,e.net_pvde,e.value_recovery_pct,
      e.cum_net_pti_3yr,e.cum_pred_pti_3yr,bk.predeal_irr,bk.net_irr,nw.predeal_irr,nw.net_irr,s.downside_protection_3yr]
      .map(v=>v==null?'':v).join(','));
  });});
  const blob=new Blob([lines.join('\n')],{type:'text/csv'});const url=URL.createObjectURL(blob);
  const aa=document.createElement('a');aa.href=url;aa.download='scenario_matrix.csv';aa.click();URL.revokeObjectURL(url);
}
function renderStmt(tid,stmt,cols,view){
  // view: 'predeal'|'ceded'|'net'
  const isNet=view==='net';
  const isCeded=view==='ceded';
  const t=document.getElementById(tid);if(!t||!stmt||!cols||!cols.length)return;
  let h='<thead><tr><th style="min-width:160px">Line Item ($M)</th>'+cols.map(p=>'<th style="min-width:75px">'+p+'</th>').join('')+'</tr></thead><tbody>';
  for(const li of LINES){
    if(li.netRevOnly&&view!=='net')continue;
    if(li.cededOnly&&view!=='ceded')continue;
    if(!li.netRevOnly&&!li.cededOnly&&(li.k==='comm1'||li.k==='comm2')&&view==='predeal')continue;
    const d=stmt[li.k]||{};
    const cls=li.cls?' class="'+li.cls+'"':'';
    h+='<tr'+cls+'><td>'+li.l+'</td>'+cols.map(p=>{
      const v=gv(d,p);
        if(li.k==='rbc_ratio_display'){
          // Pull RBC ratio from formulaic compute results (rbc_predeal_result/rbc_net_result/rbc_orig_computed)
          var rbcAdj={};
          if(view==='net'){rbcAdj=(S.out&&S.out.rbc_net_result&&S.out.rbc_net_result.net_adjustments)||{};}
          else{rbcAdj=(S.out&&S.out.rbc_predeal_result&&S.out.rbc_predeal_result.predeal_adjustments)||{};}
          var origVals=(S.out&&S.out.rbc_orig_computed&&S.out.rbc_orig_computed.original_values)||{};
          var aRow=rbcAdj[p]||rbcAdj[String(p)];
          var rbcV=aRow?aRow.ratio_w_margin:null;
          if(rbcV==null){var oRow=origVals[p]||origVals[String(p)];if(oRow)rbcV=oRow.ratio_w_margin;}
          return '<td>'+(rbcV!=null?Number(rbcV).toFixed(2)+'x':'-')+'</td>';
        }
        return '<td>'+(li.raw?fmtRaw(v):fmtNum(v))+'</td>';
    }).join('')+'</tr>';
  }
  const prem=stmt.premium||{},clm=stmt.claims||{},rev=stmt.revenue||{},sell=stmt.selling_expense||{},opx=stmt.op_expense||{};
  h+='<tr class="ratio"><td>Loss Ratio (Claims/Premium)</td>'+cols.map(p=>{const pm=gv(prem,p),cl=gv(clm,p);return '<td>'+(pm&&Math.abs(pm)>0?(Math.abs(cl||0)/Math.abs(pm)*100).toFixed(1)+'%':'-')+'</td>';}).join('')+'</tr>';
  h+='<tr class="ratio"><td>Expense Ratio ((Sell+Op)/Rev)</td>'+cols.map(p=>{const rv=gv(rev,p),sl=gv(sell,p),op=gv(opx,p);return '<td>'+(rv&&Math.abs(rv)>0?((Math.abs(sl||0)+Math.abs(op||0))/Math.abs(rv)*100).toFixed(1)+'%':'-')+'</td>';}).join('')+'</tr>';
  h+='<tr class="ratio"><td>Combined Ratio</td>'+cols.map(p=>{const rv=gv(rev,p),cl=gv(clm,p),sl=gv(sell,p),op=gv(opx,p);return '<td>'+(rv&&Math.abs(rv)>0?((Math.abs(cl||0)+Math.abs(sl||0)+Math.abs(op||0))/Math.abs(rv)*100).toFixed(1)+'%':'-')+'</td>';}).join('')+'</tr>';
  h+='</tbody>';t.innerHTML=h;
}

function setView(v,mode,btn){
  btn.closest('.vtog').querySelectorAll('button').forEach(b=>b.classList.remove('active'));btn.classList.add('active');
  if(!S.out)return;
  const by=S.out.first_yr||2026;
  let cols;
  if(mode==='annual'){cols=(S.out['annual_'+v]?.periods||[]).filter(y=>y>=by).slice(0,30);}
  else{cols=Object.keys(S.out['stmt_'+v]?.premium||{}).map(Number).sort((a,b)=>a-b).filter(p=>p>0).slice(0,60);}
  renderStmt('tbl-'+v,mode==='annual'?S.out['annual_'+v]:S.out['stmt_'+v],cols,v);
}

function renderResults(out){
  document.getElementById('res-empty').style.display='none';
  document.getElementById('res-content').style.display='';
  const firstYr=(out.annual_predeal?.periods||[2026])[0]||2026;S.out.first_yr=firstYr;
  const getPeriods=v=>(out['annual_'+v]?.periods||[]).filter(y=>y>=firstYr).slice(0,30);
  renderStmt('tbl-predeal',out.annual_predeal,getPeriods('predeal'),'predeal');
  renderStmt('tbl-ceded',out.annual_ceded,getPeriods('ceded'),'ceded');
  renderStmt('tbl-net',out.annual_net,getPeriods('net'),'net');
  renderBalancePlan(out,getPeriods('predeal'));renderRBC(out.rbc_data,out.rbc_net,out);
  renderSummary();
  showSub('summary'); // land on the headline view
}

// -- EV BROWSER (5 frozen cols with solid background) --
function populateFilters(evData){
  if(!evData)return;
  ['ev','evc'].forEach(pfx=>{
    const s=(id,opts)=>{const el=document.getElementById(id);if(el)el.innerHTML='<option value="">All</option>'+opts.map(o=>'<option>'+o+'</option>').join('');};
    s(pfx+'-vn',evData.all_varnames||[]);s(pfx+'-iy',evData.all_iys||[]);
  });
}

function renderEV(which){
  const pfx=which==='d'?'ev':'evc';const t=document.getElementById('tbl-'+pfx);if(!t||!S.evData)return;
  const rows=S.evData.browse_rows||[];
  const vf=document.getElementById(pfx+'-vn')?.value||'';const iyf=document.getElementById(pfx+'-iy')?.value||'';
  let filtered=rows;
  if(vf)filtered=filtered.filter(r=>r.vn===vf);if(iyf)filtered=filtered.filter(r=>r.iy===parseInt(iyf));
  const show=filtered.slice(0,200);
  const periodSet=new Set();show.forEach(r=>Object.keys(r.vals).forEach(p=>periodSet.add(parseInt(p))));
  const periods=Array.from(periodSet).sort((a,b)=>a-b).slice(0,25);
  // Frozen col widths: IssYr=45, VarName=140 = 185px total
  const W=[45,140];const lefts=[0,45];
  const bgCol='background:var(--navy)';const bgCell='background:var(--wh)';
  const bgEven='background:var(--off)';
  let h='<thead><tr>'+
    ['IssYr','VarName'].map((l,i)=>`<th style="position:sticky;left:${lefts[i]}px;z-index:5;${bgCol};min-width:${W[i]}px;max-width:${W[i]}px">${l}</th>`).join('')+
    periods.map(p=>'<th>V'+String(p).padStart(3,'0')+'</th>').join('')+'</tr></thead><tbody>';
  for(let ri=0;ri<show.length;ri++){
    const r=show[ri];const bg=ri%2===0?bgCell:bgEven;
    h+='<tr>'+
      [''+r.iy,r.vn].map((v,i)=>`<td style="position:sticky;left:${lefts[i]}px;z-index:2;${bg};min-width:${W[i]}px;max-width:${W[i]}px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap${i===1?';font-weight:600':''}">` +v+'</td>').join('')+
      periods.map(p=>{const v=r.vals[p];return '<td>'+(v!=null?Number(v).toFixed(2):'-')+'</td>';}).join('')+'</tr>';
  }
  h+='</tbody>';
  if(filtered.length>200)h+='<caption style="caption-side:bottom;font-size:.65rem;color:var(--mu);padding:4px">Showing 200 of '+filtered.length+'. Use filters.</caption>';
  t.innerHTML=h;
}

// -- BALANCE PLAN --
function renderBalancePlan(out,years){
  const t=document.getElementById('tbl-balance');if(!t)return;
  const bp=out.bp_data||{},ap=out.annual_predeal||{};
  if(!years||!years.length){t.innerHTML='<thead><tr><th>No data</th></tr></thead>';return;}
  const yrs=years.slice(0,10).map(String);
  const bpMap={Premium:'premium',NII:'nii',Revenue:'revenue',Claims:'claims',Reserves:'delta_reserves',Benefits:'benefits',Commissions:'commissions','Prem Tax':'prem_tax','Selling Expense':'selling_expense',Acquisition:'acq_expense',Maintenance:'maint_expense','Operating Expense':'op_expense','Pretax Income':'pretax_income'};
  const subRows=new Set(['Revenue','Benefits','Selling Expense','Operating Expense','Pretax Income']);
  let h='<thead><tr><th>Line Item</th>'+yrs.map(y=>'<th colspan="3" style="text-align:center">'+y+'</th>').join('')+'</tr><tr><th></th>'+yrs.map(()=>'<th>Plan</th><th>Model</th><th>Var</th>').join('')+'</tr></thead><tbody>';
  for(const[lbl,key] of Object.entries(bpMap)){
    const bpVals=bp[lbl],mVals=ap[key];const isSub=subRows.has(lbl);
    h+='<tr'+(isSub?' class="sub"':'')+'><td>'+lbl+'</td>'+yrs.map(y=>{
      const bv=bpVals?(bpVals[y]||0)*1e6:null,mv=mVals?gv(mVals,y)||0:null;
      const diff=bv!=null&&mv!=null?mv-bv:null;
      return '<td>'+(bv!=null?fmtNum(bv):'-')+'</td><td>'+(mv!=null?fmtNum(mv):'-')+'</td><td>'+(diff!=null?fmtNum(diff):'-')+'</td>';
    }).join('')+'</tr>';
  }
  h+='</tbody>';t.innerHTML=h;
}

// -- RBC --
function showRBCTab(tab,btn){
  ['orig','predeal','net','lift'].forEach(function(t){
    var d=document.getElementById('rbc-'+t+'-content');if(d)d.style.display=t===tab?'':'none';
  });
  var p=btn&&btn.parentNode;if(p)p.querySelectorAll('button').forEach(function(b){b.classList.remove('active');});
  if(btn)btn.classList.add('active');
}

function renderRBC(rbc,rbcNet,out){
  var origFull=out&&out.rbc_orig_full||null;
  var origComputed=out&&out.rbc_orig_computed||null;
  var predealRes=out&&out.rbc_predeal_result||null;
  var netRes=out&&out.rbc_net_result||null;
  var hasOrig=origComputed&&origComputed.original_values&&Object.keys(origComputed.original_values).length;
  
  // ── ORIGINAL TAB ──
  var dOrig=document.getElementById('rbc-orig-content');
  if(dOrig){
    if(!hasOrig){dOrig.innerHTML='<p style="color:var(--mu);padding:8px">No RBC data. Upload an RBC file on the Upload tab.</p>';}
    else{dOrig.innerHTML=buildRBCOrigTable(origComputed,origFull);}
  }
  
  // ── PREDEAL TAB ──
  var dPredeal=document.getElementById('rbc-predeal-content');
  if(dPredeal){
    dPredeal.innerHTML=buildRBCSection(origFull||rbc,'Predeal',out&&out.annual_predeal,predealRes);
  }
  
  // ── NET TAB ──
  var dNet=document.getElementById('rbc-net-content');
  if(dNet){
    dNet.innerHTML=buildRBCSection(origFull||rbcNet||rbc,'Net (Post-Deal)',out&&out.annual_net,netRes);
  }
  
  // ── LIFT TAB ──
  var dLift=document.getElementById('rbc-lift-content');
  if(dLift){dLift.innerHTML=buildRBCLiftTable(rbc,rbcNet,out);}
}


function buildRBCOrigTable(origComputed, origFull){
  if(!origComputed||!origComputed.original_values||!Object.keys(origComputed.original_values).length)
    return '<p style="color:var(--mu);padding:8px">No RBC data. Upload an RBC file on the Upload tab.</p>';
  var values=origComputed.original_values;
  var years=(origComputed.years||[]).filter(function(y){return Number(y)>=2025;}).map(Number).sort(function(a,b){return a-b;});
  
  function v(yr){return values[yr]||values[String(yr)]||{};}
  function f3(x){if(x==null||isNaN(x))return '0.0000';return Number(x).toFixed(4);}
  function fx(x){if(x==null||isNaN(x))return '0.00x';return Number(x).toFixed(2)+'x';}
  function th(){return '<thead><tr><th style="min-width:200px">Line Item</th>'+years.map(function(y){return '<th style="min-width:72px">'+y+'</th>';}).join('')+'</tr></thead>';}
  function rowVals(lbl,fn,cls){
    var vals=years.map(function(y){var x=fn(v(y));return x!=null?x:0;});
    return '<tr'+(cls?' class="'+cls+'"':'')+'><td>'+lbl+'</td>'+vals.map(function(x){return '<td>'+f3(x)+'</td>';}).join('')+'</tr>';
  }
  function rowVx(lbl,fn){
    var vals=years.map(function(y){var x=fn(v(y));return x!=null?x:0;});
    return '<tr><td>'+lbl+'</td>'+vals.map(function(x){return '<td>'+fx(x)+'</td>';}).join('')+'</tr>';
  }
  function sect(title,rows){return '<div class="sect-hdr">'+title+'</div><div class="tw"><table class="bbt" style="font-size:.68rem">'+th()+'<tbody>'+rows+'</tbody></table></div>';}
  function tsc(key){return function(o){var ms=o.ms_tsc||{};return ms[key];};}
  function atsc(key){return function(o){var ap=o.allprod_tsc||{};return ap[key];};}
  
  return '<div class="card"><div class="ch">Original RBC & Surplus</div><div class="cb">'+
    // 1. MS Income (Claims displayed as negative since stored as positive)
    sect('Med Sup Income',
      rowVals('Premium',function(o){return o.ms_prem;})+
      rowVals('Claims',function(o){return o.ms_claims!=null?-o.ms_claims:null;})+
      rowVals('Pre-Tax Income',function(o){return o.ms_income;}))+
    // 2. MS RBC Charges
    sect('Med Sup RBC Charges',
      rowVals('TSC0',tsc('TSC0'))+rowVals('TSC1',tsc('TSC1'))+rowVals('TSLR016',tsc('TSLR016'))+
      rowVals('TSC1CS',tsc('TSC1CS'))+rowVals('TSC2',tsc('TSC2'))+rowVals('TSC3',tsc('TSC3'))+
      rowVals('TSC4a',tsc('TSC4a'))+rowVals('TSC4b',tsc('TSC4b')))+
    // 3. All Product RBC Charges + Total subtotal (calculated)
    sect('All Product RBC Charges',
      rowVals('TSC0',atsc('TSC0'))+rowVals('TSC1',atsc('TSC1'))+rowVals('TSLR016',atsc('TSLR016'))+
      rowVals('TSC1CS',atsc('TSC1CS'))+rowVals('TSC2',atsc('TSC2'))+rowVals('TSC3',atsc('TSC3'))+
      rowVals('TSC4a',atsc('TSC4a'))+rowVals('TSC4b',atsc('TSC4b'))+
      rowVals('Total (Pre Covariance)',function(o){return o.pre_cov;},'sub'))+
    // 4. Post-Covariance (calculated from formulas)
    sect('Post-Covariance',
      rowVals('Post Covariance',function(o){return o.post_cov;})+
      rowVals('Adding 1% conservatism',function(o){return o.add_1pct;})+
      rowVals('Adding 3% loss of covariance',function(o){return o.add_3pct;})+
      rowVals('Total (Post Covariance)',function(o){return o.total_post_cov;},'sub'))+
    // 5. Surplus & Capital
    sect('Surplus & Capital',
      rowVals('Total Surplus',function(o){return o.total_surplus;})+
      rowVals('In LOBs',function(o){return o.in_lobs;})+
      rowVals('Portion in non-ins cos',function(o){return o.portion_non_ins;})+
      rowVals('AVR',function(o){return o.avr;})+
      rowVals('TAC (C&S + AVR)',function(o){return o.tac;},'sub'))+
    // 6. RBC Ratios (calculated)
    sect('RBC Ratios',
      rowVx('RBC Ratio pre-covariance',function(o){return o.ratio_pre_cov;})+
      rowVx('RBC Ratio post-covariance w margin',function(o){return o.ratio_w_margin;})+
      rowVx('RBC Ratio post-covariance w/o margin',function(o){return o.ratio_wo_margin;}))+
    '</div></div>';
}


function buildRBCSection(origRbc, label, annualStmt, rbcResult){
  var adjKey=label.toLowerCase().includes('net')?'net_adjustments':'predeal_adjustments';
  var adjDict=(rbcResult&&rbcResult[adjKey])||{};
  if(!origRbc||(!origRbc.data&&!Object.keys(origRbc).length))
    return '<p style="color:var(--mu);padding:8px">Upload RBC data and run the model.</p>';
  var allData=origRbc.data||origRbc;
  var years=origRbc.years?origRbc.years.filter(function(y){return y>=2025;}).map(Number):
    Object.keys(Object.values(allData)[0]||{}).map(Number).filter(function(y){return y>=2025;}).sort(function(a,b){return a-b;});
  function a(yr){return adjDict[yr]||adjDict[String(yr)]||{};}
  function f3(v){if(v==null||isNaN(v))return '0.0000';return Number(v).toFixed(4);}
  function fx(v){if(v==null||isNaN(v))return '0.00x';return Number(v).toFixed(2)+'x';}
  function th(){return '<thead><tr><th style="min-width:200px">Line Item</th>'+years.map(function(y){return '<th style="min-width:72px">'+y+'</th>';}).join('')+'</tr></thead>';}
  // Always show row, populate with 0 if missing
  function rowV(lbl,fn,cls){
    var vals=years.map(function(y){var v=fn(y);return v!=null?v:0;});
    return '<tr'+(cls?' class="'+cls+'"':'')+'><td>'+lbl+'</td>'+vals.map(function(v){return '<td>'+f3(v)+'</td>';}).join('')+'</tr>';
  }
  function rowX(lbl,fn){
    var vals=years.map(function(y){var v=fn(y);return v!=null?v:0;});
    return '<tr><td>'+lbl+'</td>'+vals.map(function(v){return '<td>'+fx(v)+'</td>';}).join('')+'</tr>';
  }
  function sect(title,rows){return '<div class="sect-hdr">'+title+'</div><div class="tw"><table class="bbt" style="font-size:.68rem">'+th()+'<tbody>'+rows+'</tbody></table></div>';}
  function tsc(key){return function(y){var ms=a(y).ms_tsc||{};return ms[key];};}
  function atsc(key){return function(y){var ap=a(y).allprod_tsc||{};return ap[key];};}
  return '<div class="card"><div class="ch">'+label+'</div><div class="cb">'+
    // 1. MS Income (claims displayed negative)
    sect('Med Sup Income ($M)',
      rowV('Premium',function(y){return a(y).ms_prem;})+
      rowV('Claims',function(y){var v=a(y).ms_claims;return v!=null?-v:null;})+
      rowV('Pre-Tax Income',function(y){return a(y).ms_income;}))+
    // 2. MS TSC Charges - all 8 lines
    sect('Med Sup RBC Charges ($M)',
      rowV('TSC0',tsc('TSC0'))+rowV('TSC1',tsc('TSC1'))+rowV('TSLR016',tsc('TSLR016'))+
      rowV('TSC1CS',tsc('TSC1CS'))+rowV('TSC2',tsc('TSC2'))+rowV('TSC3',tsc('TSC3'))+
      rowV('TSC4a',tsc('TSC4a'))+rowV('TSC4b',tsc('TSC4b')))+
    // 3. All Product TSC Charges - all 8 lines + Total subtotal
    sect('All Product RBC Charges ($M)',
      rowV('TSC0',atsc('TSC0'))+rowV('TSC1',atsc('TSC1'))+rowV('TSLR016',atsc('TSLR016'))+
      rowV('TSC1CS',atsc('TSC1CS'))+rowV('TSC2',atsc('TSC2'))+rowV('TSC3',atsc('TSC3'))+
      rowV('TSC4a',atsc('TSC4a'))+rowV('TSC4b',atsc('TSC4b'))+
      rowV('Total (Pre Covariance)',function(y){return a(y).pre_cov;},'sub'))+
    // 4. Post-Covariance
    sect('Post-Covariance ($M)',
      rowV('Post Covariance',function(y){return a(y).post_cov;})+
      rowV('Adding 1% conservatism',function(y){return a(y).add_1pct;})+
      rowV('Adding 3% loss of covariance',function(y){return a(y).add_3pct;})+
      rowV('Total (Post Covariance)',function(y){return a(y).total_post_cov;},'sub'))+
    // 5. Surplus & Capital
    sect('Surplus & Capital ($M)',
      rowV('Diff in MS income (Cumul.)',function(y){return a(y).cum_surplus_adj;})+
      rowV('Total Surplus',function(y){return a(y).total_surplus;})+
      rowV('In LOBs',function(y){return a(y).in_lobs;})+
      rowV('Portion in non-ins cos',function(y){return a(y).portion_non_ins;})+
      rowV('AVR',function(y){return a(y).avr;})+
      rowV('TAC (C&S + AVR)',function(y){return a(y).tac;},'sub'))+
    // 6. RBC Ratios - pre-cov first
    sect('RBC Ratios',
      rowX('RBC Ratio pre-covariance',function(y){return a(y).ratio_pre_cov;})+
      rowX('RBC Ratio post-covariance w margin',function(y){return a(y).ratio_w_margin;})+
      rowX('RBC Ratio post-covariance w/o margin',function(y){return a(y).ratio_wo_margin;}))+
    '</div></div>';
}


function buildRBCLiftTable(rbc,rbcNet,out){
  if(!out||(!out.rbc_orig_computed&&!out.rbc_predeal_result&&!out.rbc_net_result))
    return '<p style="color:var(--mu);padding:8px">Run the model first.</p>';
  var origVals=(out.rbc_orig_computed&&out.rbc_orig_computed.original_values)||{};
  var predAdj=(out.rbc_predeal_result&&out.rbc_predeal_result.predeal_adjustments)||{};
  var netAdj=(out.rbc_net_result&&out.rbc_net_result.net_adjustments)||{};
  var allYears=((out.rbc_orig_computed&&out.rbc_orig_computed.years)||[]).map(Number).filter(function(y){return y>=2026&&y<=2035;}).sort(function(a,b){return a-b;});
  if(!allYears.length){
    var keys=Object.keys(origVals).concat(Object.keys(predAdj)).concat(Object.keys(netAdj));
    allYears=[].concat.apply([],keys.map(function(k){var n=Number(k);return n>=2026&&n<=2035?[n]:[];}));
    allYears=Array.from(new Set(allYears)).sort(function(a,b){return a-b;});
  }
  if(!allYears.length)return '<p style="color:var(--mu);padding:8px">No RBC data available.</p>';
  function gAdj(d,yr){return d[yr]||d[String(yr)]||{};}
  function fx(x){return(x==null||isNaN(x))?'-':Number(x).toFixed(2)+'x';}
  function fxLift(x){if(x==null||isNaN(x))return '-';return(x>=0?'+':'')+Number(x).toFixed(2)+'x';}
  function liftColor(x){return x==null?'':x>0?'#28a745':x<0?'#C0392B':'';}
  var h='<div class="card"><div class="ch">RBC Ratio Post-Covariance (w/ Margin) — 2026-2035</div><div class="cb">';
  h+='<div class="tw"><table class="bbt" style="font-size:.72rem">';
  h+='<thead><tr><th style="min-width:130px">Scenario</th>'+allYears.map(function(y){return '<th style="min-width:60px">'+y+'</th>';}).join('')+'</tr></thead><tbody>';
  h+='<tr><td style="font-weight:600">Original</td>'+allYears.map(function(y){return '<td>'+fx(gAdj(origVals,y).ratio_w_margin)+'</td>';}).join('')+'</tr>';
  h+='<tr><td style="font-weight:600">Predeal</td>'+allYears.map(function(y){return '<td>'+fx(gAdj(predAdj,y).ratio_w_margin)+'</td>';}).join('')+'</tr>';
  h+='<tr><td style="font-weight:600">Net (Post-Deal)</td>'+allYears.map(function(y){return '<td>'+fx(gAdj(netAdj,y).ratio_w_margin)+'</td>';}).join('')+'</tr>';
  h+='<tr class="sub"><td>Lift (Net − Original)</td>'+allYears.map(function(y){var o=gAdj(origVals,y).ratio_w_margin,n=gAdj(netAdj,y).ratio_w_margin;var lift=(o!=null&&n!=null)?n-o:null;return '<td style="color:'+liftColor(lift)+'">'+fxLift(lift)+'</td>';}).join('')+'</tr>';
  h+='</tbody></table></div></div></div>';
  return h;
}

function showRBCSub(which,btn){['rbc-original','rbc-net','rbc-lift'].forEach(id=>{const el=document.getElementById(id);if(el)el.style.display='none';});const el=document.getElementById('rbc-'+which);if(el)el.style.display='';btn.closest('.stabs').querySelectorAll('.stab').forEach(b=>b.classList.remove('active'));btn.classList.add('active');}

// -- IY DIAGNOSTIC --
function renderIYDiag(){
  const iy=parseInt(document.getElementById('iy-select').value);const div=document.getElementById('iy-diag-content');
  if(!iy||!S.out){div.innerHTML='';return;}
  const diag=(S.out.iy_diagnostic||{})[String(iy)];if(!diag){div.innerHTML='<p style="color:var(--mu)">No data.</p>';return;}
  const active=diag.filter(d=>d.expected_rate&&d.expected_rate>0);
  const passAll=active.every(cy=>cy.lines.every(l=>l.ok));
  let html='<div style="margin-bottom:10px;padding:7px 12px;border-radius:var(--r);background:'+(passAll?'#D4EDDA':'#FEF0EE')+';border:1px solid '+(passAll?'#A8D5B5':'#FACACA')+';font-size:.76rem;font-weight:bold;color:'+(passAll?'#155724':'#721c24')+'">'+(passAll?'All pass':'Some fail')+' - IssYr '+iy+'</div>';
  for(const cd of diag){
    const exp=cd.expected_rate;const ep=exp!=null?(exp*100).toFixed(1)+'%':'-';const rel=exp&&exp>0;const ok=cd.lines.every(l=>l.ok);
    html+='<div style="margin-bottom:12px"><div style="font-size:.7rem;font-weight:bold;margin-bottom:4px;display:flex;align-items:center;gap:8px;flex-wrap:wrap"><span>CalYr '+cd.cal_year+'</span><span style="font-weight:normal;color:var(--mu)">Expected: '+ep+'</span>'+(rel?(ok?'<span class="tag tgy">Pass</span>':'<span class="tag" style="background:#C0392B;color:#fff">Fail</span>'):'<span class="tag tgg">N/A</span>')+'</div>'+
    '<div class="tw"><table class="bbt" style="font-size:.65rem"><thead><tr><th>Line</th><th>Direct ($)</th><th>Ceded ($)</th><th>Actual %</th><th>Expected</th><th>Status</th></tr></thead><tbody>';
    for(const l of cd.lines){
      const ar=l.actual_rate!=null?(l.actual_rate*100).toFixed(1)+'%':'-';
      const lep=l.expected_rate!=null?(l.expected_rate*100).toFixed(1)+'%':ep;
      const isSkip=l.zero_skip;
      const st=isSkip?'<span class="tag tgg">N/A (zero)</span>':!rel?'<span class="tag tgg">N/A</span>':l.ok?'<span class="tag tgy">OK</span>':'<span class="tag" style="background:#C0392B;color:#fff">!</span>';
      html+='<tr'+(rel&&!isSkip?'':' style="opacity:.5"')+'><td style="font-weight:600">'+l.line+'</td><td>'+fmtDol(l.direct)+'</td><td>'+fmtDol(l.ceded)+'</td><td style="font-weight:bold">'+ar+'</td><td>'+lep+'</td><td style="text-align:center">'+st+'</td></tr>';
    }
    html+='</tbody></table></div></div>';
  }
  div.innerHTML=html;
}

// -- DOCS --
function renderDocs(){
  const div=document.getElementById('docs-content');if(!div)return;
  div.dataset.rendered='1';
  const text=DOC_TEXT||'No documentation loaded.';
  function inline(s){
    return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
      .replace(/\*\*(.+?)\*\*/g,'<strong style="color:var(--navy)">$1</strong>')
      .replace(/`(.+?)`/g,'<code style="background:var(--gray);padding:1px 5px;border-radius:3px;font-size:.92em;font-family:var(--fm)">$1</code>');
  }
  let html='<div style="max-width:900px">';
  text.split('\n').forEach(raw=>{
    const t=raw.replace(/\s+$/,'');
    if(/^#\s+/.test(t)){html+='<div style="font-family:var(--ft);font-size:1.4rem;color:var(--navy);font-weight:bold;margin:0 0 2px">'+inline(t.slice(2))+'</div>';}
    else if(/^##\s+/.test(t)){html+='<div style="font-family:var(--ft);font-size:1.0rem;color:var(--yel);background:var(--navy);padding:7px 13px;border-radius:4px;margin:22px 0 9px;letter-spacing:.02em">'+inline(t.slice(3))+'</div>';}
    else if(/^###\s+/.test(t)){html+='<div style="font-size:.85rem;font-weight:bold;color:var(--navy);margin:13px 0 4px;border-bottom:1px solid var(--bdr);padding-bottom:3px">'+inline(t.slice(4))+'</div>';}
    else if(/^>\s+/.test(t)){html+='<div style="font-size:.76rem;color:var(--navy);background:#EEF2FF;border-left:3px solid var(--navy);padding:7px 11px;margin:7px 0;border-radius:3px;line-height:1.5">'+inline(t.slice(2))+'</div>';}
    else if(/^\s*-\s+/.test(t)){const sub=/^\s\s+-/.test(t);html+='<div style="display:flex;gap:8px;padding-left:'+(sub?'30':'14')+'px;margin-bottom:3px;font-size:.79rem;line-height:1.5"><span style="color:var(--gold)">&#9642;</span><span>'+inline(t.replace(/^\s*-\s+/,''))+'</span></div>';}
    else if(t.trim()===''){html+='<div style="height:7px"></div>';}
    else{html+='<p style="margin-bottom:5px;font-size:.79rem;line-height:1.62;color:var(--dk)">'+inline(t)+'</p>';}
  });
  html+='</div>';div.innerHTML=html;
}

// ── SCENARIOS ──
function saveScenario(){
  if(!S.out){alert('Run model first.');return;}
  const name=prompt('Scenario name:','Scenario '+(S.scenarios.length+1));if(!name)return;
  S.scenarios.unshift({id:S.runId,name,saved_at:new Date().toISOString(),
    assumptions:JSON.parse(JSON.stringify(S.assumptions)),
    metrics_predeal:S.out.metrics_predeal,
    metrics_ceded:S.out.metrics_ceded,
    metrics_net:S.out.metrics_net,
    annual_predeal:S.out.annual_predeal,
    annual_ceded:S.out.annual_ceded,
    annual_net:S.out.annual_net,
    rbc_data:S.out.rbc_data,
    rbc_net:S.out.rbc_net,
    rbc_orig_computed:S.out.rbc_orig_computed,
    rbc_predeal_result:S.out.rbc_predeal_result,
    rbc_net_result:S.out.rbc_net_result});
  localStorage.setItem('bb_sc',JSON.stringify(S.scenarios));
  document.getElementById('sccount').textContent=S.scenarios.length;renderScenarios();
}
// SCENARIO COMPARISON CODE - inject into renderScenarios

function renderScenarios(){
  const div=document.getElementById('sc-list');if(!div)return;
  const scenarios=S.scenarios||[];
  if(!scenarios.length){
    div.innerHTML='<div class="empty"><div class="ei">&#x1F4CB;</div><h3>No saved scenarios</h3><p>Run the model and click Save Scenario to compare runs</p></div>';
    document.getElementById('sc-compare').innerHTML='';
    return;
  }
  const metrics=[{k:'pvde_pre',l:'Predeal EV ($M)'},{k:'pvde_net',l:'Net EV ($M)'},
    {k:'irr_pre',l:'Predeal IRR'},{k:'irr_net',l:'Net IRR'},
    {k:'pvde_ced',l:'Ceded EV ($M)'},{k:'max_neg',l:'Max Neg DE ($M)'}];
  let h='<div style="overflow-x:auto"><table class="bbt"><thead><tr>'+
    '<th></th><th>Scenario</th><th>Date</th>'+
    metrics.map(m=>'<th>'+m.l+'</th>').join('')+'<th>Actions</th></tr></thead><tbody>';
  const sel=S._scSelSet||new Set();
  scenarios.forEach(function(sc,i){
    const mp=sc.metrics_predeal||{},mc2=sc.metrics_ceded||{},mn=sc.metrics_net||{};
    const isChecked=sel.has(i);
    h+='<tr>'+
      '<td><input type="checkbox"'+(isChecked?' checked':'')+' onchange="toggleScCompare('+i+',this.checked)"></td>'+
      '<td style="font-weight:600;max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'+sc.name+'</td>'+
      '<td style="font-size:.68rem;color:var(--mu)">'+new Date(sc.saved_at||sc.ts||'').toLocaleDateString()+'</td>'+
      '<td>'+fmtM(mp.pvde)+'M</td><td>'+fmtM(mn.pvde)+'M</td>'+
      '<td>'+fmtPct(mp.irr)+'</td><td>'+fmtPct(mn.irr)+'</td>'+
      '<td>'+fmtM(mc2.pvde)+'M</td><td>'+fmtM(mn.max_neg_cum_de)+'M</td>'+
      '<td><button class="btn btn-o btn-sm" onclick="loadSc('+i+')" style="margin-right:4px">Load</button>'+
      '<button onclick="delSc('+i+')" style="color:var(--err);background:none;border:none;cursor:pointer;font-size:.75rem">&#10005;</button></td></tr>';
  });
  h+='</tbody></table></div>';
  div.innerHTML=h;
  renderScCompare();
}

function toggleScCompare(i,checked){
  if(!S._scSelSet)S._scSelSet=new Set();
  if(checked){
    if(S._scSelSet.size>=2){
      // Deselect oldest, add new
      const oldest=[...S._scSelSet][0];
      S._scSelSet.delete(oldest);
    }
    S._scSelSet.add(i);
  } else { S._scSelSet.delete(i); }
  renderScenarios();
}

function rbcLiftSection(sc,color){
  var rbn=sc.rbc_net||{},rbco=sc.rbc_data||{};
  function rr(src,yr){var v=src['RBC Ratio post-covariance w margin']||{};return v[yr]||v[String(yr)]||null;}
  var yrs=Object.keys((rbco['RBC Ratio post-covariance w margin']||{})).map(Number).filter(function(y){return y>=2026;}).sort(function(a,b){return a-b;}).slice(0,5);
  if(!yrs.length)return '<div style="font-size:.72rem;color:var(--mu)">RBC data not available</div>';
  var rows='<tr><td style="color:var(--mu)">Original RBC</td>'+
    yrs.map(function(y){var v=rr(rbco,y);return '<td>'+(v?v.toFixed(2)+'x':'-')+'</td>';}).join('')+'</tr>'+
    '<tr><td style="color:var(--mu)">Post-Deal RBC</td>'+
    yrs.map(function(y){var v=rr(rbn,y);return '<td>'+(v?v.toFixed(2)+'x':'-')+'</td>';}).join('')+'</tr>'+
    '<tr class="sub"><td>Lift</td>'+
    yrs.map(function(y){
      var o=rr(rbco,y),n=rr(rbn,y);var lift=o&&n?n-o:null;
      return '<td style="color:'+(lift&&lift>0?'#28a745':lift&&lift<0?'#C0392B':'var(--mu)')+'">'+
        (lift!=null?(lift>=0?'+':'')+lift.toFixed(2)+'x':'-')+'</td>';
    }).join('')+'</tr>';
  return '<div style="margin-top:4px"><div style="font-size:.71rem;font-weight:bold;padding:2px 8px;background:'+color+';color:#fff;border-radius:3px;margin-bottom:3px">RBC Ratio History</div>'+
    '<div class="tw"><table class="bbt" style="font-size:.67rem"><thead><tr><th>Metric</th>'+
    yrs.map(function(y){return '<th>'+y+'</th>';}).join('')+'</tr></thead><tbody>'+rows+'</tbody></table></div></div>';
}

function renderScCompare(){
  const cDiv=document.getElementById('sc-compare');if(!cDiv)return;
  const sel=S._scSelSet;
  if(!sel||sel.size<2){
    if(sel&&sel.size===1)cDiv.innerHTML='<div style="color:var(--mu);font-size:.76rem;margin-top:8px;padding:6px 0">Select one more scenario to compare.</div>';
    else cDiv.innerHTML='';
    return;
  }
  const scenarios=S.scenarios||[];
  const [i1,i2]=[...sel].sort((a,b)=>a-b);
  const sc1=scenarios[i1],sc2=scenarios[i2];
  if(!sc1||!sc2){cDiv.innerHTML='';return;}

  // Get 5 years from each scenario
  function getYrs(sc){
    var p=sc.annual_predeal&&sc.annual_predeal.periods||[];
    if(!p.length)p=Object.keys((sc.annual_predeal&&sc.annual_predeal.premium)||{}).map(Number).filter(function(n){return n>2000;}).sort();
    return p.slice(0,5);
  }
  var yrs1=getYrs(sc1),yrs2=getYrs(sc2);
  var allYrs=[...new Set([...yrs1,...yrs2])].sort(function(a,b){return a-b;}).slice(0,5);

  function scVal(sc,stmtKey,lineKey,yr){
    var stmt=sc[stmtKey]||{};
    var line=stmt[lineKey]||{};
    return line[yr]||line[String(yr)]||0;
  }

  // Six metric cards for each scenario
  function metricCards(sc,label,borderColor){
    var mp=sc.metrics_predeal||{},mc2=sc.metrics_ceded||{},mn=sc.metrics_net||{};
    var strain=(mn.max_neg_cum_de||0)-(mp.max_neg_cum_de||0);
    var cards=[
      {l:'Predeal EV',v:fmtM(mp.pvde)+'M',s:'PVDE'},
      {l:'Net EV',v:fmtM(mn.pvde)+'M',s:'PVDE'},
      {l:'Predeal IRR',v:fmtPct(mp.irr),s:'IRR'},
      {l:'Net IRR',v:fmtPct(mn.irr),s:'IRR'},
      {l:'Ceded EV',v:fmtM(mc2.pvde)+'M',s:'PVDE'},
      {l:'Strain Relief',v:fmtM(strain)+'M',s:'Max neg DE'}
    ];
    return '<div style="border-top:3px solid '+borderColor+';padding-top:8px">'+
      '<div style="font-size:.78rem;font-weight:bold;color:var(--navy);margin-bottom:6px">'+label+'</div>'+
      '<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:6px">'+
      cards.map(function(c){
        return '<div class="mtile" style="margin:0"><div class="ml">'+c.l+'</div>'+
          '<div class="mv" style="font-size:.88rem">'+c.v+'</div>'+
          '<div class="ms">'+c.s+'</div></div>';
      }).join('')+'</div></div>';
  }

  var SC1_COLOR='#2c5f9e', SC2_COLOR='#1a7a4a';

  // Income statement table for one scenario, one view
  function incomeTable(sc,stmtKey,titleLabel,yrs,borderColor){
    var pairs=[
      ['Total Revenue','revenue'],['Total Benefits','benefits'],
      ['Selling Expense','selling_expense'],['Op Expense','op_expense'],
      ['Pre-Tax Income','pretax_income'],['Distrib. Earnings','distributable_earnings']
    ];
    var subLines=new Set(['Total Revenue','Pre-Tax Income']);
    // RBC ratio: use net rbc for net stmt, original rbc for predeal
    // Pull from saved compute results: rbc_net_result for Net, rbc_predeal_result for Predeal
    var scAdj={};
    if(stmtKey==='annual_net'){scAdj=(sc.rbc_net_result&&sc.rbc_net_result.net_adjustments)||{};}
    else if(stmtKey==='annual_predeal'){scAdj=(sc.rbc_predeal_result&&sc.rbc_predeal_result.predeal_adjustments)||{};}
    var scOrigVals=(sc.rbc_orig_computed&&sc.rbc_orig_computed.original_values)||{};
    var rbcRow='<tr class="ratio"><td>RBC Ratio (w/ Margin)</td>'+
      yrs.map(function(y){
        var a=scAdj[y]||scAdj[String(y)];
        var v=a?a.ratio_w_margin:null;
        if(v==null){var o=scOrigVals[y]||scOrigVals[String(y)];if(o)v=o.ratio_w_margin;}
        return '<td>'+(v!=null?Number(v).toFixed(2)+'x':'-')+'</td>';
      }).join('')+'</tr>';
    return '<div style="margin-bottom:8px">'+
      '<div style="font-size:.71rem;font-weight:bold;padding:2px 8px;background:'+borderColor+';color:#fff;border-radius:3px;margin-bottom:3px">'+titleLabel+'</div>'+
      '<div class="tw"><table class="bbt" style="font-size:.67rem">'+
      '<thead><tr><th>Line</th>'+yrs.map(function(y){return '<th>'+y+'</th>';}).join('')+'</tr></thead><tbody>'+
      pairs.map(function(p){
        var lbl=p[0],key=p[1];
        return '<tr'+(subLines.has(lbl)?' class="sub"':'')+'><td>'+lbl+'</td>'+
          yrs.map(function(y){return '<td>'+fmtNum(scVal(sc,stmtKey,key,y))+'</td>';}).join('')+'</tr>';
      }).join('')+
      rbcRow+
      '</tbody></table></div></div>';
  }

  var html='<div style="border-top:2px solid var(--navy);margin-top:16px;padding-top:14px">'+
    '<div style="font-size:.88rem;font-weight:bold;color:var(--navy);margin-bottom:12px">Scenario Comparison</div>'+
    // Two-column layout: sc1 | divider | sc2
    '<div style="display:grid;grid-template-columns:1fr 2px 1fr;gap:0 16px">'+
      // Left: Scenario 1
      '<div>'+
        metricCards(sc1,sc1.name,SC1_COLOR)+
        '<div style="margin-top:12px">'+
          incomeTable(sc1,'annual_predeal','Predeal',yrs1,SC1_COLOR)+
          incomeTable(sc1,'annual_ceded','Ceded',yrs1,SC1_COLOR)+
          incomeTable(sc1,'annual_net','Net',yrs1,SC1_COLOR)+
        '</div>'+
      '</div>'+
      // Divider
      '<div style="background:var(--bdr)"></div>'+
      // Right: Scenario 2
      '<div>'+
        metricCards(sc2,sc2.name,SC2_COLOR)+
        '<div style="margin-top:12px">'+
          incomeTable(sc2,'annual_predeal','Predeal',yrs2,SC2_COLOR)+
          incomeTable(sc2,'annual_ceded','Ceded',yrs2,SC2_COLOR)+
          incomeTable(sc2,'annual_net','Net',yrs2,SC2_COLOR)+
        '</div>'+
      '</div>'+
    // RBC Lift comparison row - same grid as tables above
    '</div></div>';
  cDiv.innerHTML=html;
}



function loadSc(i){S.assumptions=JSON.parse(JSON.stringify(S.scenarios[i].assumptions));renderAssumUI();showTab('assumptions');}
function delSc(i){if(!confirm('Delete?'))return;S.scenarios.splice(i,1);localStorage.setItem('bb_sc',JSON.stringify(S.scenarios));renderScenarios();}
function clearScenarios(){if(!confirm('Clear all?'))return;S.scenarios=[];localStorage.removeItem('bb_sc');renderScenarios();}

// ── AUDIT ──
// ── REVIEW ──
function updateReview(out){
  document.getElementById('rev-content').style.display='';
  const a=S.assumptions||{};
  const ev=(id,txt)=>{const el=document.getElementById(id);if(el)el.textContent=txt;};
  ev('re1',(out.ev_records_count||0).toLocaleString()+' rows');
  const iys=out.iss_years||[];
  ev('re2',iys.length>10?iys.slice(0,5).join(',')+' ... '+iys.slice(-3).join(','):iys.join(', '));
  ev('re3','Max: '+out.max_period+(out.max_period===360?' OK':' !'));
  const rp=a.reins_pct||{};
  ev('re4',Object.keys(rp).length+' IssYrs defined');
  const iyd=out.iy_diagnostic||{};
  const total=Object.values(iyd).flatMap(rows=>rows.flatMap(r=>r.lines.filter(l=>r.expected_rate&&r.expected_rate>0))).length;
  const pass=Object.values(iyd).flatMap(rows=>rows.flatMap(r=>r.lines.filter(l=>r.expected_rate&&r.expected_rate>0&&l.ok))).length;
  ev('re5',pass+'/'+total+' pass');
  ev('re6','Ongoing: $'+(a.ceding_comm_ongoing||'-')+'/pol/yr');
  const ap=out.annual_predeal||{};const fy=(ap.periods||[])[0];
  const prem=gv(ap.premium||{},fy)||0;const clms=Math.abs(gv(ap.claims||{},fy)||0);
  ev('re7',prem?((clms/Math.abs(prem))*100).toFixed(1)+'% yr '+fy:'-');
  ev('re8','Net: '+fmtM(out.metrics_net?.pvde)+'M vs Pre: '+fmtM(out.metrics_predeal?.pvde)+'M');
  ev('re9',prem?'Model: '+(prem/1e6).toFixed(1)+'M':'-');
  document.getElementById('rev-summary').innerHTML=
    '<table style="width:100%;font-size:.7rem;border-collapse:collapse">'+
    [['Run ID',S.runId||'-'],['Time',new Date().toLocaleString()],['EV Records',(out.ev_records_count||0).toLocaleString()],
     ['Max Period',out.max_period+''],['Predeal EV',fmtM(out.metrics_predeal?.pvde)+'M'],
     ['Net EV',fmtM(out.metrics_net?.pvde)+'M'],['Net IRR',fmtPct(out.metrics_net?.irr)],
     ['Discount Rate',fmtPct(a.discount_rate)],
    ].map(([l,v])=>'<tr style="border-bottom:1px solid var(--bdr)"><td style="padding:3px 0;color:var(--mu);width:50%">'+l+'</td><td style="font-family:var(--fm);font-weight:bold">'+v+'</td></tr>').join('')+'</table>';
}

// ── EXPORT ──
function exportCSV(view){
  const t=document.getElementById('tbl-'+view);if(!t)return;
  const rows=Array.from(t.querySelectorAll('tr')).map(tr=>
    Array.from(tr.querySelectorAll('th,td')).map(td=>'"'+td.textContent.trim()+'"').join(',')
  ).join('\n');
  const b=new Blob([rows],{type:'text/csv'});
  const a=document.createElement('a');a.href=URL.createObjectURL(b);
  a.download='bb_'+view+'_'+new Date().toISOString().slice(0,10)+'.csv';a.click();
}

function downloadRBCTemplate(fmt){
  if(fmt==='csv'){
    var hdr='Row Label,2025,2026,2027,2028,2029,2030,2031,2032,2033,2034,2035';
    var rows=[hdr,
      '// === Med Sup Section (rows used for MS charges) ===',
      '// Labels below are exactly as parsed by the model',
      'Premium,464.57,683.97,918.98,1152.78,1384.94,1593.53,1764.65,1904.27,2013.34,2103.73,2173.35',
      'Claims,-430.80,-625.12,-805.61,-972.70,-1118.28,-1248.67,-1350.74,-1429.31,-1493.20,-1542.38,-1582.44',
      'Income,-93.28,-119.34,-98.36,-51.43,17.48,78.61,134.05,189.36,226.03,258.24,280.55',
      'TSC0,0,0,0,0,0,0,0,0,0,0,0',
      'TSC1,1.25,1.70,2.19,2.64,3.06,3.39,3.64,3.83,3.97,4.07,4.14',
      'TSLR016,0.07,0.06,0.06,0.06,0.05,0.05,0.04,0.04,0.04,0.03,0.03',
      'TSC1CS,0,0,0,0,0,0,0,0,0,0,0',
      'TSC2,26.28,38.58,49.72,60.03,69.02,77.06,83.36,88.21,92.16,95.19,97.66',
      'TSC3,0.51,0.51,0.51,0.51,0.51,0.51,0.51,0.51,0.51,0.51,0.51',
      'TSC4a,2.42,3.40,4.57,5.74,6.89,7.93,8.78,9.48,10.02,10.47,10.82',
      'TSC4b,0.20,0.20,0.20,0.20,0.20,0.20,0.20,0.20,0.20,0.20,0.20',
      '// === All product Section (parser detects this header and prefixes keys with All_) ===',
      'All product',
      'TSC0,1.82,2.09,2.36,2.63,2.89,3.16,3.43,3.70,3.96,3.96,3.96',
      'TSC1,35.58,38.25,40.64,43.46,46.76,50.07,52.99,56.03,60.16,65.68,71.79',
      'TSLR016,1.04,0.94,0.86,0.79,0.73,0.68,0.63,0.59,0.55,0.52,0.50',
      'TSC1CS,9.54,10.51,11.67,11.67,11.80,11.98,12.48,13.58,14.76,16.63,13.67',
      'TSC2,44.72,56.91,68.08,78.54,87.75,95.94,102.37,107.46,111.70,115.13,118.13',
      'TSC3,21.45,22.21,22.95,23.87,24.96,26.17,27.43,28.82,30.39,32.09,33.92',
      'TSC4a,8.44,9.46,10.75,12.29,13.94,15.56,17.09,18.53,19.88,21.21,22.52',
      'TSC4b,0.20,0.20,0.20,0.20,0.20,0.20,0.20,0.20,0.20,0.20,0.20',
      'Total (Pre Covariance),122.79,140.56,157.52,173.46,189.03,203.75,216.62,228.91,241.60,255.42,264.70',
      '// === Post-Covariance ===',
      'Post Covariance,84.18,95.92,107.59,119.54,131.24,142.27,151.69,160.18,168.74,177.46,185.93',
      'Conservatism factor,0,0.00714,0.02143,0.03571,0.05,0.05,0.05,0.05,0.05,0.05,0.05',
      'Adding 1% conservatism,0,0.69,2.31,4.27,6.56,7.11,7.58,8.01,8.44,8.87,9.30',
      'Adding 3% loss of covariance,0,2.88,3.23,3.59,3.94,4.27,4.55,4.81,5.06,5.32,5.58',
      'Total (Post Covariance),84.18,99.49,113.12,127.40,141.74,153.65,163.82,173.00,182.24,191.66,200.81',
      '// === Surplus & Capital ===',
      'Total Surplus,649.40,553.29,479.42,451.00,490.84,599.99,745.20,943.54,1181.53,1457.35,1763.88',
      'In LOBs,372.19,430.87,485.14,536.14,585.98,633.09,674.27,713.58,751.96,789.44,810.88',
      'Portion in non-ins cos,118.28,110.66,22.50,22.50,22.50,22.50,83.43,188.71,236.31,291.47,352.78',
      'Portion in ins co Corp,158.93,11.76,-28.22,-107.64,-117.63,-55.60,-12.50,41.26,193.26,376.44,600.22',
      'AVR,34.01,34.01,34.01,34.01,34.01,34.01,34.01,34.01,34.01,34.01,34.01',
      'TAC (C&S + AVR),565.12,476.64,490.93,462.51,502.35,611.49,695.78,788.84,979.23,1199.89,1445.11',
      '// === RBC Ratios ===',
      'RBC Ratio pre-covariance,4.60,3.39,3.12,2.67,2.66,3.00,3.21,3.45,4.05,4.70,5.46',
      'RBC Ratio post-covariance w margin,6.71,4.79,4.34,3.63,3.54,3.98,4.25,4.56,5.37,6.26,7.20',
      'RBC Ratio post-covariance w/o margin,6.71,4.82,4.43,3.76,3.72,4.17,4.45,4.78,5.63,6.56,7.55'];
    _downloadCSV(rows.join('\n'),'RBC_Original_template.csv');
  } else {
    alert('Download the CSV template and populate with your actual values.');
  }
}

function downloadBPTemplate(fmt){
  if(fmt==='csv'){
    var rows=['Row Label,2026,2027,2028,2029,2030,2031,2032,2033,2034,2035',
      '// Balance Plan Template - annual plan figures by year',
      '// Rows can include: EarnedPrem, IncClaims, Commissions, etc.',
      'EarnedPrem,0,0,0,0,0,0,0,0,0,0',
      'IncClaims,0,0,0,0,0,0,0,0,0,0'];
    _downloadCSV(rows.join('\n'),'BalancePlan_template.csv');
  } else {
    alert('For XLSX template: use the CSV template structure with the same columns, saved as .xlsx');
  }
}

function _downloadCSV(content,filename){
  var blob=new Blob([content],{type:'text/csv;charset=utf-8;'});
  var a=document.createElement('a');a.href=URL.createObjectURL(blob);
  a.download=filename;a.click();setTimeout(function(){URL.revokeObjectURL(a.href);},1000);
}

function downloadEVTemplate(){
  // Generate EV Data CSV template with correct headers
  // KEY columns + Value000...Value360 for 30-year projection
  const keyCols=['ScenId','SensId','ck.IssYear',
    'ProjPer','ProjMode','VarName','Description','Module','FormType','Group',
    'Width','Decimals','Scale'];
  const valCols=[];
  for(let i=0;i<=360;i++)valCols.push('Value'+String(i).padStart(3,'0'));
  const header=keyCols.concat(valCols).join(',');
  // Sample row showing key variables
  const sampleRows=[
    '// EV Data CSV Template - Remove this comment line before loading',
    '// Required columns: VarName, ck.IssYear, Value000-Value360',
    '// KEY_VARS: EarnedPrem,ReinsPrem,IncClaims,ReinsClaims,TabRes,CededALRstat,CLRes,TS,Comm,PremTax,LivesIssued,AEGAdminPolCount',
    '// Value000=Dec baseline stock, Value001=Jan yr1, Value012=Dec yr1, Value013=Jan yr2, etc.',
    '1,1,2019,1,Y,MSP,1,Monthly,EarnedPrem,Earned Premium,EV,A00,A,4,2,No'
      +',0'.repeat(361)  // placeholder zeros
  ];
  const csvContent=[header,...sampleRows].join('\n');
  const blob=new Blob([csvContent],{type:'text/csv;charset=utf-8;'});
  const a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download='EV_Data_template.csv';
  a.click();
  setTimeout(()=>URL.revokeObjectURL(a.href),1000);
}

// ── INIT ──
document.addEventListener('DOMContentLoaded', function() { initPy(); });

function cedantCard(out){
  var ca=(out&&out.cedant_analytics)||null;if(!ca)return '';
  var mp=out.metrics_predeal||{},mn=out.metrics_net||{};
  function f(v,d){return (v==null||isNaN(v))?'-':Number(v).toFixed(d==null?1:d);}
  function pct(v,dd){return (v==null||isNaN(v))?'-':(v*100).toFixed(dd==null?0:dd)+'%';}
  function signM(v){return v==null||isNaN(v)?'-':(v>=0?'+':'')+Number(v).toFixed(1);}
  // Upfront ceding commission schedule (nominal $M by year) from the current assumptions
  var ccf=(S.assumptions&&S.assumptions.ceding_comm_front)||{};
  var ccfYrs=Object.keys(ccf).map(Number).sort(function(a,b){return a-b;}).filter(function(y){return ccf[y];});
  var upfrontSched=ccfYrs.length?ccfYrs.map(function(y){return (+ccf[y]).toFixed(0);}).join('-'):'-';
  // Early-strain relief ($M): improvement in peak negative cumulative DE (predeal vs net)
  var strainRelief=((mp.max_neg_cum_de||0)-(mn.max_neg_cum_de||0))/1e6;
  // RBC ratio lift at the predeal trough year (computed RBC results)
  var pa=((out.rbc_predeal_result||{}).predeal_adjustments)||{};
  var na=((out.rbc_net_result||{}).net_adjustments)||{};
  function rat(adj,yr){var r=adj[yr]||adj[String(yr)];return r&&r.ratio_w_margin!=null?r.ratio_w_margin:null;}
  var rbcLift=null,troughYr=null,troughR=Infinity;
  Object.keys(pa).map(Number).filter(function(y){return y>=2026;}).forEach(function(y){
    var r=rat(pa,y);if(r!=null&&r<troughR){troughR=r;troughYr=y;}});
  if(troughYr!=null){var nr=rat(na,troughYr);if(nr!=null)rbcLift=nr-troughR;}
  var irrChg=(mp.irr!=null&&mn.irr!=null)?(mn.irr-mp.irr):null;
  var mb=ca.metrics_back||{},mnew=ca.metrics_new||{};
  function irrPair(a,b){return (a==null?'-':pct(a,1))+' &rarr; '+(b==null?'-':pct(b,1));}
  function mM(v){return v==null||isNaN(v)?'-':(v<0?'('+Math.abs(v).toFixed(0)+')':v.toFixed(0))+'M';}  // value already in $M
  function pvPair(a,b){return mM(a)+' &rarr; '+mM(b);}
  var tile=function(lbl,val,desc,accent){return '<div style="flex:1 1 170px;min-width:160px;padding:8px 10px;background:var(--off);border-radius:5px;border-left:3px solid '+(accent||'var(--bdr)')+'">'+
    '<div style="font-size:.6rem;color:var(--mu);text-transform:uppercase;letter-spacing:.3px">'+lbl+'</div>'+
    '<div style="font-size:1.02rem;font-weight:bold;color:var(--navy);margin:1px 0">'+val+'</div>'+
    '<div style="font-size:.62rem;color:var(--mu);line-height:1.3">'+desc+'</div></div>';};
  var GIVE='#B5332333',GET='#1A8F5A';  // give-up vs benefit accent
  return '<div style="margin-bottom:14px;border:1px solid var(--bdr);border-radius:6px;padding:10px">'+
    '<div style="font-size:.78rem;font-weight:bold;color:var(--navy);margin-bottom:8px">Cedant Economics <span style="font-weight:normal;color:var(--mu);font-size:.64rem">(all $M unless noted; Cost of Capital '+pct(ca.cost_of_capital)+')</span></div>'+
    '<div style="font-size:.6rem;color:var(--mu);text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px">What you give up</div>'+
    '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px">'+
      tile('Ceded PVDE',f(ca.reinsurer_pvde),'Profit (PV of distributable earnings) handed to the reinsurer.','#B53323')+
      tile('IRR change',(mp.irr==null?'-':pct(mp.irr,1))+' &rarr; '+(mn.irr==null?'-':pct(mn.irr,1)),'Book IRR before &rarr; after the deal ('+(irrChg==null?'-':signM(irrChg*100)+' pts')+') &mdash; return given up for the relief.','#B53323')+
    '</div>'+
    '<div style="font-size:.6rem;color:var(--mu);text-transform:uppercase;letter-spacing:.4px;margin-bottom:4px">What you get</div>'+
    '<div style="display:flex;gap:8px;flex-wrap:wrap">'+
      tile('Upfront commission',upfrontSched,'Front-end ceding commission paid by the reinsurer, $M by year (e.g. 10-5-5 = $10M/$5M/$5M in 2026/27/28) &mdash; a component of the early-strain relief.',GET)+
      tile('Early-strain relief',signM(strainRelief),'Reduction in peak negative cumulative earnings &mdash; early capital you no longer must hold (incl. the upfront commission above).',GET)+
      tile('RBC ratio lift',(rbcLift==null?'-':signM(rbcLift)+'x'),'Improvement in RBC ratio'+(troughYr?' at the '+troughYr+' trough':'')+' &mdash; regulatory capital freed.',GET)+
      tile('Cedant recovery',pct(ca.value_recovery_pct),'Commissions received as a share of the value ceded &mdash; already netted out of the Ceded PVDE above (not additive to it).',GET)+
    '</div>'+
    '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:8px">'+
      tile('Portfolio EV',(mp.pvde==null?'-':fmtM(mp.pvde)+'M')+' &rarr; '+(mn.pvde==null?'-':fmtM(mn.pvde)+'M'),'Total book value of new business (PV distributable earnings) before &rarr; after the deal.','var(--bdr)')+
      tile('Back-book (IssYr&le;2025)','EV '+pvPair(mb.predeal_pvde,mb.net_pvde)+'<br>IRR '+irrPair(mb.predeal_irr,mb.net_irr),'In-force book before &rarr; after the deal. EV (book value) is the value signal &mdash; it falls as profit is ceded; the IRR is an outlay-free stream rate and can move counter-intuitively.','var(--bdr)')+
      tile('New-issue (IssYr&ge;2026)','EV '+pvPair(mnew.predeal_pvde,mnew.net_pvde)+'<br>IRR '+irrPair(mnew.predeal_irr,mnew.net_irr),'New business before &rarr; after the deal. Read EV (book value) as the value given up; IRR is a stream rate shown for reference.','var(--bdr)')+
    '</div></div>';
}
function renderSummary(){
  var div=document.getElementById('summary-content');
  if(!div)return;
  if(!S.out){div.innerHTML='<div class="empty"><div class="ei">&#x25B6;</div><h3>Run the model first</h3></div>';return;}
  var out=S.out,ap=out.annual_predeal||{},ac=out.annual_ceded||{},an=out.annual_net||{};
  var mp=out.metrics_predeal||{},mc=out.metrics_ceded||{},mn=out.metrics_net||{};
  var a=S.assumptions||{};
  var by=out.first_yr||2026;
  var yrs=(ap.periods||[]).filter(function(y){return y>=by;}).slice(0,5);
  var rbn=out.rbc_net||{},rbco=out.rbc_data||{};
  var fy=yrs[0]||by;
  function rr(src,yr){var v=src['RBC Ratio post-covariance w margin']||{};return v[yr]||v[String(yr)]||null;}
  var rpIys=Object.keys(a.reins_pct||{}).map(Number).sort(function(a,b){return a-b;});
  var rpPct=rpIys.length?(Object.values((a.reins_pct||{})[rpIys[0]]||{})[0]||0)+'%':'--';
  var rpRange=rpIys.length?((rpIys[0]<=2019?'2019 & Prior':String(rpIys[0]))+' to '+rpIys[rpIys.length-1]):'--';
  var ccf=a.ceding_comm_front||{};
  var ccfStr=Object.entries(ccf).sort(function(a,b){return a[0]-b[0];}).filter(function(e){return e[1];}).map(function(e){return e[0]+': $'+e[1].toFixed(0)+'M';}).join(' | ')||'None';
  var strain=(mn.max_neg_cum_de||0)-(mp.max_neg_cum_de||0);
  var rbcLiftVec=yrs.map(function(y){var o=rr(rbco,y),n=rr(rbn,y);return o&&n?((n-o>=0?'+':'')+((n-o).toFixed(2))+'x'):'-';});
  var today=new Date().toLocaleDateString('en-US',{month:'long',day:'numeric',year:'numeric'});
  var disc=((a.discount_rate||0.08)*100).toFixed(1);
  function mkRow(stmt,key,lbl,isSub){
    return '<tr'+(isSub?' class="sub"':'')+'><td>'+lbl+'</td>'+yrs.map(function(y){return '<td>'+fmtNum(gv(stmt[key]||{},y))+'</td>';}).join('')+'</tr>';
  }
  function incomeTbl(stmt,label,stmtKey){
    var pairs=[['Total Revenue','revenue'],['Total Benefits','benefits'],
      ['Selling Expense','selling_expense'],['Op Expense','op_expense'],
      ['Pre-Tax Income','pretax_income'],['Distrib. Earnings','distributable_earnings']];
    var subs=new Set(['Total Revenue','Pre-Tax Income']);
    // RBC ratio row from formulaic results
    var rbcRow='';
    if(stmtKey==='predeal'||stmtKey==='net'){
      var adjDict={};
      if(stmtKey==='predeal')adjDict=(out.rbc_predeal_result&&out.rbc_predeal_result.predeal_adjustments)||{};
      else if(stmtKey==='net')adjDict=(out.rbc_net_result&&out.rbc_net_result.net_adjustments)||{};
      // Fallback to original computed for predeal if predeal_result unavailable
      var origVals=(out.rbc_orig_computed&&out.rbc_orig_computed.original_values)||{};
      function getRatio(y){
        var a=adjDict[y]||adjDict[String(y)];
        if(a&&a.ratio_w_margin!=null)return a.ratio_w_margin;
        var o=origVals[y]||origVals[String(y)];
        if(o&&o.ratio_w_margin!=null)return o.ratio_w_margin;
        return null;
      }
      rbcRow='<tr class="ratio"><td>RBC Ratio (w/ Margin)</td>'+
        yrs.map(function(y){var r=getRatio(y);return '<td>'+(r!=null?Number(r).toFixed(2)+'x':'-')+'</td>';}).join('')+'</tr>';
    }
    return '<div style="margin-bottom:10px"><div style="font-size:.76rem;font-weight:bold;color:var(--navy);padding:3px 8px;background:var(--off);border-radius:3px;margin-bottom:3px">'+label+'</div>'+
      '<div class="tw"><table class="bbt" style="font-size:.69rem"><thead><tr><th>Line</th>'+yrs.map(function(y){return '<th>'+y+'</th>';}).join('')+'</tr></thead><tbody>'+
      pairs.map(function(p){return mkRow(stmt,p[1],p[0],subs.has(p[0]));}).join('')+
      rbcRow+
      '</tbody></table></div></div>';
  }
  var rbcOrigSurp=(rbco['Total Surplus']||{})[fy]||(rbco['Total Surplus']||{})[String(fy)]||null;
  var rbcNetSurp=(rbn['Total Surplus']||{})[fy]||(rbn['Total Surplus']||{})[String(fy)]||null;
  var rbcOrigTAC=(rbco['TAC (C&S + AVR)']||{})[fy]||(rbco['TAC (C&S + AVR)']||{})[String(fy)]||null;
  var rbcNetTAC=(rbn['TAC (C&S + AVR)']||{})[fy]||(rbn['TAC (C&S + AVR)']||{})[String(fy)]||null;
  div.innerHTML=
    '<div style="max-width:880px;font-family:var(--ft);color:var(--dk)">'+
    cedantCard(out)+
    '<div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:12px;padding-bottom:8px;border-bottom:3px solid var(--navy)">'+
      '<div><div style="font-size:1.2rem;font-weight:bold;color:var(--navy)">REINSURANCE QUOTA-SHARE DEAL SUMMARY</div>'+
      '<div style="font-size:.72rem;color:var(--mu);margin-top:2px">EXECUTIVE BRIEFING — '+today+'</div></div>'+
      '<div style="font-size:.7rem;color:var(--mu);text-align:right"><div>Discount: '+disc+'%</div><div>Base year: '+(by-1)+'</div></div></div>'+
    '<div style="margin-bottom:12px">'+
      '<div style="background:var(--off);border-radius:var(--r);padding:10px 14px;font-size:.74rem">'+
        '<div style="font-weight:bold;color:var(--navy);margin-bottom:5px">Deal Structure</div>'+
        '<div><span style="color:var(--mu)">Treaty:</span> <b>Quota-Share</b></div>'+
        '<div><span style="color:var(--mu)">LOB:</span> <b>Medicare Supplement</b></div>'+
        '<div><span style="color:var(--mu)">Issue Year Scope:</span> <b>'+rpRange+'</b></div>'+
        '<div><span style="color:var(--mu)">Ceding %:</span> <b>'+rpPct+'</b></div>'+
        '<div><span style="color:var(--mu)">Front-End Comm:</span> <b>'+ccfStr+'</b></div>'+
        '<div><span style="color:var(--mu)">Ongoing Comm:</span> <b>$'+(a.ceding_comm_ongoing||200).toLocaleString()+'/policy/yr</b></div>'+
      '</div>'+
    '</div>'+
    '<div style="margin-bottom:12px">'+
      '<div style="font-size:.8rem;font-weight:bold;color:var(--navy);margin-bottom:5px">5-Year Income Statement ($M)</div>'+
      incomeTbl(ap,'Direct Carrier — Predeal','predeal')+
      incomeTbl(ac,'Direct Carrier — Ceded','ceded')+
      incomeTbl(an,'Direct Carrier — Net','net')+
    '</div>'+
    
    '<div style="font-size:.61rem;color:var(--mu);border-top:1px solid var(--bdr);padding-top:5px">'+
      'Source: EV Data (EV_Data_template.csv) | Discount: '+disc+'% | Updates on each model run'+
    '</div></div>';
}

// ── TEST LOG MODULE ──
function getDefaultTestCases(){
  return [
    {id:'t001',name:'Claim scalar 1.1 increases claims ~10%',type:'scalar',param:'claim_scalar',value:1.1,target_line:'claims',expected_change:-0.10,tolerance:0.005,check_unchanged:['premium','nii','commissions','prem_tax'],enabled:true},
    {id:'t002',name:'Claim scalar 0.9 decreases claims ~10%',type:'scalar',param:'claim_scalar',value:0.9,target_line:'claims',expected_change:0.10,tolerance:0.005,check_unchanged:['premium','nii'],enabled:true},
    {id:'t003',name:'Acq expense +20% only changes acq expense',type:'expense',param:'acq_exp',multiplier:1.2,target_line:'acq_expense',expected_change:-0.20,tolerance:0.005,check_unchanged:['premium','claims','nii'],enabled:true},
    {id:'t004',name:'Maint expense +15% only changes maint expense',type:'expense',param:'maint_exp',multiplier:1.15,target_line:'maint_expense',expected_change:-0.15,tolerance:0.005,check_unchanged:['premium','claims','nii'],enabled:true},
    {id:'t005',name:'Zero ceding gives zero ceded premium',type:'reins_zero',target_line:'premium',expected_ceded:0,tolerance:1000,enabled:true},
    {id:'t006',name:'Ceding active: ceded premium > 0 and < predeal',type:'reins_direction',target_line:'premium',enabled:true},
    {id:'t007',name:'Lower discount rate increases predeal PVDE magnitude',type:'direction',param:'discount_rate',to_val:0.06,metric:'pvde',direction:'decrease',target_stmt:'predeal',enabled:true},
    {id:'t008',name:'Higher maint allowance increases ceded maint expense',type:'direction',param:'maint_exp_allowance',multiplier:1.5,metric:'maint_expense',direction:'decrease',target_stmt:'ceded',enabled:true},
    {id:'t009',name:'Sales scalar 1.1 for IssYr 2026 increases premium ~10% for that IssYr',type:'sales_scalar',iy:2026,scalar:1.1,target_line:'premium',enabled:true},
  ];
}

async function runAllTests(){
  if(!S.py||!S.evData){alert('Load the model first.');return;}
  var btn=document.getElementById('run-tests-btn');
  btn.disabled=true;btn.textContent='Running...';
  S.testResults=[];
  var cases=(S.testCases||[]).filter(function(t){return t.enabled;});
  document.getElementById('test-log-output').innerHTML='<div style="color:var(--mu);font-size:.75rem">Running '+cases.length+' tests...</div>';
  for(var i=0;i<cases.length;i++){
    var r=await runOneTest(cases[i]);
    S.testResults.push(r);
  }
  renderTestLog();
  btn.disabled=false;btn.textContent='Run Tests';
}

async function runOneTest(tc){
  try{
    var baseOut=S.out;
    if(!baseOut)return{id:tc.id,name:tc.name,pass:false,error:'No baseline',evidence:{}};
    var modAssum=JSON.parse(JSON.stringify(S.assumptions||{}));
    if(tc.type==='scalar'&&(tc.param==='claim_scalar'||tc.param==='lapse_scalar')){
      modAssum[tc.param]=tc.value||1.1;
    } else if(tc.type==='expense'){
      var ann=modAssum.annual||{};
      Object.keys(ann).forEach(function(y){if(ann[y]&&ann[y][tc.param]!=null)ann[y][tc.param]*=(tc.multiplier||1.2);});
      // Rebuild top-level so engine receives updated values
      modAssum[tc.param]={};
      Object.keys(ann).forEach(function(y){if(ann[y]&&ann[y][tc.param]!=null)modAssum[tc.param][parseInt(y)]=ann[y][tc.param];});
    } else if(tc.type==='direction'){
      if(tc.param==='discount_rate'){modAssum.discount_rate=tc.to_val||0.06;}
      else if(tc.param==='acq_exp_allowance'||tc.param==='maint_exp_allowance'){
        var ann2=modAssum.annual||{};var ak=tc.param;
        Object.keys(ann2).forEach(function(y){if(ann2[y]&&ann2[y][ak]!=null)ann2[y][ak]*=(tc.multiplier||1.5);});
        // Rebuild top-level so engine gets updated values
        modAssum[ak]={};
        Object.keys(ann2).forEach(function(y){if(ann2[y]&&ann2[y][ak]!=null)modAssum[ak][parseInt(y)]=ann2[y][ak];});
      }
    } else if(tc.type==='sales_scalar'){
      modAssum.sales_scalar=modAssum.sales_scalar||{};
      modAssum.sales_scalar[tc.iy]=tc.scalar;
    } else if(tc.type==='reins_zero'){
      modAssum.reins_pct_decimal={};
    }
    var modOut=await runModelWith(modAssum);
    var by=baseOut.first_yr||2026;
    var evidence={};var pass=true;var detail='';
    if(tc.type==='scalar'||tc.type==='expense'){
      var bv=gv(baseOut.annual_predeal[tc.target_line]||{},by);
      var mv=gv(modOut.annual_predeal[tc.target_line]||{},by);
      var ch=bv!==0?(mv-bv)/Math.abs(bv):null;
      // Per-year evidence for target line
      var byrs=(baseOut.annual_predeal.periods||[]).filter(function(y){return y>=by;}).slice(0,5);
      var targetEvidence={};
      var ptiEvidence={};
      byrs.forEach(function(y){
        var bvy=gv(baseOut.annual_predeal[tc.target_line]||{},y);
        var mvy=gv(modOut.annual_predeal[tc.target_line]||{},y);
        var bpti=gv(baseOut.annual_predeal.pretax_income||{},y);
        var mpti=gv(modOut.annual_predeal.pretax_income||{},y);
        targetEvidence[y]={base:(bvy/1e6).toFixed(3),mod:(mvy/1e6).toFixed(3),diff_M:((bvy-mvy)/1e6).toFixed(3),pct:bvy?((mvy-bvy)/Math.abs(bvy)*100).toFixed(2)+'%':'n/a'};
        ptiEvidence[y]={base:(bpti/1e6).toFixed(3),mod:(mpti/1e6).toFixed(3),diff_M:((bpti-mpti)/1e6).toFixed(3)};
      });
      evidence[tc.target_line+'_by_year']=targetEvidence;
      evidence.pti_diff_by_year=ptiEvidence;
      evidence.base_yr1=(bv/1e6).toFixed(3);evidence.mod_yr1=(mv/1e6).toFixed(3);evidence.actual_pct=ch;evidence.expected_pct=tc.expected_change;
      if(ch==null||Math.abs(ch-tc.expected_change)>tc.tolerance){pass=false;detail='Change '+(ch!=null?(ch*100).toFixed(2)+'%':'null')+' expected '+(tc.expected_change*100).toFixed(1)+'%';}
      (tc.check_unchanged||[]).forEach(function(line){
        var bvl=gv(baseOut.annual_predeal[line]||{},by);var mvl=gv(modOut.annual_predeal[line]||{},by);
        var c2=bvl!==0?Math.abs((mvl-bvl)/bvl):0;evidence['unchanged_'+line+'_pct']=(c2*100).toFixed(4)+'%';
        if(c2>0.001){pass=false;detail=(detail?detail+'; ':'')+line+' changed '+((c2*100).toFixed(2))+'%';}
      });
    } else if(tc.type==='reins_zero'){
      var byrs5=(baseOut.annual_predeal.periods||[]).filter(function(y){return y>=by;}).slice(0,5);
      evidence.ceded_premium_by_year={};
      byrs5.forEach(function(y){
        evidence.ceded_premium_by_year[y]={
          ceded:(gv(modOut.annual_ceded[tc.target_line]||{},y)/1e6).toFixed(3),
          predeal:(gv(modOut.annual_predeal[tc.target_line]||{},y)/1e6).toFixed(3)
        };
      });
      var cv=Math.abs(gv(modOut.annual_ceded[tc.target_line]||{},by));
      evidence.ceded_yr1_value=cv.toFixed(0);
      if(cv>tc.tolerance){pass=false;detail='Ceded='+cv.toFixed(0)+' expected ~0';}
    } else if(tc.type==='reins_check'){
      var pred=gv(baseOut.annual_predeal[tc.target_line]||{},by);
      var ced=gv(baseOut.annual_ceded[tc.target_line]||{},by);
      var ar=pred!==0?Math.abs(ced)/Math.abs(pred):null;
      evidence.pred=pred;evidence.ced=ced;evidence.rate=ar;
      if(ar==null||Math.abs(ar-tc.expected_rate)>tc.tolerance){pass=false;detail='Rate '+(ar!=null?(ar*100).toFixed(1)+'%':'null')+' expected '+(tc.expected_rate*100).toFixed(0)+'%';}
    } else if(tc.type==='sales_scalar'){
      // Test: IssYr premium should scale by ~tc.scalar, total premium by less
      var by2=tc.iy+1;  // First full projection year for this IssYr
      // Use month 13 (Jan of following year) which is first full year for IssYr
      // Approximation: compare annual predeal premium base vs mod
      var bvss=gv(baseOut.annual_predeal[tc.target_line]||{},by);
      var mvss=gv(modOut.annual_predeal[tc.target_line]||{},by);
      var ch_ss=bvss!==0?(mvss-bvss)/Math.abs(bvss):null;
      evidence.base_total_premium=(bvss/1e6).toFixed(3);
      evidence.mod_total_premium=(mvss/1e6).toFixed(3);
      evidence.total_pct_change=ch_ss!=null?(ch_ss*100).toFixed(3)+'%':'null';
      evidence.note='IssYr '+tc.iy+' scaled by '+tc.scalar+'. Total premium should increase.';
      // Show premium by year
      var byrss=(baseOut.annual_predeal.periods||[]).filter(function(y){return y>=by;}).slice(0,5);
      evidence.premium_by_year={};
      byrss.forEach(function(y){
        evidence.premium_by_year[y]={base:(gv(baseOut.annual_predeal.premium||{},y)/1e6).toFixed(3),
          mod:(gv(modOut.annual_predeal.premium||{},y)/1e6).toFixed(3)};
      });
      if(ch_ss===null||ch_ss<=0){pass=false;detail='Premium did not increase: '+(ch_ss!=null?(ch_ss*100).toFixed(2)+'%':'null');}
    } else if(tc.type==='reins_direction'){
      var pred2=gv(baseOut.annual_predeal[tc.target_line]||{},by);
      var ced2=gv(baseOut.annual_ceded[tc.target_line]||{},by);
      evidence.pred=pred2;evidence.ced=ced2;
      evidence.note='Ceded should be nonzero and same sign as predeal';
      var byrs6=(baseOut.annual_predeal.periods||[]).filter(function(y){return y>=by;}).slice(0,5);
      evidence.ceded_vs_predeal={};
      byrs6.forEach(function(y){
        var pred6=gv(baseOut.annual_predeal[tc.target_line]||{},y);
        var ced6=gv(baseOut.annual_ceded[tc.target_line]||{},y);
        evidence.ceded_vs_predeal[y]={predeal:(pred6/1e6).toFixed(3),ceded:(ced6/1e6).toFixed(3),
          pct:pred6?(Math.abs(ced6/pred6)*100).toFixed(1)+'%':'-'};
      });
      var ced_nonzero=Math.abs(ced2)>0;
      var same_sign=(pred2>0&&ced2>0)||(pred2<0&&ced2<0);
      var ced_lt_pred=Math.abs(ced2)<Math.abs(pred2);
      if(!ced_nonzero||!same_sign||!ced_lt_pred){pass=false;detail='Ceding not active or ceded>=predeal';}
    } else if(tc.type==='direction'){
      var stmt=tc.target_stmt==='ceded'?'annual_ceded':'annual_predeal';
      var bv2=gv(baseOut[stmt][tc.metric]||{},by);var mv2=gv(modOut[stmt][tc.metric]||{},by);
      evidence.base_yr1=(bv2/1e6).toFixed(3);evidence.mod_yr1=(mv2/1e6).toFixed(3);
      // Per-year evidence
      var byrsd=(baseOut[stmt]&&baseOut[stmt].periods||[]).filter(function(y){return y>=by;}).slice(0,5);
      evidence.by_year={};
      byrsd.forEach(function(y){
        var bvd=gv(baseOut[stmt][tc.metric]||{},y);var mvd=gv(modOut[stmt][tc.metric]||{},y);
        evidence.by_year[y]={base:(bvd/1e6).toFixed(3),mod:(mvd/1e6).toFixed(3),diff:((mvd-bvd)/1e6).toFixed(3)};
      });
      var inc=mv2>bv2;
      if(tc.direction==='increase'&&!inc){pass=false;detail='Did not increase: '+bv2.toFixed(0)+' -> '+mv2.toFixed(0);}
      if(tc.direction==='decrease'&&inc){pass=false;detail='Did not decrease: '+bv2.toFixed(0)+' -> '+mv2.toFixed(0);}
    }
    return{id:tc.id,name:tc.name,pass:pass,detail:detail,evidence:evidence,ts:new Date().toISOString()};
  }catch(e){return{id:tc.id,name:tc.name,pass:false,error:e.message,evidence:{}};}
}

async function runModelWith(assum){
  var a=Object.assign({},assum);
  var apython=Object.assign({},a,{
    reins_pct:a.reins_pct_decimal||{},
    ceding_comm_front:a.ceding_comm_front_dollars||{},
    sales_scalar:a.sales_scalar||{}
  });
  var ev_for_py={agg:S.evData.agg,agg_iy:S.evData.agg_iy,periods:S.evData.periods,iss_years:S.evData.iss_years,row_count:S.evData.row_count};
  S.py.globals.set('_evj',JSON.stringify(ev_for_py));
  S.py.globals.set('_asj',JSON.stringify(apython));
  S.py.globals.set('_by',a.base_year||2025);
  var code=[
    "import json","ev_agg=json.loads(_evj)","assum=json.loads(_asj)",
    "for k in ['acq_exp','maint_exp','nier','ceding_comm_front','acq_exp_allowance','maint_exp_allowance']:",
    "    if k in assum and isinstance(assum[k],dict):",
    "        assum[k]={int(kk):vv for kk,vv in assum[k].items()}",
    "if 'reins_pct' in assum:",
    "    assum['reins_pct']={int(iy):{int(cy):v for cy,v in yd.items()} for iy,yd in assum['reins_pct'].items()}",
    "ev_agg['agg']={k:{int(p):v for p,v in pv.items()} for k,pv in ev_agg['agg'].items()}",
    "ev_agg['agg_iy']={str(iy):{k:{int(p):v for p,v in pv.items()} for k,pv in vm.items()} for iy,vm in ev_agg['agg_iy'].items()}",
    "ev_agg['periods']=[int(p) for p in ev_agg['periods']]",
    "result=run_model(ev_agg,assum,int(_by),_rbc_rows if '_rbc_rows' in dir() else None,_bp_rows if '_bp_rows' in dir() else None,_surplus_rows if '_surplus_rows' in dir() else None)",
    "json.dumps(result)"
  ].join("\n");
  var rj=await S.py.runPythonAsync(code);
  var out=JSON.parse(rj);
  out.first_yr=(out.annual_predeal&&out.annual_predeal.periods&&out.annual_predeal.periods[0])||2026;
  return out;
}

function renderTestCases(){
  var div=document.getElementById('test-cases-list');if(!div)return;
  var cases=S.testCases||[];
  div.innerHTML=cases.map(function(tc,i){
    return '<div style="display:flex;align-items:center;gap:6px;padding:4px 6px;border-radius:3px;background:var(--off);margin-bottom:3px">'+
      '<input type="checkbox"'+(tc.enabled?' checked':'')+' onchange="S.testCases['+i+'].enabled=this.checked" style="cursor:pointer">'+
      '<span style="flex:1;font-size:.72rem">'+tc.name+'</span>'+
      '<span style="font-size:.63rem;color:var(--mu)">'+tc.type+'</span>'+
      '<button onclick="removeTestCase('+i+')" style="color:var(--err);background:none;border:none;cursor:pointer;font-size:.68rem">x</button></div>';
  }).join('');
}

function removeTestCase(i){S.testCases.splice(i,1);renderTestCases();}

function addNLTestCase(){
  var inp=document.getElementById('nl-test-input');
  var text=inp.value.trim();if(!text)return;
  var tc=parseNLTest(text);
  if(!tc){alert('Could not interpret. Try: "claim scalar 1.1 increases claims by 10%" or "acq expense +20% only changes acq expense"');return;}
  if(!S.testCases)S.testCases=[];
  S.testCases.push(tc);renderTestCases();
  localStorage.setItem('bb_tc',JSON.stringify(S.testCases));inp.value='';
}

function parseNLTest(text){
  var t=text.toLowerCase();var id='tc_'+Date.now();
  var m;
  m=t.match(/claim.*scalar.*?([\d.]+)/);
  if(m){var v=parseFloat(m[1])||1.1;return{id:id,name:text,type:'scalar',param:'claim_scalar',value:v,target_line:'claims',expected_change:-(v-1),tolerance:0.01,check_unchanged:['premium','nii','commissions'],enabled:true};}
  m=t.match(/acq.*?(increase|decrease|[+\-]).*?([\d.]+)%/);
  if(m){var pct=parseFloat(m[2])/100;var mult=t.includes('decrease')||t.includes('-')?1-pct:1+pct;return{id:id,name:text,type:'expense',param:'acq_exp',multiplier:mult,target_line:'acq_expense',expected_change:t.includes('decrease')?pct:-pct,tolerance:0.01,check_unchanged:['premium','claims','nii'],enabled:true};}
  m=t.match(/maint.*?(increase|decrease|[+\-]).*?([\d.]+)%/);
  if(m){var pct2=parseFloat(m[2])/100;var mult2=t.includes('decrease')||t.includes('-')?1-pct2:1+pct2;return{id:id,name:text,type:'expense',param:'maint_exp',multiplier:mult2,target_line:'maint_expense',expected_change:t.includes('decrease')?pct2:-pct2,tolerance:0.01,check_unchanged:['premium','claims','nii'],enabled:true};}
  if(t.includes('zero ceding')||t.includes('no reinsurance')){return{id:id,name:text,type:'reins_zero',target_line:'premium',expected_ceded:0,tolerance:1000,enabled:true};}
  return null;
}

function buildEvidenceHtml(ev,testName){
  if(!ev||!Object.keys(ev).length)return '';
  var lines=[];
  // Part 1: per-year target line diff vs PTI diff table
  var lineKey=Object.keys(ev).find(function(k){return k.endsWith('_by_year')&&k!=='pti_diff_by_year';});
  var ptiKey='pti_diff_by_year';
  if(lineKey&&ev[lineKey]&&ev[ptiKey]){
    var lineName=lineKey.replace('_by_year','').replace(/_/g,' ');
    var yrs=Object.keys(ev[lineKey]).sort();
    var lineDat=ev[lineKey]; var ptiDat=ev[ptiKey];
    var tbl='<div style="margin-top:6px;font-size:.67rem">';
    tbl+='<div style="font-weight:bold;color:var(--mu);margin-bottom:3px">Evidence Part 1: '+lineName+' change vs PTI change by year ($M)</div>';
    tbl+='<div class="tw"><table class="bbt" style="font-size:.65rem">';
    tbl+='<thead><tr><th>Metric</th>'+yrs.map(function(y){return '<th>'+y+'</th>';}).join('')+'</tr></thead><tbody>';
    tbl+='<tr><td>'+lineName+' — Base</td>'+yrs.map(function(y){return '<td>'+((lineDat[y]&&lineDat[y].base)||'-')+'</td>';}).join('')+'</tr>';
    tbl+='<tr><td>'+lineName+' — Test</td>'+yrs.map(function(y){return '<td>'+((lineDat[y]&&lineDat[y].mod)||'-')+'</td>';}).join('')+'</tr>';
    tbl+='<tr class="sub"><td>Base - Test ('+lineName+')</td>'+yrs.map(function(y){return '<td>'+((lineDat[y]&&lineDat[y].diff_M)||'-')+'</td>';}).join('')+'</tr>';
    tbl+='<tr><td>PTI — Base</td>'+yrs.map(function(y){return '<td>'+((ptiDat[y]&&ptiDat[y].base)||'-')+'</td>';}).join('')+'</tr>';
    tbl+='<tr><td>PTI — Test</td>'+yrs.map(function(y){return '<td>'+((ptiDat[y]&&ptiDat[y].mod)||'-')+'</td>';}).join('')+'</tr>';
    tbl+='<tr class="sub"><td>PTI Base - Test</td>'+yrs.map(function(y){return '<td>'+((ptiDat[y]&&ptiDat[y].diff_M)||'-')+'</td>';}).join('')+'</tr>';
    tbl+='</tbody></table></div>';
    // Part 2: % change table
    tbl+='<div style="font-weight:bold;color:var(--mu);margin:6px 0 3px">Evidence Part 2: '+lineName+' % change by year</div>';
    tbl+='<div class="tw"><table class="bbt" style="font-size:.65rem">';
    tbl+='<thead><tr><th></th>'+yrs.map(function(y){return '<th>'+y+'</th>';}).join('')+'</tr></thead><tbody>';
    tbl+='<tr><td>Base ($M)</td>'+yrs.map(function(y){return '<td>'+((lineDat[y]&&lineDat[y].base)||'-')+'</td>';}).join('')+'</tr>';
    tbl+='<tr><td>Test ($M)</td>'+yrs.map(function(y){return '<td>'+((lineDat[y]&&lineDat[y].mod)||'-')+'</td>';}).join('')+'</tr>';
    tbl+='<tr class="sub"><td>% Change</td>'+yrs.map(function(y){return '<td>'+((lineDat[y]&&lineDat[y].pct)||'-')+'</td>';}).join('')+'</tr>';
    tbl+='</tbody></table></div></div>';
    lines.push(tbl);
  }
  // Unchanged lines check
  var unchKeys=Object.keys(ev).filter(function(k){return k.startsWith('unchanged_');});
  if(unchKeys.length){
    var uchRow='<div style="margin-top:4px;font-size:.67rem;color:var(--mu)">Unchanged lines: '+
      unchKeys.map(function(k){return k.replace('unchanged_','').replace('_pct','')+'='+ev[k];}).join(', ')+'</div>';
    lines.push(uchRow);
  }
  // Render by_year / ceded_vs_predeal / ceded_premium_by_year / premium_by_year tables
  var tableEvKeys=['by_year','ceded_vs_predeal','ceded_premium_by_year','premium_by_year'];
  tableEvKeys.forEach(function(tkey){
    if(!ev[tkey])return;
    var yrs=Object.keys(ev[tkey]).sort();
    if(!yrs.length)return;
    var cols=Object.keys(ev[tkey][yrs[0]]);if(!cols.length)return;
    var tbl='<div style="margin-top:6px;font-size:.67rem"><div style="font-weight:bold;color:var(--mu);margin-bottom:2px">'+tkey.replace(/_/g,' ')+'</div>';
    tbl+='<div class="tw"><table class="bbt" style="font-size:.64rem"><thead><tr><th>Year</th>'+cols.map(function(c){return '<th>'+c+'</th>';}).join('')+'</tr></thead><tbody>';
    yrs.forEach(function(y){
      tbl+='<tr><td>'+y+'</td>'+cols.map(function(c){return '<td>'+ev[tkey][y][c]+'</td>';}).join('')+'</tr>';
    });
    tbl+='</tbody></table></div></div>';
    lines.push(tbl);
  });
  // Show simple key/value evidence if no structured data
  var simpleKeys=Object.keys(ev).filter(function(k){return typeof ev[k]==='string'||typeof ev[k]==='number';}).slice(0,8);
  if(simpleKeys.length&&!lineKey){
    var kv='<div style="margin-top:4px;font-size:.67rem;color:var(--mu)">'+
      simpleKeys.map(function(k){return '<b>'+k+'</b>: '+ev[k];}).join(' &nbsp;|&nbsp; ')+'</div>';
    lines.push(kv);
  }
  if(!lines.length)return '';
  return '<details style="margin-top:4px"><summary style="cursor:pointer;font-size:.67rem;color:var(--mu)">Show Evidence</summary>'+lines.join('')+'</details>';
}

function renderTestLog(){
  var div=document.getElementById('test-log-output');if(!div)return;
  var results=S.testResults||[];
  if(!results.length){div.innerHTML='<div style="color:var(--mu);font-size:.75rem">No results yet. Click Run Tests.</div>';return;}
  var pass=results.filter(function(r){return r.pass;}).length;
  var fail=results.length-pass;
  var h='<div style="margin-bottom:10px;padding:6px 10px;border-radius:var(--r);background:'+(fail?'#FEF0EE':'#D4EDDA')+';border:1px solid '+(fail?'#FACACA':'#A8D5B5')+';font-size:.76rem;font-weight:bold;color:'+(fail?'#721c24':'#155724')+'">'+pass+' passed, '+fail+' failed &nbsp; '+new Date().toLocaleTimeString()+'</div>';
  h+=results.map(function(r){
    return '<div style="margin-bottom:6px;padding:6px 10px;border-radius:var(--r);background:var(--off);border-left:3px solid '+(r.pass?'#28a745':'#C0392B')+'">'+
      '<div style="display:flex;align-items:center;gap:6px"><span style="font-weight:bold;font-size:.74rem">'+(r.pass?'&#10003;':'&#10007;')+' '+r.name+'</span>'+
      (r.pass?'':'<span style="color:#C0392B;font-size:.68rem">'+r.detail+'</span>')+
      '</div>'+
      (buildEvidenceHtml(r.evidence,r.name))+
      '</div>';
  }).join('');
  div.innerHTML=h;
}



// =========================================================
// SLIDING-SCALE CEDING COMMISSION TABLE (UI helpers)
// =========================================================
function initCCOTable(rows){
  var tb=document.getElementById('cco-tbody');
  if(!tb)return;
  tb.innerHTML='';
  (rows||[[0,0.75,250],[0.75,0.85,200],[0.85,0.95,150],[0.95,'Inf',100]]).forEach(function(r){
    tb.appendChild(ccoMakeRow(r[0],r[1],r[2]));
  });
}
function ccoMakeRow(lo,hi,comm){
  var tr=document.createElement('tr');
  var hiVal=(hi==='Inf'||hi===Infinity||Number(hi)>=999)?'':hi;
  tr.innerHTML='<td><input type="number" step="0.01" min="0" max="1" value="'+lo+'" style="width:60px;border:1px solid #ccc;padding:2px 4px;font-size:.72rem;border-radius:3px"></td>'+
    '<td><input type="text" value="'+hiVal+'" placeholder="Inf" style="width:60px;border:1px solid #ccc;padding:2px 4px;font-size:.72rem;border-radius:3px"></td>'+
    '<td><input type="number" step="1" min="0" value="'+comm+'" style="width:70px;border:1px solid #ccc;padding:2px 4px;font-size:.72rem;border-radius:3px"></td>'+
    '<td><button onclick="this.closest(\'tr\').remove()" style="font-size:.65rem;padding:1px 5px;cursor:pointer;color:#C0392B;background:none;border:none">&#x2715;</button></td>';
  return tr;
}
function ccoAddRow(){
  var tb=document.getElementById('cco-tbody');
  if(tb)tb.appendChild(ccoMakeRow(0.95,'Inf',100));
}
function readCCOTable(){
  var rows=[];
  document.querySelectorAll('#cco-tbody tr').forEach(function(tr){
    var ins=tr.querySelectorAll('input');
    if(ins.length>=3){
      var lo=parseFloat(ins[0].value)||0;
      var hiRaw=ins[1].value.trim();
      var hi=(hiRaw===''||hiRaw.toLowerCase()==='inf'||hiRaw==='\u221e')?'Inf':parseFloat(hiRaw);
      var comm=parseFloat(ins[2].value)||0;
      rows.push([lo,hi,comm]);
    }
  });
  return rows.length?rows:[[0,0.75,250],[0.75,0.85,200],[0.85,0.95,150],[0.95,'Inf',100]];
}

// =========================================================
// STRESS TEST UI
// =========================================================
function d3f(v,dec){return isNaN(v)||v==null?'\u2014':Number(v).toFixed(dec==null?1:dec);}

function buildHistSVG2(data,label,sfx){
  sfx=sfx||'';
  var mn=Math.min.apply(null,data),mx=Math.max.apply(null,data);
  var nbins=Math.min(20,Math.max(8,Math.round(data.length/4)));
  var bw=(mx-mn)/nbins||1;
  var counts=new Array(nbins).fill(0);
  data.forEach(function(v){counts[Math.min(nbins-1,Math.floor((v-mn)/bw))]++;});
  var maxC=Math.max.apply(null,counts)||1;
  var W=280,H=120,pl=8,pr=8,pt=14,pb=28;
  var bwPx=(W-pl-pr)/nbins;
  var meanV=data.reduce(function(a,b){return a+b;},0)/data.length;
  var meanX=pl+(meanV-mn)/((mx-mn)||1)*(W-pl-pr);
  var svg='<div><div style="font-size:.7rem;font-weight:600;color:var(--navy);text-align:center;margin-bottom:2px">'+label+'</div>'+
    '<svg width="'+W+'" height="'+H+'" style="overflow:visible">';
  counts.forEach(function(c,i){
    var x=pl+i*bwPx,h=(c/maxC)*(H-pt-pb),y=H-pb-h;
    var pFill=i/nbins;
    var r=Math.round(0+pFill*192),gb=Math.round(35+pFill*20);
    svg+='<rect x="'+x.toFixed(1)+'" y="'+y.toFixed(1)+'" width="'+(bwPx-1).toFixed(1)+'" height="'+h.toFixed(1)+'" fill="rgb('+r+','+gb+',102)" opacity="0.8"/>';
  });
  svg+='<line x1="'+pl+'" y1="'+(H-pb)+'" x2="'+(W-pr)+'" y2="'+(H-pb)+'" stroke="#aaa" stroke-width="0.8"/>';
  svg+='<line x1="'+meanX.toFixed(1)+'" y1="'+pt+'" x2="'+meanX.toFixed(1)+'" y2="'+(H-pb)+'" stroke="#E74C3C" stroke-width="1.5" stroke-dasharray="3,2"/>';
  svg+='<text x="'+meanX.toFixed(1)+'" y="'+(pt+1)+'" font-size="8" fill="#E74C3C" text-anchor="middle" dominant-baseline="hanging">\u03bc='+meanV.toFixed(2)+sfx+'</text>';
  svg+='<text x="'+pl+'" y="'+(H-pb+10)+'" font-size="7" fill="#888">'+mn.toFixed(2)+sfx+'</text>';
  svg+='<text x="'+(W-pr)+'" y="'+(H-pb+10)+'" font-size="7" fill="#888" text-anchor="end">'+mx.toFixed(2)+sfx+'</text>';
  svg+='</svg></div>';
  return svg;
}

// =========================================================
// REINSURANCE SCENARIOS MODULE
// =========================================================
const REIN_SCENARIOS = [{"id": "base", "name": "Base Case", "desc": "Full deal as negotiated \u2014 10% cede, IssYrs 2019-2030, $20M front-end commission, tiered ongoing CC", "changes": {}}, {"id": "no_frontend", "name": "No Front-End Commission", "desc": "GenRe pays $0 upfront ceding commission (vs $20M in base case)", "changes": {"ceding_comm_front": {}}}, {"id": "low_ongoing", "name": "Reduced Ongoing Commissions", "desc": "All commission tiers cut 40%: $150/$120/$90/$60 per policy per year", "changes": {"ceding_comm_table": [[0, 0.75, 150], [0.75, 0.85, 120], [0.85, 0.95, 90], [0.95, "Inf", 60]]}}, {"id": "no_2025iy", "name": "Exclude 2025 Issue Year", "desc": "GenRe excludes 2025 issue year only; all other issue years (2019-2024, 2026-2030) still ceded", "changes": {"reins_pct_exclude_iy": [2025]}}, {"id": "pre2022_only", "name": "Pre-2022 Block Only", "desc": "GenRe only accepts IssYrs 2019-2021 \u2014 very limited participation", "changes": {"reins_pct_override": [2019, 2022]}}, {"id": "half_cede", "name": "5% Cede (Half Rate)", "desc": "GenRe accepts only 5% quota share instead of negotiated 10%", "changes": {"cede_pct": 0.05}}, {"id": "short_term", "name": "3-Year Treaty", "desc": "Cession period shortened to 2026-2028 only (vs 2026-2031 in base)", "changes": {"treaty_end": 2029}}, {"id": "worst_case", "name": "Worst Case Combined", "desc": "No front-end commission, reduced ongoing CC, pre-2022 block only, 3-year term", "changes": {"ceding_comm_front": {}, "ceding_comm_table": [[0, 0.75, 150], [0.75, 0.85, 120], [0.85, 0.95, 90], [0.95, "Inf", 60]], "reins_pct_override": [2019, 2022], "treaty_end": 2029}}];
const REIN_RESULTS_CACHE = [{"id": "base", "name": "Base Case", "desc": "Full deal: 10%, IssYrs 2019-2030, $20M front-end, tiered CC", "pred_pvde": 1037.7158087448108, "net_pvde": 971.8792505626329, "ced_pvde": 65.83655818217791, "pred_irr": 0.24649912679401353, "net_irr": 0.25446981951146785, "rbc26": 5.04127042414617, "rbc29": 3.8072877572111334, "pti26": -121.10212603589761, "pti_net26": -101.55257570649579, "pti30": 76.35647022741183, "pti_net30": 68.5087762768544, "front_end": 10.0, "cede_pct": 0.1}, {"id": "no_frontend", "name": "No Front-End Commission", "desc": "GenRe pays $0 upfront vs $20M", "pred_pvde": 1037.7158087448108, "net_pvde": 964.5644357478182, "ced_pvde": 73.15137299699272, "pred_irr": 0.24649912679401353, "net_irr": 0.25011101784424383, "rbc26": 4.960107573121354, "rbc29": 3.749421437939834, "pti26": -121.10212603589761, "pti_net26": -111.5525757064958, "pti30": 76.35647022741183, "pti_net30": 68.5087762768544, "front_end": 0.0, "cede_pct": 0.1}, {"id": "low_ongoing", "name": "Reduced Ongoing Commissions", "desc": "All CC tiers cut 40%: $150/$120/$90/$60/pol", "pred_pvde": 1037.7158087448108, "net_pvde": 969.8168972049932, "ced_pvde": 67.89891153981756, "pred_irr": 0.24649912679401353, "net_irr": 0.2536319836627112, "rbc26": 5.04127042414617, "rbc29": 3.7946471228861607, "pti26": -121.10212603589761, "pti_net26": -101.55257570649579, "pti30": 76.35647022741183, "pti_net30": 67.80645910678442, "front_end": 10.0, "cede_pct": 0.1}, {"id": "no_2025iy", "name": "Exclude 2025 Issue Year", "desc": "GenRe excludes 2025 issue year only; all other IssYrs (2019-2024, 2026-2030) still ceded", "pred_pvde": 1037.7158087448108, "net_pvde": 974.2837575383231, "ced_pvde": 63.43205120648774, "pred_irr": 0.24649912679401353, "net_irr": 0.2511332683082701, "rbc26": 4.983153392750958, "rbc29": 3.7383508463378634, "pti26": -121.10212603589761, "pti_net26": -104.96262179701812, "pti30": 76.35647022741183, "pti_net30": 68.74147659178361, "front_end": 10.0, "cede_pct": 0.1}, {"id": "pre2022_only", "name": "Pre-2022 Block Only", "desc": "GenRe only takes IssYrs 2019-2021 \u2014 minimal new business", "pred_pvde": 1037.7158087448108, "net_pvde": 1023.9144121880024, "ced_pvde": 13.801396556808465, "pred_irr": 0.24649912679401353, "net_irr": 0.24649151784752066, "rbc26": 4.881465739451368, "rbc29": 3.5100243441449646, "pti26": -121.10212603589761, "pti_net26": -111.67772789794347, "pti30": 76.35647022741183, "pti_net30": 72.92312591178037, "front_end": 10.0, "cede_pct": 0.1}, {"id": "half_cede", "name": "5% Cede Percentage", "desc": "GenRe accepts only 5% quota share vs 10%", "pred_pvde": 1037.7158087448108, "net_pvde": 1008.4549370611294, "ced_pvde": 29.26087168368155, "pred_irr": 0.24649912679401353, "net_irr": 0.25229569892358983, "rbc26": 4.945521760425781, "rbc29": 3.6753356775893167, "pti26": -121.10212603589761, "pti_net26": -106.3273508711967, "pti30": 76.35647022741183, "pti_net30": 72.43262325213308, "front_end": 10.0, "cede_pct": 0.05}, {"id": "short_term", "name": "Shorter Treaty (3-Year)", "desc": "Cession period 2026-2028 only vs 2026-2031", "pred_pvde": 1037.7158087448108, "net_pvde": 997.5715087559487, "ced_pvde": 40.1442999888622, "pred_irr": 0.24649912679401353, "net_irr": 0.25626194257279966, "rbc26": 5.04127042414617, "rbc29": 3.7782641400801267, "pti26": -121.10212603589761, "pti_net26": -101.55257570649579, "pti30": 76.35647022741183, "pti_net30": 69.32110898176144, "front_end": 10.0, "cede_pct": 0.1}, {"id": "worst_case", "name": "Worst Case Combined", "desc": "No front-end, low CC, legacy block only, 3-year term", "pred_pvde": 1037.7158087448108, "net_pvde": 1016.5995973731876, "ced_pvde": 21.116211371623276, "pred_irr": 0.24649912679401353, "net_irr": 0.24272472424241975, "rbc26": 4.801573353427433, "rbc29": 3.4541086470067177, "pti26": -121.10212603589761, "pti_net26": -121.67772789794347, "pti30": 76.35647022741183, "pti_net30": 72.92312591178037, "front_end": 0.0, "cede_pct": 0.1}];

function buildReinScenAssumptions(scenChanges, baseAssum) {
  var a = JSON.parse(JSON.stringify(baseAssum));
  // Apply override: ceding_comm_front
  if(scenChanges.ceding_comm_front !== undefined) a.ceding_comm_front = scenChanges.ceding_comm_front;
  // Apply override: ceding_comm_table
  if(scenChanges.ceding_comm_table !== undefined) a.ceding_comm_table = scenChanges.ceding_comm_table;
  // Apply override: cede percentage
  if(scenChanges.cede_pct !== undefined) {
    var rp = {};
    for(var iy=2019;iy<=2030;iy++) {
      rp[iy]={};
      var start=Math.max(2026,iy+1);
      for(var cy=start;cy<=2031;cy++) rp[iy][cy]=scenChanges.cede_pct;
    }
    a.reins_pct = rp; a.reins_pct_decimal = rp;
  }
  // Apply override: reins_pct_override = [iy_start, iy_end_exclusive]
  if(scenChanges.reins_pct_override !== undefined) {
    var bounds = scenChanges.reins_pct_override;
    var rp2 = {};
    for(var iy=bounds[0];iy<bounds[1];iy++) {
      rp2[iy]={};
      var start=Math.max(2026,iy+1);
      for(var cy=start;cy<=2031;cy++) rp2[iy][cy]=0.10;
    }
    a.reins_pct = rp2; a.reins_pct_decimal = rp2;
  }
  // Apply override: reins_pct_exclude_iy = [array of IssYrs to exclude]
  if(scenChanges.reins_pct_exclude_iy !== undefined) {
    var excludeSet = new Set(scenChanges.reins_pct_exclude_iy.map(Number));
    var rp4 = {};
    for(var iy=2019;iy<=2030;iy++) {
      if(excludeSet.has(iy)) continue; // skip excluded issue years
      rp4[iy]={};
      var start=Math.max(2026,iy+1);
      for(var cy=start;cy<=2031;cy++) rp4[iy][cy]=0.10;
    }
    a.reins_pct = rp4; a.reins_pct_decimal = rp4;
  }
  // Apply override: treaty_end
  if(scenChanges.treaty_end !== undefined) {
    var rp3 = JSON.parse(JSON.stringify(a.reins_pct || a.reins_pct_decimal || {}));
    Object.keys(rp3).forEach(function(iy) {
      Object.keys(rp3[iy]).forEach(function(cy) {
        if(Number(cy) >= scenChanges.treaty_end) delete rp3[iy][cy];
      });
    });
    a.reins_pct = rp3; a.reins_pct_decimal = rp3;
  }
  return a;
}

function renderReinScenUI() {
  var el = document.getElementById('reinscen-content');
  if(!el) return;
  
  var NAVY='#002366',GREEN='#1A7A4A',RED='#C0392B',GOLD='#E8A020',MUTED='#64748B';
  function fv(v,d,sfx){ return (v==null||isNaN(v))?'—':Number(v).toFixed(d==null?1:d)+(sfx||''); }
  function fM(v){ return '$'+fv(v,0)+'M'; }
  function fx(v){ return fv(v,2)+'x'; }
  function fp(v){ return fv(v*100,1)+'%'; }
  function col(v, base, goodIfHigher) {
    if(v==null||isNaN(v)) return MUTED;
    var delta = v - (base||0);
    if(Math.abs(delta) < 0.01) return MUTED;
    return (goodIfHigher ? delta > 0 : delta < 0) ? GREEN : RED;
  }
  
  // Build summary table
  var base = REIN_RESULTS_CACHE[0];
  var metrics = [
    {k:'net_pvde',   lbl:'Net PVDE ($M)',        fmt:fM,   good:false},
    {k:'net_irr',    lbl:'Net IRR',               fmt:fp,   good:false},
    {k:'ced_pvde',   lbl:'Ceded PVDE ($M)',       fmt:fM,   good:false},
    {k:'pred_pvde',  lbl:'Predeal PVDE ($M)',     fmt:fM,   good:true},
    {k:'rbc26',      lbl:'Net RBC 2026 (x)',      fmt:fx,   good:true},
    {k:'rbc29',      lbl:'Net RBC 2029 (x)',      fmt:fx,   good:true},
  ];
  
  var tbl='<div class="tw"><table class="bbt" style="font-size:.71rem">';
  tbl+='<thead><tr><th style="min-width:180px">Scenario</th>';
  metrics.forEach(function(m){ tbl+='<th>'+m.lbl+'</th>'; });
  tbl+='<th>vs. Base</th></tr></thead><tbody>';
  
  REIN_RESULTS_CACHE.forEach(function(r,i) {
    var isBase = i===0;
    tbl+='<tr'+(isBase?' class="hdr"':'')+'>';
    tbl+='<td style="font-weight:'+(isBase?'700':'500')+'"><div style="font-size:.75rem">'+r.name+'</div>';
    tbl+='<div style="font-size:.64rem;color:'+MUTED+'">'+r.desc+'</div></td>';
    metrics.forEach(function(m) {
      var v = r[m.k]; var c = isBase ? MUTED : col(v, base[m.k], m.good);
      tbl+='<td style="color:'+c+'"><b>'+m.fmt(v)+'</b>';
      if(!isBase) {
        var delta = v - base[m.k];
        var sign = delta >= 0 ? '+' : '';
        var deltaStr = (m.k.includes('irr') ? (sign+fv(delta*100,1)+'%') : (sign+fv(delta,1)+(m.k.includes('rbc')?'x':'M')));
        tbl+='<div style="font-size:.62rem;color:'+c+'">'+deltaStr+'</div>';
      }
      tbl+='</td>';
    });
    // Net PVDE delta % vs base
    var pvdeDelta = isBase ? '' : ((r.net_pvde-base.net_pvde)/Math.abs(base.net_pvde)*100);
    tbl+='<td>'+(isBase?'—':'<span style="color:'+col(r.net_pvde,base.net_pvde,false)+'">'+fv(pvdeDelta,1)+'%</span>')+'</td>';
    tbl+='</tr>';
  });
  tbl+='</tbody></table></div>';
  
  // Comparison selector
  var compSel = '<div style="display:flex;gap:12px;align-items:center;margin-bottom:8px;flex-wrap:wrap">';
  compSel += '<span style="font-size:.78rem;font-weight:600;color:#002366">Compare:</span>';
  compSel += '<select id="rsc-left" style="font-size:.75rem;padding:3px 6px" onchange="renderReinComp()">';
  REIN_RESULTS_CACHE.forEach(function(r,i){ compSel+='<option value="'+i+'"'+((i===0)?' selected':'')+'>'+r.name+'</option>'; });
  compSel += '</select>';
  compSel += '<span style="font-size:.78rem">vs.</span>';
  compSel += '<select id="rsc-right" style="font-size:.75rem;padding:3px 6px" onchange="renderReinComp()">';
  REIN_RESULTS_CACHE.forEach(function(r,i){ compSel+='<option value="'+i+'"'+((i===7)?' selected':'')+'>'+r.name+'</option>'; });
  compSel += '</select></div>';
  compSel += '<div id="rsc-compare"></div>';
  
  el.innerHTML = '<div class="card"><div class="ch">Reinsurance Scenario Analysis — 8 Scenarios</div><div class="cb">'+
    '<p style="font-size:.74rem;color:'+MUTED+';margin-bottom:10px">Each scenario tests a less favorable term from Wellabe\'s perspective. '+
    'Predeal PVDE is unchanged (it doesn\'t depend on treaty terms) — differences are in Net PVDE, RBC lift, and ceded value.</p>'+
    tbl+'</div></div>'+
    '<div class="card"><div class="ch">Side-by-Side Comparison</div><div class="cb">'+compSel+'</div></div>';
  
  renderReinComp();
}

function renderReinComp() {
  var li = parseInt(document.getElementById('rsc-left')&&document.getElementById('rsc-left').value)||0;
  var ri = parseInt(document.getElementById('rsc-right')&&document.getElementById('rsc-right').value)||7;
  var L = REIN_RESULTS_CACHE[li]; var R = REIN_RESULTS_CACHE[ri];
  var el = document.getElementById('rsc-compare');
  if(!el||!L||!R) return;
  
  var NAVY='#002366',GREEN='#1A7A4A',RED='#C0392B',MUTED='#64748B';
  function fv(v,d,sfx){ return (v==null||isNaN(v))?'—':Number(v).toFixed(d==null?1:d)+(sfx||''); }
  
  var rows = [
    ['Net PVDE ($M)',       function(r){return '$'+fv(r.net_pvde,0)+'M';}, true],
    ['Net IRR',             function(r){return fv(r.net_irr*100,1)+'%';}, true],
    ['Predeal PVDE ($M)',   function(r){return '$'+fv(r.pred_pvde,0)+'M';}, true],
    ['Ceded PVDE ($M)',     function(r){return '$'+fv(r.ced_pvde,0)+'M';}, false],
    ['Net RBC 2026',        function(r){return fv(r.rbc26,2)+'x';}, true],
    ['Net RBC 2029',        function(r){return fv(r.rbc29,2)+'x';}, true],
    ['PTI 2026 (Net $M)',   function(r){return '$'+fv(r.pti_net26,0)+'M';}, true],
    ['PTI 2030 (Net $M)',   function(r){return '$'+fv(r.pti_net30,0)+'M';}, true],
    ['Front-End Comm',      function(r){return '$'+fv(r.front_end,0)+'M';}, true],
  ];
  
  var h='<table class="bbt" style="font-size:.75rem;max-width:640px"><thead><tr>'+
    '<th style="min-width:160px">Metric</th>'+
    '<th style="min-width:160px;color:'+NAVY+'">'+L.name+'</th>'+
    '<th style="min-width:160px;color:'+NAVY+'">'+R.name+'</th>'+
    '<th>Δ</th></tr></thead><tbody>';
  rows.forEach(function(row) {
    var lv=row[1](L), rv=row[1](R);
    var lNum=parseFloat(lv), rNum=parseFloat(rv);
    var delta=isNaN(lNum)||isNaN(rNum)?'—':(rNum-lNum>=0?'+':'')+fv(rNum-lNum,1);
    var dc=isNaN(lNum)||isNaN(rNum)?MUTED:(row[2]?(rNum>=lNum?GREEN:RED):(rNum<=lNum?GREEN:RED));
    h+='<tr><td>'+row[0]+'</td><td>'+lv+'</td><td>'+rv+'</td><td style="color:'+dc+'">'+delta+'</td></tr>';
  });
  h+='</tbody></table>';
  
  el.innerHTML=h;
}



// =========================================================
// DEAL FRONTIER MODULE
// =========================================================
function renderFrontierUI(){
  var el=document.getElementById('frontier-content');if(!el)return;
  var hint=function(t){return '<div style="font-size:.63rem;color:var(--mu);margin-top:2px">'+t+'</div>';};
  el.innerHTML='<div class="card"><div class="ch">Deal Efficient Frontier</div><div class="cb">'+
    '<div style="font-size:.72rem;color:var(--mu);margin-bottom:12px;max-width:840px">Sweeps reinsurance <b>structures</b> (cede %, upfront-commission schedule, ongoing commission level, issue-year scope, years of new business) as deterministic points, and tests each under a grid of <b>claims &times; lapse</b> sensitivities. Each structure is plotted by <b>cost</b> (Ceded PVDE handed to the reinsurer) versus <b>benefit</b> (capital relief + early-strain relief). The upper-left envelope is the efficient frontier; the faint cloud around each point is its claims/lapse sensitivity.</div>'+
    '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;max-width:660px">'+
    '<div><label class="fl">Cede % values (comma-sep)</label><input id="fl-cede" type="text" value="0.05,0.10,0.15,0.20" class="fi">'+hint('Quota-share percentages to sweep (0.10 = 10%)')+'</div>'+
    '<div><label class="fl">Years of new business (comma-sep)</label><input id="fl-dur" type="text" value="3,5" class="fi">'+hint('Cedes new-business issue years 2026..2025+N (3 = 2026-2028) alongside the in-force scope below')+'</div>'+
    '<div><label class="fl">Ongoing commission level (fraction of tiers)</label><input id="fl-cc" type="text" value="0.8,1.0" class="fi">'+hint('1.0 = full $250/$200/$150/$100 per-policy tiers; 0.8 = 80%')+'</div>'+
    '</div>'+
    '<div style="margin-top:8px;max-width:660px"><label class="fl">Upfront commission schedules ($M, year1-year2-year3 = 2026/2027/2028)</label>'+
    '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-top:3px">'+
    '<input id="fl-fe-a" type="text" value="10,5,5" class="fi">'+
    '<input id="fl-fe-b" type="text" value="15,0,0" class="fi">'+
    '<input id="fl-fe-c" type="text" value="10,0,0" class="fi">'+
    '</div>'+hint('Three front-end schedules swept as structures: default 10-5-5, sensitivity 15-0-0, and a custom slot (default 10-0-0)')+'</div>'+
    '<div style="margin-top:8px;max-width:660px"><label class="fl">Issue-year scope</label>'+
    '<div style="display:flex;gap:14px;flex-wrap:wrap;margin-top:3px">'+
    '<label style="font-size:.73rem;cursor:pointer"><input type="checkbox" id="iy-full" checked> Full book (2019-2030)</label>'+
    '<label style="font-size:.73rem;cursor:pointer"><input type="checkbox" id="iy-no2025"> Exclude IssYr 2025</label>'+
    '<label style="font-size:.73rem;cursor:pointer"><input type="checkbox" id="iy-legacy"> Legacy only (pre-2022)</label>'+
    '</div></div>'+
    '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;max-width:660px;margin-top:10px">'+
    '<div><label class="fl">Claims sensitivities (scalars, comma-sep)</label><input id="fl-claim" type="text" value="1.0,1.05,1.10" class="fi">'+hint('1.0 = base; 1.05 = +5% claims')+'</div>'+
    '<div><label class="fl">Lapse sensitivities (scalars, comma-sep)</label><input id="fl-lapse" type="text" value="1.0,1.10" class="fi">'+hint('1.0 = base; 1.10 = +10% lapse')+'</div>'+
    '</div>'+
    '<div style="margin-top:14px"><button id="frontier-run-btn" class="btn btn-y" onclick="runFrontier()">&#x25b6; Run Frontier</button> '+
    '<span id="frontier-progress" style="font-size:.74rem;color:var(--mu);margin-left:10px"></span></div>'+
    '<div id="frontier-warning" style="font-size:.7rem;color:var(--err);margin-top:4px"></div>'+
    '</div></div><div id="frontier-chart-area"></div>';
}

async function runFrontier(){
  if(!S.out){alert('Run the model first so there is loaded data and a current deal to compare against.');return;}
  var btn=document.getElementById('frontier-run-btn'),prog=document.getElementById('frontier-progress'),warn=document.getElementById('frontier-warning');
  function plist(id){return (document.getElementById(id).value||'').split(',').map(function(x){return parseFloat(x);}).filter(function(v){return!isNaN(v);});}
  function sched(id){var v=plist(id);while(v.length<3)v.push(0);return v.slice(0,3);}
  var cede=plist('fl-cede'),cc=plist('fl-cc'),dur=plist('fl-dur').map(function(d){return Math.round(d);});
  var feRaw=[sched('fl-fe-a'),sched('fl-fe-b'),sched('fl-fe-c')].filter(function(s){return s.some(function(x){return x>0;});});
  var feSeen={},fe=[];feRaw.forEach(function(s){var k=s.join('-');if(!feSeen[k]){feSeen[k]=1;fe.push(s);}});
  if(!fe.length)fe=[[0,0,0]];
  var claims=plist('fl-claim'),lapses=plist('fl-lapse');
  if(!claims.length)claims=[1.0];if(!lapses.length)lapses=[1.0];
  function rng(a,b,ex){var s=[];for(var i=a;i<=b;i++)if(i!==ex)s.push(i);return s;}
  var scopes=[];
  if(document.getElementById('iy-full').checked)scopes.push(rng(2019,2030));
  if(document.getElementById('iy-no2025').checked)scopes.push(rng(2019,2030,2025));
  if(document.getElementById('iy-legacy').checked)scopes.push(rng(2019,2021));
  if(!scopes.length)scopes.push(rng(2019,2030));
  if(!cede.length||!fe.length||!cc.length||!dur.length){if(warn)warn.textContent='Fill in all four lever ranges.';return;}
  if(warn)warn.textContent='';
  var total=cede.length*fe.length*cc.length*dur.length*scopes.length*claims.length*lapses.length;
  if(btn)btn.disabled=true;if(prog)prog.textContent='Preparing '+total+' scenarios…';
  try{
    collectAssumptions();
    var py=S.py,a=S.assumptions,by=a.base_year||2025;
    var ev_for_py={agg:S.evData.agg,agg_iy:S.evData.agg_iy,periods:S.evData.periods,iss_years:S.evData.iss_years,row_count:S.evData.row_count};
    py.globals.set('_fev',JSON.stringify(ev_for_py));
    py.globals.set('_fba',JSON.stringify(a));
    py.globals.set('_FBY',by);
    py.globals.set('_fclaims',JSON.stringify(claims));
    py.globals.set('_flapses',JSON.stringify(lapses));
    // One-time setup: parse ev/base into Python globals (heavy payload crosses once),
    // then compute the no-deal RBC baselines per (claim,lapse) environment.
    var setup=["import json",
      "_ev=json.loads(_fev)","_base=json.loads(_fba)",
      "for k in ['acq_exp','maint_exp','nier','acq_exp_allowance','maint_exp_allowance']:",
      "    if k in _base and isinstance(_base[k],dict): _base[k]={int(kk):vv for kk,vv in _base[k].items()}",
      "_ev['agg']={k:{int(p):v for p,v in pv.items()} for k,pv in _ev['agg'].items()}",
      "_ev['agg_iy']={str(iy):{k:{int(p):v for p,v in pv.items()} for k,pv in vm.items()} for iy,vm in _ev['agg_iy'].items()}",
      "_ev['periods']=[int(p) for p in _ev['periods']]",
      "_FSURP=_surplus_rows if '_surplus_rows' in dir() else None",
      "_fr_nodeal=frontier_nodeal_baselines(_ev,_base,int(_FBY),_FSURP,[float(x) for x in json.loads(_fclaims)],[float(x) for x in json.loads(_flapses)])"
      ].join("\n");
    await py.runPythonAsync(setup);
    // Build the scenario combos (same product as `total`).
    var combos=[],n=0;
    claims.forEach(function(cs){lapses.forEach(function(ls){cede.forEach(function(cd){fe.forEach(function(sch){cc.forEach(function(ccm){scopes.forEach(function(sc){dur.forEach(function(nby){
      combos.push({cede:cd,sched:sch,ccm:ccm,scope:sc,nby:nby,cs:cs,ls:ls,n:++n});
    });});});});});});});
    // JS-driven loop: one run_frontier_one per scenario, live counter, yield each step.
    var perScenario=["import json","_p=json.loads(_fone)",
      "try:",
      "    _rec=run_frontier_one(_ev,_base,int(_FBY),_FSURP,float(_p['cede']),tuple(float(x) for x in _p['sched']),float(_p['ccm']),tuple(int(x) for x in _p['scope']),int(_p['nby']),float(_p['cs']),float(_p['ls']),_fr_nodeal.get((float(_p['cs']),float(_p['ls'])),0),int(_p['n']))",
      "    _out=json.dumps(_rec)",
      "except Exception as _e:",
      "    _out=json.dumps(None)",
      "_out"].join("\n");
    S.frontier=[];
    for(var i=0;i<combos.length;i++){
      py.globals.set('_fone',JSON.stringify(combos[i]));
      var rj=await py.runPythonAsync(perScenario);
      var rec=JSON.parse(rj);
      if(rec)S.frontier.push(rec);
      if(prog)prog.textContent=(i+1)+' / '+total+' scenarios run…';
      await new Promise(function(r){setTimeout(r);});  // yield so the UI repaints
    }
    if(prog)prog.textContent=S.frontier.length+' / '+total+' scenarios done ('+S.frontier.filter(function(r){return r.is_base;}).length+' structures).';
    renderFrontierChart();
  }catch(e){if(warn)warn.textContent='Error: '+e.message;console.error(e);}
  finally{if(btn)btn.disabled=false;}
}

function fpareto(pts){
  return pts.filter(function(p){
    return !pts.some(function(o){return o!==p && o.benefit>=p.benefit && o.cost<=p.cost && (o.benefit>p.benefit||o.cost<p.cost);});
  });
}
function currentDealPoint(){
  if(!S.out)return null;var ca=S.out.cedant_analytics||{};
  var ap=S.out.annual_predeal||{},an=S.out.annual_net||{};
  function gy(d,y){d=d||{};var v=(d[y]!=null?d[y]:d[String(y)]);return (v||0)/1e6;}
  var strain=[2026,2027,2028].reduce(function(s,y){return s+(gy(an.pretax_income,y)-gy(ap.pretax_income,y));},0);
  var cap=ca.cap_relief_value||0;
  return {cost:ca.reinsurer_pvde||0,benefit:cap+strain};
}
function fmtPt(r){return {x:r.cost,y:r.benefit,r:r};}

function renderFrontierChart(){
  var area=document.getElementById('frontier-chart-area');if(!area)return;
  var all=S.frontier||[];
  if(!all.length){area.innerHTML='<div class="empty"><h3>No scenarios</h3></div>';return;}
  var base=all.filter(function(r){return r.is_base;});
  var sens=all.filter(function(r){return !r.is_base;});
  var frontier=fpareto(base);
  var frSet=new Set(frontier.map(function(p){return p.n_run;}));
  var cur=currentDealPoint();
  area.innerHTML='<div class="card"><div class="ch">Cost vs Benefit Frontier</div><div class="cb">'+
    '<div style="font-size:.66rem;color:var(--mu);margin-bottom:6px">X = cost (Ceded PVDE given up, lower = better) &middot; Y = benefit (capital + early-strain relief, higher = better). Teal line = efficient frontier; gray = dominated structures; faint = claims&times;lapse sensitivities; &#9733; = current deal.</div>'+
    '<div style="position:relative;height:460px"><canvas id="frontier-canvas"></canvas></div>'+
    '<div id="frontier-table"></div></div></div>';
  var ds=[
    {label:'Sensitivities (claims×lapse)',data:sens.map(fmtPt),backgroundColor:'rgba(120,140,170,.16)',pointRadius:3,order:5},
    {label:'Dominated structures',data:base.filter(function(p){return!frSet.has(p.n_run);}).map(fmtPt),backgroundColor:'rgba(110,110,110,.55)',pointRadius:5,order:3},
    {label:'Efficient frontier',data:frontier.slice().sort(function(a,b){return a.cost-b.cost;}).map(fmtPt),backgroundColor:'#0a7080',borderColor:'#0a7080',pointRadius:6,showLine:true,borderWidth:2,fill:false,tension:.1,order:1}
  ];
  if(cur)ds.push({label:'Current deal',data:[{x:cur.cost,y:cur.benefit}],backgroundColor:'#E6C200',borderColor:'#1C2B6B',borderWidth:2,pointStyle:'star',pointRadius:14,order:0});
  if(S.frontierChart){S.frontierChart.destroy();}
  S.frontierChart=new Chart(document.getElementById('frontier-canvas'),{
    type:'scatter',data:{datasets:ds},
    options:{animation:false,responsive:true,maintainAspectRatio:false,
      plugins:{legend:{labels:{boxWidth:10,font:{size:10}}},
        tooltip:{callbacks:{label:function(c){var d=c.raw;if(!d||!d.r)return 'Cost $'+c.parsed.x.toFixed(1)+'M | Benefit $'+c.parsed.y.toFixed(1)+'M';var r=d.r;
          return ['cede '+(r.cede_pct*100).toFixed(0)+'% | upfront '+r.front_label+' | comm '+(r.cc_mult*100).toFixed(0)+'% | new-bus '+r.nb_years+'y',
            'IY '+r.iy_min+'-'+r.iy_max+(r.iy_excluded&&r.iy_excluded.length?' (excl '+r.iy_excluded.join(',')+')':''),
            'env: claims '+r.claim_scalar+' / lapse '+r.lapse_scalar,
            'Cost (ceded PVDE) $'+r.cost.toFixed(1)+'M | Benefit $'+r.benefit.toFixed(1)+'M',
            'cap relief $'+r.cap_relief.toFixed(1)+'M | strain $'+r.strain_relief.toFixed(1)+'M | RBC lift '+r.rbc_lift.toFixed(2)+'x',
            'net PVDE $'+r.net_pvde.toFixed(1)+'M | net IRR '+(r.net_irr*100).toFixed(1)+'%'];}}}},
      scales:{x:{title:{display:true,text:'Cost — Ceded PVDE given up ($M, lower = better)'}},
              y:{title:{display:true,text:'Benefit — capital + early-strain relief ($M, higher = better)'}}}}
  });
  renderFrontierTable(base,frSet);
}

function renderFrontierTable(base,frSet){
  var el=document.getElementById('frontier-table');if(!el)return;
  var rows=base.slice().sort(function(a,b){var af=frSet.has(a.n_run)?0:1,bf=frSet.has(b.n_run)?0:1;return af-bf||b.benefit-a.benefit;});
  var cols=['Structure','Cede%','Upfront $M','Comm%','NewBus yrs','IY','Cost $M','Cap Relief','Strain','Benefit $M','Net IRR'];
  el.innerHTML='<div style="font-size:.7rem;color:var(--mu);margin:10px 0 4px">Base-environment structures (claims 1.0, lapse 1.0), frontier rows highlighted. Benefit = capital relief + early-strain relief; cost = Ceded PVDE.</div>'+
    '<div class="tw" style="max-height:340px"><table class="bbt"><thead><tr>'+cols.map(function(c){return '<th>'+c+'</th>';}).join('')+'</tr></thead><tbody>'+
    rows.map(function(r){var fr=frSet.has(r.n_run);
      return '<tr'+(fr?' style="background:#EEF8F2"':'')+'><td>'+(fr?'★ ':'')+'#'+r.n_run+'</td><td>'+(r.cede_pct*100).toFixed(0)+'%</td><td>'+r.front_label+'</td><td>'+(r.cc_mult*100).toFixed(0)+'%</td><td>'+r.nb_years+'</td><td>'+r.iy_min+'-'+r.iy_max+'</td><td>'+r.cost.toFixed(1)+'</td><td>'+r.cap_relief.toFixed(1)+'</td><td>'+r.strain_relief.toFixed(1)+'</td><td>'+r.benefit.toFixed(1)+'</td><td>'+(r.net_irr*100).toFixed(1)+'%</td></tr>';
    }).join('')+'</tbody></table></div>';
}
