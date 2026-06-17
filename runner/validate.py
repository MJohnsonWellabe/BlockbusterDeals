#!/usr/bin/env python3
"""Headless tie-out gate: run the engine on the source-workbook deal
(MS_Reins_Projection_04302026Slim.xlsx) and assert it reproduces the Excel
Predeal/Ceded/Net pretax-income vectors, their PVs, and the after-deal RBC
ratios, plus the back-book/new-issue cohort reconciliation.

Run:  python3 runner/validate.py      (exit 0 = PASS, 1 = FAIL)

The deal mirrors the Excel Assumptions tab: 10% quota-share on issue years
2019-2030 (carried to every calendar year), FLAT $200 ongoing ceding
commission (the workbook is not loss-ratio split; the online model keeps the
sliding scale), and a 10-5-5 front-end schedule. EV input is data/EV_Data_Final.csv
(already equal to the workbook EV tab); the new surplus methodology reads
data/input_surplus.json + data/surplus_ts.json (Surplus Calc = predeal RBC,
Surplus Recalc = net RBC).
"""
import csv, json, math, os, re, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))
import engine  # noqa: E402

DISC = 0.08
KEY_VARS = {"EarnedPrem", "ReinsPrem", "IncClaims", "ReinsClaims", "TabRes",
            "CededALRstat", "CLRes", "TS", "Comm", "PremTax", "LivesIssued",
            "AEGAdminPolCount"}

# ---- Excel tie-out targets (extracted from the workbook) -------------------
# Predeal PTI row 22, Ceded PTI row 24, Net PTI row 24; cols C..AF = CY 2026..2055.
T_PRED_PV, T_CED_PV, T_NET_PV = 1455214209.5104718, 94090444.69550136, 1361123764.8149707
T_PRED_PTI = [-121101932.04611076, -100616411.43403262, -53679465.31813805, 15381391.345005073, 76358086.97474386, 130989187.301274, 186353851.1525973, 223120550.39628565, 255488349.46755308, 278009465.6039558, 314926651.8174915, 305619564.299919, 291974708.26535904, 278870135.0184477, 268318801.31718102, 260904989.9107015, 246793222.33700934, 231182920.75200757, 218373447.8507943, 206937698.58638757, 195359592.8803895, 183943975.6419137, 173371012.6857284, 163764425.83055335, 154520622.28677547, 145299661.11419195, 136352613.3664125, 127247949.62126507, 118106478.4063903, 108800195.56305732]
T_CED_PTI = [-19549512.56450016, -13665475.05250083, -9625856.293084448, 2245870.7268634713, 8198969.9405804565, 13520472.698784925, 18331326.02679724, 19911834.262310997, 21139449.227930672, 21537513.55564603, 22239145.15473425, 21429780.58458346, 20172686.20630191, 19104795.770407684, 18125210.994539157, 17121725.500708044, 16239288.24301054, 15513774.791541934, 14891089.215387275, 14333311.574388972, 13831776.401069818, 13278080.136416249, 12663239.919148393, 12048937.952838896, 11400323.92648054, 10696280.287475713, 9965138.63100611, 9172451.151036179, 8352039.171711634, 7545591.286351269]
T_NET_PTI = [-101552419.48161039, -86950936.38153177, -44053609.02505342, 13135520.618141562, 68159117.0341641, 117468714.60248867, 168022525.1257998, 203208716.1339743, 234348900.2396223, 256471952.0483098, 292687506.6627578, 284189783.7153348, 271802022.059057, 259765339.24804002, 250193590.3226419, 243783264.4099931, 230553934.09399873, 215669145.96046564, 203482358.63540688, 192604387.01199868, 181527816.47931984, 170665895.5054977, 160707772.76657987, 151715487.87771437, 143120298.36029476, 134603380.82671627, 126387474.73540637, 118075498.47022882, 109754439.23467857, 101254604.27670605]
# RBC ratios from the new surplus methodology: Surplus Calc r52 (predeal) and
# Surplus Recalc r52 (net), CY 2025..2034.
T_RBC_PREDEAL = [7.194, 5.674, 5.148, 4.351, 4.273, 4.723, 4.984, 5.306, 6.208, 7.239]
T_RBC_NET = [7.194, 6.040, 5.688, 4.928, 4.800, 5.193, 5.353, 5.554, 6.361, 7.304]
PTI_YEARS = list(range(2026, 2056))   # 30 cols C..AF
RBC_YEARS = list(range(2025, 2035))


def load_ev(path):
    agg, agg_iy = {}, {}
    with open(path) as f:
        rd = csv.reader(f)
        hdr = next(rd)
        vnI, iyI = hdr.index("VarName"), hdr.index("ck.IssYear")
        valcols = sorted([(i, int(h[5:])) for i, h in enumerate(hdr) if re.match(r"^Value\d+$", h)],
                         key=lambda x: x[1])
        periods = [p for _, p in valcols]
        for row in rd:
            if not row or len(row) <= vnI:
                continue
            vn = row[vnI].strip().strip('"')
            if not vn:
                continue
            if vn == "LivesInForce1":
                vn = "AEGAdminPolCount"
            m = re.search(r"\d{4}", str(row[iyI]))
            iy = int(m.group(0)) if m else 0
            if vn not in KEY_VARS:
                continue
            agg.setdefault(vn, {})
            iys = str(iy)
            agg_iy.setdefault(iys, {}).setdefault(vn, {})
            for ci, p in valcols:
                try:
                    fv = float(row[ci])
                except (TypeError, ValueError):
                    continue
                if not fv:
                    continue
                agg[vn][p] = agg[vn].get(p, 0) + fv
                agg_iy[iys][vn][p] = agg_iy[iys][vn].get(p, 0) + fv
    return {"agg": agg, "agg_iy": agg_iy, "periods": periods,
            "iss_years": sorted(int(k) for k in agg_iy)}


def deal_assumptions():
    ann = json.load(open(os.path.join(ROOT, "data", "ann_assum.json")))
    yd = lambda key: {int(y): ann[y][key] for y in ann}
    rp = {}
    for iy in range(2019, 2031):
        rp[iy] = {cal: 0.10 for cal in range(max(2026, iy + 1), 2036)}
    return {"premium_tax": 0.021, "discount_rate": DISC, "cost_of_capital": 0.10,
            "ceding_comm_ongoing": 200, "base_year": 2025, "claim_scalar": 1.0,
            "lapse_scalar": 1.0, "reins_pct": rp,
            "ceding_comm_table": [[0, float("inf"), 200]],         # flat $200 (validation)
            "ceding_comm_front": {2026: 10e6, 2027: 5e6, 2028: 5e6},
            "nier": yd("nier"), "acq_exp": yd("acq_exp"), "maint_exp": yd("maint_exp"),
            "acq_exp_allowance": yd("acq_exp_allowance"),
            "maint_exp_allowance": yd("maint_exp_allowance")}


def gy(d, yr):
    return d.get(yr, d.get(str(yr), 0)) or 0


def pv(vec):
    return sum(v / (1 + DISC) ** i for i, v in enumerate(vec, 1))


def check_vec(name, got, want, rtol=2e-3, atol=5e3, fails=None):
    md = 0.0; worst = None
    for i, (g, w) in enumerate(zip(got, want)):
        d = abs(g - w)
        if d > md:
            md, worst = d, (PTI_YEARS[i], g, w)
        if d > atol + rtol * abs(w):
            fails.append("%s[%d]: got %.1f want %.1f" % (name, PTI_YEARS[i], g, w))
    print("  %-12s max abs diff %.1f at %s" % (name, md, worst[0] if worst else "-"))


def main():
    ev = load_ev(os.path.join(ROOT, "data", "EV_Data_Final.csv"))
    surplus = {"input_surplus": json.load(open(os.path.join(ROOT, "data", "input_surplus.json"))),
               "surplus_ts": json.load(open(os.path.join(ROOT, "data", "surplus_ts.json")))}
    out = engine.run_model(ev, deal_assumptions(), 2025, surplus_rows=surplus, lite=False)
    fails = []

    print("Pretax income vectors (engine vs Excel, $):")
    for name, stmt, target, tpv in [
            ("predeal", out["annual_predeal"], T_PRED_PTI, T_PRED_PV),
            ("ceded", out["annual_ceded"], T_CED_PTI, T_CED_PV),
            ("net", out["annual_net"], T_NET_PTI, T_NET_PV)]:
        got = [gy(stmt["pretax_income"], y) for y in PTI_YEARS]
        check_vec(name, got, target, fails=fails)
        gpv = pv(got)
        rel = abs(gpv - tpv) / abs(tpv)
        print("    PV: engine %.0f  excel %.0f  rel %.2e %s"
              % (gpv, tpv, rel, "OK" if rel < 2e-3 else "FAIL"))
        if rel >= 2e-3:
            fails.append("%s PV rel %.2e" % (name, rel))

    print("RBC ratios vs workbook (Surplus Calc=predeal, Surplus Recalc=net):")
    pa = out.get("rbc_predeal_result", {}).get("predeal_adjustments", {})
    na = out.get("rbc_net_result", {}).get("net_adjustments", {})
    # predeal ties exactly (1e-2); net within ~1% (engine computes ceded lives/claims
    # proportionally vs the workbook's separately-projected EV_Ceded tab).
    for label, adjs, target, tol in [("predeal", pa, T_RBC_PREDEAL, 0.01),
                                     ("net", na, T_RBC_NET, 0.08)]:
        md = 0.0; worst = None
        for i, y in enumerate(RBC_YEARS):
            g = (adjs.get(y) or adjs.get(str(y)) or {}).get("ratio_w_margin")
            w = target[i]
            if g is None:
                fails.append("RBC %s %d: missing" % (label, y)); continue
            if abs(g - w) > md:
                md, worst = abs(g - w), y
            if abs(g - w) > tol:
                fails.append("RBC %s %d: got %.3f want %.3f" % (label, y, g, w))
        print("  %-8s max abs diff %.3f at %s (tol %.2f)" % (label, md, worst, tol))

    print("Cohort reconciliation (back + new == portfolio, $M):")
    ca = out["cedant_analytics"]; mb, mn = ca["metrics_back"], ca["metrics_new"]
    port_pred, port_net = out["metrics_predeal"]["pvde"] / 1e6, out["metrics_net"]["pvde"] / 1e6
    for lbl, b, n, port in [("predeal_pvde", mb["predeal_pvde"], mn["predeal_pvde"], port_pred),
                            ("net_pvde", mb["net_pvde"], mn["net_pvde"], port_net)]:
        s = b + n
        d = abs(s - port)
        ok = d < 0.5  # $0.5M tolerance (monthly-vs-annual DE method)
        print("  %-12s back %.1f + new %.1f = %.1f vs portfolio %.1f  (diff %.2f) %s"
              % (lbl, b, n, s, port, d, "OK" if ok else "FAIL"))
        if not ok:
            fails.append("%s reconcile diff %.2f" % (lbl, d))
    print("  back-book EV %.1f -> %.1f | front-comm PV %.1f | back ceded PVDE %.1f"
          % (mb["predeal_pvde"], mb["net_pvde"], ca["comm_front_pv"], mb["ceded_pvde"]))

    print()
    if fails:
        print("FAIL (%d):" % len(fails))
        for f in fails[:20]:
            print("  -", f)
        sys.exit(1)
    print("PASS — engine ties out to the Excel workbook.")


if __name__ == "__main__":
    main()
