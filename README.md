# Blockbuster Deals: Reinsurance

A fully browser-based actuarial projection model for reinsurance deal analysis.

## Deployment (GitHub Pages)

1. Push `index.html` to your GitHub repository's `main` or `gh-pages` branch
2. Enable GitHub Pages in repository Settings → Pages
3. The app loads at `https://yourusername.github.io/your-repo/`

**No server required.** All computation runs in-browser via Pyodide (Python WASM).

## Usage

1. **Upload** your EV tab Excel file (required). Assumptions are auto-extracted.
2. **Assumptions** — review and edit the Reins % matrix, NIER, expenses, discount rate, etc.
3. **Run Model** — computes Predeal, Ceded (EV × Reins%), and Net views
4. **Results** — view Annual/Monthly income statements, summary metrics (EV, IRR, PVDE), and charts
5. **Scenarios** — save runs for comparison, reload assumptions from prior runs
6. **Audit & Review** — peer reviewer signs off with name and comment; immutable run log

## Calculation Logic

| Metric | Formula |
|--------|---------|
| NII | avg(TotalReserve_t, TotalReserve_{t-1}) × NIER[CalendarYear] / 12 |
| EV_Ceded | EV × Reins%[IssueYear][CalendarYear] (lower-triangle aware) |
| Net view | Predeal − Ceded + Ceding Comms |
| DE | Pretax Income (surplus adjustment applied on monthly view) |
| PVDE (EV) | Σ DE_t / (1 + disc/12)^t |
| IRR | Monthly IRR of DE series, annualized |

## First-Load Time

~15–20 seconds on first load (Pyodide + openpyxl download). Subsequent loads are faster with browser cache.
