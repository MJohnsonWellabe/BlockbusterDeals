# Blockbuster Deals: Reinsurance

A browser-based actuarial projection model for Medicare Supplement quota-share
reinsurance deal analysis, from the ceding company's perspective. Computation
runs in-browser via Pyodide (Python/WASM); no server or build step is required.

## Structure (decomposed)

```
index.html            Password landing page (gate) -> forwards to viewer/
viewer/index.html     App shell: markup, libraries, auth guard, loads app.js
viewer/app.js         UI / rendering / interaction layer
src/engine.py         Actuarial engine (income statements, EV/PVDE/IRR, RBC,
                      cedant analytics, scenario matrix)
data/                 Default inputs fetched at startup:
  EV_Data_Final.csv     30-year monthly EV projection (issue-year only, slim)
  surplus.json          RBC / surplus rows
  balanceplan.xlsx      Balance plan
  ann_assum.json        NIER / expense assumptions
  doc.txt               In-app documentation text
```

`viewer/app.js` fetches `../src/engine.py` and `../data/*` on load, so the app
must be served over **http(s)** (a `file://` open is blocked by fetch). The
engine and viewer share one source of compute truth: `src/engine.py`.

## Access

Client-side password gate (SHA-256 hash in `sessionStorage`) — obfuscation
only on a static site. To change the password, recompute its SHA-256 and
replace `EXPECTED` in `index.html` and the guard in `viewer/index.html`.

## Run locally (no Node/Python needed)

Serve the repo root over http and open `index.html`. On Windows with no
toolchain, the built-in PowerShell HttpListener works from the repo root:

```bat
powershell -NoProfile -Command "$root=(Get-Location).Path; $l=[System.Net.HttpListener]::new(); $l.Prefixes.Add('http://localhost:8000/'); $l.Start(); Write-Host 'http://localhost:8000/'; while($l.IsListening){ $c=$l.GetContext(); $p=$c.Request.Url.LocalPath.TrimStart('/'); if([string]::IsNullOrEmpty($p)){$p='index.html'}; $f=Join-Path $root $p; if(Test-Path $f -PathType Leaf){ $b=[System.IO.File]::ReadAllBytes($f); $ext=[System.IO.Path]::GetExtension($f).ToLower(); $m=@{'.html'='text/html';'.js'='text/javascript';'.json'='application/json';'.csv'='text/csv';'.css'='text/css';'.xlsx'='application/octet-stream'}; $ct=$m[$ext]; if(-not $ct){$ct='application/octet-stream'}; $c.Response.ContentType=$ct; $c.Response.OutputStream.Write($b,0,$b.Length) } else { $c.Response.StatusCode=404 }; $c.Response.Close() }"
```

Then open `http://localhost:8000/` and enter the password.

## Deploy (GitHub Pages)

Serve the repository root. Pages loads `index.html` (the gate), which forwards
to `viewer/index.html` after authentication.

## Key analytics

- Predeal / Ceded / Net income statements, EV (PVDE), IRR, RBC & surplus rebuild.
- Cedant Economics: Net Deal Value (capital-relief value - reinsurer PVDE),
  cost-of-capital-based capital relief, reinsurer profit, cedant recovery %,
  back-book vs new-issue IRR.
- Scenario Matrix: combinatoric cedant sweep ranked by Net Deal Value.
