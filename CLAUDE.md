# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the App

No build step. Open directly in a browser:

```bash
open rfq-estimate.html
```

## Repository Overview

This repo contains two independent tools for the same purpose — RFQ cost estimation:

| File | Platform | Use case |
|------|----------|----------|
| `rfq-estimate.html` | Browser (any OS) | Interactive, real-time cost estimator |
| `rfq_quote.ps1` | Windows PowerShell | Batch processor: reads RFQ Excel sheets, writes 報價彙整 summary |

The Excel files in `RFQ_outupt/` are output samples from `rfq_quote.ps1`. `估價成本計算_2025.05.07.xlsx` is the machine rate reference table required by the PS1 script.

---

## rfq-estimate.html — Architecture

### Two modes

Toggled by `setMode()`:
- **單工件** (`MODE = 'single'`): state lives in `S`, calculates via `calcQty(qty)`
- **組合件** (`MODE = 'assembly'`): state lives in `A` (array of workpieces), calculates via `calcAssemblySet(qty)`

### State objects

```js
S = { qtys, procs, nextId, activeQty, matTiers, purchItems, nextPurchId }
A = { workpieces: [{ id, name, partNo, material, qtyPerUnit, procs, matTiers, matSetupQty }], purchItems }
```

Every data mutation calls `recalc()` which re-renders the results sidebar.

### Calculation formula

```
totalCostNTD = mat + purchased + inhouse + outsource + setupAmort + oneTimeAmort + commission
priceUSD = (totalCostNTD / exchangeRate) / (1 - margin)
```

- **In-house machining**: `timeMin × machine.perMin (NTD/min)`
- **Setup amortization**: `setupHrs × machine.setupHrUSD × exchangeRate / qty`
- **Outsourced processes**: tiered pricing, looked up by `getTierPrice(tiers, qty)`
- **Material cost**: also tiered, with optional batch over-quantity (`matSetupQty`)
- **Commission**: optional flat 5% on base cost

### Margin levels (difficulty)

| Level | Margin |
|-------|--------|
| 0 (default) | 30% |
| 1–5 | 25% / 27.5% / 32.5% / 37.5% / 42.5% |
| custom | user input |

### Hardcoded data

- `MACHINES`: machine IDs, NTD/min rates, setup costs in USD — edit directly in the `<script>` block
- `OS_PRESETS`: outsourcing autocomplete list
- `CUSTOMERS`: same list as the `quotation` project — must be kept in sync manually if changed

### Persistence

- **Save/Load**: JSON file download/upload via `saveCalc()` / `loadCalc()`
- **RFQ number**: `RFQ{YYYYMMDD}{seq:03}`, daily counter in `localStorage` key `achb_rfq_counter`
- **Excel export**: SheetJS (CDN) via `exportExcel()` — requires internet

---

## rfq_quote.ps1 — Architecture

Windows-only, requires Excel installed (uses COM automation).

### How to run

```powershell
# Auto-detect RFQ*.xlsx in current folder
.\rfq_quote.ps1

# Explicit paths
.\rfq_quote.ps1 -RfqPath "RFQ_FBD01.xlsx" -CostTablePath "估價成本計算_2025.05.07.xlsx" -ExchangeRate 32 -SetupMode first
```

`-SetupMode` accepts `production` (量產, default) or `first` (首次製作).

### Excel template structure expected per sheet

| Rows | Content |
|------|---------|
| Row 3 | Customer (col B), qty tiers (cols D–N) |
| Row 4 | Part No (col D), Revision (col F), Material (col H), Surface (col J) |
| Rows 10/13/16/19/22 | Material vendors + prices (col D) |
| Rows 25–34 | Purchased parts |
| Rows 35+ | `OP01`, `OP02`… operation rows: machine (col B), time min (col C), setup notes (col D), ext. vendor/price (col H/I) |
| Rows 90+ | One-time costs (tooling, fixture, mold…) |

### Processing logic

1. `Load-MachineRates` reads rates from `估價成本計算_2025.05.07.xlsx` sheet 1, rows 3–40
2. `Process-PartSheet` scans each non-template sheet and returns a cost struct per qty tier
3. `Write-SummarySheet` creates/replaces the `報價彙整` sheet with all parts side by side
4. If a machine name can't be matched, it's flagged as `UnmappedMachines` (warning in output)

Machine name matching uses `Get-MachineRate`: exact key lookup → CNC number regex → keyword switch.
