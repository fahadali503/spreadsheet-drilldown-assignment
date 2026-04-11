# Fortune Sheet — first-level drilldown

React + TypeScript demo for **@fortune-sheet/react** (^1.0.4), modeled on **Excel Pivot Table drill-down** (“Show Details”): double-click a summary cell to open a **temporary detail view** without changing the underlying summary.

Reference walkthrough (Excel2016 pivot drill-down): [YouTube — Pivot Table drill-down](https://www.youtube.com/watch?v=jYQ8Bj3ki0g).

## Workbook layout

- **Summary** — Aggregated regional sales (the “pivot-style” view).
- **Source data** — Fixed row-level transactions; amounts **roll up** to the Summary cells (same idea as the raw table behind a pivot). In **Open Source tab (Excel-style)** mode, double-click drill **switches to this tab**.
- **Same-sheet demo** — Replaces the Summary grid in place with generated detail rows (not the Source data tab).

## Requirements

- Node.js 18+ (recommended)

## Install

```bash
npm install
```

## Run (development)

```bash
npm run dev
```

Open the URL shown in the terminal (usually `http://localhost:5173`).

## Build

```bash
npm run build
npm run preview
```

## Behavior (Excel-style)

1. **Double-click** cell **A1** on the **Summary** sheet (row 0, column 0). A single click only selects the cell (like Excel).

2. **Open Source tab (Excel-style)** — Default. **Activates the Source data** tab (the underlying rows); the Summary sheet is unchanged. Use **Back to summary** to return to the Summary tab.

3. **Same sheet (demo)** — Replaces the visible grid in place with detail data; **Back to summary** restores the summary via a saved snapshot and workbook remount.

## Assumptions

- Drilldown is **first-level only** (no chaining from the detail view).
- Trigger is **double-click** on **(0,0)** on the **Summary** sheet only (capture-phase `dblclick` on the grid + `getSelection()` for range `[0,0]×[0,0]`, with `afterActivateSheet` tracking the active tab).
- Mock data uses `Math.random()` for amounts, IDs, and row counts so each drill looks different (not filtered per cell like real Excel).

## Stack

- [Vite](https://vite.dev/) + React + TypeScript
- [@fortune-sheet/react](https://www.npmjs.com/package/@fortune-sheet/react)
