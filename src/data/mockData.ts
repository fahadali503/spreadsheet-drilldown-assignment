import type { CellWithRowAndCol, Sheet } from '@fortune-sheet/core'

/** Stable id for the summary sheet (used for activateSheet / deleteSheet). */
export const SUMMARY_SHEET_ID = 'summary-sheet'

/** Row-level “raw” data behind the summary (like Excel pivot source / drill-down records). */
export const SOURCE_DETAIL_SHEET_ID = 'source-detail-sheet'

const CT_G = { fa: 'General' as const, t: 'g' as const }
const CT_N = { fa: 'General' as const, t: 'n' as const }

function money(): number {
  return Math.round((Math.random() * 5000 + 100) * 100) / 100
}

type SourceRow = {
  region: string
  quarter: string
  id: string
  product: string
  amount: number
  rep: string
}

/**
 * Fixed transaction rows whose amounts roll up to the Summary sheet
 * (North/South/East × Q1/Q2 match the totals in `createInitialWorkbook`).
 */
function buildSourceDetailCelldata(): CellWithRowAndCol[] {
  const rows: SourceRow[] = [
    { region: 'North', quarter: 'Q1', id: 'TX-50101', product: 'Widget A', amount: 4500, rep: 'Alex' },
    { region: 'North', quarter: 'Q1', id: 'TX-50102', product: 'License', amount: 7500, rep: 'Jordan' },
    { region: 'North', quarter: 'Q2', id: 'TX-50103', product: 'Service Pack', amount: 5000, rep: 'Sam' },
    { region: 'North', quarter: 'Q2', id: 'TX-50104', product: 'Widget B', amount: 10000, rep: 'Taylor' },
    { region: 'South', quarter: 'Q1', id: 'TX-50201', product: 'Addon', amount: 4800, rep: 'Riley' },
    { region: 'South', quarter: 'Q1', id: 'TX-50202', product: 'Widget A', amount: 5000, rep: 'Alex' },
    { region: 'South', quarter: 'Q2', id: 'TX-50203', product: 'License', amount: 5200, rep: 'Jordan' },
    { region: 'South', quarter: 'Q2', id: 'TX-50204', product: 'Widget B', amount: 6000, rep: 'Sam' },
    { region: 'East', quarter: 'Q1', id: 'TX-50301', product: 'Service Pack', amount: 7300, rep: 'Taylor' },
    { region: 'East', quarter: 'Q1', id: 'TX-50302', product: 'Widget A', amount: 7000, rep: 'Riley' },
    { region: 'East', quarter: 'Q2', id: 'TX-50303', product: 'License', amount: 6100, rep: 'Alex' },
    { region: 'East', quarter: 'Q2', id: 'TX-50304', product: 'Addon', amount: 7000, rep: 'Jordan' },
  ]

  const headerLabels = ['Region', 'Quarter', 'Txn ID', 'Product', 'Amount', 'Rep']
  const celldata: CellWithRowAndCol[] = headerLabels.map((label, c) => ({
    r: 0,
    c,
    v: {
      v: label,
      m: label,
      ct: CT_G,
      bl: 1,
      bg: '#dcfce7',
    },
  }))

  rows.forEach((row, idx) => {
    const r = idx + 1
    celldata.push(
      { r, c: 0, v: { v: row.region, m: row.region, ct: CT_G } },
      { r, c: 1, v: { v: row.quarter, m: row.quarter, ct: CT_G } },
      { r, c: 2, v: { v: row.id, m: row.id, ct: CT_G } },
      { r, c: 3, v: { v: row.product, m: row.product, ct: CT_G } },
      { r, c: 4, v: { v: row.amount, m: String(row.amount), ct: CT_N } },
      { r, c: 5, v: { v: row.rep, m: row.rep, ct: CT_G } },
    )
  })

  return celldata
}

/** Initial summary workbook: double-click cell A1 (0,0) to drill (Excel-style). */
export function createInitialWorkbook(): Sheet[] {
  const celldata: CellWithRowAndCol[] = [
    {
      r: 0,
      c: 0,
      v: {
        v: 'Double-click to drill (A1)',
        m: 'Double-click to drill (A1)',
        ct: CT_G,
        bg: '#bfdbfe',
        bl: 1,
        fc: '#1e3a8a',
      },
    },
    { r: 0, c: 1, v: { v: 'Q1 Sales', m: 'Q1 Sales', ct: CT_G, bl: 1 } },
    { r: 0, c: 2, v: { v: 'Q2 Sales', m: 'Q2 Sales', ct: CT_G, bl: 1 } },
    { r: 0, c: 3, v: { v: 'Total', m: 'Total', ct: CT_G, bl: 1 } },
    {
      r: 1,
      c: 0,
      v: { v: 'North', m: 'North', ct: CT_G },
    },
    { r: 1, c: 1, v: { v: 12000, m: '12000', ct: CT_N } },
    { r: 1, c: 2, v: { v: 15000, m: '15000', ct: CT_N } },
    { r: 1, c: 3, v: { v: 27000, m: '27000', ct: CT_N } },
    {
      r: 2,
      c: 0,
      v: { v: 'South', m: 'South', ct: CT_G },
    },
    { r: 2, c: 1, v: { v: 9800, m: '9800', ct: CT_N } },
    { r: 2, c: 2, v: { v: 11200, m: '11200', ct: CT_N } },
    { r: 2, c: 3, v: { v: 21000, m: '21000', ct: CT_N } },
    {
      r: 3,
      c: 0,
      v: { v: 'East', m: 'East', ct: CT_G },
    },
    { r: 3, c: 1, v: { v: 14300, m: '14300', ct: CT_N } },
    { r: 3, c: 2, v: { v: 13100, m: '13100', ct: CT_N } },
    { r: 3, c: 3, v: { v: 27400, m: '27400', ct: CT_N } },
  ]

  return [
    {
      id: SUMMARY_SHEET_ID,
      name: 'Summary',
      celldata,
      row: 24,
      column: 12,
      order: 0,
    },
    {
      id: SOURCE_DETAIL_SHEET_ID,
      name: 'Source data',
      celldata: buildSourceDetailCelldata(),
      row: 36,
      column: 10,
      order: 1,
    },
  ]
}

/** Random detail rows as sparse celldata (for replacing the same sheet). */
export function buildDrilldownCelldata(): CellWithRowAndCol[] {
  const rowCount = 6 + Math.floor(Math.random() * 8)
  const headers: { c: number; label: string }[] = [
    { c: 0, label: 'Txn ID' },
    { c: 1, label: 'Product' },
    { c: 2, label: 'Amount' },
    { c: 3, label: 'Rep' },
  ]

  const products = ['Widget A', 'Widget B', 'Service Pack', 'License', 'Addon']
  const reps = ['Alex', 'Jordan', 'Sam', 'Taylor', 'Riley']

  const celldata: CellWithRowAndCol[] = headers.map((h) => ({
    r: 0,
    c: h.c,
    v: {
      v: h.label,
      m: h.label,
      ct: CT_G,
      bl: 1,
      bg: '#e0e7ff',
    },
  }))

  for (let r = 1; r <= rowCount; r += 1) {
    const id = `TX-${10000 + Math.floor(Math.random() * 89999)}`
    const product = products[Math.floor(Math.random() * products.length)]
    const rep = reps[Math.floor(Math.random() * reps.length)]
    const amt = money()
    celldata.push(
      { r, c: 0, v: { v: id, m: id, ct: CT_G } },
      { r, c: 1, v: { v: product, m: product, ct: CT_G } },
      { r, c: 2, v: { v: amt, m: String(amt), ct: CT_N } },
      { r, c: 3, v: { v: rep, m: rep, ct: CT_G } },
    )
  }

  return celldata
}
