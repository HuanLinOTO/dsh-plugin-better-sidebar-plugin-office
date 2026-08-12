/**
 * Convert a SheetJS (xlsx) workbook into Univer's {@link IWorkbookData} so
 * {@link XlsxView} can render it through the Univer sheets preset.
 *
 * v1 scope (matches docs/plans/2026-08-10-office-preview-design.md §3.6):
 * - Sheet names + order
 * - Cell values, typed (string / number / boolean / error → force-string)
 * - Cell formulas (text form; Univer's formula engine computes them)
 * - Merged cells (!merges → mergeData)
 * - Column widths (!cols → columnData) and row heights (!rows → rowData)
 *
 * Out of scope for v1 (SheetJS community edition does not parse them and the
 * Pro edition is commercial): cell styles (font/fill/border/alignment),
 * conditional formatting, charts. Users who need style fidelity get the
 * download button on the binary placeholder.
 *
 * The function is pure (no DOM, no Univer imports) so it is unit-testable
 * without jsdom or canvas.
 */
import type * as XLSX from 'xlsx'
// Types come from @univerjs/presets (which re-exports @univerjs/core's type
// surface). The enum values BooleanNumber / CellValueType are imported as
// runtime values from the same module — the bundler inlines them, and the
// test environment has @univerjs/presets available as a dependency.
import { BooleanNumber, CellValueType, type LocaleType, type ICellData, type IColumnData, type IRowData, type IWorkbookData, type IWorksheetData, type IObjectArrayPrimitiveType, type IObjectMatrixPrimitiveType, type IRange } from '@univerjs/presets'

/** SheetJS workbook (the slice of the module we use). */
type XLSXWorkbook = XLSX.WorkBook

/**
 * Build a Univer workbook snapshot from a SheetJS workbook.
 *
 * @param wb - the parsed workbook from `XLSX.read(buf, { type: 'array' })`
 * @param appVersion - the Univer package version, written into `appVersion`
 *   (Univer requires the field but does not gate on it).
 * @param locale - the locale Univer should render in.
 */
export function xlsxWorkbookToUniver(
  wb: XLSXWorkbook,
  appVersion: string,
  locale: LocaleType,
): IWorkbookData {
  const sheetOrder: string[] = []
  const sheets: Record<string, Partial<IWorksheetData>> = {}

  wb.SheetNames.forEach((name, index) => {
    const sheetId = `sheet-${index}`
    sheetOrder.push(sheetId)
    sheets[sheetId] = worksheetToUniver(wb.Sheets[name] ?? {}, name, sheetId)
  })

  // Univer requires at least one sheet; an empty workbook gets a placeholder.
  if (sheetOrder.length === 0) {
    const fallbackId = 'sheet-0'
    sheetOrder.push(fallbackId)
    sheets[fallbackId] = {
      id: fallbackId,
      name: 'Sheet1',
      rowCount: 20,
      columnCount: 26,
      cellData: {},
      rowData: {},
      columnData: {},
      mergeData: [],
    }
  }

  return {
    id: 'workbook',
    name: wb.Props?.Title || 'Workbook',
    appVersion,
    locale,
    styles: {},
    sheetOrder,
    sheets,
  }
}

/** Convert one SheetJS worksheet to a Univer IWorksheetData. */
function worksheetToUniver(
  ws: XLSX.WorkSheet,
  name: string,
  sheetId: string,
): Partial<IWorksheetData> {
  const cellData: IObjectMatrixPrimitiveType<ICellData> = {}
  const merges: IRange[] = []

  // Cell values + formulas: iterate every cell address (skip the !-prefixed
  // special keys SheetJS uses for sheet-level metadata).
  for (const addr of Object.keys(ws)) {
    if (addr.startsWith('!')) continue
    const cell = ws[addr] as XLSX.CellObject | undefined
    if (cell === undefined) continue
    const { r, c } = decodeAddr(addr)
    const converted = convertCell(cell)
    if (converted === null) continue
    let row = cellData[r]
    if (row === undefined) {
      row = {}
      cellData[r] = row
    }
    row[c] = converted
  }

  // Merged cells: SheetJS's Range uses { s: { r, c }, e: { r, c } } (0-indexed),
  // which is exactly Univer's IRange shape — copy verbatim.
  if (Array.isArray(ws['!merges'])) {
    for (const m of ws['!merges']!) {
      if (m === undefined || m === null) continue
      merges.push({ startRow: m.s.r, endRow: m.e.r, startColumn: m.s.c, endColumn: m.e.c })
    }
  }

  const { rowCount, columnCount } = dimensions(ws, cellData)

  return {
    id: sheetId,
    name,
    tabColor: '',
    hidden: BooleanNumber.FALSE,
    freeze: { startRow: -1, startColumn: -1, xSplit: 0, ySplit: 0 },
    rowCount,
    columnCount,
    defaultColumnWidth: 96,
    defaultRowHeight: 22,
    mergeData: merges,
    cellData,
    rowData: convertRows(ws['!rows']),
    columnData: convertCols(ws['!cols']),
    rowHeader: { width: 46 },
    columnHeader: { height: 20 },
    showGridlines: BooleanNumber.TRUE,
    rightToLeft: BooleanNumber.FALSE,
  }
}

/** Map a SheetJS CellObject onto Univer's ICellData (or null for empty). */
function convertCell(cell: XLSX.CellObject): ICellData | null {
  const out: ICellData = {}
  // Type mapping: b→BOOLEAN, n→NUMBER, s→STRING, d→STRING, e→FORCE_STRING, z→skip
  switch (cell.t) {
    case 'b':
      out.t = CellValueType.BOOLEAN
      out.v = cell.v === true
      break
    case 'n':
      out.t = CellValueType.NUMBER
      out.v = typeof cell.v === 'number' ? cell.v : Number(cell.v)
      break
    case 'e':
      // Excel error strings like #REF! — display verbatim, no formula engine.
      out.t = CellValueType.FORCE_STRING
      out.v = String(cell.w ?? cell.v ?? '')
      break
    case 's':
    case 'd':
    case 'z':
    default:
      // String / date / stub. SheetJS renders dates as ISO strings in `w`
      // when cellDates is on; otherwise `v` is the serial number. Use `w`
      // (formatted text) when present so users see what Excel showed.
      out.t = CellValueType.STRING
      out.v = cell.w ?? (cell.v !== undefined ? String(cell.v) : '')
      break
  }
  // Formula: keep as text; Univer's formula engine computes on load.
  if (typeof cell.f === 'string' && cell.f !== '') {
    out.f = cell.f
  }
  return out
}

/** Convert SheetJS `!cols` (ColInfo[]) to Univer columnData (sparse, by index). */
function convertCols(cols: XLSX.ColInfo[] | undefined): IObjectArrayPrimitiveType<Partial<IColumnData>> {
  if (!Array.isArray(cols)) return {}
  const out: IObjectArrayPrimitiveType<Partial<IColumnData>> = {}
  cols.forEach((info, index) => {
    if (info === undefined || info === null) return
    const w = info.wpx ?? (info.width !== undefined ? Math.round(info.width * 7) : undefined)
    out[index] = {
      ...(w !== undefined ? { w } : {}),
      ...(info.hidden === true ? { hd: BooleanNumber.TRUE } : {}),
    }
  })
  return out
}

/** Convert SheetJS `!rows` (RowInfo[]) to Univer rowData (sparse, by index). */
function convertRows(rows: XLSX.RowInfo[] | undefined): IObjectArrayPrimitiveType<Partial<IRowData>> {
  if (!Array.isArray(rows)) return {}
  const out: IObjectArrayPrimitiveType<Partial<IRowData>> = {}
  rows.forEach((info, index) => {
    if (info === undefined || info === null) return
    const h = info.hpx ?? info.hpt
    out[index] = {
      ...(h !== undefined ? { h: Math.round(h) } : {}),
      ...(info.hidden === true ? { hd: BooleanNumber.TRUE } : {}),
    }
  })
  return out
}

/**
 * Resolve the worksheet dimensions. SheetJS's `!ref` is authoritative when
 * present; otherwise fall back to the max row/col we actually wrote, with a
 * small floor so the canvas has somewhere to paint.
 */
function dimensions(
  ws: XLSX.WorkSheet,
  cellData: IObjectMatrixPrimitiveType<ICellData>,
): { rowCount: number; columnCount: number } {
  const ref = ws['!ref']
  if (typeof ref === 'string' && ref !== '') {
    const range = decodeRange(ref)
    return {
      rowCount: Math.max(range.endRow + 1, 20),
      columnCount: Math.max(range.endColumn + 1, 10),
    }
  }
  let maxRow = -1
  let maxCol = -1
  for (const r of Object.keys(cellData)) {
    const rowNum = Number(r)
    if (rowNum > maxRow) maxRow = rowNum
    const rowObj = cellData[rowNum]!
    for (const c of Object.keys(rowObj)) {
      const colNum = Number(c)
      if (colNum > maxCol) maxCol = colNum
    }
  }
  return {
    rowCount: Math.max(maxRow + 1, 20),
    columnCount: Math.max(maxCol + 1, 10),
  }
}

/** Decode an A1-style address (e.g. "AB12") to 0-indexed {row, col}. */
function decodeAddr(addr: string): { r: number; c: number } {
  let col = 0
  let i = 0
  while (i < addr.length) {
    const ch = addr.charCodeAt(i)
    if (ch >= 65 && ch <= 90) {
      col = col * 26 + (ch - 64)
      i += 1
    } else {
      break
    }
  }
  const row = Number(addr.slice(i)) - 1
  return { r: row, c: col - 1 }
}

/** Decode an A1-style range (e.g. "A1:AB12") to 0-indexed bounds. */
function decodeRange(range: string): { startRow: number; endRow: number; startColumn: number; endColumn: number } {
  const at = range.indexOf(':')
  const start = decodeAddr(at === -1 ? range : range.slice(0, at))
  const end = at === -1 ? start : decodeAddr(range.slice(at + 1))
  return { startRow: start.r, endRow: end.r, startColumn: start.c, endColumn: end.c }
}
