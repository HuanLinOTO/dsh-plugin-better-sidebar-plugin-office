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
import type * as XLSX from 'xlsx';
import { type LocaleType, type IWorkbookData } from '@univerjs/presets';
/** SheetJS workbook (the slice of the module we use). */
type XLSXWorkbook = XLSX.WorkBook;
/**
 * Build a Univer workbook snapshot from a SheetJS workbook.
 *
 * @param wb - the parsed workbook from `XLSX.read(buf, { type: 'array' })`
 * @param appVersion - the Univer package version, written into `appVersion`
 *   (Univer requires the field but does not gate on it).
 * @param locale - the locale Univer should render in.
 */
export declare function xlsxWorkbookToUniver(wb: XLSXWorkbook, appVersion: string, locale: LocaleType): IWorkbookData;
export {};
