/**
 * xlsx → Univer workbook conversion tests. The conversion function is pure
 * (no DOM, no canvas), so it is unit-testable without jsdom or Univer mounts.
 */
import { describe, expect, it } from 'vitest'
import { CellValueType, LocaleType } from '@univerjs/presets'
import type * as XLSX from 'xlsx'
import { xlsxWorkbookToUniver } from '../src/client/xlsx-to-univer.ts'

/** Build a minimal SheetJS workbook from a map of sheet name → cell rows. */
function makeWorkbook(sheets: Record<string, Record<string, { t: string; v: unknown; f?: string; w?: string }>>): XLSX.WorkBook {
  return {
    SheetNames: Object.keys(sheets),
    Sheets: Object.fromEntries(
      Object.entries(sheets).map(([name, cells]) => {
        const refs = Object.keys(cells)
        const ref = refs.length > 0
          ? `A1:${refs[refs.length - 1]}`
          : 'A1'
        return [name, { ...cells, '!ref': ref } as XLSX.WorkSheet]
      }),
    ),
    Props: {},
  }
}

describe('xlsxWorkbookToUniver', () => {
  it('preserves sheet names and order', () => {
    const wb = makeWorkbook({
      'Data': {},
      'Report': {},
    })
    const result = xlsxWorkbookToUniver(wb, '0.25.1', LocaleType.EN_US)
    expect(result.sheetOrder).toEqual(['sheet-0', 'sheet-1'])
    expect(result.sheets['sheet-0']?.name).toBe('Data')
    expect(result.sheets['sheet-1']?.name).toBe('Report')
  })

  it('maps typed cell values', () => {
    const wb = makeWorkbook({
      'Data': {
        A1: { t: 'n', v: 42 },
        B1: { t: 's', v: 'hello' },
        C1: { t: 'b', v: true },
      },
    })
    const result = xlsxWorkbookToUniver(wb, '0.25.1', LocaleType.EN_US)
    const cells = result.sheets['sheet-0']?.cellData ?? {}
    expect(cells[0]?.[0]?.t).toBe(CellValueType.NUMBER)
    expect(cells[0]?.[0]?.v).toBe(42)
    expect(cells[0]?.[1]?.t).toBe(CellValueType.STRING)
    expect(cells[0]?.[1]?.v).toBe('hello')
    expect(cells[0]?.[2]?.t).toBe(CellValueType.BOOLEAN)
    expect(cells[0]?.[2]?.v).toBe(true)
  })

  it('carries formulas as text for the Univer formula engine', () => {
    const wb = makeWorkbook({
      'Data': {
        A1: { t: 'n', v: 1, f: 'A1+B1' },
      },
    })
    const result = xlsxWorkbookToUniver(wb, '0.25.1', LocaleType.EN_US)
    const cell = result.sheets['sheet-0']?.cellData?.[0]?.[0]
    expect(cell?.f).toBe('A1+B1')
  })

  it('turns error cells into force-strings', () => {
    const wb = makeWorkbook({
      'Data': {
        A1: { t: 'e', v: null, w: '#REF!' },
      },
    })
    const result = xlsxWorkbookToUniver(wb, '0.25.1', LocaleType.EN_US)
    const cell = result.sheets['sheet-0']?.cellData?.[0]?.[0]
    expect(cell?.t).toBe(CellValueType.FORCE_STRING)
    expect(cell?.v).toBe('#REF!')
  })

  it('provides an empty-workbook placeholder', () => {
    const wb = makeWorkbook({})
    const result = xlsxWorkbookToUniver(wb, '0.25.1', LocaleType.EN_US)
    expect(result.sheetOrder).toEqual(['sheet-0'])
    expect(result.sheets['sheet-0']?.name).toBe('Sheet1')
  })
})
