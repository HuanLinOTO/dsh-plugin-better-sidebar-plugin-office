/**
 * Registration tests: the three Office viewer descriptors must match the
 * better-sidebar built-ins they replace (ids docx/xlsx/pptx, exts, mediaUrl
 * strategy, priority 0) so existing files route the same way.
 */
import { describe, expect, it } from 'vitest'
import { officeViewers } from '../src/client/index.tsx'

describe('officeViewers descriptors', () => {
  it('registers exactly the three Office viewers', () => {
    const viewers = officeViewers()
    expect(viewers.map(v => v.id).sort()).toEqual(['docx', 'pptx', 'xlsx'])
  })

  it('claims docx/xlsx/pptx through mediaUrl at priority 0', () => {
    for (const v of officeViewers()) {
      expect(v.fetchStrategy).toBe('mediaUrl')
      expect(v.priority ?? 0).toBe(0)
      expect(v.component).toBeTypeOf('function')
    }
    const byId = Object.fromEntries(officeViewers().map(v => [v.id, v]))
    expect(byId.docx?.exts).toEqual(['docx'])
    expect(byId.xlsx?.exts).toEqual(['xlsx'])
    expect(byId.pptx?.exts).toEqual(['pptx'])
  })

  it('carries the declarative settings surface (title + icon)', () => {
    for (const v of officeViewers()) {
      expect(v.title, v.id).toBeDefined()
      expect(v.icon, v.id).toBeDefined()
    }
  })
})
