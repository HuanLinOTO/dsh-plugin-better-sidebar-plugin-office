/**
 * Client half of @huanlin/dsh-plugin-better-sidebar-plugin-office: registers
 * the three Office file viewers (docx / xlsx / pptx) through better-sidebar's
 * `ctx.betterSidebar.registerFileViewer`. The descriptors are identical in
 * shape to the built-ins they replace (same viewer ids, exts, mediaUrl
 * strategy, priority 0) so existing files route the same way, and the Side
 * card settings inventory shows them with their own title + icon.
 *
 * The heavy render libraries (docx-preview / Univer + SheetJS / pptx-renderer)
 * live in THIS bundle only, keeping better-sidebar's own client bundle small.
 */
import { createElement } from 'react'
import type { ClientContext } from '@deepseek-ai/dsh-client-runtime/client'
// Type-only: pulls better-sidebar's `declare module 'cordis'` Context merge
// (ctx.betterSidebar) and the FileViewerDescriptor type. Erased at build time,
// so it never hits the client-bundle purity gate.
import type {} from 'dsh-better-sidebar/client'
import type { FileViewerDescriptor } from 'dsh-better-sidebar'
import { DocxView, XlsxView } from './office-view.tsx'
import { PptxView } from './PptxView.tsx'
import { IconDocxOutline16, IconPptxOutline16, IconXlsxOutline16 } from './icons.tsx'
import { t } from './locales.ts'

/** Services required before mounting: better-sidebar's client service. */
export const inject = ['betterSidebar']

/** The three Office viewer descriptors (id / exts match the former built-ins). */
export function officeViewers(): readonly FileViewerDescriptor[] {
  return [
    {
      id: 'docx',
      title: () => t('viewerDocx'),
      icon: (size: number) => createElement(IconDocxOutline16, { size }),
      exts: ['docx'],
      fetchStrategy: 'mediaUrl',
      component: ({ scope, path, title }) => createElement(DocxView, { scope, path, title }),
    },
    {
      id: 'xlsx',
      title: () => t('viewerXlsx'),
      icon: (size: number) => createElement(IconXlsxOutline16, { size }),
      exts: ['xlsx'],
      fetchStrategy: 'mediaUrl',
      component: ({ scope, path, title }) => createElement(XlsxView, { scope, path, title }),
    },
    {
      id: 'pptx',
      title: () => t('viewerPptx'),
      icon: (size: number) => createElement(IconPptxOutline16, { size }),
      exts: ['pptx'],
      fetchStrategy: 'mediaUrl',
      component: ({ scope, path, title }) => createElement(PptxView, { scope, path, title }),
    },
  ]
}

/**
 * Client plugin body.
 * @param ctx - the client cordis context (betterSidebar service).
 */
export function apply(ctx: ClientContext): void {
  const betterSidebar = (ctx as unknown as {
    betterSidebar?: { registerFileViewer(descriptor: FileViewerDescriptor): () => void }
  }).betterSidebar
  if (betterSidebar === undefined) return
  // Register through the service; the disposer unregisters on fiber disposal
  // (HMR-safe). Skipped when better-sidebar is absent (the whole plugin is a
  // viewer provider for its editor).
  for (const viewer of officeViewers()) {
    ctx.effect(() => betterSidebar.registerFileViewer(viewer), `dsh-better-sidebar-plugin-office: viewer ${viewer.id}`)
  }
}
