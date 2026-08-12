/**
 * Office preview components: {@link DocxView} renders .docx via docx-preview
 * (preserved styles/images/tables), {@link XlsxView} renders .xlsx via the
 * Univer sheets preset (data + formulas + formatting). Both load the file
 * bytes through the better-sidebar `/sidebar/file` media route and own their
 * library lifecycle (dispose on unmount so canvases/workers don't leak —
 * mirrors the better-sidebar TerminalView dispose discipline).
 *
 * Errors degrade to the same download-button affordance the binary
 * placeholder uses, so a corrupted / encrypted / oversized file always
 * leaves the user with a way to get the file.
 *
 * Both libraries pull in several MB; tsdown's `codeSplitting: false` inlines
 * them into the single client bundle (the CJS __ModuleLoader__ wrapper's
 * `require` cannot load relative chunk URLs in the browser, so dynamic
 * import() collapses to a synchronous inline — see tsdown.config.ts).
 */
import { useEffect, useRef, useState } from 'react'
import { downloadUrl, mediaUrl, type SessionScope } from './urls.ts'
import { t } from './locales.ts'
import { xlsxWorkbookToUniver } from './xlsx-to-univer.ts'
import css from './office.module.css'
// Univer's stylesheets ride the same dsh-css-inline pipeline as xterm's CSS
// (one <style data-plugin-css> tag, idempotent). Static import so the styles
// are present before the first .xlsx opens.
import '@univerjs/preset-sheets-core/lib/index.css'

/** Loading / ready / error state shared by both views. */
type LoadState =
  | { status: 'loading' }
  | { status: 'ready' }
  | { status: 'error'; message: string }

/** Shared props. */
interface OfficeViewProps {
  scope: SessionScope
  path: string
  title: string
}

/**
 * Render a .docx file via docx-preview. The library renders into a container
 * div (no canvas); images and styles are inlined. Unmounting clears the
 * container's innerHTML — docx-preview has no dispose API, but tearing down
 * the DOM is enough.
 */
export function DocxView(props: OfficeViewProps): JSX.Element {
  const { scope, path, title } = props
  const viewportRef = useRef<HTMLDivElement>(null)
  const wrapRef = useRef<HTMLDivElement>(null)
  const [load, setLoad] = useState<LoadState>({ status: 'loading' })
  const [zoom, setZoom] = useState(100)

  useEffect(() => {
    let cancelled = false
    const container = viewportRef.current
    const wrap = wrapRef.current
    if (container === null || wrap === null) return
    setZoom(100)
    void (async () => {
      try {
        const response = await fetch(mediaUrl(scope, path))
        if (cancelled) return
        if (!response.ok) {
          throw new Error(`HTTP ${response.status}`)
        }
        const buf = await response.arrayBuffer()
        if (cancelled) return
        // docx-preview ships its own CSS through the className option; the
        // wrapper div scopes its render output.
        const { renderAsync } = await import('docx-preview')
        await renderAsync(buf, wrap, undefined, {
          className: 'docx',
          inWrapper: true,
          ignoreWidth: false,
          ignoreHeight: false,
          breakPages: true,
          experimental: false,
        })
        if (!cancelled) setLoad({ status: 'ready' })
      } catch (error) {
        if (!cancelled) {
          setLoad({ status: 'error', message: error instanceof Error ? error.message : String(error) })
        }
      }
    })()
    return () => {
      cancelled = true
      // Tear down the rendered DOM so a reopen starts clean.
      if (wrap !== null) wrap.innerHTML = ''
    }
  }, [scope.sessionId, scope.cwd, path])

  useEffect(() => {
    const viewport = viewportRef.current
    if (viewport === null) return
    const onWheel = (event: WheelEvent): void => {
      if (!event.altKey) return
      event.preventDefault()
      const delta = event.deltaY < 0 ? 10 : -10
      setZoom(current => Math.max(50, Math.min(200, current + delta)))
    }
    viewport.addEventListener('wheel', onWheel, { passive: false })
    return () => { viewport.removeEventListener('wheel', onWheel) }
  }, [])

  return (
    <div className={css.editorDocx}>
      <div className={css.editorDocxViewport} ref={viewportRef}>
        {load.status === 'loading' && <div className={css.editorPlaceholder}>{t('loading')}</div>}
        {load.status === 'error' && <BinaryFallback scope={scope} path={path} message={load.message} />}
        {load.status !== 'error' && (
          <div
            className={css.editorDocxWrap}
            ref={wrapRef}
            aria-label={title}
            style={{ zoom: zoom / 100 }}
          />
        )}
      </div>
      <div className={css.editorDocxZoom}>
        <span className={css.editorDocxZoomHint}>{t('zoomHint')}</span>
        <input
          className={css.editorDocxZoomRange}
          type="range"
          min={50}
          max={200}
          step={10}
          value={zoom}
          aria-label={t('zoom')}
          onChange={(event) => { setZoom(Number(event.currentTarget.value)) }}
        />
        <span className={css.editorDocxZoomValue}>{zoom}%</span>
      </div>
    </div>
  )
}

/**
 * Render a .xlsx file via Univer. The sheets preset creates a canvas-based
 * spreadsheet (formula bar, sheet tabs, formula engine) sized to its
 * container, so the host fills the pane. Unmounting calls `univer.dispose()`
 * — without it the canvas, workers, and DOM listeners leak (mirrors the
 * xterm dispose discipline in the better-sidebar TerminalView).
 */
export function XlsxView(props: OfficeViewProps): JSX.Element {
  const { scope, path, title } = props
  const hostRef = useRef<HTMLDivElement>(null)
  const univerRef = useRef<{ dispose: () => void } | null>(null)
  const [load, setLoad] = useState<LoadState>({ status: 'loading' })

  useEffect(() => {
    let cancelled = false
    const host = hostRef.current
    if (host === null) return
    void (async () => {
      try {
        const response = await fetch(mediaUrl(scope, path))
        if (cancelled) return
        if (!response.ok) {
          throw new Error(`HTTP ${response.status}`)
        }
        const buf = await response.arrayBuffer()
        if (cancelled) return

        // Dynamic imports — collapsed into the bundle by codeSplitting:false.
        // The dynamic form keeps the source readable: each lib is only pulled
        // in when an .xlsx is actually opened (semantically; the bundle still
        // contains all of them).
        const XLSX = await import('xlsx')
        const { createUniver, LocaleType, mergeLocales } = await import('@univerjs/presets')
        const { UniverSheetsCorePreset } = await import('@univerjs/preset-sheets-core')
        // Locales pick the browser language; falls back to en-US.
        const isZh = typeof navigator !== 'undefined' && navigator.language.toLowerCase().startsWith('zh')
        const localePack = await (isZh
          ? import('@univerjs/preset-sheets-core/locales/zh-CN').then(m => m.default).catch(() => null)
          : import('@univerjs/preset-sheets-core/locales/en-US').then(m => m.default).catch(() => null))

        const wb = XLSX.read(buf, { type: 'array' })
        const locale = isZh ? LocaleType.ZH_CN : LocaleType.EN_US
        const workbookData = xlsxWorkbookToUniver(wb, '0.25.1', locale)

        if (cancelled) return
        const { univer, univerAPI } = createUniver({
          locale,
          locales: localePack !== null ? { [locale]: mergeLocales(localePack) } : {},
          presets: [UniverSheetsCorePreset({ container: host })],
        })
        univerRef.current = univer
        univerAPI.createWorkbook(workbookData)
        if (!cancelled) setLoad({ status: 'ready' })
      } catch (error) {
        if (!cancelled) {
          try {
            univerRef.current?.dispose()
          } catch {
            // Partially initialized instance — ignore disposal errors.
          }
          univerRef.current = null
          host.innerHTML = ''
          setLoad({ status: 'error', message: error instanceof Error ? error.message : String(error) })
        }
      }
    })()
    return () => {
      cancelled = true
      // Critical: dispose the Univer instance (canvas + workers + listeners).
      try {
        univerRef.current?.dispose()
      } catch {
        // Already torn down — ignore.
      }
      univerRef.current = null
      // Clear the host in case dispose left DOM behind.
      if (host !== null) host.innerHTML = ''
    }
  }, [scope.sessionId, scope.cwd, path])

  return (
    <div className={css.editorXlsx} aria-label={title}>
      {/* Univer exclusively owns this element's descendants. React only owns
          the sibling overlay, so Univer cannot remove a React child. */}
      <div className={css.editorUniverHost} ref={hostRef} />
      {load.status !== 'ready' && (
        <div className={css.editorOfficeOverlay}>
          {load.status === 'loading' && <div className={css.editorPlaceholder}>{t('loading')}</div>}
          {load.status === 'error' && <BinaryFallback scope={scope} path={path} message={load.message} />}
        </div>
      )}
    </div>
  )
}

/**
 * The shared error / fallback affordance: the failure reason plus a download
 * link, so the user always has a path to the file. Used by both DocxView and
 * XlsxView.
 */
function BinaryFallback(props: { scope: SessionScope; path: string; message: string }): JSX.Element {
  const { scope, path, message } = props
  return (
    <div className={css.editorBinary}>
      <span className={css.editorBinaryNotice}>{message}</span>
      <a className={css.editorDownloadLink} href={downloadUrl(scope, path)} download>
        {t('downloadToView')}
      </a>
    </div>
  )
}
