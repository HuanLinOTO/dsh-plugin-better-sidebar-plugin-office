/** High-fidelity browser-native PPTX preview with slide navigation. */
import { useEffect, useRef, useState } from 'react'
import { downloadUrl, mediaUrl, type SessionScope } from './urls.ts'
import { t } from './locales.ts'
import css from './office.module.css'

type PptxViewerInstance = import('@aiden0z/pptx-renderer').PptxViewer

type LoadState =
  | { status: 'loading' }
  | { status: 'ready'; count: number }
  | { status: 'error'; message: string }

export function PptxView(props: { scope: SessionScope; path: string; title: string }) {
  const { scope, path, title } = props
  const hostRef = useRef<HTMLDivElement>(null)
  const viewerRef = useRef<PptxViewerInstance | null>(null)
  const [load, setLoad] = useState<LoadState>({ status: 'loading' })
  const [slide, setSlide] = useState(0)
  const [navigating, setNavigating] = useState(false)

  useEffect(() => {
    const controller = new AbortController()
    const host = hostRef.current
    if (host === null) return
    setLoad({ status: 'loading' })
    setSlide(0)
    void (async () => {
      try {
        const response = await fetch(mediaUrl(scope, path), { signal: controller.signal })
        if (!response.ok) throw new Error(`HTTP ${response.status}`)
        const bytes = await response.arrayBuffer()
        if (controller.signal.aborted) return
        const { PptxViewer, RECOMMENDED_ZIP_LIMITS } = await import('@aiden0z/pptx-renderer')
        if (controller.signal.aborted) return
        const viewer = await PptxViewer.open(bytes, host, {
          renderMode: 'slide',
          fitMode: 'contain',
          lazyMedia: true,
          lazySlides: true,
          pdfjs: false,
          signal: controller.signal,
          zipLimits: RECOMMENDED_ZIP_LIMITS,
          onSlideChange: (index) => {
            if (!controller.signal.aborted) setSlide(index)
          },
        })
        if (controller.signal.aborted) {
          viewer.destroy()
          return
        }
        viewerRef.current = viewer
        setSlide(viewer.currentSlideIndex)
        setLoad({ status: 'ready', count: viewer.slideCount })
      } catch (error) {
        if (controller.signal.aborted) return
        try {
          viewerRef.current?.destroy()
        } catch {
          // Partially initialized viewer — ignore disposal errors.
        }
        viewerRef.current = null
        host.innerHTML = ''
        setLoad({ status: 'error', message: error instanceof Error ? error.message : String(error) })
      }
    })()
    return () => {
      controller.abort()
      try {
        viewerRef.current?.destroy()
      } catch {
        // Already destroyed.
      }
      viewerRef.current = null
      host.innerHTML = ''
    }
  }, [scope.sessionId, scope.cwd, path])

  const navigate = (target: number): void => {
    const viewer = viewerRef.current
    if (viewer === null || load.status !== 'ready' || navigating) return
    const next = Math.max(0, Math.min(load.count - 1, target))
    if (next === slide) return
    setNavigating(true)
    void viewer.goToSlide(next, { behavior: 'auto' }).then(() => {
      setSlide(viewer.currentSlideIndex)
    }).catch((error: unknown) => {
      setLoad({ status: 'error', message: error instanceof Error ? error.message : String(error) })
    }).finally(() => {
      setNavigating(false)
    })
  }

  return (
    <div className={css.editorPptx} aria-label={title}>
      <div className={css.editorPptxToolbar}>
        <button
          type="button"
          className={css.editorPptxButton}
          disabled={load.status !== 'ready' || slide <= 0 || navigating}
          onClick={() => { navigate(slide - 1) }}
        >
          {t('previousSlide')}
        </button>
        <span className={css.editorPptxPosition}>
          {load.status === 'ready' ? `${slide + 1} / ${load.count}` : '– / –'}
        </span>
        <button
          type="button"
          className={css.editorPptxButton}
          disabled={load.status !== 'ready' || slide >= load.count - 1 || navigating}
          onClick={() => { navigate(slide + 1) }}
        >
          {t('nextSlide')}
        </button>
        <a className={css.editorDownloadLink} href={downloadUrl(scope, path)} download>
          {t('downloadToView')}
        </a>
      </div>
      <div className={css.editorPptxStage}>
        <div className={css.editorPptxHost} ref={hostRef} />
        {load.status !== 'ready' && (
          <div className={css.editorOfficeOverlay}>
            {load.status === 'loading' && <div className={css.editorPlaceholder}>{t('loading')}</div>}
            {load.status === 'error' && <div className={css.editorError}>{load.message}</div>}
          </div>
        )}
      </div>
    </div>
  )
}
