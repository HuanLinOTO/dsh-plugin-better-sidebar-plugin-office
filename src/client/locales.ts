/**
 * Copy for the Office previewers. Follows the DSH i18n convention: the
 * dictionaries are registered into the DSH locale registry under {@link NS}
 * (en + zh carry identical key sets), and each key is also translated into
 * the 19 better-locale override languages in `./dictionaries.ts`. The apply
 * function attaches the DSH locale service via {@link attachLocale} (and the
 * better-locale override store via {@link attachBetterLocale}), so the
 * module-level {@link t} follows DSH's active locale instead of the browser
 * language — the theme better-locale patches LocaleRuntime.lookup is
 * mirrored here so the viewer chrome switches too.
 */

/** All copy keys for the dsh-better-sidebar-plugin-office namespace. */
export type OfficeKey =
  | 'loading'
  | 'downloadToView'
  | 'previousSlide'
  | 'nextSlide'
  | 'zoom'
  | 'zoomHint'
  | 'viewerDocx'
  | 'viewerXlsx'
  | 'viewerPptx'

/** Locale namespace id (matches the cordis.patch.yml plugin id). */
export const NS = 'dsh-better-sidebar-plugin-office'

/** English dictionary. */
export const en: Record<OfficeKey, string> = {
  loading: 'Loading…',
  downloadToView: 'Download to view',
  previousSlide: 'Previous',
  nextSlide: 'Next',
  zoom: 'Zoom',
  zoomHint: 'Alt + wheel',
  viewerDocx: 'Word',
  viewerXlsx: 'Excel',
  viewerPptx: 'PowerPoint',
}

/** Chinese dictionary (key-set-equal to en, enforced by the type annotation). */
export const zh: Record<OfficeKey, string> = {
  loading: '加载中…',
  downloadToView: '下载查看',
  previousSlide: '上一页',
  nextSlide: '下一页',
  zoom: '缩放',
  zoomHint: 'Alt + 滚轮',
  viewerDocx: 'Word 文档',
  viewerXlsx: 'Excel 表格',
  viewerPptx: 'PPT 演示',
}

/**
 * The DSH locale service face (mirror of `@deepseek-ai/dsh-client-locale`'s
 * LocaleRuntime — only the slice the copy needs). Attached by the apply
 * function via {@link attachLocale}; absent in standalone/test compositions,
 * where the browser language is used instead.
 */
let localeService: { getSnapshot(): { active: string } } | undefined

/**
 * The better-locale override store face (mirror of `BetterLocaleStore`'s
 * `getOverride`). Attached by the apply function via
 * {@link attachBetterLocale}; absent when dsh-plugin-better-locale is not
 * installed, in which case the zh/en chain runs unchanged.
 */
let betterLocaleStore: { getOverride(dshActive: string, ns: string, key: string): string | undefined } | undefined

/** Attach (or detach, with undefined) the DSH locale service. */
export function attachLocale(service: typeof localeService): void {
  localeService = service
}

/** Attach (or detach, with undefined) the better-locale override store. */
export function attachBetterLocale(store: typeof betterLocaleStore): void {
  betterLocaleStore = store
}

/**
 * The active locale id ('zh' | 'en'): the DSH locale service's snapshot when
 * attached, else the browser language.
 */
function activeLocale(): string {
  return localeService?.getSnapshot().active
    ?? (typeof navigator !== 'undefined' ? navigator.language : '')
    ?? 'en'
}

/** Translate a copy key in the active locale (zh → zh, else en). */
export function t(key: OfficeKey): string {
  // A better-locale override (e.g. ja) wins when an override is active AND
  // DSH's active locale is 'en' (the override borrows the English slot);
  // `getOverride` returns undefined otherwise and the zh/en chain runs.
  const dshActive = localeService?.getSnapshot().active ?? ''
  const override = betterLocaleStore?.getOverride(dshActive, NS, key)
  if (override !== undefined) return override
  const dict = activeLocale().toLowerCase().startsWith('zh') ? zh : en
  return dict[key]
}