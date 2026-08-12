/**
 * Minimal zh/en copy for the office previewers. The plugin registers its
 * viewer titles and the preview UIs with these strings, reading the browser
 * language directly (the DSH locale preference and the browser language
 * agree in practice; both default from navigator.language anyway).
 */

const zh = {
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

const en: Record<keyof typeof zh, string> = {
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

/** Translate a copy key in the browser's language (zh → zh, else en). */
export type CopyKey = keyof typeof zh

/** Translate a copy key. */
export function t(key: CopyKey): string {
  const lang = typeof navigator !== 'undefined' ? navigator.language : 'en'
  const dict = lang.toLowerCase().startsWith('zh') ? zh : en
  return dict[key]
}
