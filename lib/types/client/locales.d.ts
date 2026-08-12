/**
 * Minimal zh/en copy for the office previewers. The plugin registers its
 * viewer titles and the preview UIs with these strings, reading the browser
 * language directly (the DSH locale preference and the browser language
 * agree in practice; both default from navigator.language anyway).
 */
declare const zh: {
    loading: string;
    downloadToView: string;
    previousSlide: string;
    nextSlide: string;
    zoom: string;
    zoomHint: string;
    viewerDocx: string;
    viewerXlsx: string;
    viewerPptx: string;
};
/** Translate a copy key in the browser's language (zh → zh, else en). */
export type CopyKey = keyof typeof zh;
/** Translate a copy key. */
export declare function t(key: CopyKey): string;
export {};
