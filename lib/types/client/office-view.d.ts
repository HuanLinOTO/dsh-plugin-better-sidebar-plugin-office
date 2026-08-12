import { type SessionScope } from './urls.ts';
import '@univerjs/preset-sheets-core/lib/index.css';
/** Shared props. */
interface OfficeViewProps {
    scope: SessionScope;
    path: string;
    title: string;
}
/**
 * Render a .docx file via docx-preview. The library renders into a container
 * div (no canvas); images and styles are inlined. Unmounting clears the
 * container's innerHTML — docx-preview has no dispose API, but tearing down
 * the DOM is enough.
 */
export declare function DocxView(props: OfficeViewProps): JSX.Element;
/**
 * Render a .xlsx file via Univer. The sheets preset creates a canvas-based
 * spreadsheet (formula bar, sheet tabs, formula engine) sized to its
 * container, so the host fills the pane. Unmounting calls `univer.dispose()`
 * — without it the canvas, workers, and DOM listeners leak (mirrors the
 * xterm dispose discipline in the better-sidebar TerminalView).
 */
export declare function XlsxView(props: OfficeViewProps): JSX.Element;
export {};
