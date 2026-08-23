/**
 * File-viewer inventory glyphs for the three Office previewers, matching the
 * outline style of better-sidebar's own icon set (1.5px stroke, currentColor).
 */
/** Shared icon props (mirror of the ui-primitives IconProps shape). */
interface IconProps {
    size?: number;
    className?: string;
}
/** Word viewer glyph: a document frame with a "W". */
export declare const IconDocxOutline16: ({ size, className }: IconProps) => import("react").JSX.Element;
/** Excel viewer glyph: a spreadsheet grid. */
export declare const IconXlsxOutline16: ({ size, className }: IconProps) => import("react").JSX.Element;
/** PowerPoint viewer glyph: a chart with rising bars. */
export declare const IconPptxOutline16: ({ size, className }: IconProps) => import("react").JSX.Element;
export {};
