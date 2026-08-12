/**
 * File-viewer inventory glyphs for the three Office previewers, matching the
 * outline style of better-sidebar's own icon set (1.5px stroke, currentColor).
 */

/** Shared icon props (mirror of the ui-primitives IconProps shape). */
interface IconProps {
  size?: number
  className?: string
}

/** Word viewer glyph: a document frame with a "W". */
export const IconDocxOutline16 = ({ size = 16, className }: IconProps) => (
  <svg width={size} height={size} className={className} viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <path d="M3.5 1.5h6.5L13.5 5v9.5h-10z" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round" />
    <path d="M9.5 1.5V5h4" stroke="currentColor" strokeWidth="1.5" strokeLinejoin="round" />
    <path d="M6.2 13.4 7.4 10l1.2 3.4M7.4 10.6l-.35-1.1c-.2-.62.2-1.25.85-1.25h.2c.65 0 1.05.63.85 1.25l-.35 1.1" stroke="currentColor" strokeWidth="1.25" strokeLinecap="round" strokeLinejoin="round" />
    <path d="M8.75 10.6 9.2 9.4" stroke="currentColor" strokeWidth="1.25" strokeLinecap="round" />
  </svg>
)

/** Excel viewer glyph: a spreadsheet grid. */
export const IconXlsxOutline16 = ({ size = 16, className }: IconProps) => (
  <svg width={size} height={size} className={className} viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="1.5" y="2" width="13" height="12" rx="2" stroke="currentColor" strokeWidth="1.5" />
    <path d="M1.5 6h13M1.5 9.5h13M6 6v8M10.5 6v8" stroke="currentColor" strokeWidth="1.25" />
    <path d="m3.8 13.2 2-3M5.8 13.2l-2-3" stroke="currentColor" strokeWidth="1.25" strokeLinecap="round" />
  </svg>
)

/** PowerPoint viewer glyph: a chart with rising bars. */
export const IconPptxOutline16 = ({ size = 16, className }: IconProps) => (
  <svg width={size} height={size} className={className} viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg">
    <rect x="1.5" y="2.5" width="13" height="11" rx="2" stroke="currentColor" strokeWidth="1.5" />
    <path d="M4 12.5v-3M7 12.5V7M10 12.5V4.5M13 12.5v-1.5" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" />
  </svg>
)
