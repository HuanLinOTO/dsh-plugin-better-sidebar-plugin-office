/**
 * @huanlin/dsh-plugin-better-sidebar-plugin-office, node half.
 *
 * A stub. Every feature this plugin provides — the .docx/.xlsx/.pptx file
 * previewers — lives in the browser half, because each one renders inside
 * better-sidebar's editor host through `ctx.betterSidebar.registerFileViewer`
 * (a client-only service). The node half exists so the bundle's `main`
 * resolves cleanly and so a future version that needs host-side work (e.g.
 * a size/encryption probe route) has somewhere to land; for now, applying it
 * does nothing.
 *
 * @module @huanlin/dsh-plugin-better-sidebar-plugin-office
 */

/** Plugin configuration; intentionally empty — every knob lives client-side. */
export interface Config {}

/** No host services are required. */
export const inject: readonly string[] = []

/**
 * Apply the (empty) node half.
 * @param _ctx - host context, unused.
 * @param _config - see {@link Config}.
 */
export function apply(_ctx: unknown, _config: Config = {}): void {
  // Intentionally empty. See module doc.
}
