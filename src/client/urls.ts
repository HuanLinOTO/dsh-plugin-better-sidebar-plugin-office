/**
 * Office-preview URL helpers for the better-sidebar `/sidebar/file` media
 * route. The better-sidebar host serves raw bytes for any path under the
 * session cwd; these builders mirror the route's contract so the viewer
 * components fetch exactly what the built-in previewers used to.
 */

/** One request's session scope (mirror of better-sidebar's SessionScope). */
export interface SessionScope {
  sessionId: string
  /** The session's working directory from the client list summary (optional). */
  cwd?: string
}

/** Absolute URL of the media route for one path (raw bytes). */
export function mediaUrl(scope: SessionScope, path: string): string {
  return fileUrl(scope, path, false)
}

/** Absolute URL of the download route (Content-Disposition: attachment). */
export function downloadUrl(scope: SessionScope, path: string): string {
  return fileUrl(scope, path, true)
}

/** Shared URL builder for the /sidebar/file route (media vs download). */
function fileUrl(scope: SessionScope, path: string, download: boolean): string {
  const params = new URLSearchParams({ sessionId: scope.sessionId, path })
  if (scope.cwd !== undefined && scope.cwd !== '') params.set('cwd', scope.cwd)
  if (download) params.set('download', '1')
  return `/sidebar/file?${params.toString()}`
}
