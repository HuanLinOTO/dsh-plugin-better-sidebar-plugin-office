/**
 * tsdown build for @huanlin/dsh-plugin-better-sidebar-plugin-office: the host
 * stub (lib/index.js, ESM node) plus one browser client bundle
 * (lib/client.js, CJS closure factory).
 *
 * The client bundle replicates the official DSH client-bundle preset (mirror
 * of dsh-aigc-canvas / dsh-better-sidebar's tsdown configs): externals
 * resolve through the loader module table at runtime (react + cordis +
 * ui-primitives), everything else — including the heavy Office render
 * libraries (docx-preview, Univer, SheetJS, pptx-renderer) — inlines into the
 * single bundle. Each artifact registers itself via
 * window.__ModuleLoader__.load({id, factory}) with the (require) => exports
 * CJS closure shape.
 *
 * The Office libs need the same browser-entry aliases the original
 * better-sidebar bundle used: Rolldown does not currently honor some
 * packages' `browser` remaps after CJS lowering (SheetJS/JSZip leave Node
 * builtin require() calls that the DSH module table refuses), and
 * pptx-renderer auto-discovers its optional PDF.js fallback via
 * import.meta.resolve (PptxView passes pdfjs:false, so it has no resolver).
 *
 * Types ship from tsc -p tsconfig.build.json, not from tsdown.
 */
import { readFile } from 'node:fs/promises'
import { basename, dirname, relative, resolve as resolvePath, sep } from 'node:path'
import { fileURLToPath } from 'node:url'
import { createRequire } from 'node:module'
import type { UserConfig } from 'tsdown'

const require = createRequire(import.meta.url)

const REPOSITORY_ROOT = fileURLToPath(new URL('.', import.meta.url))

/** Bundle id (= package name; the client-modules compose keys on it). */
const CLIENT_ID = '@huanlin/dsh-plugin-better-sidebar-plugin-office'

/** Module specifiers the web shell shares into the frozen module table. */
const CLIENT_EXTERNALS = [
  'react',
  'react/jsx-runtime',
  'react-dom',
  'react-dom/client',
  '@deepseek-ai/cordis',
  '@deepseek-ai/dsh-client-ui-primitives',
  '@deepseek-ai/dsh-client-ui-primitives/client',
  '@deepseek-ai/dsh-client-ui-slots',
  '@deepseek-ai/dsh-client-ui-slots/client',
]

/** Browser-only standalone entries for the Office libraries whose package
 *  `browser` remaps Rolldown does not honor after CJS lowering. */
const DOCX_PREVIEW_ENTRY = require.resolve('docx-preview')
const JSZIP_BROWSER_ENTRY = resolvePath(
  dirname(require.resolve('jszip/package.json', { paths: [dirname(DOCX_PREVIEW_ENTRY)] })),
  'dist/jszip.min.js',
)
const XLSX_BROWSER_ENTRY = resolvePath(
  dirname(require.resolve('xlsx/package.json')),
  'dist/xlsx.full.min.js',
)

/** Virtual-id wrapper keeping module CSS away from tsdown's own css pipeline. */
const CSS_VIRTUAL_PREFIX = '\0dsh-css:'
const CSS_VIRTUAL_SUFFIX = '.mjs'

/** The style-injection prologue shared by module css and plain css loads. */
function injectTag(pluginId: string, fileId: string, cssText: string): string {
  const tagId = `${pluginId}/${basename(fileId)}`
  return [
    `const css = ${JSON.stringify(cssText)};`,
    `const tagId = ${JSON.stringify(tagId)};`,
    `if (typeof document !== 'undefined' && document.querySelector('style[data-plugin-css=' + JSON.stringify(tagId) + ']') === null) {`,
    `  const tag = document.createElement('style');`,
    `  tag.dataset.plugin = ${JSON.stringify(pluginId)};`,
    `  tag.dataset.pluginCss = tagId;`,
    `  tag.textContent = css;`,
    `  document.head.appendChild(tag);`,
    `}`,
  ].join('\n')
}

/** Simple CSS Modules transform (mirror of dsh-aigc-canvas's css-modules). */
function transformCssModules(filename: string, source: Buffer): { classMap: Record<string, string>; cssText: string } {
  const hash = Array.from(filename).reduce((acc, ch) => ((acc << 5) - acc + ch.charCodeAt(0)) | 0, 0).toString(36).replace('-', '')
  const cssText = source.toString('utf8')
  const classMap: Record<string, string> = {}
  const classPattern = /\.([a-zA-Z_][a-zA-Z0-9_-]*)/g
  let match: RegExpExecArray | null
  while ((match = classPattern.exec(cssText)) !== null) {
    const local = match[1]
    if (local !== undefined && classMap[local] === undefined) {
      classMap[local] = `${hash}_${local}`
    }
  }
  const transformedCss = cssText.replace(/\.([a-zA-Z_][a-zA-Z0-9_-]*)/g, (full, name: string) => {
    if (classMap[name] !== undefined) return `.${classMap[name]}`
    return full
  })
  return { classMap, cssText: transformedCss }
}

/** Rebase a physical lib-relative source onto the repository-shaped URL tree. */
function browserSourcePath(source: string, sourcemapPath: string): string {
  if (!source.startsWith('.')) return source
  const physicalSource = resolvePath(dirname(sourcemapPath), source)
  const repositoryPath = relative(REPOSITORY_ROOT, physicalSource).split(sep).join('/')
  return `../../../${repositoryPath}`
}

/** The host-half build (lib/index.js, ESM node). */
const hostConfig: UserConfig = {
  entry: { index: 'src/index.ts' },
  outDir: 'lib',
  format: ['esm'],
  platform: 'node',
  target: 'es2024',
  fixedExtension: false,
  dts: false,
  clean: true,
}

/** The client bundle build (lib/client.js, CJS closure factory). */
const clientConfig: UserConfig = {
  entry: { client: 'src/client/index.tsx' },
  outDir: 'lib',
  format: 'cjs',
  platform: 'browser',
  target: 'es2024',
  dts: false,
  sourcemap: true,
  clean: false,
  external: [...CLIENT_EXTERNALS],
  define: {
    'process.env.NODE_ENV': JSON.stringify(process.env.NODE_ENV ?? 'production'),
    'import.meta.env.MODE': JSON.stringify(process.env.NODE_ENV ?? 'production'),
    'import.meta.env': JSON.stringify({ MODE: process.env.NODE_ENV ?? 'production' }),
    // pptx-renderer only uses this to auto-discover its optional PDF.js
    // fallback. PptxView passes pdfjs:false, so browser CJS has no resolver.
    'import.meta.resolve': 'undefined',
  },
  alias: {
    jszip: JSZIP_BROWSER_ENTRY,
    xlsx: XLSX_BROWSER_ENTRY,
  },
  // CJS output otherwise makes some transitive packages (notably nanoid
  // through Univer Core) resolve their Node entry even though this bundle
  // runs in the browser. Keep browser conditional exports authoritative for
  // both source import() and generated require() edges.
  inputOptions: {
    resolve: {
      conditionNames: ['browser', 'import', 'require', 'default'],
    },
  },
  noExternal: (id: string) => (CLIENT_EXTERNALS.includes(id) ? undefined : true),
  plugins: [
    {
      name: 'dsh-css-inline',
      resolveId(source: string, importer: string | undefined) {
        if (!source.endsWith('.module.css') && !source.endsWith('.css')) return null
        // Relative/absolute paths resolve against the importer; bare
        // specifiers (e.g. '@univerjs/preset-sheets-core/lib/index.css')
        // resolve from the package through node's require resolution.
        let abs: string
        if (source.startsWith('.') || source.startsWith('/') || /^[A-Za-z]:[\\/]/.test(source)) {
          abs = importer === undefined ? source : resolvePath(dirname(importer), source)
        } else {
          abs = require.resolve(source)
        }
        return CSS_VIRTUAL_PREFIX + abs + CSS_VIRTUAL_SUFFIX
      },
      async load(virtualId: string) {
        if (!virtualId.startsWith(CSS_VIRTUAL_PREFIX)) return null
        const fileId = virtualId.slice(CSS_VIRTUAL_PREFIX.length, -CSS_VIRTUAL_SUFFIX.length)
        this.addWatchFile(fileId)
        const source = await readFile(fileId)
        if (fileId.endsWith('.module.css')) {
          const { classMap, cssText } = transformCssModules(fileId, source)
          return [
            injectTag(CLIENT_ID, fileId, cssText),
            `export default ${JSON.stringify(classMap)};`,
          ].join('\n')
        }
        return [
          injectTag(CLIENT_ID, fileId, source.toString('utf8')),
          'export default "";',
        ].join('\n')
      },
    },
  ],
  outputOptions: {
    entryFileNames: 'client.js',
    sourcemapPathTransform: browserSourcePath,
    banner: `window.__ModuleLoader__.load({ id: ${JSON.stringify(CLIENT_ID)}, factory: (require) => {`,
    footer: `return module.exports; } });`,
    intro: 'var module = { exports: {} }; var exports = module.exports;',
    // The CJS wrapper factory's `require` only resolves module-table entries
    // (react, cordis, ...); it cannot load relative chunk URLs in the browser.
    // Disable code splitting so dynamic import() inlines into the single
    // factory chunk (the Office preview libs — docx-preview, Univer — pull
    // in several MB but only when first opened, the trade-off for wrapper
    // safety).
    codeSplitting: false,
  },
}

export default [hostConfig, clientConfig] satisfies UserConfig[]
