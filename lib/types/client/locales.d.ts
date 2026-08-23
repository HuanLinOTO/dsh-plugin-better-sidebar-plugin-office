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
export type OfficeKey = 'loading' | 'downloadToView' | 'previousSlide' | 'nextSlide' | 'zoom' | 'zoomHint' | 'viewerDocx' | 'viewerXlsx' | 'viewerPptx';
/** Locale namespace id (matches the cordis.patch.yml plugin id). */
export declare const NS = "dsh-better-sidebar-plugin-office";
/** English dictionary. */
export declare const en: Record<OfficeKey, string>;
/** Chinese dictionary (key-set-equal to en, enforced by the type annotation). */
export declare const zh: Record<OfficeKey, string>;
/**
 * The DSH locale service face (mirror of `@deepseek-ai/dsh-client-locale`'s
 * LocaleRuntime — only the slice the copy needs). Attached by the apply
 * function via {@link attachLocale}; absent in standalone/test compositions,
 * where the browser language is used instead.
 */
declare let localeService: {
    getSnapshot(): {
        active: string;
    };
} | undefined;
/**
 * The better-locale override store face (mirror of `BetterLocaleStore`'s
 * `getOverride`). Attached by the apply function via
 * {@link attachBetterLocale}; absent when dsh-plugin-better-locale is not
 * installed, in which case the zh/en chain runs unchanged.
 */
declare let betterLocaleStore: {
    getOverride(dshActive: string, ns: string, key: string): string | undefined;
} | undefined;
/** Attach (or detach, with undefined) the DSH locale service. */
export declare function attachLocale(service: typeof localeService): void;
/** Attach (or detach, with undefined) the better-locale override store. */
export declare function attachBetterLocale(store: typeof betterLocaleStore): void;
/** Translate a copy key in the active locale (zh → zh, else en). */
export declare function t(key: OfficeKey): string;
export {};
