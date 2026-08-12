import type { ClientContext } from '@deepseek-ai/dsh-client-runtime/client';
import type { FileViewerDescriptor } from 'dsh-better-sidebar';
/** Services required before mounting: better-sidebar's client service. */
export declare const inject: string[];
/** The three Office viewer descriptors (id / exts match the former built-ins). */
export declare function officeViewers(): readonly FileViewerDescriptor[];
/**
 * Client plugin body.
 * @param ctx - the client cordis context (betterSidebar service).
 */
export declare function apply(ctx: ClientContext): void;
