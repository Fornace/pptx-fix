/**
 * XML round-trip writer: parse → mutate → serialize → repack ZIP.
 *
 * Transforms operate on fast-xml-parser parsed objects (preserving unknown elements),
 * then XMLBuilder serializes them back to XML. JSZip repacks the modified ZIP.
 */
import { type TransformName } from "./transforms/index.js";
export interface FixOptions {
    /** Apply only these transforms (default: all) */
    transforms?: TransformName[];
    /** Return a human-readable report of changes */
    report?: boolean;
}
export interface FixResult {
    buffer: Buffer;
    report?: string;
}
export declare function fix(pptxBuffer: Buffer, options?: FixOptions): Promise<FixResult>;
