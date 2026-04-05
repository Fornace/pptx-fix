/**
 * pptx-fix — Fix PPTX files for macOS QuickLook / iOS preview rendering.
 *
 * Transforms operate on raw XML (preserving all unknown elements) so the
 * output is a valid PPTX that looks correct in Apple's OfficeImport pipeline.
 */

export { fix, type FixOptions, type FixResult } from "./writer.js";
export { analyze, type Issue } from "./analyze.js";
export { type TransformName, ALL_TRANSFORMS } from "./transforms/index.js";
