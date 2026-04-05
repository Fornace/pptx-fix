/**
 * pptx-fix — Fix PPTX files for macOS QuickLook / iOS preview rendering.
 *
 * Transforms operate on raw XML (preserving all unknown elements) so the
 * output is a valid PPTX that looks correct in Apple's OfficeImport pipeline.
 */
export { fix } from "./writer.js";
export { analyze } from "./analyze.js";
export { ALL_TRANSFORMS } from "./transforms/index.js";
