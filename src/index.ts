/**
 * pptx-fix — Fix PPTX files for macOS QuickLook / iOS preview rendering.
 *
 * Detection delegates to quicklook-pptx-renderer's linter.
 * Transforms operate on raw XML (preserving all unknown elements) so the
 * output is a valid PPTX that looks correct in Apple's OfficeImport pipeline.
 */

export { fix, type FixOptions, type FixResult } from "./writer.js";
export { analyze, formatIssues, type LintResult, type LintIssue, type LintFix } from "./analyze.js";
export { type TransformName, ALL_TRANSFORMS } from "./transforms/index.js";
export {
  FONT_METRICS, FONT_SUBSTITUTIONS, SUPPORTED_GEOMETRIES, GEOMETRY_FALLBACKS,
  findClosestFont, widthDelta,
  APPLE_SYSTEM_FONTS, APPLE_SYSTEM_FONT_LIST,
  MACOS_SYSTEM_FONTS, MACOS_SYSTEM_FONT_LIST,
  type FontMetrics, type FontMatch, type FontCategory,
} from "quicklook-pptx-renderer";
