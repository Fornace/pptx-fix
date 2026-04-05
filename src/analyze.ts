/**
 * Analyze a PPTX for issues that will cause QuickLook rendering problems.
 * Delegates to quicklook-pptx-renderer's linter — single source of truth.
 */

export { lint as analyze, formatIssues, type LintResult, type LintIssue } from "quicklook-pptx-renderer";
