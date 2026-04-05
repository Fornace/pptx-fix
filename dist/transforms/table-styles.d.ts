/**
 * table-styles transform — Resolve table style conditional formatting into
 * explicit borders on every <a:tcPr>.
 *
 * Problem: python-pptx and other generators set a tableStyleId but don't
 * inline the actual border/fill properties. PowerPoint resolves these at
 * render time via the table style XML. OfficeImport (QuickLook) does NOT —
 * it only sees explicit properties, so tables render without borders/styling.
 *
 * Fix: Read the table style definition, resolve bandRow/firstRow/lastCol
 * conditional formatting, and write explicit <a:ln> on every <a:tcPr>.
 */
import type { Transform } from "./index.js";
export declare const tableStyles: Transform;
