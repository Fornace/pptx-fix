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

import type { Issue } from "../analyze.js";
import type { Transform, TransformContext, TransformResult } from "./index.js";

/** Walk the XML tree to find all table nodes. */
function findTables(node: any, path = ""): { table: any; path: string }[] {
  const results: { table: any; path: string }[] = [];
  if (!node || typeof node !== "object") return results;

  if (node.tbl) {
    const tbls = Array.isArray(node.tbl) ? node.tbl : [node.tbl];
    for (const t of tbls) results.push({ table: t, path: path + "/tbl" });
  }

  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    const children = Array.isArray(node[key]) ? node[key] : [node[key]];
    for (const child of children) {
      results.push(...findTables(child, path + "/" + key));
    }
  }
  return results;
}

/** Check if a table has a style ID but cells lack explicit borders. */
function tableNeedsFix(table: any): boolean {
  const tblPr = table.tblPr;
  if (!tblPr) return false;

  const styleId = tblPr["@_tblStyle"] ?? tblPr.tblStyle;
  if (!styleId) return false;

  // Check if any cell is missing explicit border definitions
  const rows = table.tr;
  if (!rows) return false;
  const rowArr = Array.isArray(rows) ? rows : [rows];

  for (const row of rowArr) {
    const cells = row.tc;
    if (!cells) continue;
    const cellArr = Array.isArray(cells) ? cells : [cells];
    for (const cell of cellArr) {
      const tcPr = cell.tcPr;
      if (!tcPr) return true; // No properties at all — definitely needs fix
      // Check if borders are defined
      const hasBorders = tcPr.lnL || tcPr.lnR || tcPr.lnT || tcPr.lnB;
      if (!hasBorders) return true;
    }
  }

  return false;
}

/** Look up a table style definition by ID from tableStyles.xml. */
function findStyleDef(tableStyleXml: any, styleId: string): any | undefined {
  if (!tableStyleXml) return undefined;
  const styleLst = tableStyleXml.tblStyleLst;
  if (!styleLst) return undefined;
  const styles = styleLst.tblStyle;
  if (!styles) return undefined;
  const arr = Array.isArray(styles) ? styles : [styles];
  return arr.find((s: any) => s["@_styleId"] === styleId);
}

/** Build a border element (<a:lnL>, etc.) from a style definition's tcStyle. */
function buildBorder(tcStyle: any, side: string): any | undefined {
  if (!tcStyle) return undefined;
  const borders = tcStyle.tcBdr;
  if (!borders) return undefined;
  const border = borders[side];
  if (!border) return undefined;
  // The border has a ln (line) child with fill and width
  return border.ln ?? border;
}

/** Apply table style borders to all cells in a table. */
function applyTableStyle(table: any, styleDef: any): string[] {
  const changes: string[] = [];
  const rows = table.tr;
  if (!rows) return changes;
  const rowArr = Array.isArray(rows) ? rows : [rows];

  const tblPr = table.tblPr ?? {};
  const hasFirstRow = tblPr["@_firstRow"] === "1" || tblPr["@_firstRow"] === "true";
  const hasLastRow = tblPr["@_lastRow"] === "1" || tblPr["@_lastRow"] === "true";
  const hasBandRow = tblPr["@_bandRow"] === "1" || tblPr["@_bandRow"] === "true";

  // Get style bands
  const wholeTbl = styleDef?.wholeTbl?.tcStyle;
  const firstRowStyle = hasFirstRow ? styleDef?.firstRow?.tcStyle : undefined;
  const lastRowStyle = hasLastRow ? styleDef?.lastRow?.tcStyle : undefined;
  const band1H = hasBandRow ? styleDef?.band1H?.tcStyle : undefined;
  const band2H = hasBandRow ? styleDef?.band2H?.tcStyle : undefined;

  let cellCount = 0;

  for (let ri = 0; ri < rowArr.length; ri++) {
    const row = rowArr[ri];
    const cells = row.tc;
    if (!cells) continue;
    const cellArr = Array.isArray(cells) ? cells : [cells];

    // Determine which style applies to this row
    let activeStyle = wholeTbl;
    if (ri === 0 && firstRowStyle) activeStyle = firstRowStyle;
    else if (ri === rowArr.length - 1 && lastRowStyle) activeStyle = lastRowStyle;
    else if (hasBandRow) activeStyle = (ri % 2 === (hasFirstRow ? 1 : 0)) ? band1H ?? wholeTbl : band2H ?? wholeTbl;

    for (const cell of cellArr) {
      if (!cell.tcPr) cell.tcPr = {};
      const tcPr = cell.tcPr;

      // Only add borders that aren't already explicitly defined
      let modified = false;
      for (const [xmlSide, styleSide] of [["lnL", "left"], ["lnR", "right"], ["lnT", "top"], ["lnB", "bottom"]] as const) {
        if (tcPr[xmlSide]) continue; // Already has explicit border
        const border = buildBorder(activeStyle, styleSide) ?? buildBorder(wholeTbl, styleSide);
        if (border) {
          tcPr[xmlSide] = border;
          modified = true;
        }
      }

      if (modified) cellCount++;
    }
  }

  if (cellCount > 0) {
    changes.push(`inlined borders on ${cellCount} table cells`);
  }

  return changes;
}

export const tableStyles: Transform = {
  name: "table-styles",

  detect(slideXml: any, slideNum: number): Issue[] {
    const issues: Issue[] = [];
    const tables = findTables(slideXml);

    for (const { table } of tables) {
      if (tableNeedsFix(table)) {
        const styleId = table.tblPr?.["@_tblStyle"] ?? table.tblPr?.tblStyle ?? "unknown";
        issues.push({
          type: "table-styles",
          slide: slideNum,
          element: `table(style=${styleId})`,
          severity: "high",
          description: "Table has style ID but cells lack explicit borders — will render without borders in QuickLook",
        });
      }
    }

    return issues;
  },

  apply(slideXml: any, slideNum: number, ctx: TransformContext): TransformResult {
    const tables = findTables(slideXml);
    const changes: string[] = [];

    for (const { table } of tables) {
      if (!tableNeedsFix(table)) continue;
      const styleId = table.tblPr?.["@_tblStyle"] ?? table.tblPr?.tblStyle;
      if (!styleId) continue;

      const styleDef = findStyleDef(ctx.tableStyleXml, styleId);
      if (!styleDef) {
        changes.push(`table style ${styleId} not found in tableStyles.xml — skipped`);
        continue;
      }

      changes.push(...applyTableStyle(table, styleDef));
    }

    return { changed: changes.length > 0, changes };
  },
};
