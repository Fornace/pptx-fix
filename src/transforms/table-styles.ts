/**
 * table-styles transform — Resolve table style conditional formatting into
 * explicit borders on every <a:tcPr>.
 *
 * OfficeImport doesn't resolve tableStyleId references — it only sees
 * explicit border properties. This inlines them from tableStyles.xml.
 */

import type { Transform, TransformContext, TransformResult } from "./index.js";

function findTables(node: any, path = ""): { table: any; path: string }[] {
  const results: { table: any; path: string }[] = [];
  if (!node || typeof node !== "object") return results;

  if (node["a:tbl"]) {
    const tbls = Array.isArray(node["a:tbl"]) ? node["a:tbl"] : [node["a:tbl"]];
    for (const t of tbls) results.push({ table: t, path: path + "/a:tbl" });
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

function tableNeedsFix(table: any): boolean {
  const tblPr = table["a:tblPr"];
  if (!tblPr) return false;

  const styleId = tblPr["@_tblStyle"];
  if (!styleId) return false;

  const rows = table["a:tr"];
  if (!rows) return false;
  const rowArr = Array.isArray(rows) ? rows : [rows];

  for (const row of rowArr) {
    const cells = row["a:tc"];
    if (!cells) continue;
    const cellArr = Array.isArray(cells) ? cells : [cells];
    for (const cell of cellArr) {
      const tcPr = cell["a:tcPr"];
      if (!tcPr) return true;
      if (!(tcPr["a:lnL"] || tcPr["a:lnR"] || tcPr["a:lnT"] || tcPr["a:lnB"])) return true;
    }
  }

  return false;
}

function findStyleDef(tableStyleXml: any, styleId: string): any | undefined {
  if (!tableStyleXml) return undefined;
  const styleLst = tableStyleXml["a:tblStyleLst"];
  if (!styleLst) return undefined;
  const styles = styleLst["a:tblStyle"];
  if (!styles) return undefined;
  const arr = Array.isArray(styles) ? styles : [styles];
  return arr.find((s: any) => s["@_styleId"] === styleId);
}

function buildBorder(tcStyle: any, side: string): any | undefined {
  if (!tcStyle) return undefined;
  const borders = tcStyle["a:tcBdr"];
  if (!borders) return undefined;
  const border = borders["a:" + side];
  if (!border) return undefined;
  return border["a:ln"] ?? border;
}

function applyTableStyle(table: any, styleDef: any): string[] {
  const changes: string[] = [];
  const rows = table["a:tr"];
  if (!rows) return changes;
  const rowArr = Array.isArray(rows) ? rows : [rows];

  const tblPr = table["a:tblPr"] ?? {};
  const hasFirstRow = tblPr["@_firstRow"] === "1" || tblPr["@_firstRow"] === "true";
  const hasLastRow = tblPr["@_lastRow"] === "1" || tblPr["@_lastRow"] === "true";
  const hasBandRow = tblPr["@_bandRow"] === "1" || tblPr["@_bandRow"] === "true";

  const wholeTbl = styleDef?.["a:wholeTbl"]?.["a:tcStyle"];
  const firstRowStyle = hasFirstRow ? styleDef?.["a:firstRow"]?.["a:tcStyle"] : undefined;
  const lastRowStyle = hasLastRow ? styleDef?.["a:lastRow"]?.["a:tcStyle"] : undefined;
  const band1H = hasBandRow ? styleDef?.["a:band1H"]?.["a:tcStyle"] : undefined;
  const band2H = hasBandRow ? styleDef?.["a:band2H"]?.["a:tcStyle"] : undefined;

  let cellCount = 0;

  for (let ri = 0; ri < rowArr.length; ri++) {
    const row = rowArr[ri];
    const cells = row["a:tc"];
    if (!cells) continue;
    const cellArr = Array.isArray(cells) ? cells : [cells];

    let activeStyle = wholeTbl;
    if (ri === 0 && firstRowStyle) activeStyle = firstRowStyle;
    else if (ri === rowArr.length - 1 && lastRowStyle) activeStyle = lastRowStyle;
    else if (hasBandRow) activeStyle = (ri % 2 === (hasFirstRow ? 1 : 0)) ? band1H ?? wholeTbl : band2H ?? wholeTbl;

    for (const cell of cellArr) {
      if (!cell["a:tcPr"]) cell["a:tcPr"] = {};
      const tcPr = cell["a:tcPr"];

      let modified = false;
      for (const [xmlSide, styleSide] of [["a:lnL", "left"], ["a:lnR", "right"], ["a:lnT", "top"], ["a:lnB", "bottom"]] as const) {
        if (tcPr[xmlSide]) continue;
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

  apply(slideXml: any, _slideNum: number, ctx: TransformContext): TransformResult {
    const tables = findTables(slideXml);
    const changes: string[] = [];

    for (const { table } of tables) {
      if (!tableNeedsFix(table)) continue;
      const styleId = table["a:tblPr"]?.["@_tblStyle"];
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
