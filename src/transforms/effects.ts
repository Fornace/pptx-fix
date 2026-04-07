/**
 * effects transform — Strip effectLst/effectDag from shape properties so
 * shapes render as CSS instead of opaque PDF blocks in QuickLook.
 *
 * Detection is handled by quicklook-pptx-renderer's linter (opaque-pdf-block rule).
 */

import type { Transform, TransformContext, TransformResult } from "./index.js";

function stripEffects(node: any, changes: string[]): void {
  if (!node || typeof node !== "object") return;

  // Shape elements: p:sp, p:cxnSp, p:pic
  const shapeKeys = ["p:sp", "p:cxnSp", "p:pic"];
  for (const key of shapeKeys) {
    if (!node[key]) continue;
    const shapes = Array.isArray(node[key]) ? node[key] : [node[key]];
    for (const shape of shapes) {
      const spPr = shape["p:spPr"];
      if (!spPr) continue;

      const name = shape["p:nvSpPr"]?.["p:cNvPr"]?.["@_name"]
        ?? shape["p:nvCxnSpPr"]?.["p:cNvPr"]?.["@_name"]
        ?? shape["p:nvPicPr"]?.["p:cNvPr"]?.["@_name"]
        ?? "shape";

      if (spPr["a:effectLst"]) {
        delete spPr["a:effectLst"];
        changes.push(`stripped effects from "${name}"`);
      }
      if (spPr["a:effectDag"]) {
        delete spPr["a:effectDag"];
        changes.push(`stripped effect DAG from "${name}"`);
      }
    }
  }

  // Groups: recurse into grpSp children
  if (node["p:grpSp"]) {
    const groups = Array.isArray(node["p:grpSp"]) ? node["p:grpSp"] : [node["p:grpSp"]];
    for (const group of groups) {
      stripEffects(group, changes);
    }
  }

  // Recurse into spTree (the slide's shape tree)
  if (node["p:spTree"]) {
    stripEffects(node["p:spTree"], changes);
  }
}

export const effects: Transform = {
  name: "effects",

  apply(slideXml: any, _slideNum: number, _ctx: TransformContext): TransformResult {
    const changes: string[] = [];
    stripEffects(slideXml, changes);
    return { changed: changes.length > 0, changes };
  },
};
