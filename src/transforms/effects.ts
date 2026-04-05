/**
 * effects transform — Strip effectLst/effectDag from shape properties so
 * shapes render as CSS instead of opaque PDF blocks in QuickLook.
 *
 * Detection is handled by quicklook-pptx-renderer's linter (opaque-pdf-block rule).
 */

import type { Transform, TransformContext, TransformResult } from "./index.js";

function stripEffects(node: any, changes: string[]): void {
  if (!node || typeof node !== "object") return;

  // Shape elements: sp, cxnSp, pic
  const shapeKeys = ["sp", "cxnSp", "pic"];
  for (const key of shapeKeys) {
    if (!node[key]) continue;
    const shapes = Array.isArray(node[key]) ? node[key] : [node[key]];
    for (const shape of shapes) {
      const spPr = shape.spPr;
      if (!spPr) continue;

      const name = shape.nvSpPr?.cNvPr?.["@_name"]
        ?? shape.nvCxnSpPr?.cNvPr?.["@_name"]
        ?? shape.nvPicPr?.cNvPr?.["@_name"]
        ?? "shape";

      if (spPr.effectLst) {
        delete spPr.effectLst;
        changes.push(`stripped effects from "${name}"`);
      }
      if (spPr.effectDag) {
        delete spPr.effectDag;
        changes.push(`stripped effect DAG from "${name}"`);
      }
    }
  }

  // Groups: recurse into grpSp children
  if (node.grpSp) {
    const groups = Array.isArray(node.grpSp) ? node.grpSp : [node.grpSp];
    for (const group of groups) {
      stripEffects(group, changes);
    }
  }

  // Recurse into spTree (the slide's shape tree)
  if (node.spTree) {
    stripEffects(node.spTree, changes);
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
