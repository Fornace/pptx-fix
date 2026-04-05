/**
 * effects transform — Detect shapes with effectLst that will become opaque
 * PDF blocks in QuickLook (covering content behind them).
 *
 * Options: strip effects, reorder z-index, or just warn.
 */

import type { Issue } from "../analyze.js";
import type { Transform, TransformContext, TransformResult } from "./index.js";

function findEffectShapes(node: any, slideNum: number): Issue[] {
  const issues: Issue[] = [];
  if (!node || typeof node !== "object") return issues;

  if (node.effectLst && Object.keys(node.effectLst).some(k => !k.startsWith("@_"))) {
    issues.push({
      type: "effects",
      slide: slideNum,
      severity: "low",
      description: "Shape with effects — will render as opaque PDF block in QuickLook (may cover content behind it)",
    });
  }

  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    const children = Array.isArray(node[key]) ? node[key] : [node[key]];
    for (const child of children) {
      issues.push(...findEffectShapes(child, slideNum));
    }
  }
  return issues;
}

export const effects: Transform = {
  name: "effects",

  detect(slideXml: any, slideNum: number): Issue[] {
    return findEffectShapes(slideXml, slideNum);
  },

  apply(_slideXml: any, _slideNum: number, _ctx: TransformContext): TransformResult {
    // TODO: implement — configurable: strip effects, reorder, or no-op
    return { changed: false, changes: [] };
  },
};
