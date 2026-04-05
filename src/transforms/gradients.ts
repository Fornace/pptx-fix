/**
 * gradients transform — Collapse 3+ stop gradients to 2-stop so QuickLook
 * renders a gradient instead of averaging to a flat color.
 *
 * Lossy: intermediate gradient stops are removed.
 */

import type { Issue } from "../analyze.js";
import type { Transform, TransformContext, TransformResult } from "./index.js";

function findGradients(node: any, slideNum: number): Issue[] {
  const issues: Issue[] = [];
  if (!node || typeof node !== "object") return issues;

  if (node.gradFill) {
    const fills = Array.isArray(node.gradFill) ? node.gradFill : [node.gradFill];
    for (const fill of fills) {
      const gsLst = fill.gsLst;
      if (!gsLst?.gs) continue;
      const stops = Array.isArray(gsLst.gs) ? gsLst.gs : [gsLst.gs];
      if (stops.length > 2) {
        issues.push({
          type: "gradients",
          slide: slideNum,
          severity: "medium",
          description: `Gradient with ${stops.length} stops — QuickLook will average to flat color`,
        });
      }
    }
  }

  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    const children = Array.isArray(node[key]) ? node[key] : [node[key]];
    for (const child of children) {
      issues.push(...findGradients(child, slideNum));
    }
  }
  return issues;
}

export const gradients: Transform = {
  name: "gradients",

  detect(slideXml: any, slideNum: number): Issue[] {
    return findGradients(slideXml, slideNum);
  },

  apply(_slideXml: any, _slideNum: number, _ctx: TransformContext): TransformResult {
    // TODO: implement — collapse 3+ stop gradients to 2-stop (endpoints only)
    return { changed: false, changes: [] };
  },
};
