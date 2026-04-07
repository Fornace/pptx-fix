/**
 * gradients transform — Collapse 3+ stop gradients to 2-stop so QuickLook
 * renders a gradient instead of averaging to a flat color.
 *
 * Detection is handled by quicklook-pptx-renderer's linter (gradient-flattened rule).
 */

import type { Transform, TransformContext, TransformResult } from "./index.js";

function collapseGradients(node: any, changes: string[], slideNum: number): void {
  if (!node || typeof node !== "object") return;

  if (node["a:gradFill"]) {
    const fills = Array.isArray(node["a:gradFill"]) ? node["a:gradFill"] : [node["a:gradFill"]];
    for (const fill of fills) {
      const gsLst = fill["a:gsLst"];
      if (!gsLst) continue;
      const stops = gsLst["a:gs"];
      if (!Array.isArray(stops) || stops.length < 3) continue;
      gsLst["a:gs"] = [stops[0], stops[stops.length - 1]];
      changes.push(`collapsed ${stops.length}-stop gradient to 2 stops`);
    }
  }

  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    const children = Array.isArray(node[key]) ? node[key] : [node[key]];
    for (const child of children) {
      collapseGradients(child, changes, slideNum);
    }
  }
}

export const gradients: Transform = {
  name: "gradients",

  apply(slideXml: any, slideNum: number, _ctx: TransformContext): TransformResult {
    const changes: string[] = [];
    collapseGradients(slideXml, changes, slideNum);
    return { changed: changes.length > 0, changes };
  },
};
