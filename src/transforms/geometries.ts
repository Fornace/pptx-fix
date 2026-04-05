/**
 * geometries transform — Replace unsupported preset geometries with the
 * visually-closest supported alternative so shapes are visible in QuickLook.
 *
 * Detection is handled by quicklook-pptx-renderer's linter (unsupported-geometry rule).
 */

import type { Transform, TransformContext, TransformResult } from "./index.js";
import { SUPPORTED_GEOMETRIES, GEOMETRY_FALLBACKS } from "quicklook-pptx-renderer";

function fixGeometries(node: any, changes: string[]): void {
  if (!node || typeof node !== "object") return;

  if (node.sp) {
    const shapes = Array.isArray(node.sp) ? node.sp : [node.sp];
    for (const shape of shapes) {
      const spPr = shape.spPr;
      if (!spPr?.prstGeom) continue;
      const prst = spPr.prstGeom["@_prst"];
      if (!prst || SUPPORTED_GEOMETRIES.has(prst)) continue;

      const fallback = GEOMETRY_FALLBACKS[prst] ?? "rect";
      const name = shape.nvSpPr?.cNvPr?.["@_name"] ?? "shape";
      spPr.prstGeom["@_prst"] = fallback;
      changes.push(`replaced unsupported geometry "${prst}" with ${fallback} on "${name}"`);
    }
  }

  if (node.grpSp) {
    const groups = Array.isArray(node.grpSp) ? node.grpSp : [node.grpSp];
    for (const group of groups) {
      fixGeometries(group, changes);
    }
  }

  if (node.spTree) {
    fixGeometries(node.spTree, changes);
  }
}

export const geometries: Transform = {
  name: "geometries",

  apply(slideXml: any, _slideNum: number, _ctx: TransformContext): TransformResult {
    const changes: string[] = [];
    fixGeometries(slideXml, changes);
    return { changed: changes.length > 0, changes };
  },
};
