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

  if (node["p:sp"]) {
    const shapes = Array.isArray(node["p:sp"]) ? node["p:sp"] : [node["p:sp"]];
    for (const shape of shapes) {
      const spPr = shape["p:spPr"];
      if (!spPr?.["a:prstGeom"]) continue;
      const prst = spPr["a:prstGeom"]["@_prst"];
      if (!prst || SUPPORTED_GEOMETRIES.has(prst)) continue;

      const fallback = GEOMETRY_FALLBACKS[prst] ?? "rect";
      const name = shape["p:nvSpPr"]?.["p:cNvPr"]?.["@_name"] ?? "shape";
      spPr["a:prstGeom"]["@_prst"] = fallback;
      changes.push(`replaced unsupported geometry "${prst}" with ${fallback} on "${name}"`);
    }
  }

  if (node["p:grpSp"]) {
    const groups = Array.isArray(node["p:grpSp"]) ? node["p:grpSp"] : [node["p:grpSp"]];
    for (const group of groups) {
      fixGeometries(group, changes);
    }
  }

  if (node["p:spTree"]) {
    fixGeometries(node["p:spTree"], changes);
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
