/**
 * geometries transform — Replace unsupported preset geometries with
 * equivalent <a:custGeom> path data.
 *
 * OfficeImport's CMCanonicalShapeBuilder silently drops ~30 presets
 * (heart, cloud, lightningBolt, etc.). This converts them to custom
 * geometry paths that OfficeImport can render.
 */

import type { Issue } from "../analyze.js";
import type { Transform, TransformContext, TransformResult } from "./index.js";

/** Presets that CMCanonicalShapeBuilder does NOT support. */
const UNSUPPORTED_PRESETS = new Set([
  "heart", "cloud", "lightningBolt", "sun", "moon",
  "irregularSeal1", "plaque", "frame", "halfFrame",
  "corner", "diagStripe", "chord", "arc",
  "bracketPair", "bracePair",
]);

function findUnsupportedGeometries(node: any, slideNum: number): Issue[] {
  const issues: Issue[] = [];
  if (!node || typeof node !== "object") return issues;

  if (node.prstGeom) {
    const prst = node.prstGeom["@_prst"];
    if (prst && UNSUPPORTED_PRESETS.has(prst)) {
      issues.push({
        type: "geometries",
        slide: slideNum,
        element: prst,
        severity: "high",
        description: `Preset '${prst}' is not supported by QuickLook — shape will be invisible`,
      });
    }
  }

  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    const children = Array.isArray(node[key]) ? node[key] : [node[key]];
    for (const child of children) {
      issues.push(...findUnsupportedGeometries(child, slideNum));
    }
  }
  return issues;
}

export const geometries: Transform = {
  name: "geometries",

  detect(slideXml: any, slideNum: number): Issue[] {
    return findUnsupportedGeometries(slideXml, slideNum);
  },

  apply(_slideXml: any, _slideNum: number, _ctx: TransformContext): TransformResult {
    // TODO: implement — replace prstGeom with custGeom containing equivalent path data
    return { changed: false, changes: [] };
  },
};
