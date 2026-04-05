/**
 * geometries transform — Replace unsupported preset geometries with
 * equivalent <a:custGeom> path data.
 *
 * Detection is handled by quicklook-pptx-renderer's linter (unsupported-geometry rule).
 */

import type { Transform, TransformContext, TransformResult } from "./index.js";

export const geometries: Transform = {
  name: "geometries",

  apply(_slideXml: any, _slideNum: number, _ctx: TransformContext): TransformResult {
    // TODO: replace prstGeom with custGeom containing equivalent path data
    return { changed: false, changes: [] };
  },
};
