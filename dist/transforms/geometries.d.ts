/**
 * geometries transform — Replace unsupported preset geometries with
 * equivalent <a:custGeom> path data.
 *
 * OfficeImport's CMCanonicalShapeBuilder silently drops ~30 presets
 * (heart, cloud, lightningBolt, etc.). This converts them to custom
 * geometry paths that OfficeImport can render.
 */
import type { Transform } from "./index.js";
export declare const geometries: Transform;
