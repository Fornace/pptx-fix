/**
 * properties transform — Resolve full inheritance chain
 * (theme → master → layout → slide) and write explicit fill, font, color
 * on each element.
 */

import type { Issue } from "../analyze.js";
import type { Transform, TransformContext, TransformResult } from "./index.js";

export const properties: Transform = {
  name: "properties",

  detect(_slideXml: any, _slideNum: number): Issue[] {
    // TODO: detect elements relying on inherited properties
    return [];
  },

  apply(_slideXml: any, _slideNum: number, _ctx: TransformContext): TransformResult {
    // TODO: implement — resolve inheritance and inline explicit properties
    return { changed: false, changes: [] };
  },
};
