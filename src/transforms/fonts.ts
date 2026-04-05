/**
 * fonts transform — Add explicit <a:latin>, <a:ea>, <a:cs> fallback
 * typefaces matching what OfficeImport's TCFontUtils would pick.
 */

import type { Issue } from "../analyze.js";
import type { Transform, TransformContext, TransformResult } from "./index.js";

export const fonts: Transform = {
  name: "fonts",

  detect(_slideXml: any, _slideNum: number): Issue[] {
    // TODO: detect text runs missing explicit font declarations
    return [];
  },

  apply(_slideXml: any, _slideNum: number, _ctx: TransformContext): TransformResult {
    // TODO: implement — add fallback typefaces
    return { changed: false, changes: [] };
  },
};
