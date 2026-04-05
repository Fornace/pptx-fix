/**
 * effects transform — Strip or adjust shapes with effectLst that render
 * as opaque PDF blocks in QuickLook.
 *
 * Detection is handled by quicklook-pptx-renderer's linter (opaque-pdf-block rule).
 */

import type { Transform, TransformContext, TransformResult } from "./index.js";

export const effects: Transform = {
  name: "effects",

  apply(_slideXml: any, _slideNum: number, _ctx: TransformContext): TransformResult {
    // TODO: configurable — strip effects, reorder z-index, or no-op
    return { changed: false, changes: [] };
  },
};
