/**
 * gradients transform — Collapse 3+ stop gradients to 2-stop so QuickLook
 * renders a gradient instead of averaging to a flat color.
 *
 * Detection is handled by quicklook-pptx-renderer's linter (gradient-flattened rule).
 */

import type { Transform, TransformContext, TransformResult } from "./index.js";

export const gradients: Transform = {
  name: "gradients",

  apply(_slideXml: any, _slideNum: number, _ctx: TransformContext): TransformResult {
    // TODO: collapse 3+ stop gradients to 2-stop (endpoints only)
    return { changed: false, changes: [] };
  },
};
