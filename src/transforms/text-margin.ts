/**
 * text-margin transform — Widen text boxes by a safety margin to prevent
 * text wrapping caused by font metric differences across platforms.
 *
 * macOS substitutes Windows fonts with slightly different widths, causing
 * text that fit snugly to overflow to a new line. Adding ~15% width to
 * text-containing shapes absorbs the difference.
 */

import type { Transform, TransformContext, TransformResult } from "./index.js";

const MARGIN = 0.15;

function widenTextBoxes(node: any, changes: string[]): void {
  if (!node || typeof node !== "object") return;

  const shapes = node.sp;
  if (shapes) {
    const list = Array.isArray(shapes) ? shapes : [shapes];
    for (const sp of list) {
      if (!sp.txBody) continue;
      const ext = sp.spPr?.xfrm?.ext;
      if (!ext?.["@_cx"]) continue;
      const cx = Number(ext["@_cx"]);
      if (cx <= 0) continue;
      ext["@_cx"] = String(Math.round(cx * (1 + MARGIN)));
    }
  }

  // Recurse into child elements (but not attributes)
  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    const children = Array.isArray(node[key]) ? node[key] : [node[key]];
    for (const child of children) {
      if (child && typeof child === "object") widenTextBoxes(child, changes);
    }
  }
}

export const textMargin: Transform = {
  name: "text-margin",

  apply(slideXml: any, slideNum: number, _ctx: TransformContext): TransformResult {
    const changes: string[] = [];
    const spTree = slideXml?.sld?.cSld?.spTree ?? slideXml?.cSld?.spTree;
    if (spTree) {
      const before = JSON.stringify(spTree);
      widenTextBoxes(spTree, changes);
      if (JSON.stringify(spTree) !== before) {
        changes.push("widened text boxes by 15%");
      }
    }
    return { changed: changes.length > 0, changes };
  },
};
