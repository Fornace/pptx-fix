/**
 * groups transform — Ungroup shape groups so children render individually
 * instead of being merged into a single opaque PDF block by QuickLook.
 *
 * Coordinate math: each child's position is relative to the group's
 * chOff/chExt (child coordinate space). To ungroup, we transform each
 * child's coordinates from group-space to slide-space.
 *
 * Detection is handled by quicklook-pptx-renderer's linter (group-as-pdf rule).
 */

import type { Transform, TransformContext, TransformResult } from "./index.js";

const CHILD_KEYS = ["p:sp", "p:pic", "p:cxnSp", "p:graphicFrame"] as const;

function getXfrm(child: any, key: string): any | undefined {
  if (key === "p:graphicFrame") return child["p:xfrm"] ?? child["a:xfrm"];
  return child["p:spPr"]?.["a:xfrm"];
}

function transformCoords(
  xfrm: any,
  groupOff: { x: number; y: number },
  groupExt: { cx: number; cy: number },
  chOff: { x: number; y: number },
  chExt: { cx: number; cy: number },
): void {
  const off = xfrm["a:off"];
  const rawExt = xfrm["a:ext"];
  const ext = Array.isArray(rawExt) ? rawExt[0] : rawExt;
  if (!off || !ext) return;

  const scaleX = chExt.cx > 0 ? groupExt.cx / chExt.cx : 1;
  const scaleY = chExt.cy > 0 ? groupExt.cy / chExt.cy : 1;

  off["@_x"] = String(Math.round(groupOff.x + (Number(off["@_x"] ?? 0) - chOff.x) * scaleX));
  off["@_y"] = String(Math.round(groupOff.y + (Number(off["@_y"] ?? 0) - chOff.y) * scaleY));
  ext["@_cx"] = String(Math.round(Number(ext["@_cx"] ?? 0) * scaleX));
  ext["@_cy"] = String(Math.round(Number(ext["@_cy"] ?? 0) * scaleY));
}

function ensureArray(parent: any, key: string): any[] {
  if (!parent[key]) parent[key] = [];
  if (!Array.isArray(parent[key])) parent[key] = [parent[key]];
  return parent[key];
}

function ungroupInto(parent: any, changes: string[]): void {
  if (!parent["p:grpSp"]) return;
  const groups = Array.isArray(parent["p:grpSp"]) ? parent["p:grpSp"] : [parent["p:grpSp"]];
  const kept: any[] = [];

  for (const group of groups) {
    // Recurse into nested groups first (ungroup from inside out)
    ungroupInto(group, changes);

    const xfrm = group["p:grpSpPr"]?.["a:xfrm"];
    if (!xfrm?.["a:off"] || !xfrm?.["a:ext"] || !xfrm?.["a:chOff"] || !xfrm?.["a:chExt"]) {
      kept.push(group);
      continue;
    }

    // Skip groups with rotation — ungrouping would lose the visual effect
    if (xfrm["@_rot"] && Number(xfrm["@_rot"]) !== 0) {
      kept.push(group);
      continue;
    }

    const gOff = xfrm["a:off"];
    const gExt = Array.isArray(xfrm["a:ext"]) ? xfrm["a:ext"][0] : xfrm["a:ext"];
    const gChOff = xfrm["a:chOff"];
    const gChExt = Array.isArray(xfrm["a:chExt"]) ? xfrm["a:chExt"][0] : xfrm["a:chExt"];

    const groupOff = { x: Number(gOff["@_x"] ?? 0), y: Number(gOff["@_y"] ?? 0) };
    const groupExt = { cx: Number(gExt["@_cx"] ?? 1), cy: Number(gExt["@_cy"] ?? 1) };
    const chOff = { x: Number(gChOff["@_x"] ?? 0), y: Number(gChOff["@_y"] ?? 0) };
    const chExt = { cx: Number(gChExt["@_cx"] ?? 1), cy: Number(gChExt["@_cy"] ?? 1) };
    const name = group["p:nvGrpSpPr"]?.["p:cNvPr"]?.["@_name"] ?? "group";

    let childCount = 0;
    for (const key of CHILD_KEYS) {
      if (!group[key]) continue;
      const children = Array.isArray(group[key]) ? group[key] : [group[key]];
      const target = ensureArray(parent, key);
      for (const child of children) {
        const childXfrm = getXfrm(child, key);
        if (childXfrm) transformCoords(childXfrm, groupOff, groupExt, chOff, chExt);
        target.push(child);
        childCount++;
      }
    }

    // Any nested groups that couldn't be ungrouped stay as top-level groups
    if (group["p:grpSp"]) {
      const nested = Array.isArray(group["p:grpSp"]) ? group["p:grpSp"] : [group["p:grpSp"]];
      for (const ng of nested) {
        const ngXfrm = ng["p:grpSpPr"]?.["a:xfrm"];
        if (ngXfrm) transformCoords(ngXfrm, groupOff, groupExt, chOff, chExt);
        kept.push(ng);
      }
    }

    if (childCount > 0) {
      changes.push(`ungrouped "${name}" (${childCount} children)`);
    }
  }

  if (kept.length > 0) {
    parent["p:grpSp"] = kept;
  } else {
    delete parent["p:grpSp"];
  }
}

export const groups: Transform = {
  name: "groups",

  apply(slideXml: any, _slideNum: number, _ctx: TransformContext): TransformResult {
    const changes: string[] = [];
    const spTree = slideXml?.["p:sld"]?.["p:cSld"]?.["p:spTree"];
    if (spTree) ungroupInto(spTree, changes);
    return { changed: changes.length > 0, changes };
  },
};
