/**
 * text-fit transform -- Measures actual rendered text dimensions using canvas
 * and reduces font sizes where text overflows its box or overlaps other text.
 *
 * Runs after the fonts transform to verify that post-replacement text still fits.
 * Uses measureTextWidth from quicklook-pptx-renderer for pixel-perfect measurement
 * with macOS system fonts.
 */

import type { Transform, TransformContext, TransformResult } from "./index.js";
import { FONT_SUBSTITUTIONS } from "quicklook-pptx-renderer";
import { createCanvas, type SKRSContext2D } from "@napi-rs/canvas";

const EMU_PER_PT = 12700;
const DEFAULT_L_INS = 91440;
const DEFAULT_R_INS = 91440;
const DEFAULT_T_INS = 45720;
const DEFAULT_B_INS = 45720;
const DEFAULT_FONT_SIZE = 1800; // 18pt in hundredths
const LINE_HEIGHT_FACTOR = 1.2;
const SAFETY_MARGIN = 0.95;

// ── Canvas text measurement ────────────────────────────────────────

let _ctx: SKRSContext2D | null = null;
function getCtx(): SKRSContext2D {
  if (!_ctx) {
    _ctx = createCanvas(1, 1).getContext("2d");
  }
  return _ctx;
}

function measureTextWidth(
  text: string,
  font: string,
  sizeHpt: number,
  opts?: { bold?: boolean; italic?: boolean },
): number {
  const ctx = getCtx();
  const sizePt = sizeHpt / 100;
  const bold = opts?.bold ? "bold " : "";
  const italic = opts?.italic ? "italic " : "";
  ctx.font = `${italic}${bold}${sizePt}px "${font}"`;
  return Math.round(ctx.measureText(text).width * EMU_PER_PT);
}

// ── Theme font resolution ──────────────────────────────────────────

function getThemeFonts(themeXml: any): { major?: string; minor?: string } {
  const fs = themeXml?.["a:theme"]?.["a:themeElements"]?.["a:fontScheme"];
  return {
    major: fs?.["a:majorFont"]?.["a:latin"]?.["@_typeface"],
    minor: fs?.["a:minorFont"]?.["a:latin"]?.["@_typeface"],
  };
}

function resolveFont(typeface: string | undefined, themeXml: any): string {
  if (!typeface) return "Helvetica";

  // Theme reference: resolve to actual font name first
  if (typeface.startsWith("+")) {
    const theme = getThemeFonts(themeXml);
    let resolved: string | undefined;
    if (typeface === "+mj-lt" || typeface === "+mj-ea" || typeface === "+mj-cs") {
      resolved = theme.major;
    } else if (typeface === "+mn-lt" || typeface === "+mn-ea" || typeface === "+mn-cs") {
      resolved = theme.minor;
    }
    if (!resolved) return "Helvetica";
    typeface = resolved;
  }

  // Check if macOS substitutes this font
  const macSub = FONT_SUBSTITUTIONS[typeface];
  return macSub ?? typeface;
}

// ── Helpers ────────────────────────────────────────────────────────

function asArray<T>(x: T | T[] | undefined): T[] {
  if (x == null) return [];
  return Array.isArray(x) ? x : [x];
}

// ── Text extraction ────────────────────────────────────────────────

interface TextRun {
  text: string;
  font: string;
  sizeHpt: number;
  bold: boolean;
  italic: boolean;
}

interface ParagraphInfo {
  runs: TextRun[];
  defaultSizeHpt?: number;
}

function isBool(v: any): boolean {
  return v === "1" || v === "true" || v === true || v === 1;
}

function extractParagraphs(txBody: any, themeXml: any): ParagraphInfo[] {
  const paras: ParagraphInfo[] = [];
  for (const p of asArray(txBody["a:p"])) {
    const pPr = p["a:pPr"];
    const defSz = pPr?.["a:defRPr"]?.["@_sz"];
    const defaultSizeHpt = defSz ? Number(defSz) : undefined;
    const runs: TextRun[] = [];

    for (const r of asArray(p["a:r"])) {
      const text = r["a:t"];
      if (text == null || text === "") continue;
      const rPr = r["a:rPr"];
      const sizeHpt = Number(rPr?.["@_sz"] ?? defaultSizeHpt ?? DEFAULT_FONT_SIZE);
      const font = resolveFont(rPr?.["a:latin"]?.["@_typeface"], themeXml);
      runs.push({
        text: String(text),
        font,
        sizeHpt,
        bold: isBool(rPr?.["@_b"]),
        italic: isBool(rPr?.["@_i"]),
      });
    }

    // Line breaks count as empty paragraphs (consume one line of height)
    if (runs.length === 0 && p["a:br"]) {
      runs.push({ text: "", font: "Helvetica", sizeHpt: Number(defSz ?? DEFAULT_FONT_SIZE), bold: false, italic: false });
    }

    paras.push({ runs, defaultSizeHpt });
  }
  return paras;
}

// ── Measurement ────────────────────────────────────────────────────

interface BoxMetrics {
  heightRatio: number; // totalTextHeight / availableHeight (>1 = overflow)
}

function measureBox(
  paras: ParagraphInfo[],
  availableWidthEmu: number,
  availableHeightEmu: number,
): BoxMetrics {
  if (availableWidthEmu <= 0 || availableHeightEmu <= 0) {
    return { heightRatio: 0 };
  }

  let totalHeightEmu = 0;

  for (const para of paras) {
    if (para.runs.length === 0) {
      // Empty paragraph still takes one line
      const sz = para.defaultSizeHpt ?? DEFAULT_FONT_SIZE;
      totalHeightEmu += (sz / 100) * EMU_PER_PT * LINE_HEIGHT_FACTOR;
      continue;
    }

    // Build a flat list of word-segments with their widths
    const maxSizeHpt = Math.max(...para.runs.map(r => r.sizeHpt));
    const lineHeightEmu = (maxSizeHpt / 100) * EMU_PER_PT * LINE_HEIGHT_FACTOR;

    // Simulate word wrap: measure run by run, splitting at spaces
    let currentLineWidth = 0;
    let lines = 1;

    for (const run of para.runs) {
      if (run.text === "") continue;
      const words = run.text.split(/(\s+)/);
      for (const word of words) {
        if (word === "") continue;
        const wordWidth = measureTextWidth(word, run.font, run.sizeHpt, {
          bold: run.bold,
          italic: run.italic,
        });
        if (currentLineWidth > 0 && currentLineWidth + wordWidth > availableWidthEmu) {
          lines++;
          currentLineWidth = /^\s+$/.test(word) ? 0 : wordWidth;
        } else {
          currentLineWidth += wordWidth;
        }
      }
    }

    totalHeightEmu += lines * lineHeightEmu;
  }

  return { heightRatio: totalHeightEmu / availableHeightEmu };
}

// ── Shape info for overlap detection ───────────────────────────────

interface ShapeRect {
  x: number;
  y: number;
  w: number;
  h: number;
  textHeight: number;
  shape: any; // reference to p:sp for mutation
  name: string;
}

function getShapeRect(sp: any, themeXml: any): ShapeRect | null {
  const xfrm = sp["p:spPr"]?.["a:xfrm"];
  const off = xfrm?.["a:off"];
  const ext = Array.isArray(xfrm?.["a:ext"]) ? xfrm["a:ext"][0] : xfrm?.["a:ext"];
  const txBody = sp["p:txBody"];
  if (!off || !ext || !txBody) return null;

  const x = Number(off["@_x"] ?? 0);
  const y = Number(off["@_y"] ?? 0);
  const cx = Number(ext["@_cx"] ?? 0);
  const cy = Number(ext["@_cy"] ?? 0);
  if (cx <= 0 || cy <= 0) return null;

  const bodyPr = txBody["a:bodyPr"];
  const lIns = Number(bodyPr?.["@_lIns"] ?? DEFAULT_L_INS);
  const rIns = Number(bodyPr?.["@_rIns"] ?? DEFAULT_R_INS);
  const tIns = Number(bodyPr?.["@_tIns"] ?? DEFAULT_T_INS);
  const bIns = Number(bodyPr?.["@_bIns"] ?? DEFAULT_B_INS);

  const paras = extractParagraphs(txBody, themeXml);
  const availW = cx - lIns - rIns;
  const availH = cy - tIns - bIns;
  const metrics = measureBox(paras, availW, availH);
  const textHeight = metrics.heightRatio * availH;

  const name = sp["p:nvSpPr"]?.["p:cNvPr"]?.["@_name"] ?? "shape";
  return { x, y, w: cx, h: cy, textHeight, shape: sp, name };
}

// ── Font size mutation ─────────────────────────────────────────────

function scaleTextBody(txBody: any, scale: number): void {
  for (const p of asArray(txBody["a:p"])) {
    // Scale default run props
    const defSz = p["a:pPr"]?.["a:defRPr"]?.["@_sz"];
    if (defSz) {
      p["a:pPr"]["a:defRPr"]["@_sz"] = String(Math.floor(Number(defSz) * scale));
    }
    // Scale each run
    for (const r of asArray(p["a:r"])) {
      const sz = r["a:rPr"]?.["@_sz"];
      if (sz) {
        r["a:rPr"]["@_sz"] = String(Math.floor(Number(sz) * scale));
      }
    }
    // Scale line break run props too
    for (const br of asArray(p["a:br"])) {
      const sz = br["a:rPr"]?.["@_sz"];
      if (sz) {
        br["a:rPr"]["@_sz"] = String(Math.floor(Number(sz) * scale));
      }
    }
  }
}

// ── Main transform ─────────────────────────────────────────────────

export const textFit: Transform = {
  name: "text-fit",

  apply(slideXml: any, _slideNum: number, ctx: TransformContext): TransformResult {
    const changes: string[] = [];
    const spTree = slideXml?.["p:sld"]?.["p:cSld"]?.["p:spTree"];
    if (!spTree) return { changed: false, changes };

    const shapes = asArray(spTree["p:sp"]);
    const rects: ShapeRect[] = [];

    // Pass 1: fix individual overflow
    for (const sp of shapes) {
      const txBody = sp["p:txBody"];
      if (!txBody) continue;

      const xfrm = sp["p:spPr"]?.["a:xfrm"];
      const ext = Array.isArray(xfrm?.["a:ext"]) ? xfrm["a:ext"][0] : xfrm?.["a:ext"];
      if (!ext) continue;

      const cx = Number(ext["@_cx"] ?? 0);
      const cy = Number(ext["@_cy"] ?? 0);
      if (cx <= 0 || cy <= 0) continue;

      const bodyPr = txBody["a:bodyPr"];
      // Skip no-wrap text (titles, one-liners that intentionally extend)
      if (bodyPr?.["@_wrap"] === "none") continue;

      const lIns = Number(bodyPr?.["@_lIns"] ?? DEFAULT_L_INS);
      const rIns = Number(bodyPr?.["@_rIns"] ?? DEFAULT_R_INS);
      const tIns = Number(bodyPr?.["@_tIns"] ?? DEFAULT_T_INS);
      const bIns = Number(bodyPr?.["@_bIns"] ?? DEFAULT_B_INS);

      const availW = cx - lIns - rIns;
      const availH = cy - tIns - bIns;
      if (availW <= 0 || availH <= 0) continue;

      const paras = extractParagraphs(txBody, ctx.themeXml);
      if (paras.every(p => p.runs.length === 0)) continue;

      const metrics = measureBox(paras, availW, availH);
      if (metrics.heightRatio > 1) {
        const scale = (1 / metrics.heightRatio) * SAFETY_MARGIN;
        const name = sp["p:nvSpPr"]?.["p:cNvPr"]?.["@_name"] ?? "shape";
        scaleTextBody(txBody, scale);
        const pct = Math.round((1 - scale) * 100);
        changes.push(`shrunk text in "${name}" by ${pct}% to fit box`);
      }
    }

    // Pass 2: fix overlap between shapes
    // Re-measure after pass 1 may have shrunk some boxes
    for (const sp of shapes) {
      const rect = getShapeRect(sp, ctx.themeXml);
      if (rect) rects.push(rect);
    }

    const OVERLAP_THRESHOLD = 5 * EMU_PER_PT; // ~5pt minimum overlap to act on

    for (let i = 0; i < rects.length; i++) {
      for (let j = i + 1; j < rects.length; j++) {
        const a = rects[i];
        const b = rects[j];

        // Compute actual text rects (position + insets + measured text height)
        const aBody = a.shape["p:txBody"]?.["a:bodyPr"];
        const bBody = b.shape["p:txBody"]?.["a:bodyPr"];
        const aTIns = Number(aBody?.["@_tIns"] ?? DEFAULT_T_INS);
        const bTIns = Number(bBody?.["@_tIns"] ?? DEFAULT_T_INS);
        const aTextTop = a.y + aTIns;
        const aTextBottom = aTextTop + a.textHeight;
        const bTextTop = b.y + bTIns;
        const bTextBottom = bTextTop + b.textHeight;

        // Check bidirectional overlap
        const overlapV = Math.min(aTextBottom, bTextBottom) - Math.max(aTextTop, bTextTop);
        if (overlapV <= OVERLAP_THRESHOLD) continue;

        const overlapH = Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x);
        if (overlapH <= OVERLAP_THRESHOLD) continue;

        // Shrink the shape whose text is taller
        const target = a.textHeight >= b.textHeight ? a : b;
        const txBody = target.shape["p:txBody"];
        if (!txBody) continue;

        const desiredHeight = target.textHeight - overlapV;
        if (desiredHeight <= 0) continue;
        const scale = (desiredHeight / target.textHeight) * SAFETY_MARGIN;
        if (scale >= 1) continue;

        scaleTextBody(txBody, scale);
        target.textHeight *= scale; // update for subsequent pair checks
        const pct = Math.round((1 - scale) * 100);
        changes.push(`shrunk text in "${target.name}" by ${pct}% to reduce overlap`);
      }
    }

    return { changed: changes.length > 0, changes };
  },
};
