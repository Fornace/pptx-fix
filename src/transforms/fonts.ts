/**
 * fonts transform — Replace high-risk Windows fonts with metrically-closest
 * cross-platform alternatives to prevent text reflow on macOS.
 *
 * macOS substitutes Windows fonts (Calibri → Helvetica Neue +14.4%, etc.)
 * with different-width fonts, causing text overflow and line breaks to shift.
 * This replaces those fonts with cross-platform safe fonts that have minimal
 * width delta on macOS, chosen by the findClosestFont similarity algorithm.
 *
 * Detection is handled by quicklook-pptx-renderer's linter (font-substitution rule).
 */

import type { Transform, TransformContext, TransformResult } from "./index.js";
import { FONT_METRICS, FONT_SUBSTITUTIONS, findClosestFont, widthDelta } from "quicklook-pptx-renderer";

const DELTA_THRESHOLD = 10;

// Fonts available on macOS — everything in the metrics DB except Windows-only
const SAFE_CANDIDATES = Object.keys(FONT_METRICS).filter(f => !FONT_SUBSTITUTIONS[f]);

const replacementCache = new Map<string, string | null>();

function getReplacement(fontName: string): string | null {
  if (replacementCache.has(fontName)) return replacementCache.get(fontName)!;

  const macSub = FONT_SUBSTITUTIONS[fontName];
  const srcMetrics = FONT_METRICS[fontName];
  const subMetrics = macSub ? FONT_METRICS[macSub] : undefined;
  if (!macSub || !srcMetrics || !subMetrics || Math.abs(widthDelta(srcMetrics, subMetrics)) < DELTA_THRESHOLD) {
    replacementCache.set(fontName, null);
    return null;
  }

  const matches = findClosestFont(fontName, {
    candidates: SAFE_CANDIDATES,
    sameCategory: true,
    limit: 1,
  });

  const result = matches.length > 0 ? matches[0].font : null;
  replacementCache.set(fontName, result);
  return result;
}

const FONT_ELEMENTS = new Set(["latin", "ea", "cs", "sym", "buFont"]);

function replaceFonts(node: any, seen: Set<string>, changes: string[]): void {
  if (!node || typeof node !== "object") return;

  for (const key of FONT_ELEMENTS) {
    if (!node[key]) continue;
    const elements = Array.isArray(node[key]) ? node[key] : [node[key]];
    for (const el of elements) {
      const typeface = el["@_typeface"];
      if (!typeface || typeface.startsWith("+")) continue;
      const replacement = getReplacement(typeface);
      if (replacement) {
        el["@_typeface"] = replacement;
        if (!seen.has(typeface)) {
          seen.add(typeface);
          changes.push(`replaced font "${typeface}" ��� "${replacement}"`);
        }
      }
    }
  }

  for (const key of Object.keys(node)) {
    if (key.startsWith("@_") || FONT_ELEMENTS.has(key)) continue;
    const children = Array.isArray(node[key]) ? node[key] : [node[key]];
    for (const child of children) {
      replaceFonts(child, seen, changes);
    }
  }
}

export const fonts: Transform = {
  name: "fonts",

  apply(slideXml: any, _slideNum: number, _ctx: TransformContext): TransformResult {
    const changes: string[] = [];
    const seen = new Set<string>();
    replaceFonts(slideXml, seen, changes);
    return { changed: changes.length > 0, changes };
  },
};
