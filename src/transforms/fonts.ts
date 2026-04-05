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
import { FONT_METRICS, findClosestFont } from "quicklook-pptx-renderer";

// Fonts available on both Windows and macOS with ≤1% substitution delta.
const SAFE_FONTS = [
  "Arial", "Verdana", "Georgia", "Trebuchet MS",
  "Times New Roman", "Courier New", "Impact",
  "Palatino Linotype", "Century Gothic",
];

// macOS substitutions with ≥10% absolute width delta — these cause reflow.
const HIGH_RISK = new Set([
  "Calibri", "Calibri Light",
  "Arial Black", "Arial Narrow",
  "Tahoma",
  "Segoe UI", "Segoe UI Light", "Segoe UI Semibold",
  "Franklin Gothic Medium",
  "Corbel", "Candara",
]);

const replacementCache = new Map<string, string | null>();

function getReplacement(fontName: string): string | null {
  if (replacementCache.has(fontName)) return replacementCache.get(fontName)!;

  if (!HIGH_RISK.has(fontName) || !FONT_METRICS[fontName]) {
    replacementCache.set(fontName, null);
    return null;
  }

  const candidates = SAFE_FONTS.filter(f => FONT_METRICS[f]);
  const matches = findClosestFont(fontName, {
    candidates,
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
