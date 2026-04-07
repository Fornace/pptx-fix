/**
 * Strip embedded fonts from a PPTX and replace with cross-platform alternatives.
 *
 * QuickLook (OfficeImport) ignores embedded fonts entirely — they bloat the
 * file and the text renders with system substitutes anyway. This:
 *  1. Finds the best cross-platform replacement for each embedded font
 *  2. Replaces all font references across slides and themes
 *  3. Removes the font data files, embeddedFontLst, rels, and content types
 */

import type JSZip from "jszip";
import { XMLParser, XMLBuilder } from "fast-xml-parser";
import { FONT_METRICS, findClosestFont, APPLE_SYSTEM_FONT_LIST } from "quicklook-pptx-renderer";

const FONT_ELEMENTS = new Set(["a:latin", "a:ea", "a:cs", "a:sym", "a:buFont"]);

const relsParserOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  removeNSPrefix: true,
  parseTagValue: false,
  parseAttributeValue: false,
  isArray: (name: string) => name === "Relationship",
};

const relsBuilderOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  format: true,
  suppressEmptyNode: false,
};

const presParserOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  removeNSPrefix: false,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
  isArray: (name: string) => name === "p:embeddedFont",
};

const presBuilderOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  format: false,
  suppressEmptyNode: false,
  suppressBooleanAttributes: false,
};

/** Slide/theme parser — must preserve namespace prefixes for round-trip. */
const slideParserOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  removeNSPrefix: false,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
  isArray: (name: string) => [
    "p:sp", "p:pic", "p:cxnSp", "p:grpSp", "p:graphicFrame",
    "a:p", "a:r", "a:br", "a:fld",
    "a:gs", "a:ln", "a:solidFill", "a:gradFill",
    "a:tr", "a:tc", "a:tblStyleLst", "a:ext",
  ].includes(name),
};

/** Find the best cross-platform replacement for a font name. */
function findReplacement(fontName: string): string | null {
  if (!FONT_METRICS[fontName]) return null;
  const matches = findClosestFont(fontName, {
    candidates: APPLE_SYSTEM_FONT_LIST,
    sameCategory: true,
    limit: 1,
  });
  return matches.length > 0 ? matches[0].font : null;
}

/** Walk an XML tree and replace typeface attributes matching the replacement map. */
function replaceTypefaces(node: any, replacements: Map<string, string>): void {
  if (!node || typeof node !== "object") return;

  for (const key of FONT_ELEMENTS) {
    if (!node[key]) continue;
    const elements = Array.isArray(node[key]) ? node[key] : [node[key]];
    for (const el of elements) {
      const typeface = el["@_typeface"];
      if (typeface && replacements.has(typeface)) {
        el["@_typeface"] = replacements.get(typeface)!;
      }
    }
  }

  for (const key of Object.keys(node)) {
    if (key.startsWith("@_") || FONT_ELEMENTS.has(key)) continue;
    const children = Array.isArray(node[key]) ? node[key] : [node[key]];
    for (const child of children) {
      replaceTypefaces(child, replacements);
    }
  }
}

export async function stripEmbeddedFonts(
  zip: JSZip,
  reportLines: string[],
): Promise<void> {
  const presFile = zip.file("ppt/presentation.xml");
  if (!presFile) return;

  const parser = new XMLParser(presParserOptions);
  const presXml = parser.parse(await presFile.async("string"));
  const presNode = presXml?.["p:presentation"] ?? presXml;
  if (!presNode) return;

  const embFontLst = presNode["p:embeddedFontLst"]?.["p:embeddedFont"];
  if (!embFontLst) return;

  const entries = Array.isArray(embFontLst) ? embFontLst : [embFontLst];
  if (entries.length === 0) return;

  // Build replacement map and collect rIds to remove
  const replacements = new Map<string, string>();
  const fontNames: string[] = [];
  const rIdsToRemove = new Set<string>();

  for (const entry of entries) {
    const typeface = entry["p:font"]?.["@_typeface"];
    if (typeface) {
      fontNames.push(typeface);
      const replacement = findReplacement(typeface);
      if (replacement) replacements.set(typeface, replacement);
    }

    for (const variant of ["p:regular", "p:bold", "p:italic", "p:boldItalic"]) {
      const rId = entry[variant]?.["@_r:id"];
      if (rId) rIdsToRemove.add(rId);
    }
  }

  if (rIdsToRemove.size === 0) return;

  // Replace font references across slides and themes
  if (replacements.size > 0) {
    const slideParser = new XMLParser(slideParserOptions);
    const xmlFiles = Object.keys(zip.files).filter(f =>
      /^ppt\/(slides\/slide|theme\/theme)\d+\.xml$/.test(f)
    );
    for (const path of xmlFiles) {
      const xml = await zip.file(path)!.async("string");
      const parsed = slideParser.parse(xml);
      replaceTypefaces(parsed, replacements);
      const builder = new XMLBuilder(presBuilderOptions);
      zip.file(path, builder.build(parsed));
    }
  }

  // Remove embeddedFontLst from presentation XML
  delete presNode["p:embeddedFontLst"];
  const builder = new XMLBuilder(presBuilderOptions);
  zip.file("ppt/presentation.xml", builder.build(presXml));

  // Remove font files via presentation rels
  const relsPath = "ppt/_rels/presentation.xml.rels";
  const relsFile = zip.file(relsPath);
  if (relsFile) {
    const relsParser = new XMLParser(relsParserOptions);
    const relsXml = relsParser.parse(await relsFile.async("string"));
    const rels = relsXml?.Relationships?.Relationship ?? [];

    const kept: any[] = [];
    for (const rel of rels) {
      if (rIdsToRemove.has(rel["@_Id"])) {
        const target: string = rel["@_Target"];
        const resolved = target.startsWith("/")
          ? target.slice(1)
          : "ppt/" + (target.startsWith("../") ? target.slice(3) : target);
        zip.remove(resolved);
      } else {
        kept.push(rel);
      }
    }

    relsXml.Relationships.Relationship = kept;
    const relsBuilder = new XMLBuilder(relsBuilderOptions);
    zip.file(relsPath, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + relsBuilder.build(relsXml));
  }

  // Clean up Content_Types.xml — remove fntdata entries if no font files remain
  const hasFontFiles = Object.keys(zip.files).some(f => f.endsWith(".fntdata"));
  if (!hasFontFiles) {
    const ctFile = zip.file("[Content_Types].xml");
    if (ctFile) {
      let ct = await ctFile.async("string");
      ct = ct.replace(/<Default[^>]*Extension="fntdata"[^>]*\/>\s*/g, "");
      ct = ct.replace(/<Override[^>]*PartName="[^"]*\/fonts\/[^"]*"[^>]*\/>\s*/g, "");
      zip.file("[Content_Types].xml", ct);
    }
  }

  // Report
  const details = fontNames.map(f => {
    const r = replacements.get(f);
    return r ? `"${f}" → "${r}"` : `"${f}" (stripped)`;
  }).join(", ");
  reportLines.push(`embedded fonts: ${details}`);
}
