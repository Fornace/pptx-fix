/**
 * XML round-trip writer: parse → mutate → serialize → repack ZIP.
 *
 * Transforms operate on fast-xml-parser parsed objects (preserving unknown elements),
 * then XMLBuilder serializes them back to XML. JSZip repacks the modified ZIP.
 */

import JSZip from "jszip";
import { XMLParser, XMLBuilder } from "fast-xml-parser";
import { xmlParserOptions, xmlBuilderOptions } from "./xml.js";
import { ALL_TRANSFORMS, type TransformName, type Transform } from "./transforms/index.js";
import { addChartFallbacks } from "./chart-fallbacks.js";
import { stripEmbeddedFonts } from "./embedded-fonts.js";

export interface FixOptions {
  /** Apply only these transforms (default: all) */
  transforms?: TransformName[];
  /** Return a human-readable report of changes */
  report?: boolean;
}

export interface FixResult {
  buffer: Buffer;
  report?: string;
}

export async function fix(pptxBuffer: Buffer, options?: FixOptions): Promise<FixResult> {
  const zip = await JSZip.loadAsync(pptxBuffer);
  const parser = new XMLParser(xmlParserOptions);
  const builder = new XMLBuilder(xmlBuilderOptions);
  const reportLines: string[] = [];

  const allNames: TransformName[] = [...ALL_TRANSFORMS.map(t => t.name), "embedded-fonts"];
  const enabledNames = new Set<TransformName>(options?.transforms ?? allNames);
  const enabled = ALL_TRANSFORMS.filter(t => enabledNames.has(t.name));

  // Find all slide XML files
  const slideFiles = Object.keys(zip.files)
    .filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => {
      const na = parseInt(a.match(/\d+/)![0]);
      const nb = parseInt(b.match(/\d+/)![0]);
      return na - nb;
    });

  // Load table style XML if needed (for table-styles transform)
  let tableStyleXml: any = undefined;
  if (enabledNames.has("table-styles")) {
    const tsPath = "ppt/tableStyles.xml";
    const tsFile = zip.file(tsPath);
    if (tsFile) {
      tableStyleXml = parser.parse(await tsFile.async("string"));
    }
  }

  // Load theme XML if needed (for text-fit transform)
  let themeXml: any = undefined;
  if (enabledNames.has("text-fit")) {
    const themeFile = zip.file("ppt/theme/theme1.xml");
    if (themeFile) {
      themeXml = parser.parse(await themeFile.async("string"));
    }
  }

  // Process each slide
  for (const slidePath of slideFiles) {
    const slideNum = parseInt(slidePath.match(/\d+/)![0]);
    const xml = await zip.file(slidePath)!.async("string");
    const parsed = parser.parse(xml);

    let changed = false;
    for (const transform of enabled) {
      const result = transform.apply(parsed, slideNum, { tableStyleXml, themeXml });
      if (result.changed) {
        changed = true;
        for (const line of result.changes) {
          reportLines.push(`Slide ${slideNum}: ${line}`);
        }
      }
    }

    if (changed) {
      zip.file(slidePath, builder.build(parsed));
    }
  }

  // Process theme XML for font replacement
  if (enabledNames.has("fonts")) {
    const fontsTransform = enabled.find(t => t.name === "fonts")!;
    const themeFiles = Object.keys(zip.files).filter(f => /^ppt\/theme\/theme\d+\.xml$/.test(f));
    for (const themePath of themeFiles) {
      const xml = await zip.file(themePath)!.async("string");
      const parsed = parser.parse(xml);
      const result = fontsTransform.apply(parsed, 0, { tableStyleXml });
      if (result.changed) {
        zip.file(themePath, builder.build(parsed));
        for (const line of result.changes) {
          reportLines.push(`Theme: ${line}`);
        }
      }
    }
  }

  // Strip embedded fonts (QuickLook ignores them, they just bloat the file)
  if (enabledNames.has("embedded-fonts")) {
    await stripEmbeddedFonts(zip, reportLines);
  }

  // Generate chart fallback images (requires Playwright — skips if not installed)
  await addChartFallbacks(zip, pptxBuffer, reportLines);

  const outBuffer = Buffer.from(await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } }));

  const result: FixResult = { buffer: outBuffer };
  if (options?.report) {
    result.report = reportLines.length > 0
      ? reportLines.join("\n")
      : "No issues found — file unchanged.";
  }

  return result;
}
