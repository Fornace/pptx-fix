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

  const enabledNames = new Set<TransformName>(options?.transforms ?? ALL_TRANSFORMS.map(t => t.name));
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

  // Process each slide
  for (const slidePath of slideFiles) {
    const slideNum = parseInt(slidePath.match(/\d+/)![0]);
    const xml = await zip.file(slidePath)!.async("string");
    const parsed = parser.parse(xml);

    let changed = false;
    for (const transform of enabled) {
      const result = transform.apply(parsed, slideNum, { tableStyleXml });
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

  const outBuffer = Buffer.from(await zip.generateAsync({ type: "nodebuffer" }));

  const result: FixResult = { buffer: outBuffer };
  if (options?.report) {
    result.report = reportLines.length > 0
      ? reportLines.join("\n")
      : "No issues found — file unchanged.";
  }

  return result;
}
