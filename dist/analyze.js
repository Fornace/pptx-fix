/**
 * Analyze a PPTX for issues that will cause QuickLook rendering problems.
 * Read-only — does not modify the file.
 */
import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";
import { xmlParserOptions } from "./xml.js";
import { ALL_TRANSFORMS } from "./transforms/index.js";
export async function analyze(pptxBuffer) {
    const zip = await JSZip.loadAsync(pptxBuffer);
    const parser = new XMLParser(xmlParserOptions);
    const issues = [];
    // Find all slide XML files
    const slideFiles = Object.keys(zip.files)
        .filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))
        .sort((a, b) => {
        const na = parseInt(a.match(/\d+/)[0]);
        const nb = parseInt(b.match(/\d+/)[0]);
        return na - nb;
    });
    for (const slidePath of slideFiles) {
        const slideNum = parseInt(slidePath.match(/\d+/)[0]);
        const xml = await zip.file(slidePath).async("string");
        const parsed = parser.parse(xml);
        for (const transform of ALL_TRANSFORMS) {
            const found = transform.detect(parsed, slideNum);
            issues.push(...found);
        }
    }
    return issues;
}
