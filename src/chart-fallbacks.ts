/**
 * Chart fallback generator — render charts to PNG and embed as fallback
 * images so QuickLook displays charts instead of blank rectangles.
 *
 * QuickLook (OfficeImport) doesn't parse chart XML — it only uses
 * pre-rendered fallback images stored in the PPTX. Tools like python-pptx
 * and PptxGenJS don't generate these fallback images.
 *
 * Uses Playwright to screenshot chart HTML rendered by quicklook-pptx-renderer.
 * Playwright is a dynamic import — if not installed, this step is skipped.
 */

import type JSZip from "jszip";
import { XMLParser, XMLBuilder } from "fast-xml-parser";

const CHART_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

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

/** Slide parser — preserves namespace prefixes for correct element access. */
const slideParserOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  removeNSPrefix: false,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
  isArray: (name: string) => [
    "p:sp", "p:pic", "p:cxnSp", "p:grpSp", "p:graphicFrame",
    "a:p", "a:r", "a:ext",
  ].includes(name),
};

interface ChartInfo {
  slideNum: number;
  chartPath: string;     // e.g. "ppt/charts/chart1.xml"
  chartRelsPath: string; // e.g. "ppt/charts/_rels/chart1.xml.rels"
  bounds: { x: number; y: number; cx: number; cy: number };
}

/** Find charts without fallback images in the PPTX. */
async function findChartsWithoutFallback(zip: JSZip): Promise<ChartInfo[]> {
  const relsParser = new XMLParser(relsParserOptions);
  const slideParser = new XMLParser(slideParserOptions);
  const results: ChartInfo[] = [];

  const slideFiles = Object.keys(zip.files)
    .filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort();

  for (const slidePath of slideFiles) {
    const slideNum = parseInt(slidePath.match(/\d+/)![0]);

    // Read slide rels
    const slideRelsPath = slidePath.replace("slides/", "slides/_rels/").replace(".xml", ".xml.rels");
    const slideRelsFile = zip.file(slideRelsPath);
    if (!slideRelsFile) continue;
    const slideRels = relsParser.parse(await slideRelsFile.async("string"));
    const rels = slideRels.Relationships?.Relationship ?? [];

    // Find chart relationships
    const chartRels = rels.filter((r: any) =>
      r["@_Type"]?.includes("/chart") || r["@_Type"] === CHART_REL_TYPE
    );

    for (const chartRel of chartRels) {
      const target = chartRel["@_Target"];
      if (!target) continue;

      // Resolve chart path relative to slide
      const chartPath = target.startsWith("../")
        ? "ppt/" + target.slice(3)
        : "ppt/slides/" + target;

      // Check if chart already has a fallback image
      const chartRelsPath = chartPath.replace(/([^/]+)$/, "_rels/$1.rels");
      const chartRelsFile = zip.file(chartRelsPath);

      if (chartRelsFile) {
        const chartRelsXml = relsParser.parse(await chartRelsFile.async("string"));
        const chartRelsList = chartRelsXml.Relationships?.Relationship ?? [];
        const hasImage = chartRelsList.some((r: any) =>
          r["@_Type"]?.includes("/image") || r["@_Type"] === IMAGE_REL_TYPE
        );
        if (hasImage) continue; // Already has fallback
      }

      // Get chart bounds from slide XML
      const slideXmlStr = await zip.file(slidePath)!.async("string");
      const slideXml = slideParser.parse(slideXmlStr);
      const bounds = findChartBounds(slideXml, chartRel["@_Id"]);
      if (!bounds) continue;

      results.push({ slideNum, chartPath, chartRelsPath, bounds });
    }
  }

  return results;
}

/** Find the graphicFrame bounds for a chart by its relationship ID. */
function findChartBounds(
  slideXml: any,
  rId: string,
): { x: number; y: number; cx: number; cy: number } | null {
  const frames = findGraphicFrames(slideXml);
  for (const frame of frames) {
    // Chart ref can be under a:graphic > a:graphicData with various prefixes
    const graphicData = frame["a:graphic"]?.["a:graphicData"]
      ?? frame["p:graphic"]?.["a:graphicData"];
    const chartRef = graphicData?.["c:chart"] ?? graphicData?.["chart"];
    if (!chartRef) continue;
    if (chartRef["@_r:id"] === rId || chartRef["@_id"] === rId) {
      // graphicFrame xfrm is p:xfrm (direct child)
      const xfrm = frame["p:xfrm"] ?? frame["a:xfrm"];
      if (!xfrm) continue;
      const off = xfrm["a:off"];
      const rawExt = xfrm["a:ext"];
      const ext = Array.isArray(rawExt) ? rawExt[0] : rawExt;
      if (!off || !ext) continue;
      return {
        x: Number(off["@_x"] ?? 0),
        y: Number(off["@_y"] ?? 0),
        cx: Number(ext["@_cx"] ?? 0),
        cy: Number(ext["@_cy"] ?? 0),
      };
    }
  }
  return null;
}

function findGraphicFrames(node: any): any[] {
  const results: any[] = [];
  if (!node || typeof node !== "object") return results;
  if (node["p:graphicFrame"]) {
    const frames = Array.isArray(node["p:graphicFrame"]) ? node["p:graphicFrame"] : [node["p:graphicFrame"]];
    results.push(...frames);
  }
  for (const key of Object.keys(node)) {
    if (key.startsWith("@_")) continue;
    const children = Array.isArray(node[key]) ? node[key] : [node[key]];
    for (const child of children) {
      results.push(...findGraphicFrames(child));
    }
  }
  return results;
}

/** Render a chart to PNG using Playwright and the renderer. */
async function renderChartToPng(
  pptxBuffer: Buffer,
  chart: ChartInfo,
  chromium: any,
): Promise<Buffer> {
  const { PptxPackage, readPresentation, generateHtml } = await import("quicklook-pptx-renderer");

  const pkg = await PptxPackage.open(pptxBuffer);
  const pres = await readPresentation(pkg);
  const { html } = await generateHtml(pres, { pkg });

  const EMU_PER_PX = 12700;
  const slideWidth = pres.slideSize.cx / EMU_PER_PX;
  const slideHeight = pres.slideSize.cy / EMU_PER_PX;

  const browser = await chromium.launch();
  try {
    const page = await browser.newPage({ viewport: { width: Math.ceil(slideWidth), height: Math.ceil(slideHeight) } });
    await page.setContent(html, { waitUntil: "networkidle" });

    // Navigate to the correct slide (slides are stacked vertically in the HTML)
    const slideSelector = `.slide:nth-child(${chart.slideNum})`;
    const slideEl = await page.$(slideSelector);
    if (slideEl) await slideEl.scrollIntoViewIfNeeded();

    const clip = {
      x: chart.bounds.x / EMU_PER_PX,
      y: chart.bounds.y / EMU_PER_PX + (chart.slideNum - 1) * slideHeight,
      width: chart.bounds.cx / EMU_PER_PX,
      height: chart.bounds.cy / EMU_PER_PX,
    };

    const png = await page.screenshot({ clip, type: "png" });
    return Buffer.from(png);
  } finally {
    await browser.close();
  }
}

/** Embed a fallback PNG for a chart into the PPTX zip. */
async function embedFallback(
  zip: JSZip,
  chart: ChartInfo,
  png: Buffer,
  index: number,
): Promise<void> {
  const parser = new XMLParser(relsParserOptions);
  const builder = new XMLBuilder(relsBuilderOptions);

  // Add PNG to media folder
  const mediaPath = `ppt/media/chart${index}-fallback.png`;
  zip.file(mediaPath, png);

  // Get or create chart rels
  const chartRelsFile = zip.file(chart.chartRelsPath);
  let relsXml: any;
  if (chartRelsFile) {
    relsXml = parser.parse(await chartRelsFile.async("string"));
  } else {
    relsXml = {
      Relationships: {
        "@_xmlns": "http://schemas.openxmlformats.org/package/2006/relationships",
        Relationship: [],
      },
    };
  }

  if (!relsXml.Relationships) relsXml.Relationships = { "@_xmlns": "http://schemas.openxmlformats.org/package/2006/relationships" };
  if (!relsXml.Relationships.Relationship) relsXml.Relationships.Relationship = [];
  if (!Array.isArray(relsXml.Relationships.Relationship)) {
    relsXml.Relationships.Relationship = [relsXml.Relationships.Relationship];
  }

  // Compute relative path from chart to media
  const relTarget = "../" + mediaPath.slice(4); // strip "ppt/"
  const rId = `rId${relsXml.Relationships.Relationship.length + 1}`;

  relsXml.Relationships.Relationship.push({
    "@_Id": rId,
    "@_Type": IMAGE_REL_TYPE,
    "@_Target": relTarget,
  });

  zip.file(chart.chartRelsPath, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + builder.build(relsXml));

  // Ensure Content_Types.xml includes png
  const ctFile = zip.file("[Content_Types].xml");
  if (ctFile) {
    let ct = await ctFile.async("string");
    if (!ct.includes('Extension="png"')) {
      ct = ct.replace(
        "</Types>",
        '  <Default Extension="png" ContentType="image/png"/>\n</Types>',
      );
      zip.file("[Content_Types].xml", ct);
    }
  }
}

/**
 * Add fallback images to charts that don't have them.
 * Requires Playwright — skips silently if not installed.
 */
export async function addChartFallbacks(
  zip: JSZip,
  pptxBuffer: Buffer,
  reportLines: string[],
): Promise<void> {
  let chromium: any;
  try {
    const pw = await import("playwright" as string);
    chromium = pw.chromium;
  } catch {
    return; // Playwright not installed — skip
  }

  const charts = await findChartsWithoutFallback(zip);
  if (charts.length === 0) return;

  for (let i = 0; i < charts.length; i++) {
    const chart = charts[i];
    try {
      const png = await renderChartToPng(pptxBuffer, chart, chromium);
      await embedFallback(zip, chart, png, i + 1);
      reportLines.push(`Slide ${chart.slideNum}: added chart fallback image (${png.length} bytes)`);
    } catch (err: any) {
      reportLines.push(`Slide ${chart.slideNum}: chart fallback failed — ${err.message}`);
    }
  }
}
