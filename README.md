# pptx-fix

**Fix PowerPoint files that render incorrectly on macOS QuickLook and iOS.** Rewrites PPTX internals so tables, shapes, gradients, and effects display correctly on Apple devices — no manual editing needed.

If your `.pptx` looks perfect in PowerPoint but looks wrong when someone opens it on a Mac or iPhone, this tool fixes it.

---

## The Problem

Apple's QuickLook (macOS Finder spacebar preview), iOS Files app, and iPadOS use a private rendering engine called `OfficeImport.framework` — a completely independent OOXML parser that renders slides differently than Microsoft PowerPoint. Presentations created by python-pptx, PptxGenJS, Google Slides, Canva, LibreOffice, Pandoc, and other tools contain patterns that PowerPoint handles fine but OfficeImport renders incorrectly.

Common artifacts:

- **Table borders vanish** — generators set a table style reference, but OfficeImport only reads explicit border properties
- **Shapes disappear** — ~120 preset geometries (heart, cloud, lightningBolt, sun, moon...) are silently dropped
- **Gradients become flat colors** — 3+ color stops are averaged to a single solid color
- **Drop shadows become opaque blocks** — shapes with effects render as opaque PDF images that cover content behind them
- **Fonts shift and text reflows** — Calibri becomes Helvetica Neue, Arial becomes Helvetica, with different metrics

`pptx-fix` rewrites the OOXML XML inside the PPTX ZIP to work around these OfficeImport quirks. The output is a valid `.pptx` that looks correct in both PowerPoint and Apple's preview.

---

## Install

```bash
npm install pptx-fix
```

## Usage

### CLI

```bash
# Fix a PPTX file
npx pptx-fix input.pptx -o output.pptx

# Apply only specific transforms
npx pptx-fix input.pptx -o output.pptx --only table-styles,gradients

# See what was changed
npx pptx-fix input.pptx -o output.pptx --report

# Analyze without fixing (dry run)
npx pptx-fix analyze input.pptx
```

### As a Library

```typescript
import { fix } from "pptx-fix";
import { readFileSync, writeFileSync } from "fs";

const input = readFileSync("presentation.pptx");
const result = await fix(input, { report: true });

writeFileSync("fixed.pptx", result.buffer);
console.log(result.report);
```

### Analyze Only

```typescript
import { analyze } from "pptx-fix";

const issues = await analyze(pptxBuffer);
for (const issue of issues) {
  console.log(`[${issue.severity}] Slide ${issue.slide}: ${issue.description}`);
}
```

---

## Transforms

| Transform | Status | What it fixes |
|-----------|--------|--------------|
| **table-styles** | Done | Resolves `tableStyleId` references and inlines explicit `<a:lnL/R/T/B>` borders on every cell. Handles `firstRow`, `lastRow`, `bandRow` conditional formatting. Only adds borders where none are explicitly defined. |
| **geometries** | Planned | Replace unsupported `<a:prstGeom>` (heart, cloud, lightningBolt, sun, moon, frame, arc, chord, etc.) with `<a:custGeom>` path data. |
| **gradients** | Planned | Collapse 3+ stop gradients to 2-stop (start + end color) so QuickLook renders a gradient instead of a flat color. |
| **effects** | Planned | Strip or adjust shapes with `<effectLst>` (drop shadow, glow, reflection) that render as opaque PDF blocks covering content. |

Detection of all issues (including font substitution, chart fallbacks, text inscription shifts, and more) is handled by [quicklook-pptx-renderer](https://www.npmjs.com/package/quicklook-pptx-renderer)'s 12-rule linter. Run `pptx-fix analyze` to see all issues, or use the linter directly in CI.

---

## Which Tools Produce Affected Files

| Tool | Primary issue on Mac/iPhone |
|------|---|
| **python-pptx** | Tables render without borders — the #1 reported issue. Style references not resolved by OfficeImport. |
| **PptxGenJS** | Missing thumbnails, shape rendering differences, effect artifacts |
| **Google Slides** export | Font substitution, formatting shifts, missing content types |
| **Canva** export | Fonts substituted, layout differences, animation artifacts |
| **LibreOffice Impress** | Table styles unresolved, gradient rendering differences |
| **Pandoc** / **Quarto** | Content type corruption, missing shapes, "PowerPoint found a problem with content" errors |
| **Apache POI** | Content type errors (`InvalidFormatException: Package should contain a content type part [M1.13]`) |
| **Aspose.Slides** | Missing thumbnail if `refresh_thumbnail` not called |
| **Open XML SDK** | Repair errors, missing relationships |

---

## How It Works

```
PPTX (ZIP) → extract XML → parse → detect issues → apply transforms → serialize → repack ZIP
```

1. **Extract** — JSZip opens the PPTX (which is a ZIP archive)
2. **Parse** — fast-xml-parser converts slide XML to objects, preserving all unknown elements
3. **Detect** — each transform scans for its class of issues
4. **Apply** — transforms mutate the XML objects (e.g., inline borders from table style definitions)
5. **Serialize** — XMLBuilder converts back to XML
6. **Repack** — JSZip produces a new valid PPTX

The round-trip preserves all XML elements the tool doesn't explicitly modify — the output is always a valid PPTX.

---

## Detecting Issues Without Fixing

For comprehensive linting with 12 rules, CI integration, cross-platform rendering, and pixel-diff comparison against actual QuickLook output, see [**quicklook-pptx-renderer**](https://www.npmjs.com/package/quicklook-pptx-renderer) — a companion renderer + linter that replicates Apple's QuickLook output pixel for pixel, runs on Linux/Docker without a Mac.

```bash
# Lint (12 rules, JSON output for CI)
npx quicklook-pptx lint presentation.pptx --json

# Render slides as PNG (see exactly what Mac users see)
npx quicklook-pptx render presentation.pptx --out ./slides/

# Fix (this package)
npx pptx-fix presentation.pptx -o fixed.pptx
```

---

## Dependencies

| Package | Purpose |
|---------|---------|
| [jszip](https://stuk.github.io/jszip/) | ZIP extraction/repacking |
| [fast-xml-parser](https://github.com/NaturalIntelligence/fast-xml-parser) | OOXML XML parsing and serialization |

Zero native dependencies. Works everywhere Node.js 20+ runs.

## License

MIT
