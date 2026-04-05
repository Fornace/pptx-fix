/**
 * geometries transform — Replace unsupported preset geometries with rect
 * so shapes are visible in QuickLook instead of silently invisible.
 *
 * Detection is handled by quicklook-pptx-renderer's linter (unsupported-geometry rule).
 */

import type { Transform, TransformContext, TransformResult } from "./index.js";

// The 60 preset geometries that OfficeImport's CMCanonicalShapeBuilder supports.
// Shapes not in this set produce nil CGPathRef and are invisible.
const SUPPORTED = new Set([
  // Basic shapes
  "rect", "roundRect", "ellipse", "diamond", "triangle", "rtTriangle",
  "parallelogram", "trapezoid", "hexagon", "octagon", "plus", "pentagon",
  "chevron", "homePlate", "cube", "can",
  // Stars
  "star4", "star5", "star6", "star8", "star10", "star12", "star24",
  // Arrows
  "rightArrow", "leftArrow", "upArrow", "downArrow",
  "leftRightArrow", "upDownArrow", "notchedRightArrow", "stripedRightArrow",
  "bentArrow", "uturnArrow", "curvedRightArrow", "curvedLeftArrow",
  "curvedUpArrow", "curvedDownArrow",
  // Callouts
  "wedgeRoundRectCallout", "wedgeRectCallout", "wedgeEllipseCallout",
  "cloudCallout",
  // Flowchart
  "flowChartProcess", "flowChartDecision", "flowChartTerminator",
  "flowChartDocument", "flowChartPreparation",
  // Block
  "heart", "lightningBolt", "sun", "moon", "cloud",
  "arc", "donut", "noSmoking", "blockArc",
  // Math
  "mathPlus", "mathMinus", "mathMultiply", "mathDivide", "mathEqual",
]);

function fixGeometries(node: any, changes: string[]): void {
  if (!node || typeof node !== "object") return;

  if (node.sp) {
    const shapes = Array.isArray(node.sp) ? node.sp : [node.sp];
    for (const shape of shapes) {
      const spPr = shape.spPr;
      if (!spPr?.prstGeom) continue;
      const prst = spPr.prstGeom["@_prst"];
      if (!prst || SUPPORTED.has(prst)) continue;

      const name = shape.nvSpPr?.cNvPr?.["@_name"] ?? "shape";
      spPr.prstGeom["@_prst"] = "rect";
      changes.push(`replaced unsupported geometry "${prst}" with rect on "${name}"`);
    }
  }

  if (node.grpSp) {
    const groups = Array.isArray(node.grpSp) ? node.grpSp : [node.grpSp];
    for (const group of groups) {
      fixGeometries(group, changes);
    }
  }

  if (node.spTree) {
    fixGeometries(node.spTree, changes);
  }
}

export const geometries: Transform = {
  name: "geometries",

  apply(slideXml: any, _slideNum: number, _ctx: TransformContext): TransformResult {
    const changes: string[] = [];
    fixGeometries(slideXml, changes);
    return { changed: changes.length > 0, changes };
  },
};
