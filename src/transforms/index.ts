/**
 * Transform registry — each transform detects and fixes one class of
 * OfficeImport rendering issues.
 */

import type { Issue } from "../analyze.js";
import { tableStyles } from "./table-styles.js";
import { gradients } from "./gradients.js";
import { geometries } from "./geometries.js";
import { properties } from "./properties.js";
import { fonts } from "./fonts.js";
import { effects } from "./effects.js";

export type TransformName =
  | "table-styles"
  | "gradients"
  | "geometries"
  | "properties"
  | "fonts"
  | "effects";

export interface TransformContext {
  tableStyleXml?: any;
}

export interface TransformResult {
  changed: boolean;
  changes: string[];
}

export interface Transform {
  name: TransformName;
  detect: (slideXml: any, slideNum: number) => Issue[];
  apply: (slideXml: any, slideNum: number, ctx: TransformContext) => TransformResult;
}

export const ALL_TRANSFORMS: Transform[] = [
  tableStyles,
  properties,
  geometries,
  gradients,
  fonts,
  effects,
];
