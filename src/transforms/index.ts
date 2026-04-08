/**
 * Transform registry — each transform fixes one class of
 * OfficeImport rendering issues by mutating raw OOXML.
 *
 * Detection is handled by quicklook-pptx-renderer's linter.
 * These transforms only apply fixes.
 */

import { tableStyles } from "./table-styles.js";
import { gradients } from "./gradients.js";
import { geometries } from "./geometries.js";
import { effects } from "./effects.js";
import { fonts } from "./fonts.js";
import { groups } from "./groups.js";
import { textFit } from "./text-fit.js";
export type TransformName =
  | "table-styles"
  | "gradients"
  | "geometries"
  | "effects"
  | "fonts"
  | "groups"
  | "text-fit"
  | "embedded-fonts";

export interface TransformContext {
  tableStyleXml?: any;
  themeXml?: any;
}

export interface TransformResult {
  changed: boolean;
  changes: string[];
}

export interface Transform {
  name: TransformName;
  apply: (slideXml: any, slideNum: number, ctx: TransformContext) => TransformResult;
}

export const ALL_TRANSFORMS: Transform[] = [
  tableStyles,
  gradients,
  geometries,
  effects,
  fonts,
  groups,
  textFit,
];
