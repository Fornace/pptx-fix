/**
 * Transform registry — each transform detects and fixes one class of
 * OfficeImport rendering issues.
 */
import { tableStyles } from "./table-styles.js";
import { gradients } from "./gradients.js";
import { geometries } from "./geometries.js";
import { properties } from "./properties.js";
import { fonts } from "./fonts.js";
import { effects } from "./effects.js";
export const ALL_TRANSFORMS = [
    tableStyles,
    properties,
    geometries,
    gradients,
    fonts,
    effects,
];
