/**
 * Shared XML parser/builder options matching OOXML conventions.
 *
 * IMPORTANT: removeNSPrefix must be false to preserve namespace prefixes
 * (p:, a:, r:) during round-trip. Without them, PowerPoint rejects the file.
 */

export const xmlParserOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  removeNSPrefix: false,
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
  isArray: (name: string) => ARRAY_ELEMENTS.has(name),
};

export const xmlBuilderOptions = {
  ignoreAttributes: false,
  attributeNamePrefix: "@_",
  format: false,
  suppressEmptyNode: false,
  suppressBooleanAttributes: false,
};

/** Elements that must always be parsed as arrays (even with a single child). */
const ARRAY_ELEMENTS = new Set([
  "p:sp", "p:pic", "p:cxnSp", "p:grpSp", "p:graphicFrame",
  "a:p", "a:r", "a:br", "a:fld",
  "a:gs", "a:ln", "a:solidFill", "a:gradFill",
  "a:tr", "a:tc",
  "a:tblStyleLst",
  "a:ext",
]);
