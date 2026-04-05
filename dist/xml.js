/**
 * Shared XML parser/builder options matching OOXML conventions.
 */
export const xmlParserOptions = {
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    removeNSPrefix: true,
    parseTagValue: false,
    parseAttributeValue: false,
    trimValues: false,
    isArray: (name) => ARRAY_ELEMENTS.has(name),
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
    "sp", "pic", "cxnSp", "grpSp", "graphicFrame",
    "p", "r", "br", "fld",
    "gs", "ln", "solidFill", "gradFill",
    "tr", "tc",
    "tblStyleLst",
    "ext",
]);
