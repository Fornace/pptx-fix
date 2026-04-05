/**
 * fonts transform — Add explicit <a:latin>, <a:ea>, <a:cs> fallback
 * typefaces matching what OfficeImport's TCFontUtils would pick.
 */
export const fonts = {
    name: "fonts",
    detect(_slideXml, _slideNum) {
        // TODO: detect text runs missing explicit font declarations
        return [];
    },
    apply(_slideXml, _slideNum, _ctx) {
        // TODO: implement — add fallback typefaces
        return { changed: false, changes: [] };
    },
};
