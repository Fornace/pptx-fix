/**
 * properties transform — Resolve full inheritance chain
 * (theme → master → layout → slide) and write explicit fill, font, color
 * on each element.
 */
export const properties = {
    name: "properties",
    detect(_slideXml, _slideNum) {
        // TODO: detect elements relying on inherited properties
        return [];
    },
    apply(_slideXml, _slideNum, _ctx) {
        // TODO: implement — resolve inheritance and inline explicit properties
        return { changed: false, changes: [] };
    },
};
