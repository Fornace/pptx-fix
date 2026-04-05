/**
 * geometries transform — Replace unsupported preset geometries with
 * equivalent <a:custGeom> path data.
 *
 * OfficeImport's CMCanonicalShapeBuilder silently drops ~30 presets
 * (heart, cloud, lightningBolt, etc.). This converts them to custom
 * geometry paths that OfficeImport can render.
 */
/** Presets that CMCanonicalShapeBuilder does NOT support. */
const UNSUPPORTED_PRESETS = new Set([
    "heart", "cloud", "lightningBolt", "sun", "moon",
    "irregularSeal1", "plaque", "frame", "halfFrame",
    "corner", "diagStripe", "chord", "arc",
    "bracketPair", "bracePair",
]);
function findUnsupportedGeometries(node, slideNum) {
    const issues = [];
    if (!node || typeof node !== "object")
        return issues;
    if (node.prstGeom) {
        const prst = node.prstGeom["@_prst"];
        if (prst && UNSUPPORTED_PRESETS.has(prst)) {
            issues.push({
                type: "geometries",
                slide: slideNum,
                element: prst,
                severity: "high",
                description: `Preset '${prst}' is not supported by QuickLook — shape will be invisible`,
            });
        }
    }
    for (const key of Object.keys(node)) {
        if (key.startsWith("@_"))
            continue;
        const children = Array.isArray(node[key]) ? node[key] : [node[key]];
        for (const child of children) {
            issues.push(...findUnsupportedGeometries(child, slideNum));
        }
    }
    return issues;
}
export const geometries = {
    name: "geometries",
    detect(slideXml, slideNum) {
        return findUnsupportedGeometries(slideXml, slideNum);
    },
    apply(_slideXml, _slideNum, _ctx) {
        // TODO: implement — replace prstGeom with custGeom containing equivalent path data
        return { changed: false, changes: [] };
    },
};
