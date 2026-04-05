/**
 * effects transform — Detect shapes with effectLst that will become opaque
 * PDF blocks in QuickLook (covering content behind them).
 *
 * Options: strip effects, reorder z-index, or just warn.
 */
function findEffectShapes(node, slideNum) {
    const issues = [];
    if (!node || typeof node !== "object")
        return issues;
    if (node.effectLst && Object.keys(node.effectLst).some(k => !k.startsWith("@_"))) {
        issues.push({
            type: "effects",
            slide: slideNum,
            severity: "low",
            description: "Shape with effects — will render as opaque PDF block in QuickLook (may cover content behind it)",
        });
    }
    for (const key of Object.keys(node)) {
        if (key.startsWith("@_"))
            continue;
        const children = Array.isArray(node[key]) ? node[key] : [node[key]];
        for (const child of children) {
            issues.push(...findEffectShapes(child, slideNum));
        }
    }
    return issues;
}
export const effects = {
    name: "effects",
    detect(slideXml, slideNum) {
        return findEffectShapes(slideXml, slideNum);
    },
    apply(_slideXml, _slideNum, _ctx) {
        // TODO: implement — configurable: strip effects, reorder, or no-op
        return { changed: false, changes: [] };
    },
};
