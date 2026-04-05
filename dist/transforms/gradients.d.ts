/**
 * gradients transform — Collapse 3+ stop gradients to 2-stop so QuickLook
 * renders a gradient instead of averaging to a flat color.
 *
 * Lossy: intermediate gradient stops are removed.
 */
import type { Transform } from "./index.js";
export declare const gradients: Transform;
