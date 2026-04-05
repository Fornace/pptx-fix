/**
 * Analyze a PPTX for issues that will cause QuickLook rendering problems.
 * Read-only — does not modify the file.
 */
import { type TransformName } from "./transforms/index.js";
export interface Issue {
    type: TransformName;
    slide: number;
    element?: string;
    severity: "high" | "medium" | "low";
    description: string;
}
export declare function analyze(pptxBuffer: Buffer): Promise<Issue[]>;
