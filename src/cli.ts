#!/usr/bin/env node

import { readFile, writeFile } from "node:fs/promises";
import { resolve } from "node:path";
import { fix } from "./writer.js";
import { analyze, formatIssues } from "./analyze.js";
import type { TransformName } from "./transforms/index.js";

const args = process.argv.slice(2);

function getFlag(name: string): string | undefined {
  const i = args.indexOf(name);
  if (i === -1 || i + 1 >= args.length) return undefined;
  return args[i + 1];
}

function hasFlag(name: string): boolean {
  return args.includes(name);
}

function positional(): string[] {
  const out: string[] = [];
  for (let i = 0; i < args.length; i++) {
    if (args[i].startsWith("--") || args[i] === "-o") { i++; continue; }
    out.push(args[i]);
  }
  return out;
}

async function main() {
  if (hasFlag("--version") || hasFlag("-v")) {
    console.log("0.2.1");
    process.exit(0);
  }

  const pos = positional();
  const command = pos[0];

  if (command === "analyze") {
    const input = pos[1];
    if (!input) {
      console.error("Usage: pptx-fix analyze <input.pptx>");
      process.exit(1);
    }
    const buf = await readFile(resolve(input));
    const result = await analyze(buf);
    console.log(formatIssues(result));
    process.exit(result.summary.errors > 0 ? 1 : 0);
  }

  // Default: fix
  const input = command;
  if (!input || input.startsWith("-")) {
    console.error("Usage: pptx-fix <input.pptx> -o <output.pptx> [--only table-styles,gradients] [--report]");
    console.error("       pptx-fix analyze <input.pptx>");
    process.exit(1);
  }

  const outPath = getFlag("-o") ?? getFlag("--out");
  if (!outPath) {
    console.error("Error: output path required (-o <output.pptx>)");
    process.exit(1);
  }

  const onlyStr = getFlag("--only");
  const transforms = onlyStr
    ? onlyStr.split(",").map(s => s.trim()) as TransformName[]
    : undefined;

  const wantReport = hasFlag("--report");

  const buf = await readFile(resolve(input));
  const result = await fix(buf, { transforms, report: wantReport });
  await writeFile(resolve(outPath), result.buffer);

  if (wantReport && result.report) {
    console.log(result.report);
  }

  console.log(`Fixed → ${outPath}`);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
