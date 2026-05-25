"use strict";

const path = require("path");
const {
  ROOT,
  runCommand,
  extractText,
  scanTextForForbiddenOutput,
  validateNotesXml,
} = require("./qa_lib");

const DEFAULT_SMOKE_OUTPUTS = [
  "output/_smoke_test_foundation_numeracy/smoke_test.pptx",
  "output/_smoke_test_y56_literacy/smoke_test.pptx",
];

async function validatePptx(pptxPath, issues) {
  const text = extractText(pptxPath);
  issues.push(...scanTextForForbiddenOutput(text, path.basename(pptxPath)));
  issues.push(...await validateNotesXml(pptxPath));
}

async function main() {
  const [buildScript, pptxPath] = process.argv.slice(2);
  const issues = [];

  if (buildScript || pptxPath) {
    if (!buildScript || !pptxPath) {
      throw new Error("Usage: node scripts/qa_notes_regression.js [build-script pptx-path]");
    }
    console.log(`Building notes regression target ${buildScript}...`);
    const output = runCommand("node", [buildScript], { timeout: 180000 });
    process.stdout.write(output);
    await validatePptx(path.resolve(ROOT, pptxPath), issues);
  } else {
    console.log("Building shared notes smoke decks...");
    const output = runCommand("node", ["tests/test_megaprompt_v3_apis.js"], { timeout: 180000 });
    process.stdout.write(output);
    for (const relativePath of DEFAULT_SMOKE_OUTPUTS) {
      await validatePptx(path.join(ROOT, relativePath), issues);
    }
  }

  if (issues.length > 0) {
    console.error("Notes regression QA failed:");
    issues.forEach((issue) => console.error(`- ${issue}`));
    process.exit(1);
  }

  console.log("Notes regression QA passed.");
}

main().catch((error) => {
  console.error(error.message || error);
  if (error.output) process.stderr.write(error.output);
  process.exit(1);
});
