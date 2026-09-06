"use strict";

/**
 * Build a lesson from a spec file.
 *
 *   node scripts/build_lesson.js builds/exemplar_foundation_numeracy_making_10.json
 *
 * Validates the spec (every problem listed with its field path and the fix),
 * builds the deck and companion PDFs into output/<outputFolder>/, and prints
 * "PPTX written to ..." so scripts/build_and_check.js can run the gates:
 *
 *   node scripts/build_and_check.js builds/exemplar_foundation_numeracy_making_10.json
 *
 * Exit codes: 0 built, 1 spec invalid or build failed, 2 bad usage.
 */

const { buildLesson, loadSpec } = require("../themes/lesson/buildLesson");

async function main() {
  const specPath = process.argv[2];
  if (!specPath) {
    console.error("Usage: node scripts/build_lesson.js <spec.json>");
    process.exit(2);
  }
  let spec;
  try {
    spec = loadSpec(specPath);
  } catch (err) {
    console.error(`Could not read spec ${specPath}: ${err.message}`);
    process.exit(2);
  }
  try {
    await buildLesson(spec);
  } catch (err) {
    if (err.specErrors) {
      console.error("SPEC INVALID:");
      err.specErrors.forEach((e) => console.error("  - " + e));
    } else {
      console.error(err && err.stack ? err.stack : err);
    }
    process.exit(1);
  }
}

main();
