"use strict";

const fs = require("fs");
const {
  getRootPath,
  buildOutputString,
} = require("./prompt_builder_data_lib");

function main() {
  const targetPath = getRootPath("prompt_builder_data.js");
  if (!fs.existsSync(targetPath)) {
    throw new Error("prompt_builder_data.js is missing. Run `node scripts/generate_prompt_builder_data.js`.");
  }

  const current = fs.readFileSync(targetPath, "utf8");
  const match = current.match(/"generatedAt": "([^"]+)"/);
  const stableTimestamp = match ? match[1] : "1970-01-01T00:00:00.000Z";
  const expected = buildOutputString(stableTimestamp);

  if (current !== expected) {
    throw new Error(
      "prompt_builder_data.js is out of sync with CLAUDE.md or megapromptlean_3.0.md. Run `node scripts/generate_prompt_builder_data.js`."
    );
  }

  console.log("Prompt Builder data is in sync.");
}

try {
  main();
} catch (error) {
  console.error(error.message || error);
  process.exit(1);
}
