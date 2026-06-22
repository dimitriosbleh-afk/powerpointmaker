"use strict";

const path = require("path");
const fs = require("fs");
const { lintTeacherNotesInFile } = require("./qa_lib");

const BUILDS_DIR = path.resolve(__dirname, "..", "builds");
const DEFAULT_FILES = fs.readdirSync(BUILDS_DIR)
  .filter((name) => name.startsWith("build_") && name.endsWith(".js"))
  .sort()
  .map((name) => path.join("builds", name));

const files = process.argv.slice(2);
const targets = files.length > 0 ? files : DEFAULT_FILES;

const issues = [];
targets.forEach((filePath) => {
  issues.push(...lintTeacherNotesInFile(filePath, {
    checkMarkdownHeaders: true,
    checkUnicodeBullets: true,
    checkSmartPunctuation: true,
    checkAscii: true,
    checkSectionStructure: true,
    maxLines: 40,
    maxChars: 2600,
    maxSayBullets: 8,
    maxDoBullets: 8,
    maxWatchForBullets: 5,
    maxTeacherNotesLines: 8,
    maxTeacherNotesChars: 1400,
  }));
});

if (issues.length > 0) {
  console.error("Teacher notes source lint failed:");
  issues.forEach((issue) => console.error(`- ${issue}`));
  process.exit(1);
}

console.log(`Teacher notes source lint passed for ${targets.map((file) => path.basename(file)).join(", ")}.`);
