"use strict";

const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const ROOT = path.resolve(__dirname, "..");
const SOURCE_FILES = [
  "CLAUDE.md",
  "megapromptlean_3.0.md",
  "runtime_spine.md",
];

function getRootPath(...parts) {
  return path.join(ROOT, ...parts);
}

function readSourceFile(fileName) {
  return fs.readFileSync(getRootPath(fileName), "utf8").replace(/^\uFEFF/, "");
}

function hashText(text) {
  return crypto.createHash("sha256").update(text, "utf8").digest("hex");
}

function buildSourceMeta(fileName) {
  const absPath = getRootPath(fileName);
  const stats = fs.statSync(absPath);
  const text = readSourceFile(fileName);
  return {
    file: fileName,
    mtimeMs: stats.mtimeMs,
    sha256: hashText(text),
  };
}

function buildPayload(nowIso) {
  return {
    generatedAt: nowIso || new Date().toISOString(),
    sources: SOURCE_FILES.map(buildSourceMeta),
    claude: readSourceFile("CLAUDE.md"),
    megaprompt: readSourceFile("megapromptlean_3.0.md"),
    runtimeSpine: readSourceFile("runtime_spine.md"),
  };
}

function buildOutputString(nowIso) {
  return `window.PROMPT_BUILDER_DATA = ${JSON.stringify(buildPayload(nowIso), null, 2)};\n`;
}

module.exports = {
  ROOT,
  SOURCE_FILES,
  getRootPath,
  readSourceFile,
  hashText,
  buildPayload,
  buildOutputString,
};
