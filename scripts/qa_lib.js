"use strict";

const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");
const JSZip = require("jszip");
const { getTeacherNotesSourceIssues } = require("../themes/core/notes");

const ROOT = path.resolve(__dirname, "..");
const PYTHON_ENV = { ...process.env, PYTHONUTF8: "1" };

const FORBIDDEN_OUTPUT_PATTERNS = [
  { regex: /\bTODO\b/i, label: "TODO marker" },
  { regex: /\bTBD\b/i, label: "TBD marker" },
  { regex: /\bFIXME\b/i, label: "FIXME marker" },
  { regex: /but wait\.\.\./i, label: "unfinished 'but wait...'" },
  { regex: /\bSR1\b|\bSR2\b|\bEXT1\b/, label: "legacy resource code" },
  { regex: /\bSupporting Resource\b/i, label: "legacy supporting-resource label" },
];

function runCommand(command, args, opts) {
  const result = spawnSync(command, args, {
    cwd: ROOT,
    encoding: "utf8",
    env: command === "python" ? PYTHON_ENV : process.env,
    ...opts,
  });

  if (result.error) {
    throw result.error;
  }

  const output = `${result.stdout || ""}${result.stderr || ""}`;
  if (result.status !== 0) {
    const error = new Error(`${command} ${args.join(" ")} failed with exit code ${result.status}`);
    error.output = output;
    throw error;
  }

  return output;
}

function lintTeacherNotesInFile(filePath, opts) {
  const absPath = path.resolve(ROOT, filePath);
  const source = fs.readFileSync(absPath, "utf8");
  const issues = [];
  const noteRegex = /const\s+(NOTES_[A-Z0-9_]+)\s*=\s*`([\s\S]*?)`;/g;
  let match = noteRegex.exec(source);

  while (match) {
    const [, name, noteBody] = match;
    const noteIssues = getTeacherNotesSourceIssues(noteBody, opts);
    noteIssues.forEach((issue) => issues.push(`${filePath}:${name}: ${issue}`));
    match = noteRegex.exec(source);
  }

  return issues;
}

function extractText(filePath) {
  return runCommand("python", ["-m", "markitdown", filePath], { timeout: 120000 });
}

function scanTextForForbiddenOutput(text, fileLabel) {
  const issues = [];
  FORBIDDEN_OUTPUT_PATTERNS.forEach(({ regex, label }) => {
    if (regex.test(text)) {
      issues.push(`${fileLabel}: found ${label}`);
    }
  });
  return issues;
}

function unescapeXml(text) {
  return text
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

async function validateNotesXml(pptxPath) {
  const zip = await JSZip.loadAsync(fs.readFileSync(pptxPath));
  const noteFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

  const issues = [];
  let sawParagraphStructuredNotes = false;

  for (const noteFile of noteFiles) {
    const xml = await zip.file(noteFile).async("string");
    const txBodyMatch = xml.match(/name="Notes Placeholder 2"[\s\S]*?<p:txBody>([\s\S]*?)<\/p:txBody>/);
    if (!txBodyMatch) {
      issues.push(`${path.basename(pptxPath)}:${noteFile}: notes placeholder body missing`);
      continue;
    }

    const txBody = txBodyMatch[1];
    const paragraphs = txBody.match(/<a:p>/g) || [];
    if (paragraphs.length > 2) {
      sawParagraphStructuredNotes = true;
    }

    const textRuns = [...txBody.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)];
    textRuns.forEach((match) => {
      const text = unescapeXml(match[1]);
      if (/[\r\n]/.test(text)) {
        issues.push(`${path.basename(pptxPath)}:${noteFile}: note text run contains embedded newline characters`);
      }
      if (/[^\x09\x20-\x7E]/.test(text)) {
        issues.push(`${path.basename(pptxPath)}:${noteFile}: note text run contains non-ASCII characters`);
      }
    });
  }

  if (!sawParagraphStructuredNotes) {
    issues.push(`${path.basename(pptxPath)}: no notes slide contained paragraph-structured notes`);
  }

  return issues;
}

function listPdfFiles(dirPath) {
  return fs.readdirSync(dirPath)
    .filter((name) => name.toLowerCase().endsWith(".pdf"))
    .map((name) => path.join(dirPath, name));
}

module.exports = {
  ROOT,
  runCommand,
  lintTeacherNotesInFile,
  extractText,
  scanTextForForbiddenOutput,
  validateNotesXml,
  listPdfFiles,
};
