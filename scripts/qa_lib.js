"use strict";

const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");
const JSZip = require("jszip");
const { getTeacherNotesSourceIssues, isGlanceFormatNotes } = require("../themes/core/notes");

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

/**
 * Extract each slide's speaker notes from a built PPTX as plain text
 * (one string per slide, paragraphs joined with \n, empty paragraphs kept
 * as blank lines). Returns an array indexed by slide order.
 */
async function extractNotesTextPerSlide(pptxPath) {
  const zip = await JSZip.loadAsync(fs.readFileSync(pptxPath));
  const noteFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

  const notesPerSlide = [];
  for (const noteFile of noteFiles) {
    const xml = await zip.file(noteFile).async("string");
    const txBodyMatch = xml.match(/name="Notes Placeholder 2"[\s\S]*?<p:txBody>([\s\S]*?)<\/p:txBody>/);
    if (!txBodyMatch) {
      notesPerSlide.push("");
      continue;
    }
    const paragraphs = [...txBodyMatch[1].matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)];
    const lines = paragraphs.map((para) =>
      [...para[1].matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)]
        .map((run) => unescapeXml(run[1]))
        .join("")
    );
    notesPerSlide.push(lines.join("\n"));
  }
  return notesPerSlide;
}

/**
 * Gate 5 checks (megaprompt v12.3 sections 45-47):
 *  - consecutive slides must not carry identical notes (a reveal slide that
 *    byte-copies its base slide's notes leaves the teacher staring at the
 *    same wall after clicking to the answer)
 *  - Glance Format live zones must respect the rendered budgets (120 words,
 *    16 words per physical line, 18 physical lines) so the notes stay
 *    glanceable on an iPad presenter view
 */
async function validateNotesFormat(pptxPath) {
  const notesPerSlide = await extractNotesTextPerSlide(pptxPath);
  const issues = [];
  const label = path.basename(pptxPath);

  notesPerSlide.forEach((notes, index) => {
    const slideNo = index + 1;
    const normalized = notes.replace(/\s+/g, " ").trim();

    if (index > 0) {
      const prev = notesPerSlide[index - 1].replace(/\s+/g, " ").trim();
      const wordCount = normalized ? normalized.split(" ").length : 0;
      if (normalized && normalized === prev && wordCount >= 15) {
        issues.push(`${label}: slide ${slideNo} duplicates slide ${slideNo - 1}'s notes verbatim - a reveal slide must carry its own post-reveal notes (withReveal opts.revealNotes / composeRevealNotes)`);
      }
    }

    if (isGlanceFormatNotes(notes)) {
      const noteIssues = getTeacherNotesSourceIssues(notes, {
        checkSectionStructure: true,
        checkResponsiveGlance: false, // beat wording is composeGlanceNotes' job; the gate enforces shape
        maxLines: 40,
        maxChars: 2600,
      });
      noteIssues.forEach((issue) => issues.push(`${label}: slide ${slideNo}: ${issue}`));
    }
  });

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
  extractNotesTextPerSlide,
  validateNotesFormat,
  listPdfFiles,
};
