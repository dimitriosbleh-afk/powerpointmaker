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

/**
 * Extract each slide's FACE text from a built PPTX, one string per slide.
 *
 * Assumes slideN.xml order is presentation order, which holds for decks
 * PptxGenJS has just written. Merged units are validated by qa_unit.js, which
 * does not use this.
 */
async function extractSlideTextPerSlide(pptxPath) {
  const zip = await JSZip.loadAsync(fs.readFileSync(pptxPath));
  const slideFiles = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));

  const out = [];
  for (const slideFile of slideFiles) {
    const xml = await zip.file(slideFile).async("string");
    const runs = [...xml.matchAll(/<a:t>([\s\S]*?)<\/a:t>/g)].map((m) => unescapeXml(m[1]));
    out.push(runs.join(" ").replace(/\s+/g, " ").trim());
  }
  return out;
}

// Scaffolding wording legitimately repeats between We Do and You Do (step
// words, criteria reminders, the unit anchor phrase). Stripping it before
// comparing is what keeps the similarity check from crying wolf.
const SCAFFOLD_STOPWORDS = new Set([
  "first", "next", "then", "finally", "i", "can", "we", "do", "you", "your",
  "the", "a", "an", "and", "or", "to", "of", "in", "on", "it", "is", "are",
  "with", "for", "this", "that", "show", "me", "write", "draw", "use", "check",
  "partner", "board", "boards", "book", "books", "success", "criteria", "task",
  "step", "steps", "now", "try", "turn", "tell", "think", "talk",
]);

function contentTokens(text) {
  return new Set(
    String(text || "")
      .toLowerCase()
      .replace(/[^a-z0-9\s.\/]/g, " ")
      .split(/\s+/)
      .filter((t) => t.length > 1 && !SCAFFOLD_STOPWORDS.has(t))
  );
}

function jaccard(a, b) {
  if (a.size === 0 || b.size === 0) return 0;
  let shared = 0;
  a.forEach((t) => { if (b.has(t)) shared += 1; });
  return shared / (a.size + b.size - shared);
}

/**
 * Gate 6 checks: lesson structure rules that were previously honour-system.
 *
 * Covers megaprompt section 0a item 19 / section 44 (Teacher Resources near
 * the front), section 0a item 23 / section 21 (Daily Review and Fluency sit
 * before the LI and SC slide), section 23 (a maths deck has Fluency), and
 * section 35 / section 64 (We Do and You Do must not use the same content).
 *
 * Only runs on decks that look like a lesson, i.e. that carry a Learning
 * Intention slide. Visual catalogues, smoke tests and fragments are skipped
 * so the gate never fires on something it was not written for.
 */
async function validateLessonStructure(pptxPath) {
  const slides = await extractSlideTextPerSlide(pptxPath);
  const label = path.basename(pptxPath);
  const issues = [];
  if (slides.length === 0) return issues;

  const findIndex = (re) => slides.findIndex((t) => re.test(t));
  const liIndex = findIndex(/Learning Intention/i);
  if (liIndex === -1) return issues; // not a lesson deck

  const resourcesIndex = findIndex(/Teacher Resources|Printable Resources/i);
  const dailyReviewIndex = findIndex(/Daily Review/i);
  const fluencyIndex = findIndex(/\bFluency\b/i);

  // Section 0a item 19: resources belong at the front, so the teacher is
  // prepared before instruction starts. Allowing the first four slides leaves
  // room for a session divider or a teacher-facing overview.
  if (resourcesIndex === -1) {
    issues.push(`${label}: no Teacher Resources slide found (megaprompt section 44 requires one, even when no printing is needed)`);
  } else if (resourcesIndex > 3) {
    issues.push(`${label}: Teacher Resources is slide ${resourcesIndex + 1}; it belongs immediately after the title slide (megaprompt section 0a item 19)`);
  }

  // Section 0a item 23: the opening order is fixed. Review and fluency
  // retrieve prior learning BEFORE today's intention is shared.
  if (dailyReviewIndex !== -1 && dailyReviewIndex > liIndex) {
    issues.push(`${label}: Daily Review (slide ${dailyReviewIndex + 1}) comes after the LI and SC slide (slide ${liIndex + 1}); the opening order in megaprompt section 0a item 23 is fixed`);
  }
  if (fluencyIndex !== -1 && fluencyIndex > liIndex) {
    issues.push(`${label}: Fluency (slide ${fluencyIndex + 1}) comes after the LI and SC slide (slide ${liIndex + 1}); Fluency sits after Daily Review and before LI and SC (megaprompt section 23)`);
  }
  // Section 23: every maths lesson includes Fluency. Daily Review is the
  // reliable marker that this is a maths deck.
  if (dailyReviewIndex !== -1 && fluencyIndex === -1) {
    issues.push(`${label}: deck has Daily Review but no Fluency slide (megaprompt section 23: every maths lesson includes Fluency, and it is not a second Daily Review)`);
  }

  // Section 35: You Do must use different content from We Do.
  //
  // Only slides badged as exactly one stage can be compared. A slide carrying a
  // combined "We Do / You Do" badge belongs to both lists, which would make it
  // its own counterexample and report 100% every time.
  const weDo = [];
  const youDo = [];
  slides.forEach((text, i) => {
    const isWe = /\bWe Do\b/i.test(text);
    const isYou = /\bYou Do\b/i.test(text);
    if (isWe && isYou) return; // combined badge - stage is ambiguous
    if (isWe) weDo.push({ i, text });
    if (isYou) youDo.push({ i, text });
  });
  weDo.forEach((w) => {
    youDo.forEach((y) => {
      if (y.i === w.i) return;              // never compare a slide with itself
      if (Math.abs(y.i - w.i) === 1) return; // adjacent slides are a reveal pair
      const score = jaccard(contentTokens(w.text), contentTokens(y.text));
      if (score >= 0.8) {
        issues.push(`${label}: You Do (slide ${y.i + 1}) reuses We Do (slide ${w.i + 1}) content (${Math.round(score * 100)}% shared after scaffolding is stripped); megaprompt section 35 requires changed numbers, text, visual or context`);
      }
    });
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
  extractSlideTextPerSlide,
  validateNotesFormat,
  validateLessonStructure,
  listPdfFiles,
};
