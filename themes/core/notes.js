"use strict";

const fs = require("fs");
const JSZip = require("jszip");

const NOTE_SECTION_HEADERS = [
  "SAY",
  "DO",
  "PACING OVERVIEW",
  "CFU CHECKPOINT",
  "TEACHER NOTES",
  "ENABLING & EXTENDING",
  "MISCONCEPTIONS",
  "SENSITIVITY ADVISORY",
  "WATCH FOR",
  "SOURCES",
];

const ASCII_REPLACEMENTS = [
  [/\u00C3\u0192\u00E2\u20AC\u201D/g, "x"],
  [/\u00C3\u0192\u00C2\u00B7/g, " divided by "],
  [/â€¢/g, "-"],
  [/â€“|â€”/g, "-"],
  [/â€˜|â€™/g, "'"],
  [/â€œ|â€�/g, '"'],
  [/â€¦/g, "..."],
  [/Ã—/g, "x"],
  [/Ã·/g, " divided by "],
  [/Â/g, ""],
  [/[\u2018\u2019\u201A\u201B\u2032]/g, "'"],
  [/[\u201C\u201D\u201E\u201F\u2033]/g, '"'],
  [/[\u2010\u2011\u2012\u2013\u2014\u2015\u2212]/g, "-"],
  [/\u2026/g, "..."],
  [/[\u2022\u2023\u25E6\u2043\u2219\u25CF\u25AA]/g, "-"],
  [/\u2192/g, "->"],
  [/\u2190/g, "<-"],
  [/\u21D2/g, "=>"],
  [/\u21D0/g, "<="],
  [/\u2265/g, ">="],
  [/\u2264/g, "<="],
  [/\u2260/g, "!="],
  [/\u00D7/g, "x"],
  [/\u00F7/g, " divided by "],
  [/\u00A0/g, " "],
  [/[\u200B-\u200D\uFEFF]/g, ""],
];

const HEADER_PATTERN = new RegExp(
  `^(${NOTE_SECTION_HEADERS.map(escapeRegex).join("|")})\\s*:?$`,
  "i"
);

let notesPatchInstalled = false;

function escapeRegex(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function stripMarkdown(value) {
  return value
    .replace(/```[\s\S]*?```/g, (block) => block.replace(/```/g, ""))
    .replace(/`([^`]+)`/g, "$1")
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/__([^_]+)__/g, "$1")
    .replace(/\*([^*\n]+)\*/g, "$1")
    .replace(/_([^_\n]+)_/g, "$1")
    .replace(/^\s{0,3}#{1,6}\s+/gm, "");
}

function toAscii(value) {
  let next = value;
  for (const [pattern, replacement] of ASCII_REPLACEMENTS) {
    next = next.replace(pattern, replacement);
  }

  return next
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^\x09\x0A\x0D\x20-\x7E]/g, "")
    .replace(/[ \t]+divided by[ \t]+/g, " divided by ");
}

function normalizeHeader(line) {
  const stripped = stripMarkdown(line).trim();
  const match = stripped.match(HEADER_PATTERN);
  return match ? `${match[1].toUpperCase()}:` : null;
}

function normalizeBullet(line) {
  // Glance Format beats are numbered ("1. POINT..."); the number IS the
  // structure, so numbered lines are preserved (normalising "1)" to "1.").
  const beatMatch = line.match(/^\s*(\d+)[.)]\s+(.*)$/);
  if (beatMatch) return `${beatMatch[1]}. ${beatMatch[2].trimEnd()}`;
  const bulletMatch = line.match(/^\s*[-*]\s+(.*)$/);
  if (!bulletMatch) return line.trimEnd();
  return `- ${bulletMatch[1].trimEnd()}`;
}

// Glance Format detection: a "---" divider line, an ANSWER: first line, or
// numbered beats opening with a CAPS anchor ("1. POINT ...", "2. ASK: ...").
function isGlanceFormatNotes(rawText) {
  const lines = String(rawText || "").replace(/\r\n?/g, "\n").split("\n");
  const nonBlank = lines.map((line) => line.trim()).filter(Boolean);
  if (nonBlank.length === 0) return false;
  if (nonBlank.some((line) => line === "---")) return true;
  if (/^ANSWER:/.test(nonBlank[0])) return true;
  return nonBlank.some((line) => /^\d+\.\s+[A-Z]{2,}/.test(line));
}

function sanitizeTeacherNotes(notes) {
  if (notes == null || notes === "") return notes;

  const rawLines = toAscii(stripMarkdown(String(notes)).replace(/\r\n?/g, "\n"))
    .split("\n")
    .map((line) => line.replace(/\t/g, "  ").trimEnd());

  const normalized = [];
  for (const rawLine of rawLines) {
    const line = rawLine.trim();
    const header = normalizeHeader(rawLine);

    if (header) {
      if (normalized.length && normalized[normalized.length - 1] !== "") {
        normalized.push("");
      }
      normalized.push(header);
      continue;
    }

    if (!line) {
      if (normalized.length && normalized[normalized.length - 1] !== "") {
        normalized.push("");
      }
      continue;
    }

    normalized.push(normalizeBullet(rawLine));
  }

  while (normalized.length && normalized[normalized.length - 1] === "") {
    normalized.pop();
  }

  return normalized.join("\n");
}

function parseNotesSections(notes) {
  const sanitized = sanitizeTeacherNotes(notes || "");
  const lines = sanitized ? sanitized.split("\n") : [];
  const sections = [];
  let current = null;

  lines.forEach((line) => {
    const trimmed = line.trim();
    if (HEADER_PATTERN.test(trimmed.replace(/:$/, "")) || NOTE_SECTION_HEADERS.some((header) => `${header}:` === trimmed)) {
      current = {
        name: trimmed.endsWith(":") ? trimmed : `${trimmed}:`,
        lines: [],
      };
      sections.push(current);
      return;
    }

    if (current) {
      current.lines.push(line);
    }
  });

  return sections;
}

const RESPONSE_ROUTINE_RE = /\b(?:boards?(?:\s+up)?|mini[- ]whiteboards?|choral(?:\s+(?:response|read))?|everyone\s+(?:points?|shows?|writes?|reads?|acts?)|fingers?|thumbs?|turn\s+and\s+tell|pair(?:\s+(?:share|check))?|partner(?:\s+(?:talk|check|share))?|cold\s+call|stand\s+if|hold\s+up|quick\s+(?:write|sketch)|sort|match|act\s+it\s+out|read\s+aloud|show\s+me|chin\s+it|write\s+it)\b/i;
const THINK_TIME_RE = /\b\d+(?:\.\d+)?\s*(?:sec|secs|second|seconds)\b/i;

// A live-zone logical unit starts with ANSWER:, a beat number, or a labelled
// anchor line. Any other physical line (indented continuation, cue line,
// EXPECT: line, pivot branch) attaches to the unit above it. This is what
// lets the airy v12.3 format split one beat across several short lines while
// validation still reasons about whole beats.
const UNIT_START_RE = /^(?:ANSWER:|REVEALED[.:]|\d+[.)]\s|TRAP:|STRETCH:|HELP:|CARE:|SCAN\b|REVEAL\b)/;

function groupGlanceUnits(liveLines) {
  const units = [];
  let current = null;
  liveLines.forEach((line) => {
    const trimmed = line.trim();
    if (!trimmed) return;
    if (UNIT_START_RE.test(trimmed) || !current) {
      current = { lines: [trimmed] };
      units.push(current);
    } else {
      current.lines.push(trimmed);
    }
  });
  units.forEach((unit) => {
    unit.text = unit.lines.join(" ");
    unit.first = unit.lines[0];
  });
  return units;
}

function countWords(text) {
  const trimmed = String(text || "").trim();
  return trimmed ? trimmed.split(/\s+/).length : 0;
}

function getResponsiveGlanceIssues(liveLines) {
  const issues = [];
  const units = groupGlanceUnits(liveLines);
  const askUnits = units.filter((unit) => /\bASK:/.test(unit.text));
  const scanUnits = units.filter((unit) => /^(?:\d+[.)]\s+)?SCAN\b/i.test(unit.first));
  const revealUnits = units.filter((unit) => /^(?:\d+[.)]\s+)?REVEAL\b/i.test(unit.first));

  askUnits.forEach((unit, index) => {
    if (!/\bEXPECT:/i.test(unit.text)) {
      issues.push(`ASK beat ${index + 1} is missing EXPECT: inside the beat`);
    }
    if (!THINK_TIME_RE.test(unit.text)) {
      issues.push(`ASK beat ${index + 1} is missing explicit think time in seconds`);
    }
    if (!RESPONSE_ROUTINE_RE.test(unit.text)) {
      issues.push(`ASK beat ${index + 1} is missing one named response routine`);
    }
    if (/\b(?:hands?\s+up|volunteers?)\b/i.test(unit.text)) {
      issues.push(`ASK beat ${index + 1} defaults to volunteer hands instead of whole-class thinking`);
    }
  });

  scanUnits.forEach((unit, index) => {
    if (askUnits.length === 0) {
      issues.push(`SCAN beat ${index + 1} has no ASK beat to generate evidence`);
    }
    if (!/80%\+[^.\n]*->/i.test(unit.text)) {
      issues.push(`SCAN beat ${index + 1} is missing the 80%+ proceed branch`);
    }
    if (!/Less\s*->/i.test(unit.text)) {
      issues.push(`SCAN beat ${index + 1} is missing the Less -> pivot branch`);
    }
    if (!/re[- ]?(?:ask|check)/i.test(unit.text)) {
      issues.push(`SCAN beat ${index + 1} is missing a fresh re-ask or re-check`);
    }
  });

  revealUnits.forEach((unit, index) => {
    if (!/\bafter\b/i.test(unit.text)) {
      issues.push(`REVEAL beat ${index + 1} does not protect thinking with an 'after' condition`);
    }
  });

  return issues;
}

function getTeacherNotesSourceIssues(notes, opts) {
  if (notes == null || notes === "") return [];

  const o = opts || {};
  const raw = String(notes).replace(/\r\n?/g, "\n");
  const lines = raw.split("\n");
  const issues = [];
  const sanitized = sanitizeTeacherNotes(raw);
  const sections = parseNotesSections(raw);

  if (o.checkMarkdownHeaders !== false &&
      /\*\*(?:SAY|DO|PACING OVERVIEW|CFU CHECKPOINT|TEACHER NOTES|ENABLING & EXTENDING|MISCONCEPTIONS|SENSITIVITY ADVISORY|WATCH FOR|SOURCES):\*\*/i.test(raw)) {
    issues.push("markdown note headers are not allowed");
  }

  if (o.checkUnicodeBullets !== false && /[\u2022\u2023\u25E6\u2043\u2219\u25CF\u25AA]/.test(raw)) {
    issues.push("unicode bullets are not allowed");
  }

  if (o.checkSmartPunctuation !== false &&
      /[\u2018\u2019\u201A\u201B\u2032\u201C\u201D\u201E\u201F\u2033\u2010\u2011\u2012\u2013\u2014\u2015\u2212\u2026]/.test(raw)) {
    issues.push("smart punctuation must be authored in ASCII");
  }

  if (o.checkAscii !== false && /[^\x09\x0A\x0D\x20-\x7E]/.test(raw)) {
    issues.push("non-ASCII/control characters are not allowed in note source");
  }

  const maxLines = o.maxLines || 40;
  const maxChars = o.maxChars || 2600;
  if (lines.length > maxLines) {
    issues.push(`note block exceeds ${maxLines} lines`);
  }
  if (sanitized && sanitized.length > maxChars) {
    issues.push(`note block exceeds ${maxChars} characters after sanitizing`);
  }

  if (o.checkSectionStructure && isGlanceFormatNotes(raw)) {
    // Glance Format (v12.3): live zone above a "---" divider, prep zone below.
    // Budgets are RENDERED budgets: logical units (ANSWER, beats, TRAP,
    // STRETCH/HELP, CARE) stay capped at 8, but the wall-of-text failure mode
    // is prevented by word caps - a live zone over ~120 words or any physical
    // line over ~16 words wraps into an unglanceable block on an iPad.
    const allLines = sanitized ? sanitized.split("\n") : [];
    const dividerIndex = allLines.findIndex((line) => line.trim() === "---");
    const liveLines = (dividerIndex === -1 ? allLines : allLines.slice(0, dividerIndex))
      .filter((line) => line.trim());
    const prepLines = (dividerIndex === -1 ? [] : allLines.slice(dividerIndex + 1))
      .filter((line) => line.trim() && line.trim() !== "---");

    const maxLiveZoneUnits = o.maxLiveZoneUnits || o.maxLiveZoneLines || 8;
    const maxLiveZonePhysicalLines = o.maxLiveZonePhysicalLines || 18;
    const maxLiveZoneWords = o.maxLiveZoneWords || 120;
    const maxLineWords = o.maxLineWords || 16;
    const maxPrepZoneLines = o.maxPrepZoneLines || 3;

    if (dividerIndex === -1 && liveLines.length > 2) {
      issues.push("glance notes missing --- divider (only 1-2 line non-teaching notes may omit it)");
    }

    const units = groupGlanceUnits(liveLines);
    // STRETCH and HELP sit on separate physical lines but count as one unit.
    const unitCount = units.length - units.filter((unit, idx) =>
      /^HELP:/.test(unit.first) && idx > 0 && /^STRETCH:/.test(units[idx - 1].first)
    ).length;
    if (unitCount > maxLiveZoneUnits) {
      issues.push(`glance live zone exceeds ${maxLiveZoneUnits} logical units (${unitCount})`);
    }
    if (liveLines.length > maxLiveZonePhysicalLines) {
      issues.push(`glance live zone exceeds ${maxLiveZonePhysicalLines} physical lines (${liveLines.length})`);
    }
    const liveWords = countWords(liveLines.join(" "));
    if (liveWords > maxLiveZoneWords) {
      issues.push(`glance live zone exceeds ${maxLiveZoneWords} words (${liveWords}) - cut rationale, split the slide, or move detail to the prep zone`);
    }
    liveLines.forEach((line) => {
      const words = countWords(line);
      if (words > maxLineWords) {
        issues.push(`glance live line exceeds ${maxLineWords} words (${words}): "${line.trim().slice(0, 48)}..." - break it into short indented lines`);
      }
    });
    if (prepLines.length > maxPrepZoneLines) {
      issues.push(`glance prep zone exceeds ${maxPrepZoneLines} lines (${prepLines.length})`);
    }
    if (liveLines.some((line) => /\bASK:/.test(line) || /EXPECT:/.test(line)) &&
        liveLines.length > 0 && !/^(?:ANSWER:|REVEALED[.:])/.test(liveLines[0].trim())) {
      issues.push("glance notes with an ASK must open with an ANSWER: line");
    }
    if (o.checkResponsiveGlance !== false) {
      issues.push(...getResponsiveGlanceIssues(liveLines));
    }
    return issues;
  }

  if (o.checkSectionStructure) {
    const sectionNames = new Set(sections.map((section) => section.name));
    ["SAY:", "DO:"].forEach((required) => {
      if (!sectionNames.has(required)) {
        issues.push(`missing required ${required} section`);
      }
    });

    const maxSayBullets = o.maxSayBullets || 16;
    const maxDoBullets = o.maxDoBullets || 8;
    const maxWatchForBullets = o.maxWatchForBullets || 5;
    const maxTeacherNotesLines = o.maxTeacherNotesLines || 8;
    const maxTeacherNotesChars = o.maxTeacherNotesChars || 1400;

    sections.forEach((section) => {
      const nonBlank = section.lines.filter((line) => line.trim());
      const bulletCount = nonBlank.filter((line) => /^-\s+/.test(line.trim())).length;
      const contentCount = bulletCount || nonBlank.length;
      const chars = nonBlank.join(" ").length;

      if (section.name === "SAY:" && contentCount > maxSayBullets) {
        issues.push(`SAY section exceeds ${maxSayBullets} bullets/lines`);
      }
      if (section.name === "DO:" && contentCount > maxDoBullets) {
        issues.push(`DO section exceeds ${maxDoBullets} bullets/lines`);
      }
      if (section.name === "WATCH FOR:" && contentCount > maxWatchForBullets) {
        issues.push(`WATCH FOR section exceeds ${maxWatchForBullets} bullets/lines`);
      }
      if (section.name === "TEACHER NOTES:") {
        if (nonBlank.length > maxTeacherNotesLines) {
          issues.push(`TEACHER NOTES section exceeds ${maxTeacherNotesLines} lines`);
        }
        if (chars > maxTeacherNotesChars) {
          issues.push(`TEACHER NOTES section exceeds ${maxTeacherNotesChars} characters`);
        }
      }
    });
  }

  return issues;
}

function getSlideNotesText(slide) {
  if (!slide || !Array.isArray(slide._slideObjects)) return "";

  return slide._slideObjects
    .filter((obj) => obj && obj._type === "notes" && Array.isArray(obj.text))
    .flatMap((obj) => obj.text)
    .map((entry) => entry && entry.text ? String(entry.text).trim() : "")
    .filter(Boolean)
    .join("\n\n");
}

// Recognised Glance Format anchors get a real bold run in the notes XML so
// they survive into PowerPoint and Google Slides presenter view (mirrors the
// OG builder, which has always bolded its note labels). Case-sensitive on
// purpose: "show" inside SAY speech must not match SHOW.
const NOTE_BOLD_TOKEN_RE = /(\d+[.)](?=\s)|ANSWER:|REVEALED[.:]|SAY:|ASK:|EXPECT:|ACCEPT:|SCAN\b|TRAP:|Fix:|Cue:|STRETCH:|HELP:|CARE:|REVEAL\b|MODEL\b|POINT\b|SHOW\b|DRAW\b|BUILD\b|COVER\b|COLLECT\b|CIRCULATE\b|TIME:|SOURCES:|WHY:)/g;

function noteLineToRunsXml(line) {
  // split() with a single capture group alternates plain (even index) and
  // captured (odd index) segments; empty plain segments keep their slot.
  const parts = String(line).split(NOTE_BOLD_TOKEN_RE);
  return parts
    .map((part, index) => {
      if (!part) return "";
      const isBold = index % 2 === 1;
      const rPr = isBold ? '<a:rPr lang="en-US" b="1" dirty="0"/>' : '<a:rPr lang="en-US" dirty="0"/>';
      return `<a:r>${rPr}<a:t>${escapeXml(part)}</a:t></a:r>`;
    })
    .join("");
}

function buildNotesParagraphsXml(notes) {
  const sanitized = sanitizeTeacherNotes(notes || "");
  const lines = sanitized ? sanitized.split("\n") : [""];
  const paragraphs = [];

  lines.forEach((line) => {
    if (!line) {
      paragraphs.push('<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>');
      return;
    }

    paragraphs.push(
      `<a:p>${noteLineToRunsXml(line)}<a:endParaRPr lang="en-US" dirty="0"/></a:p>`
    );
  });

  if (paragraphs.length === 0) {
    paragraphs.push('<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>');
  }

  return `<a:bodyPr/><a:lstStyle/>${paragraphs.join("")}`;
}

function rewriteNotesSlideXml(noteXml, notes) {
  const replacementBody = buildNotesParagraphsXml(notes);
  const pattern = /(<p:cNvPr id="3" name="Notes Placeholder 2"\/>[\s\S]*?<p:txBody>)[\s\S]*?(<\/p:txBody>)/;

  if (!pattern.test(noteXml)) {
    throw new Error("Unable to locate notes placeholder body while rewriting notes XML.");
  }

  // Function-form replacement: avoids `$N` inside replacementBody being
  // parsed as a backreference (corrupts notes that contain "$25" etc.).
  return noteXml.replace(pattern, (_match, p1, p2) => `${p1}${replacementBody}${p2}`);
}

async function rewriteSpeakerNotesInFile(pptxPath, slides) {
  if (!pptxPath || !Array.isArray(slides) || slides.length === 0) return;

  const input = await fs.promises.readFile(pptxPath);
  const zip = await JSZip.loadAsync(input);

  for (let index = 0; index < slides.length; index += 1) {
    const xmlPath = `ppt/notesSlides/notesSlide${index + 1}.xml`;
    const file = zip.file(xmlPath);
    if (!file) continue;

    const noteXml = await file.async("string");
    const noteText = getSlideNotesText(slides[index]);
    zip.file(xmlPath, rewriteNotesSlideXml(noteXml, noteText));
  }

  const output = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  await fs.promises.writeFile(pptxPath, output);
}

function cleanTargetText(value) {
  return toAscii(stripMarkdown(String(value || ""))).replace(/\s+/g, " ").trim();
}

function ensureSentence(value) {
  if (!value) return value;
  return /[.!?]$/.test(value) ? value : `${value}.`;
}

function ensureICan(value) {
  if (!value) return value;
  return /^i can\b/i.test(value) ? value : `I can ${value.charAt(0).toLowerCase()}${value.slice(1)}`;
}

function toArray(values) {
  if (values == null) return [];
  return Array.isArray(values) ? values : [values];
}

function stripSourcesSection(notes) {
  const sanitized = sanitizeTeacherNotes(notes || "");
  if (!sanitized) return "";

  const lines = sanitized.split("\n");
  const nextLines = [];
  let inSources = false;

  lines.forEach((line) => {
    const trimmed = line.trim();
    if (trimmed === "SOURCES:") {
      inSources = true;
      return;
    }

    if (inSources && HEADER_PATTERN.test(trimmed.replace(/:$/, ""))) {
      inSources = false;
    }

    if (!inSources) {
      nextLines.push(line);
    }
  });

  while (nextLines.length > 0 && nextLines[nextLines.length - 1] === "") {
    nextLines.pop();
  }

  return nextLines.join("\n");
}

function appendSourcesToNotes(notes, sources) {
  const cleanedNotes = stripSourcesSection(notes);
  const existingSections = parseNotesSections(notes);
  const existingSources = existingSections
    .filter((section) => section.name === "SOURCES:")
    .flatMap((section) => section.lines)
    .map((line) => line.replace(/^\s*-\s*/, "").trim())
    .filter(Boolean);
  const nextSources = toArray(sources)
    .map((source) => cleanTargetText(source))
    .filter(Boolean);
  const combined = [...new Set([...existingSources, ...nextSources])];
  if (combined.length === 0) {
    return sanitizeTeacherNotes(cleanedNotes);
  }

  const sourceBlock = ["SOURCES:", ...combined.map((source) => `- ${source}`)].join("\n");
  const joined = cleanedNotes ? `${cleanedNotes}\n\n${sourceBlock}` : sourceBlock;
  return sanitizeTeacherNotes(joined);
}

function normalizeLessonTargets(liItems, scItems) {
  const li = toArray(liItems).map(cleanTargetText).filter(Boolean);
  const sc = toArray(scItems).map(cleanTargetText).filter(Boolean);
  const warnings = [];

  if (li.length !== 1) {
    warnings.push(`expected exactly 1 Learning Intention, received ${li.length}`);
  }
  if (sc.length !== 3) {
    warnings.push(`expected exactly 3 Success Criteria, received ${sc.length}`);
  }

  return {
    liItems: li.slice(0, 1).map(ensureSentence),
    scItems: sc.slice(0, 3).map(ensureICan),
    warnings,
  };
}

// Targeted slide-face text hygiene. Unlike notes (full ASCII), slide faces
// legitimately use maths glyphs (x, divided-by, ticks), so only the banned
// characters are converted: dash family -> "-", smart quotes -> straight,
// ellipsis -> "...", double hyphens -> single.
const SLIDE_TEXT_REPLACEMENTS = [
  [/[‘’‚‛′]/g, "'"],
  [/[“”„‟″]/g, '"'],
  [/…/g, "..."],
  [/\s*[‐‑‒–—―]\s*/g, " - "],
  [/(^|[^-])--(?!-)([^-]|$)/g, "$1 - $2"],
  [/ {2,}- {2,}/g, " - "],
];

function sanitizeSlideText(value) {
  if (typeof value !== "string" || value === "") return value;
  let next = value;
  for (const [pattern, replacement] of SLIDE_TEXT_REPLACEMENTS) {
    next = next.replace(pattern, replacement);
  }
  return next;
}

let slideTextPatchInstalled = false;

function installSlideTextPatch(PptxGenJS) {
  if (slideTextPatchInstalled || typeof PptxGenJS !== "function") return;

  const probe = new PptxGenJS();
  const slide = probe.addSlide();
  const proto = Object.getPrototypeOf(slide);
  const originalAddText = proto && proto.addText;

  if (typeof originalAddText !== "function" || originalAddText.__slideTextPatched) {
    slideTextPatchInstalled = true;
    return;
  }

  function patchedAddText(text, options) {
    let cleanText = text;
    if (typeof text === "string") {
      cleanText = sanitizeSlideText(text);
    } else if (Array.isArray(text)) {
      cleanText = text.map((run) => {
        if (run && typeof run === "object" && typeof run.text === "string") {
          return { ...run, text: sanitizeSlideText(run.text) };
        }
        return run;
      });
    }
    return originalAddText.call(this, cleanText, options);
  }

  patchedAddText.__slideTextPatched = true;
  proto.addText = patchedAddText;
  slideTextPatchInstalled = true;
}

function installNotesPatch(PptxGenJS) {
  if (notesPatchInstalled || typeof PptxGenJS !== "function") return;

  const probe = new PptxGenJS();
  const slide = probe.addSlide();
  const proto = Object.getPrototypeOf(slide);
  const originalAddNotes = proto && proto.addNotes;
  const presProto = PptxGenJS.prototype;
  const originalWriteFile = presProto && presProto.writeFile;

  if (typeof originalAddNotes !== "function" || typeof originalWriteFile !== "function" ||
      originalAddNotes.__teacherNotesPatched || originalWriteFile.__teacherNotesPatched) {
    notesPatchInstalled = true;
    return;
  }

  function patchedAddNotes(notes) {
    return originalAddNotes.call(this, sanitizeTeacherNotes(notes));
  }

  patchedAddNotes.__teacherNotesPatched = true;
  proto.addNotes = patchedAddNotes;

  // The single post-write pass. PptxGenJS has finished the file at this point,
  // so anything the library cannot express itself (structured notes XML, click
  // build timing trees) is written here.
  async function patchedWriteFile(props) {
    const filePath = await originalWriteFile.call(this, props);
    if (typeof filePath === "string") {
      const slides = this.slides || [];
      await rewriteSpeakerNotesInFile(filePath, slides);
      // Lazy require: animations.js needs nothing from this module, so there
      // is no cycle, but keeping it lazy means decks that never call
      // clickBuild() do not pay for loading it.
      const { injectClickBuildsInFile } = require("./animations");
      await injectClickBuildsInFile(filePath, slides);
    }
    return filePath;
  }

  patchedWriteFile.__teacherNotesPatched = true;
  presProto.writeFile = patchedWriteFile;
  notesPatchInstalled = true;
}

module.exports = {
  NOTE_SECTION_HEADERS,
  isGlanceFormatNotes,
  sanitizeSlideText,
  installSlideTextPatch,
  sanitizeTeacherNotes,
  parseNotesSections,
  getTeacherNotesSourceIssues,
  getResponsiveGlanceIssues,
  getSlideNotesText,
  buildNotesParagraphsXml,
  rewriteNotesSlideXml,
  rewriteSpeakerNotesInFile,
  normalizeLessonTargets,
  appendSourcesToNotes,
  installNotesPatch,
};
