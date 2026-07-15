"use strict";

const {
  sanitizeTeacherNotes,
  appendSourcesToNotes,
  getTeacherNotesSourceIssues,
} = require("./notes");

/**
 * LEGACY (pre-v11.0 sectioned format). New lessons use the Glance Format via
 * composeGlanceNotes() below. composeNotes() is kept so existing build
 * scripts still rebuild byte-identically.
 *
 * Compose teacher notes in the megaprompt pre-v11 mandated section order.
 *
 * Section order (omitting any sections that are empty):
 *   SAY -> DO -> CFU CHECKPOINT -> TEACHER NOTES -> ENABLING & EXTENDING
 *   -> MISCONCEPTIONS -> SENSITIVITY ADVISORY -> WATCH FOR
 *   then optional SOURCES, then optional trailing single-line tag.
 *
 * The output is plain ASCII (PowerPoint-iPad-safe), uses hyphen bullets,
 * and is run through sanitizeTeacherNotes() before return so it picks up
 * the same hardening as direct addNotes() calls.
 *
 * Each section accepts:
 *  - a string (used as-is, split on newlines, bulleted if it isn't already)
 *  - an array of strings (each becomes a `- ...` bullet)
 *  - an object with .text and .bullets for mixed prose + bulleted detail
 *
 * cfu accepts a structured object: { technique, script, scanFor, proceed, pivot }
 * and renders the §47 template.
 *
 * @param {object} input
 * @param {string|string[]|object}            [input.say]
 * @param {string|string[]|object}            [input.do]
 * @param {string|string[]|object}            [input.cfu]    structured CFU or free-text
 * @param {string|string[]|object}            [input.teacherNotes]
 * @param {string|string[]|object}            [input.enabling]
 * @param {string|string[]|object}            [input.misconceptions]
 * @param {string|string[]|object}            [input.sensitivity]
 * @param {string|string[]|object}            [input.watchFor]
 * @param {string|string[]}                   [input.sources]
 * @param {string}                            [input.tag]
 * @param {object}                            [opts]
 * @param {boolean}                           [opts.requireSay=true]   warn if SAY missing
 * @param {boolean}                           [opts.requireDo=true]    warn if DO missing
 * @returns {string} composed, sanitized notes block
 */
function composeNotes(input, opts) {
  const i = input || {};
  const o = opts || {};
  const requireSay = o.requireSay !== false;
  const requireDo = o.requireDo !== false;

  const blocks = [];

  function pushSection(header, body) {
    if (body == null || body === "") return;
    blocks.push(`${header}:`);
    blocks.push(body);
    blocks.push("");
  }

  function renderItems(value) {
    if (value == null) return "";
    if (Array.isArray(value)) {
      const lines = value
        .map((entry) => String(entry == null ? "" : entry).trim())
        .filter(Boolean)
        .map((entry) => (/^-\s+/.test(entry) ? entry : `- ${entry}`));
      return lines.join("\n");
    }
    if (typeof value === "object") {
      const out = [];
      if (value.text) out.push(String(value.text).trim());
      if (Array.isArray(value.bullets)) {
        const bulletLines = value.bullets
          .map((entry) => String(entry == null ? "" : entry).trim())
          .filter(Boolean)
          .map((entry) => (/^-\s+/.test(entry) ? entry : `- ${entry}`));
        if (bulletLines.length) out.push(bulletLines.join("\n"));
      }
      return out.join("\n");
    }
    return String(value).trim();
  }

  function renderCfu(value) {
    if (value == null || value === "") return "";
    if (typeof value === "string" || Array.isArray(value)) {
      return renderItems(value);
    }
    const cfu = value || {};
    const parts = [];
    if (cfu.technique) parts.push(`Technique: ${String(cfu.technique).trim()}`);
    if (cfu.script) {
      parts.push("Script:");
      parts.push(renderItems(cfu.script));
    }
    if (cfu.scanFor) parts.push(`- Scan for: ${String(cfu.scanFor).trim()}`);
    if (cfu.proceed) {
      parts.push("PROCEED:");
      parts.push(renderItems(cfu.proceed));
    }
    if (cfu.pivot) {
      const pivot = cfu.pivot;
      parts.push("PIVOT:");
      if (typeof pivot === "string" || Array.isArray(pivot)) {
        parts.push(renderItems(pivot));
      } else {
        if (pivot.misconception) parts.push(`- Most likely misconception: ${String(pivot.misconception).trim()}`);
        if (pivot.reteach) parts.push(`- Reteach: ${String(pivot.reteach).trim()}`);
        if (pivot.recheck) parts.push(`- Re-check: ${String(pivot.recheck).trim()}`);
      }
    }
    return parts.filter(Boolean).join("\n");
  }

  const saySection = renderItems(i.say);
  const doSection = renderItems(i.do);
  const cfuSection = renderCfu(i.cfu);
  const teacherSection = renderItems(i.teacherNotes);
  const enablingSection = renderItems(i.enabling);
  const misconceptionsSection = renderItems(i.misconceptions);
  const sensitivitySection = renderItems(i.sensitivity);
  const watchForSection = renderItems(i.watchFor);

  if (requireSay && !saySection) {
    console.warn("[composeNotes] SAY section is empty; megaprompt §45 requires SAY on every slide.");
  }
  if (requireDo && !doSection) {
    console.warn("[composeNotes] DO section is empty; megaprompt §45 requires DO on every slide.");
  }

  pushSection("SAY", saySection);
  pushSection("DO", doSection);
  pushSection("CFU CHECKPOINT", cfuSection);
  pushSection("TEACHER NOTES", teacherSection);
  pushSection("ENABLING & EXTENDING", enablingSection);
  pushSection("MISCONCEPTIONS", misconceptionsSection);
  pushSection("SENSITIVITY ADVISORY", sensitivitySection);
  pushSection("WATCH FOR", watchForSection);

  if (i.tag) {
    blocks.push(String(i.tag).trim());
  }

  while (blocks.length && blocks[blocks.length - 1] === "") blocks.pop();

  const joined = blocks.join("\n");
  // appendSourcesToNotes already runs sanitizeTeacherNotes on its output,
  // so skip the second pass when sources are appended.
  if (i.sources) return appendSourcesToNotes(joined, i.sources);
  return sanitizeTeacherNotes(joined);
}

const GLANCE_VALIDATION_OPTS = {
  checkSectionStructure: true,
  checkResponsiveGlance: true,
  maxLines: 34,
  maxChars: 1800,
  maxLiveZoneUnits: 8,
  maxLiveZonePhysicalLines: 18,
  maxLiveZoneWords: 120,
  maxLineWords: 16,
  maxPrepZoneLines: 3,
};

const CONTINUATION_INDENT = "   ";

/**
 * Compose Glance Format notes (megaprompt v12.3 sections 45-47).
 *
 * Live zone: ANSWER, numbered beats, TRAP, STRETCH/HELP, CARE - max 8 logical
 * units, one blank line between units, and RENDERED budgets: 120 words for
 * the whole live zone, 16 words per physical line. A beat may be an array of
 * short lines; the first line is numbered and the rest render indented, so
 * speech (SAY), think time, cue script and EXPECT each sit on their own line.
 * Prep zone (max 3 lines) below a "---" divider: purpose line(s) + tag,
 * then optional SOURCES.
 *
 * @param {object} input
 * @param {string}                    [input.answer]  ANSWER: content (omit when the slide asks nothing)
 * @param {Array<string|string[]>}    [input.beats]   2-5 beats in teaching order; numbering is applied.
 *                                                    A string[] beat renders as numbered first line +
 *                                                    indented continuation lines.
 * @param {string|string[]}           [input.trap]    "error. Fix: move, student redoes" (TRAP: prefixed;
 *                                                    an array renders as continuation lines)
 * @param {string}                    [input.stretch] STRETCH content (core teaching slides)
 * @param {string}                    [input.help]    HELP content (core teaching slides)
 * @param {string}                    [input.care]    CARE: content (sensitive content only)
 * @param {string|string[]}           [input.prep]    prep-zone purpose line(s)
 * @param {string|string[]}           [input.sources] prep-zone SOURCES entries
 * @param {string}                    [input.tag]     "[Stage | VTLM element | SC | HITS n]"
 * @returns {string} composed, sanitized notes block
 */
function composeGlanceNotes(input, opts) {
  const i = input || {};
  const o = opts || {};
  const list = (value) => (value == null ? [] : Array.isArray(value) ? value : [value])
    .map((entry) => String(entry).trim())
    .filter(Boolean);
  const toBeatLines = (value) => (Array.isArray(value) ? value : [value])
    .map((entry) => String(entry).trim())
    .filter(Boolean);

  // Each block is one logical unit; blocks join with a blank line between
  // them so the live zone reads as separated moments, not a wall.
  const blocks = [];
  if (i.answer) blocks.push([`ANSWER: ${String(i.answer).trim()}`]);

  (i.beats == null ? [] : i.beats).forEach((beat, idx) => {
    const lines = toBeatLines(beat);
    if (!lines.length) return;
    const first = /^\d+[.)]\s/.test(lines[0]) ? lines[0] : `${idx + 1}. ${lines[0]}`;
    blocks.push([first, ...lines.slice(1).map((line) => `${CONTINUATION_INDENT}${line}`)]);
  });

  const trapLines = toBeatLines(i.trap == null ? [] : i.trap);
  if (trapLines.length) {
    const first = /^TRAP:/i.test(trapLines[0]) ? trapLines[0] : `TRAP: ${trapLines[0]}`;
    blocks.push([first, ...trapLines.slice(1).map((line) => `${CONTINUATION_INDENT}${line}`)]);
  }

  const stretchHelp = [];
  if (i.stretch) stretchHelp.push(`STRETCH: ${String(i.stretch).trim()}`);
  if (i.help) stretchHelp.push(`HELP: ${String(i.help).trim()}`);
  if (stretchHelp.length) blocks.push(stretchHelp);

  if (i.care) {
    const care = String(i.care).trim();
    blocks.push([/^CARE:/i.test(care) ? care : `CARE: ${care}`]);
  }

  if (blocks.length > 8) {
    console.warn(`[composeGlanceNotes] live zone has ${blocks.length} logical units; megaprompt section 46 caps it at 8.`);
  }

  const lines = [];
  blocks.forEach((block, idx) => {
    if (idx > 0) lines.push("");
    lines.push(...block);
  });

  const prep = list(i.prep);
  const tag = i.tag ? String(i.tag).trim() : "";
  const sources = list(i.sources);
  if (prep.length || tag || sources.length) {
    lines.push("---");
    prep.forEach((line, idx) => {
      const isLast = idx === prep.length - 1;
      lines.push(isLast && tag ? `${line} ${tag}` : line);
    });
    if (!prep.length && tag) lines.push(tag);
    if (sources.length) lines.push(`SOURCES: ${sources.join("; ")}`);
  }

  const composed = sanitizeTeacherNotes(lines.join("\n"));
  if (o.validate !== false) {
    const issues = getTeacherNotesSourceIssues(composed, GLANCE_VALIDATION_OPTS);
    if (issues.length) {
      throw new Error(`[composeGlanceNotes] ${issues.join("; ")}`);
    }
  }
  return composed;
}

/**
 * Compose reveal-slide notes (megaprompt v12.3 section 47). A reveal slide
 * must NEVER duplicate its base slide's notes: when the teacher clicks to the
 * answer, the notes advance with the slide. The live zone is the post-reveal
 * script only - the answer restated, the tick-and-fix move, and at most one
 * follow-up.
 *
 * @param {object} input
 * @param {string}                 input.answer   what is now on screen, student voice
 * @param {Array<string|string[]>} [input.beats]  1-3 post-reveal beats (tick and fix,
 *                                                cold-call follow-up, transition)
 * @param {string|string[]}        [input.prep]   prep-zone line(s); defaults to "Reveal slide."
 * @param {string}                 [input.tag]    optional prep-zone tag
 * @returns {string} composed, sanitized notes block
 */
function composeRevealNotes(input, opts) {
  const i = input || {};
  const o = opts || {};
  if (!i.answer) {
    throw new Error("[composeRevealNotes] answer is required - the reveal slide restates what is now on screen");
  }
  const toBeatLines = (value) => (Array.isArray(value) ? value : [value])
    .map((entry) => String(entry).trim())
    .filter(Boolean);

  const lines = [`REVEALED: ${String(i.answer).trim()}`];
  (i.beats == null ? [] : i.beats).forEach((beat, idx) => {
    const beatLines = toBeatLines(beat);
    if (!beatLines.length) return;
    lines.push("");
    const first = /^\d+[.)]\s/.test(beatLines[0]) ? beatLines[0] : `${idx + 1}. ${beatLines[0]}`;
    lines.push(first, ...beatLines.slice(1).map((line) => `${CONTINUATION_INDENT}${line}`));
  });

  lines.push("---");
  const prep = (i.prep == null || (Array.isArray(i.prep) && !i.prep.length)) ? ["Reveal slide."] :
    (Array.isArray(i.prep) ? i.prep : [i.prep]).map((entry) => String(entry).trim()).filter(Boolean);
  const tag = i.tag ? String(i.tag).trim() : "";
  prep.forEach((line, idx) => {
    const isLast = idx === prep.length - 1;
    lines.push(isLast && tag ? `${line} ${tag}` : line);
  });

  const composed = sanitizeTeacherNotes(lines.join("\n"));
  if (o.validate !== false) {
    const issues = getTeacherNotesSourceIssues(composed, GLANCE_VALIDATION_OPTS);
    if (issues.length) {
      throw new Error(`[composeRevealNotes] ${issues.join("; ")}`);
    }
  }
  return composed;
}

module.exports = { composeNotes, composeGlanceNotes, composeRevealNotes };
