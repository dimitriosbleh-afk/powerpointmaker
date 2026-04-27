"use strict";

const { sanitizeTeacherNotes, appendSourcesToNotes } = require("./notes");

/**
 * Compose teacher notes in the megaprompt §45 mandated section order.
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

module.exports = { composeNotes };
