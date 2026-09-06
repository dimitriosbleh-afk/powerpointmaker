"use strict";

/**
 * Lint the teacher notes in a lesson spec BEFORE building.
 *
 *   node scripts/check_spec_notes.js builds/exemplar_foundation_numeracy_making_10.json
 *
 * For every slide with Glance-object notes it reports the live-zone word
 * count (budget 120), any physical line over 16 words (which the build would
 * auto-wrap with an ADVISORY), and any composer error (missing routine cue,
 * missing think time, too many units). Fix everything it lists and the build
 * will produce zero note advisories.
 */

const path = require("path");
const { loadSpec } = require("../themes/lesson/buildLesson");
const { composeGlanceNotes } = require("../themes/core/composeNotes");

const specPath = process.argv[2];
if (!specPath) {
  console.error("Usage: node scripts/check_spec_notes.js <spec.json>");
  process.exit(2);
}

const spec = loadSpec(path.resolve(specPath));
let problems = 0;

(spec.slides || []).forEach((slide, i) => {
  const label = `slide ${i + 1} (${slide.kind}${slide.title ? `: ${slide.title}` : ""})`;
  const n = slide.notes;
  if (!n || typeof n === "string") return;

  // Word budget and line lengths from the author's own lines, before wrapping.
  const flat = [];
  if (n.answer) flat.push(`ANSWER: ${n.answer}`);
  (n.beats || []).forEach((b) => (Array.isArray(b) ? b : [b]).forEach((l) => flat.push(String(l))));
  (Array.isArray(n.trap) ? n.trap : (n.trap ? [n.trap] : [])).forEach((l) => flat.push(String(l)));
  if (n.stretch) flat.push(`STRETCH: ${n.stretch}`);
  if (n.help) flat.push(`HELP: ${n.help}`);
  if (n.care) flat.push(`CARE: ${n.care}`);
  // Count the way the build gate counts: the composed live zone (numbers and
  // labels included), everything above the --- divider.
  let words = 0;
  try {
    const silent = console.log; console.log = () => {};
    const composed = composeGlanceNotes(n, { validate: false });
    console.log = silent;
    const live = composed.split("\n---")[0].trim();
    words = live ? live.split(/\s+/).length : 0;
  } catch (err) { /* reported below */ }
  const longLines = flat.filter((l) => l.trim().split(/\s+/).filter(Boolean).length > 16);

  const issues = [];
  if (words > 120) issues.push(`live zone ${words} words (budget 120): cut a beat or move detail to prep`);
  longLines.forEach((l) => issues.push(`line over 16 words, split it into an array of short lines: "${l.slice(0, 60)}..."`));
  const origLog = console.log;
  console.log = () => {};   // the composer's wrap advisories duplicate the long-line report above
  try {
    composeGlanceNotes(n);
  } catch (err) {
    issues.push(String(err.message).replace(/^\[composeGlanceNotes\]\s*/, ""));
  } finally {
    console.log = origLog;
  }

  if (issues.length) {
    problems += issues.length;
    console.log(`\n${label}`);
    issues.forEach((m) => console.log(`  - ${m}`));
  } else {
    console.log(`ok   ${label}  (${words} words)`);
  }
});

console.log(problems ? `\n${problems} note problem(s). Fix before building.` : "\nAll notes within budget.");
process.exit(problems ? 1 : 0);
