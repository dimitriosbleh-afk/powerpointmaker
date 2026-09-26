"use strict";

/**
 * Taught log: what this machine's lessons have taught, in teaching order.
 *
 * Every lesson spec that passes build_and_check.js is recorded in
 * records/taught_<yearLevel>_<subject>.json (gitignored: each teacher keeps
 * their own log on their own computer). Entries are ordered by teaching
 * position (year, term, week, session), never by build date, because a whole
 * term is often generated in one sitting.
 *
 * Daily Review reads it (megaprompt 83): every dailyReview slide names the
 * logged lesson it retrieves, and the validator checks that lesson came
 * earlier and that the review is spaced.
 */

const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..", "..");
// TAUGHT_LOG_DIR redirects the log, so tests and trial builds never touch a real one.
const DEFAULT_DIR = process.env.TAUGHT_LOG_DIR ? path.resolve(process.env.TAUGHT_LOG_DIR) : path.join(ROOT, "records");

// Sources a dailyReview slide may name instead of a log key.
const TEACHER_SOURCE = "teacher";       // the teacher named the focus in the request
const BEFORE_LOG_SOURCE = "before log"; // taught before this log began

function resolveYear(L) {
  return Number.isInteger(L.year) ? L.year : new Date().getFullYear();
}

/** Weeks since an arbitrary origin; terms are treated as 12 weeks so gaps across a holiday read as spaced. */
function weekIndex(year, term, week) {
  return (year * 4 + (term - 1)) * 12 + week;
}

function positionOf(L) {
  const year = resolveYear(L);
  return {
    year, term: L.term, week: L.week, session: L.session || 1,
    weekIndex: weekIndex(year, L.term, L.week),
    order: weekIndex(year, L.term, L.week) * 10 + (L.session || 1),
    key: `${year}-T${L.term}-W${L.week}-S${L.session || 1}`,
  };
}

function logPath(yearLevel, subject, dir) {
  return path.join(dir || DEFAULT_DIR, `taught_${yearLevel}_${subject}.json`);
}

function readLog(yearLevel, subject, dir) {
  const p = logPath(yearLevel, subject, dir);
  if (!fs.existsSync(p)) return [];
  const data = JSON.parse(fs.readFileSync(p, "utf8"));
  return Array.isArray(data.entries) ? data.entries : [];
}

function writeLog(yearLevel, subject, entries, dir) {
  const p = logPath(yearLevel, subject, dir);
  fs.mkdirSync(path.dirname(p), { recursive: true });
  const sorted = entries.slice().sort((a, b) => a.order - b.order);
  fs.writeFileSync(p, JSON.stringify({ yearLevel, subject, entries: sorted }, null, 2) + "\n");
  return p;
}

function toList(v) { return v == null ? [] : (Array.isArray(v) ? v : [v]); }

/**
 * Items worth retrieving later, best first: the exit ticket (the core
 * criterion on a new item), then worksheet items, then practice round items.
 */
function reviewItemsFrom(spec) {
  const exit = [];
  const sheet = [];
  const rounds = [];
  (spec.slides || []).forEach((s) => {
    if (!s) return;
    if (s.kind === "exitTicket") {
      exit.push({
        from: "exit ticket",
        prompt: toList(s.questions).join(" "),
        visual: s.visual,
        answer: s.notes && typeof s.notes === "object" ? s.notes.answer : undefined,
      });
    }
    if (s.kind === "practice") {
      toList(s.items).forEach((it) => rounds.push({
        from: `practice: ${s.title}`,
        prompt: it.extract || it.text || s.title,
        visual: it.visual,
        answer: it.answer,
      }));
    }
  });
  (spec.resources || []).forEach((r) => {
    if (r && r.kind === "worksheet") {
      toList(r.items).forEach((it) => sheet.push({ from: `worksheet: ${r.label}`, prompt: it.prompt, visual: it.visual, answer: it.answer }));
    }
  });
  return exit.concat(sheet, rounds);
}

function entryFromSpec(spec, specPath) {
  const L = spec.lesson || {};
  const pos = positionOf(L);
  const li = (spec.slides || []).find((s) => s && s.kind === "li") || {};
  return {
    key: pos.key,
    order: pos.order,
    weekIndex: pos.weekIndex,
    year: pos.year, term: pos.term, week: pos.week, session: pos.session,
    title: L.title,
    learningIntention: li.learningIntention,
    successCriteria: li.successCriteria,
    keyWords: (spec.slides || []).filter((s) => s && s.kind === "keyWord").map((s) => s.word),
    reviews: (spec.slides || []).filter((s) => s && s.kind === "dailyReview").map((s) => s.from).filter(Boolean),
    items: reviewItemsFrom(spec),
    spec: specPath ? path.relative(ROOT, path.resolve(specPath)) : undefined,
    builtAt: new Date().toISOString(),
  };
}

/** Add or replace this lesson's entry (a rebuild replaces, never duplicates). Returns the log path. */
function recordLesson(spec, specPath, dir) {
  const L = spec.lesson || {};
  const entry = entryFromSpec(spec, specPath);
  const entries = readLog(L.yearLevel, L.subject, dir).filter((e) => e.key !== entry.key);
  entries.push(entry);
  return writeLog(L.yearLevel, L.subject, entries, dir);
}

function removeEntry(yearLevel, subject, key, dir) {
  const entries = readLog(yearLevel, subject, dir);
  const kept = entries.filter((e) => e.key !== key);
  if (kept.length === entries.length) return false;
  writeLog(yearLevel, subject, kept, dir);
  return true;
}

/** Specs that must never write to a teacher's log: the golden exemplars. */
function isLoggable(specPath) {
  return !/^exemplar_/.test(path.basename(String(specPath || "")));
}

module.exports = {
  TEACHER_SOURCE, BEFORE_LOG_SOURCE, DEFAULT_DIR,
  positionOf, weekIndex, logPath, readLog, writeLog, entryFromSpec, recordLesson, removeEntry, isLoggable,
};
