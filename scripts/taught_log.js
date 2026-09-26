"use strict";

/**
 * Read this machine's taught log (megaprompt 83).
 *
 *   node scripts/taught_log.js builds/<spec>.json
 *       What lessons before this one taught, grouped by how long ago, with
 *       each lesson's review items and answers and how often it has already
 *       been reviewed. Use it to choose Daily Review items when the request
 *       does not name a focus: one from each older group where possible.
 *
 *   node scripts/taught_log.js --list <yearLevel> <subject>
 *   node scripts/taught_log.js --remove <yearLevel> <subject> <key>
 *   node scripts/taught_log.js --add <yearLevel> <subject> <key> "<what was taught>" ["<detail>"]
 *       Record learning taught outside this pipeline (a term planner, last
 *       term's decks) so Daily Review can reach back to it.
 *
 * The log lives in records/ (gitignored). build_and_check.js writes to it
 * whenever a lesson spec passes every gate.
 */

const path = require("path");
const { loadSpec } = require("../themes/lesson/buildLesson");
const { readLog, writeLog, removeEntry, positionOf, logPath } = require("../themes/lesson/taughtLog");

function line(e) {
  const sc2 = Array.isArray(e.successCriteria) ? e.successCriteria[1] : "";
  return `${e.key}  ${e.title}\n    LI: ${e.learningIntention || "-"}\n    SC2: ${sc2 || "-"}`;
}

function reviewCounts(entries) {
  const counts = new Map();
  entries.forEach((e) => (e.reviews || []).forEach((k) => {
    const c = counts.get(k) || { times: 0, last: null };
    c.times += 1;
    c.last = e.key;
    counts.set(k, c);
  }));
  return counts;
}

function showFor(specPath) {
  const spec = loadSpec(path.resolve(specPath));
  const L = spec.lesson || {};
  if (!Number.isInteger(L.term) || !Number.isInteger(L.week)) {
    console.error("The spec needs lesson.term and lesson.week to place it in teaching order.");
    process.exit(2);
  }
  const here = positionOf(L);
  const entries = readLog(L.yearLevel, L.subject);
  const earlier = entries.filter((e) => e.order < here.order);
  console.log(`Taught log for ${L.yearLevel} ${L.subject}: ${logPath(L.yearLevel, L.subject)}`);
  console.log(`This lesson: ${here.key}. ${earlier.length} earlier lesson(s) logged.`);
  if (!earlier.length) {
    console.log(`\nNothing logged before this lesson. Use from: "teacher" if the request names a focus, otherwise from: "before log" with review drawn from the previous term's content.`);
    return;
  }
  const counts = reviewCounts(entries);
  const groups = [
    ["Earlier this week", (gap) => gap === 0],
    ["Last week", (gap) => gap === 1],
    ["2 to 4 weeks ago", (gap) => gap >= 2 && gap <= 4],
    ["5 or more weeks ago", (gap) => gap >= 5],
  ];
  groups.forEach(([name, test]) => {
    const inGroup = earlier.filter((e) => test(here.weekIndex - e.weekIndex));
    if (!inGroup.length) return;
    console.log(`\n== ${name} ==`);
    inGroup.forEach((e) => {
      const c = counts.get(e.key);
      console.log(line(e));
      console.log(`    Reviewed: ${c ? `${c.times} time(s), last in ${c.last}` : "never"}`);
      (e.items || []).slice(0, 4).forEach((it) => {
        console.log(`    Item (${it.from}): ${it.prompt || "-"}${it.visual ? ` [visual ${JSON.stringify(it.visual)}]` : ""} -> ${it.answer || "-"}`);
      });
    });
  });
  console.log(`\nChoose Daily Review items that retrieve, spaced: one from last week, one from 2 to 4 weeks ago, one older where the log has it. Prefer lessons reviewed least. Write a NEW item of the same kind (new numbers, same skill), and set each dailyReview slide's "from" to the lesson key.`);
}

function main() {
  const args = process.argv.slice(2);
  if (args[0] === "--list" && args.length === 3) {
    const entries = readLog(args[1], args[2]);
    console.log(`${entries.length} lesson(s) in ${logPath(args[1], args[2])}`);
    entries.forEach((e) => console.log(line(e)));
    return;
  }
  if (args[0] === "--remove" && args.length === 4) {
    const ok = removeEntry(args[1], args[2], args[3]);
    console.log(ok ? `Removed ${args[3]}.` : `${args[3]} is not in the log.`);
    process.exit(ok ? 0 : 1);
  }
  if (args[0] === "--add" && (args.length === 5 || args.length === 6)) {
    const [, yearLevel, subject, key, title, detail] = args;
    const m = /^(\d{4})-T([1-4])-W(\d{1,2})-S(\d{1,2})$/.exec(key);
    if (!m) { console.error(`Key must look like 2026-T3-W7-S1 (got ${key}).`); process.exit(2); }
    const pos = positionOf({ year: Number(m[1]), term: Number(m[2]), week: Number(m[3]), session: Number(m[4]) });
    const entries = readLog(yearLevel, subject).filter((e) => e.key !== key);
    entries.push({
      key, order: pos.order, weekIndex: pos.weekIndex, year: pos.year, term: pos.term, week: pos.week, session: pos.session,
      title, learningIntention: detail || undefined, successCriteria: [], keyWords: [], reviews: [], items: [],
      recordedBy: "manual", builtAt: new Date().toISOString(),
    });
    console.log(`Recorded ${key} in ${writeLog(yearLevel, subject, entries)}`);
    return;
  }
  if (args.length === 1 && args[0].endsWith(".json")) { showFor(args[0]); return; }
  console.error("Usage:\n  node scripts/taught_log.js builds/<spec>.json\n  node scripts/taught_log.js --list <yearLevel> <subject>\n  node scripts/taught_log.js --remove <yearLevel> <subject> <key>\n  node scripts/taught_log.js --add <yearLevel> <subject> <key> \"<what was taught>\" [\"<detail>\"]");
  process.exit(2);
}

main();
