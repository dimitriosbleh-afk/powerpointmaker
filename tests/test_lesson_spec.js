"use strict";

/**
 * Lesson spec pipeline regression guard.
 *
 *  - every golden exemplar in builds/ validates with zero errors, builds
 *    into a scratch folder, and produces a PPTX plus its declared PDFs
 *  - the validator names the field for the mistakes a weaker model makes
 *    most: a typo'd kind, an unknown field, a missing routine cue, a word
 *    card without a picture, a reveal on a prompt slide, a bad pictogram
 *
 * Run:  node tests/test_lesson_spec.js
 */

const assert = require("assert");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { validateLessonSpec, yourTurnSteps, followsPlanningRules } = require("../themes/lesson/validate");
const { buildLesson, loadSpec } = require("../themes/lesson/buildLesson");

const ROOT = path.resolve(__dirname, "..");
let passed = 0;
function ok(label) { console.log("  PASS " + label); passed += 1; }

const exemplars = fs.readdirSync(path.join(ROOT, "builds"))
  .filter((n) => /^exemplar_.*\.json$/.test(n))
  .map((n) => path.join(ROOT, "builds", n));

async function testExemplarsBuild() {
  assert(exemplars.length >= 3, "expected at least three exemplar specs in builds/");
  const scratch = fs.mkdtempSync(path.join(os.tmpdir(), "lesson-spec-"));
  const silent = console.log;
  for (const file of exemplars) {
    const spec = loadSpec(file);
    const { errors, warnings } = validateLessonSpec(spec);
    assert.deepStrictEqual(errors, [], `${path.basename(file)} should validate: ${errors.join("; ")}`);
    assert.deepStrictEqual(warnings, [], `${path.basename(file)} should have no spec warnings: ${warnings.join("; ")}`);
    console.log = () => {};
    let result;
    try {
      result = await buildLesson(spec, { outRoot: scratch });
    } finally {
      console.log = silent;
    }
    assert(fs.existsSync(result.pptxPath), `${path.basename(file)}: PPTX written`);
    result.resources.forEach((r) => {
      assert(fs.existsSync(path.join(result.outDir, r.fileName)), `${path.basename(file)}: resource ${r.fileName} written`);
    });
  }
  fs.rmSync(scratch, { recursive: true, force: true });
  ok(`${exemplars.length} exemplar specs validate clean and build with their resources`);
}

function testValidatorNamesTheMistake() {
  const base = loadSpec(exemplars[0]);
  const clone = () => JSON.parse(JSON.stringify(base));

  const badKind = clone();
  badKind.slides[3].kind = "heroVisua";
  assert(validateLessonSpec(badKind).errors.some((e) => /kind: "heroVisua" is not a slide kind/.test(e)), "typo'd kind is named");

  const badField = clone();
  badField.slides[3].visuals = badField.slides[3].visual;
  assert(validateLessonSpec(badField).errors.some((e) => /\.visuals: unknown field/.test(e)), "unknown field is named");

  const noPicture = clone();
  const kw = noPicture.slides.find((s) => s.kind === "keyWord");
  delete kw.pictogram;
  assert(validateLessonSpec(noPicture).errors.some((e) => /word card needs a picture/.test(e)), "word card without a picture is an error");

  const badPicto = clone();
  noPicture.slides.find((s) => s.kind === "keyWord").pictogram = "frog";
  assert(validateLessonSpec(noPicture).errors.some((e) => /"frog" is not a pictogram/.test(e)), "unknown pictogram is named");
  void badPicto;

  const promptAndReveal = clone();
  const hero = promptAndReveal.slides.find((s) => s.kind === "heroVisual" && s.reveal);
  hero.prompt = "Show me";
  assert(validateLessonSpec(promptAndReveal).errors.some((e) => /cannot also have a prompt bar/.test(e)), "prompt plus reveal is an error");

  const dash = clone();
  dash.slides[0].notes = "Open the lesson — with a dash";
  assert(validateLessonSpec(dash).errors.some((e) => /em\/en dash/.test(e)), "banned characters are named with their path");

  const noLaunch = clone();
  noLaunch.slides = noLaunch.slides.filter((s) => s.kind !== "launch");
  assert(validateLessonSpec(noLaunch).errors.some((e) => /no launch after the LI/.test(e)), "missing launch is an error");

  ok("validator names the common mistakes with their field paths");
}

/** The evidence rules: mapped distractors on checks, and an exit ticket that is a new, individual item. */
function testEvidenceRules() {
  const byName = (n) => loadSpec(exemplars.find((f) => path.basename(f).includes(n)));
  const errorsOf = (spec) => validateLessonSpec(spec).errors;

  const found = byName("foundation");
  const check = found.slides.find((s) => s.kind === "choice" && s.answer != null);
  delete check.options.find((o, i) => i !== check.answer).misconception;
  assert(errorsOf(found).some((e) => /misconception: every wrong option/.test(e)), "a wrong option with no misconception is an error");

  const lit = byName("year2");
  const exitAt = lit.slides.findIndex((s) => s.kind === "exitTicket");
  const extract = lit.slides.find((s) => s.kind === "textExtract").extract;
  const reusedName = JSON.parse(JSON.stringify(lit));
  reusedName.slides[exitAt].visual = { type: "text", text: "Point to how Mia felt at the pool." };
  assert(errorsOf(reusedName).some((e) => /reuses "Mia"/.test(e)), "an exit ticket reusing the I Do character is an error");

  const copied = JSON.parse(JSON.stringify(lit));
  copied.slides[exitAt].questions = [extract];
  assert(errorsOf(copied).some((e) => /six or more words in a row/.test(e)), "an exit ticket copying an earlier text is an error");

  const partner = JSON.parse(JSON.stringify(lit));
  partner.slides[exitAt].notes.beats[1][1] = "10 sec. Cue: turn and tell, partner A first.";
  assert(errorsOf(partner).some((e) => /partner talk or a choral answer/.test(e)), "exit evidence through partner talk is an error");

  const sameVisual = JSON.parse(JSON.stringify(found));
  const fExit = sameVisual.slides.find((s) => s.kind === "exitTicket");
  fExit.visual = JSON.parse(JSON.stringify(sameVisual.slides.find((s) => s.kind === "heroVisual" && /I Do/.test(s.badge)).visual));
  assert(errorsOf(sameVisual).some((e) => /the same visual as/.test(e)), "an exit ticket reusing a modelled visual is an error");

  const sci = byName("science");
  const sExit = sci.slides.find((s) => s.kind === "exitTicket");
  sExit.questions = [sci.slides.find((s) => s.kind === "cfu").question];
  assert(errorsOf(sci).some((e) => /asks the same question as/.test(e)), "an exit ticket repeating a check question is an error");

  const noPlan = JSON.parse(JSON.stringify(lit));
  delete noPlan.lesson.plan;
  assert(errorsOf(noPlan).some((e) => /lesson\.plan: required/.test(e)), "a lesson without a plan is an error");
  const thinPlan = JSON.parse(JSON.stringify(lit));
  thinPlan.lesson.plan.decisionPoints = ["the exit ticket"];
  delete thinPlan.lesson.plan.criticalFeature;
  assert(errorsOf(thinPlan).some((e) => /decisionPoints: 2 to 4/.test(e)), "one decision point is not enough");
  assert(errorsOf(thinPlan).some((e) => /criticalFeature/.test(e)), "the critical feature is required");

  ok("checks need mapped distractors; exit tickets must be a new item with individual evidence; every lesson carries its plan");
}

/** Practice rounds expand to one slide per item with notes that pass the Glance gate. */
function testPracticeRounds() {
  const { expandSpec } = require("../themes/lesson/practice");
  const { composeGlanceNotes } = require("../themes/core/composeNotes");
  const spec = loadSpec(exemplars.find((f) => path.basename(f).includes("year2")));
  const rounds = spec.slides.filter((s) => s.kind === "practice");
  assert(rounds.length >= 1, "the Year 2 exemplar carries a practice round");
  const expanded = expandSpec(spec);
  const itemCount = rounds.reduce((n, r) => n + r.items.length, 0);
  assert.strictEqual(expanded.slides.length, spec.slides.length - rounds.length + itemCount, "each round becomes one slide per item");
  expanded.slides.filter((s) => s.practiceRound).forEach((s) => {
    const notes = composeGlanceNotes(s.notes);
    assert(/ASK:/.test(notes) && /SCAN/.test(notes) && /REVEAL after/.test(notes), "every item slide asks, scans and reveals");
    assert(s.reveal && s.reveal.answers.length, "every item slide reveals its answer on click");
  });

  const bad = JSON.parse(JSON.stringify(spec));
  const r = bad.slides.find((s) => s.kind === "practice");
  r.items = r.items.slice(0, 2);
  r.routine = "hands";
  delete r.pivot;
  const errs = validateLessonSpec(bad).errors;
  assert(errs.some((e) => /3 to 8 items/.test(e)), "a round needs 3 to 8 items");
  assert(errs.some((e) => /routine: use one of/.test(e)), "an unknown routine is named");
  assert(errs.some((e) => /\.pivot: /.test(e)), "a round needs a pivot");
  ok("practice rounds expand to hero slides with composed notes, and malformed rounds are named");
}

/** Practice volume: whole-class responses and independent items have floors. */
function testPracticeVolume() {
  const { practiceCounts } = require("../themes/lesson/validate");
  const lit = loadSpec(exemplars.find((f) => path.basename(f).includes("year2")));
  const noRounds = JSON.parse(JSON.stringify(lit));
  noRounds.slides = noRounds.slides.filter((s) => s.kind !== "practice");
  const errs = validateLessonSpec(noRounds).errors;
  assert(errs.some((e) => /planned whole-class responses/.test(e)), "a thin lesson fails the response floor");
  assert(errs.some((e) => /independent items/.test(e)), "a lesson with no independent practice fails");

  const short = JSON.parse(JSON.stringify(lit));
  short.lesson.minutes = 45;
  assert(practiceCounts(short).responses >= Math.ceil(45 / 3), "the floor scales with lesson length");

  const found = loadSpec(exemplars.find((f) => path.basename(f).includes("foundation")));
  found.resources[0].items = found.resources[0].items.slice(0, 2);
  assert(validateLessonSpec(found).errors.some((e) => /foundation needs at least 4/.test(e)), "the independent floor is per band");
  const extended = JSON.parse(JSON.stringify(found));
  extended.slides.find((s) => s.kind === "youDo").extendedTask = "draw and explain one model";
  assert(!validateLessonSpec(extended).errors.some((e) => /independent items/.test(e)), "one extended task waives the item count");
  ok("response and independent-practice floors hold, scale with minutes, and waive for an extended task");
}

/** The taught log orders by teaching position and checks Daily Review sources against it. */
function testTaughtLog() {
  const { recordLesson, readLog, removeEntry } = require("../themes/lesson/taughtLog");
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "taught-log-"));
  const base = loadSpec(exemplars.find((f) => path.basename(f).includes("foundation")));
  const at = (term, week, froms) => {
    const s = JSON.parse(JSON.stringify(base));
    Object.assign(s.lesson, { year: 2026, term, week, session: 1 });
    s.slides.filter((x) => x.kind === "dailyReview").forEach((d, i) => { d.from = froms[i]; });
    return s;
  };
  recordLesson(at(3, 1, ["before log", "before log", "before log"]), null, dir);
  recordLesson(at(2, 5, ["before log", "before log", "before log"]), null, dir);
  recordLesson(at(3, 1, ["before log", "before log", "before log"]), null, dir);
  const log = readLog("foundation", "numeracy", dir);
  assert.deepStrictEqual(log.map((e) => e.key), ["2026-T2-W5-S1", "2026-T3-W1-S1"], "teaching order, and a rebuild replaces its entry");
  assert(log[0].items[0].from === "exit ticket", "the exit ticket is the first review item");

  const errorsWith = (spec) => validateLessonSpec(spec, { taughtLog: log }).errors;
  assert(errorsWith(at(3, 3, ["2026-T3-W1-S1", "2026-T2-W5-S1", "teacher"])).every((e) => !/dailyReview/.test(e)), "earlier, spaced sources pass");
  assert(errorsWith(at(3, 3, ["2026-T4-W1-S1", "teacher", "teacher"])).some((e) => /is not in the taught log/.test(e)), "an unknown key is named");
  assert(errorsWith(at(2, 5, ["2026-T3-W1-S1", "teacher", "teacher"])).some((e) => /is not earlier than this lesson/.test(e)), "a later lesson cannot be reviewed");
  assert(errorsWith(at(3, 2, ["2026-T3-W1-S1", "2026-T3-W1-S1", "2026-T3-W1-S1"])).some((e) => /two or more weeks ago/.test(e)), "review must reach back when older learning exists");
  const noFrom = at(3, 3, ["teacher", "teacher", "teacher"]);
  delete noFrom.slides.find((x) => x.kind === "dailyReview").from;
  assert(validateLessonSpec(noFrom).errors.some((e) => /\.from: required/.test(e)), "every Daily Review slide names its source");
  const noTerm = at(3, 3, ["teacher", "teacher", "teacher"]);
  delete noTerm.lesson.term;
  assert(validateLessonSpec(noTerm).errors.some((e) => /lesson\.term/.test(e)), "a lesson needs its term");

  const { positionOf } = require("../themes/lesson/taughtLog");
  assert(positionOf({ year: 2026, term: 4, week: 3, session: 12 }).order < positionOf({ year: 2026, term: 4, week: 4, session: 1 }).order,
    "a unit that numbers sessions past 9 still sorts by week first");

  assert(removeEntry("foundation", "numeracy", "2026-T2-W5-S1", dir), "an entry can be removed");
  assert.strictEqual(readLog("foundation", "numeracy", dir).length, 1, "removal persists");
  fs.rmSync(dir, { recursive: true, force: true });
  ok("taught log keeps teaching order and Daily Review sources are checked against it");
}

/** The Years 5-6 maths planning team's rules (megaprompt 85). */
function testPlanner56() {
  const base = loadSpec(exemplars.find((f) => path.basename(f).includes("year56_numeracy")));
  const clone = () => JSON.parse(JSON.stringify(base));
  const errorsOf = (spec) => validateLessonSpec(spec).errors;
  assert.deepStrictEqual(errorsOf(base), [], "the Years 5-6 maths exemplar meets every planner rule");

  const noExtension = clone();
  noExtension.resources = noExtension.resources.filter((r) => r.role !== "extension");
  assert(errorsOf(noExtension).some((e) => /role "extension"/.test(e)), "every session needs an Extension");

  const ownSteps = clone();
  ownSteps.slides.find((s) => s.kind === "youDo").steps = ["Read", "Solve", "Check"];
  assert(errorsOf(ownSteps).some((e) => /standard steps/.test(e)), "Your turn uses the standard steps");

  const launchFirst = clone();
  const li = launchFirst.slides.findIndex((s) => s.kind === "li");
  const launch = launchFirst.slides.findIndex((s) => s.kind === "launch");
  const [launchSlide] = launchFirst.slides.splice(launch, 1);
  launchFirst.slides.splice(li, 0, launchSlide);
  assert(errorsOf(launchFirst).some((e) => /cannot come before the LI/.test(e)), "the launch comes after the LI and SC");

  const lessonWord = clone();
  lessonWord.slides.find((s) => s.kind === "li").notes.prep = "The core of the lesson.";
  assert(errorsOf(lessonWord).some((e) => /the word "lesson"/.test(e)), "teacher notes never use the word lesson");

  const repeated = clone();
  const slideVisual = repeated.slides.find((s) => s.kind === "practice").items[0].visual;
  repeated.resources.find((r) => r.role === "main").sections[1].items[0].visual = JSON.parse(JSON.stringify(slideVisual));
  assert(errorsOf(repeated).some((e) => /same numbers/.test(e)), "a slide question may not reappear on a sheet");

  const short = clone();
  short.resources.find((r) => r.role === "main").sections.forEach((sec) => { sec.items = sec.items.slice(0, 2); });
  assert(errorsOf(short).some((e) => /at least 12/.test(e)), "the main sheet is long enough for fast finishers");

  const bland = clone();
  bland.resources.find((r) => r.role === "main").sections.forEach((sec) => {
    sec.items = sec.items.map((it) => ({ prompt: "Measure this angle.", answer: it.answer || "40" }));
  });
  const blandErrors = errorsOf(bland);
  assert(blandErrors.some((e) => /question format/.test(e)), "a main sheet mixes question kinds");
  assert(blandErrors.some((e) => /four questions in a row ask the same thing/.test(e)), "four identical questions in a row are a drill, not a worksheet");
  ok("Years 5-6 maths: three sheets, standard Your turn, LI before launch, no repeated questions, long and varied main sheet");
}

/** The planning team's rules beyond Years 5-6 maths: every deck, and the Years 3-6 package in every subject but wellbeing (megaprompt 85). */
function testPlanningRulesEverySubject() {
  const byName = (n) => loadSpec(exemplars.find((f) => path.basename(f).includes(n)));
  const errorsOf = (spec) => validateLessonSpec(spec).errors;

  // Every deck, Foundation included: LI before the launch, and an extension task is called Extension.
  const found = byName("foundation");
  const li = found.slides.findIndex((s) => s.kind === "li");
  const [launch] = found.slides.splice(found.slides.findIndex((s) => s.kind === "launch"), 1);
  found.slides.splice(li, 0, launch);
  assert(errorsOf(found).some((e) => /cannot come before the LI/.test(e)), "Foundation also puts LI and SC before the launch");
  const challenge = byName("foundation");
  challenge.resources.push({ kind: "page", label: "Challenge Card", blocks: [{ heading: "Make 10 two ways" }] });
  assert(errorsOf(challenge).some((e) => /is an extension task/.test(e)), "a Challenge resource must be named Extension");
  assert(!errorsOf(byName("foundation")).some((e) => /role "main"/.test(e)), "Foundation keeps its hands-on default: no three-sheet package");

  // Years 3-6 science carries the package and the task wording of the Your turn steps.
  const sci = byName("science");
  assert.deepStrictEqual(errorsOf(sci), [], "the science exemplar meets the package");
  assert.deepStrictEqual(yourTurnSteps("science"), yourTurnSteps("literacy"), "science uses the task wording of the steps");
  const noSupported = byName("science");
  noSupported.resources = noSupported.resources.filter((r) => r.role !== "supported");
  assert(errorsOf(noSupported).some((e) => /role "supported"/.test(e)), "Years 5-6 science needs a supported sheet");
  const grade34 = byName("science");
  grade34.lesson.yearLevel = "grade34";
  grade34.resources = [];
  assert(errorsOf(grade34).some((e) => /role "main"/.test(e)), "Years 3-4 carry the package too");
  const wellbeing = byName("science");
  wellbeing.lesson.subject = "wellbeing";
  assert(!followsPlanningRules(wellbeing), "wellbeing is discussion-led and outside the package");

  // The same sentence on a slide and a sheet is a repeat outside maths.
  const copied = byName("science");
  const slideText = copied.slides.find((s) => s.kind === "practice").items[0].text;
  copied.resources.find((r) => r.role === "main").sections[0].items.push({ prompt: slideText, answer: "precipitation" });
  assert(errorsOf(copied).some((e) => /word for word/.test(e)), "a slide sentence may not reappear on a sheet");
  ok("every deck: LI before launch, Extension naming; Years 3-6 package in science; no repeated sentences");
}

function testCrossword() {
  const { buildCrossword } = require("../themes/lesson/crossword");
  const words = ["rebuttal", "audience", "modality", "evidence", "emotive", "hook", "slogan", "perspective"]
    .map((answer) => ({ answer, clue: "clue" }));
  const cw = buildCrossword({ label: "test", words });
  assert.strictEqual(cw.starts.length, words.length, "every word is placed");
  // Each placed word reads back from the grid, and words only meet where letters match.
  cw.starts.forEach((s) => {
    const read = s.word.split("").map((_, i) => cw.cells.get(`${s.r + (s.dir === "down" ? i : 0)},${s.c + (s.dir === "across" ? i : 0)}`)).join("");
    assert.strictEqual(read, s.word, `${s.word} reads back from the grid`);
  });
  const again = buildCrossword({ label: "test", words });
  assert.deepStrictEqual([...again.cells], [...cw.cells], "the layout is deterministic");
  const spec = loadSpec(exemplars.find((f) => /literacy/.test(f)));
  spec.resources = [{ kind: "crossword", label: "Bad Crossword", words: [{ answer: "xyz", clue: "a" }, { answer: "qqq", clue: "b" }, { answer: "www", clue: "c" }, { answer: "vvv", clue: "d" }] }];
  const { errors } = validateLessonSpec(spec);
  assert(errors.some((e) => /could not fit/.test(e)), "a word that cannot cross is named at validation");
  ok("crossword: every word placed and readable, deterministic, unplaceable words named");
}

(async () => {
  await testExemplarsBuild();
  testPlanner56();
  testPlanningRulesEverySubject();
  testValidatorNamesTheMistake();
  testEvidenceRules();
  testPracticeRounds();
  testPracticeVolume();
  testTaughtLog();
  testCrossword();
  console.log(`${passed} check(s) passed.`);
})().catch((err) => {
  console.error("FAIL " + (err && err.stack ? err.stack : err));
  process.exit(1);
});
