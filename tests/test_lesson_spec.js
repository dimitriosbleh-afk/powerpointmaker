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
const { validateLessonSpec } = require("../themes/lesson/validate");
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
  assert(validateLessonSpec(noLaunch).errors.some((e) => /no launch before the LI/.test(e)), "missing launch is an error");

  ok("validator names the common mistakes with their field paths");
}

(async () => {
  await testExemplarsBuild();
  testValidatorNamesTheMistake();
  console.log(`${passed} check(s) passed.`);
})().catch((err) => {
  console.error("FAIL " + (err && err.stack ? err.stack : err));
  process.exit(1);
});
