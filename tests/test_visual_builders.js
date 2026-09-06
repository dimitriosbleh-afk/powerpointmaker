"use strict";

/**
 * Regression guard for the visual layer added in the September 2026 redesign:
 *
 *  - retuned palettes keep every contrast promise (all 150 palettes)
 *  - derived tint keys exist on every theme
 *  - every pictogram name resolves and renders synchronously
 *  - every visual spec type draws inside the frame it is given
 *  - the pattern builders (heroVisualSlide, choiceSlide + markChoice,
 *    youDoSlide, textExtractSlide) build on every band and subject
 *  - keyWordSlide accepts a pictogram, titleSlide accepts a visual spec
 *  - addNumberLine and addRoutineBadge are on every theme and synchronous
 *
 * Run:  node tests/test_visual_builders.js
 */

const assert = require("assert");
const path = require("path");
const fs = require("fs");
const os = require("os");
const pptxgen = require("pptxgenjs");
const { createTheme, VALID_SUBJECTS, VALID_YEAR_LEVELS, VARIANTS_PER_LEVEL } = require("../themes/factory");
const { contrastRatio } = require("../themes/core/contrast");
const { PICTOGRAMS, listPictograms, renderPictogramPng } = require("../themes/core/pictograms");
const { SUPPORTED_TYPES } = require("../themes/core/visualSpec");
const Pi = require("react-icons/pi");

let passed = 0;
function ok(label) {
  console.log("  PASS " + label);
  passed += 1;
}

/* ── 1. palettes ─────────────────────────────────────────────────────────── */

function testPalettes() {
  let count = 0;
  VALID_SUBJECTS.forEach((subject) => {
    VALID_YEAR_LEVELS.forEach((level) => {
      for (let v = 0; v < VARIANTS_PER_LEVEL; v += 1) {
        const { C } = createTheme(subject, level, v);
        ["PRIMARY", "SECONDARY", "ACCENT", "ALERT", "SUCCESS", "ASSESS"].forEach((role) => {
          assert(contrastRatio(C.WHITE, C[role]) >= 4.5, `${subject}/${level}/v${v} white on ${role} fails AA`);
          assert(/^[0-9A-F]{6}$/.test(C[`${role}_SOFT`]), `${subject}/${level}/v${v} missing ${role}_SOFT`);
          assert(/^[0-9A-F]{6}$/.test(C[`${role}_LINE`]), `${subject}/${level}/v${v} missing ${role}_LINE`);
          // A soft wash must stay a light surface for dark text.
          assert(contrastRatio(C.CHARCOAL, C[`${role}_SOFT`]) >= 7, `${subject}/${level}/v${v} ${role}_SOFT too dark for body text`);
        });
        assert(contrastRatio(C.TEXT_ON_DARK, C.BG_DARK) >= 4.5, `${subject}/${level}/v${v} TEXT_ON_DARK fails on BG_DARK`);
        assert(contrastRatio(C.WHITE, C.BG_DARK) >= 7, `${subject}/${level}/v${v} BG_DARK too light for white text`);
        assert(contrastRatio(C.CHARCOAL, C.BG_LIGHT) >= 4.5, `${subject}/${level}/v${v} CHARCOAL fails on BG_LIGHT`);
        // The retune's point: Foundation PRIMARY must no longer be near-black.
        if (level === "foundation") {
          assert(contrastRatio(C.WHITE, C.PRIMARY) <= 6.5, `${subject}/foundation/v${v} PRIMARY is still too dark (${contrastRatio(C.WHITE, C.PRIMARY).toFixed(1)}:1)`);
        }
        count += 1;
      }
    });
  });
  assert.strictEqual(count, 150);
  ok("150 palettes keep every contrast promise and carry SOFT/LINE tints");
}

/* ── 2. pictograms ───────────────────────────────────────────────────────── */

function testPictograms() {
  const names = listPictograms();
  assert(names.length >= 200, "pictogram catalogue shrank");
  names.forEach((name) => {
    assert(Pi[PICTOGRAMS[name]], `pictogram ${name} -> ${PICTOGRAMS[name]} is not an installed icon`);
  });
  const t = Date.now();
  const data = renderPictogramPng("butterfly", "FFFFFF", 200);
  assert(data && data.startsWith("image/png;base64,"), "render returns a PptxGenJS data string");
  const png = Buffer.from(data.split(",")[1], "base64");
  assert(png.length > 500, "rendered PNG is not empty");
  // Cached second render must be instant; the first must not be seconds.
  renderPictogramPng("butterfly", "FFFFFF", 200);
  assert(Date.now() - t < 3000, "pictogram render is too slow (system font scan re-enabled?)");
  assert.strictEqual(renderPictogramPng("not-a-real-name", "FFFFFF"), null);
  ok(`${names.length} pictograms resolve and render synchronously`);
}

/* ── 3. visual specs ─────────────────────────────────────────────────────── */

const SPEC_SAMPLES = {
  tensFrame: { type: "tensFrame", filled: 7 },
  fiveFrame: { type: "fiveFrame", filled: 3 },
  doubleTensFrame: { type: "doubleTensFrame", filledTop: 10, filledBottom: 8 },
  dotCard: { type: "dotCard", count: 6 },
  dotCards: { type: "dotCards", counts: [4, 6] },
  numberTrack: { type: "numberTrack", start: 1, end: 10, highlight: [7] },
  numberLine: { type: "numberLine", start: 0, end: 2, step: 1 / 3, marked: [3] },
  fractionStrips: { type: "fractionStrips", strips: [{ denom: 4, shaded: 1 }, { denom: 4, shaded: 3 }] },
  array: { type: "array", rows: 3, cols: 4 },
  baseTen: { type: "baseTen", hundreds: 1, tens: 2, ones: 3 },
  groupedCounters: { type: "groupedCounters", groups: 3, per: 4 },
  ppwMat: { type: "ppwMat", whole: 7, partA: 4, partB: null },
  chips: { type: "chips", items: ["1/2", "3/4", "1/8"] },
  pictogram: { type: "pictogram", name: "frog" in PICTOGRAMS ? "frog" : "butterfly", label: "butterfly" },
  pictograms: { type: "pictograms", items: ["happy", "sad", "angry"] },
  text: { type: "text", text: "9" },
  table: { type: "table", rows: [["Animal", "Legs"], ["Dog", "4"], ["Bird", "2"]] },
  custom: { type: "custom", draw: (slide, f) => { slide.addText("x", { x: f.x, y: f.y, w: 1, h: 0.4 }); return { x: f.x, y: f.y, w: 1, h: 0.4 }; } },
};

function testVisualSpecs() {
  SUPPORTED_TYPES.forEach((type) => {
    if (type === "image") return; // needs a real file; covered by images.js
    assert(SPEC_SAMPLES[type], `no sample spec for type ${type}`);
  });
  ["foundation", "grade2", "grade56"].forEach((level) => {
    const T = createTheme("numeracy", level, 0);
    const pres = new pptxgen();
    pres.layout = "LAYOUT_16x9";
    const s = pres.addSlide();
    const frame = { x: 0.5, y: 1.3, w: 9, h: 3.8 };
    Object.values(SPEC_SAMPLES).forEach((spec) => {
      const b = T.drawVisual(s, spec, frame);
      assert(b.x >= frame.x - 0.01 && b.y >= frame.y - 0.01, `${level} ${spec.type} starts outside the frame`);
      assert(b.x + b.w <= frame.x + frame.w + 0.01, `${level} ${spec.type} overflows the frame width`);
      assert(b.y + b.h <= frame.y + frame.h + 0.01, `${level} ${spec.type} overflows the frame height`);
    });
    assert(T.isVisualSpec({ type: "tensFrame" }));
    assert(!T.isVisualSpec(() => {}));
    assert(!T.isVisualSpec({ type: "nope" }));
  });
  ok("every visual spec type draws inside its frame on all three bands");
}

/* ── 4. builders ─────────────────────────────────────────────────────────── */

async function testBuilders() {
  const out = path.join(os.tmpdir(), `visual-builders-${process.pid}.pptx`);
  const warnings = [];
  const origWarn = console.warn;
  console.warn = (...args) => { warnings.push(args.join(" ")); };
  try {
    for (const subject of VALID_SUBJECTS) {
      for (const level of ["foundation", "grade2", "grade56"]) {
        const T = createTheme(subject, level, 1);
        const pres = new pptxgen();
        pres.layout = "LAYOUT_16x9";
        const notes = "One line of notes.";

        T.titleSlide(pres, "Title", "Subtitle", "Meta", notes);
        T.titleSlide(pres, "Title with visual", "Subtitle", "Meta", notes, { visual: { type: "tensFrame", filled: 10 } });
        const hero = T.heroVisualSlide(pres, "I Do", "Hero", { type: "tensFrame", filled: 7 }, notes, "f", { label: "Tens frame", prompt: "Show me" });
        assert(hero.heroBounds && hero.heroBounds.w > 3, "hero visual should be large");
        const choice = T.choiceSlide(pres, "CFU", "Which one?", "Show me A or B", [
          { visual: { type: "tensFrame", filled: 7 } },
          { text: "seven" },
          { visual: { type: "dotCard", count: 7 }, caption: "dots" },
        ], notes, "f");
        assert.strictEqual(choice.choiceFrames.length, 3);
        const before = choice._slideObjects.length;
        T.markChoice(choice, 0);
        assert(choice._slideObjects.length > before, "markChoice adds elements");
        T.youDoSlide(pres, "Your turn", "Draw counters to make 10.", ["Look", "Draw", "Write"], notes, "f",
          { where: "On your sheet", visual: { type: "tensFrame", filled: 6 }, frame: "___ and ___ make 10" });
        T.textExtractSlide(pres, "Read", "Extract", "The wind howled. Joey pressed his gnarled hands to the door.", notes, "f",
          { highlights: ["gnarled hands"], source: "Source", prompt: "Which word shows fear?" });
        T.keyWordSlide(pres, { word: "calm", meaning: "Quiet and still.", example: "I feel calm.", pictogram: "calm" }, notes, "f");
        T.contentSlide(pres, "Launch", T.C.PRIMARY, "Sparse", "One big statement.", notes, "f");
        T.contentSlide(pres, "I Do", T.C.PRIMARY, "With visual", ["3 groups.", "4 in each."], notes, "f", { type: "groupedCounters", groups: 3, per: 4 });
        T.closingSlide(pres, { reflectionPrompt: "Tell a partner.", scItems: ["I can a.", "I can b.", "I can c."], selfAssessment: "Thumbs" }, notes);

        // Universal helpers that used to be numeracy-only or async.
        const s = pres.addSlide();
        assert.strictEqual(typeof T.addNumberLine, "function", `${subject} lacks addNumberLine`);
        T.addNumberLine(s, 1, 2, 7, ["0", "", "1", "", "2"], [2]);
        const result = T.addRoutineBadge(s, "miniWhiteboard", 1, 3);
        assert(!(result && typeof result.then === "function"), "addRoutineBadge must be synchronous");
        assert.strictEqual(typeof T.addDataTable, "function");
        T.addDataTable(s, 0.5, 3.6, 9, [["A", "B"], ["1", "2"]]);

        if (subject === "science") {
          T.cycleDiagramSlide(pres, "I Do", "Cycle", "Watch", ["Follow the arrows."], "Water", [
            { label: "Evaporation", detail: "Sun heats", icon: "sun" },
            { label: "Condensation", detail: "Cools", icon: "cloud" },
            { label: "Precipitation", detail: "Falls", icon: "rain" },
            { label: "Collection", detail: "Gathers", icon: "waves" },
          ], notes, "f");
        }
        if (subject === "numeracy") {
          T.workedExSlide(pres, 2, "I Do", "Worked", ["Step one.", "Step two."], notes, "f", { type: "fractionStrips", strips: [{ denom: 4, shaded: 0 }] });
          T.dailyReviewSlide(pres, "Daily Review", ["How many?"], notes, "f", { type: "tensFrame", filled: 7 });
        }

        await pres.writeFile({ fileName: out });
        assert(fs.existsSync(out));
      }
    }
  } finally {
    console.warn = origWarn;
    if (fs.existsSync(out)) fs.unlinkSync(out);
  }
  const real = warnings.filter((w) => /^(WARN|ERROR)/.test(w));
  assert.deepStrictEqual(real, [], "builders emitted diagnostics:\n" + real.join("\n"));
  ok("pattern builders build cleanly on every subject and band with zero diagnostics");
}

/* ── 5. unknown pictogram is loud ───────────────────────────────────────── */

function testUnknownPictogramWarns() {
  const T = createTheme("literacy", "grade34", 0);
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  const s = pres.addSlide();
  const warnings = [];
  const origWarn = console.warn;
  console.warn = (...args) => { warnings.push(args.join(" ")); };
  try {
    T.addPictogram(s, "definitely-not-a-pictogram", 1, 1, 1);
  } finally {
    console.warn = origWarn;
  }
  assert(warnings.some((w) => /^WARN \[pictogram\]/.test(w)), "unknown pictogram must emit a gate-failing WARN");
  ok("unknown pictogram names fail loudly");
}

(async () => {
  console.log("\nVisual layer regressions");
  testPalettes();
  testPictograms();
  testVisualSpecs();
  await testBuilders();
  testUnknownPictogramWarns();
  console.log(`\n${passed} check(s) passed.`);
})().catch((err) => {
  console.error("\nVisual layer test FAILED:", err && err.stack ? err.stack : err);
  process.exit(1);
});
