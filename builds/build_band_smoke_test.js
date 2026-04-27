"use strict";

// Smoke test: render LI/SC, content, CFU, workedEx and closing slides at all
// three grade bands so the title-overflow, tier-label, bullet-overflow and
// CFU-overflow fixes can be verified visually. Stress-tests with dense
// content (8-9 bullets with empty-string spacers, multi-line CFU questions)
// matching the PPWI lesson density that originally regressed.
//
// Not a teaching artefact. Delete after verification.

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const { createTheme } = require("../themes/factory");

const SAMPLE_LI = "We are learning to investigate fractions and reason about equal parts.";
const SAMPLE_SC = [
  "I can name the parts of a fraction.",
  "I can show a fraction on a fraction strip.",
  "I can compare two fractions and explain which is larger.",
];

// Dense bullet list with empty-string spacers — this is the content shape
// that overflowed in PPWI Foundation lessons before the fix.
const DENSE_BULLETS = [
  "I have 5 counters.",
  "5 is the whole.",
  "",
  "I shake and spill.",
  "",
  "4 are red.",
  "1 is yellow.",
  "",
  "Whole 5.",
  "Parts 4 and 1.",
];

// First/Next/Then bullet pattern used by You Do slides — 8 items with two
// spacers, similar to PPWI lesson 3 "Your Turn: 6 and 7".
const FIRST_NEXT_THEN = [
  "First: 6 counters.",
  "Find one way to make 6.",
  "",
  "Next: 7 counters.",
  "Find one way to make 7.",
  "",
  "Tell your partner:",
  "Whole ___, parts ___ and ___.",
];

// Multi-line CFU question similar to PPWI lesson 2 "Is This Right?" — the
// previous build overflowed both above (overlapping the technique pill)
// and below the question card.
const DENSE_CFU = "Look at the counters.\n\nWhole is 5.\nParts are 2 and 1.\n\nIs that right?";

const REFLECTION = "Turn and tell: which fraction is larger and how do you know?";
const SELF_ASSESS = {
  prompt: "Self-assess: thumbs up, sideways, or down.",
  options: ["Got it", "Getting there", "Need more"],
};

(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  for (const band of ["foundation", "grade2", "grade56"]) {
    const T = createTheme("numeracy", band, 0);
    const label = `Band: ${band}`;

    T.titleSlide(pres, `Smoke Test - ${band}`, "LI/SC, content, CFU, workedEx, closing", label, "Internal smoke test.");
    T.liSlide(pres, SAMPLE_LI, SAMPLE_SC, "Internal smoke test.", label);

    // Dense bullets with empty-string spacers — left card stress test.
    T.contentSlide(pres, "I Do", T.C.PRIMARY,
      "Dense Bullets with Empty Spacers",
      DENSE_BULLETS,
      "Internal smoke test.",
      label);

    // Dense workedEx with right-side visual placeholder — narrow column.
    T.workedExSlide(pres, 2, "I Do", "Dense Worked Example",
      DENSE_BULLETS,
      "Internal smoke test.",
      label,
      (slide, lg) => {
        slide.addShape("roundRect", {
          x: lg.rightX, y: lg.panelTopPadded,
          w: lg.rightW, h: 1.2, rectRadius: 0.08,
          fill: { color: T.C.PRIMARY },
        });
        slide.addText("Whole = 5", {
          x: lg.rightX, y: lg.panelTopPadded, w: lg.rightW, h: 1.2,
          fontFace: T.FONT_H, fontSize: 32, color: T.C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      });

    // First/Next/Then You Do pattern — narrow column with spacers.
    T.workedExSlide(pres, 4, "You Do", "Your Turn: First, Next, Then",
      FIRST_NEXT_THEN,
      "Internal smoke test.",
      label,
      (slide, lg) => {
        slide.addShape("roundRect", {
          x: lg.rightX, y: lg.panelTopPadded,
          w: lg.rightW, h: 1.0, rectRadius: 0.08,
          fill: { color: T.C.PRIMARY },
        });
        slide.addText("Whole = 7", {
          x: lg.rightX, y: lg.panelTopPadded, w: lg.rightW, h: 1.0,
          fontFace: T.FONT_H, fontSize: 28, color: T.C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      });

    // Multi-line CFU question — previously overflowed above and below.
    T.cfuSlide(pres, "CFU", "Is This Right?",
      "Thumbs Up or Thumbs Down",
      DENSE_CFU,
      "Internal smoke test.",
      label);

    T.closingSlide(pres, {
      reflectionPrompt: REFLECTION,
      scItems: SAMPLE_SC,
      selfAssessment: SELF_ASSESS,
    }, "Internal smoke test.");
  }

  const outDir = path.join(__dirname, "..", "output", "BandSmokeTest");
  fs.mkdirSync(outDir, { recursive: true });
  const outPath = path.join(outDir, "BandSmokeTest.pptx");
  await pres.writeFile({ fileName: outPath });
  console.log("PPTX written to " + outPath);
})();
