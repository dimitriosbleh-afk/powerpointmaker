"use strict";

// Smoke test: render LI/SC, content, CFU and closing slides at all three
// grade bands so the title-overflow and tier-label fixes can be verified
// visually. Single deck, three sections (F / Y12 / Y36).
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

    T.titleSlide(pres, `Smoke Test - ${band}`, "LI/SC and closing render", label, "Internal smoke test.");
    T.liSlide(pres, SAMPLE_LI, SAMPLE_SC, "Internal smoke test.", label);
    T.contentSlide(pres, "I Do", T.C.PRIMARY,
      "Title with a Long Phrase to Stress the Title Bar Geometry",
      [
        "First content line that should sit cleanly inside the card.",
        "Second content line.",
        "Third content line.",
      ],
      "Internal smoke test.",
      label);
    T.cfuSlide(pres, "CFU", "Hinge Question Title That Is Reasonably Long",
      "Show Me Boards",
      "Which fraction strip shows one half? Hold up your card.",
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
