/* eslint-disable no-console */
"use strict";

// Smoke test for the megaprompt v3.0 alignment work. Exercises the new
// builders, manipulatives, routine icons, placeholders, and notes
// composer end-to-end so that any wiring break shows up before the
// next bulk-output run.

const PptxGenJS = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant, composeNotes } = require("./themes/factory");

function ensureDir(p) {
  fs.mkdirSync(path.dirname(p), { recursive: true });
}

async function buildFoundationNumeracy() {
  const T = createTheme("numeracy", "foundation", weekToVariant(1));
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  const notes = composeNotes({
    say: ["Today we are counting to 10.", "Watch me build the tens frame."],
    do: ["Hand each pair a tens frame.", "Demonstrate 10 and 8 more is 18."],
    cfu: {
      technique: "Show Me Boards",
      script: ["Ask: How many counters do I have?", "Scan all responses."],
      scanFor: "students writing a number-name (eight)",
      proceed: ["Move to We Do."],
      pivot: {
        misconception: "Students count the empty cells.",
        reteach: "Re-anchor with the five frame first.",
        recheck: "Show 5 with one extra counter.",
      },
    },
    teacherNotes: "Use blue counters for the teaching, red for student practice.",
    watchFor: ["Reverse counting (10, 9, 8...) — gently redirect."],
    tag: "Stage I Do | VTLM 2.0 explicit teaching",
  });

  T.titleSlide(pres, "Tens Frame to 10", "Foundation | Week 1", "Counting collections", notes);

  T.liSlide(pres, "I can count to 10 using a tens frame.", [
    "I can place counters 1-by-1 on a tens frame.",
    "I can name how many counters are on the frame.",
    "I can describe the missing parts to make 10.",
  ], notes, "Foundation | Week 1");

  T.dailyReviewSlide(pres, "Daily Review: counting", [
    "How many fingers am I holding up?",
    "Count to 5 with me.",
  ], notes, "Daily Review", (s) => {
    T.addFiveFrame(s, 5.5, 1.6, 3.6, 3, { fillColor: T.C.PRIMARY });
  });

  T.fluencySlide(pres, "Fluency: numbers to 5", [
    "3",
    "5",
  ], notes);

  T.contentSlide(pres, "I Do", T.C.PRIMARY, "Watch me build 8", [
    "Place one counter at a time.",
    "Count out loud.",
  ], notes, "Foundation | Week 1", (s) => {
    T.addTensFrame(s, 5.4, 1.8, 3.6, 8);
  });

  T.cfuSlide(pres, "CFU", "Show me 6", {
    technique: "Show Me Boards",
    question: "Build 6 on your tens frame.",
  }, notes, "Foundation | Week 1");

  T.boardBuildSlide(pres, "Build Together", "Build 9 on the giant frame",
    "Build this together",
    notes, "Foundation | Week 1",
    {
      promptText: "Teacher draws an empty frame and class fills it.",
      prefilledHints: ["1", "2", "?", "?", "?"],
    });

  T.exitTicketSlide(pres, [
    "Show me 7 on your tens frame.",
  ], notes, "Foundation | Week 1", { assessesSc: 2 });

  T.closingSlide(pres, {
    reflectionPrompt: "Tell your partner one thing you can count to 10.",
    scItems: [
      "Place counters 1-by-1 on a tens frame.",
      "Name how many counters are on the frame.",
      "Describe the missing parts to make 10.",
    ],
    selfAssessment: {
      prompt: "Show me with your fingers.",
      options: ["Got it", "Getting there", "Need help"],
    },
  }, notes);

  const out1 = path.join("output", "_smoke_test_foundation_numeracy", "smoke_test.pptx");
  ensureDir(out1);
  await pres.writeFile({ fileName: out1 });
  console.log("Foundation numeracy smoke test PPTX written.");
}

async function buildY56Literacy() {
  const T = createTheme("literacy", "grade56", weekToVariant(3));
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  T.titleSlide(pres, "Word Study: discontent", "Year 5/6 | Week 3", "Vocabulary focus", "SAY:\n- Welcome.\n\nDO:\n- Project the slide.");

  T.liSlide(pres, "I can use the word discontent in a sentence.", [
    "I can match the word to its meaning.",
    "I can use the word in a sentence.",
    "I can compare two synonyms for discontent.",
  ], "SAY:\n- Read the LI.\n\nDO:\n- Point to each SC.", "Year 5/6", { tieredSc: true });

  // Note: No image so vocabSlide takes legacy single-column path.
  T.vocabSlide(pres, "discontent", "noun",
    "A feeling of not being satisfied.",
    "Her discontent grew as she waited longer.",
    "SAY:\n- Say it with me.\n\nDO:\n- Point to the word.",
    "Year 5/6");

  const placeholderSlide = pres.addSlide();
  T.addTeacherProvidedPlaceholder(placeholderSlide, "quote", { x: 0.5, y: 1.4, w: 9, h: 3.2 });

  const out2 = path.join("output", "_smoke_test_y56_literacy", "smoke_test.pptx");
  ensureDir(out2);
  await pres.writeFile({ fileName: out2 });
  console.log("Y5/6 literacy smoke test PPTX written.");
}

async function buildY12NumeracyManipulatives() {
  const T = createTheme("numeracy", "grade2", weekToVariant(2));
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  const s1 = pres.addSlide();
  T.addTopBar(s1, T.C.PRIMARY);
  T.addTitle(s1, "Manipulative library check");
  T.addTensFrame(s1, 0.5, 1.4, 3.0, 7);
  T.addFiveFrame(s1, 4.0, 1.4, 2.5, 4);
  T.addDotCard(s1, 6.8, 1.4, 1.6, 8);
  T.addNumberTrack(s1, 0.5, 3.0, 7.5, 1, 10, [4, 7]);
  T.addArray(s1, 0.5, 3.7, 3, 4, { cellSize: 0.32 });
  T.addBaseTenBlocks(s1, 4.5, 3.7, 1, 3, 5);
  T.addFractionStripSet(s1, 0.5, 4.6, 8.0, 0.5, [
    { denom: 1, shaded: 1, label: "1 whole" },
    { denom: 2, shaded: 1 },
  ]);

  // Routine icon (async)
  const s2 = pres.addSlide();
  T.addTopBar(s2, T.C.PRIMARY);
  T.addTitle(s2, "Routine icon check");
  await T.addRoutineBadge(s2, "miniWhiteboard", 1.0, 2.0);
  await T.addRoutineBadge(s2, "partnerTalk", 3.0, 2.0);
  await T.addRoutineBadge(s2, "thumbsUp", 5.0, 2.0);
  await T.addRoutineBadge(s2, "exitTicket", 7.0, 2.0);

  const out3 = path.join("output", "_smoke_test_y12_numeracy", "smoke_test.pptx");
  ensureDir(out3);
  await pres.writeFile({ fileName: out3 });
  console.log("Y1/2 manipulatives smoke test PPTX written.");
}

(async () => {
  fs.mkdirSync("output", { recursive: true });
  await buildFoundationNumeracy();
  await buildY56Literacy();
  await buildY12NumeracyManipulatives();
  console.log("\nAll smoke tests passed.");
})().catch((err) => {
  console.error("\nSmoke test FAILED:", err);
  process.exit(1);
});
