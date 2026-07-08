"use strict";

/**
 * The Water Cycle - Year 5/6 Science (single lesson, first in the unit)
 *
 * Content: evaporation -> condensation -> precipitation -> collection, taught
 * as a repeating cycle. 60 minutes, mixed readiness. Core visual is a BUILT
 * labelled cycle diagram (cycleDiagramSlide) - no external images required.
 *
 * General-subject explicit-teaching order (megaprompt s32 / s0a item 23):
 *   Title -> Teacher Resources -> Launch -> LI/SC -> Key Vocabulary
 *   -> I Do (cycle) -> CFU (reveal) -> We Do attempt (labels faded)
 *   -> We Do check -> CFU hinge (reveal) -> You Do (science journal)
 *   -> Exit Ticket -> Closing
 *
 * Visual fade through the GRR: the cycle diagram is the anchor throughout.
 * I Do shows the full labelled cycle; We Do fades the LABELS (students name the
 * parts from clues) before confirming; You Do is independent in journals.
 *
 * You Do is a science-journal task (per the brief), so no core worksheet is
 * generated. One enabling scaffold (form change for students who need it) plus
 * its answer key are the only companion PDFs.
 *
 * Notes use the v11.0 Glance Format (composeGlanceNotes).
 *
 * Companion PDFs (resources-session1/):
 *   - Session 1 Water Cycle Scaffold (label the cycle + word bank + explain)
 *   - Session 1 Answer Key
 */

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  addResourceSlide,
  makeSessionResource,
  createPdf,
  writePdf,
  addPdfHeader,
  addSectionHeading,
  addBodyText,
  addTipBox,
  addLinedArea,
  addWriteLine,
  addCycleDiagramPdf,
  addPdfFooter,
} = require("../themes/pdf_helpers");

// First lesson in the unit -> week 1 -> variant 0. All lessons in this unit
// MUST reuse this variant for theme cohesion.
const T = createTheme("science", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  cycleDiagramSlide,
  withReveal, addRevealAnswerBar,
  composeGlanceNotes,
} = T;

const FOOTER = "Science | Year 5/6 | The Water Cycle";
const OUT_DIR = "output/Water_Cycle_Lesson1";
fs.mkdirSync(OUT_DIR, { recursive: true });

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "The Water Cycle - Year 5/6 Science";
pres.author = "Year 5/6 Science";

/* Success criteria - plain unlabelled "I can..." list (no tier labels on any
 * student-facing surface). Internal tiers: SC1 name (floor everyone reaches),
 * SC2 order and describe the movement (exit ticket target), SC3 explain why it
 * repeats (depth). */
const LI = "We are learning how water moves and changes through the water cycle.";
const SC = [
  "I can name the four parts of the water cycle.",
  "I can put the four parts in order and describe how the water moves.",
  "I can explain why the water cycle repeats and never runs out.",
];

/* The four parts, used by both the slide cycle diagrams and the PDFs.
 * Order is clockwise from the top: evaporation -> condensation ->
 * precipitation -> collection -> back to evaporation. */
const STAGE_COLORS = [C.PRIMARY, C.SECONDARY, C.ACCENT, C.SUCCESS];
const STAGES_FULL = [
  { label: "Evaporation",   detail: "Sun heats the water",   color: STAGE_COLORS[0] },
  { label: "Condensation",  detail: "Vapour cools to cloud", color: STAGE_COLORS[1] },
  { label: "Precipitation", detail: "Rain or snow falls",    color: STAGE_COLORS[2] },
  { label: "Collection",    detail: "Gathers in rivers, sea", color: STAGE_COLORS[3] },
];
const STAGES_BLANK = STAGES_FULL.map((s) => ({ label: "", detail: s.detail, color: s.color }));

/* Resource objects (PDFs generated at the end of this script). */
const scaffold = makeSessionResource(
  1,
  "Water Cycle Scaffold",
  "Label the cycle with a word bank, then explain why it repeats. Print a few for students who need a different entry point."
);
const answerKey = makeSessionResource(
  1,
  "Answer Key",
  "Completed cycle and model answers for the teacher and self-marking."
);

/* ─────────────────────────────────────────────────────────────────────────
 *  1. Title
 * ───────────────────────────────────────────────────────────────────────── */

titleSlide(
  pres,
  "The Water Cycle",
  "Evaporation, condensation, precipitation, collection - and around again",
  "Year 5/6 Science  |  Mixed readiness  |  Lesson 1 of the unit",
  "Today we begin the water cycle. Keep this deck on the I Do cycle slide so you can flip back to it during practice."
);

/* ─────────────────────────────────────────────────────────────────────────
 *  2. Teacher Resources (immediately after title)
 * ───────────────────────────────────────────────────────────────────────── */

addResourceSlide(
  pres,
  {
    resources: [scaffold, answerKey],
    studentTools: [
      "Science journals and pencils (You Do)",
      "Mini-whiteboards and markers (launch, CFU, exit)",
    ],
    manipulatives: [
      "Optional: four water cycle word cards per group (evaporation, condensation, precipitation, collection)",
    ],
    boardSetup: [
      "Print a few scaffold sheets for students who need a different entry point",
      "Print a couple of answer keys for self-marking corners",
      "Have the I Do cycle slide ready to flip back to during practice",
    ],
  },
  T,
  FOOTER,
  "Set up before students arrive: whiteboards out, science journals ready, a few scaffold sheets printed. Everything students need is taught on the slides first."
);

/* ─────────────────────────────────────────────────────────────────────────
 *  3. Launch - prior knowledge to new learning
 * ───────────────────────────────────────────────────────────────────────── */

contentSlide(
  pres,
  "Launch",
  C.SECONDARY,
  "Where does rain come from, and where does it go?",
  [
    "Turn and tell your partner: where do you think rain comes from?",
    "A puddle dries up on a sunny day. Where does that water go?",
  ],
  composeGlanceNotes({
    answer: "open - listen for the idea that rain falls from clouds and water rises, soaks in, or runs to rivers and the sea",
    beats: [
      "ASK: Where does rain come from, and where does it go after it lands? 30 sec, turn and tell. EXPECT: from clouds, then soaks in or runs to rivers and the sea.",
      "COLLECT two or three ideas, do not correct yet. SAY: You already notice water moving around us. Today we follow its whole journey and name each part.",
      "POINT to the puddle question. SAY: A puddle dries up on a hot day - hold onto where you think that water went.",
    ],
    trap: "thinking clouds make brand new water. Fix: name that it is the same water moving and changing, we will prove it on the cycle.",
    prep: "Bridges everyday weather to today's target: the named water cycle. Keep it short and curious. First lesson in the unit.",
    tag: "[Launch | Retention and recall | HITS 2, 6]",
  }),
  FOOTER
);

/* ─────────────────────────────────────────────────────────────────────────
 *  4. Learning Intention & Success Criteria
 * ───────────────────────────────────────────────────────────────────────── */

liSlide(
  pres,
  LI,
  SC,
  composeGlanceNotes({
    beats: [
      "SAY: Here is what we are learning today and how you will know you have got it.",
      "POINT to each success criterion. SAY: The first one is for everyone - we will all name the four parts together.",
    ],
    prep: "Read the LI, then each I can statement. Internal only (not shown): SC1 name, SC2 order and describe the movement (exit target), SC3 explain why it repeats.",
    tag: "[LI and SC | Planning made visible | HITS 1]",
  }),
  FOOTER
);

/* ─────────────────────────────────────────────────────────────────────────
 *  5. Key Vocabulary (after LI/SC)
 * ───────────────────────────────────────────────────────────────────────── */

contentSlide(
  pres,
  "Key Vocabulary",
  C.PRIMARY,
  "Words for today",
  [
    "evaporation - the sun heats water and it rises as an invisible gas (water vapour)",
    "condensation - the vapour cools high up and forms tiny cloud droplets",
    "precipitation - water falls from clouds as rain, hail, sleet or snow",
    "collection - water gathers in oceans, rivers and lakes, ready to start again",
  ],
  composeGlanceNotes({
    beats: [
      "SHOW each word. SAY: Four new science words today. Say each one with me.",
      "SAY: Evaporation is water rising as an invisible gas. Condensation is that gas cooling and coming together into cloud.",
      "SAY: Precipitation is water falling as rain or snow. Collection is water gathering, ready to start again.",
      "ASK: Which word means water rising into the air? Choral response. EXPECT: evaporation.",
    ],
    trap: "mixing up evaporation and condensation. Fix: evaporation goes UP as gas, condensation comes together into cloud - students say each with a hand action.",
    prep: "Brisk. These words are tools for the cycle diagram next, not a spelling list.",
    tag: "[Key Vocabulary | Knowledge and memory | HITS 3, 6]",
  }),
  FOOTER
);

/* ─────────────────────────────────────────────────────────────────────────
 *  6. I Do - the labelled cycle (the model)
 * ───────────────────────────────────────────────────────────────────────── */

cycleDiagramSlide(
  pres,
  "I Do",
  "The Water Cycle",
  "Watch and listen",
  [
    "I will name each part and how the water moves.",
    "Follow the arrows around the cycle.",
  ],
  "Water Cycle",
  STAGES_FULL,
  composeGlanceNotes({
    answer: "the four parts in order: evaporation, condensation, precipitation, collection, then round again",
    beats: [
      "POINT to evaporation at the top. SAY: Let us follow one drop of water. The sun heats the ocean and the water rises as invisible vapour.",
      "MODEL round the arrows. SAY: High up it cools and condenses into clouds. When the clouds are full, precipitation falls as rain or snow.",
      "SAY: The water collects in rivers, lakes and the sea, then the sun heats it again. Watch the arrow go back to the start.",
      "SAY: Same water, changing and moving. That return arrow is why we call it a cycle, not a straight line.",
    ],
    trap: "thinking each part is different water. Fix: trace one drop all the way round, naming it as the same water changing.",
    stretch: "ask where the energy that drives the whole cycle comes from (the sun).",
    help: "give the four word cards and have the student place each beside its arrow as you narrate.",
    prep: "This is the model - narrate it, do not ask students to read the slide. The diagram is the lesson. Center label is Water Cycle.",
    tag: "[I Do | Explicit teaching | HITS 3, 4]",
  }),
  FOOTER
);

/* ─────────────────────────────────────────────────────────────────────────
 *  7. CFU 1 - name the process (reveal pair)
 * ───────────────────────────────────────────────────────────────────────── */

withReveal(
  () => cfuSlide(
    pres,
    "CFU",
    "Quick check",
    "Show me on your whiteboards",
    "Which part turns liquid water into water vapour that rises?",
    composeGlanceNotes({
      answer: "Evaporation",
      beats: [
        "ASK: Which part turns liquid water into water vapour that rises? 10 sec, boards up. EXPECT: evaporation.",
        "SCAN boards, back row first. 80%+ -> reveal and move to the We Do. Less -> retrace evaporation on the cycle with a rising hand, re-ask.",
        "REVEAL after boards scanned. SAY: Tick yours, fix it if you need to.",
      ],
      trap: "writing condensation. Fix: condensation is cooling DOWN into cloud - point to both arrows, student redoes.",
      prep: "Checks SC1 naming before guided practice. Show me on whiteboards.",
      tag: "[CFU | Supported application | HITS 7, 8]",
    }),
    FOOTER
  ),
  (slide) => {
    addRevealAnswerBar(slide, "Evaporation", {
      label: "Answer",
      color: C.SUCCESS,
    });
  }
);

/* ─────────────────────────────────────────────────────────────────────────
 *  8. We Do - name each part (labels faded), then check
 * ───────────────────────────────────────────────────────────────────────── */

cycleDiagramSlide(
  pres,
  "We Do",
  "Your turn: name each part",
  "With your partner",
  [
    "Use the clues to name parts 1 to 4.",
    "Write them on your whiteboard.",
  ],
  "Water Cycle",
  STAGES_BLANK,
  composeGlanceNotes({
    answer: "1 evaporation, 2 condensation, 3 precipitation, 4 collection",
    beats: [
      "SAY: Your turn. The names are gone but the clues stay. With your partner, name parts 1 to 4.",
      "TIME about a minute. CIRCULATE and listen. Prompt stuck pairs: part 1 is the sun heating water - what is that called?",
      "SAY: Write all four on your whiteboard, then hold it up before we check.",
    ],
    trap: "right names, wrong order. Fix: start at the top arrow and go clockwise, student re-numbers.",
    stretch: "add what drives the cycle and where it could speed up (more heat, more evaporation).",
    help: "give the four word cards to place on the numbers instead of writing.",
    prep: "Labels faded, clues and diagram stay - the visual is still the thing they reason about. Check slide follows.",
    tag: "[We Do | Supported application | HITS 5, 10]",
  }),
  FOOTER
);

cycleDiagramSlide(
  pres,
  "We Do",
  "Let's check",
  "Check together",
  [
    "Did you get the order right?",
    "Start at evaporation and follow the arrows.",
  ],
  "Water Cycle",
  STAGES_FULL,
  composeGlanceNotes({
    answer: "1 evaporation, 2 condensation, 3 precipitation, 4 collection",
    beats: [
      "REVEAL each name in order, pointing to its clue. SAY: Part 1 evaporation, part 2 condensation, part 3 precipitation, part 4 collection.",
      "SAY: Tick your whiteboard if you matched each clue. If you swapped two, fix them now.",
    ],
    trap: "confusing precipitation and collection. Fix: precipitation FALLS, collection GATHERS - re-link each clue, student fixes.",
    prep: "This is the answer to the previous slide. Keep it up while students self-correct.",
    tag: "[We Do | Supported application | HITS 8]",
  }),
  FOOTER
);

/* ─────────────────────────────────────────────────────────────────────────
 *  9. CFU 2 - hinge on the cycle idea (reveal pair)
 * ───────────────────────────────────────────────────────────────────────── */

withReveal(
  () => cfuSlide(
    pres,
    "CFU",
    "Hinge question",
    "Turn and tell, then show me",
    "What makes this a cycle and not a straight line?",
    composeGlanceNotes({
      answer: "the water collects and evaporates again, so it repeats and never runs out",
      beats: [
        "ASK: What makes this a cycle and not a straight line? 15 sec, turn and tell, then thumbs up. EXPECT: the water goes round and starts again.",
        "SCAN thumbs and take one shared answer. 80%+ -> move to the You Do. Less -> trace the return arrow from collection back to evaporation, re-ask.",
        "REVEAL after you take one answer. SAY: The same water is used again and again.",
      ],
      trap: "saying it just stops after the rain. Fix: point to collection then the arrow back up, student explains where the water goes next.",
      prep: "Hinge checks SC3 (why it repeats). Turn and tell then thumbs up.",
      tag: "[CFU | Supported application | HITS 7, 9]",
    }),
    FOOTER
  ),
  (slide) => {
    addRevealAnswerBar(slide, "It collects, then evaporates again - the cycle repeats", {
      label: "Answer",
      color: C.SUCCESS,
    });
  }
);

/* ─────────────────────────────────────────────────────────────────────────
 *  10. You Do - independent in the science journal (different format)
 * ───────────────────────────────────────────────────────────────────────── */

contentSlide(
  pres,
  "You Do",
  C.SUCCESS,
  "Show what you know in your journal",
  [
    "First: draw the water cycle and label all four parts.",
    "Next: add arrows to show the water moving around.",
    "Then: write one sentence on why it is a cycle.",
  ],
  composeGlanceNotes({
    answer: "open - a labelled cycle in the right order, arrows showing movement, and a sentence on why it repeats",
    beats: [
      "SAY: Now it is just you, in your science journal. Draw the water cycle and label all four parts.",
      "SAY: Add arrows to show the water moving, then write one sentence on why it is a cycle.",
      "CIRCULATE. If a student is stuck on the names, point them to the order, not the answer. Offer the scaffold sheet if it helps.",
    ],
    trap: "drawing the parts with no arrows. Fix: prompt for the arrows - they show the water moving and make the loop.",
    stretch: "label where the sun's energy powers the cycle, or add groundwater soaking in at the collection stage.",
    help: "give the Water Cycle Scaffold - diagram pre-drawn, word bank, and a sentence to finish.",
    prep: "Journal task, a different form from the We Do. Note one or two strong journals to share at the close.",
    tag: "[You Do | Mastery and application | HITS 4, 10]",
  }),
  FOOTER
);

/* ─────────────────────────────────────────────────────────────────────────
 *  11. Exit Ticket - assesses SC2 (order), no SC label on face
 * ───────────────────────────────────────────────────────────────────────── */

contentSlide(
  pres,
  "Exit Ticket",
  C.ALERT,
  "Before you go",
  [
    "Write the four parts of the water cycle in order.",
    "Start with evaporation.",
  ],
  composeGlanceNotes({
    answer: "evaporation, condensation, precipitation, collection",
    beats: [
      "SAY: Last job before you go. On your whiteboard, write the four parts of the water cycle in order, starting with evaporation.",
      "SCAN whiteboards for the correct order. Jot who still needs support so you can pick them up next session.",
    ],
    trap: "right words, wrong order. Fix: note these students - they have naming (SC1) but not yet the sequence (SC2).",
    prep: "Checks SC2 (order and movement). The SC target stays in these notes only, not on the slide.",
    tag: "[Exit Ticket | Mastery and application | HITS 1, 8]",
  }),
  FOOTER
);

/* ─────────────────────────────────────────────────────────────────────────
 *  12. Closing reflection
 * ───────────────────────────────────────────────────────────────────────── */

closingSlide(
  pres,
  {
    reflectionPrompt: "Tell your partner the four parts of the water cycle in order, from memory.",
    scItems: SC,
    selfAssessment: {
      prompt: "Thumbs up, sideways, or down: how sure are you of the four parts in order?",
      options: ["Sure", "Getting there", "Need more"],
    },
  },
  composeGlanceNotes({
    beats: [
      "SAY: Let us look back at what we set out to do today.",
      "ASK: Tell your partner the four parts in order, from memory. EXPECT: evaporation, condensation, precipitation, collection.",
      "SAY: Now show me a thumb - up, sideways or down - for how sure you are of the order.",
    ],
    prep: "Read the three I can statements and have students self-check. Use the thumbs data to plan who to revisit. Acknowledge progress - most can now name and order the cycle.",
    tag: "[Closing | Retention and recall | HITS 1, 9]",
  })
);

/* ─────────────────────────────────────────────────────────────────────────
 *  Companion PDFs (resources-session1/)
 * ───────────────────────────────────────────────────────────────────────── */

const PDF_STAGES = STAGES_FULL.map((s) => ({ label: s.label, color: s.color }));

function buildScaffold() {
  const doc = createPdf({ title: "Water Cycle Scaffold" });

  let y = addPdfHeader(doc, "The Water Cycle", {
    subtitle: "Label the cycle, then explain why it repeats.",
    color: C.PRIMARY,
    showNameDate: true,
  });

  y = addSectionHeading(doc, "Part 1: Label the cycle", y, { color: C.PRIMARY });
  y = addTipBox(doc, "Word bank: evaporation, condensation, precipitation, collection", y, { color: C.SECONDARY });
  y = addBodyText(doc, "Look at the numbered cycle. Write each part on the matching line. Start at the top and go clockwise.", y);
  y = addCycleDiagramPdf(doc, y, PDF_STAGES, {
    centerLabel: "Water Cycle",
    centerColor: C.PRIMARY,
    showStageNames: false,
    numberedLines: false,
    height: 200,
  });
  y = addWriteLine(doc, "1.", y, {});
  y = addWriteLine(doc, "2.", y, {});
  y = addWriteLine(doc, "3.", y, {});
  y = addWriteLine(doc, "4.", y, {});

  y = addSectionHeading(doc, "Part 2: Why is it a cycle?", y + 4, { color: C.SECONDARY });
  y = addBodyText(doc, "Explain why the water cycle repeats and never runs out.", y);
  y = addLinedArea(doc, y + 4, 2);

  y = addSectionHeading(doc, "Challenge (optional)", y + 6, { color: C.ACCENT });
  y = addBodyText(doc, "Where is the water cycle happening near your home right now? Write one example.", y);
  y = addLinedArea(doc, y + 4, 1);

  addPdfFooter(doc, FOOTER);
  return writePdf(doc, path.join(OUT_DIR, scaffold.fileName));
}

function buildAnswerKey() {
  const doc = createPdf({ title: "Water Cycle Answer Key" });

  let y = addPdfHeader(doc, "The Water Cycle - Answer Key", {
    subtitle: "For the teacher and for self-marking.",
    color: C.PRIMARY,
    showNameDate: false,
  });

  y = addSectionHeading(doc, "Part 1: Label the cycle (answers)", y, { color: C.PRIMARY });
  y = addCycleDiagramPdf(doc, y, PDF_STAGES, {
    centerLabel: "Water Cycle",
    centerColor: C.PRIMARY,
    showStageNames: true,
    numberedLines: false,
  });
  y = addBodyText(doc, "Order (clockwise from the top): 1. Evaporation   2. Condensation   3. Precipitation   4. Collection.", y);

  y = addSectionHeading(doc, "Part 2: Why is it a cycle?", y + 4, { color: C.SECONDARY });
  y = addBodyText(doc, "Sample answer: It is a cycle because the same water keeps moving around. The water collects in rivers and the sea, then the sun heats it and it evaporates again, so it repeats instead of stopping.", y);

  y = addSectionHeading(doc, "Challenge (sample answers)", y + 4, { color: C.ACCENT });
  y = addBodyText(doc, "Accept any real example: puddles drying after rain, dew on the grass, clouds forming, rain filling a creek. Look for water changing state or moving between the ground and the sky.", y);

  addPdfFooter(doc, FOOTER);
  return writePdf(doc, path.join(OUT_DIR, answerKey.fileName));
}

/* ─────────────────────────────────────────────────────────────────────────
 *  Write everything
 * ───────────────────────────────────────────────────────────────────────── */

const outFile = path.join(OUT_DIR, "The Water Cycle Year 5-6 Science.pptx");

Promise.all([
  pres.writeFile({ fileName: outFile }),
  buildScaffold(),
  buildAnswerKey(),
]).then(() => {
  console.log("PPTX written to " + outFile);
  console.log("PDF written to " + path.join(OUT_DIR, scaffold.fileName));
  console.log("PDF written to " + path.join(OUT_DIR, answerKey.fileName));
}).catch((err) => {
  console.error("Build failed:", err);
  process.exitCode = 1;
});
