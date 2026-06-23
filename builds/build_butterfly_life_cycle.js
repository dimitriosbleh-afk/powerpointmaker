"use strict";

/**
 * The Butterfly Life Cycle - Year 3/4 Science (single lesson)
 *
 * Content: egg -> caterpillar -> chrysalis -> butterfly, taught as a repeating
 * life cycle. Mixed readiness. Core visual is a BUILT labelled cycle diagram
 * (cycleDiagramSlide) - no external images required, per the brief.
 *
 * General-subject explicit-teaching order (megaprompt s32):
 *   Title -> Teacher Resources -> Launch -> LI/SC -> Key Vocabulary
 *   -> I Do (cycle) -> CFU (reveal) -> We Do (name the stages, then check)
 *   -> CFU hinge (reveal) -> You Do (recording sheet) -> Exit Ticket -> Closing
 *
 * Visual fade through the GRR: the cycle diagram is the anchor throughout.
 * I Do shows the full labelled cycle; We Do fades the LABELS (students name the
 * stages from clues) before confirming; You Do is independent on paper.
 *
 * Companion PDFs (resources-session1/):
 *   - Session 1 Life Cycle Recording Sheet (label the cycle + explain + challenge)
 *   - Session 1 Answer Key
 */

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, composeNotes } = require("../themes/factory");
const {
  addResourceSlide,
  makeSessionResource,
  createPdf,
  writePdf,
  addPdfHeader,
  addSectionHeading,
  addBodyText,
  addLinedArea,
  addWriteLine,
  addCycleDiagramPdf,
  addPdfFooter,
} = require("../themes/pdf_helpers");

const VARIANT = 2;
const T = createTheme("science", "grade34", VARIANT);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  cycleDiagramSlide,
  withReveal, addRevealAnswerBar,
} = T;

const FOOTER = "Science | Year 3/4 | The Butterfly Life Cycle";
const OUT_DIR = "output/Butterfly_Life_Cycle_Lesson";
fs.mkdirSync(OUT_DIR, { recursive: true });

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "The Butterfly Life Cycle - Year 3/4 Science";
pres.author = "Year 3/4 Science";

/* Success criteria - plain unlabelled "I can..." list (no tier labels on any
 * student-facing surface). Internal tiers: SC1 foundation, SC2 core (exit
 * ticket target = ordering), SC3 depth (explain the cycle). */
const SC = [
  "I can name the four stages of a butterfly's life cycle.",
  "I can put the four stages in the correct order.",
  "I can explain why we call it a life cycle.",
];
const LI = "We are learning how a butterfly grows and changes through its life cycle.";

/* The four stages, used by both the slide cycle diagrams and the PDFs. */
const STAGE_COLORS = [C.PRIMARY, C.SECONDARY, C.ACCENT, C.SUCCESS];
const STAGES_FULL = [
  { label: "Egg", detail: "Laid on a leaf", color: STAGE_COLORS[0] },
  { label: "Caterpillar", detail: "Eats and grows", color: STAGE_COLORS[1] },
  { label: "Chrysalis", detail: "Hard case forms", color: STAGE_COLORS[2] },
  { label: "Butterfly", detail: "Wings unfold", color: STAGE_COLORS[3] },
];
const STAGES_BLANK = STAGES_FULL.map((s) => ({ label: "", detail: s.detail, color: s.color }));

/* Resource objects (PDFs generated at the end of this script). */
const recordingSheet = makeSessionResource(
  1,
  "Life Cycle Recording Sheet",
  "Label the cycle, explain why it repeats, plus a challenge. Print one per student."
);
const answerKey = makeSessionResource(
  1,
  "Answer Key",
  "Completed cycle and model answers for self-marking and the teacher."
);

/* ─────────────────────────────────────────────────────────────────────────
 *  1. Title
 * ───────────────────────────────────────────────────────────────────────── */

titleSlide(
  pres,
  "The Butterfly Life Cycle",
  "Egg, caterpillar, chrysalis, butterfly - and around again",
  "Year 3/4 Science  |  Mixed readiness  |  One lesson",
  composeNotes({
    say: [
      "Today we are going to follow a butterfly through its whole life, from a tiny egg all the way to a flying butterfly.",
      "The big idea is that it happens in stages, and those stages repeat. We call that a life cycle.",
    ],
    do: [
      "Have the recording sheets printed and mini-whiteboards ready before students arrive.",
      "Keep this deck on the I Do cycle slide so you can flip back to it during practice.",
    ],
    teacherNotes: [
      "Australian Curriculum 2.0 link: Science (Biological sciences) - living things have life cycles and depend on each other and the environment.",
      "No external images are needed. The labelled cycle is built into the slides and the recording sheet.",
    ],
    tag: "[Butterfly Life Cycle | Year 3/4 | Title]",
  }, { requireSay: false, requireDo: false })
);

/* ─────────────────────────────────────────────────────────────────────────
 *  2. Teacher Resources (immediately after title)
 * ───────────────────────────────────────────────────────────────────────── */

addResourceSlide(
  pres,
  {
    resources: [recordingSheet, answerKey],
    studentTools: [
      "Mini-whiteboards and markers (launch, CFU, exit)",
      "Pencils and the recording sheet (You Do)",
    ],
    manipulatives: [
      "Optional: butterfly life cycle model figures or picture cards if your room has them",
    ],
    boardSetup: [
      "Print the recording sheet, one per student",
      "Print a few answer keys for self-marking corners",
      "Have the I Do cycle slide ready to flip back to",
    ],
  },
  T,
  FOOTER,
  composeNotes({
    say: [
      "Quick set-up check before we start: whiteboards out, recording sheets ready, pencils on desks.",
    ],
    do: [
      "Hand out mini-whiteboards as students settle.",
      "Keep recording sheets face down until the You Do so they do not preview the answers.",
    ],
    teacherNotes: [
      "Everything students need to label the cycle is taught on the slides first. The sheet is for the independent stage only.",
    ],
    tag: "[Butterfly Life Cycle | Resources]",
  }, { requireSay: false, requireDo: false })
);

/* ─────────────────────────────────────────────────────────────────────────
 *  3. Launch - prior knowledge to new learning
 * ───────────────────────────────────────────────────────────────────────── */

contentSlide(
  pres,
  "Launch",
  C.SECONDARY,
  "How does a caterpillar become a butterfly?",
  [
    "Have you ever seen a caterpillar or a butterfly? What did it look like?",
    "Turn and tell your partner: how do you think a caterpillar changes into a butterfly?",
  ],
  composeNotes({
    say: [
      "Picture a caterpillar munching on a leaf. Now picture a butterfly floating past a flower. Believe it or not, they can be the same animal at different times in its life.",
      "Some of you may have seen this happen, and if this feels brand new, that is okay - we will build it together.",
      "Turn to your partner and share one idea: how do you think a caterpillar changes into a butterfly?",
    ],
    do: [
      "Give partners about 30 to 45 seconds to talk.",
      "Take two or three ideas from the room. Do not correct yet - just collect their thinking.",
      "Bridge: 'Let us find out the real order it happens in.'",
    ],
    teacherNotes: [
      "This launch connects what students already notice in the garden to today's target. Keep it short and curious.",
    ],
    watchFor: [
      "Students who think the caterpillar and butterfly are two completely different animals - that is a great misconception to surface now and resolve during the I Do.",
    ],
    tag: "[Butterfly Life Cycle | Launch]",
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
  composeNotes({
    say: [
      "Here is what we are learning and how you will know you have got it.",
      "Learning intention: we are learning how a butterfly grows and changes through its life cycle.",
      "By the end you should be able to name the four stages, put them in the right order, and explain why we call it a cycle.",
    ],
    do: [
      "Read the learning intention aloud, then each success criterion.",
      "Tell students the first one is for everyone and we will all reach it together.",
    ],
    teacherNotes: [
      "Internal tiers only (not shown): SC1 naming is the floor everyone reaches; SC2 ordering is the core and is what the exit ticket checks; SC3 explaining the cycle is the stretch.",
    ],
    tag: "[Butterfly Life Cycle | LI and SC]",
  }, { requireSay: false, requireDo: false }),
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
    "life cycle - the stages a living thing goes through, again and again",
    "metamorphosis - a big change in body shape as an animal grows",
    "larva and pupa - the science names for the caterpillar and the chrysalis",
  ],
  composeNotes({
    say: [
      "Three words will help us today. Say each one with me.",
      "Life cycle means the stages a living thing goes through, over and over.",
      "Metamorphosis is our big science word - it means a big change in body shape as an animal grows. A caterpillar does not just get bigger, it changes shape completely.",
      "Scientists also call the caterpillar a larva, and the chrysalis a pupa. You can use either name.",
    ],
    do: [
      "Choral read each word once.",
      "Do a quick action for metamorphosis - hands together then open wide - so the word sticks.",
    ],
    teacherNotes: [
      "Keep this brisk. The words are tools for the cycle diagram coming next, not a spelling list.",
    ],
    watchFor: [
      "Students saying metamorphosis just means getting bigger - clarify it means changing shape, not only growing.",
    ],
    tag: "[Butterfly Life Cycle | Vocabulary]",
  }),
  FOOTER
);

/* ─────────────────────────────────────────────────────────────────────────
 *  6. I Do - the labelled cycle (the model)
 * ───────────────────────────────────────────────────────────────────────── */

cycleDiagramSlide(
  pres,
  "I Do",
  "The Butterfly Life Cycle",
  "Watch and listen",
  [
    "I will name each stage and what happens.",
    "Follow the arrows around the cycle.",
  ],
  "Life Cycle",
  STAGES_FULL,
  composeNotes({
    say: [
      "Let us look at this together. I am going to start at the egg, because every butterfly starts as a tiny egg laid on a leaf.",
      "Watch how I follow the arrows. The egg hatches into a caterpillar. The caterpillar's whole job is to eat leaves and grow bigger and bigger.",
      "When it is big enough, it makes a hard case around itself called a chrysalis. This is where the amazing change, the metamorphosis, happens inside.",
      "Then out comes the butterfly, with wings that unfold and dry. And here is the clever part - the butterfly lays new eggs, so the whole thing starts again. That is why the arrows go in a circle.",
    ],
    do: [
      "Point to each stage as you name it, then trace the arrows all the way round and back to the egg.",
      "Emphasise the return arrow from butterfly to egg - that is what makes it a cycle.",
      "Link back to the launch: 'So the caterpillar and the butterfly are the SAME animal at different stages.'",
    ],
    teacherNotes: [
      "This is the model. Narrate it; do not ask students to read the slide. The diagram is the lesson.",
    ],
    misconceptions: [
      "Misconception: the caterpillar and butterfly are different animals. Why: they look nothing alike. Quick correction: trace the cycle and stress it is one animal changing shape.",
      "Misconception: the cycle has a start and an end. Quick correction: show the loop has no end - eggs lead to butterflies which lead to eggs.",
    ],
    watchFor: [
      "Students who want to jump straight to butterfly - slow them down and keep the order egg, caterpillar, chrysalis, butterfly.",
    ],
    tag: "[Butterfly Life Cycle | I Do]",
  }, { requireSay: false, requireDo: false }),
  FOOTER
);

/* ─────────────────────────────────────────────────────────────────────────
 *  7. CFU 1 - order check (reveal pair)
 * ───────────────────────────────────────────────────────────────────────── */

withReveal(
  () => cfuSlide(
    pres,
    "CFU",
    "Quick check",
    "Show me on your fingers",
    "Which stage comes straight after the egg?",
    composeNotes({
      say: [
        "Quick check. On your fingers or whiteboard, show or write which stage comes straight after the egg.",
      ],
      do: [
        "Give 10 seconds of thinking time, then say 'show me'.",
        "Scan the room before you click to reveal the answer.",
      ],
      cfu: {
        technique: "Show me (fingers or whiteboard)",
        script: [
          "Ask: which stage comes straight after the egg?",
          "Hold up your answer on 'show me'.",
        ],
        scanFor: "Most students answering caterpillar (or larva).",
        proceed: ["If about 80 percent or more are correct, move on to the We Do."],
        pivot: {
          misconception: "Students naming chrysalis or butterfly - they have the right stages but the wrong order.",
          reteach: "Slow-trace the cycle again from the egg, saying each stage out loud and tapping it.",
          recheck: "Re-ask with a partner check: 'tell your partner the stage right after the egg.'",
        },
      },
      teacherNotes: [
        "Use the optional re-teach move in the pivot only if the show-me is below about 80 percent.",
      ],
      tag: "[Butterfly Life Cycle | CFU order]",
    }),
    FOOTER
  ),
  (slide) => {
    addRevealAnswerBar(slide, "The caterpillar (larva)", {
      label: "Answer",
      color: C.SUCCESS,
    });
  }
);

/* ─────────────────────────────────────────────────────────────────────────
 *  8. We Do - name each stage (labels faded), then check
 * ───────────────────────────────────────────────────────────────────────── */

cycleDiagramSlide(
  pres,
  "We Do",
  "Your turn: name each stage",
  "With your partner",
  [
    "Use the clues to name stages 1 to 4.",
    "Write them on your whiteboard.",
  ],
  "Life Cycle",
  STAGES_BLANK,
  composeNotes({
    say: [
      "Now you try. The names are gone, but the clues are still here. With your partner, use each clue to name stages 1 to 4.",
      "Start at stage 1 and work your way around. Write the four names on your whiteboard.",
    ],
    do: [
      "Give partners about a minute. Walk and listen.",
      "Prompt stuck pairs with a clue: 'Stage 2 eats and grows - what is that?'",
      "Tell students to hold up their whiteboards before you advance to the check slide.",
    ],
    enabling: [
      "ENABLING: let students who need it flip back to the I Do cycle in their mind, or quietly point them to the order egg, caterpillar, chrysalis, butterfly.",
    ],
    extending: [
      "EXTENDING: ask early finishers to add what the butterfly does to make the cycle start again.",
    ],
    watchFor: [
      "Pairs writing the right names in the wrong order - remind them to start at stage 1 and follow the arrows.",
    ],
    tag: "[Butterfly Life Cycle | We Do attempt]",
  }, { requireSay: false, requireDo: false }),
  FOOTER
);

cycleDiagramSlide(
  pres,
  "We Do",
  "Let's check",
  "Check together",
  [
    "Did you get the order right?",
    "Egg first, then around we go.",
  ],
  "Life Cycle",
  STAGES_FULL,
  composeNotes({
    say: [
      "Let us check together. Stage 1 is the egg, stage 2 is the caterpillar, stage 3 is the chrysalis, stage 4 is the butterfly.",
      "Tick your whiteboard if you matched each clue to the right stage. If you swapped two, fix them now.",
    ],
    do: [
      "Reveal each name in order, pointing to the clue beside it.",
      "Celebrate correct order, then have students fix any swaps before moving on.",
    ],
    teacherNotes: [
      "This is the answer to the previous slide. Keep it up while students self-correct.",
    ],
    watchFor: [
      "Students who confused chrysalis and caterpillar - re-link the clue 'hard case forms' to chrysalis.",
    ],
    tag: "[Butterfly Life Cycle | We Do check]",
  }, { requireSay: false, requireDo: false }),
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
    "After the butterfly, what happens next in the cycle?",
    composeNotes({
      say: [
        "Here is the big one. After the butterfly stage, what happens next in the cycle?",
        "Turn and tell your partner first, then show me a thumb up when you have an answer.",
      ],
      do: [
        "Give partners 15 seconds, then take one answer before revealing.",
      ],
      cfu: {
        technique: "Turn and tell, then thumbs up",
        script: [
          "Ask: after the butterfly, what happens next in the cycle?",
          "Partners agree on an answer, then thumbs up.",
        ],
        scanFor: "Students saying the butterfly lays eggs and it starts again.",
        proceed: ["If most pairs say the cycle starts again with eggs, students are ready for the You Do."],
        pivot: {
          misconception: "Students saying 'it ends' or 'the butterfly dies and that is it' - they are missing the loop.",
          reteach: "Trace the return arrow from butterfly to egg and ask 'where do the new eggs come from?'",
          recheck: "Re-ask: 'so does the cycle stop, or start again?'",
        },
      },
      teacherNotes: [
        "This hinge checks SC3 (why it is a cycle). Use the optional re-teach only if pairs are below about 80 percent.",
      ],
      tag: "[Butterfly Life Cycle | CFU hinge]",
    }),
    FOOTER
  ),
  (slide) => {
    addRevealAnswerBar(slide, "It lays eggs and the cycle starts again", {
      label: "Answer",
      color: C.SUCCESS,
    });
  }
);

/* ─────────────────────────────────────────────────────────────────────────
 *  10. You Do - independent on the recording sheet (different format)
 * ───────────────────────────────────────────────────────────────────────── */

contentSlide(
  pres,
  "You Do",
  C.SUCCESS,
  "Show what you know on your sheet",
  [
    "First: label all four stages on your recording sheet.",
    "Next: draw the arrow from the butterfly back to the egg.",
    "Then: write one sentence on why it is a life cycle.",
  ],
  composeNotes({
    say: [
      "Now it is just you. On your recording sheet, label all four stages in the cycle.",
      "Then draw the arrow that goes from the butterfly back to the egg - that is the part that makes it a cycle.",
      "Finally, write one sentence to explain why we call it a life cycle.",
    ],
    do: [
      "Hand out the recording sheets now.",
      "Circulate. If a student is stuck on naming, point them to the order, not the answer.",
      "Note one or two strong sentences to share at the close.",
    ],
    enabling: [
      "ENABLING: students who need it may whisper the stages to you first, then write. The challenge box is optional for them.",
    ],
    extending: [
      "EXTENDING: the challenge box asks students to name another animal that changes as it grows and describe its first and last stage.",
    ],
    watchFor: [
      "Students who skip the return arrow - it is the key evidence for SC3, so prompt for it.",
    ],
    tag: "[Butterfly Life Cycle | You Do]",
  }),
  FOOTER
);

/* ─────────────────────────────────────────────────────────────────────────
 *  11. Exit Ticket - assesses SC2 (ordering), no SC label on face
 * ───────────────────────────────────────────────────────────────────────── */

contentSlide(
  pres,
  "Exit Ticket",
  C.ALERT,
  "Before you go",
  [
    "Write the four stages of the butterfly life cycle in order.",
    "You can use words or a quick drawing.",
  ],
  composeNotes({
    say: [
      "Last job before you go. On your whiteboard, write the four stages of the butterfly life cycle in the correct order.",
      "Words or a quick drawing - your choice. Hold it up when you are done.",
    ],
    do: [
      "Scan whiteboards for the correct order: egg, caterpillar, chrysalis, butterfly.",
      "Jot who still needs support so you can pick them up next session.",
    ],
    teacherNotes: [
      "This exit ticket checks SC2 (correct order). The SC target stays in these notes only - it is not shown to students.",
    ],
    watchFor: [
      "Right names, wrong order - note these students; they have SC1 but not yet SC2.",
    ],
    tag: "[Butterfly Life Cycle | Exit Ticket]",
  }),
  FOOTER
);

/* ─────────────────────────────────────────────────────────────────────────
 *  12. Closing reflection
 * ───────────────────────────────────────────────────────────────────────── */

closingSlide(
  pres,
  {
    reflectionPrompt: "Tell your partner the four stages in order, from memory.",
    scItems: SC,
    selfAssessment: {
      prompt: "Thumbs up, sideways, or down: how sure are you of the four stages in order?",
      options: ["Sure", "Getting there", "Need more"],
    },
  },
  composeNotes({
    say: [
      "Let us look back at what we set out to do today.",
      "Turn and tell your partner the four stages in order, from memory: egg, caterpillar, chrysalis, butterfly.",
      "Now show me a thumb - up, sideways, or down - for how sure you are of the order.",
    ],
    do: [
      "Read the three success criteria aloud and have students self-check against them.",
      "Use the thumbs data to decide who to revisit next session.",
    ],
    teacherNotes: [
      "Acknowledge progress - most students can now name and order the stages, which is the core goal.",
    ],
    tag: "[Butterfly Life Cycle | Closing]",
  }, { requireSay: false, requireDo: false })
);

/* ─────────────────────────────────────────────────────────────────────────
 *  Companion PDFs (resources-session1/)
 * ───────────────────────────────────────────────────────────────────────── */

const PDF_STAGES = STAGES_FULL.map((s) => ({ label: s.label, color: s.color }));

function buildRecordingSheet() {
  const doc = createPdf({ title: "Butterfly Life Cycle Recording Sheet" });

  let y = addPdfHeader(doc, "Butterfly Life Cycle", {
    subtitle: "Label the cycle, then explain why it repeats.",
    color: C.PRIMARY,
    showNameDate: true,
  });

  y = addSectionHeading(doc, "Part 1: Label the cycle", y, { color: C.PRIMARY });
  y = addBodyText(doc, "Look at the numbered cycle. Write each stage name on the matching line.", y);
  y = addCycleDiagramPdf(doc, y, PDF_STAGES, {
    centerLabel: "Life Cycle",
    centerColor: C.PRIMARY,
    showStageNames: false,
    numberedLines: false,
    height: 200,
  });
  y = addWriteLine(doc, "1.", y, {});
  y = addWriteLine(doc, "2.", y, {});
  y = addWriteLine(doc, "3.", y, {});
  y = addWriteLine(doc, "4.", y, {});

  y = addSectionHeading(doc, "Part 2: Why is it a life cycle?", y + 4, { color: C.SECONDARY });
  y = addBodyText(doc, "Explain why we call it a cycle and not a straight line.", y);
  y = addLinedArea(doc, y + 4, 2);

  y = addSectionHeading(doc, "Challenge (optional)", y + 6, { color: C.ACCENT });
  y = addBodyText(doc, "Name another animal that changes as it grows. Draw or describe its first stage and its last stage in the box.", y);
  const CONTENT_BOTTOM = 771.89;
  const boxH = Math.min(140, CONTENT_BOTTOM - y - 4);
  if (boxH >= 70) {
    doc.save();
    doc.roundedRect(50, y, 495.28, boxH, 6)
      .lineWidth(0.9).strokeColor("#" + C.MUTED).stroke();
    doc.restore();
  }

  addPdfFooter(doc, FOOTER);
  return writePdf(doc, path.join(OUT_DIR, recordingSheet.fileName));
}

function buildAnswerKey() {
  const doc = createPdf({ title: "Butterfly Life Cycle Answer Key" });

  let y = addPdfHeader(doc, "Butterfly Life Cycle - Answer Key", {
    subtitle: "For the teacher and for self-marking.",
    color: C.PRIMARY,
    showNameDate: false,
  });

  y = addSectionHeading(doc, "Part 1: Label the cycle (answers)", y, { color: C.PRIMARY });
  y = addCycleDiagramPdf(doc, y, PDF_STAGES, {
    centerLabel: "Life Cycle",
    centerColor: C.PRIMARY,
    showStageNames: true,
    numberedLines: false,
  });
  y = addBodyText(doc, "Order: 1. Egg   2. Caterpillar (larva)   3. Chrysalis (pupa)   4. Butterfly (adult).", y);

  y = addSectionHeading(doc, "Part 2: Why is it a life cycle?", y + 4, { color: C.SECONDARY });
  y = addBodyText(doc, "Sample answer: It is a life cycle because the stages repeat. The butterfly lays eggs, and those eggs grow into new caterpillars, so it goes around again instead of stopping.", y);

  y = addSectionHeading(doc, "Challenge (sample answers)", y + 4, { color: C.ACCENT });
  y = addBodyText(doc, "Accept any animal that changes as it grows. For example a frog (egg, tadpole, froglet, frog) or a chicken (egg, chick, hen). Look for a clear first stage and last stage.", y);

  addPdfFooter(doc, FOOTER);
  return writePdf(doc, path.join(OUT_DIR, answerKey.fileName));
}

/* ─────────────────────────────────────────────────────────────────────────
 *  Write everything
 * ───────────────────────────────────────────────────────────────────────── */

const outFile = path.join(OUT_DIR, "Butterfly Life Cycle Year 3-4 Science.pptx");

Promise.all([
  pres.writeFile({ fileName: outFile }),
  buildRecordingSheet(),
  buildAnswerKey(),
]).then(() => {
  console.log("PPTX written to " + outFile);
  console.log("PDF written to " + path.join(OUT_DIR, recordingSheet.fileName));
  console.log("PDF written to " + path.join(OUT_DIR, answerKey.fileName));
}).catch((err) => {
  console.error("Build failed:", err);
  process.exitCode = 1;
});
