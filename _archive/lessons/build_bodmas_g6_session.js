"use strict";

// BODMAS - order of operations. Grade 6 Numeracy, single session.
//
// Scope (megaprompt section 8): BODMAS is unit-sized. This session teaches
// applying the order to expressions with brackets and two or three
// operations, with the crux being that Divide/Multiply share ONE tier worked
// left to right, and Add/Subtract share ONE tier worked left to right. That
// is the misconception the BODMAS mnemonic itself creates, and it is the
// mathematically important correction for Year 6.
//
// Unit anchor (section 79): the four-tier BODMAS ladder plus the anchor
// phrase "Same tier, left to right." Restated in every I Do.
//
// "BODMAS" is the user's own word for this and is locked (section 5b).
// It is used throughout in preference to PEMDAS or "order of operations".
//
// Reveals use clickBuild (section 20b): one slide, one element per click.
// No duplicate-slide reveal pairs anywhere in this deck.

const path = require("path");
const fs = require("fs");
const pptxgen = require("pptxgenjs");
const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf,
  addPdfHeader, addSectionHeading, addBodyText, addTipBox,
  addPdfFooter, addLinedArea,
  addResourceSlide, makeSessionResource, getSessionResourceFolder,
} = require("../themes/pdf_helpers");

// Variant 2: numeracy/grade56 ACCENT fails WCAG AA on white for variants
// 4 and 5, and this deck puts text on ACCENT fills.
const T = createTheme("numeracy", "grade56", 2);
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, cfuSlide, closingSlide, exitTicketSlide,
  dailyReviewSlide, fluencySlide, workedExSlide,
  addCard, addTextOnShape, addFooter, addChipRow, addArray,
  clickBuild, composeGlanceNotes, runSlideDiagnostics,
  CONTENT_TOP, SAFE_BOTTOM,
} = T;

const UNIT = "BODMAS_G6_Order_Of_Operations";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "BODMAS Order of Operations.pptx";
const FOOTER = "BODMAS | Order of Operations | Year 6 Numeracy";
const SESSION = 1;
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

const WORKSHEET = makeSessionResource(
  SESSION,
  "BODMAS Practice",
  "Ladder reminder, four worked problems with space, and a create-your-own challenge."
);
const ANSWER_KEY = makeSessionResource(
  SESSION,
  "BODMAS Answer Key",
  "Worked steps for every worksheet item, not just final answers."
);
const RESOURCE_ITEMS = [WORKSHEET, ANSWER_KEY];

fs.mkdirSync(RES_DIR, { recursive: true });

/* ══════════════════════════════════════════════════════════════════════
 * Shared visual: the BODMAS ladder (the unit anchor)
 * ══════════════════════════════════════════════════════════════════════ */

const TIERS = [
  { key: "B", label: "Brackets first", color: () => C.PRIMARY },
  { key: "O", label: "Orders - powers like 3 squared", color: () => C.SECONDARY },
  { key: "DM", label: "Divide and Multiply - left to right", color: () => C.ACCENT },
  { key: "AS", label: "Add and Subtract - left to right", color: () => C.ALERT },
];

// Draw one ladder tier. Returns nothing; callers wrap it in a clickBuild step
// so the ladder assembles one rung per click.
function drawTier(s, i, x, y, w, h) {
  const t = TIERS[i];
  const keyW = 1.15;
  addTextOnShape(s, t.key, {
    x, y, w: keyW, h, rectRadius: 0.08,
    fill: { color: t.color() },
  }, {
    fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  addCard(s, x + keyW + 0.12, y, w - keyW - 0.12, h, { fill: C.WHITE, strip: t.color() });
  s.addText(t.label, {
    x: x + keyW + 0.36, y, w: w - keyW - 0.55, h,
    fontSize: 17, fontFace: FONT_B, color: C.CHARCOAL,
    valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
  });
}

/* ══════════════════════════════════════════════════════════════════════
 * Teacher notes (Glance Format v12.3, authored - zero advisories)
 * ══════════════════════════════════════════════════════════════════════ */

const NOTES_TITLE =
  "Display as students arrive. Have mini-whiteboards and markers out before you start.";

const NOTES_RESOURCES = composeGlanceNotes({
  beats: [
    "PREP: print the BODMAS Practice sheet, one per student.",
    "SHOW the ladder poster or keep slide 7 handy to flip back to.",
  ],
  prep: [
    "Materials: mini-whiteboards, markers, workbooks. Answer key is teacher-only.",
    "Anchor: the four-tier ladder and the phrase Same tier, left to right.",
  ],
  tag: "[Setup | Planning | HITS 1]",
});

const NOTES_DR = composeGlanceNotes({
  answer: "48, and 5 rows of 6",
  beats: [
    "SAY: Two quick ones from last week to warm up your number facts.",
    [
      "ASK: What is 6 times 8, and how many rows are in the array?",
      "20 sec. Cue: Write it... Chin it... Show me.",
      "EXPECT: 48, and 5 rows.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> click through the answers and move on.",
      "Less -> count the array rows together, re-ask.",
    ],
  ],
  trap: "counting array columns as rows. Fix: run a finger down one column, student recounts.",
  prep: "Prior learning only. Multiplication facts feed today's Divide and Multiply tier. Whole block under 4 minutes.",
  tag: "[Stage 1 | Daily Review | Retention and recall | HITS 6]",
});

const NOTES_FLUENCY = composeGlanceNotes({
  answer: "9, 8, 7",
  beats: [
    "SAY: Division facts, fast as you can. No working, just the answer.",
    [
      "ASK: 72 divided by 8, then 56 divided by 7, then 63 divided by 9?",
      "15 sec. Cue: Write it... Chin it... Show me.",
      "EXPECT: 9, 8, 7.",
    ],
    "REVEAL after boards are up. SAY: Tick it or fix it.",
  ],
  prep: "Automaticity only, not new teaching. Division is the tier students slip on today. Whole block under 3 minutes.",
  tag: "[Stage 1 | Fluency | Retention and recall | HITS 6]",
});

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "12. Mia is right - the times is done before the add.",
  beats: [
    "SHOW both answers. SAY: Two students, same question, different answers.",
    [
      "ASK: Who is right, Mia or Sam? 30 sec, turn and tell.",
      "EXPECT: Mia, because you multiply before you add.",
      "ACCEPT: Mia, without a reason yet.",
    ],
    [
      "SCAN the room as pairs talk.",
      "80%+ -> reveal Mia and move to the intention.",
      "Less -> leave it open, say we will settle it today, re-ask after the ladder.",
    ],
  ],
  trap: [
    "reading strictly left to right, so 6 plus 2 first.",
    "Fix: name that operations have a rank, prove it after the ladder.",
  ],
  prep: "Problem-first shape: the disagreement creates the need for a rule before any rule is given.",
  tag: "[Launch | Attention and knowledge | HITS 2, 7]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    "POINT to the intention. SAY: Today we settle the order once and for all.",
    "SAY: Read the I can statements with me. Choral read.",
    "SAY: The first one everyone gets today - name which part goes first.",
  ],
  prep: "SC1 achievable by all. SC2 is the core target and the exit ticket. SC3 stretches to explaining the tier rule.",
  tag: "[LI and SC | Planning made visible | HITS 1]",
});

const NOTES_LADDER = composeGlanceNotes({
  answer: "four tiers, and the bottom two go left to right",
  beats: [
    "BUILD one rung per click. SAY: BODMAS is four tiers, not six steps.",
    "POINT to D and M. SAY: These share one tier. Neither beats the other.",
    [
      "SAY: Same tier, left to right. Say it with me.",
      "Everyone, together, on three.",
    ],
    [
      "ASK: In 20 divided by 4 times 2, which comes first?",
      "15 sec, boards up.",
      "EXPECT: the divide, it is further left.",
    ],
  ],
  trap: [
    "reading BODMAS as divide always beating multiply.",
    "Fix: cover the M, show the tier holds both, student re-says it.",
  ],
  stretch: "write an expression where multiply first gives a different answer.",
  help: "give the ladder card to hold and point to.",
  prep: "The anchor for the whole session. Every I Do restates Same tier, left to right in these exact words.",
  tag: "[Stage 2 | I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_IDO1 = composeGlanceNotes({
  answer: "12",
  beats: [
    "SAY: Back to Mia and Sam. Watch me use the ladder.",
    [
      "MODEL down the ladder. SAY: No brackets, no orders.",
      "SAY: 2 times 3 is 6. The line becomes 6 plus 6.",
    ],
    "SAY: Now the bottom tier. 6 plus 6 is 12. Mia was right.",
    [
      "ASK: Why was 24 wrong? 15 sec, turn and tell.",
      "EXPECT: they added first instead of multiplying.",
    ],
  ],
  trap: [
    "adding left to right regardless of tier.",
    "Fix: point up the ladder, student redoes the first step.",
  ],
  stretch: "change one number so the answer becomes 20.",
  help: "ladder card, student ticks each tier as it is checked.",
  prep: "First model. Anchor restated: Same tier, left to right. Rewrite the whole line each step, never just the bit that changed.",
  tag: "[Stage 2 | I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_IDO2 = composeGlanceNotes({
  answer: "10",
  beats: [
    "SAY: This one is the trap the mnemonic sets. Watch carefully.",
    [
      "MODEL left to right. SAY: Divide and multiply share a tier.",
      "SAY: 20 divided by 4 is 5. Then 5 times 2 is 10.",
    ],
    [
      "SHOW the wrong path. SAY: Times first gives 20 divided by 8.",
      "SAY: That is 2 point 5, and it is wrong.",
    ],
    [
      "ASK: Say the rule back to me. 10 sec, choral.",
      "EXPECT: same tier, left to right.",
    ],
  ],
  trap: [
    "doing multiply first because M is later in BODMAS.",
    "Fix: cover the M, re-read the tier, student redoes it.",
  ],
  stretch: "does the same trap exist for add and subtract?",
  help: "underline the leftmost operation before starting.",
  prep: "The decision-grade concept of the lesson. Anchor restated in the same words. Do not rush this slide.",
  tag: "[Stage 2 | I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B - 11",
  beats: [
    [
      "ASK: Which is correct for 30 divided by 5 times 2 minus 1?",
      "30 sec. Cue: Write A, B or C... chin it... show me.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: How do you know? Then reveal.",
      "Less -> rebuild it on the ladder tier by tier, re-ask.",
    ],
    "REVEAL after every board is up. SAY: Tick it or fix it.",
  ],
  trap: [
    "A means multiply was done before divide. C means minus went first.",
    "Fix: name the tier, student redoes step one.",
  ],
  stretch: "write a fourth option that would also be wrong.",
  help: "ladder card, work one tier at a time with the student.",
  prep: "The hinge of the lesson and the release point into guided practice.",
  tag: "[Stage 2 | CFU hinge | Supported application | HITS 7, 8]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "23",
  beats: [
    "SAY: Your turn, with me. Brackets are in play now.",
    [
      "ASK: What is inside the brackets? 15 sec, boards up.",
      "EXPECT: 5 plus 3 is 8.",
    ],
    [
      "ASK: Now the whole line. 45 sec, boards up.",
      "EXPECT: 23.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> reveal, then cold call one board: Convince us.",
      "Less -> rewrite the line after the brackets step, re-ask.",
    ],
  ],
  trap: [
    "solving the brackets then adding left to right.",
    "Fix: point to the times tier, student redoes it.",
  ],
  stretch: "move the brackets so the answer changes, and say it.",
  help: "pre-solve the bracket, student continues from there.",
  prep: "Guided practice with brackets added. Fading from the I Do: students do each tier, teacher confirms.",
  tag: "[Stage 3 | We Do | Supported application | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "worksheet items: 11, 7, 30, 16",
  beats: [
    "SAY: Now on your own. Different numbers, same ladder.",
    "SAY: Rewrite the whole line at every step. Do not skip working.",
    [
      "CIRCULATE. Check the first step only on the first pass.",
      "If the first step is right, the rest usually follows.",
    ],
  ],
  trap: "skipping working and guessing. Fix: cover the answer, student writes tier by tier.",
  stretch: "worksheet Section 3 - write an expression for a partner with answer 12.",
  help: "worksheet Section 1 rebuilds the ladder before any problem.",
  prep: "Different content from the We Do: no shared numbers or contexts. Note two strong books for the close.",
  tag: "[Stage 4 | You Do | Mastery and application | HITS 4, 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "9",
  beats: [
    "SAY: One question, on your own, in your book. This one I am collecting.",
    [
      "ASK: What is 3 plus 24 divided by 4?",
      "60 sec. Cue: quick write in your book, working shown.",
      "EXPECT: 9, with the divide done first.",
    ],
    "COLLECT books or scan as students hold them up.",
  ],
  trap: [
    "adding first to get 6 point 75.",
    "Fix: mark the tier used first, reteach that pair tomorrow.",
  ],
  prep: "Assesses SC2. Sort into secure, first-step-wrong, and no-working piles to plan tomorrow.",
  tag: "[Stage 5 | Exit Ticket | Mastery and application | HITS 1, 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "same tier, left to right",
  beats: [
    "POINT to each criterion. SAY: Thumbs up, sideways or down for each one.",
    [
      "ASK: Say the anchor phrase one last time.",
      "5 sec, choral. Cue: Everyone, together, on three.",
      "EXPECT: same tier, left to right.",
    ],
    "SAY: You settled an argument with a rule today. That is what maths is for.",
  ],
  prep: "Self-assessment against the criteria. Any sideways or down on the second criterion joins tomorrow's small group.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

/* ══════════════════════════════════════════════════════════════════════
 * Deck
 * ══════════════════════════════════════════════════════════════════════ */

function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  /* 1. Title */
  titleSlide(
    pres,
    "BODMAS",
    "Getting the order right, every time",
    "Year 6 Numeracy  |  Order of Operations",
    NOTES_TITLE
  );

  /* 2. Teacher Resources - immediately after the title (section 0a item 19) */
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  /* 3. Daily Review - answers click-revealed on the same slide */
  const sDR = dailyReviewSlide(
    pres,
    "Daily Review",
    ["6 x 8 =", "How many rows?"],
    NOTES_DR,
    FOOTER,
    (s, lg) => {
      addArray(s, lg.rightX + 0.35, lg.panelTopPadded + 0.30, 5, 6);
    }
  );
  clickBuild(sDR, [
    () => { T.addRevealAnswerBar(sDR, ["6 x 8 = 48", "5 rows of 6"], { y: 4.25, h: 0.7, fontSize: 24 }); },
  ]);
  runSlideDiagnostics(sDR, pres);

  /* 4. Fluency - answers click-revealed */
  const sFl = fluencySlide(
    pres,
    "Fluency: Division facts",
    ["72 ÷ 8", "56 ÷ 7", "63 ÷ 9"],
    NOTES_FLUENCY,
    FOOTER
  );
  clickBuild(sFl, [
    () => { T.addRevealAnswerBar(sFl, ["9", "8", "7"], { y: 4.35, h: 0.62, fontSize: 26 }); },
  ]);

  /* 5. Launch - two students disagree */
  const sLaunch = pres.addSlide();
  T.addTopBar(sLaunch, C.PRIMARY);
  T.addBadge(sLaunch, "Launch");
  T.addTitle(sLaunch, "Who is right?");
  addCard(sLaunch, 0.5, CONTENT_TOP, 9, 0.95, { fill: C.WHITE, strip: C.PRIMARY });
  sLaunch.addText("6 + 2 × 3", {
    x: 0.7, y: CONTENT_TOP + 0.05, w: 8.6, h: 0.85,
    fontSize: 40, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  addTextOnShape(sLaunch, "Mia says 12", {
    x: 0.7, y: 2.55, w: 4.0, h: 0.95, rectRadius: 0.1,
    fill: { color: C.SECONDARY },
  }, {
    fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  addTextOnShape(sLaunch, "Sam says 24", {
    x: 5.3, y: 2.55, w: 4.0, h: 0.95, rectRadius: 0.1,
    fill: { color: C.ALERT },
  }, {
    fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  sLaunch.addText("Turn and tell your partner who you agree with, and why.", {
    x: 0.5, y: 3.75, w: 9, h: 0.5,
    fontSize: 17, fontFace: FONT_B, color: C.MUTED,
    align: "center", valign: "middle", margin: 0,
  });
  addFooter(sLaunch, FOOTER);
  sLaunch.addNotes(NOTES_LAUNCH);
  clickBuild(sLaunch, [
    () => {
      addTextOnShape(sLaunch, "Mia is right. We will prove it.", {
        x: 0.5, y: 4.35, w: 9, h: 0.7, rectRadius: 0.1,
        fill: { color: C.SUCCESS },
      }, {
        fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    },
  ]);
  runSlideDiagnostics(sLaunch, pres);

  /* 6. LI and SC */
  liSlide(
    pres,
    "We are learning to work out an expression in the right order using BODMAS.",
    [
      "I can name which part of an expression to work out first.",
      "I can solve an expression with brackets and two or more operations.",
      "I can explain why divide and multiply are worked left to right.",
    ],
    NOTES_LI,
    FOOTER
  );

  /* 7. The BODMAS ladder - the unit anchor, built one rung per click */
  const sLadder = pres.addSlide();
  T.addTopBar(sLadder, C.PRIMARY);
  T.addBadge(sLadder, "The Ladder");
  T.addTitle(sLadder, "BODMAS is four tiers, not six steps");
  const tierH = 0.62;
  const tierGap = 0.10;
  const tierTop = CONTENT_TOP + 0.10;
  addFooter(sLadder, FOOTER);
  sLadder.addNotes(NOTES_LADDER);
  clickBuild(sLadder, TIERS.map((_, i) => () => {
    drawTier(sLadder, i, 0.7, tierTop + i * (tierH + tierGap), 8.6, tierH);
  }).concat([
    () => {
      addTextOnShape(sLadder, "Same tier, left to right", {
        x: 0.7, y: tierTop + 4 * (tierH + tierGap) + 0.04, w: 8.6, h: 0.55,
        rectRadius: 0.1, fill: { color: C.SUCCESS },
      }, {
        fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    },
  ]));
  runSlideDiagnostics(sLadder, pres);

  /* 8. I Do 1 - settle the launch, built step by step */
  buildWorkedStep(
    pres,
    "I Do",
    "Settling the argument",
    "6 + 2 × 3",
    [
      { text: "No brackets, no orders. Drop to Divide and Multiply.", color: C.PRIMARY },
      { text: "2 × 3 = 6, so the line becomes 6 + 6", color: C.ACCENT },
    ],
    "= 12",
    NOTES_IDO1
  );

  /* 9. I Do 2 - the equal-rank trap */
  buildWorkedStep(
    pres,
    "I Do",
    "The trap the mnemonic sets",
    "20 ÷ 4 × 2",
    [
      { text: "Divide and Multiply share a tier. Read left to right.", color: C.PRIMARY },
      { text: "20 ÷ 4 = 5, so the line becomes 5 × 2", color: C.ACCENT },
    ],
    "= 10",
    NOTES_IDO2,
    "Times first would give 20 ÷ 8 = 2.5. Wrong."
  );

  /* 10. CFU hinge - built directly: cfuSlide's question box is density-aware
   * and expands to hero size, which leaves no room for an options row. */
  const sCfu = pres.addSlide();
  T.addTopBar(sCfu, C.ALERT);
  T.addBadge(sCfu, "CFU", { color: C.ALERT });
  T.addTitle(sCfu, "Which one is correct?", { color: C.ALERT });
  addTextOnShape(sCfu, "Show Me Boards", {
    x: 0.5, y: CONTENT_TOP, w: 2.6, h: 0.42, rectRadius: 0.08,
    fill: { color: C.ALERT },
  }, {
    fontSize: 14, fontFace: FONT_B, color: T.getContrastColor(C.ALERT), bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  addCard(sCfu, 0.5, 1.85, 9, 1.0, { fill: C.WHITE, strip: C.ALERT });
  sCfu.addText("30 ÷ 5 × 2 - 1", {
    x: 0.7, y: 1.9, w: 8.6, h: 0.9,
    fontSize: 36, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  // Every distractor maps to a real misconception (section 37):
  // A = multiply before divide, C = subtract before the times tier.
  addChipRow(sCfu, 0.7, 3.05, 8.6, ["A = 2", "B = 11", "C = 6"]);
  addFooter(sCfu, FOOTER);
  sCfu.addNotes(NOTES_CFU);
  clickBuild(sCfu, [
    () => { T.addRevealAnswerBar(sCfu, ["B = 11"], { y: 3.9, h: 0.68, fontSize: 26 }); },
  ]);
  runSlideDiagnostics(sCfu, pres);

  /* 11. We Do - brackets added, different numbers from the I Do */
  const sWe = workedExSlide(
    pres,
    3,
    "We Do",
    "Now with brackets",
    [
      "Work out the brackets first.",
      "Then Divide and Multiply, left to right.",
      "Then Add and Subtract, left to right.",
      "Rewrite the whole line each step.",
    ],
    NOTES_WEDO,
    FOOTER,
    (s, lg) => {
      addCard(s, lg.rightX, lg.panelTopPadded, lg.rightW, 0.9, { fill: C.WHITE, strip: C.PRIMARY });
      s.addText("(5 + 3) × 2 + 7", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.05, w: lg.rightW - 0.3, h: 0.8,
        fontSize: 26, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );
  clickBuild(sWe, [
    () => {
      addTextOnShape(sWe, "8 × 2 + 7", {
        x: 5.3, y: 2.45, w: 4.2, h: 0.6, rectRadius: 0.08,
        fill: { color: C.ACCENT },
      }, {
        fontSize: 22, fontFace: FONT_H, color: T.getContrastColor(C.ACCENT), bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    },
    () => {
      addTextOnShape(sWe, "16 + 7", {
        x: 5.3, y: 3.15, w: 4.2, h: 0.6, rectRadius: 0.08,
        fill: { color: C.ACCENT },
      }, {
        fontSize: 22, fontFace: FONT_H, color: T.getContrastColor(C.ACCENT), bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    },
    () => {
      addTextOnShape(sWe, "= 23", {
        x: 5.3, y: 3.85, w: 4.2, h: 0.7, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, {
        fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    },
  ]);
  runSlideDiagnostics(sWe, pres);

  /* 12. You Do - worksheet, different content again */
  const sYou = pres.addSlide();
  T.addTopBar(sYou, C.SUCCESS);
  T.addBadge(sYou, "You Do", { color: C.SUCCESS });
  T.addTitle(sYou, "Your turn - BODMAS Practice sheet");
  addCard(sYou, 0.5, CONTENT_TOP, 4.5, 2.6, { fill: C.WHITE, strip: C.SUCCESS });
  sYou.addText([
    { text: "First: rebuild the ladder in Section 1.", options: { bullet: true, breakLine: true } },
    { text: "Next: solve Section 2, rewriting every line.", options: { bullet: true, breakLine: true } },
    { text: "Then: try the challenge in Section 3.", options: { bullet: true } },
  ], {
    x: 0.75, y: CONTENT_TOP + 0.2, w: 4.0, h: 2.2,
    fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
  });
  addCard(sYou, 5.2, CONTENT_TOP, 4.3, 2.6, { fill: C.BG_CARD, strip: C.PRIMARY });
  sYou.addText("Same tier,\nleft to right", {
    x: 5.4, y: CONTENT_TOP + 0.2, w: 3.9, h: 2.2,
    fontSize: 24, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  sYou.addText("Show every step. Do not skip working.", {
    x: 0.5, y: 4.1, w: 9, h: 0.5,
    fontSize: 17, fontFace: FONT_B, color: C.MUTED,
    align: "center", valign: "middle", margin: 0,
  });
  addFooter(sYou, FOOTER);
  sYou.addNotes(NOTES_YOUDO);
  runSlideDiagnostics(sYou, pres);

  /* 13. Exit ticket */
  exitTicketSlide(
    pres,
    ["3 + 24 ÷ 4"],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2 }
  );

  /* 14. Closing */
  closingSlide(
    pres,
    {
      prompt: "Which tier tripped you up most today?",
      criteria: [
        "I can name which part of an expression to work out first.",
        "I can solve an expression with brackets and two or more operations.",
        "I can explain why divide and multiply are worked left to right.",
      ],
      selfAssess: "Thumbs up, sideways or down for each one.",
    },
    NOTES_CLOSING,
    FOOTER
  );

  return pres;
}

/**
 * A worked example built one step per click: hero expression, then reasoning
 * steps, then the answer. Nothing on this slide is visible until the teacher
 * chooses, so the class never reads ahead of the modelling.
 */
function buildWorkedStep(pres, badge, title, expression, steps, answer, notes, trapText) {
  const s = pres.addSlide();
  T.addTopBar(s, C.PRIMARY);
  T.addBadge(s, badge);
  T.addTitle(s, title);

  addCard(s, 0.5, CONTENT_TOP, 9, 0.9, { fill: C.WHITE, strip: C.PRIMARY });
  s.addText(expression, {
    x: 0.7, y: CONTENT_TOP + 0.05, w: 8.6, h: 0.8,
    fontSize: 38, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });

  addFooter(s, FOOTER);
  s.addNotes(notes);

  const stepTop = CONTENT_TOP + 1.05;
  const stepH = 0.62;
  const stepGap = 0.12;

  const stepFns = steps.map((st, i) => () => {
    addTextOnShape(s, st.text, {
      x: 0.7, y: stepTop + i * (stepH + stepGap), w: 8.6, h: stepH, rectRadius: 0.08,
      fill: { color: st.color },
    }, {
      fontSize: 18, fontFace: FONT_B, color: T.getContrastColor(st.color), bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });

  const answerY = stepTop + steps.length * (stepH + stepGap);
  stepFns.push(() => {
    addTextOnShape(s, answer, {
      x: 2.2, y: answerY, w: 5.6, h: 0.72, rectRadius: 0.1,
      fill: { color: C.SUCCESS },
    }, {
      fontSize: 30, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  });

  if (trapText) {
    stepFns.push(() => {
      s.addText(trapText, {
        x: 0.7, y: answerY + 0.8, w: 8.6, h: 0.42,
        fontSize: 15, fontFace: FONT_B, color: C.ALERT, italic: true,
        align: "center", valign: "middle", margin: 0,
      });
    });
  }

  clickBuild(s, stepFns);
  runSlideDiagnostics(s, pres);
  return s;
}

/* ══════════════════════════════════════════════════════════════════════
 * Companion PDFs
 * ══════════════════════════════════════════════════════════════════════ */

// Teacher copy: the complete ladder.
const LADDER_ROWS = [
  "B - Brackets first",
  "O - Orders (powers)",
  "D M - Divide and Multiply, left to right",
  "A S - Add and Subtract, left to right",
];

// Student copy: the same ladder with the load-bearing words removed, so
// Section 1 is a real re-grounding task rather than a printed answer.
const LADDER_GAPS = [
  "B - ______________ first",
  "O - Orders (powers)",
  "D M - Divide and ______________, ________ to ________",
  "A S - Add and ______________, ________ to ________",
];

const PROBLEMS = [
  "5 + 2 x 3 =",
  "18 ÷ 3 + 1 =",
  "(4 + 6) x 3 =",
  "20 - 24 ÷ (3 + 3) =",
];

const ANSWERS = [
  "5 + 2 x 3. Times tier first: 2 x 3 = 6. Then 5 + 6 = 11.",
  "18 ÷ 3 + 1. Divide tier first: 18 ÷ 3 = 6. Then 6 + 1 = 7.",
  "(4 + 6) x 3. Brackets first: 4 + 6 = 10. Then 10 x 3 = 30.",
  "20 - 24 ÷ (3 + 3). Brackets: 3 + 3 = 6. Divide: 24 ÷ 6 = 4. Then 20 - 4 = 16.",
];

function buildWorksheet() {
  const doc = createPdf({ title: WORKSHEET.name });
  let y = addPdfHeader(doc, "BODMAS Practice", {
    color: C.PRIMARY,
    subtitle: "Year 6 Numeracy",
  });

  y = addSectionHeading(doc, "Section 1: Rebuild the ladder", y, { color: C.PRIMARY });
  y = addBodyText(doc, "Fill the gaps. This is your reminder for every question below.", y);
  LADDER_GAPS.forEach((row) => { y = addBodyText(doc, row, y); });

  y = addSectionHeading(doc, "Section 2: Solve. Rewrite the whole line each step.", y, { color: C.PRIMARY });
  y = addBodyText(doc, "Worked example: 8 + 4 x 2. Times tier first: 4 x 2 = 8. Then 8 + 8 = 16.", y);
  PROBLEMS.forEach((p) => {
    y = addBodyText(doc, p, y);
    y = addLinedArea(doc, y, 2);
  });

  y = addSectionHeading(doc, "Section 3: Challenge", y, { color: C.ACCENT });
  y = addBodyText(doc, "Write an expression with brackets that equals 12. Swap with a partner.", y);
  y = addLinedArea(doc, y, 2);

  addPdfFooter(doc, FOOTER);
  return writePdf(doc, path.join(RES_DIR, path.basename(WORKSHEET.fileName)));
}

function buildAnswerKey() {
  const doc = createPdf({ title: ANSWER_KEY.name });
  let y = addPdfHeader(doc, "BODMAS Answer Key", {
    color: C.ALERT,
    subtitle: "Teacher copy - Year 6 Numeracy",
    showNameDate: false,
  });

  y = addSectionHeading(doc, "Section 1", y, { color: C.ALERT });
  LADDER_ROWS.forEach((row) => { y = addBodyText(doc, row, y); });

  y = addSectionHeading(doc, "Section 2 - full working", y, { color: C.ALERT });
  ANSWERS.forEach((a) => { y = addBodyText(doc, a, y); });

  y = addSectionHeading(doc, "Section 3 - what to look for", y, { color: C.ALERT });
  y = addBodyText(doc, "Any expression that uses brackets and evaluates to 12, with correct working shown. Example: (2 + 4) x 2 = 12.", y);
  y = addTipBox(doc, "Mark the FIRST step, not just the final answer. A wrong first step is the tier misconception.", y, { color: C.ALERT });

  addPdfFooter(doc, FOOTER);
  return writePdf(doc, path.join(RES_DIR, path.basename(ANSWER_KEY.fileName)));
}

/* ══════════════════════════════════════════════════════════════════════ */

(async function main() {
  const pres = build();
  const out = path.join(LESSON_FOLDER, PPTX_NAME);
  await pres.writeFile({ fileName: out });
  console.log("PPTX written to " + out);

  buildWorksheet();
  console.log("PDF written: " + WORKSHEET.fileName);
  buildAnswerKey();
  console.log("PDF written: " + ANSWER_KEY.fileName);

  console.log("Lesson build complete.");
})();
