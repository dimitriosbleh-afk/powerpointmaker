"use strict";

// Fractions, Decimals & Percentages (Year 5/6 composite Numeracy) - Session 2 of 2.
//
// "ONE LINE, ONE WHOLE." Apply knowledge of equivalence to compare, order and
// represent common fractions - halves, thirds and quarters - on the SAME number
// line, and justify the order.
//
// Victorian Curriculum 2.0: Mathematics, Number, Year 5/6.
// Codes supplied by the teacher: VC2M5N04, VC2M6N03.
//
// The VC2M6N03 elaborations drive the two models: a fraction wall built from
// equal-length wholes, then the same fractions located as POINTS on one number
// line, so students meet both models and can say how they differ.
//
// UNIT ANCHOR (megaprompt section 79, held identical to Session 1):
//   "Same whole. Same line. Then compare."
//   Drawn by builds/fdpeq_lib.js so it cannot drift between sessions.
//
// Daily Review: Metric Measurement and Unit Conversion, plus Practical Problem
// Solving and Financial Reasoning (the remaining teacher-supplied focus areas).
// Fluency: subtraction vertical algorithm with decimals (teacher-supplied focus).
//
// Catch-up: the launch is answerable by eye, the recap slide rebuilds the wall
// from scratch, and Section 1 of the worksheet prints the wall on paper.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const { composeGlanceNotes } = require("../themes/core/composeNotes");
const { createLib, ANCHOR_PHRASE } = require("./fdpeq_lib");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addFractionStripsPdf, addNumberLinePdf,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

// Same variant as Session 1: one unit, one palette (CLAUDE.md theme cohesion).
const UNIT_VARIANT = 3;
const T = createTheme("numeracy", "grade56", UNIT_VARIANT);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, closingSlide, keyWordSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  dailyReviewSlide, addRevealAnswerBar,
  addFractionStripSet, addNumberLine,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  clickBuild, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 2;
const FOOTER = "Fractions & Decimals | Session 2 of 2 | Year 5/6 Numeracy";
const OUT_DIR = "output/FDPEQ_Lesson2_Comparing_Fractions_On_One_Line";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const lib = createLib(T, FOOTER);

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Read equivalences off a printed fraction wall, compare by renaming, then mark and order on one number line.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Session 2 worksheet, with the errors to watch for.");
const SCAFFOLD_RES = makeSessionResource(SESSION,
  "Fraction Wall Helper",
  "The wall and the number line drawn ready, with the first comparison already worked. Glue into books.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, SCAFFOLD_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

/* ── Teacher notes (Glance Format, megaprompt sections 45-47) ─────────────── */

const NOTES_TITLE =
  "Open Session 2 of 2. Name today's focus in one line, then move to Teacher Resources.";

const NOTES_RESOURCES =
  "Prep slide. Print one worksheet and one answer key per student, and the Fraction Wall " +
  "Helper for anyone who needs the wall in front of them.\n" +
  "The Year 8 extension handed out in Session 1 is still the early-finisher task today.\n" +
  "CATCH-UP: missed yesterday? No problem. Today's launch is answerable just by looking, " +
  "the recap slide rebuilds the wall from nothing, and Section 1 of the worksheet prints " +
  "the wall on paper. Hand a returner the Helper card and start them at Section 1.";

const NOTES_OVERVIEW =
  "For you, not the class. Skip past it when projecting.\n" +
  "Day two of a five-day plan taught in two. Today is fraction equivalence, comparing and " +
  "ordering on ONE number line. Mixed numerals appear only in the challenge and extension.\n" +
  "Decision points: the boards check after the two models, the hinge after guided practice, " +
  "and the exit ticket. Between them keep the pace brisk.\n" +
  "SOURCES: teacher-supplied content descriptions VC2M5N04 and VC2M6N03, and the supplied " +
  "Daily Review and Number Fluency focus lists.";

const NOTES_DR = composeGlanceNotes({
  answer: "2400 g, 3.5 L, and two dollars thirty five.",
  beats: [
    "POINT to the handy facts. SAY: Decide multiply or divide before you write.",
    [
      "ASK: Write all three answers.",
      "90 sec. Cue: Write it... chin it... show me.",
      "EXPECT: 2400 g, 3.5 L, two dollars thirty five.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> reveal, tick and fix, straight into fluency.",
      "Less -> count up from the drink price to five dollars, re-ask.",
    ],
  ],
  trap: [
    "dividing to reach grams because grams sound small.",
    "Fix: ask more or fewer, student re-decides.",
  ],
  prep: "Retrieval of metric conversion and money problems, the two remaining supplied review focuses. Whole block under 5 minutes.",
  tag: "[Stage 1 | Daily Review | Retention and recall | HITS 6]",
});

const NOTES_FLUENCY = composeGlanceNotes({
  answer: "2.45, 8.53 and 6.45.",
  beats: [
    "SAY: Fluency again. Columns, and the decimal points in a straight line.",
    [
      "SAY: Fill any empty place with a zero first.",
      "SAY: 12 becomes 12.00, so every column has a digit.",
    ],
    [
      "ASK: What are the three answers?",
      "90 sec. Boards up on cue.",
      "EXPECT: 2.45, 8.53, 6.45.",
    ],
    [
      "SCAN boards.",
      "80%+ -> reveal, tick and fix.",
      "Less -> rebuild 12 minus 3.47 with the zeros written in, re-ask.",
    ],
  ],
  trap: [
    "lining the numbers up on the right instead of on the point.",
    "Fix: rule down the decimal points, student redoes it.",
  ],
  prep: "Same fluency focus as Session 1: vertical subtraction with decimals. Brisk retrieval, under 4 minutes, not new teaching.",
  tag: "[Stage 1 | Fluency | Retention and recall | HITS 6]",
});

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "the second bar, 3/4. Its shaded part reaches further.",
  beats: [
    "POINT to both bars. SAY: Same whole. The two bars are exactly the same length.",
    [
      "ASK: Which bar has more shaded?",
      "20 sec. Thumbs only, voices off. Show me... now.",
      "EXPECT: the bottom one, 3/4.",
    ],
    [
      "SCAN the room.",
      "80%+ -> cold call one: How did you decide without counting?",
      "Less -> run a finger along both shaded parts, re-ask.",
    ],
    [
      "REVEAL after every thumb is up.",
      "SAY: More pieces does not mean more amount.",
    ],
  ],
  trap: [
    "choosing 2/3 because thirds are bigger pieces than quarters.",
    "Fix: compare how far each shaded part reaches, student re-decides.",
  ],
  prep: "Low-coupling launch: answerable by looking, so a student who missed Session 1 joins here. It sets up the need for one shared line.",
  tag: "[Launch | Knowledge and memory | HITS 2, 6]",
});

const NOTES_LI_SC = composeGlanceNotes({
  beats: [
    "POINT to the learning intention. SAY: Today we put fractions on one line.",
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: The first one everyone can do - find a fraction equal to one half.",
  ],
  prep: "SC1 is reachable by all; SC2 is the core target the exit ticket assesses; SC3 stretches to ordering and justifying. Tier labels stay off the slide.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_VOCAB = composeGlanceNotes({
  answer: "equivalent means equal in value - one amount, written a different way.",
  beats: [
    [
      "SAY: Equivalent means equal in value.",
      "SAY: Same amount of the bar, just cut into different pieces.",
    ],
    [
      "POINT to 1/2 and 3/6.",
      "SAY: These are not two amounts. They are one amount, named twice.",
    ],
    [
      "ASK: Is 2/4 the same amount as 1/2?",
      "10 sec. Thumbs only, voices off. Show me... now.",
      "EXPECT: thumbs up.",
    ],
  ],
  trap: [
    "thinking 3/6 must be more than 1/2 because the numbers are bigger.",
    "Fix: line the two rows up on the wall, student re-decides.",
  ],
  prep: "The word that carries the session. A student who trusts equivalence stops fearing fractions with unfamiliar denominators.",
  tag: "[Key Vocabulary | Knowledge and memory | HITS 6]",
});

const NOTES_RECAP = composeGlanceNotes({
  answer: "every row is the same whole, just cut into a different number of pieces.",
  beats: [
    "POINT across the three moves. SAY: Same whole. Same line. Then compare.",
    [
      "RUN a finger down the left edge. SAY: Every row starts at nothing.",
      "SAY: Every row ends at one whole. Only the cuts change.",
    ],
    [
      "ASK: Which row has the smallest pieces?",
      "10 sec. Everyone points to your screen.",
      "EXPECT: the bottom row, twelfths.",
    ],
  ],
  trap: [
    "thinking a row with more pieces holds more altogether.",
    "Fix: compare the two end points, student re-decides.",
  ],
  stretch: "which two rows could you use to make 1/2 three different ways?",
  help: "the Fraction Wall Helper card, finger on one row at a time.",
  prep: "The catch-up re-entry slide. A student who missed Session 1 can do everything that follows from this wall alone.",
  tag: "[Recap | Knowledge and memory | HITS 3, 6]",
});

const NOTES_IDO1 = composeGlanceNotes({
  answer: "1/2 = 2/4 = 3/6 = 6/12. All stop in the same place.",
  beats: [
    "SAY: Same whole. Watch where the shading stops.",
    [
      "SAY: I shade one of the two halves, reaching the middle.",
      "SAY: Now quarters. I need two to reach that same place.",
    ],
    "REVEAL sixths and twelfths after the halves and quarters are said.",
    [
      "SAY: Three sixths, six twelfths. All stop in the same spot.",
      "SAY: More pieces, but each piece is smaller.",
    ],
  ],
  trap: [
    "reading 6/12 as more than 1/2 because six beats one.",
    "Fix: drop a finger down the stopping line, student re-reads.",
  ],
  stretch: "write two more fractions equal to 1/2. Prove them.",
  help: "the Helper card, one row at a time.",
  prep: "First model. The anchor is restated in full so a returning student can follow from this slide alone.",
  tag: "[Stage 2 | I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_IDO2 = composeGlanceNotes({
  answer: "1/2, then 2/3, then 3/4. In twelfths: 6, 8 and 9.",
  beats: [
    "SAY: The wall shows area. A line shows a point.",
    [
      "SAY: Twelfths work for halves, thirds and quarters.",
      "SAY: One half is six twelfths, two thirds eight, three quarters nine.",
    ],
    "REVEAL the marks after I have named all three in twelfths.",
    [
      "POINT along the line. SAY: Now they share one line.",
      "SAY: Further right means larger. Read the order straight off.",
    ],
  ],
  trap: [
    "spacing the marks by guess instead of counting twelfths.",
    "Fix: count the ticks aloud from zero, student re-marks.",
  ],
  stretch: "where would 1 1/2 sit on this line?",
  help: "the Helper card line with twelfths already numbered.",
  prep: "Second model, the number line the elaboration asks for. Say aloud how it differs from the wall: a point, not an area.",
  tag: "[Stage 2 | I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_CFU1 = composeGlanceNotes({
  answer: "5/6. In twelfths it is 10/12, and 3/4 is only 9/12.",
  beats: [
    "SAY: Check on your boards. Rename them both before you choose.",
    [
      "ASK: Which is bigger, 5/6 or 3/4?",
      "45 sec. Cue: Write it... chin it... show me.",
      "EXPECT: 5/6, shown as 10/12.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one: Why did you choose twelfths?",
      "Less -> line up sixths and twelfths on the wall, re-ask 2/3 or 3/4.",
    ],
    "REVEAL only after every board is up.",
  ],
  trap: [
    "choosing 3/4 because quarters are bigger pieces than sixths.",
    "Fix: rename both to twelfths, student re-decides.",
  ],
  prep: "Decision point one. Checks that renaming happens before comparing, before releasing into guided practice.",
  tag: "[Stage 2 | CFU | Supported application | HITS 7, 8]",
});

const NOTES_WEDO1 = composeGlanceNotes({
  answer: "5/6. Two thirds is four sixths, and five sixths is more.",
  beats: [
    "SAY: Your turn. These two are kinder - sixths will do it.",
    [
      "ASK: Which is bigger, 2/3 or 5/6?",
      "60 sec with your partner. Boards up on cue.",
      "EXPECT: 5/6, with 2/3 renamed as 4/6.",
    ],
    [
      "CIRCULATE. Look for the rename written before the choice.",
      "80%+ -> reveal, then bounce: Do you agree? Add one thing.",
      "Less -> shade thirds and sixths, re-ask.",
    ],
  ],
  trap: [
    "renaming only one of the two fractions.",
    "Fix: both must have the same bottom number, student redoes it.",
  ],
  stretch: "how much bigger is 5/6 than 2/3?",
  help: "the wall with thirds and sixths already lined up.",
  prep: "Guided practice, faded from the I Do: nothing worked stays on screen and the denominator choice is theirs.",
  tag: "[Stage 3 | We Do | Supported application | HITS 4, 5]",
});

const NOTES_WEDO2 = composeGlanceNotes({
  answer: "1/3, then 1/2, then 2/3, then 3/4.",
  beats: [
    "SAY: Four fractions now. Same line, same whole, then read the order.",
    [
      "ASK: Rename all four in twelfths, then order them.",
      "90 sec with your partner. Boards up on cue.",
      "EXPECT: 4/12, 6/12, 8/12, 9/12.",
    ],
    [
      "CIRCULATE. Check all four are renamed before any ordering.",
      "80%+ -> reveal, then cold call one: How do you know 1/3 is smallest?",
      "Less -> mark 1/3 and 1/2 together, re-ask.",
    ],
  ],
  trap: [
    "ordering by the bottom numbers, so 1/2 lands before 1/3.",
    "Fix: mark both on the line, student re-orders.",
  ],
  stretch: "add 5/6 to the list and place it too.",
  help: "twelfths written under each fraction already.",
  prep: "Flex slide - cut this if the models ran long. The exit ticket still reaches ordering and justifying.",
  tag: "[Stage 3 | We Do | Supported application | HITS 4, 5]",
});

const NOTES_CFU2 = composeGlanceNotes({
  answer: "B. Cut one whole into more pieces and each piece gets smaller.",
  beats: [
    "SAY: Hinge question. A student said this. Decide if they are right.",
    [
      "ASK: Is it A, B or C?",
      "30 sec, voices off. Cue: Write it... chin it... show me.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: Convince us. Then reveal.",
      "Less -> line halves against thirds on the wall, re-ask 1/4 or 1/5.",
    ],
    "REVEAL only after every board is up.",
  ],
  trap: [
    "reading the bottom numbers as ordinary whole numbers.",
    "Fix: share one bar between two, then three, student re-decides.",
  ],
  prep: "Decision point two, the hinge of the unit. A means denominators read as whole numbers. C is the right answer for a wrong reason - a third is not 0.3.",
  tag: "[Stage 2 | CFU hinge | Supported application | HITS 7, 8]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "open - listen for renaming said out loud before any comparing.",
  beats: [
    "SAY: On your own now. The wall is printed at the top for you.",
    "SAY: Every comparison, rename to the same bottom number first.",
    [
      "CIRCULATE. Sit with the catch-up group first.",
      "COLLECT: check Section 1 before anyone starts Section 3.",
    ],
    "TIME: about 12 minutes, then stop for the exit ticket.",
  ],
  stretch: "the Challenge box, then the Year 8 sheet from Session 1.",
  help: "the Fraction Wall Helper, Section 1 worked with you.",
  prep: "Section 1 is the re-grounding task: doable from the printed wall alone, with no earlier session needed.",
  tag: "[Stage 4 | You Do | Mastery and application | HITS 4, 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "any two of 2/4, 3/6, 4/8, 6/12. Then 1/4, 1/2, 2/3.",
  beats: [
    "SAY: On your own, on your whiteboard. No partner talk for this one.",
    [
      "SAY: One, write two fractions equal to one half.",
      "SAY: Two, order the three fractions and say how you know.",
    ],
    [
      "TIME: about 3 minutes.",
      "COLLECT boards, or photograph them for your records.",
    ],
  ],
  trap: [
    "ordering by the bottom numbers alone.",
    "Fix: note the name, revisit the wall in the next session.",
  ],
  prep: "Assesses the core target (using equivalence to compare) and reaches into justifying. Keep the SC number off the slide.",
  tag: "[Stage 5 | Exit Ticket | Mastery and application | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - listen for rename to the same bottom number, then compare.",
  beats: [
    "SAY: Look back at the three I can statements. Thumbs for each one.",
    [
      "ASK: What do you always do before comparing two fractions?",
      "30 sec. Turn and tell. Partner A first.",
      "EXPECT: put them on the same line, or rename them.",
    ],
    "COLLECT: note who showed need more practice, for a short revisit.",
  ],
  prep: "Close of the two-day unit. The keeper: same whole, same line, then compare - and it works for fractions, decimals and percentages alike.",
  tag: "[Closing | Retention and recall | HITS 9]",
});

/* ── Local visual: fractions as points on one shared number line ──────────── */

// Used by the I Do and again, faded, by the We Do. One helper so the model and
// the guided practice are visibly the same representation (megaprompt section 79).
function sameLineSlide(pres, cfg) {
  const stageColor = STAGE_COLORS[String(cfg.stage)];
  const s = pres.addSlide();
  addTopBar(s, stageColor);
  addStageBadge(s, cfg.stage, cfg.stageLabel);
  addTitle(s, cfg.title, { color: stageColor });

  const items = cfg.items;
  const n = items.length;
  const chipGap = 0.35;
  const chipW = (8.2 - chipGap * (n - 1)) / n;

  items.forEach((it, i) => {
    addTextOnShape(s, it.frac, {
      x: 0.9 + i * (chipW + chipGap), y: 1.45, w: chipW, h: 0.60, rectRadius: 0.08,
      fill: { color: stageColor },
    }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
  });

  // The number line is drawn from the start: the marks, not the line, are what
  // the click protects.
  const lineX = 0.9;
  const lineW = 8.2;
  const lineY = 3.45;
  const labels = [];
  for (let i = 0; i <= 12; i += 1) labels.push(i === 0 ? "0" : (i === 12 ? "1" : ""));
  const geo = addNumberLine(s, lineX, lineY, lineW, labels, null);

  addFooter(s, FOOTER);
  s.addNotes(cfg.notes);

  clickBuild(s, [
    () => {
      items.forEach((it, i) => {
        addTextOnShape(s, "= " + it.twelfths, {
          x: 0.9 + i * (chipW + chipGap), y: 2.15, w: chipW, h: 0.55, rectRadius: 0.08,
          fill: { color: C.SECONDARY },
        }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
    },
    () => {
      items.forEach((it) => {
        const mx = lineX + it.index * geo.intervalW;
        s.addShape("roundRect", {
          x: mx - 0.08, y: lineY - 0.08, w: 0.16, h: 0.16, rectRadius: 0.08,
          fill: { color: C.ALERT },
        });
        s.addText(it.frac, {
          x: mx - 0.275, y: 2.90, w: 0.55, h: 0.40,
          fontSize: 15, fontFace: FONT_B, color: C.ALERT, bold: true,
          align: "center", valign: "bottom", margin: 0,
        });
      });
    },
    () => {
      addRevealAnswerBar(s, cfg.answer, {
        color: C.SUCCESS, label: "Order", fontSize: 22, y: 4.28, h: 0.80,
      });
    },
  ]);

  runSlideDiagnostics(s, pres);
  return s;
}

/* ── Build ────────────────────────────────────────────────────────────────── */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  titleSlide(pres, "One Line, One Whole",
    "Comparing and ordering fractions using equivalence",
    "Year 5/6 Numeracy | Session 2 of 2", NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3. Teacher-facing overview
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.MUTED);
    addBadge(s, "For the teacher", { color: C.MUTED, w: 2.4 });
    addTitle(s, "Day two of two", { color: C.MUTED });

    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;

    addCard(s, 0.5, cardY, 4.4, cardH, { strip: C.PRIMARY });
    s.addText("What today carries", {
      x: 0.7, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Yesterday  100% is the whole", options: { bold: true, breakLine: true } },
      { text: "Fractions, decimals and percentages named the same amount.", options: { breakLine: true } },
      { text: "Today  One line, one whole", options: { bold: true, breakLine: true } },
      { text: "Equivalence to compare and order fractions on one number line.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Two models, on purpose:", options: { color: C.MUTED, breakLine: true } },
      { text: "the wall shows an AREA of the whole; the number line shows a POINT. Say that difference out loud.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Unit anchor:", options: { color: C.MUTED, breakLine: true } },
      { text: ANCHOR_PHRASE, options: { color: C.PRIMARY, bold: true } },
    ], {
      x: 0.7, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addCard(s, 5.1, cardY, 4.4, cardH, { strip: C.ACCENT });
    s.addText("How today works", {
      x: 5.3, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
    });
    s.addText([
      { text: "Shape: example-first. Wall model, line model, guided practice, then independent.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Decision points:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "boards check after the models, hinge after guided practice, exit ticket.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Catch-up:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "nothing today needs Session 1. The recap rebuilds the wall from nothing.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Flex: cut the ordering We Do if the models ran long.", options: { italic: true, color: C.MUTED, breakLine: true } },
      { text: "Curriculum: Mathematics 2.0, Number, Year 5/6.", options: { italic: true, color: C.MUTED } },
    ], {
      x: 5.3, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_OVERVIEW);
    runSlideDiagnostics(s, pres);
  })();

  // 4. Daily Review - metric conversion and money (supplied focuses)
  (() => {
    const s = dailyReviewSlide(pres, "Daily Review: Measurement and money",
      [
        "Convert 2.4 kg into grams.",
        "Convert 3500 mL into litres.",
        "A drink costs $2.65. You pay with $5. What is your change?",
      ],
      NOTES_DR, FOOTER,
      (slide, lg) => {
        // Panel stops at 4.15 so the click-revealed answer bar (top 4.30)
        // keeps its clearance.
        const y0 = lg.panelTopPadded;
        const cardBottom = 4.15;
        addCard(slide, lg.rightX, y0, lg.rightW, cardBottom - y0, { strip: C.ACCENT });
        slide.addText("Handy facts", {
          x: lg.rightX + 0.15, y: y0 + 0.10, w: lg.rightW - 0.30, h: 0.28,
          fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, align: "center", margin: 0,
        });
        const facts = ["1 km = 1000 m", "1 kg = 1000 g", "1 L = 1000 mL", "$1 = 100 c"];
        facts.forEach((f, i) => {
          addTextOnShape(slide, f, {
            x: lg.rightX + 0.25, y: y0 + 0.52 + i * 0.58, w: lg.rightW - 0.50, h: 0.48,
            rectRadius: 0.08, fill: { color: i === 3 ? C.SECONDARY : C.PRIMARY },
          }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
        });
      });
    clickBuild(s, [
      () => { addRevealAnswerBar(s, ["2400 g", "3.5 L", "$2.35"], { color: C.SUCCESS }); },
    ]);
    runSlideDiagnostics(s, pres);
  })();

  // 5. Fluency - subtraction vertical algorithm with decimals (supplied focus)
  lib.fluencySlide(pres, "Fluency: Subtraction, set out in columns",
    ["7.3 - 4.85", "12 - 3.47", "9.05 - 2.6"],
    ["2.45", "8.53", "6.45"],
    NOTES_FLUENCY);

  // 6. Launch - which bar has more shaded? (answerable by eye)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "Launch", { color: C.SECONDARY });
    addTitle(s, "Which bar has more shaded?", { color: C.SECONDARY });

    s.addText("Two bars, exactly the same size.", {
      x: 0.5, y: 1.32, w: 9, h: 0.30,
      fontSize: 15, fontFace: FONT_B, color: C.MUTED, bold: true,
      align: "center", margin: 0,
    });

    addFractionStripSet(s, 1.2, 1.75, 7.6, 0.85,
      [{ denom: 3, shaded: 2, label: "2/3", color: C.PRIMARY }],
      { labelW: 0.9, labelFontSize: 20 });
    addFractionStripSet(s, 1.2, 3.05, 7.6, 0.85,
      [{ denom: 4, shaded: 3, label: "3/4", color: C.SECONDARY }],
      { labelW: 0.9, labelFontSize: 20 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    clickBuild(s, [
      () => {
        addTextOnShape(s, "3/4 reaches further. More pieces does not mean more amount.", {
          x: 0.5, y: 4.15, w: 9, h: 0.82, rectRadius: 0.1,
          fill: { color: C.SUCCESS },
        }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      },
    ]);
    runSlideDiagnostics(s, pres);
  })();

  // 7. Learning Intention and Success Criteria
  liSlide(pres,
    "We are learning to use equivalent fractions to compare and order fractions on the same number line.",
    [
      "I can find a fraction that is equal to one half.",
      "I can rename two fractions so they have the same bottom number and say which is larger.",
      "I can put several fractions in order on one number line and justify the order.",
    ],
    NOTES_LI_SC, FOOTER);

  // 8. Key vocabulary
  keyWordSlide(pres, {
    word: "equivalent",
    meaning: "Equal in value. The same amount, written a different way.",
    example: "1/2, 2/4 and 3/6 are equivalent. One amount, three names.",
    routine: ["Say it", "Show it", "Use it"],
    color: C.PRIMARY,
  }, NOTES_VOCAB, FOOTER);

  // 9. Recap - the fraction wall (catch-up re-entry slide)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Recap", { color: C.PRIMARY });
    addTitle(s, "One whole in every row", { color: C.PRIMARY });

    lib.anchorMoves(s, CONTENT_TOP);

    addFractionStripSet(s, 0.8, 2.05, 8.4, 2.85, [
      { denom: 1, shaded: 0, label: "1 whole" },
      { denom: 2, shaded: 0, label: "halves" },
      { denom: 3, shaded: 0, label: "thirds" },
      { denom: 4, shaded: 0, label: "quarters" },
      { denom: 6, shaded: 0, label: "sixths" },
      { denom: 12, shaded: 0, label: "twelfths" },
    ], { labelW: 1.3, labelFontSize: 14 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_RECAP);
    runSlideDiagnostics(s, pres);
  })();

  // 10. I Do - one half has many names
  (() => {
    const rowX = 5.45;
    const rowW = 3.9;
    const rowH = 0.50;
    const rowGap = 0.12;
    const rowTop = 1.92;
    let slideRef;
    const rowAt = (i) => rowTop + i * (rowH + rowGap);
    const drawRow = (slide, i, denom, shaded, label) => {
      addFractionStripSet(slide, rowX, rowAt(i), rowW, rowH,
        [{ denom, shaded, label, color: C.SECONDARY }],
        { labelW: 0.85, labelFontSize: 14 });
    };

    const s = workedExSlide(pres, 2, "I Do", "One half has many names",
      [
        "One whole in every row.",
        "Shade half the top row.",
        "Now match it in quarters.",
        "Two quarters reach it.",
        "Three sixths reach it too.",
        "Six twelfths as well.",
      ],
      NOTES_IDO1, FOOTER,
      (slide, lg) => {
        slideRef = slide;
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW,
          lg.safeBottom - lg.panelTopPadded, { strip: C.PRIMARY });
        slide.addText("Where does the shading stop?", {
          x: lg.rightX + 0.18, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.36, h: 0.28,
          fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
          align: "center", margin: 0,
        });
        drawRow(slide, 0, 2, 1, "1/2");
      });

    clickBuild(s, [
      () => { drawRow(slideRef, 1, 4, 2, "2/4"); },
      () => { drawRow(slideRef, 2, 6, 3, "3/6"); drawRow(slideRef, 3, 12, 6, "6/12"); },
      () => {
        addTextOnShape(slideRef, "1/2 = 2/4 = 3/6 = 6/12", {
          x: 5.45, y: 4.45, w: 3.9, h: 0.52, rectRadius: 0.08,
          fill: { color: C.SUCCESS },
        }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      },
    ]);
    runSlideDiagnostics(s, pres);
  })();

  // 11. I Do - the same fractions as points on one line
  sameLineSlide(pres, {
    stage: 2, stageLabel: "I Do",
    title: "Put them on the same line",
    items: [
      { frac: "1/2", twelfths: "6/12", index: 6 },
      { frac: "2/3", twelfths: "8/12", index: 8 },
      { frac: "3/4", twelfths: "9/12", index: 9 },
    ],
    answer: "1/2, then 2/3, then 3/4",
    notes: NOTES_IDO2,
  });

  // 12. CFU - decision point one
  lib.checkSlide(pres, {
    title: "Rename, then compare",
    technique: "Show Me Boards",
    question: "Which is bigger: 5/6 or 3/4?",
    answer: "5/6 = 10/12.  3/4 = 9/12.",
    notes: NOTES_CFU1,
  });

  // 13. We Do - compare by renaming
  (() => {
    let revealRename;
    let revealVerdict;
    const s = workedExSlide(pres, 3, "We Do", "Which is bigger: 2/3 or 5/6?",
      [
        "With your partner.",
        "",
        "Same whole for both.",
        "Sixths are smaller pieces.",
        "Rename thirds into sixths.",
        "Then compare the tops.",
      ],
      NOTES_WEDO1, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW,
          lg.safeBottom - lg.panelTopPadded, { strip: C.SECONDARY });
        slide.addText("Line them up", {
          x: lg.rightX + 0.18, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.36, h: 0.28,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        addFractionStripSet(slide, 5.45, 1.95, 3.9, 1.40, [
          { denom: 3, shaded: 2, label: "2/3", color: C.PRIMARY },
          { denom: 6, shaded: 5, label: "5/6", color: C.SECONDARY },
        ], { labelW: 0.85, labelFontSize: 14 });

        revealRename = () => {
          addTextOnShape(slide, "2/3 = 4/6", {
            x: 5.45, y: 3.50, w: 3.9, h: 0.55, rectRadius: 0.08,
            fill: { color: C.PRIMARY },
          }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
        };
        revealVerdict = () => {
          addTextOnShape(slide, "5 sixths beats 4 sixths", {
            x: 5.45, y: 4.20, w: 3.9, h: 0.60, rectRadius: 0.08,
            fill: { color: C.SUCCESS },
          }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
        };
      });
    clickBuild(s, [
      () => revealRename(),
      () => revealVerdict(),
    ]);
    runSlideDiagnostics(s, pres);
  })();

  // 14. We Do - order four fractions on one line (flex slide)
  sameLineSlide(pres, {
    stage: 3, stageLabel: "We Do",
    title: "Put these in order, smallest first",
    items: [
      { frac: "3/4", twelfths: "9/12", index: 9 },
      { frac: "1/3", twelfths: "4/12", index: 4 },
      { frac: "1/2", twelfths: "6/12", index: 6 },
      { frac: "2/3", twelfths: "8/12", index: 8 },
    ],
    answer: "1/3, 1/2, 2/3, 3/4",
    notes: NOTES_WEDO2,
  });

  // 15. CFU hinge - decision point two
  lib.checkSlide(pres, {
    title: "Spot the error",
    technique: "Show Me Boards",
    lead: "A student says:",
    question: "1/3 is bigger than 1/2, because 3 is bigger than 2.",
    questionSize: 25,
    options: [
      "A. They are right",
      "B. No: thirds are smaller pieces",
      "C. No: 1/3 is 0.3, so it is smaller",
    ],
    answer: "B.  More pieces means smaller pieces.",
    answerColor: C.ALERT,
    answerSize: 20,
    notes: NOTES_CFU2,
  });

  // 16. You Do
  lib.youDoSlide(pres, {
    title: "Your turn: rename, compare, order",
    steps: [
      { label: "First:", text: "Section 1, use the wall." },
      { label: "Next:", text: "Section 2, compare." },
      { label: "Then:", text: "Section 3, mark and order." },
    ],
    panelTitle: "Every comparison, the same three moves",
    reminders: [
      { text: ANCHOR_PHRASE, color: C.PRIMARY },
      { text: "Rename to the same bottom number first.", color: C.SECONDARY },
      { text: "Finished? Try the Challenge, then the Year 8 sheet.", color: C.ALERT },
    ],
    notes: NOTES_YOUDO,
  });

  // 17. Exit Ticket
  exitTicketSlide(pres,
    [
      "Write two fractions that are equivalent to 1/2.",
      "Put these in order, smallest first: 2/3, 1/4, 1/2. Explain how you know.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // 18. Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what do you always do before comparing two fractions?",
      scItems: [
        "I can find a fraction that is equal to one half.",
        "I can rename two fractions so they have the same bottom number and say which is larger.",
        "I can put several fractions in order on one number line and justify the order.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Session 2 - One Line One Whole.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  await generatePdfs();
  console.log("Session 2 build complete.");
}

/* ── PDFs ─────────────────────────────────────────────────────────────────── */

const ANCHOR_TIP =
  "Same whole. Same line. Then compare. Every row of the wall is the same whole, cut into a " +
  "different number of pieces. To compare, rename both so they share a bottom number.";

// The printed twin of the recap slide's wall. Rows are drawn one call each so
// every row is a separate whole with its own label.
function drawWallPdf(doc, y, opts) {
  const o = opts || {};
  const rows = [
    { parts: 1, label: "1 whole" },
    { parts: 2, label: "halves" },
    { parts: 3, label: "thirds" },
    { parts: 4, label: "quarters" },
    { parts: 6, label: "sixths" },
    { parts: 12, label: "twelfths" },
  ];
  let cy = y;
  rows.forEach((r) => {
    cy = addFractionStripsPdf(doc, cy, 1, r.parts, {
      x: o.x || 60, stripW: o.stripW || 340, stripH: o.stripH || 16,
      gap: 0, label: r.label,
    }) - 2;
  });
  return cy + 6;
}

async function generatePdfs() {
  // -- Worksheet -------------------------------------------------------------
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Use the wall, rename to compare, then mark and order on one line.",
      color: C.PRIMARY,
      lessonInfo: "Session 2 of 2 | Year 5/6 Numeracy",
    });

    y = addTipBox(doc, ANCHOR_TIP, y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 - Read it off the wall", y, { color: C.PRIMARY });
    y = drawWallPdf(doc, y + 2);

    const writeItem = (label) => { y = addWriteLine(doc, label, y) + 1; };
    y = addBodyText(doc,
      "Worked for you:   1/2 = 2/4   (two quarters stop in exactly the same place as one half)",
      y, { italic: true, color: "5F6E7A" });
    writeItem("a)   1/2 = ___ sixths = ___ twelfths");
    writeItem("b)   1/3 = ___ sixths = ___ twelfths");
    writeItem("c)   3/4 = ___ twelfths");

    y = addSectionHeading(doc, "Section 2 - Which is bigger? Rename first", y, { color: C.PRIMARY });
    y = addBodyText(doc,
      "Give both fractions the same bottom number, then circle the bigger one.",
      y, { italic: true, color: "5F6E7A" });
    writeItem("d)   2/3 or 3/5?   Rename both, then circle the bigger one:");
    writeItem("e)   5/6 or 3/4?   Rename both, then circle the bigger one:");
    writeItem("f)   1/4 or 2/6?   Rename both, then circle the bigger one:");

    y = addSectionHeading(doc, "Section 3 - Mark them, then order them", y, { color: C.PRIMARY });
    y = addBodyText(doc,
      "Mark 1/2, 1/4, 5/6 and 2/3 on the line below. The small ticks are twelfths.", y);
    y = addNumberLinePdf(doc, y + 14, 1, 12) + 4;
    writeItem("g)   Smallest to largest:");

    y = addSectionHeading(doc, "Challenge", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "A number line keeps going past 1. Write 7/4 as a mixed numeral, then say where it would sit.", y);
    writeItem("      7/4 as a mixed numeral, and where it sits:");

    addPdfFooter(doc, "Session 2 | Fractions & Decimals | Year 5/6 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // -- Answer key ------------------------------------------------------------
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Session 2 worksheet.",
      color: C.PRIMARY,
      lessonInfo: "Session 2 of 2 | Year 5/6 Numeracy",
    });

    y = addSectionHeading(doc, "Section 1 - Read it off the wall", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1/2 = 3 sixths = 6 twelfths", y);
    y = addBodyText(doc, "b)  1/3 = 2 sixths = 4 twelfths", y);
    y = addBodyText(doc, "c)  3/4 = 9 twelfths", y);

    y = addSectionHeading(doc, "Section 2 - Which is bigger?", y, { color: C.PRIMARY });
    y = addBodyText(doc, "d)  2/3 is bigger.   In fifteenths: 10/15 against 9/15.", y);
    y = addBodyText(doc, "e)  5/6 is bigger.   In twelfths: 10/12 against 9/12.", y);
    y = addBodyText(doc, "f)  2/6 is bigger.   In twelfths: 4/12 against 3/12.", y);

    y = addSectionHeading(doc, "Section 3 - Mark them, then order them", y, { color: C.PRIMARY });
    y = addBodyText(doc,
      "In twelfths: 1/4 = 3/12, 1/2 = 6/12, 2/3 = 8/12, 5/6 = 10/12.", y);
    y = addBodyText(doc, "g)  Smallest to largest: 1/4, then 1/2, then 2/3, then 5/6.", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "7/4 is 1 3/4, because 4/4 makes one whole and 3/4 is left over. On a number line it sits " +
      "three quarters of the way between 1 and 2. Accept a drawn line with the point marked.", y);

    y = addTipBox(doc,
      "Watch for: comparing the bottom numbers as if they were whole numbers (1/3 read as bigger than " +
      "1/2); renaming only one of the two fractions in Section 2; and spacing the marks by eye in " +
      "Section 3 instead of counting twelfths. Section 1 is the re-grounding task - a student who " +
      "missed Session 1 should be able to do it from the printed wall alone.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, "Session 2 | Answer Key | Year 5/6 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // -- Fraction Wall Helper (enabling scaffold + glue-in reference) ----------
  await (async () => {
    const doc = createPdf({ title: SCAFFOLD_RES.name });
    let y = addPdfHeader(doc, SCAFFOLD_RES.name, {
      subtitle: "Same whole. Same line. Then compare.",
      color: C.SECONDARY,
      lessonInfo: "Session 2 of 2 | Year 5/6 Numeracy",
    });

    y = addSectionHeading(doc, "Your fraction wall", y, { color: C.SECONDARY });
    y = drawWallPdf(doc, y + 2, { stripH: 20, stripW: 360 });

    y = addSectionHeading(doc, "The same fractions, as points on one line", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "The wall shows how much of the whole is coloured. The line shows where the fraction sits. " +
      "The small ticks are twelfths.", y, { italic: true, color: "5F6E7A" });
    y = addNumberLinePdf(doc, y + 14, 1, 12) + 4;

    y = addTipBox(doc,
      "To compare two fractions, give them the same bottom number first. Twelfths work for halves, " +
      "thirds, quarters and sixths, so twelfths are a safe choice every time.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "One worked comparison", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Which is bigger, 2/3 or 3/4?   Run a finger down the wall. 2/3 is 8 twelfths. 3/4 is 9 twelfths. " +
      "9 twelfths is further along, so 3/4 is bigger.", y);

    y = addSectionHeading(doc, "Now your turn - the wall is drawn for you above", y, { color: C.PRIMARY });
    const writeItem = (label) => { y = addWriteLine(doc, label, y) + 2; };
    y = addBodyText(doc,
      "For each pair: rename both in twelfths, then circle the bigger one.",
      y, { italic: true, color: "5F6E7A" });
    writeItem("a)   1/2 or 1/3?   1/2 = 6 twelfths, and 1/3 = ___ twelfths.   Bigger:");
    writeItem("b)   3/4 or 5/6?   3/4 = ___ twelfths, and 5/6 = ___ twelfths.   Bigger:");
    writeItem("c)   2/3 or 1/2?   2/3 = ___ twelfths, and 1/2 = ___ twelfths.   Bigger:");

    y = addTipBox(doc,
      "Teacher answers: a) 1/3 = 4 twelfths, so 1/2 is bigger. b) 3/4 = 9 and 5/6 = 10, so 5/6 is bigger. " +
      "c) 2/3 = 8 and 1/2 = 6, so 2/3 is bigger. Glue this card into the maths book.",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, "Session 2 | Fraction Wall Helper | Year 5/6 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, SCAFFOLD_RES.fileName));
    console.log("PDF written: " + SCAFFOLD_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
