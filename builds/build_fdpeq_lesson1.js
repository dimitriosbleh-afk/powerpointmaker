"use strict";

// Fractions, Decimals & Percentages (Year 5/6 composite Numeracy) - Session 1 of 2.
//
// "100% IS THE WHOLE." Recognise that 100% represents the complete whole, use
// percentages to describe and compare relative size, and connect familiar
// percentages to their decimal and fraction equivalents.
//
// Victorian Curriculum 2.0: Mathematics, Number, Year 5/6.
// Codes supplied by the teacher: VC2M5N04, VC2M6N03.
//
// COMPRESSION: the teacher planned five days and has two. Session 1 carries the
// percentage strand; Session 2 carries fraction equivalence and ordering. What
// was cut is named on the teacher-facing overview slide, not hidden.
//
// UNIT ANCHOR (megaprompt section 79, held identical in both sessions):
//   "Same whole. Same line. Then compare."
//   Drawn by builds/fdpeq_lib.js so it cannot drift between sessions.
//
// Daily Review: Fractions & Decimals (teacher-supplied focus, days 1-2).
// Fluency: subtraction vertical algorithm with decimals (teacher-supplied focus).
//
// Catch-up: the launch is answerable by eye, the anchor slide rebuilds the whole
// idea from scratch, and Section 1 of the worksheet re-grounds it on paper.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const { composeGlanceNotes } = require("../themes/core/composeNotes");
const { createLib, ANCHOR_PHRASE } = require("./fdpeq_lib");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addHundredGridPdf,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

// Variant 3 (Ocean Logic), fixed for both sessions of the unit. Variants 0-3
// keep ACCENT contrast-safe for text on fills at this band.
const UNIT_VARIANT = 3;
const T = createTheme("numeracy", "grade56", UNIT_VARIANT);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, closingSlide, keyWordSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  dailyReviewSlide, addRevealAnswerBar,
  addPlaceValueChart, addAreaModel, addTenthsStrip,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  clickBuild, runSlideDiagnostics,
} = T;

const SESSION = 1;
const FOOTER = "Fractions & Decimals | Session 1 of 2 | Year 5/6 Numeracy";
const OUT_DIR = "output/FDPEQ_Lesson1_Percentages_Are_Out_Of_100";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const lib = createLib(T, FOOTER);

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Shade the hundred grid, then name the same amount as a fraction, a decimal and a percentage.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Session 1 worksheet, with the errors to watch for.");
const EXT_RES = makeSessionResource(SESSION,
  "Year 8 Extension",
  "Percentage increase and decrease, reverse percentages, and recurring decimals. Use across both sessions.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXT_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

/* ── Teacher notes (Glance Format, megaprompt sections 45-47) ─────────────── */

const NOTES_TITLE =
  "Open Session 1 of 2. Name the two-day focus in one line, then move to Teacher Resources.";

const NOTES_RESOURCES =
  "Prep slide. Print one worksheet and one answer key per student. Keep the Year 8 " +
  "extension on hand for early finishers across BOTH sessions - it is not reprinted tomorrow.\n" +
  "Have whiteboards, markers and a blank hundred grid on the board.\n" +
  "CATCH-UP: a student who has missed work can start here. The launch is answerable just " +
  "by looking, the anchor slide rebuilds the whole idea, and Section 1 of the worksheet " +
  "re-grounds it on paper. Nothing today assumes an earlier lesson.";

const NOTES_OVERVIEW =
  "For you, not the class. Skip past it when projecting.\n" +
  "Five planned days compressed into two. Today is the percentage strand; tomorrow is " +
  "fraction equivalence and ordering on one number line. What was cut is listed on the slide.\n" +
  "Decision points: the boards check after the two models, the hinge after guided practice, " +
  "and the exit ticket. Between them keep the pace brisk.\n" +
  "SOURCES: teacher-supplied content descriptions VC2M5N04 and VC2M6N03, and the supplied " +
  "Daily Review and Number Fluency focus lists.";

const NOTES_DR = composeGlanceNotes({
  answer: "0.6, then 4/10 which is 2/5, then 3.7.",
  beats: [
    "POINT to the place value chart. SAY: Line the places up before you decide.",
    [
      "ASK: Write all three answers.",
      "90 sec. Cue: Write it... chin it... show me.",
      "EXPECT: 0.6, 4/10 or 2/5, 3.7.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> reveal, tick and fix, straight into fluency.",
      "Less -> stack 0.60 above 0.56, compare tenths, re-ask.",
    ],
  ],
  trap: [
    "choosing 0.56 because it has more digits.",
    "Fix: write 0.6 as 0.60, student re-decides.",
  ],
  prep: "Retrieval of Year 5 fractions and decimals. Comparing decimals feeds straight into today's percentage work. Whole block under 5 minutes.",
  tag: "[Stage 1 | Daily Review | Retention and recall | HITS 6]",
});

const NOTES_FLUENCY = composeGlanceNotes({
  answer: "4.7, 7.75 and 13.62.",
  beats: [
    "SAY: Fluency. Set each one out in columns and line up the points.",
    [
      "SAY: Fill any empty place with a zero before you start.",
      "SAY: 20 becomes 20.00, so every column has a digit.",
    ],
    [
      "ASK: What are the three answers?",
      "90 sec. Boards up on cue.",
      "EXPECT: 4.7, 7.75, 13.62.",
    ],
    [
      "SCAN boards.",
      "80%+ -> reveal, tick and fix.",
      "Less -> rebuild 20 minus 6.38 with the zeros written in, re-ask.",
    ],
  ],
  trap: [
    "subtracting with the empty places left blank.",
    "Fix: write the zeros, student redoes the column.",
  ],
  prep: "The fluency focus for both sessions: vertical subtraction with decimals. Brisk retrieval, under 4 minutes, not new teaching.",
  tag: "[Stage 1 | Fluency | Retention and recall | HITS 6]",
});

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "Grid B. It has 65 squares shaded and Grid A has 40.",
  beats: [
    "POINT to both grids. SAY: Same whole. Each grid is 100 squares.",
    [
      "ASK: Which grid shows more of the whole?",
      "20 sec. Thumbs only, voices off. Show me... now.",
      "EXPECT: Grid B, the one on the right.",
    ],
    [
      "SCAN the room.",
      "80%+ -> cold call one: How did you decide so fast?",
      "Less -> count the full columns in tens together, re-ask.",
    ],
    [
      "REVEAL after every thumb is up.",
      "SAY: 40 out of 100. And 65 out of 100.",
    ],
  ],
  trap: [
    "counting the shaded squares one at a time.",
    "Fix: count full columns in tens, student recounts.",
  ],
  prep: "Low-coupling launch: answerable by looking, so a student who missed earlier lessons joins here. It hands you the words out of 100.",
  tag: "[Launch | Knowledge and memory | HITS 2, 6]",
});

const NOTES_LI_SC = composeGlanceNotes({
  beats: [
    "POINT to the learning intention. SAY: Today, 100% means the whole thing.",
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: The first one everyone can do - say what 100% means.",
  ],
  prep: "SC1 is reachable by all; SC2 is the core target the exit ticket assesses; SC3 stretches to comparing and explaining. Tier labels stay off the slide.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_VOCAB = composeGlanceNotes({
  answer: "per cent means out of 100.",
  beats: [
    [
      "SAY: Cent means hundred. A century is 100 years.",
      "SAY: So per cent means for every hundred.",
    ],
    [
      "POINT to 50%.",
      "SAY: That says 50 for every 100. Fifty of the hundred squares.",
    ],
    [
      "ASK: What does 100% mean?",
      "10 sec. Choral response: everyone, together, on three.",
      "EXPECT: all of it, the whole thing.",
    ],
  ],
  trap: [
    "hearing per cent as a brand new kind of number.",
    "Fix: cover the % sign, read out of 100, student re-says it.",
  ],
  prep: "The one word that unlocks the session. Students who hear out of 100 inside the word stop treating percentages as separate from fractions.",
  tag: "[Key Vocabulary | Knowledge and memory | HITS 6]",
});

const NOTES_ANCHOR = composeGlanceNotes({
  answer: "one point on the line has three names: a percent, a decimal, a fraction.",
  beats: [
    "POINT across the three moves. SAY: Same whole. Same line. Then compare.",
    [
      "POINT to the middle mark.",
      "SAY: One point. Fifty per cent, nought point five, one half.",
    ],
    [
      "ASK: Say the three moves back to me.",
      "5 sec. Choral response: everyone, together, on three.",
      "EXPECT: same whole, same line, then compare.",
    ],
  ],
  trap: [
    "thinking a percent must be bigger because its number is bigger.",
    "Fix: point to the one mark, student names all three.",
  ],
  stretch: "where would 10% sit on this line? Point to it.",
  help: "cover the decimal and fraction rows, read the percents only.",
  prep: "The unit anchor, introduced here and restated in every I Do across both sessions. A returning student can re-enter the unit from this slide alone.",
  tag: "[Anchor | Knowledge and memory | HITS 3, 6]",
});

const NOTES_IDO1 = composeGlanceNotes({
  answer: "1/2 = 0.5 = 50%. Fifty of the hundred squares.",
  beats: [
    "SAY: Watch how I name this shaded part three ways.",
    [
      "SAY: Same whole first. The whole is all 100 squares.",
      "SAY: Half of 100 is 50, so 50 are shaded.",
    ],
    [
      "SAY: As a fraction that is 50 out of 100.",
      "SAY: As a decimal, 50 hundredths is 0.5.",
    ],
    "REVEAL the chips after the working is said. SAY: Out of 100, so 50%.",
  ],
  trap: [
    "reading 0.50 as fifty rather than a half.",
    "Fix: shade the same half on the bar, student re-reads.",
  ],
  stretch: "name the unshaded half all three ways too.",
  help: "count the shaded columns in tens first.",
  prep: "First model. The anchor is restated in full so a returning student can follow from this slide alone.",
  tag: "[Stage 2 | I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_IDO2 = composeGlanceNotes({
  answer: "1/10 = 0.1 = 10%. One column of the hundred.",
  beats: [
    "SAY: Same routine, smaller piece. One column this time.",
    [
      "SAY: One column is 10 squares out of the 100.",
      "SAY: Ten of these columns make the whole, so it is one tenth.",
    ],
    "REVEAL the chips after the working is said. SAY: 10 out of 100, so 10%.",
    [
      "SAY: Ten per cent is the one to remember.",
      "SAY: Ten of them make one hundred per cent.",
    ],
  ],
  trap: [
    "writing 1/10 as 0.10 and reading it as ten.",
    "Fix: point to one column of the hundred, student re-reads.",
  ],
  stretch: "how many of these columns make 40%?",
  help: "the hundred grid with the columns already ruled off.",
  prep: "Second model, deliberately a smaller piece so the routine is copied, not the number. Ten per cent is the workhorse for tomorrow.",
  tag: "[Stage 2 | I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_CFU1 = composeGlanceNotes({
  answer: "30%. Nought point two five is 25%, and 25 is less than 30.",
  beats: [
    "SAY: Check on your boards. Put both onto the same line first.",
    [
      "ASK: Which is more, 0.25 or 30%?",
      "40 sec. Cue: Write it... chin it... show me.",
      "EXPECT: 30%.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one: How did you change 0.25 into a percent?",
      "Less -> shade 25 squares and 30 squares, re-ask with 0.4 or 35%.",
    ],
    "REVEAL only after every board is up.",
  ],
  trap: [
    "reading 0.25 as bigger because it looks like twenty five.",
    "Fix: shade both on the grid, student re-decides.",
  ],
  prep: "Decision point one. Checks that a decimal and a percent can be compared before releasing into guided practice.",
  tag: "[Stage 2 | CFU | Supported application | HITS 7, 8]",
});

const NOTES_WEDO1 = composeGlanceNotes({
  answer: "3/4 = 0.75 = 75%. Seventy five of the hundred squares.",
  beats: [
    "SAY: Your turn with me. Same whole, same line, then name it.",
    [
      "ASK: Write this amount three ways.",
      "60 sec with your partner. Boards up on cue.",
      "EXPECT: 3/4, 0.75 and 75%.",
    ],
    [
      "CIRCULATE. Look for the count written before the answer.",
      "80%+ -> reveal, then bounce: Do you agree? Add one thing.",
      "Less -> count the shaded columns in tens together, re-ask.",
    ],
  ],
  trap: [
    "writing 75/100 then stalling on the decimal.",
    "Fix: read it as seventy five hundredths, student writes 0.75.",
  ],
  stretch: "what percent is NOT shaded? Prove it two ways.",
  help: "the fraction 75/100 already written, student names the rest.",
  prep: "Guided practice, faded from the I Do: no worked steps stay on screen and the count is theirs to do.",
  tag: "[Stage 3 | We Do | Supported application | HITS 4, 5]",
});

const NOTES_WEDO2 = composeGlanceNotes({
  answer: "Class A. Three fifths is 60%, and 60 beats 55.",
  beats: [
    [
      "SAY: Two classes, two different ways of saying it.",
      "SAY: I cannot compare 3/5 with 55% yet.",
    ],
    "SAY: Rename the fifths into tenths, then into hundredths.",
    [
      "ASK: Who read more, Class A or Class B?",
      "60 sec with your partner. Boards up on cue.",
      "EXPECT: Class A, because 3/5 is 60%.",
    ],
    [
      "CIRCULATE. Check the rename happens first.",
      "80%+ -> reveal, then on to the hinge.",
      "Less -> shade 6 tenths, count to 60, re-ask.",
    ],
  ],
  trap: [
    "picking Class B because 55 looks bigger than 3.",
    "Fix: rename 3/5 to 60%, student re-decides.",
  ],
  stretch: "what percent beats 3/4?",
  help: "tenths strip with 6 parts shaded.",
  prep: "Flex slide - cut this if the models ran long. The exit ticket still reaches comparing across forms.",
  tag: "[Stage 3 | We Do | Supported application | HITS 4, 5]",
});

const NOTES_CFU2 = composeGlanceNotes({
  answer: "B. One quarter is 0.25, which is 25 out of 100.",
  beats: [
    "SAY: Hinge question. A student wrote this. Decide if it is right.",
    [
      "ASK: Is it A, B or C?",
      "30 sec, voices off. Cue: Write it... chin it... show me.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: Convince us. Then reveal.",
      "Less -> shade a quarter of the grid, count 25, re-ask with 3/4.",
    ],
    "REVEAL only after every board is up.",
  ],
  trap: [
    "reading the 4 in 1/4 as a tenths digit, giving 0.4.",
    "Fix: quarter the grid, count the squares, student rewrites.",
  ],
  prep: "Decision point two, the hinge of the lesson. A means no check at all. C means the digits were shuffled, not converted.",
  tag: "[Stage 2 | CFU hinge | Supported application | HITS 7, 8]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "open - listen for out of 100 said before any conversion is written.",
  beats: [
    "SAY: On your own now. Section 1 rebuilds it, then it builds up.",
    "SAY: Every question, same three moves. Shade it if you are stuck.",
    [
      "CIRCULATE. Sit with the catch-up group first.",
      "COLLECT: check Section 1 before anyone starts Section 3.",
    ],
    "TIME: about 12 minutes, then stop for the exit ticket.",
  ],
  stretch: "the Challenge box, then the Year 8 extension sheet.",
  help: "Section 1 worked with you, using the printed hundred grid.",
  prep: "Section 1 is the re-grounding task: doable from the anchor slide alone, with no earlier session needed.",
  tag: "[Stage 4 | You Do | Mastery and application | HITS 4, 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "0.75 and 75%. Then 0.7, because 0.7 is 70% and 70 beats 65.",
  beats: [
    "SAY: On your own, on your whiteboard. No partner talk for this one.",
    [
      "SAY: One, write three quarters as a decimal and a percentage.",
      "SAY: Two, decide which is more and say how you know.",
    ],
    [
      "TIME: about 3 minutes.",
      "COLLECT boards, or photograph them for your records.",
    ],
  ],
  trap: [
    "choosing 65% because 65 looks bigger than 0.7.",
    "Fix: note the name, reteach on the grid at the start of Session 2.",
  ],
  prep: "Assesses the core target (connecting the three forms) and reaches into comparing. Keep the SC number off the slide.",
  tag: "[Stage 5 | Exit Ticket | Mastery and application | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - listen for 100% is the whole, and per cent means out of 100.",
  beats: [
    "SAY: Look back at the three I can statements. Thumbs for each one.",
    [
      "ASK: What does 100% mean?",
      "20 sec. Turn and tell. Partner B first.",
      "EXPECT: the whole thing, all of it.",
    ],
    "COLLECT: note who showed need more practice, ready for tomorrow.",
  ],
  prep: "Close of day one. The keeper: per cent means out of 100, so every amount can sit on the same line. Tomorrow uses that line for fractions.",
  tag: "[Closing | Retention and recall | HITS 9]",
});

/* ── Local visual: the hundred-grid naming panel ──────────────────────────── */

// A hundred grid with the same amount named three ways underneath. The three
// chips hold back until the click, so students commit before they see the
// conversion. Returns reveal handles so the click build can stage them.
function fdpPanel(slide, lg, opts) {
  const o = opts || {};
  const x = lg.rightX;
  const w = lg.rightW;
  const y0 = lg.panelTopPadded;
  const bottom = lg.safeBottom;

  addCard(slide, x, y0, w, bottom - y0, { strip: C.PRIMARY });
  slide.addText(o.title, {
    x: x + 0.18, y: y0 + 0.10, w: w - 0.36, h: 0.28,
    fontSize: 14.5, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
  });

  const gridSize = 1.75;
  const gridX = x + (w - gridSize) / 2;
  const gridY = y0 + 0.48;
  addAreaModel(slide, gridX, gridY, gridSize, o.cols, o.extra || 0);

  slide.addText(o.caption, {
    x: x + 0.18, y: gridY + gridSize + 0.08, w: w - 0.36, h: 0.28,
    fontSize: 13.5, fontFace: FONT_B, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
  });

  const chipY = y0 + 2.67;
  const chipW = 1.20;
  const chipGap = 0.12;
  const chipX0 = x + (w - (chipW * 3 + chipGap * 2)) / 2;
  const chipFills = [C.SECONDARY, C.CHARCOAL, C.SUCCESS];

  return {
    revealChip(i) {
      addTextOnShape(slide, o.chips[i], {
        x: chipX0 + i * (chipW + chipGap), y: chipY, w: chipW, h: 0.55, rectRadius: 0.08,
        fill: { color: chipFills[i] },
      }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
    },
    revealStrap() {
      slide.addText("The same amount, three ways.", {
        x: x + 0.18, y: chipY + 0.66, w: w - 0.36, h: 0.30,
        fontSize: 13.5, fontFace: FONT_B, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
    },
  };
}

/* ── Build ────────────────────────────────────────────────────────────────── */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  titleSlide(pres, "100% is the Whole",
    "Fractions, decimals and percentages: three ways to say the same amount",
    "Year 5/6 Numeracy | Session 1 of 2", NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3. Teacher-facing overview
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.MUTED);
    addBadge(s, "For the teacher", { color: C.MUTED, w: 2.4 });
    addTitle(s, "Five days into two", { color: C.MUTED });

    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;

    addCard(s, 0.5, cardY, 4.4, cardH, { strip: C.PRIMARY });
    s.addText("What the two days carry", {
      x: 0.7, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Session 1  100% is the whole", options: { bold: true, breakLine: true } },
      { text: "Connect fractions, decimals and percentages.", options: { breakLine: true } },
      { text: "Session 2  One line, one whole", options: { bold: true, breakLine: true } },
      { text: "Compare and order fractions using equivalence.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Cut to fit two days:", options: { color: C.MUTED, breakLine: true } },
      { text: "separate lessons on tenths and on hundredths; mixed numerals now sit in the challenge and extension only.", options: { breakLine: true } },
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
      { text: "Shape: example-first. Two models, guided practice, then independent.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Decision points:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "boards check after the models, hinge after guided practice, exit ticket.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Catch-up:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "the launch reads by eye and the anchor slide rebuilds the whole idea.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Flex: cut the second We Do if the models ran long.", options: { italic: true, color: C.MUTED, breakLine: true } },
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

  // 4. Daily Review - fractions and decimals (supplied focus)
  (() => {
    const s = dailyReviewSlide(pres, "Daily Review: Fractions and decimals",
      [
        "Which is larger: 0.6 or 0.56?",
        "Write 0.4 as a fraction.",
        "Round 3.68 to the nearest tenth.",
      ],
      NOTES_DR, FOOTER,
      (slide, lg) => {
        // Panel stops at 4.15 so the click-revealed answer bar (top 4.30)
        // keeps its clearance.
        const y0 = lg.panelTopPadded;
        const cardBottom = 4.15;
        addCard(slide, lg.rightX, y0, lg.rightW, cardBottom - y0, { strip: C.ACCENT });
        slide.addText("Line the places up", {
          x: lg.rightX + 0.15, y: y0 + 0.10, w: lg.rightW - 0.30, h: 0.28,
          fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, align: "center", margin: 0,
        });
        addPlaceValueChart(slide, lg.rightX + 0.25, y0 + 0.52,
          ["ones", "tenths", "hundredths"],
          [null, null, null],
          { w: lg.rightW - 0.50, hdrH: 0.44, valH: 0.75, headerColor: C.PRIMARY });
        addTextOnShape(slide, "0.6 = 0.60", {
          x: lg.rightX + 0.60, y: y0 + 1.92, w: lg.rightW - 1.20, h: 0.60, rectRadius: 0.08,
          fill: { color: C.SECONDARY },
        }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
    clickBuild(s, [
      () => { addRevealAnswerBar(s, ["0.6", "4/10 = 2/5", "3.7"], { color: C.SUCCESS }); },
    ]);
    runSlideDiagnostics(s, pres);
  })();

  // 5. Fluency - subtraction vertical algorithm with decimals (supplied focus)
  lib.fluencySlide(pres, "Fluency: Subtraction, set out in columns",
    ["8.6 - 3.9", "15.4 - 7.65", "20 - 6.38"],
    ["4.7", "7.75", "13.62"],
    NOTES_FLUENCY);

  // 6. Launch - which grid shows more? (answerable by eye)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "Launch", { color: C.SECONDARY });
    addTitle(s, "Which shows more of the whole?", { color: C.SECONDARY });

    addAreaModel(s, 1.45, 1.50, 2.05, 4, 0);
    addAreaModel(s, 6.45, 1.50, 2.05, 6, 5);

    ["Grid A", "Grid B"].forEach((label, i) => {
      s.addText(label, {
        x: 1.45 + i * 5.0, y: 3.64, w: 2.05, h: 0.30,
        fontSize: 17, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    // One chip per grid rather than a single answer bar: the count belongs
    // beside the grid it counts, and a shared bar wraps to two lines here.
    clickBuild(s, [
      () => {
        ["40 out of 100 = 40%", "65 out of 100 = 65%"].forEach((txt, i) => {
          addTextOnShape(s, txt, {
            x: 1.05 + i * 5.0, y: 4.15, w: 2.85, h: 0.70, rectRadius: 0.1,
            fill: { color: C.SUCCESS },
          }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
        });
      },
    ]);
    runSlideDiagnostics(s, pres);
  })();

  // 7. Learning Intention and Success Criteria
  liSlide(pres,
    "We are learning that 100% means one whole, so a fraction, a decimal and a percentage can name the same amount.",
    [
      "I can say what 100% means.",
      "I can write the same amount as a fraction, a decimal and a percentage.",
      "I can compare two amounts written in different ways and explain how I know.",
    ],
    NOTES_LI_SC, FOOTER);

  // 8. Key vocabulary
  keyWordSlide(pres, {
    word: "per cent",
    meaning: "Out of 100. The % sign is a short way of writing out of 100.",
    example: "50% means 50 out of every 100. 100% means the whole thing.",
    routine: ["Say it", "Show it", "Use it"],
    color: C.PRIMARY,
  }, NOTES_VOCAB, FOOTER);

  // 9. Anchor - same whole, same line (the catch-up re-entry slide)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Anchor", { color: C.PRIMARY });
    addTitle(s, "Same whole. Same line.", { color: C.PRIMARY });

    lib.anchorMoves(s, CONTENT_TOP);

    s.addText("The same point. Three ways to say it.", {
      x: 0.5, y: 2.02, w: 9, h: 0.30,
      fontSize: 15, fontFace: FONT_B, color: C.MUTED, bold: true,
      align: "center", margin: 0,
    });

    lib.tripleScale(s, 0.8, 2.50, 8.4);

    addTextOnShape(s, "100%, 1 whole and 1.0 all name the same amount.", {
      x: 0.5, y: 4.38, w: 9, h: 0.60, rectRadius: 0.1,
      fill: { color: C.SUCCESS },
    }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_ANCHOR);
    runSlideDiagnostics(s, pres);
  })();

  // 10. I Do - one half, three ways
  (() => {
    let panel;
    const s = workedExSlide(pres, 2, "I Do", "One half, three ways",
      [
        "The whole is 100 squares.",
        "Half of it is 50 squares.",
        "50 out of 100 is 50/100.",
        "50 hundredths is 0.5.",
        "Out of 100 means per cent.",
        "So it is 50%.",
      ],
      NOTES_IDO1, FOOTER,
      (slide, lg) => {
        panel = fdpPanel(slide, lg, {
          title: "50 squares shaded", cols: 5, extra: 0,
          caption: "50 out of 100",
          chips: ["1/2", "0.5", "50%"],
        });
      });
    clickBuild(s, [
      () => { panel.revealChip(0); panel.revealChip(1); },
      () => { panel.revealChip(2); panel.revealStrap(); },
    ]);
    runSlideDiagnostics(s, pres);
  })();

  // 11. I Do - one tenth, three ways
  (() => {
    let panel;
    const s = workedExSlide(pres, 2, "I Do", "One tenth, three ways",
      [
        "Same whole: 100 squares.",
        "One column is shaded.",
        "That is 10 out of 100.",
        "Ten columns make the whole.",
        "So the fraction is 1/10.",
        "As a decimal, 0.1.",
      ],
      NOTES_IDO2, FOOTER,
      (slide, lg) => {
        panel = fdpPanel(slide, lg, {
          title: "10 squares shaded", cols: 1, extra: 0,
          caption: "10 out of 100",
          chips: ["1/10", "0.1", "10%"],
        });
      });
    clickBuild(s, [
      () => { panel.revealChip(0); panel.revealChip(1); },
      () => { panel.revealChip(2); panel.revealStrap(); },
    ]);
    runSlideDiagnostics(s, pres);
  })();

  // 12. CFU - decision point one
  lib.checkSlide(pres, {
    title: "Compare across forms",
    technique: "Show Me Boards",
    question: "Which is more: 0.25 or 30%?",
    answer: "30%.  0.25 is only 25 out of 100.",
    notes: NOTES_CFU1,
  });

  // 13. We Do - three quarters, three ways
  (() => {
    let panel;
    const s = workedExSlide(pres, 3, "We Do", "Name this amount three ways",
      [
        "With your partner.",
        "",
        "Same whole: 100 squares.",
        "Count the shaded squares.",
        "Write the fraction first.",
        "Then the decimal, then the percent.",
      ],
      NOTES_WEDO1, FOOTER,
      (slide, lg) => {
        panel = fdpPanel(slide, lg, {
          title: "Your turn", cols: 7, extra: 5,
          caption: "How many out of 100?",
          chips: ["3/4", "0.75", "75%"],
        });
      });
    clickBuild(s, [
      () => { panel.revealChip(0); panel.revealChip(1); panel.revealChip(2); panel.revealStrap(); },
    ]);
    runSlideDiagnostics(s, pres);
  })();

  // 14. We Do - compare relative size across forms (flex slide)
  (() => {
    let revealA;
    let revealVerdict;
    const s = workedExSlide(pres, 3, "We Do", "Who read more: 3/5 or 55%?",
      [
        "Class A read 3/5.",
        "Class B read 55%.",
        "",
        "Same whole: the book.",
        "Rename 3/5 into tenths.",
        "Then into hundredths.",
        "Now compare the percents.",
      ],
      NOTES_WEDO2, FOOTER,
      (slide, lg) => {
        const x = lg.rightX;
        const w = lg.rightW;
        const y0 = lg.panelTopPadded;
        addCard(slide, x, y0, w, lg.safeBottom - y0, { strip: C.SECONDARY });

        slide.addText("Class A: 3 out of 5", {
          x: x + 0.18, y: y0 + 0.10, w: w - 0.36, h: 0.28,
          fontSize: 14.5, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        addTenthsStrip(slide, x + 0.22, y0 + 0.55, 3.0, 6);

        slide.addText("Class B", {
          x: x + 0.18, y: y0 + 1.92, w: w - 0.36, h: 0.28,
          fontSize: 14.5, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        addTextOnShape(slide, "55%", {
          x: x + 0.22, y: y0 + 2.24, w: w - 0.44, h: 0.55, rectRadius: 0.08,
          fill: { color: C.CHARCOAL },
        }, { fontSize: 21, fontFace: FONT_H, color: C.WHITE, bold: true });

        revealA = () => {
          addTextOnShape(slide, "6/10 = 60/100 = 60%", {
            x: x + 0.22, y: y0 + 1.20, w: w - 0.44, h: 0.55, rectRadius: 0.08,
            fill: { color: C.SECONDARY },
          }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
        };
        revealVerdict = () => {
          addTextOnShape(slide, "60% beats 55%. Class A.", {
            x: x + 0.22, y: y0 + 2.95, w: w - 0.44, h: 0.60, rectRadius: 0.08,
            fill: { color: C.SUCCESS },
          }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
        };
      });
    clickBuild(s, [
      () => revealA(),
      () => revealVerdict(),
    ]);
    runSlideDiagnostics(s, pres);
  })();

  // 15. CFU hinge - decision point two
  lib.checkSlide(pres, {
    title: "Spot the error",
    technique: "Show Me Boards",
    lead: "A student wrote:",
    question: "1/4 = 0.4 = 40%",
    questionSize: 34,
    options: [
      "A. Correct",
      "B. Wrong: 1/4 is 0.25 and 25%",
      "C. Wrong: 1/4 is 0.14 and 14%",
    ],
    answer: "B.  1/4 = 25/100 = 0.25 = 25%",
    answerColor: C.ALERT,
    notes: NOTES_CFU2,
  });

  // 16. You Do
  lib.youDoSlide(pres, {
    title: "Your turn: name it and compare it",
    steps: [
      { label: "First:", text: "Section 1, shade and name." },
      { label: "Next:", text: "Section 2, fill the gaps." },
      { label: "Then:", text: "Section 3, compare." },
    ],
    panelTitle: "Every question, the same three moves",
    reminders: [
      { text: ANCHOR_PHRASE, color: C.PRIMARY },
      { text: "Stuck? Shade it on the hundred grid first.", color: C.SECONDARY },
      { text: "Finished? Try the Challenge, then the Year 8 sheet.", color: C.ALERT },
    ],
    notes: NOTES_YOUDO,
  });

  // 17. Exit Ticket
  exitTicketSlide(pres,
    [
      "Write 3/4 as a decimal and as a percentage.",
      "Which is more: 0.7 or 65%? Explain how you know.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // 18. Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what does 100% mean, and how do you know?",
      scItems: [
        "I can say what 100% means.",
        "I can write the same amount as a fraction, a decimal and a percentage.",
        "I can compare two amounts written in different ways and explain how I know.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Session 1 - 100 Per Cent is the Whole.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  await generatePdfs();
  console.log("Session 1 build complete.");
}

/* ── PDFs ─────────────────────────────────────────────────────────────────── */

const ANCHOR_TIP =
  "Same whole. Same line. Then compare. Per cent means out of 100, so 100% is the whole. " +
  "50 out of 100 = 50/100 = 0.5 = 50%.";

async function generatePdfs() {
  // -- Worksheet -------------------------------------------------------------
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Name the same amount as a fraction, a decimal and a percentage.",
      color: C.PRIMARY,
      lessonInfo: "Session 1 of 2 | Year 5/6 Numeracy",
    });

    y = addTipBox(doc, ANCHOR_TIP, y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 - Shade it, then name it", y, { color: C.PRIMARY });
    y = addBodyText(doc,
      "The grid on the left is done for you. Shade 25 squares on the grid on the right.",
      y, { italic: true, color: "5F6E7A" });

    const gridTop = y + 4;
    addHundredGridPdf(doc, gridTop, 50, {
      x: 60, size: 105, fillColor: "#" + C.SECONDARY,
      label: "50/100 = 0.5 = 50%",
    });
    const rightBottom = addHundredGridPdf(doc, gridTop, 0, {
      x: 300, size: 105, fillColor: "#" + C.SECONDARY,
      label: "Shade 25 squares",
    });
    y = rightBottom + 6;

    const writeItem = (label) => { y = addWriteLine(doc, label, y) + 1; };
    writeItem("a)   Your shaded grid, written as a fraction, a decimal and a percentage:");

    y = addSectionHeading(doc, "Section 2 - Fill the gaps", y, { color: C.PRIMARY });
    y = addBodyText(doc,
      "Worked for you:   1/2 = 0.5 = 50%   (one half of 100 squares is 50)",
      y, { italic: true, color: "5F6E7A" });
    writeItem("b)   Write 1/4 as a decimal and as a percentage:");
    writeItem("c)   Write 0.75 as a fraction and as a percentage:");
    writeItem("d)   Write 10% as a fraction and as a decimal:");
    writeItem("e)   Write 3/5 as a percentage. Rename it in hundredths first:");

    y = addSectionHeading(doc, "Section 3 - Which is more?", y, { color: C.PRIMARY });
    y = addBodyText(doc,
      "Put both onto the same line first, then circle the larger one and say how you know.",
      y, { italic: true, color: "5F6E7A" });
    writeItem("f)   0.3   or   35%?   Larger:");
    writeItem("g)   1/2   or   45%?   Larger:");
    writeItem("h)   7/10  or   65%?   Larger:");

    y = addSectionHeading(doc, "Challenge", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Change these into percentages, then order smallest to largest:   0.4,   45%,   1/2,   0.38",
      y);
    writeItem("      Smallest to largest:");

    addPdfFooter(doc, "Session 1 | Fractions & Decimals | Year 5/6 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // -- Answer key ------------------------------------------------------------
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Session 1 worksheet.",
      color: C.PRIMARY,
      lessonInfo: "Session 1 of 2 | Year 5/6 Numeracy",
    });

    y = addSectionHeading(doc, "Section 1 - Shade it, then name it", y, { color: C.PRIMARY });
    y = addBodyText(doc,
      "a)  25 squares shaded.   25/100 = 1/4 = 0.25 = 25%.   Accept 25/100 without simplifying.", y);

    y = addSectionHeading(doc, "Section 2 - Fill the gaps", y, { color: C.PRIMARY });
    y = addBodyText(doc, "b)  1/4 = 0.25 = 25%   (a quarter of 100 squares is 25)", y);
    y = addBodyText(doc, "c)  0.75 = 75/100 = 3/4 = 75%", y);
    y = addBodyText(doc, "d)  10% = 10/100 = 1/10 = 0.1", y);
    y = addBodyText(doc, "e)  3/5 = 6/10 = 60/100 = 60%", y);

    y = addSectionHeading(doc, "Section 3 - Which is more?", y, { color: C.PRIMARY });
    y = addBodyText(doc, "f)  35% is larger.   0.3 is 30%, and 30 is less than 35.", y);
    y = addBodyText(doc, "g)  1/2 is larger.   1/2 is 50%, and 50 is more than 45.", y);
    y = addBodyText(doc, "h)  7/10 is larger.  7/10 is 70%, and 70 is more than 65.", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "As percentages: 0.38 = 38%, 0.4 = 40%, 45%, 1/2 = 50%. " +
      "Smallest to largest: 0.38, then 0.4, then 45%, then 1/2.", y);

    y = addTipBox(doc,
      "Watch for: reading 0.4 as 4% instead of 40%; reading 1/4 as 0.4; and choosing the amount " +
      "with the bigger-looking digits without converting first (Section 3 is where this shows up). " +
      "Section 1 is the re-grounding task - a student who missed earlier work should be able to do " +
      "it from the hundred grid alone.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, "Session 1 | Answer Key | Year 5/6 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // -- Year 8 extension ------------------------------------------------------
  await (async () => {
    const doc = createPdf({ title: EXT_RES.name });
    let y = addPdfHeader(doc, EXT_RES.name, {
      subtitle: "For students ready to push past Year 6. Keep this sheet for both sessions.",
      color: C.ASSESS,
      lessonInfo: "Sessions 1 and 2 | Year 5/6 Numeracy",
    });

    y = addTipBox(doc,
      "Year 8 works with percentage CHANGE, not just percentage of a whole, and with fractions " +
      "that do not convert to neat decimals. Show your working for every question.",
      y, { color: C.ASSESS });

    y = addSectionHeading(doc, "Part A - Percentage increase and decrease", y, { color: C.ASSESS });
    y = addBodyText(doc,
      "Worked for you:   A $40 shirt rises by 10%.   10% of 40 = 4, so the new price is $44. " +
      "You can also do 40 x 1.10 = 44 in one step.",
      y, { italic: true, color: "5F6E7A" });

    // Question on its own line, then a full-width rule: a Year 8 percentage-change
    // question needs room for working, not the tail end of a line.
    const workItem = (question) => {
      y = addBodyText(doc, question, y);
      y = addWriteLine(doc, "      Working and answer:", y) + 2;
    };
    workItem("1)   A jacket costs $80. The price rises by 15%. What is the new price?");
    workItem("2)   A phone is reduced by 20% and now costs $360. What was the original price?");
    workItem("3)   A town grows from 250 people to 290 people. What is the percentage increase?");

    y = addSectionHeading(doc, "Part B - Fractions that are not so neat", y, { color: C.ASSESS });
    y = addBodyText(doc,
      "A fraction is a division. 3/8 means 3 divided by 8. Some divisions stop, and some repeat forever.",
      y, { italic: true, color: "5F6E7A" });
    workItem("4)   Use division to write 3/8 as a decimal. Does it stop or repeat?");
    workItem("5)   Use division to write 2/3 as a decimal. Does it stop or repeat?");
    y = addWriteLine(doc, "6)   Order from smallest to largest:   5/8,   0.6,   62%,   7/12", y) + 2;
    y = addWriteLine(doc, "7)   Find a fraction between 1/3 and 1/2. Prove it:", y) + 2;

    y = addSectionHeading(doc, "Challenge", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "The decimal 0.4444... repeats forever. Call it x. Work out what 10x is, subtract x from it, " +
      "and use the result to write 0.4444... as a fraction.",
      y);
    y = addWriteLine(doc, "      0.4444... as a fraction:", y) + 2;

    y = addTipBox(doc,
      "Teacher answers. 1) $92. 2) $450, because 360 is 80% of the original. 3) 16%, because 40 out of 250 " +
      "is 16 out of 100. 4) 0.375, it stops. 5) 0.666..., it repeats. 6) 7/12 (0.583), 0.6, 62%, 5/8 (0.625). " +
      "7) 5/12 works, because 1/3 is 4/12 and 1/2 is 6/12. Challenge: 10x = 4.4444..., so 9x = 4 and x = 4/9. " +
      "A strong response shows the working, not just the answer.",
      y, { color: C.ASSESS });

    addPdfFooter(doc, "Year 8 Extension | Fractions & Decimals | Year 5/6 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, EXT_RES.fileName));
    console.log("PDF written: " + EXT_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
