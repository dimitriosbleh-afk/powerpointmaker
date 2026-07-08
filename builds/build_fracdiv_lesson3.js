"use strict";

// Year 6 Numeracy — Choosing the operation and solving worded problems that divide
//   with fractions (both types from Sessions 1 and 2).
// Victorian Curriculum 2.0: Mathematics, Number, Year 6 — solve problems involving
//   division with fractions, choosing and justifying the operation.
// Session 3 of the Dividing With Fractions unit. Error-analysis-led shape (Sessions 1
//   and 2 were example-first) so the lesson body varies deliberately (megaprompt s72).
// Centrepiece: the "divide BY a half" (how many halves fit -> bigger) versus "divide IN
//   half" (share into 2 -> smaller) confusion, addressed head on in the I Do, the CFU
//   hinge, SC3, and a spot-the-error task.
// Daily Review: equivalent fractions and simplifying, with visual models.
// Fluency: multiplication and division fact families.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const { composeGlanceNotes } = require("../themes/core/composeNotes");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine, addLinedArea,
  addFractionStripsPdf, addNumberLinePdf,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  dailyReviewSlide, fluencySlide, addRevealAnswerBar,
  addFractionStripSet, addNumberLine, addGroupedCounters,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 3;
const FOOTER = "Fractions | Choosing and Solving | Year 6 Numeracy";
const OUT_DIR = "output/FracDiv_Lesson3_Choosing_And_Solving";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Problem Set",
  "Decide the operation, then solve each fraction division word problem.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers and chosen operations for the problem set.");
const SCAFFOLD_RES = makeSessionResource(SESSION,
  "Enabling Scaffold",
  "Decision chart and pre-drawn models - tick the question type, then solve.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, SCAFFOLD_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher notes (Glance Format, megaprompt sections 45-47) ────────────────

const NOTES_TITLE =
  "Open the lesson. Say the title, then move straight to Teacher Resources.";

const NOTES_RESOURCES =
  "Prep slide. Print the problem set, answer key and enabling scaffold. " +
  "Have paper strips, whiteboards and markers ready. Keep the decision chart on the board.";

const NOTES_DR = composeGlanceNotes({
  answer: "3/4 = 6/8 (six eighths shade the same as three quarters). 9/12 in simplest form is 3/4.",
  beats: [
    "POINT to the two strips. SAY: Equal fractions cover the same length. Three quarters and six eighths line up.",
    "ASK: 3/4 equals how many eighths? 20 sec, boards up. EXPECT: 6, so 3/4 = 6/8.",
    "ASK: Write 9/12 as simply as you can. 30 sec, boards up. EXPECT: 3/4 - divide top and bottom by 3.",
    "SCAN boards. 80%+ -> reveal. Less -> shade 9/12 and 3/4 on strips, match them, re-ask.",
  ],
  trap: "scaling only one number when making eighths. Fix: point to the strips, both parts scale by 2, student redoes.",
  prep: "Prior learning: equivalent fractions and simplifying with visual models. Third day of spaced retrieval on the equal-parts idea.",
  tag: "[Stage 1 | Daily Review | Retention and recall | HITS 6]",
});

const NOTES_FLUENCY = composeGlanceNotes({
  answer: "63, 7, 64.",
  beats: [
    "SAY: Fluency. Say the whole fact family in your head, then write the answer.",
    "ASK: What are the answers? 45 sec, boards up. EXPECT: 63, 7, 64.",
    "SCAN boards, back row first. 80%+ -> reveal. Less -> build 9 x 7 = 63, so 63 divided by 9 = 7, re-ask.",
  ],
  trap: "guessing 63 divided by 9. Fix: use the linked fact 9 x 7 = 63, student states the division.",
  prep: "Brisk fact-family recall (Number). Fluent facts free working memory for choosing operations today.",
  tag: "[Stage 1 | Fluency | Mastery and application | HITS 6]",
});

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "open - listen for 'how many fit makes it bigger; sharing a fraction makes it smaller'.",
  beats: [
    "POINT to the two questions. SAY: We met both of these this week. Same numbers, very different answers.",
    "ASK: Which makes the BIGGER answer - how many halves fit in 4, or share 1/2 among 4? 20 sec, boards up. EXPECT: how many halves fit.",
    "SAY: Today we decide which one a word problem is really asking, then solve it.",
    "SCAN boards. 80%+ -> next slide. Less -> recall Session 1 grew, Session 2 shrank, re-ask.",
  ],
  trap: "assuming every division makes a smaller answer. Fix: recall 3 divided by 1/4 = 12 from Session 1, student re-decides.",
  prep: "Launch retrieves both meanings from Sessions 1 and 2 and bridges to choosing the operation. Assumes both were taught this week.",
  tag: "[Launch | Retention and recall | HITS 2, 6]",
});

const NOTES_LI_SC = composeGlanceNotes({
  beats: [
    "POINT to the learning intention. SAY: Today we choose the right division and solve fraction word problems.",
    "SAY: Read the I can statements with me. Choral read.",
    "SAY: The first one everyone can do today - decide what the problem is asking.",
  ],
  prep: "SC1 achievable by all; SC2 is the core target and the exit ticket; SC3 stretches to explaining the by-a-half versus halving difference.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_IDO = composeGlanceNotes({
  answer: "4 ÷ 1/2 = 8 (eight halves fit in 4 - bigger). 4 ÷ 2 = 2 (4 shared into 2 - smaller). They are different.",
  beats: [
    "SHOW both problems. SAY: These look alike but they are not. Watch the difference.",
    "MODEL 4 ÷ 1/2. SAY: Dividing BY a half asks how many halves fit in 4. Each whole holds 2, so 4 holds 8.",
    "MODEL 4 ÷ 2. SAY: Dividing by 2 shares 4 into 2 equal groups, so each group is 2.",
    "ASK: Which one made the answer bigger? 10 sec, choral. EXPECT: dividing BY a half - 4 ÷ 1/2 = 8.",
  ],
  trap: "reading 4 ÷ 1/2 as 'halve 4' and writing 2. Fix: point to the 8 halves that fit on the line, student restates 4 ÷ 1/2 = 8.",
  stretch: "explain 6 ÷ 1/2 and 6 ÷ 2 in your own words.",
  help: "hand the pre-cut halves for 4 ÷ 1/2, student counts how many fit.",
  prep: "Error-analysis I Do targeting the by-a-half versus halving confusion head on. Number line for fitting, groups for sharing.",
  tag: "[Stage 2 | I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "12 - twelve halves fit in 6, so 6 ÷ 1/2 = 12. Bigger than 6, not smaller.",
  beats: [
    "POINT to the reminder: 1 whole = 2 halves. SAY: Same trap as the I Do. Think how many halves fit.",
    "ASK: 6 ÷ 1/2 - bigger or smaller than 6, and what is it? 30 sec, boards up. EXPECT: bigger, 12. ACCEPT: 6 x 2 = 12.",
    "SCAN boards, back row first. 80%+ -> reveal. Less -> count halves in 1, 2, 3 wholes together, re-ask.",
  ],
  trap: "answering 3 (halved 6). Fix: point to how many halves fit in one whole, count up to 6, student redoes.",
  stretch: "what is 6 ÷ 1/4? Explain how you know.",
  help: "count the halves in one whole first, then in two, then keep going.",
  prep: "Hinge on the by-a-half misconception. Wrong answer 3 exposes halving; correct 12 shows fitting.",
  tag: "[Stage 2 | CFU | Supported application | HITS 7, 8]",
});

const NOTES_CHOOSE = composeGlanceNotes({
  answer: "how many parts fit -> divide BY the fraction (bigger). Share a fraction into groups -> divide the fraction BY the whole (smaller).",
  beats: [
    "POINT to the two branches. SAY: Every fraction-division problem is one of these two questions.",
    "SAY: How many parts fit inside? Divide the amount by the fraction. The answer grows.",
    "SAY: Sharing a fraction into equal groups? Divide the fraction by the whole. The answer is a smaller fraction.",
    "ASK: 3 ÷ 1/4 - which branch? 15 sec, boards up. EXPECT: how many quarters fit - divide by the fraction.",
  ],
  trap: "picking the operation from the numbers, not the meaning. Fix: reread what the problem asks, student re-sorts.",
  stretch: "write one word problem for each branch.",
  help: "match two given word problems to the two branches with a partner.",
  prep: "Explicit strategy for choosing the operation before the guided word problems. Meaning drives the choice, not the numbers.",
  tag: "[Stage 3 | Explicit teaching | Knowledge and memory | HITS 3, 7]",
});

const NOTES_WEDO1 = composeGlanceNotes({
  answer: "6 bows - how many thirds fit in 2 metres, so 2 ÷ 1/3 = 6.",
  beats: [
    "POINT to the problem. SAY: Your turn with support. First, what is it really asking?",
    "ASK: How many parts fit, or share a fraction? 15 sec, boards up. EXPECT: how many 1/3 fit - divide by the fraction.",
    "MODEL 2 ÷ 1/3. SAY: Each metre holds 3 thirds, so 2 metres hold 6. Count them on the strips.",
    "SCAN boards. 80%+ -> next slide. Less -> count thirds in one strip, then the second, re-ask.",
  ],
  trap: "sharing 2 into 3 and writing a smaller answer. Fix: reread - it asks how many bows fit, count the thirds, student redoes.",
  stretch: "how many bows from 5 metres? Explain.",
  help: "strip with thirds pre-drawn, student counts how many fit.",
  prep: "Guided word problem, type A (how many parts fit). Decide first, then solve on strips.",
  tag: "[Stage 3 | We Do | Supported application | HITS 4, 5]",
});

const NOTES_WEDO2 = composeGlanceNotes({
  answer: "1/8 of a metre - share 1/2 m among 4, so 1/2 ÷ 4 = 1/8.",
  beats: [
    "POINT to the problem. SAY: New problem. Decide the question first.",
    "ASK: How many parts fit, or share a fraction? 15 sec, boards up. EXPECT: share 1/2 among 4 - divide the fraction by the whole.",
    "MODEL 1/2 ÷ 4. SAY: Cut the half metre into 4 equal shares. The whole is in eighths, so each child gets 1/8 m.",
    "SCAN boards, back row first. 80%+ -> reveal. Less -> cut the half into 4 on the strip, count all 8, re-ask.",
  ],
  trap: "dividing BY the fraction and getting a bigger answer. Fix: reread - we SHARE the half, so each piece is smaller, student redoes.",
  stretch: "share 1/2 m among 5 - how much each?",
  help: "strip with eighths pre-drawn, student shades one share.",
  prep: "Guided word problem, type B (share a fraction). Contrasts the operation choice with the previous slide.",
  tag: "[Stage 3 | We Do | Supported application | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "Section 1: 4 ÷ 1/2 = 8 pieces. Section 2: 1/3 ÷ 2 = 1/6. Section 3: 6 ÷ 1/2 = 12 (Sam halved instead of counting halves). Challenge: open.",
  beats: [
    "COLLECT the problem set. SAY: On your own. For each one, decide the question first, then solve.",
    "SAY: Write which branch you chose - how many fit, or share a fraction.",
    "CIRCULATE, back row first. Watch for students halving when the problem divides BY a fraction.",
  ],
  trap: "choosing the operation from the numbers, not the meaning. Fix: student rereads and names what the problem asks.",
  stretch: "Section 4 - write your own problem for a partner and solve it.",
  help: "use the decision chart from the board and the strip scaffold.",
  prep: "Independent application: choose and solve both problem types, plus spot the by-a-half error.",
  tag: "[Stage 4 | You Do | Mastery and application | HITS 4, 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "4 pieces - how many 1/2 m fit in 2 m, so 2 ÷ 1/2 = 4. Chosen because it asks how many pieces fit.",
  beats: [
    "SHOW the two prompts. SAY: Last task, on your own. Decide the question, then solve.",
    "TIME 5 minutes. COLLECT boards or slips.",
  ],
  trap: "answer of 1 (halved 2) instead of 4. Note who for tomorrow's small group.",
  prep: "Assesses SC2 (choose and solve the right division). SC number stays in notes, not on the slide.",
  tag: "[Stage 5 | Exit Ticket | Evidence of learning | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - listen for 'how many fit makes it bigger; sharing a fraction makes it smaller; dividing by a half is not halving'.",
  beats: [
    "POINT to the I can statements. SAY: Show me thumbs for each one - up, sideways or down.",
    "ASK: Why is dividing by a half not the same as halving? Turn and tell. 30 sec. EXPECT: by a half asks how many halves fit, so it grows; halving shares into 2.",
    "SAY: You can now choose the operation and explain the by-a-half difference - strong work this week.",
  ],
  prep: "Revisits all three success criteria and the unit's big idea: the meaning of the problem decides the operation.",
  tag: "[Closing | Retention and reflection | HITS 9]",
});

// ─── Lesson targets ──────────────────────────────────────────────────────────

const LI = "We are learning to choose the right operation and solve word problems that divide with fractions.";
const SC = [
  "I can decide if a problem asks how many parts fit or how to share a fraction.",
  "I can choose and solve the right division for a word problem.",
  "I can explain why dividing by a half is not the same as halving.",
];

// ─── Build ───────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, "Choosing the Operation",
    "Which division does the problem ask for, and what is the answer?",
    "Year 6 Numeracy | Fractions", NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, {
    resources: RESOURCE_ITEMS,
    manipulatives: [
      "Paper strips (3-4 per student) to fold into equal parts",
      "Fraction strip sets (optional, if available)",
    ],
    studentTools: [
      "Mini-whiteboards and markers",
      "Printed problem set and decision-chart scaffold",
    ],
    boardSetup: [
      "Decision chart drawn on the board: how many fit vs share a fraction",
      "A long strip and a 0 to 4 number line ready to use live",
    ],
  }, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal (equivalent fractions + simplifying, visual)
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Equal fractions",
      [
        "Fill the gap:  3/4 = ?/8",
        "Write 9/12 in its simplest form.",
      ],
      NOTES_DR, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.34, { strip: C.SECONDARY });
        slide.addText("Equal fractions", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.28,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
        });
        addFractionStripSet(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.48, lg.rightW - 0.5, 0.72,
          [
            { denom: 4, shaded: 3, label: "3/4", color: C.SECONDARY },
            { denom: 8, shaded: 6, label: "?/8", color: C.SECONDARY },
          ], { labelW: 0.7 });

        const simplifyY = lg.panelTopPadded + 1.52;
        addCard(slide, lg.rightX, simplifyY, lg.rightW, 1.14, { strip: C.ACCENT });
        slide.addText("Simplify", {
          x: lg.rightX + 0.2, y: simplifyY + 0.10, w: lg.rightW - 0.4, h: 0.28,
          fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, align: "center", margin: 0,
        });
        addFractionStripSet(slide, lg.rightX + 0.25, simplifyY + 0.50, lg.rightW - 0.5, 0.46,
          [{ denom: 12, shaded: 9, label: "9/12", color: C.ACCENT }], { labelW: 0.7 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, [
        "3/4 = 6/8",
        "9/12 = 3/4",
      ], { y: 4.35, h: 0.72, fontSize: 18 });
    }
  );

  // Slides 5-6: Fluency + reveal (multiplication and division fact families)
  withReveal(
    () => fluencySlide(pres, "Fluency: Fact families",
      ["9 × 7", "63 ÷ 9", "8 × 8"], NOTES_FLUENCY, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, ["9 × 7 = 63", "63 ÷ 9 = 7", "8 × 8 = 64"],
        { y: 4.42, h: 0.66, fontSize: 19 });
    }
  );

  // Slide 7: Launch — retrieve both meanings, same numbers, different answers
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Launch", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Same numbers, different question");

    // Left: Session 1 type (how many fit)
    addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
    s.addText("This week we asked:", {
      x: 0.72, y: CONTENT_TOP + 0.16, w: 3.9, h: 0.34,
      fontSize: 15, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("How many 1/2s fit in 4?", {
      x: 0.72, y: CONTENT_TOP + 0.55, w: 3.9, h: 1.1,
      fontSize: 24, fontFace: FONT_H, color: C.CHARCOAL, bold: true, margin: 0,
    });
    s.addText("counting how many parts fit", {
      x: 0.72, y: CONTENT_TOP + 2.35, w: 3.9, h: 0.5,
      fontSize: 14, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
    });

    // Right: Session 2 type (share a fraction)
    addCard(s, 5.1, CONTENT_TOP, 4.4, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
    s.addText("And we also asked:", {
      x: 5.32, y: CONTENT_TOP + 0.16, w: 4.0, h: 0.34,
      fontSize: 15, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText("Share 1/2 among 4.", {
      x: 5.32, y: CONTENT_TOP + 0.55, w: 4.0, h: 1.1,
      fontSize: 24, fontFace: FONT_H, color: C.CHARCOAL, bold: true, margin: 0,
    });
    s.addText("sharing a fraction into groups", {
      x: 5.32, y: CONTENT_TOP + 2.35, w: 4.0, h: 0.5,
      fontSize: 14, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
    });

    // Bottom estimate banner (full width)
    addTextOnShape(s, "Boards up: which one makes a BIGGER answer?", {
      x: 0.5, y: SAFE_BOTTOM - 0.62, w: 9.0, h: 0.56, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 8: LI / SC
  liSlide(pres, LI, SC, NOTES_LI_SC, FOOTER);

  // Slide 9: I Do — error analysis: divide BY a half vs divide IN half
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Divide BY a half, or IN half?", { color: STAGE_COLORS["2"] });

    const cardTop = CONTENT_TOP;
    const cardH = 2.9;

    // Left: 4 divided by 1/2 -> how many halves fit (bigger)
    addCard(s, 0.5, cardTop, 4.3, cardH, { strip: C.PRIMARY });
    s.addText("4 ÷ 1/2", {
      x: 0.72, y: cardTop + 0.12, w: 3.9, h: 0.5,
      fontSize: 26, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("How many halves fit in 4?", {
      x: 0.72, y: cardTop + 0.68, w: 3.9, h: 0.34,
      fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });
    addNumberLine(s, 0.95, cardTop + 1.55, 3.25,
      ["0", "", "1", "", "2", "", "3", "", "4"], [], { labelFontSize: 13 });
    addTextOnShape(s, "= 8 halves (bigger)", {
      x: 0.9, y: cardTop + 2.18, w: 3.5, h: 0.5, rectRadius: 0.08,
      fill: { color: C.SUCCESS },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    // Right: 4 divided by 2 -> share into 2 groups (smaller)
    addCard(s, 5.1, cardTop, 4.4, cardH, { strip: C.SECONDARY });
    s.addText("4 ÷ 2", {
      x: 5.32, y: cardTop + 0.12, w: 4.0, h: 0.5,
      fontSize: 26, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText("Share 4 into 2 equal groups", {
      x: 5.32, y: cardTop + 0.68, w: 4.0, h: 0.34,
      fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });
    addGroupedCounters(s, 6.05, cardTop + 1.55, 2, 2, { dot: 0.42, groupGap: 0.5 });
    addTextOnShape(s, "= 2 in each group (smaller)", {
      x: 5.5, y: cardTop + 2.18, w: 3.6, h: 0.5, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    // Bottom banner (full width)
    addTextOnShape(s, "Dividing BY 1/2 doubles it. Halving (÷ 2) makes it smaller.", {
      x: 0.5, y: cardTop + cardH + 0.14, w: 9.0, h: 0.55, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slides 10-11: CFU — 6 divided by 1/2 (hinge on the misconception) + reveal
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "6 ÷ 1/2 = ?", { color: C.ALERT });
      addTextOnShape(s, "✓  CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      // Left prompt
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 19, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "6 ÷ 1/2", options: { fontSize: 30, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Bigger or smaller than 6?", options: { fontSize: 18, color: C.CHARCOAL, breakLine: true } },
        { text: "What is the answer?", options: { fontSize: 18, color: C.CHARCOAL } },
      ], {
        x: 0.72, y: CONTENT_TOP + 0.22, w: 3.9, h: SAFE_BOTTOM - CONTENT_TOP - 0.4,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right: reminder anchor (does not give away the count)
      addCard(s, 5.1, CONTENT_TOP, 4.4, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      s.addText("Remember", {
        x: 5.3, y: CONTENT_TOP + 0.14, w: 4.0, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      addFractionStripSet(s, 5.5, CONTENT_TOP + 0.66, 3.6, 0.5,
        [{ denom: 2, shaded: 0, label: "", color: C.PRIMARY }], { showLabels: false });
      s.addText("1 whole = 2 halves.", {
        x: 5.3, y: CONTENT_TOP + 1.35, w: 4.0, h: 0.34,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
      });
      s.addText("How many halves fit in 6?", {
        x: 5.3, y: CONTENT_TOP + 1.78, w: 4.0, h: 0.34,
        fontSize: 15, fontFace: FONT_B, color: C.MUTED, italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU);
      runSlideDiagnostics(s, pres);
      return s;
    },
    (slide) => {
      addRevealAnswerBar(slide, ["6 ÷ 1/2 = 12", "twelve halves fit in 6 - bigger, not smaller"],
        { y: 4.42, h: 0.66, fontSize: 18 });
    }
  );

  // Slide 12: How to choose the operation — decision chart anchor
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "Strategy", { color: C.SECONDARY, w: 1.7 });
    addTitle(s, "How to choose the operation");

    // Left: the strategy steps
    addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
    s.addText([
      { text: "Read the problem.", options: { fontSize: 17, color: C.CHARCOAL, bold: true, bullet: true, breakLine: true, paraSpaceAfter: 10 } },
      { text: "What is it really asking?", options: { fontSize: 17, color: C.CHARCOAL, bullet: true, breakLine: true, paraSpaceAfter: 10 } },
      { text: "How many parts fit? Divide by the fraction.", options: { fontSize: 17, color: C.CHARCOAL, bullet: true, breakLine: true, paraSpaceAfter: 10 } },
      { text: "Share a fraction into groups? Divide the fraction.", options: { fontSize: 17, color: C.CHARCOAL, bullet: true, breakLine: true, paraSpaceAfter: 10 } },
      { text: "Check: is my answer a sensible size?", options: { fontSize: 17, color: C.CHARCOAL, bullet: true } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.2, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.4,
      fontFace: FONT_B, margin: 0, valign: "top",
    });

    // Right: two-branch decision chart
    const rx = 5.3;
    const rw = 4.2;
    const branchH = 1.72;
    // Branch A: how many fit -> divide by the fraction (bigger)
    addCard(s, rx, CONTENT_TOP, rw, branchH, { strip: C.PRIMARY });
    s.addText("How many parts fit inside?", {
      x: rx + 0.2, y: CONTENT_TOP + 0.12, w: rw - 0.4, h: 0.34,
      fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("Divide the amount BY the fraction. The answer gets bigger.", {
      x: rx + 0.2, y: CONTENT_TOP + 0.52, w: rw - 0.4, h: 0.6,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });
    addTextOnShape(s, "3 ÷ 1/4 = 12", {
      x: rx + 0.2, y: CONTENT_TOP + 1.14, w: rw - 0.4, h: 0.44, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    // Branch B: share a fraction -> divide the fraction (smaller)
    const bY = CONTENT_TOP + branchH + 0.14;
    addCard(s, rx, bY, rw, branchH, { strip: C.ACCENT });
    s.addText("Sharing a fraction into groups?", {
      x: rx + 0.2, y: bY + 0.12, w: rw - 0.4, h: 0.34,
      fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
    });
    s.addText("Divide the fraction BY the whole number. The answer is a smaller fraction.", {
      x: rx + 0.2, y: bY + 0.52, w: rw - 0.4, h: 0.6,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });
    addTextOnShape(s, "1/2 ÷ 4 = 1/8", {
      x: rx + 0.2, y: bY + 1.14, w: rw - 0.4, h: 0.44, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CHOOSE);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 13: We Do (type A) — how many bows fit
  workedExSlide(pres, 3, "We Do", "How many bows can we make?",
    [
      "2 m of ribbon.",
      "Each bow uses 1/3 m.",
      "",
      "Asking: how many 1/3 fit?",
      "Divide BY the fraction: 2 ÷ 1/3.",
      "Each metre holds 3 thirds.",
      "2 × 3 = 6 bows.",
    ],
    NOTES_WEDO1, FOOTER,
    (slide, lg) => {
      const cardH = SAFE_BOTTOM - lg.panelTopPadded;
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, cardH, { strip: C.PRIMARY });
      slide.addText("2 metres in thirds", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      addFractionStripSet(slide, lg.rightX + 0.3, lg.panelTopPadded + 0.65, lg.rightW - 0.55, 1.25,
        [
          { denom: 3, shaded: 0, label: "1 m", color: C.SECONDARY },
          { denom: 3, shaded: 0, label: "1 m", color: C.SECONDARY },
        ], { labelW: 0.75 });
      addTextOnShape(slide, "6 thirds fit = 6 bows", {
        x: lg.rightX + 0.5, y: lg.panelTopPadded + 2.35, w: lg.rightW - 1.0, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do (type B) — share half a metre + reveal
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "How much does each child get?",
      [
        "1/2 m of ribbon.",
        "Shared between 4 children.",
        "",
        "Asking: share a fraction?",
        "Divide the fraction: 1/2 ÷ 4.",
        "Cut the half into 4 shares.",
        "Each child gets ______ m.",
      ],
      NOTES_WEDO2, FOOTER,
      (slide, lg) => {
        const cardH = SAFE_BOTTOM - lg.panelTopPadded;
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, cardH, { strip: C.SECONDARY });
        slide.addText("Share 1/2 m among 4", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.28,
          fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
        });
        addFractionStripSet(slide, lg.rightX + 0.3, lg.panelTopPadded + 0.65, lg.rightW - 0.55, 1.25,
          [
            { denom: 2, shaded: 1, label: "1/2 m", color: C.PRIMARY },
            { denom: 8, shaded: 1, label: "1/8 m", color: C.SUCCESS },
          ], { labelW: 0.85 });
        slide.addText("Your answer:  1/2 ÷ 4 = ______", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.35, w: lg.rightW - 0.4, h: 0.4,
          fontSize: 15, fontFace: FONT_B, color: C.MUTED, italic: true, align: "center", valign: "middle", margin: 0,
        });
      }),
    (slide) => {
      addRevealAnswerBar(slide, ["1/2 ÷ 4 = 1/8 m", "each child gets one eighth of a metre"],
        { y: 4.42, h: 0.66, fontSize: 18 });
    }
  );

  // Slide 16: You Do — problem set (choose + solve, both types + spot the error)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: choose, then solve", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.5, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 17, color: C.ALERT, bold: true } },
      { text: "For each problem, decide the question - how many fit, or share a fraction.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
      { text: "Next: ", options: { fontSize: 17, color: C.ALERT, bold: true } },
      { text: "Choose the division and solve. Use a strip or number line if it helps.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
      { text: "Then: ", options: { fontSize: 17, color: C.ALERT, bold: true } },
      { text: "Fix Sam's mistake and explain the by-a-half difference.", options: { fontSize: 16, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.14, w: 8.5, h: 1.24,
      fontFace: FONT_B, valign: "middle", margin: 0, paraSpaceAfter: 5, fit: "shrink", shrinkText: true,
    });

    // Decision reminder panel
    const panelY = CONTENT_TOP + 1.66;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("Decision reminder", {
      x: 0.7, y: panelY + 0.12, w: 8.6, h: 0.30,
      fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
    });
    addTextOnShape(s, "How many parts fit?  ->  divide BY the fraction  (bigger)", {
      x: 0.8, y: panelY + 0.5, w: 8.4, h: 0.44, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 14.5, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(s, "Share a fraction into groups?  ->  divide the fraction  (smaller)", {
      x: 0.8, y: panelY + 1.0, w: 8.4, h: 0.44, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    }, { fontSize: 14.5, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "A 2 m plank is cut into 1/2 m pieces. How many pieces? Show or explain your choice.",
      "How did you decide which division to use?",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 18: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: why is dividing by a half not the same as halving?",
      scItems: SC,
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FracDiv_Lesson3_Choosing_And_Solving.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Companion PDFs ─────────────────────────────────────────────────────────

  // 1) Problem set (choose + solve, both types + spot the error)
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Decide the question first, then choose and solve the right division.",
      color: C.PRIMARY,
      lessonInfo: "Year 6 Numeracy | Fractions",
    });
    y = addTipBox(doc,
      "How many parts fit inside? Divide BY the fraction (answer bigger). Sharing a fraction into groups? Divide the fraction BY the whole (answer a smaller fraction).",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Problem 1 - Cutting rope", y, { color: C.PRIMARY });
    y = addBodyText(doc, "A 4 m rope is cut into 1/2 m pieces. How many pieces are there?", y);
    y = addBodyText(doc, "Tick the question:   [  ] how many parts fit      [  ] share a fraction", y);
    y = addWriteLine(doc, "Number sentence and answer:", y);

    y = addSectionHeading(doc, "Problem 2 - Sharing cake", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1/3 of a cake is shared equally between 2 people. How much does each person get?", y);
    y = addBodyText(doc, "Tick the question:   [  ] how many parts fit      [  ] share a fraction", y);
    y = addWriteLine(doc, "Number sentence and answer:", y);

    y = addSectionHeading(doc, "Problem 3 - Spot the mistake", y, { color: C.ALERT });
    y = addBodyText(doc, "Sam says 6 ÷ 1/2 = 3, because dividing means halving. What did Sam get wrong? Write the correct answer and explain.", y);
    y = addLinedArea(doc, y + 4, 3);

    y = addSectionHeading(doc, "Problem 4 - Challenge", y, { color: C.ACCENT });
    y = addBodyText(doc, "Write your own word problem that needs 8 ÷ 1/4. Then solve it.", y);
    y = addLinedArea(doc, y + 4, 2);

    addPdfFooter(doc, "Session 3 | Problem Set | Year 6 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // 2) Answer key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers and chosen operations for the problem set.",
      color: C.PRIMARY,
      lessonInfo: "Year 6 Numeracy | Fractions",
      showNameDate: false,
    });

    y = addSectionHeading(doc, "Problem 1 - Cutting rope", y, { color: C.PRIMARY });
    y = addBodyText(doc, "How many parts fit. 4 ÷ 1/2 = 8. Each metre holds 2 half-metre pieces, so 4 metres give 8 pieces.", y);

    y = addSectionHeading(doc, "Problem 2 - Sharing cake", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Share a fraction. 1/3 ÷ 2 = 1/6. The third is cut into 2 shares, so the whole is in sixths. Each person gets 1/6.", y);

    y = addSectionHeading(doc, "Problem 3 - Spot the mistake", y, { color: C.ALERT });
    y = addBodyText(doc, "6 ÷ 1/2 = 12, not 3. Sam halved 6 instead of asking how many halves fit in 6. Dividing BY a half asks how many halves fit, so the answer is bigger. Halving would be 6 ÷ 2 = 3.", y);

    y = addSectionHeading(doc, "Problem 4 - Challenge (sample)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Sample: A ribbon is 8 m long. Each bookmark needs 1/4 m. How many bookmarks? 8 ÷ 1/4 = 32 (each metre holds 4 quarters, 8 x 4 = 32).", y);

    y = addTipBox(doc,
      "Watch for: students who halve when a problem divides BY a fraction (Problem 1 and 3). Redirect to the meaning: how many of these parts fit inside?",
      y, { color: C.ACCENT });

    addPdfFooter(doc, "Session 3 | Answer Key | Year 6 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // 3) Enabling scaffold — decision chart + pre-drawn models
  await (async () => {
    const doc = createPdf({ title: SCAFFOLD_RES.name });
    let y = addPdfHeader(doc, SCAFFOLD_RES.name, {
      subtitle: "Tick the question type, then use the model that is drawn for you.",
      color: C.SECONDARY,
      lessonInfo: "Year 6 Numeracy | Fractions",
    });
    y = addTipBox(doc,
      "How many parts fit? Divide BY the fraction and the answer gets bigger. Share a fraction into groups? Divide the fraction and the answer is smaller.",
      y, { color: C.SECONDARY });

    y = addSectionHeading(doc, "Problem 1 - A 4 m rope cut into 1/2 m pieces", y, { color: C.SECONDARY });
    y = addBodyText(doc, "Tick:   [  ] how many parts fit      [  ] share a fraction", y);
    y = addBodyText(doc, "The number line shows 4 metres in halves. Count the half-jumps.", y);
    y = addNumberLinePdf(doc, y + 10, 4, 2);
    y = addWriteLine(doc, "4 ÷ 1/2 =  ______ pieces", y);

    y = addSectionHeading(doc, "Problem 2 - 1/3 of a cake shared between 2", y, { color: C.SECONDARY });
    y = addBodyText(doc, "Tick:   [  ] how many parts fit      [  ] share a fraction", y);
    y = addBodyText(doc, "The whole is cut into sixths. Shade ONE share, then count all the parts.", y);
    y = addFractionStripsPdf(doc, y + 4, 1, 6, { label: "whole (in sixths)" });
    y = addWriteLine(doc, "1/3 ÷ 2 =  ______", y);

    y = addSectionHeading(doc, "Finish the sentence", y, { color: C.SECONDARY });
    y = addWriteLine(doc, "Dividing BY a half asks how many halves  ____________.", y);

    addPdfFooter(doc, "Session 3 | Enabling Scaffold | Year 6 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, SCAFFOLD_RES.fileName));
    console.log("PDF written: " + SCAFFOLD_RES.fileName);
  })();

  console.log("Lesson build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
