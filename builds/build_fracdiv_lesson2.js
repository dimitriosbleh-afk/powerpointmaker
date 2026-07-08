"use strict";

// Year 6 Numeracy — Dividing a unit fraction by a whole number (e.g. 1/2 divided by 3)
// Victorian Curriculum 2.0: Mathematics, Number, Year 6 — solve problems that require
//   dividing a unit fraction by a whole number, using visual models.
// Session 2 of the Dividing With Fractions unit. Meaning first (share the fraction into
//   equal parts), then the denominator x whole pattern. Fraction strips are the anchor.
// Deliberately contrasts with Session 1 (dividing BY a fraction grew the answer; sharing
//   a fraction shrinks it) to seed the by-vs-in confusion addressed fully in Session 3.
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
  addFractionStripsPdf,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  dailyReviewSlide, fluencySlide, addRevealAnswerBar,
  addFractionStripSet,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 2;
const FOOTER = "Fractions | Sharing a Unit Fraction | Year 6 Numeracy";
const OUT_DIR = "output/FracDiv_Lesson2_Divide_Unit_Fraction_By_Whole";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Practice Sheet",
  "Share a unit fraction into equal parts using fraction strips.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the practice sheet.");
const SCAFFOLD_RES = makeSessionResource(SESSION,
  "Enabling Scaffold",
  "Strips already cut into the equal parts - shade one share and count.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, SCAFFOLD_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher notes (Glance Format, megaprompt sections 45-47) ────────────────

const NOTES_TITLE =
  "Open the lesson. Say the title, then move straight to Teacher Resources.";

const NOTES_RESOURCES =
  "Prep slide. Print the practice sheet, answer key and enabling scaffold. " +
  "Have paper strips (3-4 each), whiteboards and markers ready.";

const NOTES_DR = composeGlanceNotes({
  answer: "2/3 = 8/12 (eight twelfths shade the same as two thirds). 6/9 in simplest form is 2/3.",
  beats: [
    "POINT to the two strips. SAY: Equal fractions cover the same length. Two thirds and eight twelfths line up.",
    "ASK: 2/3 equals how many twelfths? 20 sec, boards up. EXPECT: 8, so 2/3 = 8/12.",
    "ASK: Write 6/9 as simply as you can. 30 sec, boards up. EXPECT: 2/3 - divide top and bottom by 3.",
    "SCAN boards. 80%+ -> reveal. Less -> shade 6/9 and 2/3 on strips, match them, re-ask.",
  ],
  trap: "scaling only the top number to make twelfths. Fix: point to the strips, both parts scale by 4, student redoes.",
  prep: "Prior learning: equivalent fractions and simplifying with visual models. Keeps the equal-parts idea sharp for today's sharing.",
  tag: "[Stage 1 | Daily Review | Retention and recall | HITS 6]",
});

const NOTES_FLUENCY = composeGlanceNotes({
  answer: "48, 6, 63.",
  beats: [
    "SAY: Fluency. Say the whole fact family in your head - the times fact pulls the divide fact with it.",
    "ASK: What are the answers? 45 sec, boards up. EXPECT: 48, 6, 63.",
    "SCAN boards, back row first. 80%+ -> reveal. Less -> build 8 x 6 = 48, so 48 divided by 8 = 6, re-ask.",
  ],
  trap: "guessing 48 divided by 8. Fix: use the linked fact 8 x 6 = 48, student states the division.",
  prep: "Brisk fact-family recall (Number). Times-and-divide links feed today's denominator x whole pattern.",
  tag: "[Stage 1 | Fluency | Mastery and application | HITS 6]",
});

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "open - listen for 'each share is smaller than 1/3 because we are sharing only half, not a whole'.",
  beats: [
    "POINT to the whole shared in 3. SAY: Share one whole among 3 friends and each gets one third.",
    "SHOW the half. SAY: Today we share only HALF among the same 3 friends.",
    "ASK: Bigger or smaller than 1/3 each? 20 sec, boards up. EXPECT: smaller. ACCEPT: there is less to share.",
    "SCAN boards. 80%+ -> next slide. Less -> fold a paper half into 3, compare one piece to a third, re-ask.",
  ],
  trap: "saying bigger because 'dividing makes more pieces'. Fix: hold the half beside the whole, less to share means smaller shares, student re-estimates.",
  prep: "Launch bridges sharing a whole into equal parts to sharing a fraction. New concept - assume no prior fraction-by-whole division.",
  tag: "[Launch | Attention and knowledge | HITS 2, 7]",
});

const NOTES_LI_SC = composeGlanceNotes({
  beats: [
    "POINT to the learning intention. SAY: Today we share a unit fraction into equal parts and name each share.",
    "SAY: Read the I can statements with me. Choral read.",
    "SAY: The first one everyone can do today - cut one part into equal shares.",
  ],
  prep: "SC1 achievable by all; SC2 is the core target and the exit ticket; SC3 stretches to explaining why each share shrinks.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_IDO = composeGlanceNotes({
  answer: "1/6 - the half splits into 3 equal shares, and each share is one sixth of the whole.",
  beats: [
    "SHOW the strip with one half shaded. SAY: Let us share this half among 3. Watch how I cut it.",
    "MODEL cutting the half into 3. SAY: The half is now 3 equal pieces. The whole is in sixths, so each piece is 1/6.",
    "SAY: So one half shared among 3 is one sixth. I am sharing the half, not the whole.",
    "ASK: How many sixths make the half? 10 sec, choral. EXPECT: three.",
  ],
  trap: "naming the share 1/3 (shared into 3 so 'thirds'). Fix: count every equal part of the whole strip, 6 in all, student renames one share 1/6.",
  stretch: "share 1/2 among 4 - what is each share? Start it on paper.",
  help: "hand the strip already cut into sixths, student shades one share and names it.",
  prep: "First model - meaning before any rule. Partitive division: share the fraction into equal parts. Strips are the anchor.",
  tag: "[Stage 2 | I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "1/8 - the quarter splits into 2 equal shares, and each share is one eighth of the whole.",
  beats: [
    "POINT to the quarter shaded strip. SAY: Same idea, new fraction. Share one quarter between 2.",
    "ASK: How big is each share? 30 sec, boards up. EXPECT: 1/8. ACCEPT: the quarter cut in 2 makes eighths.",
    "SCAN boards, back row first. 80%+ -> reveal. Less -> cut the quarter into 2 on the strip, count all 8 parts, re-ask.",
  ],
  trap: "answering 1/2 (halved the quarter so 'a half'). Fix: count every equal part of the WHOLE, 8 in all, student renames 1/8.",
  stretch: "share 1/4 among 3 - what is each share?",
  help: "cover all but the quarter, cut it into 2, count the eighths in the whole.",
  prep: "Checks the sharing meaning transfers to a new unit fraction before we generalise.",
  tag: "[Stage 2 | CFU | Supported application | HITS 7, 8]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "1/6 - the third splits into 2 equal shares, and each share is one sixth of the whole.",
  beats: [
    "POINT to the third shaded strip. SAY: Your turn with support. Share one third between 2 people.",
    "MODEL the first cut. SAY: I cut the third into 2. Keep going with your partner and name each share.",
    "ASK: How big is each share? 45 sec, boards up. EXPECT: 1/6. ACCEPT: the third cut in 2 makes sixths.",
    "SCAN boards, back row first. 80%+ -> reveal. Less -> cut the third into 2 on the strip, count all 6 parts, re-ask.",
  ],
  trap: "counting only the 2 shares, not the whole. Fix: count every equal part of the whole strip, student renames 1/6.",
  stretch: "share 1/3 among 4 - what is each share? Use the strip.",
  help: "strip with sixths pre-drawn, student shades one share.",
  prep: "Second guided example, faded from the I Do. Same strip model, students name the share.",
  tag: "[Stage 3 | We Do | Supported application | HITS 4, 5]",
});

const NOTES_CONNECT = composeGlanceNotes({
  answer: "bottom number (denominator) x the whole number = the new bottom number. Each share is 1 over that.",
  beats: [
    "POINT to the three answers. SAY: Look for the pattern. Sharing cuts each part into more, smaller pieces.",
    "ASK: 1/2 divided by 3 - what is 2 x 3? 10 sec, choral. EXPECT: 6, so each share is 1/6.",
    "SAY: Multiply the bottom number by the whole. The strip still shows why each share is smaller.",
    "ASK: Why is this answer smaller than last lesson's? 20 sec, turn and tell. EXPECT: we are sharing, so the pieces get smaller.",
  ],
  trap: "expecting a bigger answer like Session 1. Fix: name the action - here we SHARE a part into pieces, so each is smaller, student explains.",
  stretch: "write the rule in your own words and test it on 1/4 divided by 5.",
  help: "keep the strip picture beside the numbers while you multiply.",
  prep: "Generalises AFTER meaning (section 25). Contrast with Session 1 (dividing BY a fraction grew the answer) seeds the by-vs-in confusion for Session 3.",
  tag: "[Stage 3 | Connect | Knowledge and memory | HITS 4, 9]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "Section 1: 1/2 divided by 4 = 1/8. Section 2: 1/3 divided by 3 = 1/9. Section 3: open - 'sharing makes each piece smaller'. Challenge: 1/5 divided by 3 = 1/15.",
  beats: [
    "COLLECT the worksheet. SAY: On your own now. Shade the fraction, cut it into equal shares, name one share.",
    "SAY: Use the strips first, then check with the bottom-number pattern.",
    "CIRCULATE, back row first. Watch for shares named after the sharing number, not the total parts.",
  ],
  trap: "naming the share after how many shares, e.g. 1/2 divided by 4 = 1/4. Fix: count every equal part of the whole, student renames.",
  stretch: "Section 4 - 1/5 divided by 3, and explain the pattern without drawing.",
  help: "use the enabling scaffold with the parts pre-drawn, student shades one share.",
  prep: "Independent application on NEW numbers. Worksheet keeps the strip model on paper.",
  tag: "[Stage 4 | You Do | Mastery and application | HITS 4, 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "1/12 - the quarter shares into 3, the whole is now twelfths, each share 1/12. Smaller than 1/4 because we cut the quarter into more pieces.",
  beats: [
    "SHOW the two prompts. SAY: Last task, on your own. Draw a strip if it helps.",
    "TIME 4 minutes. COLLECT boards or slips.",
  ],
  trap: "answer of 1/3, or anything larger than 1/4. Note who for tomorrow's small group.",
  prep: "Assesses SC2 (divide a unit fraction by a whole number). SC number stays in notes, not on the slide.",
  tag: "[Stage 5 | Exit Ticket | Evidence of learning | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - listen for 'sharing a fraction into equal parts makes each part smaller'.",
  beats: [
    "POINT to the I can statements. SAY: Show me thumbs for each one - up, sideways or down.",
    "ASK: Why is 1/2 divided by 3 smaller than 1/2? Turn and tell. 30 sec. EXPECT: we shared the half into 3, so each piece is smaller.",
    "SAY: You explained that sharing makes smaller pieces - that is the big idea today.",
  ],
  prep: "Revisits all three success criteria and the key idea that dividing a fraction by a whole shrinks each share.",
  tag: "[Closing | Retention and reflection | HITS 9]",
});

// ─── Lesson targets ──────────────────────────────────────────────────────────

const LI = "We are learning to divide a unit fraction by a whole number by sharing it into equal parts.";
const SC = [
  "I can cut one fraction part into equal shares.",
  "I can divide a unit fraction by a whole number using a strip.",
  "I can explain why each share is smaller than the fraction I started with.",
];

// ─── Build ───────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, "Sharing a Unit Fraction",
    "How big is each share when we split a fraction into equal parts?",
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
      "Printed strip practice sheets",
    ],
    boardSetup: [
      "One long strip drawn on the board, shade one half to cut live",
      "Colour markers to show one share in a different colour",
    ],
  }, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal (equivalent fractions + simplifying, visual)
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Equal fractions",
      [
        "Fill the gap:  2/3 = ?/12",
        "Write 6/9 in its simplest form.",
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
            { denom: 3, shaded: 2, label: "2/3", color: C.SECONDARY },
            { denom: 12, shaded: 8, label: "?/12", color: C.SECONDARY },
          ], { labelW: 0.7 });

        const simplifyY = lg.panelTopPadded + 1.52;
        addCard(slide, lg.rightX, simplifyY, lg.rightW, 1.14, { strip: C.ACCENT });
        slide.addText("Simplify", {
          x: lg.rightX + 0.2, y: simplifyY + 0.10, w: lg.rightW - 0.4, h: 0.28,
          fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, align: "center", margin: 0,
        });
        addFractionStripSet(slide, lg.rightX + 0.25, simplifyY + 0.50, lg.rightW - 0.5, 0.46,
          [{ denom: 9, shaded: 6, label: "6/9", color: C.ACCENT }], { labelW: 0.7 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, [
        "2/3 = 8/12",
        "6/9 = 2/3",
      ], { y: 4.35, h: 0.72, fontSize: 18 });
    }
  );

  // Slides 5-6: Fluency + reveal (multiplication and division fact families)
  withReveal(
    () => fluencySlide(pres, "Fluency: Fact families",
      ["8 × 6", "48 ÷ 8", "7 × 9"], NOTES_FLUENCY, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, ["8 × 6 = 48", "48 ÷ 8 = 6", "7 × 9 = 63"],
        { y: 4.42, h: 0.66, fontSize: 19 });
    }
  );

  // Slide 7: Launch — sharing a whole bridges to sharing a fraction
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Launch", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Sharing a smaller amount");

    // Left: known sharing question (whole among 3)
    addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
    s.addText("We already know:", {
      x: 0.72, y: CONTENT_TOP + 0.16, w: 3.9, h: 0.34,
      fontSize: 16, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText("Share 1 whole among 3.", {
      x: 0.72, y: CONTENT_TOP + 0.52, w: 3.9, h: 0.6,
      fontSize: 21, fontFace: FONT_H, color: C.CHARCOAL, bold: true, margin: 0,
    });
    addFractionStripSet(s, 0.75, CONTENT_TOP + 1.55, 3.7, 0.55,
      [{ denom: 3, shaded: 3, label: "", color: C.SECONDARY }], { showLabels: false });
    s.addText("Each friend gets 1/3.", {
      x: 0.72, y: CONTENT_TOP + 2.35, w: 3.9, h: 0.5,
      fontSize: 15, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
    });

    // Right: today's question (hero) + estimate routine
    addCard(s, 5.1, CONTENT_TOP, 4.4, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
    s.addText("Today:", {
      x: 5.32, y: CONTENT_TOP + 0.16, w: 4.0, h: 0.34,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("Share HALF among 3 friends.", {
      x: 5.32, y: CONTENT_TOP + 0.52, w: 4.0, h: 0.95,
      fontSize: 25, fontFace: FONT_H, color: C.CHARCOAL, bold: true, margin: 0,
    });
    addFractionStripSet(s, 5.35, CONTENT_TOP + 1.7, 3.9, 0.5,
      [{ denom: 2, shaded: 1, label: "1/2", color: C.PRIMARY }], { labelW: 0.7 });
    addTextOnShape(s, "Boards up: bigger or smaller than 1/3?", {
      x: 5.32, y: CONTENT_TOP + 2.5, w: 4.0, h: 0.62, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 8: LI / SC
  liSlide(pres, LI, SC, NOTES_LI_SC, FOOTER);

  // Slide 9: I Do — 1/2 divided by 3 with fraction strips (share the half)
  workedExSlide(pres, 2, "I Do", "Share one half among 3",
    [
      "Take one half.",
      "Cut the half into 3 equal shares.",
      "",
      "The whole is now in sixths.",
      "Each share is 1 of 6 parts.",
      "",
      "So 1/2 ÷ 3 = 1/6.",
    ],
    NOTES_IDO, FOOTER,
    (slide, lg) => {
      const cardH = SAFE_BOTTOM - lg.panelTopPadded;
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, cardH, { strip: C.PRIMARY });
      slide.addText("Share 1/2 among 3", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      addFractionStripSet(slide, lg.rightX + 0.3, lg.panelTopPadded + 0.6, lg.rightW - 0.55, 1.85,
        [
          { denom: 2, shaded: 1, label: "1/2", color: C.SECONDARY },
          { denom: 6, shaded: 3, label: "3/6", color: C.SECONDARY },
          { denom: 6, shaded: 1, label: "1/6", color: C.SUCCESS },
        ], { labelW: 0.8 });
      addTextOnShape(slide, "1/2 ÷ 3 = 1/6", {
        x: lg.rightX + 0.6, y: lg.panelTopPadded + 2.7, w: lg.rightW - 1.2, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 10-11: CFU — share 1/4 between 2 (strips) + reveal
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Share one quarter between 2", { color: C.ALERT });
      addTextOnShape(s, "✓  CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      // Left prompt
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 19, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Share 1/4 between 2.", options: { fontSize: 24, color: C.CHARCOAL, breakLine: true } },
        { text: "How big is each share?", options: { fontSize: 22, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Count all the equal parts of the whole.", options: { fontSize: 15, color: C.MUTED, italic: true } },
      ], {
        x: 0.72, y: CONTENT_TOP + 0.22, w: 3.9, h: SAFE_BOTTOM - CONTENT_TOP - 0.4,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right: one quarter, then cut into 2
      addCard(s, 5.1, CONTENT_TOP, 4.4, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      s.addText("Share 1/4 between 2", {
        x: 5.3, y: CONTENT_TOP + 0.14, w: 4.0, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      addFractionStripSet(s, 5.35, CONTENT_TOP + 0.62, 3.9, 1.15,
        [
          { denom: 4, shaded: 1, label: "1/4", color: C.SECONDARY },
          { denom: 8, shaded: 2, label: "2/8", color: C.SECONDARY },
        ], { labelW: 0.8 });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU);
      runSlideDiagnostics(s, pres);
      return s;
    },
    (slide) => {
      addRevealAnswerBar(slide, ["1/4 ÷ 2 = 1/8", "the quarter cut in 2 makes eighths"],
        { y: 4.42, h: 0.66, fontSize: 18 });
    }
  );

  // Slides 12-13: We Do — share 1/3 between 2 (strips) + reveal
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Share one third between 2",
      [
        "With your partner.",
        "",
        "Take one third.",
        "Cut the third into 2 equal shares.",
        "Count all the parts of the whole.",
        "",
        "Name one share.",
      ],
      NOTES_WEDO, FOOTER,
      (slide, lg) => {
        const cardH = SAFE_BOTTOM - lg.panelTopPadded;
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, cardH, { strip: C.SECONDARY });
        slide.addText("Share 1/3 among 2", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.28,
          fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
        });
        addFractionStripSet(slide, lg.rightX + 0.3, lg.panelTopPadded + 0.6, lg.rightW - 0.55, 1.2,
          [
            { denom: 3, shaded: 1, label: "1/3", color: C.PRIMARY },
            { denom: 6, shaded: 2, label: "2/6", color: C.PRIMARY },
          ], { labelW: 0.8 });
        slide.addText("Your answer:  1/3 ÷ 2 = ______", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.35, w: lg.rightW - 0.4, h: 0.4,
          fontSize: 15, fontFace: FONT_B, color: C.MUTED, italic: true, align: "center", valign: "middle", margin: 0,
        });
      }),
    (slide) => {
      addRevealAnswerBar(slide, ["1/3 ÷ 2 = 1/6", "the third cut in 2 makes sixths"],
        { y: 4.42, h: 0.66, fontSize: 18 });
    }
  );

  // Slide 14: Connect to the rule — denominator x whole, and why smaller
  workedExSlide(pres, 3, "Connect", "A quicker way to check",
    [
      "Look back at our answers.",
      "",
      "1/2 ÷ 3 = 1/6.",
      "1/3 ÷ 2 = 1/6.",
      "1/4 ÷ 2 = 1/8.",
      "",
      "Bottom number × whole = new bottom.",
      "Sharing makes each part smaller.",
    ],
    NOTES_CONNECT, FOOTER,
    (slide, lg) => {
      const cardH = SAFE_BOTTOM - lg.panelTopPadded;
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, cardH, { strip: C.ACCENT });
      slide.addText("Same answer, faster", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.ACCENT, bold: true, align: "center", margin: 0,
      });
      const rows = [
        "1/2 ÷ 3  =  1 / (2 × 3)  =  1/6",
        "1/3 ÷ 2  =  1 / (3 × 2)  =  1/6",
        "1/4 ÷ 5  =  1 / (4 × 5)  =  1/20",
      ];
      rows.forEach((r, i) => {
        addTextOnShape(slide, r, {
          x: lg.rightX + 0.25, y: lg.panelTopPadded + 0.65 + i * 0.72, w: lg.rightW - 0.5, h: 0.55, rectRadius: 0.08,
          fill: { color: i === 2 ? C.PRIMARY : C.SECONDARY },
        }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
      addTextOnShape(slide, "Session 1 grew. Sharing shrinks.", {
        x: lg.rightX + 0.25, y: lg.panelTopPadded + 2.85, w: lg.rightW - 0.5, h: 0.5, rectRadius: 0.08,
        fill: { color: C.ALERT },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 15: You Do — practice sheet (new numbers)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: how big is each share?", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.35, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 17, color: C.ALERT, bold: true } },
      { text: "Share 1/2 among 4. Shade, cut, name one share.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
      { text: "Next: ", options: { fontSize: 17, color: C.ALERT, bold: true } },
      { text: "Share 1/3 among 3. Shade, cut, name one share.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
      { text: "Then: ", options: { fontSize: 17, color: C.ALERT, bold: true } },
      { text: "Explain why each share is smaller than the fraction you started with.", options: { fontSize: 16, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.5, h: 1.12,
      fontFace: FONT_B, valign: "middle", margin: 0, paraSpaceAfter: 4, fit: "shrink", shrinkText: true,
    });

    // Tool reminder panel with the strip model
    const panelY = CONTENT_TOP + 1.52;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("Remember the strip model", {
      x: 0.7, y: panelY + 0.14, w: 8.6, h: 0.30,
      fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
    });
    s.addText("Shade the fraction, then cut that part into equal shares.", {
      x: 0.8, y: panelY + 0.52, w: 8.4, h: 0.3,
      fontSize: 13, fontFace: FONT_B, color: C.MUTED, margin: 0,
    });
    addFractionStripSet(s, 0.8, panelY + 0.92, 8.4, 0.95,
      [
        { denom: 3, shaded: 1, label: "1/3", color: C.SECONDARY },
        { denom: 9, shaded: 1, label: "1/9 each", color: C.SUCCESS },
      ], { labelW: 1.1 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 16: Exit Ticket
  exitTicketSlide(pres,
    [
      "Share 1/4 among 3. How big is each share? You may draw a strip.",
      "Explain why each share is smaller than 1/4.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 17: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: why is 1/2 ÷ 3 smaller than 1/2?",
      scItems: SC,
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FracDiv_Lesson2_Divide_Unit_Fraction_By_Whole.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Companion PDFs ─────────────────────────────────────────────────────────

  // 1) Practice sheet (core You Do task, strip model on paper)
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Share a unit fraction into equal parts. How big is each share?",
      color: C.PRIMARY,
      lessonInfo: "Year 6 Numeracy | Fractions",
    });
    y = addTipBox(doc,
      "To share a fraction, cut that part into equal pieces. Then count every equal part of the WHOLE strip to name one share.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 - Share 1/2 among 4", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Shade one half. Cut the half into 4 equal shares. Count all the parts, then name one share.", y);
    y = addFractionStripsPdf(doc, y + 4, 1, 2, { label: "whole (in halves)" });
    y = addWriteLine(doc, "1/2 ÷ 4 =  ______", y);

    y = addSectionHeading(doc, "Section 2 - Share 1/3 among 3", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Shade one third. Cut the third into 3 equal shares. Count all the parts, then name one share.", y);
    y = addFractionStripsPdf(doc, y + 4, 1, 3, { label: "whole (in thirds)" });
    y = addWriteLine(doc, "1/3 ÷ 3 =  ______", y);

    y = addSectionHeading(doc, "Section 3 - Explain your thinking", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Why is each share smaller than the fraction you started with?", y);
    y = addLinedArea(doc, y + 4, 2);

    y = addSectionHeading(doc, "Section 4 - Challenge", y, { color: C.ACCENT });
    y = addWriteLine(doc, "Share 1/5 among 3.    1/5 ÷ 3 =  ______", y + 4);
    y = addBodyText(doc, "Explain how the bottom-number pattern gives the answer without drawing every part.", y);
    y = addLinedArea(doc, y + 2, 2);

    addPdfFooter(doc, "Session 2 | Practice Sheet | Year 6 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // 2) Answer key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the practice sheet.",
      color: C.PRIMARY,
      lessonInfo: "Year 6 Numeracy | Fractions",
      showNameDate: false,
    });

    y = addSectionHeading(doc, "Section 1 - 1/2 ÷ 4", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1/2 ÷ 4 = 1/8.  The half is cut into 4 shares, so the whole is in eighths (2 × 4 = 8). Each share is 1/8.", y);

    y = addSectionHeading(doc, "Section 2 - 1/3 ÷ 3", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1/3 ÷ 3 = 1/9.  The third is cut into 3 shares, so the whole is in ninths (3 × 3 = 9). Each share is 1/9.", y);

    y = addSectionHeading(doc, "Section 3 - Explain (sample answer)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "We take one part and cut it into more equal pieces. Cutting a part into more pieces makes each piece smaller, so each share is smaller than the fraction we started with.", y);

    y = addSectionHeading(doc, "Section 4 - Challenge", y, { color: C.ACCENT });
    y = addBodyText(doc, "1/5 ÷ 3 = 1/15.  Bottom number × whole = 5 × 3 = 15, so each share is 1/15.", y);

    y = addTipBox(doc,
      "Watch for: students who name the share after how many shares (e.g. 1/2 ÷ 4 = 1/4). Redirect to the strip: count every equal part of the whole, not just the shares.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, "Session 2 | Answer Key | Year 6 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // 3) Enabling scaffold — parts pre-cut, student shades one share and counts
  await (async () => {
    const doc = createPdf({ title: SCAFFOLD_RES.name });
    let y = addPdfHeader(doc, SCAFFOLD_RES.name, {
      subtitle: "The equal parts are drawn for you. Shade ONE share, then count all the parts.",
      color: C.SECONDARY,
      lessonInfo: "Year 6 Numeracy | Fractions",
    });
    y = addTipBox(doc,
      "Shade one share, then count every equal part of the whole strip. That number is the bottom of your answer.",
      y, { color: C.SECONDARY });

    y = addSectionHeading(doc, "Section 1 - Share 1/2 among 4", y, { color: C.SECONDARY });
    y = addBodyText(doc, "The whole is cut into eighths. Shade ONE share and count all the parts.", y);
    y = addFractionStripsPdf(doc, y + 4, 1, 8, { label: "whole (in eighths)" });
    y = addWriteLine(doc, "1/2 ÷ 4 =  ______", y);

    y = addSectionHeading(doc, "Section 2 - Share 1/3 among 3", y, { color: C.SECONDARY });
    y = addBodyText(doc, "The whole is cut into ninths. Shade ONE share and count all the parts.", y);
    y = addFractionStripsPdf(doc, y + 4, 1, 9, { label: "whole (in ninths)" });
    y = addWriteLine(doc, "1/3 ÷ 3 =  ______", y);

    y = addSectionHeading(doc, "Section 3 - Finish the sentences", y, { color: C.SECONDARY });
    y = addWriteLine(doc, "When I share a fraction, each piece gets  ____________.", y);
    y = addWriteLine(doc, "1/2 shared among 4 is  ______.", y);

    addPdfFooter(doc, "Session 2 | Enabling Scaffold | Year 6 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, SCAFFOLD_RES.fileName));
    console.log("PDF written: " + SCAFFOLD_RES.fileName);
  })();

  console.log("Lesson build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
