"use strict";

// Year 6 Numeracy — Dividing a whole number by a unit fraction (e.g. 3 divided by 1/4)
// Victorian Curriculum 2.0: Mathematics, Number, Year 6 — solve problems that require
//   dividing a whole number by a unit fraction, using visual models.
// First lesson on this concept. Meaning first (how many unit fractions fit), then the
//   whole x denominator pattern. Fraction strips + number lines are the anchors.
// Daily Review: unit fractions and placing fractions on a number line.
// Fluency: multiplication facts.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const { composeGlanceNotes } = require("../themes/core/composeNotes");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource, hex,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  dailyReviewSlide, fluencySlide, addRevealAnswerBar,
  addNumberLine, addFractionStripSet,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const FOOTER = "Fractions | Dividing by a Unit Fraction | Year 6 Numeracy";
const OUT_DIR = "output/FracDiv_Lesson1_Divide_Whole_By_Unit_Fraction";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Practice Sheet",
  "Use fraction strips and number lines to find how many unit fractions fit.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the practice sheet.");
const SCAFFOLD_RES = makeSessionResource(SESSION,
  "Enabling Scaffold",
  "Strips and number lines with the parts already drawn - count and write.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, SCAFFOLD_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher notes (Glance Format, megaprompt sections 45-47) ────────────────

const NOTES_TITLE =
  "Open the lesson. Say the title, then move straight to Teacher Resources.";

const NOTES_RESOURCES =
  "Prep slide. Print the practice sheet, answer key and enabling scaffold. " +
  "Have paper strips (3-4 each), whiteboards and markers ready.";

const NOTES_DR = composeGlanceNotes({
  answer: "1/2 = 3/6 (three sixths shade the same as one half). 4/8 in simplest form is 1/2.",
  beats: [
    "POINT to the two strips. SAY: Equal fractions cover the same amount. One half and three sixths line up.",
    "ASK: 1/2 equals how many sixths? 20 sec, boards up. EXPECT: 3, so 1/2 = 3/6.",
    "ASK: Write 4/8 as simply as you can. 30 sec, boards up. EXPECT: 1/2 - divide top and bottom by 4.",
    "SCAN boards. 80%+ -> reveal. Less -> shade 4/8 and 1/2 on strips, match them, re-ask.",
  ],
  trap: "changing only the bottom number, e.g. 1/2 = 1/6. Fix: point to the strips, both parts grow the same way, student redoes.",
  prep: "Prior learning: equivalent fractions and simplifying with visual models. Warms the equal-parts idea today's division rests on.",
  tag: "[Stage 1 | Daily Review | Retention and recall | HITS 6]",
});

const NOTES_FLUENCY = composeGlanceNotes({
  answer: "42, 6, 72.",
  beats: [
    "SAY: Fluency. Multiplication and division live in fact families. Say the whole family in your head, then write.",
    "ASK: What are the answers? 45 sec, boards up. EXPECT: 42, 6, 72.",
    "SCAN boards, back row first. 80%+ -> reveal. Less -> build 6 x 7 = 42, so 42 divided by 7 = 6, re-ask.",
  ],
  trap: "stuck on 42 divided by 7. Fix: use the linked fact 6 x 7 = 42, student states the division.",
  prep: "Brisk fact-family recall (Number). Division facts feed today's 'how many groups fit' thinking.",
  tag: "[Stage 1 | Fluency | Mastery and application | HITS 6]",
});

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "open - listen for a link between 'how many 2s fit in 6' and 'how many quarters fit in 3'; estimates above 3.",
  beats: [
    "POINT to the six counters in groups of two. SAY: We can ask how many 2s fit into 6. Here that is three groups.",
    "ASK: How many 2s fit into 6? 10 sec, choral. EXPECT: three.",
    "SHOW the new question. SAY: Today we ask the same kind of question - how many quarters fit into 3 wholes?",
    "ASK: Estimate - more than 3 or fewer? 15 sec, boards up. EXPECT: more than 3. ACCEPT: a quarter is small so many fit.",
    "SCAN boards. 80%+ -> next slide. Less -> fold one strip into quarters, count 4 in one whole, re-ask.",
  ],
  trap: "estimating fewer than 3 because dividing 'makes it smaller'. Fix: fold one whole into 4, count 4 in ONE whole, student re-estimates.",
  prep: "Launch bridges whole-number grouping division to fitting unit fractions. New concept - assume no prior fraction division.",
  tag: "[Launch | Attention and knowledge | HITS 2, 7]",
});

const NOTES_LI_SC = composeGlanceNotes({
  beats: [
    "POINT to the learning intention. SAY: Today we find how many equal parts fit inside a whole number.",
    "SAY: Read the I can statements with me. Choral read.",
    "SAY: The first one everyone can do today - count how many equal parts are in one whole.",
  ],
  prep: "SC1 achievable by all; SC2 is the core target and the exit ticket; SC3 stretches to explaining why the answer grows.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_IDO = composeGlanceNotes({
  answer: "12 - twelve quarters fit into 3 wholes, so 3 divided by 1/4 = 12.",
  beats: [
    "SHOW three whole strips. SAY: Let us work this out together. Watch how I split each whole into quarters.",
    "MODEL splitting and counting. SAY: One whole holds four quarters. I count 4, 8, 12 - twelve quarters in all.",
    "SAY: So three divided by one quarter is twelve. I am asking how many quarters fit, not sharing into groups.",
    "ASK: How many quarters in ONE whole? 10 sec, choral. EXPECT: four.",
  ],
  trap: "thinking the answer must be smaller than 3. Fix: point to the 12 quarters on the strips, student says 'more parts fit'.",
  stretch: "how many quarters fit into 5? Start it on paper.",
  help: "hand the pre-split strip scaffold, student counts the quarters aloud.",
  prep: "First model - meaning before any rule. Quotitive division: how many 1/4 fit. Strips are the anchor.",
  tag: "[Stage 2 | I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "4 - four halves fit into 2 wholes, so 2 divided by 1/2 = 4.",
  beats: [
    "POINT to the two strips split into halves. SAY: Same question, new fraction. How many halves fit into two wholes?",
    "ASK: Show me the count on your board. 30 sec, boards up. EXPECT: 4. ACCEPT: 2 halves in each whole, doubled.",
    "SCAN boards, back row first. 80%+ -> reveal. Less -> count halves on ONE strip, then both, re-ask.",
  ],
  trap: "answering 2 (halves in one whole only). Fix: run a finger across BOTH strips, count all four, student recounts.",
  stretch: "how many halves fit into 3?",
  help: "cover one strip, count halves in one whole first, then add the second.",
  prep: "Checks the 'how many fit' meaning transfers to a new unit fraction before we move to the number line.",
  tag: "[Stage 2 | CFU | Supported application | HITS 7, 8]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "6 - six thirds fit into 2 wholes, so 2 divided by 1/3 = 6.",
  beats: [
    "POINT to 0 on the line. SAY: Your turn with support. We jump one third at a time and count the jumps.",
    "MODEL the first jump. SAY: Zero to one third is jump one. Keep going with your partner.",
    "ASK: How many thirds fit into 2? 45 sec, boards up. EXPECT: 6. ACCEPT: three in each whole, two wholes.",
    "SCAN boards, back row first. 80%+ -> reveal. Less -> count thirds in ONE whole on the line, then the second, re-ask.",
  ],
  trap: "counting the tick marks (7) instead of the jumps. Fix: count the spaces between marks, student recounts jumps.",
  stretch: "how many thirds fit into 4? Use the line.",
  help: "number-line scaffold with thirds pre-ticked, student counts the jumps.",
  prep: "Second representation - the number line. Guided count of unit-fraction jumps, faded from the I Do strips.",
  tag: "[Stage 3 | We Do | Supported application | HITS 4, 5]",
});

const NOTES_CONNECT = composeGlanceNotes({
  answer: "whole number x the bottom number (denominator) = how many unit fractions fit.",
  beats: [
    "POINT to the three answers. SAY: Look for a pattern. Each whole splits into 'bottom number' parts.",
    "ASK: 3 divided by 1/4 - what is 3 x 4? 10 sec, choral. EXPECT: 12, the same as we counted.",
    "SAY: So we can multiply the whole by the bottom number. The picture still shows why it works.",
    "ASK: Try 4 divided by 1/5. 20 sec, boards up. EXPECT: 20 (4 x 5).",
  ],
  trap: "multiplying by the top number instead of the bottom. Fix: point to the denominator = size of each part, student redoes 4 x 5.",
  stretch: "write the rule in your own words.",
  help: "keep the strip picture beside the numbers while you multiply.",
  prep: "Generalises AFTER meaning (section 25). Rule sits on top of the strip and line, never replaces them.",
  tag: "[Stage 3 | Connect | Knowledge and memory | HITS 4, 9]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "Section 1: 3 divided by 1/2 = 6. Section 2: 2 divided by 1/4 = 8. Section 3: open - 'parts are smaller so more fit'.",
  beats: [
    "COLLECT the worksheet. SAY: On your own now. Use the strips first, then the number line, just like we practised.",
    "SAY: Split each whole into the unit fraction, then count how many fit.",
    "CIRCULATE, back row first. Watch for students counting parts, not sharing into groups.",
  ],
  trap: "writing an answer smaller than the whole. Fix: point to the drawn parts, student counts every part.",
  stretch: "Section 4 - how many eighths fit into 2, and explain why.",
  help: "use the enabling scaffold with the parts pre-drawn, student counts.",
  prep: "Independent application on NEW numbers (3 divided by 1/2, 2 divided by 1/4). Worksheet keeps both models on paper.",
  tag: "[Stage 4 | You Do | Mastery and application | HITS 4, 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "15 - five wholes, three thirds in each, 5 x 3 = 15. More than 5 because each whole makes three smaller parts.",
  beats: [
    "SHOW the two prompts. SAY: Last task, on your own. Draw a strip or number line if it helps.",
    "TIME 4 minutes. COLLECT boards or slips.",
  ],
  trap: "answer below 5 (treated it like sharing). Note who for tomorrow's small group.",
  prep: "Assesses SC2 (divide a whole number by a unit fraction). SC number stays in notes, not on the slide.",
  tag: "[Stage 5 | Exit Ticket | Evidence of learning | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - listen for 'a quarter is smaller than a whole, so many of them fit into 3'.",
  beats: [
    "POINT to the I can statements. SAY: Show me thumbs for each one - up, sideways or down.",
    "ASK: Why is 3 divided by 1/4 more than 3? Turn and tell. 30 sec. EXPECT: quarters are small, so lots fit.",
    "SAY: You explained that smaller parts means more fit - that is the big idea today.",
  ],
  prep: "Revisits all three success criteria and the key idea that dividing by a unit fraction grows the count.",
  tag: "[Closing | Retention and reflection | HITS 9]",
});

// ─── Lesson targets ──────────────────────────────────────────────────────────

const LI = "We are learning to divide a whole number by a unit fraction by finding how many equal parts fit inside.";
const SC = [
  "I can count how many equal parts are in one whole.",
  "I can divide a whole number by a unit fraction using a strip or a number line.",
  "I can explain why dividing by a unit fraction gives a bigger number.",
];

// ─── Small drawing helpers (slide-side) ──────────────────────────────────────

// Grouped counters for the launch: `groups` rows of `per` dots, each group framed.
function drawGroupedCounters(slide, x, y, groups, per, opts) {
  const o = opts || {};
  const dot = o.dot || 0.26;
  const gap = o.gap || 0.12;
  const framePad = 0.10;
  const groupW = per * dot + (per - 1) * gap + framePad * 2;
  const groupGap = o.groupGap || 0.24;
  for (let g = 0; g < groups; g += 1) {
    const gx = x + g * (groupW + groupGap);
    slide.addShape("roundRect", {
      x: gx, y: y - framePad, w: groupW, h: dot + framePad * 2, rectRadius: 0.06,
      fill: { color: C.BG_LIGHT }, line: { color: C.SECONDARY, width: 1.2 },
    });
    for (let i = 0; i < per; i += 1) {
      const cx = gx + framePad + i * (dot + gap);
      slide.addShape("roundRect", {
        x: cx, y, w: dot, h: dot, rectRadius: dot / 2,
        fill: { color: C.ACCENT }, line: { color: C.ACCENT, width: 0.2 },
      });
    }
  }
  return { totalW: groups * groupW + (groups - 1) * groupGap };
}

// Number line 0..end with a tick at every unit fraction (denom parts per whole).
// Labels whole numbers only. Optionally prints jump counts above each interval.
function drawUnitFractionLine(slide, x, y, w, end, denom, opts) {
  const o = opts || {};
  const intervals = end * denom;
  const labels = [];
  for (let i = 0; i <= intervals; i += 1) {
    labels.push(i % denom === 0 ? String(i / denom) : "");
  }
  const geo = addNumberLine(slide, x, y, w, labels, [], { labelFontSize: o.labelFontSize || 15 });
  if (o.showJumps) {
    for (let i = 0; i < intervals; i += 1) {
      const cx = geo.x + (i + 0.5) * geo.intervalW;
      slide.addText(String(i + 1), {
        x: cx - 0.18, y: geo.y - 0.40, w: 0.36, h: 0.26,
        fontSize: 12, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  }
  return geo;
}

// ─── Small drawing helpers (PDF-side, pdfkit primitives) ─────────────────────

// Draw `count` whole strips. parts=0 -> blank whole (student partitions);
// parts>1 -> pre-divided into equal cells (enabling scaffold). Returns next y.
function drawPdfStrips(doc, y, count, parts, opts) {
  const o = opts || {};
  const x = 55;
  const w = o.width || 380;
  const h = o.stripH || 26;
  const gap = o.gap || 12;
  for (let s = 0; s < count; s += 1) {
    const sy = y + s * (h + gap);
    doc.save();
    doc.lineWidth(1.2).strokeColor("#000000");
    doc.rect(x, sy, w, h).stroke();
    if (parts > 1) {
      const cw = w / parts;
      for (let i = 1; i < parts; i += 1) {
        doc.moveTo(x + i * cw, sy).lineTo(x + i * cw, sy + h).stroke();
      }
    }
    doc.restore();
    doc.fontSize(10).font("Sans").fillColor("#6B7280");
    doc.text("1 whole", x + w + 10, sy + h / 2 - 6, { lineBreak: false });
  }
  return y + count * (h + gap) + 6;
}

// Draw a 0..end number line. tickDenom = ticks per whole (1 = whole ticks only,
// student marks the parts; 3/4 = parts pre-ticked for the scaffold). Returns y.
function drawPdfNumberLine(doc, y, end, tickDenom, opts) {
  const o = opts || {};
  const x = 65;
  const w = o.width || 430;
  const totalTicks = end * tickDenom;
  const stepW = w / totalTicks;
  doc.save();
  doc.lineWidth(1.4).strokeColor("#000000");
  doc.moveTo(x, y).lineTo(x + w, y).stroke();
  for (let i = 0; i <= totalTicks; i += 1) {
    const tx = x + i * stepW;
    const whole = i % tickDenom === 0;
    const th = whole ? 9 : 6;
    doc.moveTo(tx, y - th).lineTo(tx, y + th).stroke();
    if (whole) {
      doc.fontSize(11).font("Sans-Bold").fillColor("#000000");
      doc.text(String(i / tickDenom), tx - 8, y + 13, { width: 16, align: "center", lineBreak: false });
    }
  }
  doc.restore();
  return y + 36;
}

// ─── Build ───────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, "Dividing by a Unit Fraction",
    "How many equal parts fit inside a whole number?",
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
      "Printed strip and number-line practice sheets",
    ],
    boardSetup: [
      "One long strip drawn on the board to fold into quarters live",
      "A 0 to 3 number line ready to build with the class",
    ],
  }, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal (equivalent fractions + simplifying, visual)
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Equal fractions",
      [
        "Fill the gap:  1/2 = ?/6",
        "Write 4/8 in its simplest form.",
      ],
      NOTES_DR, FOOTER,
      (slide, lg) => {
        // Equivalence: 1/2 lines up with 3/6 on matching strips
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.34, { strip: C.SECONDARY });
        slide.addText("Equal fractions", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.28,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
        });
        addFractionStripSet(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.48, lg.rightW - 0.5, 0.72,
          [
            { denom: 2, shaded: 1, label: "1/2", color: C.SECONDARY },
            { denom: 6, shaded: 3, label: "?/6", color: C.SECONDARY },
          ], { labelW: 0.7 });

        // Simplify: 4 of 8 parts shaded
        const simplifyY = lg.panelTopPadded + 1.52;
        addCard(slide, lg.rightX, simplifyY, lg.rightW, 1.14, { strip: C.ACCENT });
        slide.addText("Simplify", {
          x: lg.rightX + 0.2, y: simplifyY + 0.10, w: lg.rightW - 0.4, h: 0.28,
          fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, align: "center", margin: 0,
        });
        addFractionStripSet(slide, lg.rightX + 0.25, simplifyY + 0.50, lg.rightW - 0.5, 0.46,
          [{ denom: 8, shaded: 4, label: "4/8", color: C.ACCENT }], { labelW: 0.7 });
      }),
    (slide) => {
      addRevealAnswerBar(slide, [
        "1/2 = 3/6",
        "4/8 = 1/2",
      ], { y: 4.35, h: 0.72, fontSize: 18 });
    }
  );

  // Slides 5-6: Fluency + reveal (multiplication and division fact families)
  withReveal(
    () => fluencySlide(pres, "Fluency: Fact families",
      ["6 × 7", "42 ÷ 7", "9 × 8"], NOTES_FLUENCY, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, ["6 × 7 = 42", "42 ÷ 7 = 6", "9 × 8 = 72"],
        { y: 4.42, h: 0.66, fontSize: 19 });
    }
  );

  // Slide 7: Launch — grouping division bridges to fitting unit fractions
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Launch", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "How many fit inside?");

    // Left: known grouping question
    addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
    s.addText("We already know:", {
      x: 0.72, y: CONTENT_TOP + 0.16, w: 3.9, h: 0.34,
      fontSize: 16, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText("How many groups of 2 are in 6?", {
      x: 0.72, y: CONTENT_TOP + 0.52, w: 3.9, h: 0.7,
      fontSize: 21, fontFace: FONT_H, color: C.CHARCOAL, bold: true, margin: 0,
    });
    drawGroupedCounters(s, 0.9, CONTENT_TOP + 1.75, 3, 2);
    s.addText("6 shared into groups of 2 makes 3 groups.", {
      x: 0.72, y: CONTENT_TOP + 2.35, w: 3.9, h: 0.6,
      fontSize: 14, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
    });

    // Right: today's question (hero) + estimate routine
    addCard(s, 5.1, CONTENT_TOP, 4.4, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
    s.addText("Today:", {
      x: 5.32, y: CONTENT_TOP + 0.16, w: 4.0, h: 0.34,
      fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("How many quarters fit into 3 wholes?", {
      x: 5.32, y: CONTENT_TOP + 0.55, w: 4.0, h: 1.2,
      fontSize: 26, fontFace: FONT_H, color: C.CHARCOAL, bold: true, margin: 0,
    });
    addTextOnShape(s, "Boards up: more or fewer than 3?", {
      x: 5.32, y: CONTENT_TOP + 2.35, w: 4.0, h: 0.62, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 8: LI / SC
  liSlide(pres, LI, SC, NOTES_LI_SC, FOOTER);

  // Slide 9: I Do — 3 divided by 1/4 with fraction strips
  workedExSlide(pres, 2, "I Do", "How many quarters fit into 3?",
    [
      "Take 3 whole strips.",
      "Split each whole into quarters.",
      "",
      "Count every quarter.",
      "4 + 4 + 4 = 12 quarters.",
      "",
      "So 3 ÷ 1/4 = 12.",
    ],
    NOTES_IDO, FOOTER,
    (slide, lg) => {
      const cardH = SAFE_BOTTOM - lg.panelTopPadded;
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, cardH, { strip: C.PRIMARY });
      slide.addText("3 wholes, each split into quarters", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      addFractionStripSet(slide, lg.rightX + 0.25, lg.panelTopPadded + 0.62, lg.rightW - 0.5, 1.7,
        [
          { denom: 4, shaded: 0, label: "1 whole", color: C.SECONDARY },
          { denom: 4, shaded: 0, label: "1 whole", color: C.SECONDARY },
          { denom: 4, shaded: 0, label: "1 whole", color: C.SECONDARY },
        ], { labelW: 0.9 });
      addTextOnShape(slide, "12 quarters fit into 3", {
        x: lg.rightX + 0.6, y: lg.panelTopPadded + 2.55, w: lg.rightW - 1.2, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 10-11: CFU — how many halves fit into 2 (strips) + reveal
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "How many halves fit into 2?", { color: C.ALERT });
      addTextOnShape(s, "✓  CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      // Left prompt
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 19, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "How many halves (1/2)", options: { fontSize: 23, color: C.CHARCOAL, breakLine: true } },
        { text: "fit into 2 wholes?", options: { fontSize: 23, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Count the parts on the strips.", options: { fontSize: 15, color: C.MUTED, italic: true } },
      ], {
        x: 0.72, y: CONTENT_TOP + 0.22, w: 3.9, h: SAFE_BOTTOM - CONTENT_TOP - 0.4,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right: 2 whole strips in halves
      addCard(s, 5.1, CONTENT_TOP, 4.4, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      s.addText("2 wholes, each split into halves", {
        x: 5.3, y: CONTENT_TOP + 0.14, w: 4.0, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      addFractionStripSet(s, 5.35, CONTENT_TOP + 0.7, 3.9, 1.5,
        [
          { denom: 2, shaded: 0, label: "1 whole", color: C.SECONDARY },
          { denom: 2, shaded: 0, label: "1 whole", color: C.SECONDARY },
        ], { labelW: 0.9 });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU);
      runSlideDiagnostics(s, pres);
      return s;
    },
    (slide) => {
      addRevealAnswerBar(slide, ["2 ÷ 1/2 = 4", "two halves in each whole: 2 + 2 = 4"],
        { y: 4.42, h: 0.66, fontSize: 18 });
    }
  );

  // Slides 12-13: We Do — 2 divided by 1/3 on a number line + reveal
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "How many thirds fit into 2?",
      [
        "With your partner.",
        "",
        "Start at 0 on the line.",
        "Jump one third each time.",
        "Count the jumps to 2.",
        "",
        "Write how many thirds fit.",
      ],
      NOTES_WEDO, FOOTER,
      (slide, lg) => {
        const cardH = SAFE_BOTTOM - lg.panelTopPadded;
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, cardH, { strip: C.SECONDARY });
        slide.addText("Number line 0 to 2 in thirds", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.28,
          fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
        });
        drawUnitFractionLine(slide, lg.rightX + 0.35, lg.panelTopPadded + 1.25, lg.rightW - 0.70, 2, 3,
          { labelFontSize: 15 });
        slide.addText("Each jump = 1/3", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.60, w: lg.rightW - 0.4, h: 0.26,
          fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true, align: "center", margin: 0,
        });
        slide.addText("Your answer:  ______ thirds", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.35, w: lg.rightW - 0.4, h: 0.4,
          fontSize: 15, fontFace: FONT_B, color: C.MUTED, italic: true, align: "center", valign: "middle", margin: 0,
        });
      }),
    (slide) => {
      addRevealAnswerBar(slide, ["2 ÷ 1/3 = 6", "six thirds fit into 2 wholes"],
        { y: 4.42, h: 0.66, fontSize: 18 });
    }
  );

  // Slide 14: Connect to the rule - whole x denominator
  workedExSlide(pres, 3, "Connect", "A quicker way to check",
    [
      "Look back at our answers.",
      "",
      "3 ÷ 1/4 = 12.",
      "2 ÷ 1/3 = 6.",
      "2 ÷ 1/2 = 4.",
      "",
      "Whole × bottom number = how many fit.",
      "The parts are smaller, so more fit.",
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
        "3 ÷ 1/4  =  3 × 4  =  12",
        "2 ÷ 1/3  =  2 × 3  =  6",
        "4 ÷ 1/5  =  4 × 5  =  20",
      ];
      rows.forEach((r, i) => {
        addTextOnShape(slide, r, {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.65 + i * 0.72, w: lg.rightW - 0.6, h: 0.55, rectRadius: 0.08,
          fill: { color: i === 2 ? C.PRIMARY : C.SECONDARY },
        }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
    }
  );

  // Slide 15: You Do — practice sheet (new numbers)
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: how many fit?", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.35, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 17, color: C.ALERT, bold: true } },
      { text: "Use the strips. How many halves fit into 3?", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
      { text: "Next: ", options: { fontSize: 17, color: C.ALERT, bold: true } },
      { text: "Use the number line. How many quarters fit into 2?", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
      { text: "Then: ", options: { fontSize: 17, color: C.ALERT, bold: true } },
      { text: "Explain why the answer is bigger than the whole.", options: { fontSize: 16, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.12, w: 8.5, h: 1.12,
      fontFace: FONT_B, valign: "middle", margin: 0, paraSpaceAfter: 4, fit: "shrink", shrinkText: true,
    });

    // Tool reminder panel with the two models
    const panelY = CONTENT_TOP + 1.52;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("Remember your two tools", {
      x: 0.7, y: panelY + 0.14, w: 8.6, h: 0.30,
      fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
    });
    // Strip reminder (a whole split into halves)
    s.addText("Strip: split each whole into equal parts, then count.", {
      x: 0.8, y: panelY + 0.52, w: 4.0, h: 0.3,
      fontSize: 12, fontFace: FONT_B, color: C.MUTED, margin: 0,
    });
    addFractionStripSet(s, 0.8, panelY + 0.86, 3.8, 0.5,
      [{ denom: 2, shaded: 0, label: "1 whole", color: C.SECONDARY }], { labelW: 0.9 });
    // Number line reminder (0-2 in quarters)
    s.addText("Number line: jump one part at a time, then count the jumps.", {
      x: 5.1, y: panelY + 0.52, w: 4.2, h: 0.3,
      fontSize: 12, fontFace: FONT_B, color: C.MUTED, margin: 0,
    });
    drawUnitFractionLine(s, 5.25, panelY + 1.05, 3.9, 2, 4, { labelFontSize: 12 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 16: Exit Ticket
  exitTicketSlide(pres,
    [
      "How many thirds (1/3) fit into 5 wholes? You may draw a strip or a number line.",
      "Explain how you know your answer is more than 5.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 17: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: why is 3 ÷ 1/4 more than 3?",
      scItems: SC,
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FracDiv_Lesson1_Divide_Whole_By_Unit_Fraction.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Companion PDFs ─────────────────────────────────────────────────────────

  // 1) Practice sheet (core You Do task, both models on paper)
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Find how many equal parts fit inside a whole number.",
      color: C.PRIMARY,
      lessonInfo: "Year 6 Numeracy | Fractions",
    });
    y = addTipBox(doc,
      "To divide by a unit fraction, ask: how many of these parts fit inside? Split each whole into equal parts, then count how many fit.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 - Fraction strips: 3 ÷ 1/2", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Split each whole strip into halves. Then count how many halves fit into 3 wholes.", y);
    y = drawPdfStrips(doc, y, 3, 0);
    y = addWriteLine(doc, "3 ÷ 1/2 =  ______ halves", y);

    y = addSectionHeading(doc, "Section 2 - Number line: 2 ÷ 1/4", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Mark each quarter on the line. Then count the jumps from 0 to 2.", y);
    y = drawPdfNumberLine(doc, y + 6, 2, 1);
    y = addWriteLine(doc, "2 ÷ 1/4 =  ______ quarters", y);

    y = addSectionHeading(doc, "Section 3 - Explain your thinking", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Why is the answer bigger than the whole number you started with?", y);
    y = addLinedArea(doc, y + 4, 2);

    y = addSectionHeading(doc, "Section 4 - Challenge", y, { color: C.ACCENT });
    y = addWriteLine(doc, "How many eighths fit into 2 wholes?    2 ÷ 1/8 =  ______", y + 4);
    y = addBodyText(doc, "Explain how you worked it out without drawing all the parts.", y);
    y = addLinedArea(doc, y + 2, 2);

    addPdfFooter(doc, "Session 1 | Practice Sheet | Year 6 Numeracy");
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

    y = addSectionHeading(doc, "Section 1 - 3 ÷ 1/2", y, { color: C.PRIMARY });
    y = addBodyText(doc, "3 ÷ 1/2 = 6.  Each whole splits into 2 halves, so 3 wholes = 3 × 2 = 6 halves.", y);

    y = addSectionHeading(doc, "Section 2 - 2 ÷ 1/4", y, { color: C.PRIMARY });
    y = addBodyText(doc, "2 ÷ 1/4 = 8.  Each whole splits into 4 quarters, so 2 wholes = 2 × 4 = 8 quarters.", y);

    y = addSectionHeading(doc, "Section 3 - Explain (sample answer)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "The parts (halves or quarters) are smaller than a whole, so several of them fit inside each whole. Dividing by a unit fraction counts how many small parts fit, so the answer is larger than the number we started with.", y);

    y = addSectionHeading(doc, "Section 4 - Challenge", y, { color: C.ACCENT });
    y = addBodyText(doc, "2 ÷ 1/8 = 16.  Each whole = 8 eighths, so 2 × 8 = 16. (Whole × bottom number = how many fit.)", y);

    y = addTipBox(doc,
      "Watch for: students who write an answer smaller than the whole - they shared into groups instead of counting how many parts fit. Redirect to the strip or number line: how many of these parts fit inside?",
      y, { color: C.ACCENT });

    addPdfFooter(doc, "Session 1 | Answer Key | Year 6 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // 3) Enabling scaffold — parts pre-drawn, student counts (changes task form)
  await (async () => {
    const doc = createPdf({ title: SCAFFOLD_RES.name });
    let y = addPdfHeader(doc, SCAFFOLD_RES.name, {
      subtitle: "The parts are drawn for you. Count how many fit, then write the number.",
      color: C.SECONDARY,
      lessonInfo: "Year 6 Numeracy | Fractions",
    });
    y = addTipBox(doc,
      "Count how many equal parts fit inside. Point to each part as you count.",
      y, { color: C.SECONDARY });

    y = addSectionHeading(doc, "Section 1 - How many halves fit into 3?", y, { color: C.SECONDARY });
    y = addBodyText(doc, "The halves are drawn. Count every half.", y);
    y = drawPdfStrips(doc, y, 3, 2);
    y = addWriteLine(doc, "3 ÷ 1/2 =  ______ halves", y);

    y = addSectionHeading(doc, "Section 2 - How many quarters fit into 2?", y, { color: C.SECONDARY });
    y = addBodyText(doc, "The quarters are marked. Count the jumps from 0 to 2.", y);
    y = drawPdfNumberLine(doc, y + 6, 2, 4);
    y = addWriteLine(doc, "2 ÷ 1/4 =  ______ quarters", y);

    y = addSectionHeading(doc, "Section 3 - Finish the sentences", y, { color: C.SECONDARY });
    y = addWriteLine(doc, "A quarter is  ____________  than a whole.", y);
    y = addWriteLine(doc, "So  ______  quarters fit into 2 wholes.", y);

    addPdfFooter(doc, "Session 1 | Enabling Scaffold | Year 6 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, SCAFFOLD_RES.fileName));
    console.log("PDF written: " + SCAFFOLD_RES.fileName);
  })();

  console.log("Lesson build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
