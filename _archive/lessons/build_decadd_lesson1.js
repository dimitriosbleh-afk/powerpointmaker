"use strict";

// Adding & Subtracting Decimals Unit (Year 5/6 Numeracy) — Lesson 1: Understanding decimals
// First-time learning. Foundational place value (tenths and hundredths) before any operations.
// Daily Review: identifying 2D shapes.
// Fluency: multiplication facts (x2).

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

// Theme: Year 5/6 numeracy. Variant fixed across all 5 lessons of this unit.
const T = createTheme("numeracy", "grade56", weekToVariant(3));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  dailyReviewSlide, fluencySlide,
  addPlaceValueChart, addDecimalDot,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const TOTAL = 5;
const UNIT_TITLE = "Adding and Subtracting Decimals";
const FOOTER = `Add & Subtract Decimals | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/DecAdd_Lesson1_Understanding_Decimals";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 1 Decimal Place Value Practice",
  "Name the place of each digit, write decimals into a chart, and shade hundredths grids.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 1 Answer Key",
  "Worked answers for the decimal place value practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome to a brand new unit on adding and subtracting decimals.
- This is the first time we are learning this together, so we are going to take it slowly.
- Today we are not adding yet. Today we are building the place value foundation that makes adding decimals easy later in the week.

DO:
- Have whiteboards, markers, and printed place value reference cards ready.
- Settle students before clicking past the title.

TEACHER NOTES:
Lesson 1 of 5. This week is the first explicit teaching of adding and subtracting decimals for this cohort. Treat this as a careful refresh of what a decimal is. Without secure tenths and hundredths, addition collapses in Lesson 2.

WATCH FOR:
- Students who look unsure - that is expected on day one. Reassure: "If this feels new, that is okay."

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The decimal place value practice sheet is for the You Do section.

DO:
- Print one copy of the practice sheet and answer key per student.
- Have whiteboards, markers, and the printed place value reference cards ready.
- Optional: MAB blocks or 10x10 hundredths grid printouts for small group support.

TEACHER NOTES:
One student resource (practice sheet) plus answer key. Most early work is on whiteboards. Hundredths grids printed at A4 can act as concrete support.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Look at each shape.
- Whisper the name to your partner, then write it on your whiteboard.

DO:
- Display the three shapes.
- 60 seconds.
- Walk and listen for correct names.

TEACHER NOTES:
Daily Review is identifying 2D shapes this week. Today we start with three common shapes. Listen for "triangle", "rectangle", "hexagon".

WATCH FOR:
- Students who name them confidently - secure.
- Students who say "diamond" for a rhombus or "stop sign" for octagon - prompt the formal name.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- The first shape is a triangle - 3 sides.
- The second shape is a rectangle - 4 sides, opposite sides equal.
- The third shape is a hexagon - 6 sides.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick-and-fix.

TEACHER NOTES:
Inverse operations are the strategy. Note any student still uncertain - small group focus this week.

WATCH FOR:
- Students who self-correct quickly - secure.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Multiplication facts.
- Today is the two times tables.
- Whisper-answer each one, then write it on your board.

DO:
- Display the three prompts.
- 45 seconds.

TEACHER NOTES:
Fluency this week is multiplication. We start with x2 to make sure every student gets a brisk success on day one. Tighter facts come in later lessons.

WATCH FOR:
- Students who answer instantly - secure.
- Students who count up - prompt them to remember "doubling."

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 2 times 6 is 12.
- 2 times 9 is 18.
- 2 times 7 is 14.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Tick and fix. Look for students who still count up rather than recall.

WATCH FOR:
- Students who self-correct - secure.
- Students who keep counting - small group focus.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to understand decimals using place value.
- Now the success criteria.

DO:
- Choral read.
- Hold up the printed place value chart.

TEACHER NOTES:
SC1 is achievable for everyone today - it is the naming move. SC2 is the core target. SC3 stretches to comparing two decimals.

WATCH FOR:
- Students who can repeat the language - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. This is a hundredths grid. It has 100 small squares.
- The whole grid is 1.
- If I shade 1 column - 10 squares - that is 1 tenth, or 0.1.
- If I shade 1 small square, that is 1 hundredth, or 0.01.
- A decimal is just a way of writing parts of a whole.

DO:
- Point to the whole grid first.
- Point to the shaded column.
- Point to the single hundredth square.
- Repeat "one tenth" and "one hundredth" three times each.

TEACHER NOTES:
This is the concrete anchor for the whole unit. The hundredths grid lets students SEE that one tenth is much bigger than one hundredth. Many students arrive thinking 0.07 is bigger than 0.7 because it has more digits. The grid breaks that misconception.

MISCONCEPTIONS:
- Misconception: Students think more digits means bigger.
  Why: They transfer whole-number thinking.
  Impact: They cannot compare decimals later this week.
  Quick correction: Use the grid. "1 tenth is a whole column. 1 hundredth is one tiny square."

WATCH FOR:
- Students who nod when you say "1 tenth is bigger" - tracking.
- Students who still look confused - we use the chart again in the CFU.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Watch me put 2.45 into a place value chart.
- The 2 sits in the ones place. That is 2 ones.
- The decimal point is the marker - it separates whole numbers from parts.
- The 4 sits in the tenths place. That is 4 tenths.
- The 5 sits in the hundredths place. That is 5 hundredths.
- I read it: "two and forty-five hundredths."

DO:
- Point to each cell as you say the name.
- Trace the decimal point as a marker, not a separator.
- Repeat the read-it sentence twice.

TEACHER NOTES:
This is the second core move. The decimal point is a marker, not a divider. Many Year 5/6 students learning this for the first time say "two point four five" - that is not wrong but it hides the place value. Push the place-value language.

MISCONCEPTIONS:
- Misconception: Students read 2.45 as "two point forty-five."
  Why: They treat the decimal point as a divider.
  Impact: They cannot compare or add decimals later.
  Quick correction: "Each digit has a place name. Read the smallest place at the end - 'two and forty-five hundredths'."

WATCH FOR:
- Students who repeat the place names - tracking.
- Students who say "two point four five" - prompt: "use the place name."

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. The number is 3.07.
- On your whiteboard, write each digit under the right place name.
- Ones, tenths, hundredths.

DO:
- Display the prompt.
- 30 seconds.
- Walk and scan.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your chart on three, two, one, show."
- Scan for: 3 in ones, 0 in tenths, 7 in hundredths.
PROCEED: If 80% have the correct chart, click to reveal and move to We Do.
PIVOT: Most likely misconception - students skip the zero and write 3, 7.
- Reteach: "The zero holds the tenths place. We do not skip it."
- Re-check: "Where does the zero go in 3.07?"

TEACHER NOTES:
Place value zero is the trap. Students who skip the zero will later misread 3.07 as 3.7.

WATCH FOR:
- Students who hold the zero in tenths - secure.
- Students who skip the zero - place value misconception.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Put 1.36 into the chart.
- And shade 1.36 on the second hundredths grid (1 whole grid plus 36 hundredths on the next).
- Whisper to your partner: how many tenths, how many hundredths?

DO:
- Display 1.36 above two blank charts on the slide.
- 90 seconds.
- Walk and listen.

TEACHER NOTES:
Same structure as the I Do but new numbers and an additional grid task. Listen for "3 tenths and 6 hundredths" - that is the relational language.

WATCH FOR:
- Pairs who say "3 tenths and 6 hundredths" - secure.
- Pairs who shade 36 squares but cannot name the tenths - reteach with the column.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- 1 one. Decimal point. 3 tenths. 6 hundredths.
- That is 1 whole grid shaded fully, plus 3 columns and 6 single squares on the next grid.
- We say: "one and thirty-six hundredths."

DO:
- Click to reveal.
- Run through "3 tenths, 6 hundredths" one more time.

TEACHER NOTES:
Reveal restates the place names. Students who got the chart right know the place. Students who got the shading right know the size.

WATCH FOR:
- Students who self-correct - secure.
- Students who shaded 136 squares on one grid - reteach with the "1 whole grid is 1 one" idea.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. Two students disagree.
- Mia says 0.5 is bigger than 0.45 because 5 tenths is bigger than 4 tenths.
- Sam says 0.45 is bigger because it has more digits.
- Thumbs UP for Mia. Thumbs DOWN for Sam.

DO:
- Display the disagreement.
- 15 seconds.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Thumbs UP if you agree with Mia. DOWN if you agree with Sam."
- Scan for: thumbs UP. Mia is right.
PROCEED: If 80% agree with Mia, click to reveal and confirm the rule.
PIVOT: Most likely misconception - students agree with Sam (more digits = bigger).
- Reteach: Use a hundredths grid. 0.5 = 50 hundredths shaded. 0.45 = 45 hundredths shaded.
- Re-check: "Which has more shaded? Which is bigger?"

TEACHER NOTES:
This hinge probes whether students apply the "compare the same place" rule. If most agree with Sam, we need another grid model before You Do.

WATCH FOR:
- Confident thumbs up for Mia - they see the place value.
- Thumbs down or sideways - they need the grid again.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Take the practice sheet.
- Section 1: name the place of the underlined digit.
- Section 2: write each decimal into a chart.
- Section 3: shade the hundredths grid to match the decimal.

DO:
- Distribute the practice sheet.
- Circulate. Listen for place-name language.
- Cold call 1-2 students to share Section 3 thinking.

TEACHER NOTES:
Different numbers from the We Do. Same strategy: place value chart, name the place, link to the grid.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use only Section 1 and Section 2 with the printed place value reference card.
- Extra Notes: Sit with these students. Do the first one together.
EXTENDING PROMPT:
- Task: Section 4 - order four decimals from smallest to largest. Write one sentence explaining how you decided.
- Extra Notes: Encourage place-value language in the explanation.

WATCH FOR:
- Students who write the place name fluently - secure.
- Students who write "point seven" instead of "seven tenths" - prompt them to use the place name.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Question 1: Write 4.08 into a place value chart. Name each digit's place.
- Question 2: Which is bigger, 0.7 or 0.65? Circle it. Explain using the words tenths and hundredths.

DO:
- Display the prompt.
- 3 minutes.
- Collect whiteboards or photograph.

TEACHER NOTES:
Exit ticket assesses SC2. Look for the zero in tenths for Q1, and "compare the tenths first" reasoning for Q2.

WATCH FOR:
- Students who write 4 in ones, 0 in tenths, 8 in hundredths - secure.
- Students who circle 0.65 - still using "more digits is bigger" thinking. Small group focus tomorrow.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways, or down for each one.
- Turn and tell your partner: what does the decimal point really mean?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call 1-2 students for the partner share.

TEACHER NOTES:
The threshold idea is "the decimal point separates the whole from the parts, and each place has a value." Students who can say this are ready for adding decimals tomorrow.

WATCH FOR:
- Strong thumbs up across all three - move at pace tomorrow.
- Sideways or down on SC2 - small group revision in Lesson 2.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Helpers ────────────────────────────────────────────────────────────────

// Draw a 10x10 hundredths grid. shadedColumns = number of full columns shaded (tenths).
// extraSquares = number of additional single squares shaded (hundredths).
function addHundredthsGrid(slide, x, y, sizeIn, shadedColumns, extraSquares, opts) {
  const o = opts || {};
  const fill = o.fillColor || C.PRIMARY;
  const lineColor = o.lineColor || C.MUTED;
  const bgColor = o.bgColor || C.WHITE;
  const cell = sizeIn / 10;

  // Background
  slide.addShape("rect", {
    x, y, w: sizeIn, h: sizeIn,
    fill: { color: bgColor },
    line: { color: lineColor, width: 0.75 },
  });

  // Shade full columns (tenths)
  if (shadedColumns > 0) {
    slide.addShape("rect", {
      x, y, w: shadedColumns * cell, h: sizeIn,
      fill: { color: fill },
      line: { color: fill, width: 0 },
    });
  }

  // Shade extra single squares (hundredths) starting on the next column, top-down
  if (extraSquares > 0) {
    const col = shadedColumns;
    for (let i = 0; i < extraSquares; i++) {
      const row = i;
      slide.addShape("rect", {
        x: x + col * cell,
        y: y + row * cell,
        w: cell, h: cell,
        fill: { color: fill },
        line: { color: fill, width: 0 },
      });
    }
  }

  // Re-draw grid lines on top so cells stay visible
  for (let i = 1; i < 10; i++) {
    slide.addShape("line", {
      x: x + i * cell, y, w: 0, h: sizeIn,
      line: { color: lineColor, width: 0.5 },
    });
    slide.addShape("line", {
      x, y: y + i * cell, w: sizeIn, h: 0,
      line: { color: lineColor, width: 0.5 },
    });
  }
}

function placePvChart(slide, x, y, headers, values, opts) {
  const o = opts || {};
  const totalW = o.totalW || 4.0;
  const valH = o.valH || 0.55;
  const hdrH = o.hdrH || 0.34;
  const chart = addPlaceValueChart(slide, x, y, headers, values, {
    totalW, valH, hdrH,
    headerColor: o.headerColor || C.PRIMARY,
  });
  if (o.dotAfter != null) {
    addDecimalDot(slide, chart, o.dotAfter, { position: "baseline" });
  }
  return chart;
}

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 1: Understanding decimals",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal — identifying 2D shapes
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Daily Review: Name each 2D shape");

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });

      // Three shapes evenly spaced
      const shapeY = CONTENT_TOP + 0.5;
      // Triangle (left)
      s.addShape("triangle", {
        x: 1.3, y: shapeY, w: 1.6, h: 1.7,
        fill: { color: C.PRIMARY }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("A", {
        x: 1.3, y: shapeY + 1.8, w: 1.6, h: 0.4,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });

      // Rectangle (centre)
      s.addShape("rect", {
        x: 4.0, y: shapeY + 0.20, w: 2.0, h: 1.3,
        fill: { color: C.SECONDARY }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("B", {
        x: 4.0, y: shapeY + 1.8, w: 2.0, h: 0.4,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });

      // Hexagon (right) - drawn as a regular hexagon shape
      s.addShape("hexagon", {
        x: 7.1, y: shapeY + 0.10, w: 1.9, h: 1.6,
        fill: { color: C.ACCENT }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("C", {
        x: 7.1, y: shapeY + 1.8, w: 1.9, h: 0.4,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });

      // Bottom prompt
      s.addText("Write the name of each shape on your whiteboard.", {
        x: 0.7, y: SAFE_BOTTOM - 0.55, w: 8.6, h: 0.35,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "A = triangle     B = rectangle     C = hexagon", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal — multiplication (x2)
  withReveal(
    () => fluencySlide(pres, "Fluency: Multiplication (x2)",
      ["2 × 6", "2 × 9", "2 × 7"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "12     18     14", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: LI/SC
  liSlide(pres,
    "We are learning to understand decimals using place value.",
    [
      "I can name the place of each digit in a decimal (ones, tenths, hundredths).",
      "I can write a decimal into a place value chart and shade it on a hundredths grid.",
      "I can compare two decimals by lining up the place values and explain which is bigger.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 8: I Do — hundredths grid showing tenths vs hundredths
  workedExSlide(pres, 2, "I Do", "1 tenth vs 1 hundredth",
    [
      "The whole grid is 1.",
      "",
      "1 column shaded = 1 tenth = 0.1",
      "1 small square = 1 hundredth = 0.01",
      "",
      "A tenth is much bigger",
      "than a hundredth.",
      "",
      "Longer is NOT bigger.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.PRIMARY });
      slide.addText("Hundredths grid", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });

      const gridSize = 2.0;
      const gridX = lg.rightX + (lg.rightW - gridSize) / 2;
      const gridY = lg.panelTopPadded + 0.50;
      addHundredthsGrid(slide, gridX, gridY, gridSize, 1, 0,
        { fillColor: C.PRIMARY });

      // Single hundredth square highlighted in a different colour
      const cell = gridSize / 10;
      slide.addShape("rect", {
        x: gridX + 5 * cell, y: gridY + 3 * cell, w: cell, h: cell,
        fill: { color: C.ALERT }, line: { color: C.ALERT, width: 0 },
      });

      slide.addText("1 tenth (blue column)", {
        x: lg.rightX + 0.10, y: gridY + gridSize + 0.10, w: lg.rightW - 0.20, h: 0.28,
        fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      slide.addText("1 hundredth (red square)", {
        x: lg.rightX + 0.10, y: gridY + gridSize + 0.38, w: lg.rightW - 0.20, h: 0.28,
        fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 9: I Do — place value chart for 2.45
  workedExSlide(pres, 2, "I Do", "Place value chart: 2.45",
    [
      "Each digit has a place name.",
      "",
      "2 sits in the ones place.",
      "4 sits in the tenths place.",
      "5 sits in the hundredths place.",
      "",
      "Read: two and forty-five hundredths.",
      "",
      "The decimal point is the marker.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.SECONDARY });
      slide.addText("2.45", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.40,
        fontSize: 28, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      const chartY = lg.panelTopPadded + 0.60;
      placePvChart(slide, lg.rightX + 0.10, chartY,
        ["Ones", "Tenths", "Hundredths"],
        ["2", "4", "5"],
        { totalW: lg.rightW - 0.20, valH: 0.80, hdrH: 0.42, dotAfter: 0 });

      slide.addText("Read: two and forty-five hundredths", {
        x: lg.rightX + 0.10, y: chartY + 1.40, w: lg.rightW - 0.20, h: 0.30,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slides 10-11: CFU + reveal — write 3.07 in the chart
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Place 3.07 in the chart", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 20, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Write each digit of 3.07", options: { fontSize: 18, color: C.CHARCOAL, breakLine: true } },
        { text: "under the correct place name.", options: { fontSize: 18, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Watch the zero!", options: { fontSize: 20, color: C.ALERT, bold: true } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.3, CONTENT_TOP, 4.2, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      s.addText("3.07", {
        x: 5.3, y: CONTENT_TOP + 0.15, w: 4.2, h: 0.45,
        fontSize: 32, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      placePvChart(s, 5.45, CONTENT_TOP + 0.95,
        ["Ones", "Tenths", "Hundredths"],
        ["", "", ""],
        { totalW: 3.90, valH: 0.95, hdrH: 0.42, dotAfter: 0 });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Ones 3   |   Tenths 0   |   Hundredths 7", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 12-13: We Do + reveal — place 1.36 in the chart and shade the grid
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Build 1.36 with a chart and a grid",
      [
        "With your partner.",
        "",
        "1.  Write 1.36 into a chart on your whiteboard.",
        "2.  Picture it shaded on grids:",
        "    1 whole grid + 36 squares on the next.",
        "",
        "How many tenths? How many hundredths?",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.SECONDARY });
        slide.addText("1.36", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.34,
          fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });

        // Two side-by-side grids
        const gSize = 1.30;
        const gap = 0.18;
        const totalW = gSize * 2 + gap;
        const gx0 = lg.rightX + (lg.rightW - totalW) / 2;
        const gy = lg.panelTopPadded + 0.55;
        // Grid 1: fully shaded = 1 whole
        addHundredthsGrid(slide, gx0, gy, gSize, 10, 0, { fillColor: C.SECONDARY });
        slide.addText("1 whole", {
          x: gx0, y: gy + gSize + 0.05, w: gSize, h: 0.25,
          fontSize: 11, fontFace: FONT_B, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        // Grid 2: 36 hundredths shaded = 3 columns + 6 extra squares
        addHundredthsGrid(slide, gx0 + gSize + gap, gy, gSize, 3, 6,
          { fillColor: C.SECONDARY });
        slide.addText("36 hundredths", {
          x: gx0 + gSize + gap, y: gy + gSize + 0.05, w: gSize, h: 0.25,
          fontSize: 11, fontFace: FONT_B, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });

        slide.addText("How many tenths? How many hundredths?", {
          x: lg.rightX + 0.10, y: gy + gSize + 0.42, w: lg.rightW - 0.20, h: 0.30,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "Ones 1   Tenths 3   Hundredths 6   =   1.36   (one and thirty-six hundredths)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 14-15: CFU hinge + reveal — Mia vs Sam
  withReveal(
    () => cfuSlide(pres, "CFU", "Who is right?", "Thumbs Up or Thumbs Down",
      "Mia:  0.5 is bigger than 0.45.\n     (5 tenths vs 4 tenths.)\n\nSam: 0.45 is bigger because it has more digits.\n\nThumbs UP for Mia.   Thumbs DOWN for Sam.",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Mia is right.   0.5 has 5 tenths.   0.45 has 4 tenths.   0.5 > 0.45.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 16: You Do — practice sheet from the table
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: practice sheet", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.30, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 1 — name the place of each underlined digit.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 2 — write each decimal into a chart.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 3 — shade the hundredths grid.", options: { fontSize: 18, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.45;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    s.addText("Remember", {
      x: 0.7, y: panelY + 0.15, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });

    addTextOnShape(s, "Each digit has a place name. The decimal point is the marker.", {
      x: 1.0, y: panelY + 0.55, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    placePvChart(s, 2.5, panelY + 1.20,
      ["Ones", "Tenths", "Hundredths"],
      [".", "", ""],
      { totalW: 5.0, valH: 0.55, hdrH: 0.34, dotAfter: 0 });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "Write 4.08 into a place value chart. Name each digit's place.",
      "Which is bigger: 0.7 or 0.65? Circle it. Explain using the words tenths and hundredths.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 18: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what does the decimal point really mean?",
      scItems: [
        "I can name the place of each digit in a decimal (ones, tenths, hundredths).",
        "I can write a decimal into a place value chart and shade it on a hundredths grid.",
        "I can compare two decimals by lining up the place values and explain which is bigger.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecAdd_Lesson1_Understanding_Decimals.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Name the place, write it in a chart, and shade the grid.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "Each digit has a place name. The decimal point is the marker between ones and tenths. Longer is NOT bigger - compare the SAME place first.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 — Name the place of the underlined digit", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  2.6_5_      The 5 is in the _________________ place.", y);
    y = addWriteLine(doc, "b)  4._3_       The 3 is in the _________________ place.", y);
    y = addWriteLine(doc, "c)  _7_.18      The 7 is in the _________________ place.", y);
    y = addWriteLine(doc, "d)  0.0_9_      The 9 is in the _________________ place.", y);

    y = addSectionHeading(doc, "Section 2 — Write each decimal into the place value chart", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Headers: Ones | Tenths | Hundredths", y, { fontSize: 10, color: C.MUTED });
    y = addWriteLine(doc, "a)  1.42       Ones: ___    Tenths: ___    Hundredths: ___", y);
    y = addWriteLine(doc, "b)  3.07       Ones: ___    Tenths: ___    Hundredths: ___", y);
    y = addWriteLine(doc, "c)  0.85       Ones: ___    Tenths: ___    Hundredths: ___", y);

    y = addSectionHeading(doc, "Section 3 — Shade the grid to match the decimal", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Use the 10x10 grids beside each decimal. Shade full columns for tenths, single squares for hundredths.", y);
    y = addWriteLine(doc, "a)  0.4   (Shade 4 full columns.)", y);
    y = addWriteLine(doc, "b)  0.23  (Shade 2 full columns plus 3 single squares.)", y);
    y = addWriteLine(doc, "c)  0.07  (Shade 7 single squares only.)", y);

    y = addSectionHeading(doc, "Section 4 — Extension (optional, order four decimals)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Order the four decimals from smallest to largest. Write one sentence to explain.", y);
    y = addWriteLine(doc, "0.5    0.45    0.50    0.05", y);
    y = addWriteLine(doc, "Order: ____________________________________________________________", y);
    y = addWriteLine(doc, "Why:   ____________________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Decimal Place Value Practice | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the decimal place value practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 — Name the place", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  hundredths      b)  tenths      c)  ones      d)  hundredths", y);

    y = addSectionHeading(doc, "Section 2 — Place value chart", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1.42    Ones 1 | Tenths 4 | Hundredths 2", y);
    y = addBodyText(doc, "b)  3.07    Ones 3 | Tenths 0 | Hundredths 7   (watch the zero in tenths)", y);
    y = addBodyText(doc, "c)  0.85    Ones 0 | Tenths 8 | Hundredths 5", y);

    y = addSectionHeading(doc, "Section 3 — Shade the grid", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  0.4  = 4 full columns shaded (40 small squares).", y);
    y = addBodyText(doc, "b)  0.23 = 2 columns plus 3 single squares (23 small squares).", y);
    y = addBodyText(doc, "c)  0.07 = 7 single squares only.", y);

    y = addSectionHeading(doc, "Section 4 — Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "Order: 0.05  <  0.45  <  0.5 = 0.50.", y);
    y = addBodyText(doc, "Why: 0.05 has zero tenths. 0.45 has 4 tenths. 0.5 and 0.50 both have 5 tenths and zero (or no) hundredths.", y);

    y = addTipBox(doc,
      "Watch for: students who skip the zero in 3.07 (the place value zero is the trap); students who pick 0.65 as bigger than 0.7 (more digits = bigger misconception); students who think 0.5 is smaller than 0.50.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 1 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
