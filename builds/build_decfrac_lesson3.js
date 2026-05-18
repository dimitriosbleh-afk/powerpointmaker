"use strict";

// Decimals & Fractions Unit (Year 5/6 Numeracy) — Lesson 3: Adding decimals using place value
// VC2M6N02 — apply knowledge of place value to add decimals; mental strategies via
//            basic facts, place value, partitioning and properties of operations.
// Daily Review: Solving equations with multiplication, division, and operations.
// Fluency: Adding and subtracting decimals.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter,
  addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  dailyReviewSlide, fluencySlide,
  addPlaceValueChart, addDecimalDot,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 3;
const TOTAL = 6;
const UNIT_TITLE = "Decimals and Fractions";
const FOOTER = `Decimals & Fractions | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/DecFrac_Lesson3_Adding_Decimals";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 3 Adding Decimals Practice",
  "Adding decimals using place value charts and the line-up strategy.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 3 Answer Key",
  "Worked answers for the adding decimals practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Today we add decimals using place value.
- We line up the decimal point so digits sit in the same place value column.
- Then we add column by column, just like whole numbers.

DO:
- Have whiteboards, markers, and printed place value reference charts ready.

TEACHER NOTES:
Lesson 3 of 6. The core strategy is "line up the decimal point" so that tenths add to tenths, hundredths to hundredths, and so on. This sits directly on Lessons 1 and 2.

WATCH FOR:
- Students who line up the right-hand digit - that is the place value trap. The estimation step from Lesson 2 catches it.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The adding decimals practice sheet is for the You Do section.

DO:
- Print one copy of the practice sheet and answer key.
- Have whiteboards, markers, and place value reference charts ready.

TEACHER NOTES:
One student resource (practice sheet) plus answer key.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Warm up. Find the missing number in each equation.
- Whisper to your partner, then write on your board.

DO:
- Display the three equations.
- 90 seconds.

TEACHER NOTES:
Daily Review continues with equation-solving from Lessons 1 and 2.

WATCH FOR:
- Students using the inverse confidently - secure.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- 8 times n equals 72. n is 9 (because 72 divided by 8 is 9).
- n divided by 6 equals 7. n is 42 (because 7 times 6 is 42).
- 13 + n = 25. n is 12.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Same retrieval. Students who still guess need small group support.

WATCH FOR:
- Students who self-correct - secure.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Adding and subtracting decimals.
- Line up the decimal point.

DO:
- Display the three prompts.
- 60 seconds.

TEACHER NOTES:
Fluency builds toward today's I Do. These are tenths and an easier hundredths. The line-up habit matters.

WATCH FOR:
- Students who line up the decimal point - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 0.7 plus 0.3 is 1.0.
- 2.4 minus 1.2 is 1.2.
- 0.5 plus 0.25 is 0.75.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
The third one mixes tenths and hundredths. Watch for students who write 0.30 (5 + 25 = 30, ignoring place value).

WATCH FOR:
- Students who line up correctly even with different decimal lengths - secure.
- Students who wrote 0.30 - reteach with the place value chart in I Do.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to add decimals using place value, lining up the decimal point.
- Now the success criteria.

DO:
- Choral read.
- Point to a printed place value chart.

TEACHER NOTES:
SC1 is achievable - lining up decimals already shown in a chart. SC2 is the core target - adding decimals where students must align themselves. SC3 stretches to mental partitioning strategies.

WATCH FOR:
- Students who can repeat the language - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me add 3.4 and 2.5.
- Step 1: I estimate first. 3 plus 3 is about 6. So my answer should be around 6.
- Step 2: I line up the decimal points so the tenths sit under the tenths and the ones sit under the ones.
- Step 3: I add column by column, starting from the right. Tenths: 4 plus 5 is 9. Ones: 3 plus 2 is 5.
- Step 4: I bring down the decimal point.
- Step 5: 5.9. I check against my estimate of 6. Yes, that is reasonable.

DO:
- Display the worked example with the place value chart.
- Point to each column as you add.
- Repeat the estimate-add-check rhythm.

TEACHER NOTES:
The estimate-first habit from Lesson 2 carries straight into Lesson 3. It catches misalignment errors before they happen.

MISCONCEPTIONS:
- Misconception: Line up the right-hand digit.
  Why: Students copy whole-number addition habits.
  Impact: 3.4 + 2.15 would be added wrongly.
  Quick correction: "Line up the decimal point, not the right-hand digit."

WATCH FOR:
- Students who add the columns in the right place - secure.
- Students who line up the right-hand digit - the trap.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now I add 3.4 and 2.15 - decimals with different lengths.
- Estimate first: 3 plus 2 is 5. So my answer should be around 5.
- Watch how I line them up. 3.4 has nothing in the hundredths. I can write a zero placeholder: 3.40.
- Now they are the same length: 3.40 and 2.15.
- Add column by column.
- Hundredths: 0 plus 5 is 5. Tenths: 4 plus 1 is 5. Ones: 3 plus 2 is 5.
- 5.55. Check against estimate of 5. Yes, reasonable.

DO:
- Display both numbers stacked with the placeholder zero shown.
- Underline the placeholder in a different colour.

TEACHER NOTES:
The placeholder zero is the key move. Without it, students misalign. Some texts call this "make the decimals the same length."

MISCONCEPTIONS:
- Misconception: 3.4 + 2.15 = 3.55 (right-shift align).
  Why: They lined up the right-hand digits.
  Impact: Wildly wrong answers.
  Quick correction: "Write the placeholder zero so both numbers have the same number of decimal places."

WATCH FOR:
- Students who write the placeholder zero - secure.
- Students who line up the right-hand digit - reteach with the chart.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. Add 4.6 + 2.3.
- On your whiteboard. Show me.

DO:
- Display the prompt.
- 30 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your answer on three, two, one, show."
- Scan for: 6.9.
PROCEED: If 80% have 6.9, click to reveal and move on.
PIVOT: Most likely misconception - students lined up wrong or carried.
- Reteach: "Tenths plus tenths first. 6 + 3 = 9 tenths. Then ones."
- Re-check: "Add 4.6 + 2.3 again."

TEACHER NOTES:
Same-length tenths. Quick check before we move to different-length decimals.

WATCH FOR:
- Students who hold up 6.9 - secure.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Add 5.27 + 1.4.
- First: ESTIMATE. About 5 + 1 = 6.
- Then: line up the decimal point. Write a placeholder zero where needed.
- Add column by column.
- Finally: check against your estimate.

DO:
- Display the two numbers and the place value chart.
- 90 seconds.
- Walk and listen for "placeholder zero."

TEACHER NOTES:
Different-length decimals. The placeholder zero is the move that separates secure from shaky.

WATCH FOR:
- Students who add a placeholder zero before adding - secure.
- Students who write 5.67 (didn't align) - reteach.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- 5.27 + 1.40 (placeholder zero on 1.4).
- Hundredths: 7 + 0 = 7.
- Tenths: 2 + 4 = 6.
- Ones: 5 + 1 = 6.
- 6.67. Estimate was 6 - yes, reasonable.

DO:
- Click to reveal.
- Run through the placeholder zero step once more.

TEACHER NOTES:
The reveal restates the placeholder-zero move. Students who got 6.67 used it. Students who got 5.67 did not align.

WATCH FOR:
- Students who self-correct - secure.
- Students who keep 5.67 - reteach with the chart.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. Look at this student's work.
- 2.3 + 0.45 = 6.8.
- Use your estimation. Is this reasonable? Thumbs UP for yes, DOWN for no.

DO:
- Display the work.
- 15 seconds.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Thumbs UP if reasonable. DOWN if not."
- Scan for: thumbs DOWN. 2 + 0 is about 2-3, not 6-7.
PROCEED: If 80% thumbs down, click to reveal.
PIVOT: Most likely misconception - students try to recalculate instead of estimating.
- Reteach: "Estimate first. 2 + 0 is about 2. The student lined up wrongly and got 6.8 because they added 23 + 45."
- Re-check: "What is the real answer?"

TEACHER NOTES:
This hinge probes whether students use estimation as a reasonableness check from Lesson 2.

WATCH FOR:
- Confident thumbs down - they used estimation.
- Students who hesitate - reteach the estimate-first habit.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own. Take the practice sheet.
- Section 1: same-length decimals.
- Section 2: different-length decimals - remember the placeholder zero.
- Section 3: word problem.

DO:
- Distribute the practice sheet.
- Circulate.
- Cold call 1-2 students to share Section 3 thinking.

TEACHER NOTES:
The practice sheet sequences from same-length to different-length, then a word problem.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use only Section 1 with the printed place value chart.
- Extra Notes: Sit with these students. Do the first one together.
EXTENDING PROMPT:
- Task: Section 4 - mental partitioning. Solve 4.7 + 2.6 mentally by partitioning 4.7 into 4 + 0.7.
- Extra Notes: Ask students to record their partition step.

WATCH FOR:
- Students who use the placeholder zero - secure on different-length decimals.
- Students who line up the right-hand digit - reteach.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task.
- Question 1: 6.3 + 2.85. Show your working.
- Question 2: A student wrote 4.2 + 3.15 = 7.17. Is this correct? If not, fix it.

DO:
- Display the prompt.
- 3 minutes.
- Collect whiteboards.

TEACHER NOTES:
Exit ticket assesses SC2 (adding decimals with different lengths). Q2 probes the placeholder-zero understanding.

WATCH FOR:
- Students who got 9.15 with placeholder zero - secure.
- Students who fix Q2 to 7.35 - they understand the alignment trap.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria. Thumbs for each.
- Turn and tell your partner: why do we write a placeholder zero?

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
The placeholder zero is the lesson's threshold concept. Students who can explain it are ready for subtraction tomorrow.

WATCH FOR:
- Students who say "to line up the place values" - secure.
- Students unsure - revisit at the start of Lesson 4.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Helpers ────────────────────────────────────────────────────────────────

function placePvChart(slide, x, y, headers, values, opts) {
  const o = opts || {};
  return addPlaceValueChart(slide, x, y, headers, values, {
    totalW: o.totalW || 4.0,
    valH: o.valH || 0.55,
    hdrH: o.hdrH || 0.34,
    headerColor: o.headerColor || C.PRIMARY,
  });
}

// Stack two decimals visually with a decimal point and an answer line.
function addStackedAddition(slide, x, y, w, top, bottom, answer, opts) {
  const o = opts || {};
  const lineColor = o.color || C.PRIMARY;
  const fontSize = o.fontSize || 28;
  const rowH = o.rowH || 0.46;

  // Top number
  slide.addText(top, {
    x, y, w, h: rowH,
    fontSize, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  // Plus sign and bottom number on same row
  slide.addText("+   " + bottom, {
    x, y: y + rowH, w, h: rowH,
    fontSize, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  // Horizontal line
  slide.addShape("line", {
    x: x + 0.4, y: y + 2 * rowH + 0.04, w: w - 0.8, h: 0,
    line: { color: lineColor, width: 2 },
  });
  // Answer (or blank)
  if (answer) {
    slide.addText(String(answer), {
      x, y: y + 2 * rowH + 0.12, w, h: rowH,
      fontSize, fontFace: FONT_H, color: lineColor, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  }
}

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 3: Adding decimals using place value",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal
  withReveal(
    () => dailyReviewSlide(pres,
      "Daily Review: Find the missing number",
      [
        "8 × n = 72",
        "n ÷ 6 = 7",
        "13 + n = 25",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "n = 9     n = 42     n = 12", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal
  withReveal(
    () => fluencySlide(pres, "Fluency: Adding & Subtracting Decimals",
      ["0.7 + 0.3", "2.4 − 1.2", "0.5 + 0.25"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "1.0     1.2     0.75", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: LI/SC
  liSlide(pres,
    "We are learning to add decimals using place value, lining up the decimal point.",
    [
      "I can use a place value chart to add two decimals.",
      "I can add two decimals (with different decimal lengths) by lining up the decimal point.",
      "I can use estimation and partitioning to check my answer makes sense.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 8: I Do — add 3.4 + 2.5 (same length)
  workedExSlide(pres, 2, "I Do", "Add 3.4 + 2.5",
    [
      "1.  ESTIMATE: 3 + 3 ≈ 6.",
      "",
      "2.  LINE UP the decimal point.",
      "",
      "3.  ADD column by column.",
      "     Tenths: 4 + 5 = 9.",
      "     Ones:   3 + 2 = 5.",
      "",
      "4.  BRING DOWN the decimal point.",
      "",
      "5.  CHECK: 5.9 ≈ 6  ✓",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("Stacked addition", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      // Column headers
      slide.addText("Ones    Tenths", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.20, h: 0.26,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });
      // Stacked addition: 3.4 + 2.5 = 5.9
      addStackedAddition(slide, lg.rightX + 0.30, lg.panelTopPadded + 0.78,
        lg.rightW - 0.60, "3.4", "2.5", "5.9",
        { color: C.PRIMARY, fontSize: 36, rowH: 0.50 });
      slide.addText("3.4 + 2.5 = 5.9", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 2.70, w: lg.rightW - 0.20, h: 0.26,
        fontSize: 14, fontFace: FONT_H, color: C.SUCCESS, bold: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 9: I Do — add 3.4 + 2.15 (different length, placeholder zero)
  workedExSlide(pres, 2, "I Do", "Add 3.4 + 2.15 (different lengths)",
    [
      "1.  ESTIMATE: 3 + 2 ≈ 5.",
      "",
      "2.  Write a PLACEHOLDER ZERO:",
      "     3.4 becomes 3.40.",
      "",
      "3.  LINE UP the decimal point.",
      "",
      "4.  ADD column by column.",
      "     Hundredths: 0 + 5 = 5.",
      "     Tenths:     4 + 1 = 5.",
      "     Ones:       3 + 2 = 5.",
      "",
      "5.  5.55 ≈ 5  ✓",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY });
      slide.addText("Placeholder zero in red", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      slide.addText("Ones    Tenths    Hundredths", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.20, h: 0.26,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });
      // Stack: 3.40 + 2.15 = 5.55
      // We render the placeholder 0 in alert color via two separate text overlays
      // to keep the visual story clear.
      const stackX = lg.rightX + 0.30;
      const stackW = lg.rightW - 0.60;
      const rowH = 0.50;
      const rowY = lg.panelTopPadded + 0.78;

      // Top: "3.40" with the placeholder 0 in red, rendered as a single text
      // element using a rich-text run array.
      slide.addText([
        { text: "3.4",
          options: { fontSize: 36, fontFace: FONT_H, color: C.CHARCOAL, bold: true } },
        { text: "0",
          options: { fontSize: 36, fontFace: FONT_H, color: C.ALERT, bold: true } },
      ], {
        x: stackX, y: rowY, w: stackW, h: rowH,
        align: "center", valign: "middle", margin: 0,
      });

      // Plus row: "+   2.15"
      slide.addText("+   2.15", {
        x: stackX, y: rowY + rowH, w: stackW, h: rowH,
        fontSize: 36, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Line
      slide.addShape("line", {
        x: stackX + 0.4, y: rowY + 2 * rowH + 0.04, w: stackW - 0.8, h: 0,
        line: { color: C.SECONDARY, width: 2 },
      });

      // Answer
      slide.addText("5.55", {
        x: stackX, y: rowY + 2 * rowH + 0.12, w: stackW, h: rowH,
        fontSize: 36, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // Slides 10-11: CFU + reveal — add 4.6 + 2.3
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Add 4.6 + 2.3", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 20, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Add 4.6 + 2.3", options: { fontSize: 28, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "1.  Estimate first.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "2.  Line up the decimal point.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "3.  Add column by column.", options: { fontSize: 16, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right card: stacked addition visualisation
      addCard(s, 5.3, CONTENT_TOP, 4.2, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      s.addText("Stacked addition", {
        x: 5.3, y: CONTENT_TOP + 0.10, w: 4.2, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      addStackedAddition(s, 5.6, CONTENT_TOP + 0.70, 3.6, "4.6", "2.3", null,
        { color: C.PRIMARY, fontSize: 38 });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "4.6 + 2.3 = 6.9", {
        x: 1.0, y: 4.55, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 12-13: We Do + reveal — add 5.27 + 1.4
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Add 5.27 + 1.4",
      [
        "With your partner.",
        "",
        "1.  ESTIMATE first.",
        "2.  Write the PLACEHOLDER ZERO.",
        "3.  LINE UP and ADD.",
        "4.  CHECK against your estimate.",
        "",
        "Watch: which number needs the placeholder zero?",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY });
        slide.addText("Build the calculation", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        addStackedAddition(slide, lg.rightX + 0.30, lg.panelTopPadded + 0.55,
          lg.rightW - 0.60, "5.27", "1.4", null,
          { color: C.SECONDARY, fontSize: 34 });
        slide.addText("(Add the placeholder zero!)", {
          x: lg.rightX + 0.20, y: lg.panelTopPadded + 2.50, w: lg.rightW - 0.40, h: 0.35,
          fontSize: 14, fontFace: FONT_B, color: C.MUTED, italic: true,
          align: "center", valign: "middle", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "5.27 + 1.40 = 6.67   (placeholder zero on 1.4 → 1.40)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 14-15: CFU hinge + reveal
  withReveal(
    () => cfuSlide(pres, "CFU", "Reasonable or not?", "Thumbs Up or Thumbs Down",
      "A student wrote:\n\n2.3 + 0.45 = 6.8\n\nIs this answer reasonable?\nUse estimation.",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "NOT reasonable.   2 + 0 ≈ 2.   The student added 23 + 45 = 68 with wrong alignment. Correct: 2.75.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 16: You Do — practice sheet
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: practice sheet", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.30, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 1 — same-length decimals.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 2 — different-length (placeholder zero).   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 3 — word problem.", options: { fontSize: 18, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.45;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    s.addText("Process — every time", {
      x: 0.7, y: panelY + 0.15, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });

    addTextOnShape(s, "1.  Estimate    2.  Line up the decimal point    3.  Add columns    4.  Check", {
      x: 0.8, y: panelY + 0.65, w: 8.4, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    addTextOnShape(s, "If lengths are different → add a PLACEHOLDER ZERO.", {
      x: 1.0, y: panelY + 1.20, w: 8.0, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "Add 6.3 + 2.85. Show your working including the placeholder zero.",
      "A student wrote 4.2 + 3.15 = 7.17. Is this correct? If not, fix it and explain the mistake.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 18: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: why do we write a placeholder zero?",
      scItems: [
        "I can use a place value chart to add two decimals.",
        "I can add two decimals (with different decimal lengths) by lining up the decimal point.",
        "I can use estimation and partitioning to check my answer makes sense.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecFrac_Lesson3_Adding_Decimals.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Add decimals using place value. Estimate, line up the decimal point, then add.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "Every time: estimate first, line up the decimal point, add column by column, then check against your estimate. Different lengths? Add a PLACEHOLDER ZERO.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 — Same-length decimals", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   3.4 + 2.5   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "b)   5.7 + 1.6   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "c)   4.25 + 1.43 =  _______      Estimate: ___", y);

    y = addSectionHeading(doc, "Section 2 — Different lengths (placeholder zero)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   5.27 + 1.4   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "b)   3.6 + 0.45   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "c)   2.1 + 4.305  =  _______      Estimate: ___", y);

    y = addSectionHeading(doc, "Section 3 — Word problem", y, { color: C.PRIMARY });
    y = addBodyText(doc, "At the school sports carnival, Mia jumped 2.45 metres in the long jump. Ben jumped 0.6 of a metre further than Mia.", y);
    y = addWriteLine(doc, "How far did Ben jump? Answer: ______ m", y);
    y = addWriteLine(doc, "What was their combined distance? Answer: ______ m", y);

    y = addSectionHeading(doc, "Section 4 — Extension (optional, mental partitioning)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Solve mentally by partitioning. Show your partition step.", y);
    y = addWriteLine(doc, "a)   4.7 + 2.6     Partition: ____________   Answer: ______", y);
    y = addWriteLine(doc, "b)   3.85 + 1.5    Partition: ____________   Answer: ______", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Adding Decimals Practice | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the adding decimals practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 — Same-length decimals", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  5.9  (est ~6)     b)  7.3  (est ~8)     c)  5.68  (est ~6)", y);

    y = addSectionHeading(doc, "Section 2 — Different lengths", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  6.67  (5.27 + 1.40, est ~6)", y);
    y = addBodyText(doc, "b)  4.05  (3.60 + 0.45, est ~4)", y);
    y = addBodyText(doc, "c)  6.405 (2.100 + 4.305, est ~6)", y);

    y = addSectionHeading(doc, "Section 3 — Word problem", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Ben's jump: 2.45 + 0.60 = 3.05 m.", y);
    y = addBodyText(doc, "Combined: 2.45 + 3.05 = 5.50 m (or 5.5 m).", y);

    y = addSectionHeading(doc, "Section 4 — Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "a)  4.7 + 2.6 = 7.3.   Partition example: 4 + 2 = 6, then 0.7 + 0.6 = 1.3, then 6 + 1.3 = 7.3.", y);
    y = addBodyText(doc, "b)  3.85 + 1.5 = 5.35.  Partition example: 3 + 1 = 4, then 0.85 + 0.5 = 1.35, then 4 + 1.35 = 5.35.", y);

    y = addTipBox(doc,
      "Watch for: students who add 5.27 + 1.4 and get 5.67 - they did not write the placeholder zero. Students who get 6.4 from 5.27 + 1.4 are dropping the 2 hundredths digit. Both errors indicate the placeholder-zero move is not yet automatic.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 3 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
