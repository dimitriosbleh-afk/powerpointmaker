"use strict";

// Adding & Subtracting Decimals Unit (Year 5/6 Numeracy) — Lesson 2: Adding decimals (same length)
// First-time learning. Tenths + tenths, hundredths + hundredths. "Line up the decimal point."
// Daily Review: identifying 3D shapes.
// Fluency: multiplication facts (x3).

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(3));
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

const SESSION = 2;
const TOTAL = 5;
const UNIT_TITLE = "Adding and Subtracting Decimals";
const FOOTER = `Add & Subtract Decimals | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/DecAdd_Lesson2_Adding_Decimals_Same_Length";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 2 Adding Decimals Practice",
  "Adding tenths to tenths and hundredths to hundredths by lining up the decimal point.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 2 Answer Key",
  "Worked answers for the adding decimals practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Yesterday we learned how decimals are built from tenths and hundredths.
- Today we use that to add decimals together.
- The one big rule we will say all lesson: line up the decimal point.

DO:
- Whiteboards, markers, and printed place value reference cards ready.

TEACHER NOTES:
Lesson 2 of 5. Today is the first time these students add decimals explicitly. Both numbers have the same number of decimal places, so the "line up the decimal point" rule works exactly like whole-number column addition. Lesson 3 adds the placeholder zero for different lengths.

WATCH FOR:
- Students still wobbly on tenths vs hundredths from Lesson 1 - reuse the grid quickly during I Do.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The adding decimals practice sheet is for the You Do section.

DO:
- Print one copy of the practice sheet and answer key per student.
- Whiteboards, markers, and printed place value reference cards ready.

TEACHER NOTES:
One student resource (practice sheet) plus answer key.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Today we are naming 3D shapes.
- Whisper the name to your partner, then write it on your whiteboard.

DO:
- Display the three shapes.
- 60 seconds.

TEACHER NOTES:
3D shape names this week. Listen for "cube", "cylinder", "cone". These are the three most common.

WATCH FOR:
- Students who confidently name them - secure.
- Students who say "square" for cube - prompt: "is it 2D or 3D?"

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check.
- A = cube.   6 square faces.
- B = cylinder.   2 circle faces and a curved face.
- C = cone.   1 circle face and a curved face to a point.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Quick tick-and-fix. Use later this week to bridge into 2D faces of 3D shapes.

WATCH FOR:
- Students who self-correct - secure.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Three times tables today.
- Whisper-answer each one, then write it on your board.

DO:
- Display the three prompts.
- 45 seconds.

TEACHER NOTES:
Fluency continues building multiplication recall. x3 facts today.

WATCH FOR:
- Students who answer instantly - secure.
- Students who count up in 3s - flag for fact recall practice.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 3 times 4 is 12.
- 3 times 8 is 24.
- 3 times 6 is 18.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Tick and fix. Look for students who still count up rather than recall.

WATCH FOR:
- Students who self-correct - secure.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to add decimals by lining up the decimal point.
- Now the success criteria.

DO:
- Choral read.
- Hold up the printed place value chart.

TEACHER NOTES:
SC1 is achievable - estimating the answer. SC2 is the core target (the addition). SC3 stretches to using estimation to check.

WATCH FOR:
- Students who can repeat the line-up rule - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me add 0.3 and 0.4.
- These are both tenths, so I add tenths to tenths.
- Step 1: Estimate. 0.3 is a little less than half. 0.4 is a little less than half. So my answer should be just under 1.
- Step 2: Line up the decimal points so the tenths sit under the tenths.
- Step 3: Add the tenths column. 3 + 4 = 7.
- Step 4: Bring the decimal point straight down.
- Step 5: 0.7. Check against my estimate of just under 1. Yes, reasonable.

DO:
- Display the worked example stacked.
- Trace down through the decimal point with a finger.
- Repeat "tenths plus tenths" three times.

TEACHER NOTES:
The first I Do is intentionally simple - both tenths, no carrying. The aim is to install the line-up habit before any complication. Use the grid model from Lesson 1 in your think-aloud if students need the size to make sense.

MISCONCEPTIONS:
- Misconception: Students drop the decimal point and write 7.
  Why: They treat addition of decimals like whole numbers without preserving place value.
  Impact: Their answer changes by a factor of 10.
  Quick correction: "Bring the decimal point straight down. It marks the place."

WATCH FOR:
- Students who say "tenths plus tenths" with you - tracking.
- Students who answer 7 instead of 0.7 - the dropped-dot trap.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now I add 1.25 and 2.34.
- Both have hundredths, so I add hundredths to hundredths.
- Estimate first: 1 + 2 = 3. So my answer should be around 3.
- Line up the decimal points.
- Add column by column, from the right.
- Hundredths: 5 + 4 = 9.
- Tenths:     2 + 3 = 5.
- Ones:       1 + 2 = 3.
- Bring the decimal point straight down.
- 3.59. Check: 3.59 is close to 3. Yes, reasonable.

DO:
- Display both numbers stacked, with the decimal point aligned.
- Point to each column as you add.
- Repeat "right to left, line up the dot."

TEACHER NOTES:
Two-decimal-place addition. Same line-up rule. We work from the right to keep it consistent with whole-number column addition.

MISCONCEPTIONS:
- Misconception: Students add from left to right.
  Why: They forget the carry rule.
  Impact: They miss any regrouping.
  Quick correction: "Always start from the right - the smallest place first."

WATCH FOR:
- Students who add column by column starting from the right - secure.
- Students who try left to right - reteach the right-to-left rule.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. Add 0.6 + 0.3.
- On your whiteboard. Show the line-up.
- Show me your answer.

DO:
- Display the prompt.
- 30 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your answer on three, two, one, show."
- Scan for: 0.9.
PROCEED: If 80% have 0.9, click to reveal and move on.
PIVOT: Most likely misconception - students wrote 9 (dropped the dot) or 0.09 (placed it wrongly).
- Reteach: "Tenths plus tenths is tenths. 6 tenths + 3 tenths = 9 tenths = 0.9."
- Re-check: "Add 0.6 + 0.3 again."

TEACHER NOTES:
Tenths-only addition. The dropped-dot or misplaced-dot errors are the ones to scan for.

WATCH FOR:
- Students who write 0.9 - secure.
- Students who write 9 - dropped dot.
- Students who write 0.09 - misplaced dot.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Add 2.36 and 1.42.
- First: estimate. About 2 + 1 = 3.
- Then: write them stacked with the decimal points aligned.
- Add column by column from the right.
- Check against your estimate.

DO:
- Display the two numbers.
- 90 seconds.
- Walk and listen for "line up the dot" and "start from the right."

TEACHER NOTES:
Two-decimal-place addition with no regrouping. The aim is fluency on the procedure.

WATCH FOR:
- Students who align and add cleanly - secure.
- Students who miscount columns - reteach with the chart.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- Hundredths: 6 + 2 = 8.
- Tenths:     3 + 4 = 7.
- Ones:       2 + 1 = 3.
- Answer: 3.78. Estimate was 3 - yes, reasonable.

DO:
- Click to reveal.
- Run through the columns once more.

TEACHER NOTES:
The reveal confirms the line-up procedure. Tick and fix.

WATCH FOR:
- Students who self-correct - secure.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. Look at this student's work.
- 0.4 + 0.5 = 9.
- Is this reasonable? Thumbs UP for yes, DOWN for no.

DO:
- Display the work.
- 15 seconds.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Thumbs UP if reasonable. DOWN if not."
- Scan for: thumbs DOWN.
PROCEED: If 80% thumbs down, click to reveal.
PIVOT: Most likely misconception - students see 4 + 5 = 9 and accept it.
- Reteach: "0.4 is less than 1. 0.5 is less than 1. The answer must be less than 2, not 9. The student dropped the decimal point."
- Re-check: "What is the real answer?"

TEACHER NOTES:
This hinge probes whether students use the size of the numbers to spot an unreasonable answer.

WATCH FOR:
- Confident thumbs down - they used estimation.
- Students who hesitate - reteach the estimate-first habit.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Take the practice sheet.
- Section 1: tenths only.
- Section 2: hundredths.
- Section 3: word problem.

DO:
- Distribute the practice sheet.
- Circulate.
- Cold call 1-2 students to share Section 3 thinking.

TEACHER NOTES:
The practice sheet stays at same-length decimals. The placeholder zero is a Lesson 3 move.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use only Section 1 with the printed place value chart.
- Extra Notes: Sit with these students. Do the first one together.
EXTENDING PROMPT:
- Task: Section 4 - addition chains. Solve 1.23 + 0.45 + 0.31.
- Extra Notes: Encourage stacking all three with the decimal points aligned.

WATCH FOR:
- Students who line up the dot consistently - secure.
- Students who write 0.9 vs 9 errors - the dropped-dot trap.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task.
- Question 1: 2.36 + 4.51. Show your line-up and the answer.
- Question 2: A student wrote 0.7 + 0.2 = 9. Is this correct? If not, fix it and explain.

DO:
- Display the prompt.
- 3 minutes.
- Collect whiteboards.

TEACHER NOTES:
Exit ticket assesses SC2. Q2 probes whether students notice the dropped-dot error.

WATCH FOR:
- Students who write 6.87 for Q1 - secure on hundredths addition.
- Students who fix Q2 to 0.9 with an explanation - secure on size reasoning.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria. Thumbs for each.
- Turn and tell your partner: what is the one big rule for adding decimals?

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
The threshold idea is "line up the decimal point." Students who can say this with the place-value reason are ready for Lesson 3.

WATCH FOR:
- Students who say "line up the decimal point" - secure.
- Students unsure - revisit at the start of Lesson 3.

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

// Stacked addition with aligned decimal point.
function addStackedAddition(slide, x, y, w, top, bottom, answer, opts) {
  const o = opts || {};
  const lineColor = o.color || C.PRIMARY;
  const fontSize = o.fontSize || 28;
  const rowH = o.rowH || 0.46;

  slide.addText(top, {
    x, y, w, h: rowH,
    fontSize, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  slide.addText("+   " + bottom, {
    x, y: y + rowH, w, h: rowH,
    fontSize, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  slide.addShape("line", {
    x: x + 0.4, y: y + 2 * rowH + 0.04, w: w - 0.8, h: 0,
    line: { color: lineColor, width: 2 },
  });
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

  titleSlide(pres, UNIT_TITLE, "Lesson 2: Adding decimals - line up the dot",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Daily Review + reveal — naming 3D shapes
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Daily Review: Name each 3D shape");

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });

      const shapeY = CONTENT_TOP + 0.5;

      // A: Cube (drawn as a square with offset square to suggest 3D)
      {
        const x = 1.3, y = shapeY + 0.15, side = 1.4, off = 0.30;
        // back face
        s.addShape("rect", {
          x: x + off, y, w: side, h: side,
          fill: { color: C.SECONDARY }, line: { color: C.CHARCOAL, width: 1 },
        });
        // front face
        s.addShape("rect", {
          x, y: y + off, w: side, h: side,
          fill: { color: C.PRIMARY }, line: { color: C.CHARCOAL, width: 1 },
        });
        // top connecting lines
        s.addShape("line", {
          x, y: y + off, w: off, h: 0,
          line: { color: C.CHARCOAL, width: 1 },
        });
        s.addShape("line", {
          x: x + side, y: y + off, w: off, h: 0,
          line: { color: C.CHARCOAL, width: 1 },
        });
        s.addText("A", {
          x: 1.1, y: shapeY + 1.95, w: 1.8, h: 0.4,
          fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
          align: "center", margin: 0,
        });
      }

      // B: Cylinder (oval on top of rectangle with another oval at the bottom)
      {
        const cx = 4.0, cy = shapeY + 0.15, w = 1.7, h = 1.6;
        // body rectangle
        s.addShape("rect", {
          x: cx, y: cy + 0.25, w, h: h - 0.50,
          fill: { color: C.ACCENT }, line: { color: C.ACCENT, width: 0 },
        });
        // top oval
        s.addShape("ellipse", {
          x: cx, y: cy, w, h: 0.45,
          fill: { color: C.ACCENT }, line: { color: C.CHARCOAL, width: 1 },
        });
        // bottom oval
        s.addShape("ellipse", {
          x: cx, y: cy + h - 0.45, w, h: 0.45,
          fill: { color: C.ACCENT }, line: { color: C.CHARCOAL, width: 1 },
        });
        // side lines
        s.addShape("line", {
          x: cx, y: cy + 0.225, w: 0, h: h - 0.45,
          line: { color: C.CHARCOAL, width: 1 },
        });
        s.addShape("line", {
          x: cx + w, y: cy + 0.225, w: 0, h: h - 0.45,
          line: { color: C.CHARCOAL, width: 1 },
        });
        s.addText("B", {
          x: cx - 0.1, y: shapeY + 1.95, w: 1.9, h: 0.4,
          fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
          align: "center", margin: 0,
        });
      }

      // C: Cone (triangle on top of an ellipse)
      {
        const cx = 7.1, cy = shapeY, w = 1.7, h = 1.8;
        s.addShape("triangle", {
          x: cx, y: cy, w, h: h - 0.30,
          fill: { color: C.ALERT }, line: { color: C.CHARCOAL, width: 1 },
        });
        s.addShape("ellipse", {
          x: cx, y: cy + h - 0.55, w, h: 0.45,
          fill: { color: C.ALERT }, line: { color: C.CHARCOAL, width: 1 },
        });
        s.addText("C", {
          x: cx - 0.1, y: shapeY + 1.95, w: 1.9, h: 0.4,
          fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
          align: "center", margin: 0,
        });
      }

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
      addTextOnShape(slide, "A = cube     B = cylinder     C = cone", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Fluency + reveal — x3 facts
  withReveal(
    () => fluencySlide(pres, "Fluency: Multiplication (x3)",
      ["3 × 4", "3 × 8", "3 × 6"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "12     24     18", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // LI/SC
  liSlide(pres,
    "We are learning to add decimals by lining up the decimal point.",
    [
      "I can estimate the answer before adding decimals.",
      "I can add two decimals (same length) by lining up the decimal point and adding column by column.",
      "I can use my estimate to check the answer is reasonable.",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do 1: 0.3 + 0.4
  workedExSlide(pres, 2, "I Do", "Add 0.3 + 0.4",
    [
      "1.  ESTIMATE: just under 1.",
      "",
      "2.  LINE UP the decimal points.",
      "",
      "3.  ADD the tenths column.",
      "     3 + 4 = 7 tenths.",
      "",
      "4.  BRING DOWN the decimal point.",
      "",
      "5.  0.7. Check: just under 1.  ✓",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("Stacked addition", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      slide.addText("Ones    Tenths", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.20, h: 0.26,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });
      addStackedAddition(slide, lg.rightX + 0.30, lg.panelTopPadded + 0.78,
        lg.rightW - 0.60, "0.3", "0.4", "0.7",
        { color: C.PRIMARY, fontSize: 36, rowH: 0.50 });
    }
  );

  // I Do 2: 1.25 + 2.34
  workedExSlide(pres, 2, "I Do", "Add 1.25 + 2.34",
    [
      "1.  ESTIMATE: 1 + 2 ≈ 3.",
      "",
      "2.  LINE UP the decimal points.",
      "",
      "3.  ADD column by column from the right.",
      "     Hundredths: 5 + 4 = 9.",
      "     Tenths:     2 + 3 = 5.",
      "     Ones:       1 + 2 = 3.",
      "",
      "4.  BRING DOWN the decimal point.",
      "",
      "5.  3.59 ≈ 3.  ✓",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY });
      slide.addText("Right to left, line up the dot", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      slide.addText("Ones    Tenths    Hundredths", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.20, h: 0.26,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });
      addStackedAddition(slide, lg.rightX + 0.30, lg.panelTopPadded + 0.78,
        lg.rightW - 0.60, "1.25", "2.34", "3.59",
        { color: C.SECONDARY, fontSize: 34, rowH: 0.50 });
    }
  );

  // CFU + reveal: 0.6 + 0.3
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Add 0.6 + 0.3", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 20, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Add 0.6 + 0.3", options: { fontSize: 30, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "1.  Stack with the dot lined up.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "2.  Add the tenths.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "3.  Bring the dot straight down.", options: { fontSize: 16, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.3, CONTENT_TOP, 4.2, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      s.addText("Stacked addition", {
        x: 5.3, y: CONTENT_TOP + 0.10, w: 4.2, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      addStackedAddition(s, 5.6, CONTENT_TOP + 0.70, 3.6, "0.6", "0.3", null,
        { color: C.PRIMARY, fontSize: 38 });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "0.6 + 0.3 = 0.9", {
        x: 1.0, y: 4.55, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // We Do + reveal: 2.36 + 1.42
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Add 2.36 + 1.42",
      [
        "With your partner.",
        "",
        "1.  ESTIMATE first.",
        "2.  LINE UP the decimal points.",
        "3.  ADD column by column (right to left).",
        "4.  CHECK against your estimate.",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY });
        slide.addText("Build it together", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        addStackedAddition(slide, lg.rightX + 0.30, lg.panelTopPadded + 0.55,
          lg.rightW - 0.60, "2.36", "1.42", null,
          { color: C.SECONDARY, fontSize: 34 });
      }),
    (slide) => {
      addTextOnShape(slide, "2.36 + 1.42 = 3.78    (Hundredths 6+2, Tenths 3+4, Ones 2+1)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Hinge + reveal: 0.4 + 0.5 = 9?
  withReveal(
    () => cfuSlide(pres, "CFU", "Reasonable or not?", "Thumbs Up or Thumbs Down",
      "A student wrote:\n\n0.4 + 0.5 = 9\n\nIs this reasonable?\nUse estimation.",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "NOT reasonable.   0.4 and 0.5 are both less than 1.   Real answer: 0.9.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // You Do — practice sheet
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: practice sheet", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.30, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 1 — tenths only.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 2 — hundredths.   ", options: { fontSize: 18, color: C.CHARCOAL } },
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

    addTextOnShape(s, "1.  Estimate    2.  Line up the dot    3.  Add right to left    4.  Check", {
      x: 0.8, y: panelY + 0.65, w: 8.4, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    addTextOnShape(s, "Same length means same number of decimal places.", {
      x: 1.0, y: panelY + 1.20, w: 8.0, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Exit ticket
  exitTicketSlide(pres,
    [
      "Add 2.36 + 4.51. Show how you line up the decimal points.",
      "A student wrote 0.7 + 0.2 = 9. Is this correct? If not, fix it and explain.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what is the one big rule for adding decimals?",
      scItems: [
        "I can estimate the answer before adding decimals.",
        "I can add two decimals (same length) by lining up the decimal point and adding column by column.",
        "I can use my estimate to check the answer is reasonable.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecAdd_Lesson2_Adding_Decimals_Same_Length.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // PDFs
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Add decimals by lining up the decimal point.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "Every time: estimate first, line up the decimal point, add column by column from the right, then check against your estimate.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 — Tenths only", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   0.3 + 0.4   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "b)   0.6 + 0.2   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "c)   0.5 + 0.4   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "d)   1.4 + 2.3   =   _______      Estimate: ___", y);

    y = addSectionHeading(doc, "Section 2 — Hundredths", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   2.36 + 1.42   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "b)   0.45 + 0.32   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "c)   3.15 + 1.24   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "d)   4.07 + 2.61   =   _______      Estimate: ___", y);

    y = addSectionHeading(doc, "Section 3 — Word problem", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Mia runs 1.45 km in the morning and 2.32 km after school.", y);
    y = addWriteLine(doc, "How far does Mia run in total? Answer: ______ km", y);

    y = addSectionHeading(doc, "Section 4 — Extension (optional, addition chain)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Stack all three numbers with the decimal points aligned, then add.", y);
    y = addWriteLine(doc, "a)   1.23 + 0.45 + 0.31   =   _______", y);
    y = addWriteLine(doc, "b)   2.4 + 1.3 + 0.2     =   _______", y);

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

    y = addSectionHeading(doc, "Section 1 — Tenths only", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  0.7     b)  0.8     c)  0.9     d)  3.7", y);

    y = addSectionHeading(doc, "Section 2 — Hundredths", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3.78    b)  0.77    c)  4.39    d)  6.68", y);

    y = addSectionHeading(doc, "Section 3 — Word problem", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Total distance: 1.45 + 2.32 = 3.77 km.", y);

    y = addSectionHeading(doc, "Section 4 — Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "a)  1.23 + 0.45 + 0.31 = 1.99.", y);
    y = addBodyText(doc, "b)  2.4 + 1.3 + 0.2 = 3.9.", y);

    y = addTipBox(doc,
      "Watch for: students who write 0.9 as 9 (dropped the decimal point); students who answer 1.7 to 1.4 + 2.3 (forgot to carry from tenths to ones is NOT the issue here - watch for missing the ones column).",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 2 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
