"use strict";

// Adding & Subtracting Decimals Unit (Year 5/6 Numeracy) — Lesson 3: Different lengths + placeholder zero
// First-time learning. Adding decimals where one number has more decimal places than the other.
// Daily Review: 2D shape properties (sides and vertices).
// Fluency: multiplication facts (x4).

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
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 3;
const TOTAL = 5;
const UNIT_TITLE = "Adding and Subtracting Decimals";
const FOOTER = `Add & Subtract Decimals | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/DecAdd_Lesson3_Adding_Decimals_Different_Lengths";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 3 Placeholder Zero Practice",
  "Adding decimals of different lengths using the placeholder zero strategy.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 3 Answer Key",
  "Worked answers for the placeholder zero practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Yesterday we added decimals that were the same length.
- Today we tackle decimals with different lengths.
- We meet a tiny trick called the placeholder zero - it keeps the place values lined up.

DO:
- Whiteboards, markers, printed place value reference cards ready.

TEACHER NOTES:
Lesson 3 of 5. The placeholder zero is the threshold concept of the unit. Without it, students align right-most digits and get wildly wrong answers. Teach it slowly today.

WATCH FOR:
- Students still dropping the decimal point - quick recap in I Do 1.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The placeholder zero practice sheet is for the You Do section.

DO:
- Print one copy of the practice sheet and answer key per student.

TEACHER NOTES:
One student resource (practice sheet) plus answer key.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Today we look at 2D shape properties.
- Look at each shape on the slide.
- On your whiteboard write the number of sides and the number of vertices.

DO:
- Display the three shapes.
- 90 seconds.

TEACHER NOTES:
Builds 2D shape language. Vertex = corner. Side = straight line edge.

WATCH FOR:
- Students who confuse sides and vertices - clarify with one shape on the board.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check.
- Pentagon: 5 sides, 5 vertices.
- Hexagon: 6 sides, 6 vertices.
- Octagon: 8 sides, 8 vertices.
- A regular shape has the same number of sides and vertices.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Tick and fix. Reinforce: sides and vertices always match in number for a simple polygon.

WATCH FOR:
- Students who self-correct - secure.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Four times tables today.
- Whisper-answer each one, then write it on your board.

DO:
- Display the three prompts.
- 45 seconds.

TEACHER NOTES:
Fluency continues building multiplication recall. x4 facts today.

WATCH FOR:
- Students who answer instantly - secure.
- Students who double-double (x2 then x2 again) - that is a valid strategy, prompt them to commit the fact.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check.
- 4 times 6 is 24.
- 4 times 8 is 32.
- 4 times 7 is 28.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Tick and fix. Note any student who still works it out vs recalls instantly.

WATCH FOR:
- Students who self-correct - secure.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to add decimals of different lengths using the placeholder zero.
- Now the success criteria.

DO:
- Choral read.

TEACHER NOTES:
SC1 is achievable - writing the placeholder zero so both decimals have the same number of decimal places. SC2 is the core target (addition). SC3 stretches to error analysis.

WATCH FOR:
- Students who can repeat "placeholder zero" - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me add 3.4 and 2.15.
- Estimate first: 3 + 2 = 5. So my answer should be around 5.
- These two numbers have DIFFERENT lengths. 3.4 has 1 decimal place. 2.15 has 2 decimal places.
- I write a PLACEHOLDER ZERO so both have 2 decimal places.
- 3.4 becomes 3.40. The value has not changed - 3.40 and 3.4 are the same.
- Now they line up neatly.
- Add column by column from the right.
- Hundredths: 0 + 5 = 5.
- Tenths:     4 + 1 = 5.
- Ones:       3 + 2 = 5.
- 5.55. Check: close to 5. Yes, reasonable.

DO:
- Display both numbers stacked.
- Underline the placeholder zero in a different colour.
- Repeat "the value has not changed" three times.

TEACHER NOTES:
The placeholder zero is the move that separates secure from shaky. Some texts call this "make the decimals the same length." Spend time on "the value has not changed" - that is the conceptual link.

MISCONCEPTIONS:
- Misconception: 3.4 + 2.15 = 3.55 (right-align the digits).
  Why: Students copy whole-number column addition habits.
  Impact: Wildly wrong answers.
  Quick correction: "Write the placeholder zero so both numbers have the same number of decimal places. Then line up."

WATCH FOR:
- Students who write 3.40 with the placeholder zero - secure.
- Students who line up the right-hand digit - the trap.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Another one. 4.6 + 0.07.
- Estimate: 4 + 0 is about 4. Maybe a tiny bit more.
- Lengths are different. 4.6 has 1 decimal place. 0.07 has 2.
- Add the placeholder zero on 4.6: it becomes 4.60.
- Stack 4.60 above 0.07 with the dots aligned.
- Hundredths: 0 + 7 = 7.
- Tenths:     6 + 0 = 6.
- Ones:       4 + 0 = 4.
- 4.67. Check against estimate of about 4. Yes, reasonable.

DO:
- Display 4.60 + 0.07 stacked.
- Highlight the placeholder zero.
- Trace down through the dot.

TEACHER NOTES:
This second model uses a small second number on purpose. It surfaces the "the small number adds a little" intuition that supports estimation.

MISCONCEPTIONS:
- Misconception: 4.6 + 0.07 = 4.13 (right-align).
  Why: Same as before.
  Impact: Wrong answer because tenths added to hundredths.
  Quick correction: "Add the placeholder zero. Then line up. 4.6 = 4.60."

WATCH FOR:
- Students who write 4.67 - secure.
- Students who write 4.13 - misaligned.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. Add 2.5 + 1.32.
- On your whiteboard. Write the placeholder zero first.
- Show me your answer.

DO:
- Display the prompt.
- 45 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your answer on three, two, one, show."
- Scan for: 3.82.
PROCEED: If 80% have 3.82, reveal.
PIVOT: Most likely misconception - students wrote 2.55 (line up right) or 1.57 (mis-add tenths).
- Reteach: "2.5 = 2.50. Then add hundredths first: 0 + 2 = 2."
- Re-check: "Add 2.5 + 1.32 again with the placeholder zero."

TEACHER NOTES:
The CFU has different lengths. Watch for the placeholder zero in students' working.

WATCH FOR:
- Students who write 3.82 - secure.
- Students who write 2.57 or similar - misaligned.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Add 5.27 + 1.4.
- First: ESTIMATE. About 5 + 1 = 6.
- Then: write the PLACEHOLDER ZERO where it is needed.
- Stack and add column by column from the right.
- Check against your estimate.

DO:
- Display the two numbers.
- 90 seconds.
- Walk and listen for "placeholder zero."

TEACHER NOTES:
The bigger number is the one with two decimal places. Students need to recognise that 1.4 is the one that needs the placeholder.

WATCH FOR:
- Students who add a placeholder zero to 1.4 - secure.
- Students who write 5.67 (didn't align) - reteach.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- 5.27 + 1.40 (placeholder zero on 1.4).
- Hundredths: 7 + 0 = 7.
- Tenths:     2 + 4 = 6.
- Ones:       5 + 1 = 6.
- 6.67. Estimate was 6 - yes, reasonable.

DO:
- Click to reveal.
- Restate the placeholder zero move.

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
- Reteach: "Estimate first. 2 + 0 is about 2. The student added 23 + 45 because they lined up the right-hand digit."
- Re-check: "What is the real answer?"

TEACHER NOTES:
This hinge probes whether students use estimation as a reasonableness check before doing the calculation.

WATCH FOR:
- Confident thumbs down - they used estimation.
- Students who hesitate - reteach the estimate-first habit.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Section 1: short + long decimals.
- Section 2: word problem.
- Section 3: optional challenge.

DO:
- Distribute the practice sheet.
- Circulate.
- Cold call 1-2 students to share thinking on the word problem.

TEACHER NOTES:
The practice sheet focuses on the placeholder zero across two sections.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use Section 1 only with the printed place value chart.
- Extra Notes: Sit with these students. Build the placeholder zero with a coloured pencil on the first one together.
EXTENDING PROMPT:
- Task: Section 3 - add three decimals of different lengths.
- Extra Notes: Encourage stacking all three with the dots aligned.

WATCH FOR:
- Students who write the placeholder zero consistently - secure.
- Students who line up the right-hand digit - reteach.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task.
- Question 1: 6.3 + 2.85. Show your placeholder zero step.
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
- Students who say "to make the lengths the same" or "to line up the place values" - secure.
- Students unsure - revisit at the start of Lesson 4.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Helpers ────────────────────────────────────────────────────────────────

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

  titleSlide(pres, UNIT_TITLE, "Lesson 3: Different lengths - the placeholder zero",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Daily Review + reveal — 2D shape properties (sides and vertices)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Daily Review: Sides and vertices");

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });

      const shapeY = CONTENT_TOP + 0.45;

      // Pentagon
      s.addShape("pentagon", {
        x: 1.3, y: shapeY, w: 1.6, h: 1.5,
        fill: { color: C.PRIMARY }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("Pentagon", {
        x: 1.0, y: shapeY + 1.60, w: 2.2, h: 0.35,
        fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });
      s.addText("Sides: ___   Vertices: ___", {
        x: 1.0, y: shapeY + 1.95, w: 2.2, h: 0.30,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });

      // Hexagon
      s.addShape("hexagon", {
        x: 4.0, y: shapeY, w: 1.9, h: 1.5,
        fill: { color: C.SECONDARY }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("Hexagon", {
        x: 3.7, y: shapeY + 1.60, w: 2.5, h: 0.35,
        fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });
      s.addText("Sides: ___   Vertices: ___", {
        x: 3.7, y: shapeY + 1.95, w: 2.5, h: 0.30,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });

      // Octagon
      s.addShape("octagon", {
        x: 7.0, y: shapeY, w: 1.7, h: 1.5,
        fill: { color: C.ACCENT }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("Octagon", {
        x: 6.6, y: shapeY + 1.60, w: 2.5, h: 0.35,
        fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });
      s.addText("Sides: ___   Vertices: ___", {
        x: 6.6, y: shapeY + 1.95, w: 2.5, h: 0.30,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Pentagon: 5 sides, 5 vertices   |   Hexagon: 6 / 6   |   Octagon: 8 / 8", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Fluency + reveal: x4 facts
  withReveal(
    () => fluencySlide(pres, "Fluency: Multiplication (x4)",
      ["4 × 6", "4 × 8", "4 × 7"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "24     32     28", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // LI/SC
  liSlide(pres,
    "We are learning to add decimals of different lengths by writing a placeholder zero.",
    [
      "I can write a placeholder zero so two decimals have the same number of decimal places.",
      "I can add two decimals (different lengths) by lining up the decimal point.",
      "I can spot when a student's answer is unreasonable and explain why.",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do 1: 3.4 + 2.15
  workedExSlide(pres, 2, "I Do", "Add 3.4 + 2.15",
    [
      "1.  ESTIMATE: 3 + 2 ≈ 5.",
      "",
      "2.  Different lengths!",
      "     Write the PLACEHOLDER ZERO:",
      "     3.4 becomes 3.40.",
      "",
      "3.  LINE UP the decimal points.",
      "",
      "4.  ADD right to left.",
      "     Hundredths: 0 + 5 = 5.",
      "     Tenths:     4 + 1 = 5.",
      "     Ones:       3 + 2 = 5.",
      "",
      "5.  5.55 ≈ 5.  ✓",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("Placeholder zero in red", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      slide.addText("Ones    Tenths    Hundredths", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.20, h: 0.26,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });

      const stackX = lg.rightX + 0.30;
      const stackW = lg.rightW - 0.60;
      const rowH = 0.50;
      const rowY = lg.panelTopPadded + 0.78;

      // Top: "3.40" with the placeholder 0 in red
      slide.addText([
        { text: "3.4",
          options: { fontSize: 36, fontFace: FONT_H, color: C.CHARCOAL, bold: true } },
        { text: "0",
          options: { fontSize: 36, fontFace: FONT_H, color: C.ALERT, bold: true } },
      ], {
        x: stackX, y: rowY, w: stackW, h: rowH,
        align: "center", valign: "middle", margin: 0,
      });

      slide.addText("+   2.15", {
        x: stackX, y: rowY + rowH, w: stackW, h: rowH,
        fontSize: 36, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      slide.addShape("line", {
        x: stackX + 0.4, y: rowY + 2 * rowH + 0.04, w: stackW - 0.8, h: 0,
        line: { color: C.PRIMARY, width: 2 },
      });

      slide.addText("5.55", {
        x: stackX, y: rowY + 2 * rowH + 0.12, w: stackW, h: rowH,
        fontSize: 36, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // I Do 2: 4.6 + 0.07
  workedExSlide(pres, 2, "I Do", "Add 4.6 + 0.07",
    [
      "1.  ESTIMATE: about 4.",
      "",
      "2.  Different lengths!",
      "     4.6 becomes 4.60.",
      "",
      "3.  LINE UP and ADD.",
      "     Hundredths: 0 + 7 = 7.",
      "     Tenths:     6 + 0 = 6.",
      "     Ones:       4 + 0 = 4.",
      "",
      "4.  4.67. Just over 4.  ✓",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY });
      slide.addText("4.6 = 4.60 (same value)", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      const stackX = lg.rightX + 0.30;
      const stackW = lg.rightW - 0.60;
      const rowH = 0.50;
      const rowY = lg.panelTopPadded + 0.55;

      slide.addText([
        { text: "4.6",
          options: { fontSize: 36, fontFace: FONT_H, color: C.CHARCOAL, bold: true } },
        { text: "0",
          options: { fontSize: 36, fontFace: FONT_H, color: C.ALERT, bold: true } },
      ], {
        x: stackX, y: rowY, w: stackW, h: rowH,
        align: "center", valign: "middle", margin: 0,
      });

      slide.addText("+   0.07", {
        x: stackX, y: rowY + rowH, w: stackW, h: rowH,
        fontSize: 36, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      slide.addShape("line", {
        x: stackX + 0.4, y: rowY + 2 * rowH + 0.04, w: stackW - 0.8, h: 0,
        line: { color: C.SECONDARY, width: 2 },
      });

      slide.addText("4.67", {
        x: stackX, y: rowY + 2 * rowH + 0.12, w: stackW, h: rowH,
        fontSize: 36, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // CFU + reveal: 2.5 + 1.32
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Add 2.5 + 1.32", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 20, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Add 2.5 + 1.32", options: { fontSize: 28, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "1.  Estimate.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "2.  Add the PLACEHOLDER ZERO.", options: { fontSize: 16, color: C.ALERT, bold: true, breakLine: true } },
        { text: "3.  Line up and add.", options: { fontSize: 16, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.3, CONTENT_TOP, 4.2, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      s.addText("Stack it carefully", {
        x: 5.3, y: CONTENT_TOP + 0.10, w: 4.2, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      addStackedAddition(s, 5.6, CONTENT_TOP + 0.70, 3.6, "2.5", "1.32", null,
        { color: C.PRIMARY, fontSize: 38 });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "2.50 + 1.32 = 3.82   (placeholder zero on 2.5 → 2.50)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // We Do + reveal: 5.27 + 1.4
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
        "Which number needs the placeholder?",
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

  // Hinge + reveal
  withReveal(
    () => cfuSlide(pres, "CFU", "Reasonable or not?", "Thumbs Up or Thumbs Down",
      "A student wrote:\n\n2.3 + 0.45 = 6.8\n\nIs this answer reasonable?\nUse estimation.",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "NOT reasonable.   2 + 0 ≈ 2.   The student added 23 + 45 = 68 (wrong alignment).   Correct: 2.75.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // You Do
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: practice sheet", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.30, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 1 — short + long decimals.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 2 — word problem.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 3 — challenge.", options: { fontSize: 18, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.45;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    s.addText("The placeholder zero rule", {
      x: 0.7, y: panelY + 0.15, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });

    addTextOnShape(s, "If lengths are different → add a PLACEHOLDER ZERO.  Then line up the dot.", {
      x: 0.8, y: panelY + 0.65, w: 8.4, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    addTextOnShape(s, "The value does not change.   3.4 = 3.40   |   4.6 = 4.60", {
      x: 1.0, y: panelY + 1.20, w: 8.0, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Exit ticket
  exitTicketSlide(pres,
    [
      "Add 6.3 + 2.85. Show your working including the placeholder zero.",
      "A student wrote 4.2 + 3.15 = 7.17. Is this correct? If not, fix it and explain the mistake.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: why do we write a placeholder zero?",
      scItems: [
        "I can write a placeholder zero so two decimals have the same number of decimal places.",
        "I can add two decimals (different lengths) by lining up the decimal point.",
        "I can spot when a student's answer is unreasonable and explain why.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecAdd_Lesson3_Adding_Decimals_Different_Lengths.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // PDFs
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Add decimals of different lengths using the placeholder zero.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "Different lengths? Write the placeholder zero so both numbers have the same number of decimal places. Then line up the dot and add column by column.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 — Placeholder zero practice", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   3.4 + 2.15   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "b)   5.27 + 1.4   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "c)   2.5 + 1.32   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "d)   4.6 + 0.07   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "e)   0.8 + 0.45   =   _______      Estimate: ___", y);

    y = addSectionHeading(doc, "Section 2 — Word problem", y, { color: C.PRIMARY });
    y = addBodyText(doc, "At the school sports carnival, Mia jumped 2.45 metres in the long jump. Ben jumped 0.6 of a metre further than Mia.", y);
    y = addWriteLine(doc, "How far did Ben jump? Answer: ______ m", y);
    y = addWriteLine(doc, "What was their combined distance? Answer: ______ m", y);

    y = addSectionHeading(doc, "Section 3 — Challenge (three decimals)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Stack all three numbers with the dots aligned. Add the placeholder zero where needed.", y);
    y = addWriteLine(doc, "a)   1.4 + 0.25 + 2.1     =   _______", y);
    y = addWriteLine(doc, "b)   3.05 + 0.6 + 1.27    =   _______", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Placeholder Zero Practice | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the placeholder zero practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 — Placeholder zero practice", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  5.55  (3.40 + 2.15, est ~5)", y);
    y = addBodyText(doc, "b)  6.67  (5.27 + 1.40, est ~6)", y);
    y = addBodyText(doc, "c)  3.82  (2.50 + 1.32, est ~4)", y);
    y = addBodyText(doc, "d)  4.67  (4.60 + 0.07, est ~4-5)", y);
    y = addBodyText(doc, "e)  1.25  (0.80 + 0.45, est ~1)", y);

    y = addSectionHeading(doc, "Section 2 — Word problem", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Ben's jump: 2.45 + 0.60 = 3.05 m.", y);
    y = addBodyText(doc, "Combined: 2.45 + 3.05 = 5.50 m (or 5.5 m).", y);

    y = addSectionHeading(doc, "Section 3 — Challenge", y, { color: C.ACCENT });
    y = addBodyText(doc, "a)  1.40 + 0.25 + 2.10 = 3.75.", y);
    y = addBodyText(doc, "b)  3.05 + 0.60 + 1.27 = 4.92.", y);

    y = addTipBox(doc,
      "Watch for: students who add 5.27 + 1.4 and get 5.67 (no placeholder zero); students who add 4.6 + 0.07 and get 4.13 (right-align misalignment). Both errors mean the placeholder-zero move is not yet automatic.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 3 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
