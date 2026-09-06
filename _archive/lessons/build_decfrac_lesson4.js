"use strict";

// Decimals & Fractions Unit (Year 5/6 Numeracy) — Lesson 4: Subtracting decimals using place value
// VC2M6N02 — apply place value to subtract decimals; use the inverse relationship with addition.
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
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 4;
const TOTAL = 6;
const UNIT_TITLE = "Decimals and Fractions";
const FOOTER = `Decimals & Fractions | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/DecFrac_Lesson4_Subtracting_Decimals";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 4 Subtracting Decimals Practice",
  "Subtracting decimals using place value, regrouping, and the inverse check.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 4 Answer Key",
  "Worked answers for the subtracting decimals practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Today we subtract decimals using place value.
- Same line-up-the-dot move as yesterday.
- We also use the inverse — check by adding back.

DO:
- Have whiteboards, markers, place value reference cards ready.

TEACHER NOTES:
Lesson 4 of 6. Subtraction uses the same alignment habit as addition. The inverse-checking move (subtraction undone by addition) is the reasonableness check we will use throughout.

WATCH FOR:
- Students still wobbly on the placeholder zero from yesterday - reteach with a quick I Do recap.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The subtracting decimals practice sheet is for the You Do section.

DO:
- Print the practice sheet and answer key.
- Have whiteboards and place value reference cards ready.

TEACHER NOTES:
One student resource (practice sheet) plus answer key.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Warm up. Find the missing number.
- Use the inverse operation.

DO:
- Display the three equations.
- 90 seconds.

TEACHER NOTES:
Daily Review continues equation-solving. By Lesson 4 students should be quicker than Lesson 1.

WATCH FOR:
- Students confident with the inverse - secure.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check.
- 9 times n = 81. n is 9.
- n divided by 7 = 6. n is 42.
- 25 minus n = 14. n is 11.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
The third one uses subtraction with an unknown subtrahend. Students who got 39 added instead of subtracted.

WATCH FOR:
- Students who got 11 for the third - secure.
- Students who got 39 - reteach the inverse.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Adding and subtracting decimals.
- Line up the decimal point.

DO:
- Display the three prompts.
- 60 seconds.

TEACHER NOTES:
Fluency. Mixing addition and subtraction so the placeholder-zero habit transfers.

WATCH FOR:
- Students who line up correctly - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check.
- 0.8 plus 0.2 is 1.0.
- 1.6 minus 0.9 is 0.7.
- 3.45 minus 1.2 is 2.25 (placeholder zero on 1.2 makes it 1.20).

DO:
- Click to reveal.
- Run the placeholder-zero step once more.

TEACHER NOTES:
The third item is today's preview - subtraction with different decimal lengths.

WATCH FOR:
- Students who got 2.25 - secure.
- Students who got 3.25 - did not align.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to subtract decimals using place value and check by adding back.
- Now the success criteria.

DO:
- Choral read.

TEACHER NOTES:
SC1 floor - subtract same-length decimals. SC2 core target - subtract different-length decimals with regrouping. SC3 depth - check by inverse.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me subtract 5.8 minus 2.3.
- Step 1: estimate. 6 - 2 = 4. Answer should be around 4.
- Step 2: line up the decimal points.
- Step 3: subtract column by column starting from the right.
- Tenths: 8 - 3 = 5. Ones: 5 - 2 = 3.
- Step 4: bring down the decimal point. 3.5.
- Step 5: check. Estimate was 4. 3.5 is close. ✓
- Step 6: inverse check. 3.5 + 2.3 = 5.8. ✓

DO:
- Display the stacked subtraction.
- Point to each column.

TEACHER NOTES:
Same alignment rules as addition. The inverse-check move at the end gives students a built-in verification tool.

MISCONCEPTIONS:
- Misconception: Subtract the smaller from the larger in each column, regardless of position.
  Why: Whole-number trap repeated.
  Impact: 5.2 - 2.7 = 3.5 (wrongly subtracts 2 from 7).
  Quick correction: "Subtract top minus bottom. If you cannot, regroup."

WATCH FOR:
- Students who subtract column by column correctly - secure.
- Students who flip the order - reteach.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now I subtract with regrouping. 6.3 minus 2.85.
- Step 1: estimate. 6 - 3 = 3.
- Step 2: write the placeholder zero. 6.3 becomes 6.30.
- Step 3: line up.
- Step 4: subtract.
- Hundredths: 0 minus 5 - I cannot. So I regroup from the tenths.
- The 3 tenths becomes 2 tenths. The hundredths becomes 10. Now 10 - 5 = 5.
- Tenths: 2 - 8 - I cannot. Regroup from the ones. 6 becomes 5. Tenths becomes 12.
- 12 - 8 = 4.
- Ones: 5 - 2 = 3.
- Bring down the decimal point. 3.45.
- Step 5: estimate check (about 3) ✓ and inverse check (3.45 + 2.85 = 6.30 = 6.3) ✓.

DO:
- Display the subtraction with regrouping arrows shown.
- Move slowly through each column.

TEACHER NOTES:
Regrouping across the decimal point is the highest cognitive load step. Move slowly and verbalise each move. The inverse-add check at the end catches errors.

MISCONCEPTIONS:
- Misconception: Students forget the placeholder zero and try to subtract from nothing.
  Why: Different decimal lengths.
  Impact: They write incorrect answers or stop.
  Quick correction: "Add the placeholder zero so both numbers have the same number of decimal places."

WATCH FOR:
- Students who write the placeholder zero before subtracting - secure.
- Students stuck on regrouping - small group with MAB blocks.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. Subtract 7.6 minus 3.2.
- On your whiteboard. Show me.

DO:
- Display the prompt.
- 30 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your answer on three, two, one, show."
- Scan for: 4.4.
PROCEED: If 80% have 4.4, click to reveal and move on.
PIVOT: Most likely misconception - students subtracted the smaller from the larger.
- Reteach: "Top minus bottom. Same place column by column."
- Re-check: "Try again."

TEACHER NOTES:
Same-length subtraction, no regrouping. Quick check before we move to harder cases.

WATCH FOR:
- Students who hold up 4.4 - secure.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Subtract 4.2 minus 1.85.
- Estimate first: 4 - 2 = 2.
- Placeholder zero on 4.2 → 4.20.
- Line up. Subtract.
- Check by adding back.

DO:
- Display the prompt.
- 90 seconds.
- Walk and listen for "regroup."

TEACHER NOTES:
Different-length subtraction with regrouping. The crucial move is the placeholder zero, then regrouping from the tenths to the hundredths.

WATCH FOR:
- Pairs who write 4.20 first - secure on the placeholder zero.
- Pairs who write 4.20 and regroup - secure.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- 4.20 minus 1.85.
- Hundredths: 0 minus 5 - regroup. 2 tenths becomes 1 tenth. Hundredths is 10. 10 - 5 = 5.
- Tenths: 1 minus 8 - regroup. 4 ones becomes 3 ones. Tenths is 11. 11 - 8 = 3.
- Ones: 3 - 1 = 2.
- Answer: 2.35.
- Inverse check: 2.35 + 1.85 = 4.20 = 4.2. ✓

DO:
- Click to reveal.
- Run through the regrouping move.

TEACHER NOTES:
The reveal restates regrouping and inverse-checking.

WATCH FOR:
- Students who self-correct - secure.
- Students with wildly wrong answers - small group focus tomorrow.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. A student wrote: 5.3 - 2.18 = 3.18.
- Use the inverse to check. Add 3.18 + 2.18.
- Does it equal 5.3?
- Thumbs UP if correct. Thumbs DOWN if not.

DO:
- Display the work.
- 30 seconds.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Thumbs UP if correct. DOWN if not."
- Scan for: thumbs DOWN. 3.18 + 2.18 = 5.36, not 5.3.
PROCEED: If 80% thumbs down, click to reveal.
PIVOT: Most likely misconception - students cannot use inverse-add to check.
- Reteach: "If subtraction is right, adding back must give the original number."
- Re-check: "What does 3.12 + 2.18 give? Yes - 5.30 = 5.3."

TEACHER NOTES:
The hinge probes whether students use inverse-add as a reasonableness check. The correct answer is 5.30 - 2.18 = 3.12.

WATCH FOR:
- Confident thumbs down - they used inverse-add.
- Students who hesitate - reteach.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own. Take the practice sheet.
- Section 1: same-length subtraction.
- Section 2: different-length subtraction (placeholder zero).
- Section 3: word problem.

DO:
- Distribute the practice sheet.
- Circulate.
- Cold call 1-2 students to share Section 3.

TEACHER NOTES:
Section 3 is the word problem - applying subtraction in context.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the printed place value chart. Do only Section 1.
- Extra Notes: Sit with these students. Do the first one together.
EXTENDING PROMPT:
- Task: Section 4 - mental partitioning. Solve 5.2 - 2.7 by partitioning.
- Extra Notes: Encourage them to use jump strategies on the number line.

WATCH FOR:
- Students who use inverse-check - secure on the threshold.
- Students who skip the placeholder zero - reteach.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task.
- Question 1: 8.4 - 3.75. Show working and inverse check.
- Question 2: A student wrote 6.0 - 2.45 = 3.65. Use inverse-add to check. Is it correct? If not, fix it.

DO:
- Display.
- 3 minutes.
- Collect whiteboards.

TEACHER NOTES:
Exit ticket assesses SC2 (regrouping subtraction) and SC3 (inverse check).

WATCH FOR:
- Students who got 4.65 with inverse check - secure.
- Students who fixed Q2 to 3.55 - they used the inverse correctly.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria. Thumbs for each.
- Turn and tell your partner: why is inverse-check useful?

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
The inverse check is the lesson's threshold concept. Students who can say "adding back must give the original" understand it.

WATCH FOR:
- Students who can explain inverse-add - secure.
- Students unsure - revisit in Lesson 5.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Helpers ────────────────────────────────────────────────────────────────

function addStackedSubtraction(slide, x, y, w, top, bottom, answer, opts) {
  const o = opts || {};
  const lineColor = o.color || C.PRIMARY;
  const fontSize = o.fontSize || 28;
  const rowH = o.rowH || 0.46;

  slide.addText(top, {
    x, y, w, h: rowH,
    fontSize, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  slide.addText("−   " + bottom, {
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

  titleSlide(pres, UNIT_TITLE, "Lesson 4: Subtracting decimals using place value",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  withReveal(
    () => dailyReviewSlide(pres,
      "Daily Review: Find the missing number",
      ["9 × n = 81", "n ÷ 7 = 6", "25 − n = 14"],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "n = 9     n = 42     n = 11", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  withReveal(
    () => fluencySlide(pres, "Fluency: Adding & Subtracting Decimals",
      ["0.8 + 0.2", "1.6 − 0.9", "3.45 − 1.2"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "1.0     0.7     2.25", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  liSlide(pres,
    "We are learning to subtract decimals using place value and check our answer by adding back.",
    [
      "I can subtract two decimals when both have the same number of decimal places.",
      "I can subtract decimals with different lengths using a placeholder zero and regrouping.",
      "I can check my subtraction by using inverse-add (the answer plus the subtrahend equals the original).",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do 1 — 5.8 - 2.3
  workedExSlide(pres, 2, "I Do", "Subtract 5.8 − 2.3",
    [
      "1.  ESTIMATE: 6 − 2 ≈ 4.",
      "",
      "2.  LINE UP the decimal point.",
      "",
      "3.  SUBTRACT column by column.",
      "     Tenths: 8 − 3 = 5.",
      "     Ones:   5 − 2 = 3.",
      "",
      "4.  Answer: 3.5.",
      "",
      "5.  CHECK by inverse-add:",
      "     3.5 + 2.3 = 5.8  ✓",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("Stacked subtraction", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      slide.addText("Ones    Tenths", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.20, h: 0.26,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });
      addStackedSubtraction(slide, lg.rightX + 0.30, lg.panelTopPadded + 0.78,
        lg.rightW - 0.60, "5.8", "2.3", "3.5",
        { color: C.PRIMARY, fontSize: 36, rowH: 0.50 });
      slide.addText("Inverse: 3.5 + 2.3 = 5.8 ✓", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 2.70, w: lg.rightW - 0.20, h: 0.26,
        fontSize: 13, fontFace: FONT_H, color: C.SUCCESS, bold: true,
        align: "center", margin: 0,
      });
    }
  );

  // I Do 2 — 6.3 - 2.85 (regrouping)
  workedExSlide(pres, 2, "I Do", "Subtract 6.3 − 2.85 (regrouping)",
    [
      "1.  ESTIMATE: 6 − 3 ≈ 3.",
      "",
      "2.  PLACEHOLDER ZERO: 6.3 → 6.30.",
      "",
      "3.  LINE UP the decimal point.",
      "",
      "4.  SUBTRACT, regrouping where needed.",
      "     Hundredths: 0 − 5 → regroup.",
      "     Tenths: 2 − 8 → regroup.",
      "     Ones: 5 − 2 = 3.",
      "",
      "5.  Answer: 3.45.",
      "",
      "6.  Inverse check: 3.45 + 2.85 = 6.30  ✓",
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
      const stackX = lg.rightX + 0.30;
      const stackW = lg.rightW - 0.60;
      const rowH = 0.50;
      const rowY = lg.panelTopPadded + 0.78;
      // Top: "6.30" with the placeholder 0 in red
      slide.addText([
        { text: "6.3", options: { fontSize: 36, fontFace: FONT_H, color: C.CHARCOAL, bold: true } },
        { text: "0", options: { fontSize: 36, fontFace: FONT_H, color: C.ALERT, bold: true } },
      ], {
        x: stackX, y: rowY, w: stackW, h: rowH,
        align: "center", valign: "middle", margin: 0,
      });
      // Minus row
      slide.addText("−   2.85", {
        x: stackX, y: rowY + rowH, w: stackW, h: rowH,
        fontSize: 36, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addShape("line", {
        x: stackX + 0.4, y: rowY + 2 * rowH + 0.04, w: stackW - 0.8, h: 0,
        line: { color: C.SECONDARY, width: 2 },
      });
      slide.addText("3.45", {
        x: stackX, y: rowY + 2 * rowH + 0.12, w: stackW, h: rowH,
        fontSize: 36, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // CFU — 7.6 - 3.2
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Subtract 7.6 − 3.2", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 20, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Subtract 7.6 − 3.2", options: { fontSize: 28, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "1.  Estimate first.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "2.  Line up the decimal point.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "3.  Subtract column by column.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "4.  Check by adding back.", options: { fontSize: 16, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.3, CONTENT_TOP, 4.2, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      s.addText("Stacked subtraction", {
        x: 5.3, y: CONTENT_TOP + 0.10, w: 4.2, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      addStackedSubtraction(s, 5.6, CONTENT_TOP + 0.70, 3.6, "7.6", "3.2", null,
        { color: C.PRIMARY, fontSize: 38 });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "7.6 − 3.2 = 4.4   (check: 4.4 + 3.2 = 7.6 ✓)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // We Do — 4.2 - 1.85
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Subtract 4.2 − 1.85",
      [
        "With your partner.",
        "",
        "1.  ESTIMATE first.",
        "2.  PLACEHOLDER ZERO on 4.2.",
        "3.  LINE UP. SUBTRACT (regroup).",
        "4.  CHECK by adding back.",
        "",
        "Show your working step by step.",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY });
        slide.addText("Build the calculation", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        addStackedSubtraction(slide, lg.rightX + 0.30, lg.panelTopPadded + 0.55,
          lg.rightW - 0.60, "4.2", "1.85", null,
          { color: C.SECONDARY, fontSize: 34 });
        slide.addText("(Placeholder zero on 4.2 → 4.20)", {
          x: lg.rightX + 0.20, y: lg.panelTopPadded + 2.50, w: lg.rightW - 0.40, h: 0.35,
          fontSize: 14, fontFace: FONT_B, color: C.MUTED, italic: true,
          align: "center", valign: "middle", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "4.20 − 1.85 = 2.35   Inverse: 2.35 + 1.85 = 4.20 ✓", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // CFU Hinge — inverse-check 5.3 - 2.18 = 3.18 ?
  withReveal(
    () => cfuSlide(pres, "CFU", "Use inverse-check", "Thumbs Up or Thumbs Down",
      "A student wrote:\n\n5.3 − 2.18 = 3.18\n\nUse inverse-add to check.\n3.18 + 2.18 = ?\n\nDoes it equal 5.3?",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "NOT correct.   3.18 + 2.18 = 5.36, not 5.3.   Correct answer: 5.30 − 2.18 = 3.12.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
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
      { text: "Section 1 — same-length subtraction.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 2 — different lengths (placeholder zero + regroup).   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 3 — word problem.", options: { fontSize: 18, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.45;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    s.addText("Always check by inverse-add", {
      x: 0.7, y: panelY + 0.15, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });

    addTextOnShape(s, "If a − b = c, then c + b = a.   Add back to check.", {
      x: 0.8, y: panelY + 0.65, w: 8.4, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    addTextOnShape(s, "Different lengths → PLACEHOLDER ZERO first, then regroup.", {
      x: 1.0, y: panelY + 1.20, w: 8.0, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  exitTicketSlide(pres,
    [
      "Calculate 8.4 − 3.75. Show your working and check by inverse-add.",
      "A student wrote 6.0 − 2.45 = 3.65. Use inverse-add to check. Is it correct? If not, fix it and explain the error.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: why is inverse-check useful?",
      scItems: [
        "I can subtract two decimals when both have the same number of decimal places.",
        "I can subtract decimals with different lengths using a placeholder zero and regrouping.",
        "I can check my subtraction by using inverse-add (the answer plus the subtrahend equals the original).",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecFrac_Lesson4_Subtracting_Decimals.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Subtract decimals using place value. Always check by inverse-add.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "Process: estimate, line up the decimal point, subtract (regroup if needed), check by inverse-add. Different lengths → placeholder zero first.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 — Same-length subtraction", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   5.8 − 2.3   =   _______      Check: ___ + 2.3 = ___", y);
    y = addWriteLine(doc, "b)   7.6 − 3.2   =   _______      Check: ___ + 3.2 = ___", y);
    y = addWriteLine(doc, "c)   9.45 − 4.23 = _______       Check: ___ + 4.23 = ___", y);

    y = addSectionHeading(doc, "Section 2 — Different lengths (placeholder zero)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   4.2 − 1.85   =   _______      Check: ___ + 1.85 = ___", y);
    y = addWriteLine(doc, "b)   6.3 − 2.85   =   _______      Check: ___ + 2.85 = ___", y);
    y = addWriteLine(doc, "c)   8.0 − 3.467  =  _______      Check: ___ + 3.467 = ___", y);

    y = addSectionHeading(doc, "Section 3 — Word problem", y, { color: C.PRIMARY });
    y = addBodyText(doc, "At the school athletics, the Year 6 record for the long jump was 3.85 metres. Maya jumped 3.4 metres at this year's carnival.", y);
    y = addWriteLine(doc, "How far off the record was Maya? Answer: ______ m", y);
    y = addWriteLine(doc, "Show your working: __________________________________________________", y);

    y = addSectionHeading(doc, "Section 4 — Extension (optional, mental partitioning)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Solve mentally using the inverse (add up) or jump strategy. Show your steps.", y);
    y = addWriteLine(doc, "a)   5.2 − 2.7    Steps: __________________   Answer: ______", y);
    y = addWriteLine(doc, "b)   8.05 − 3.6   Steps: __________________   Answer: ______", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Subtracting Decimals Practice | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the subtracting decimals practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 — Same length", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3.5   (check: 3.5 + 2.3 = 5.8 ✓)", y);
    y = addBodyText(doc, "b)  4.4   (check: 4.4 + 3.2 = 7.6 ✓)", y);
    y = addBodyText(doc, "c)  5.22  (check: 5.22 + 4.23 = 9.45 ✓)", y);

    y = addSectionHeading(doc, "Section 2 — Different lengths", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  2.35  (4.20 − 1.85; check: 2.35 + 1.85 = 4.20 ✓)", y);
    y = addBodyText(doc, "b)  3.45  (6.30 − 2.85; check: 3.45 + 2.85 = 6.30 ✓)", y);
    y = addBodyText(doc, "c)  4.533 (8.000 − 3.467; check: 4.533 + 3.467 = 8.000 ✓)", y);

    y = addSectionHeading(doc, "Section 3 — Word problem", y, { color: C.PRIMARY });
    y = addBodyText(doc, "3.85 − 3.40 = 0.45 m. Maya was 0.45 m off the record.", y);

    y = addSectionHeading(doc, "Section 4 — Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "a)  5.2 − 2.7 = 2.5.   Example steps: 2.7 + 0.3 → 3.0; 3.0 + 2.0 → 5.0; 5.0 + 0.2 → 5.2. Total jumps: 0.3 + 2.0 + 0.2 = 2.5.", y);
    y = addBodyText(doc, "b)  8.05 − 3.6 = 4.45. Example steps: 3.6 + 0.4 → 4.0; 4.0 + 4.0 → 8.0; 8.0 + 0.05 → 8.05. Total: 0.4 + 4.0 + 0.05 = 4.45.", y);

    y = addTipBox(doc,
      "Watch for: students who write 4.2 − 1.85 = 3.85 or 3.65 - they skipped the placeholder zero or did not regroup. Students who write 6.0 − 2.45 = 3.65 are still on yesterday's misalignment trap. The inverse-add check is the fix - if it does not equal the original, the answer is wrong.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 4 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
