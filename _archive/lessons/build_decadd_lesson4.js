"use strict";

// Adding & Subtracting Decimals Unit (Year 5/6 Numeracy) — Lesson 4: Subtracting decimals
// First-time learning. Same line-up rule, same placeholder zero, plus regrouping across the decimal point.
// Daily Review: faces of 3D shapes (links 2D and 3D from earlier in the week).
// Fluency: multiplication facts (x6).

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

const SESSION = 4;
const TOTAL = 5;
const UNIT_TITLE = "Adding and Subtracting Decimals";
const FOOTER = `Add & Subtract Decimals | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/DecAdd_Lesson4_Subtracting_Decimals";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 4 Subtracting Decimals Practice",
  "Subtracting decimals using line-up, placeholder zero, and trading.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 4 Answer Key",
  "Worked answers for the subtracting decimals practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Today we move from adding decimals to subtracting them.
- The big rule is exactly the same: line up the decimal point.
- We still use the placeholder zero when lengths are different.
- And we trade between columns when the top digit is too small to take from.

DO:
- Whiteboards, markers, printed place value reference cards ready.

TEACHER NOTES:
Lesson 4 of 5. Subtraction uses the same alignment habit as addition. Trading across the decimal point is the only new move - and it works exactly like trading in whole-number subtraction.

WATCH FOR:
- Students wobbly on the placeholder zero from Lesson 3 - quick recap in I Do 1.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The subtracting decimals practice sheet is for the You Do section.

DO:
- Print one copy of the practice sheet and answer key per student.

TEACHER NOTES:
One student resource (practice sheet) plus answer key.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Match the 2D face to the 3D shape it comes from.
- On your whiteboard, write the letter of the matching 2D face for each 3D shape.

DO:
- Display three 3D shapes and three 2D faces.
- 90 seconds.

TEACHER NOTES:
This links 2D and 3D from earlier in the week. Listen for "the cube's face is a square."

WATCH FOR:
- Students who can match - secure.
- Students who guess - revisit faces in next lesson's review.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check.
- Cube - matches the square (A).
- Cylinder - matches the circle (B).
- Square pyramid - matches the triangle (C) for its side faces.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
A cube has 6 square faces. A cylinder has 2 circle faces and a curved face. A square pyramid has a square base and 4 triangle faces.

WATCH FOR:
- Students who self-correct - secure.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Six times tables today.
- Whisper-answer each one, then write it on your board.

DO:
- Display the three prompts.
- 45 seconds.

TEACHER NOTES:
x6 facts. These are trickier than x2, x3, x4. Use doubling strategies if students stall: x6 = double x3.

WATCH FOR:
- Students who answer instantly - secure.
- Students who double x3 - valid strategy, prompt them to recall.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check.
- 6 times 4 is 24.
- 6 times 7 is 42.
- 6 times 8 is 48.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Tick and fix.

WATCH FOR:
- Students who self-correct - secure.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to subtract decimals by lining up the decimal point.
- Now the success criteria.

DO:
- Choral read.

TEACHER NOTES:
SC1 is achievable - subtracting two decimals of the same length. SC2 is the core target (different lengths with placeholder zero). SC3 stretches to checking with addition.

WATCH FOR:
- Students who can repeat the line-up rule - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me subtract 0.7 minus 0.3.
- These are both tenths.
- Estimate: 0.7 is just under 1, 0.3 is small, so the answer is just under 0.5.
- Line up the decimal points.
- Subtract the tenths: 7 - 3 = 4 tenths.
- Bring the decimal point straight down.
- 0.4. Check against estimate. Yes, reasonable.

DO:
- Display the worked example stacked.
- Trace through the decimal point.
- Repeat "tenths minus tenths" three times.

TEACHER NOTES:
This first I Do is intentionally simple - both tenths, no trading. The aim is to install the line-up habit for subtraction before any complication.

MISCONCEPTIONS:
- Misconception: Students drop the decimal point and write 4.
  Why: They strip the dot when subtracting.
  Impact: Answer changes by a factor of 10.
  Quick correction: "Bring the dot straight down. It marks the place."

WATCH FOR:
- Students who say "tenths minus tenths" with you - tracking.
- Students who answer 4 instead of 0.4 - dropped-dot trap.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now a trickier one: 4.5 minus 1.27.
- Different lengths. 4.5 has 1 decimal place. 1.27 has 2.
- Write the placeholder zero: 4.5 becomes 4.50.
- Estimate: 4 - 1 = 3.
- Line up the decimal points.
- Start from the right. Hundredths: 0 - 7. I cannot do that. I need to TRADE.
- Trade 1 tenth into 10 hundredths. The 5 in tenths becomes 4. The 0 in hundredths becomes 10.
- Now hundredths: 10 - 7 = 3.
- Tenths: 4 - 2 = 2.
- Ones: 4 - 1 = 3.
- 3.23. Check against estimate of 3. Yes, reasonable.

DO:
- Display 4.50 above 1.27 with the dots aligned.
- Cross out the 5 and write 4 above it.
- Above the 0 write a small 10.
- Walk through the trade slowly.

TEACHER NOTES:
Trading across the decimal point is the only new move in subtraction. It works exactly like trading in whole-number subtraction. The place value chart from Lesson 1 makes this easy to see.

MISCONCEPTIONS:
- Misconception: Students cannot do 0 - 7 and write 7 (just take small from big).
  Why: They reverse the digits to avoid the trade.
  Impact: Wrong answer.
  Quick correction: "We cannot subtract more from less. We trade 1 tenth for 10 hundredths."

WATCH FOR:
- Students who trade correctly - secure.
- Students who flip the digits (7 - 0 = 7) - the trade trap.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. 0.9 minus 0.4.
- On your whiteboard. Show your line-up.

DO:
- Display the prompt.
- 30 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your answer on three, two, one, show."
- Scan for: 0.5.
PROCEED: If 80% have 0.5, reveal.
PIVOT: Most likely misconception - dropped-dot (wrote 5) or wrong place value.
- Reteach: "Tenths minus tenths. 9 tenths minus 4 tenths = 5 tenths = 0.5."
- Re-check: "Try 0.9 - 0.4 again."

TEACHER NOTES:
Tenths-only subtraction with no trading. Quick check before we add trading complexity.

WATCH FOR:
- Students who write 0.5 - secure.
- Students who write 5 - dropped dot.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. 3.2 minus 1.45.
- First: ESTIMATE. About 3 - 1 = 2.
- Then: write the PLACEHOLDER ZERO. 3.2 becomes 3.20.
- Line up the decimal points.
- Subtract right to left. You will need to TRADE.
- Check against your estimate.

DO:
- Display the two numbers.
- 90 seconds.
- Walk and listen for "trade" or "regroup."

TEACHER NOTES:
Different lengths AND trading. This is the most complex subtraction we will model today.

WATCH FOR:
- Students who add a placeholder zero, then trade correctly - secure.
- Students who write 2.25 (no placeholder, no trade) - reteach.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- 3.20 - 1.45.
- Hundredths: 0 - 5. Cannot do. Trade 1 tenth for 10 hundredths.
- Now tenths becomes 1 and hundredths becomes 10.
- Hundredths: 10 - 5 = 5.
- Tenths: 1 - 4. Cannot do. Trade 1 one for 10 tenths.
- Now ones becomes 2 and tenths becomes 11.
- Tenths: 11 - 4 = 7.
- Ones: 2 - 1 = 1.
- 1.75. Estimate was 2 - yes, reasonable.

DO:
- Click to reveal.
- Walk through the two trades slowly.

TEACHER NOTES:
The reveal shows the two-trade sequence. Some students will need to see this 2-3 times before it clicks.

WATCH FOR:
- Students who self-correct - secure.
- Students who keep 2.25 - reteach with the chart.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. Look at this student's work.
- 5.0 minus 1.4 = 4.4.
- Is this reasonable? Thumbs UP for yes, DOWN for no.

DO:
- Display the work.
- 15 seconds.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Thumbs UP if reasonable. DOWN if not."
- Scan for: thumbs DOWN. 5 - 1 = 4, but we have to subtract a bit more (0.4), so answer should be 3.6.
PROCEED: If 80% thumbs down, click to reveal.
PIVOT: Most likely misconception - students subtract big-from-small in the tenths column (4 - 0 = 4).
- Reteach: "Trade 1 one for 10 tenths. Now ones is 4 and tenths is 10. 10 - 4 = 6 tenths."
- Re-check: "What is the real answer?"

TEACHER NOTES:
This hinge surfaces the "I cannot trade across the dot" misconception.

WATCH FOR:
- Confident thumbs down - they used estimation.
- Students who hesitate or thumbs up - need the trade reteach.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Section 1: same-length subtraction (no trade).
- Section 2: different lengths (placeholder zero).
- Section 3: trading across the decimal point.
- Section 4: word problem.

DO:
- Distribute the practice sheet.
- Circulate.
- Cold call 1-2 students to share Section 4 thinking.

TEACHER NOTES:
The practice sheet sequences from easy to hardest.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use Sections 1 and 2 only with the printed place value chart.
- Extra Notes: Sit with these students. Do the first trading example together.
EXTENDING PROMPT:
- Task: Section 5 - check each answer by adding it back to the smaller number to get the bigger number (inverse).
- Extra Notes: Encourage students to write the addition check beside each subtraction.

WATCH FOR:
- Students who trade correctly - secure.
- Students who flip digits in the trading column - reteach.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task.
- Question 1: 5.6 - 2.34. Show your placeholder zero and any trading.
- Question 2: A student wrote 4.0 - 1.7 = 3.3. Is this correct? If not, fix it and explain.

DO:
- Display the prompt.
- 3 minutes.
- Collect whiteboards.

TEACHER NOTES:
Exit ticket assesses SC2 (different lengths with possible trading). Q2 probes the trading misconception.

WATCH FOR:
- Students who write 3.26 - secure.
- Students who fix Q2 to 2.3 with an explanation - they understand trading.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria. Thumbs for each.
- Turn and tell your partner: what is the only new move in subtracting decimals?

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
The threshold idea is "trade across the decimal point exactly like in whole-number subtraction." Students who can say this are ready for word problems tomorrow.

WATCH FOR:
- Students who say "trading" or "regrouping" - secure.
- Students unsure - revisit at the start of Lesson 5.

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

  titleSlide(pres, UNIT_TITLE, "Lesson 4: Subtracting decimals",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Daily Review + reveal — match 2D faces to 3D shapes
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Daily Review: Match the 2D face to the 3D shape");

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });

      const y3D = CONTENT_TOP + 0.35;
      const y2D = CONTENT_TOP + 2.10;
      const label = (x, y, w, txt) =>
        s.addText(txt, {
          x, y, w, h: 0.32,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
          align: "center", margin: 0,
        });

      s.addText("3D shapes", {
        x: 0.5, y: y3D - 0.30, w: 9.0, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, italic: true, bold: true,
        align: "center", margin: 0,
      });

      // Cube
      {
        const x = 1.3, y = y3D, side = 1.0, off = 0.22;
        s.addShape("rect", {
          x: x + off, y, w: side, h: side,
          fill: { color: C.SECONDARY }, line: { color: C.CHARCOAL, width: 1 },
        });
        s.addShape("rect", {
          x, y: y + off, w: side, h: side,
          fill: { color: C.PRIMARY }, line: { color: C.CHARCOAL, width: 1 },
        });
        label(1.0, y + side + 0.30, 1.8, "Cube");
      }

      // Cylinder
      {
        const cx = 4.2, cy = y3D, w = 1.3, h = 1.2;
        s.addShape("rect", {
          x: cx, y: cy + 0.20, w, h: h - 0.40,
          fill: { color: C.ACCENT }, line: { color: C.ACCENT, width: 0 },
        });
        s.addShape("ellipse", {
          x: cx, y: cy, w, h: 0.35,
          fill: { color: C.ACCENT }, line: { color: C.CHARCOAL, width: 1 },
        });
        s.addShape("ellipse", {
          x: cx, y: cy + h - 0.35, w, h: 0.35,
          fill: { color: C.ACCENT }, line: { color: C.CHARCOAL, width: 1 },
        });
        label(3.9, y3D + h + 0.30, 1.9, "Cylinder");
      }

      // Square pyramid (triangle with a base line and slanted edges)
      {
        const cx = 7.0, cy = y3D, w = 1.5, h = 1.2;
        s.addShape("triangle", {
          x: cx, y: cy, w, h,
          fill: { color: C.ALERT }, line: { color: C.CHARCOAL, width: 1 },
        });
        // base ellipse to suggest 3D
        s.addShape("ellipse", {
          x: cx - 0.15, y: cy + h - 0.18, w: w + 0.30, h: 0.30,
          fill: { color: C.ALERT }, line: { color: C.CHARCOAL, width: 1 },
        });
        label(6.7, y3D + h + 0.30, 2.0, "Square pyramid");
      }

      // 2D faces
      s.addText("2D faces", {
        x: 0.5, y: y2D - 0.30, w: 9.0, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, italic: true, bold: true,
        align: "center", margin: 0,
      });

      // A: Square
      s.addShape("rect", {
        x: 1.6, y: y2D, w: 1.0, h: 1.0,
        fill: { color: C.WHITE }, line: { color: C.CHARCOAL, width: 1.2 },
      });
      label(1.2, y2D + 1.0, 1.8, "A");

      // B: Circle
      s.addShape("ellipse", {
        x: 4.5, y: y2D, w: 1.0, h: 1.0,
        fill: { color: C.WHITE }, line: { color: C.CHARCOAL, width: 1.2 },
      });
      label(4.1, y2D + 1.0, 1.8, "B");

      // C: Triangle
      s.addShape("triangle", {
        x: 7.3, y: y2D, w: 1.0, h: 1.0,
        fill: { color: C.WHITE }, line: { color: C.CHARCOAL, width: 1.2 },
      });
      label(6.9, y2D + 1.0, 1.8, "C");

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Cube → A (square)   |   Cylinder → B (circle)   |   Square pyramid → C (triangle, side face)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Fluency + reveal: x6
  withReveal(
    () => fluencySlide(pres, "Fluency: Multiplication (x6)",
      ["6 × 4", "6 × 7", "6 × 8"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "24     42     48", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // LI/SC
  liSlide(pres,
    "We are learning to subtract decimals by lining up the decimal point.",
    [
      "I can subtract two decimals (same length) by lining up the decimal point.",
      "I can subtract decimals of different lengths by using a placeholder zero and trading.",
      "I can check my subtraction is right by adding the answer back to the smaller number.",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do 1: 0.7 - 0.3
  workedExSlide(pres, 2, "I Do", "Subtract 0.7 - 0.3",
    [
      "1.  ESTIMATE: just under 0.5.",
      "",
      "2.  LINE UP the decimal points.",
      "",
      "3.  SUBTRACT the tenths.",
      "     7 - 3 = 4 tenths.",
      "",
      "4.  BRING DOWN the decimal point.",
      "",
      "5.  0.4. Check: just under 0.5.  ✓",
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
        lg.rightW - 0.60, "0.7", "0.3", "0.4",
        { color: C.PRIMARY, fontSize: 36, rowH: 0.50 });
    }
  );

  // I Do 2: 4.5 - 1.27 (trading)
  workedExSlide(pres, 2, "I Do", "Subtract 4.5 - 1.27 (with trading)",
    [
      "1.  ESTIMATE: 4 - 1 = 3.",
      "",
      "2.  Placeholder zero: 4.5 → 4.50.",
      "",
      "3.  Hundredths: 0 - 7. TRADE!",
      "     1 tenth → 10 hundredths.",
      "     Tenths becomes 4. Hundredths becomes 10.",
      "",
      "4.  Hundredths: 10 - 7 = 3.",
      "     Tenths:     4 - 2 = 2.",
      "     Ones:       4 - 1 = 3.",
      "",
      "5.  3.23 ≈ 3.  ✓",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY });
      slide.addText("Trade across the dot", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      const stackX = lg.rightX + 0.30;
      const stackW = lg.rightW - 0.60;
      const rowH = 0.50;
      const rowY = lg.panelTopPadded + 0.78;

      // Top: 4.50 with placeholder zero in red, and the trade annotation above
      slide.addText("4 . 4 10", {
        x: stackX, y: rowY - 0.30, w: stackW, h: 0.28,
        fontSize: 16, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      slide.addText([
        { text: "4.5",
          options: { fontSize: 36, fontFace: FONT_H, color: C.CHARCOAL, bold: true } },
        { text: "0",
          options: { fontSize: 36, fontFace: FONT_H, color: C.ALERT, bold: true } },
      ], {
        x: stackX, y: rowY, w: stackW, h: rowH,
        align: "center", valign: "middle", margin: 0,
      });

      slide.addText("−   1.27", {
        x: stackX, y: rowY + rowH, w: stackW, h: rowH,
        fontSize: 36, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      slide.addShape("line", {
        x: stackX + 0.4, y: rowY + 2 * rowH + 0.04, w: stackW - 0.8, h: 0,
        line: { color: C.SECONDARY, width: 2 },
      });

      slide.addText("3.23", {
        x: stackX, y: rowY + 2 * rowH + 0.12, w: stackW, h: rowH,
        fontSize: 36, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // CFU + reveal: 0.9 - 0.4
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Subtract 0.9 - 0.4", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 20, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Subtract 0.9 - 0.4", options: { fontSize: 28, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "1.  Stack with the dot lined up.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "2.  Subtract the tenths.", options: { fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "3.  Bring the dot straight down.", options: { fontSize: 16, color: C.CHARCOAL } },
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
      addStackedSubtraction(s, 5.6, CONTENT_TOP + 0.70, 3.6, "0.9", "0.4", null,
        { color: C.PRIMARY, fontSize: 38 });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "0.9 - 0.4 = 0.5", {
        x: 1.0, y: 4.55, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // We Do + reveal: 3.2 - 1.45
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Subtract 3.2 - 1.45",
      [
        "With your partner.",
        "",
        "1.  ESTIMATE first.",
        "2.  Write the PLACEHOLDER ZERO.",
        "3.  LINE UP and SUBTRACT right to left.",
        "4.  TRADE where needed.",
        "5.  CHECK against your estimate.",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY });
        slide.addText("Two trades needed", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        addStackedSubtraction(slide, lg.rightX + 0.30, lg.panelTopPadded + 0.55,
          lg.rightW - 0.60, "3.2", "1.45", null,
          { color: C.SECONDARY, fontSize: 34 });
        slide.addText("(Add placeholder zero, then trade twice.)", {
          x: lg.rightX + 0.20, y: lg.panelTopPadded + 2.50, w: lg.rightW - 0.40, h: 0.35,
          fontSize: 14, fontFace: FONT_B, color: C.MUTED, italic: true,
          align: "center", valign: "middle", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "3.20 - 1.45 = 1.75    (trade 1 tenth → 10 hundredths, then 1 one → 10 tenths)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Hinge + reveal
  withReveal(
    () => cfuSlide(pres, "CFU", "Reasonable or not?", "Thumbs Up or Thumbs Down",
      "A student wrote:\n\n5.0 - 1.4 = 4.4\n\nIs this answer reasonable?\nUse estimation.",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "NOT reasonable.   5 - 1 ≈ 4, then subtract a bit more (0.4) → about 3.6.   Correct: 3.6.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
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
      { text: "Sections: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "1 — same length   |   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "2 — placeholder zero   |   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "3 — trading   |   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "4 — word problem.", options: { fontSize: 18, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.45;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    s.addText("Subtraction routine", {
      x: 0.7, y: panelY + 0.15, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });

    addTextOnShape(s, "1.  Estimate   2.  Line up the dot   3.  Placeholder zero   4.  Trade if needed   5.  Check", {
      x: 0.7, y: panelY + 0.65, w: 8.6, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });

    addTextOnShape(s, "Cannot subtract? TRADE 1 from the left column → 10 in this column.", {
      x: 1.0, y: panelY + 1.20, w: 8.0, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Exit ticket
  exitTicketSlide(pres,
    [
      "Subtract 5.6 - 2.34. Show your placeholder zero and any trading.",
      "A student wrote 4.0 - 1.7 = 3.3. Is this correct? If not, fix it and explain the mistake.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what is the only new move in subtracting decimals?",
      scItems: [
        "I can subtract two decimals (same length) by lining up the decimal point.",
        "I can subtract decimals of different lengths by using a placeholder zero and trading.",
        "I can check my subtraction is right by adding the answer back to the smaller number.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecAdd_Lesson4_Subtracting_Decimals.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // PDFs
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Subtract decimals by lining up the decimal point. Use placeholder zero and trading where needed.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "Every time: estimate first, line up the decimal point, write the placeholder zero, subtract right to left, trade where needed, check against your estimate.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 — Same length, no trading", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   0.7 - 0.3   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "b)   2.6 - 1.4   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "c)   3.78 - 1.25 =   _______      Estimate: ___", y);

    y = addSectionHeading(doc, "Section 2 — Different lengths (placeholder zero)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   3.6 - 1.25   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "b)   5.4 - 0.27   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "c)   2.9 - 1.07   =   _______      Estimate: ___", y);

    y = addSectionHeading(doc, "Section 3 — With trading", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   4.5 - 1.27   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "b)   3.2 - 1.45   =   _______      Estimate: ___", y);
    y = addWriteLine(doc, "c)   6.0 - 2.38   =   _______      Estimate: ___", y);

    y = addSectionHeading(doc, "Section 4 — Word problem", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Ms Patel's water bottle holds 1.5 litres. She drinks 0.65 litres at recess.", y);
    y = addWriteLine(doc, "How much water is left? Answer: ______ L", y);

    y = addSectionHeading(doc, "Section 5 — Inverse check (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Pick any problem from Section 3. Check your answer by adding it back to the smaller number. You should get the bigger number.", y);
    y = addWriteLine(doc, "Problem chosen: ______ + ______ = ______", y);

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

    y = addSectionHeading(doc, "Section 1 — Same length, no trading", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  0.4     b)  1.2     c)  2.53", y);

    y = addSectionHeading(doc, "Section 2 — Different lengths", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3.60 - 1.25 = 2.35     b)  5.40 - 0.27 = 5.13     c)  2.90 - 1.07 = 1.83", y);

    y = addSectionHeading(doc, "Section 3 — With trading", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  4.50 - 1.27 = 3.23     b)  3.20 - 1.45 = 1.75     c)  6.00 - 2.38 = 3.62", y);

    y = addSectionHeading(doc, "Section 4 — Word problem", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1.50 - 0.65 = 0.85 L left.", y);

    y = addSectionHeading(doc, "Section 5 — Inverse check", y, { color: C.ACCENT });
    y = addBodyText(doc, "Example: 3.23 + 1.27 = 4.50. The answer adds back to the bigger number, so the subtraction is correct.", y);

    y = addTipBox(doc,
      "Watch for: students who do 5.4 - 0.27 = 5.17 (right-align without placeholder zero); students who do 4.5 - 1.27 by flipping 0 - 7 into 7 - 0 (avoiding the trade); students who drop the decimal point in the answer.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 4 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
