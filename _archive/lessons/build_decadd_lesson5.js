"use strict";

// Adding & Subtracting Decimals Unit (Year 5/6 Numeracy) — Lesson 5: Decimal word problems (money)
// First-time learning unit consolidation. Real-world add/subtract decimals (money + measurement).
// Daily Review: mixed 2D + 3D shape names.
// Fluency: multiplication facts (x7) - mixed times tables review.

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

const SESSION = 5;
const TOTAL = 5;
const UNIT_TITLE = "Adding and Subtracting Decimals";
const FOOTER = `Add & Subtract Decimals | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/DecAdd_Lesson5_Decimal_Word_Problems";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 5 Decimal Word Problems Practice",
  "Money and measurement word problems using add and subtract decimals.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 5 Answer Key",
  "Worked answers for the decimal word problems practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Last lesson of the unit. Today we use everything we have learned.
- We solve money and measurement word problems by adding and subtracting decimals.
- We use the same rules: estimate, line up the dot, placeholder zero, trade if needed, check.

DO:
- Whiteboards, markers, printed place value reference cards ready.

TEACHER NOTES:
Lesson 5 of 5. Consolidation lesson. The mathematics is what students have learned in lessons 2-4. The new move is selecting whether to add or subtract from a word problem. Most contexts are money - dollars and cents read very naturally as decimals.

WATCH FOR:
- Students who can do the arithmetic but freeze on word problems - that is normal. The DRAW-WRITE-CHECK frame helps.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The decimal word problems practice sheet is for the You Do section.

DO:
- Print one copy of the practice sheet and answer key per student.

TEACHER NOTES:
One student resource (practice sheet) plus answer key.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Today is mixed 2D and 3D shapes.
- For each picture, write the name and write "2D" or "3D".

DO:
- Display the three shapes.
- 90 seconds.

TEACHER NOTES:
Last review of the shape strand. Students have seen each shape this week.

WATCH FOR:
- Students who confidently name and label - secure.
- Students confused between sphere and circle - reteach: sphere is 3D, circle is 2D.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check.
- A = pentagon, 2D.
- B = sphere, 3D.
- C = rectangular prism, 3D.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Tick and fix. A rectangular prism has 6 rectangular faces (some may be squares).

WATCH FOR:
- Students who self-correct - secure.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Seven times tables today.
- These are some of the trickiest. Take your best go.

DO:
- Display the three prompts.
- 45 seconds.

TEACHER NOTES:
x7 facts. The trickiest set. Strategies: x7 = x5 + x2. Or commute (7 x 4 = 4 x 7).

WATCH FOR:
- Students who answer instantly - secure.
- Students who count on - flag for fact recall practice next week.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check.
- 7 times 4 is 28.
- 7 times 6 is 42.
- 7 times 8 is 56.

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
- We are learning to solve word problems by adding and subtracting decimals.
- Now the success criteria.

DO:
- Choral read.

TEACHER NOTES:
SC1 is achievable - choosing to add or subtract from key words. SC2 is the core target (solving the problem and explaining). SC3 stretches to multi-step problems.

WATCH FOR:
- Students who can repeat the DRAW-WRITE-CHECK frame - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me solve this problem.
- "Liam buys a book for $12.45 and a pen for $2.80. How much did he spend in total?"
- Step 1: DRAW. I draw a bar with two parts: $12.45 and $2.80, with a total above.
- Step 2: WRITE the calculation. "In total" means add. 12.45 + 2.80.
- Estimate: 12 + 3 = 15.
- Add placeholder zero on 2.80 if needed - it already has 2 decimal places. Good.
- Line up the decimal points. Hundredths: 5 + 0 = 5. Tenths: 4 + 8 = 12, write 2 carry 1. Ones: 2 + 2 + 1 = 5. Tens: 1 + 0 = 1.
- 15.25.
- Step 3: CHECK. $15.25 is close to my estimate of $15. Answer: $15.25.

DO:
- Display the problem.
- Draw the bar model live on a whiteboard or screen.
- Walk through DRAW-WRITE-CHECK aloud.

TEACHER NOTES:
The bar model is the bridge from words to maths. "In total" / "altogether" / "combined" signal addition.

MISCONCEPTIONS:
- Misconception: Students see two numbers and just add without reading.
  Why: They short-circuit the problem.
  Impact: Wrong operation chosen.
  Quick correction: "Read the question word for word. Picture what is happening."

WATCH FOR:
- Students who can label the bar model - they understand the situation.
- Students who guess - they need the bar model first.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now a subtraction problem.
- "A jug holds 1.5 litres of juice. We pour out 0.65 litres for the class. How much is left?"
- DRAW: a bar of 1.5 with a part of 0.65 removed and the unknown labelled "?".
- WRITE: "How much is left" means subtract. 1.5 - 0.65.
- Estimate: 1.5 - 0.65 is about 0.85.
- Placeholder zero! 1.5 becomes 1.50.
- 1.50 - 0.65. Hundredths: 0 - 5. Trade. Tenths becomes 4, hundredths becomes 10.
- Hundredths: 10 - 5 = 5. Tenths: 4 - 6. Trade. Ones becomes 0, tenths becomes 14.
- Tenths: 14 - 6 = 8. Ones: 0 - 0 = 0.
- 0.85 L. Check against estimate of 0.85. Yes!

DO:
- Display the problem.
- Draw the "missing piece" bar model.
- Walk through the trade slowly.

TEACHER NOTES:
"How much is left" / "how much more" / "difference" signal subtraction. The two-trade sequence appears here - revisit slowly.

MISCONCEPTIONS:
- Misconception: 1.5 - 0.65 = 1.15 (no placeholder, no trade - just flip the digits).
  Why: Students avoid the trade by reversing the subtraction.
  Impact: Wrong answer.
  Quick correction: "Placeholder zero first. Then trade properly."

WATCH FOR:
- Students who get 0.85 - secure.
- Students who get 1.15 - reteach with the place value chart.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check.
- "Mia has $5.00. She buys a juice for $2.45. How much does she have left?"
- On your whiteboard. Show your subtraction.

DO:
- Display the prompt.
- 90 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your answer on three, two, one, show."
- Scan for: $2.55.
PROCEED: If 80% have $2.55, reveal.
PIVOT: Most likely misconception - students forget the placeholder zero or the trade.
- Reteach: "5.00 - 2.45. Trade 1 one for 10 tenths. Then 1 tenth for 10 hundredths."
- Re-check: "Try it again with the trades written above."

TEACHER NOTES:
Subtraction with trading and a placeholder zero on 5.00. A common money problem.

WATCH FOR:
- Students who write $2.55 - secure.
- Students who write $2.45 - they did 5.00 - 2.55 (mixed up).
- Students who write $3.55 or similar - trading errors.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me.
- "Tom buys a pie for $4.80 and a milk for $2.50. He pays with a $10 note. How much change does he get?"
- This is a TWO-STEP problem. First add. Then subtract.
- DRAW. WRITE. CHECK.

DO:
- Display the problem.
- 3 minutes.
- Walk and listen.

TEACHER NOTES:
Two-step problem. Students need to (1) add the two prices, then (2) subtract from $10. Watch for students who try to subtract one item at a time and lose track.

WATCH FOR:
- Students who add first, then subtract - secure.
- Students who subtract twice from $10 - acceptable as a method, but slower.
- Students who add all three numbers - misread the problem.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- Step 1: 4.80 + 2.50.
- Hundredths: 0 + 0 = 0. Tenths: 8 + 5 = 13, write 3 carry 1. Ones: 4 + 2 + 1 = 7. = $7.30.
- Step 2: 10.00 - 7.30. Trade. = $2.70.
- Tom gets $2.70 change.

DO:
- Click to reveal.
- Walk through both steps.

TEACHER NOTES:
The two-step structure is the main learning move. Confirm the steps in order.

WATCH FOR:
- Students who self-correct - secure.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. Look at this student's work.
- Problem: "A bottle holds 2 L. We pour out 0.4 L. How much is left?"
- Student answer: 2 - 0.4 = 1.4 L.
- Thumbs UP if you agree. DOWN if you do not.

DO:
- Display the work.
- 15 seconds.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Thumbs UP if you agree. DOWN if you do not."
- Scan for: thumbs UP. The answer is correct (1.6 L is also a valid alt - check details).
PROCEED: If 80% thumbs up with reasoning, reveal.
PIVOT: Most likely misconception - some students may compute 2 - 0.4 = 2.4 (added) or 2.4 (dropped trade).
- Reteach: "2.0 - 0.4. Trade 1 one for 10 tenths. 10 tenths - 4 tenths = 6 tenths. Answer 1.6 L."
- Re-check: "Is 1.4 right? Use estimation."

Actually, 2 - 0.4 = 1.6 (not 1.4). The student's answer is INCORRECT. Adjust your script: scan for thumbs DOWN.

TEACHER NOTES:
The "student answer" deliberately shows the common error of subtracting tenths incorrectly without trading. The correct answer is 1.6 L. Thumbs DOWN is the secure response.

WATCH FOR:
- Confident thumbs down - they spotted the trading error.
- Thumbs up - they accepted the wrong answer; trade reteach needed.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Section 1: one-step add or subtract.
- Section 2: two-step problems (add THEN subtract).
- Section 3: measurement problems.

DO:
- Distribute the practice sheet.
- Circulate.
- Cold call 1-2 students to share Section 2 thinking.

TEACHER NOTES:
The practice sheet sequences one-step then two-step.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Section 1 only with a bar model template for each problem.
- Extra Notes: Sit with these students. Draw the first bar model together.
EXTENDING PROMPT:
- Task: Section 4 - design your own two-step money problem and solve it.
- Extra Notes: Encourage students to swap problems with a partner.

WATCH FOR:
- Students who DRAW first - secure.
- Students who guess add vs subtract - prompt the bar model.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task of the unit.
- Question 1: A water bottle holds 1.25 L. You drink 0.4 L. How much is left?
- Question 2: A sandwich costs $6.50 and a drink costs $3.75. How much change from $15?

DO:
- Display the prompt.
- 5 minutes.
- Collect whiteboards.

TEACHER NOTES:
Exit ticket assesses SC2. Q1 tests subtraction with placeholder zero and trading. Q2 tests two-step problem solving.

WATCH FOR:
- Q1: students who get 0.85 L - secure.
- Q2: students who get $4.75 - secure on the two-step structure.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria. Thumbs for each.
- Turn and tell your partner: what is one thing you can do now that you could not do at the start of the week?

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
End-of-unit reflection. Acknowledge progress. Most students should now confidently add and subtract decimals using place value.

WATCH FOR:
- Students who share growth - celebrate.
- Students still unsure - keep on the small-group radar for revisit next week.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Helpers ────────────────────────────────────────────────────────────────

function addStackedAddition(slide, x, y, w, top, bottom, answer, opts) {
  const o = opts || {};
  const lineColor = o.color || C.PRIMARY;
  const fontSize = o.fontSize || 28;
  const rowH = o.rowH || 0.46;
  const symbol = o.symbol || "+";

  slide.addText(top, {
    x, y, w, h: rowH,
    fontSize, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  slide.addText(symbol + "   " + bottom, {
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

// Draw a bar model: total bar split into two parts (or part + unknown).
function addBarModel(slide, x, y, w, h, parts, opts) {
  const o = opts || {};
  const totalWidth = parts.reduce((acc, p) => acc + p.weight, 0);
  let cursor = x;
  parts.forEach((p, i) => {
    const partW = (p.weight / totalWidth) * w;
    slide.addShape("rect", {
      x: cursor, y, w: partW, h,
      fill: { color: p.color || (i % 2 === 0 ? C.PRIMARY : C.SECONDARY) },
      line: { color: C.CHARCOAL, width: 0.75 },
    });
    slide.addText(String(p.label), {
      x: cursor, y, w: partW, h,
      fontSize: o.fontSize || 14, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    cursor += partW;
  });
}

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, UNIT_TITLE, "Lesson 5: Decimal word problems (money + measurement)",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Daily Review + reveal — mixed 2D and 3D
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Daily Review: Name and label (2D or 3D)");

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });

      const shapeY = CONTENT_TOP + 0.45;

      // A: Pentagon (2D)
      s.addShape("pentagon", {
        x: 1.3, y: shapeY, w: 1.6, h: 1.5,
        fill: { color: C.PRIMARY }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("A", {
        x: 1.0, y: shapeY + 1.60, w: 2.2, h: 0.35,
        fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });

      // B: Sphere (3D - circle with shading suggesting a sphere)
      {
        const cx = 4.0, cy = shapeY, w = 1.6, h = 1.5;
        s.addShape("ellipse", {
          x: cx, y: cy, w, h,
          fill: { color: C.SECONDARY }, line: { color: C.CHARCOAL, width: 1 },
        });
        // highlight oval
        s.addShape("ellipse", {
          x: cx + 0.30, y: cy + 0.20, w: 0.55, h: 0.30,
          fill: { color: C.WHITE }, line: { color: C.WHITE, width: 0 },
        });
        s.addText("B", {
          x: 3.7, y: shapeY + 1.60, w: 2.2, h: 0.35,
          fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
          align: "center", margin: 0,
        });
      }

      // C: Rectangular prism (3D)
      {
        const x = 6.9, y = shapeY + 0.20, w = 1.6, h = 1.1, off = 0.30;
        s.addShape("rect", {
          x: x + off, y, w, h,
          fill: { color: C.ACCENT }, line: { color: C.CHARCOAL, width: 1 },
        });
        s.addShape("rect", {
          x, y: y + off, w, h,
          fill: { color: C.ALERT }, line: { color: C.CHARCOAL, width: 1 },
        });
        s.addShape("line", {
          x, y: y + off, w: off, h: 0,
          line: { color: C.CHARCOAL, width: 1 },
        });
        s.addShape("line", {
          x: x + w, y: y + off, w: off, h: 0,
          line: { color: C.CHARCOAL, width: 1 },
        });
        s.addText("C", {
          x: 6.6, y: shapeY + 1.60, w: 2.2, h: 0.35,
          fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
          align: "center", margin: 0,
        });
      }

      s.addText("For each shape: name + 2D or 3D.", {
        x: 0.7, y: SAFE_BOTTOM - 0.55, w: 8.6, h: 0.35,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "A = pentagon (2D)   |   B = sphere (3D)   |   C = rectangular prism (3D)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Fluency + reveal: x7
  withReveal(
    () => fluencySlide(pres, "Fluency: Multiplication (x7)",
      ["7 × 4", "7 × 6", "7 × 8"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "28     42     56", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // LI/SC
  liSlide(pres,
    "We are learning to solve word problems by adding and subtracting decimals.",
    [
      "I can decide if a word problem needs adding or subtracting.",
      "I can draw a bar model, write the calculation and solve a money or measurement problem.",
      "I can solve a two-step problem (add then subtract) and check the answer is reasonable.",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do 1: Liam buys a book and a pen
  workedExSlide(pres, 2, "I Do", "$12.45 book + $2.80 pen = ?",
    [
      "Liam buys a book for $12.45",
      "and a pen for $2.80.",
      "How much did he spend in total?",
      "",
      "1.  DRAW the bar model.",
      "2.  WRITE: \"in total\" → ADD.",
      "     12.45 + 2.80.",
      "3.  Estimate: 12 + 3 = 15.",
      "4.  Add with the dots lined up.",
      "5.  Answer: $15.25.  ✓",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("Bar model + calculation", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });

      // Bar model showing $12.45 + $2.80 = ? total
      addBarModel(slide,
        lg.rightX + 0.15, lg.panelTopPadded + 0.50,
        lg.rightW - 0.30, 0.50,
        [
          { weight: 12.45, label: "$12.45", color: C.PRIMARY },
          { weight: 2.80,  label: "$2.80",  color: C.SECONDARY },
        ],
        { fontSize: 13 });
      slide.addText("Total = ?", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.05, w: lg.rightW - 0.30, h: 0.30,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });

      // Stacked addition
      addStackedAddition(slide, lg.rightX + 0.40, lg.panelTopPadded + 1.45,
        lg.rightW - 0.80, "$12.45", "$2.80", "$15.25",
        { color: C.PRIMARY, fontSize: 22, rowH: 0.36 });
    }
  );

  // I Do 2: 1.5L juice - 0.65L poured out
  workedExSlide(pres, 2, "I Do", "1.5 L − 0.65 L = ?",
    [
      "A jug holds 1.5 L of juice.",
      "We pour out 0.65 L.",
      "How much is left?",
      "",
      "1.  DRAW the bar with the missing piece.",
      "2.  WRITE: \"how much is left\" → SUBTRACT.",
      "3.  Placeholder zero: 1.5 → 1.50.",
      "4.  Trade twice.",
      "5.  Answer: 0.85 L.  ✓",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY });
      slide.addText("Bar model with missing part", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      // Whole = 1.5; part1 = 0.65 poured out; part2 = ? left
      addBarModel(slide,
        lg.rightX + 0.15, lg.panelTopPadded + 0.50,
        lg.rightW - 0.30, 0.50,
        [
          { weight: 0.65, label: "Poured 0.65 L", color: C.ALERT },
          { weight: 0.85, label: "? L",           color: C.SECONDARY },
        ],
        { fontSize: 12 });
      slide.addText("Whole jug = 1.5 L", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.05, w: lg.rightW - 0.30, h: 0.30,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", margin: 0,
      });

      addStackedAddition(slide, lg.rightX + 0.40, lg.panelTopPadded + 1.45,
        lg.rightW - 0.80, "1.50", "0.65", "0.85",
        { color: C.SECONDARY, fontSize: 22, rowH: 0.36, symbol: "−" });
    }
  );

  // CFU + reveal: $5.00 - $2.45
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Mia's change", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 20, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Mia has $5.00.", options: { fontSize: 20, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "She buys a juice for $2.45.", options: { fontSize: 18, color: C.CHARCOAL, breakLine: true } },
        { text: "How much does she have left?", options: { fontSize: 18, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Show the subtraction.", options: { fontSize: 16, color: C.ALERT, bold: true } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.3, CONTENT_TOP, 4.2, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      s.addText("Subtraction", {
        x: 5.3, y: CONTENT_TOP + 0.10, w: 4.2, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      addStackedAddition(s, 5.6, CONTENT_TOP + 0.70, 3.6, "$5.00", "$2.45", null,
        { color: C.PRIMARY, fontSize: 32, symbol: "−" });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "$5.00 − $2.45 = $2.55", {
        x: 1.0, y: 4.55, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // We Do + reveal: two-step problem
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Tom's change",
      [
        "Tom buys a pie for $4.80",
        "and a milk for $2.50.",
        "He pays with a $10 note.",
        "How much change does he get?",
        "",
        "TWO steps:",
        "1.  ADD the items.",
        "2.  SUBTRACT from $10.",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY });
        slide.addText("Plan it together", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });

        // Step 1 box
        addTextOnShape(slide, "Step 1:  $4.80  +  $2.50  =  ?", {
          x: lg.rightX + 0.20, y: lg.panelTopPadded + 0.50, w: lg.rightW - 0.40, h: 0.55,
          rectRadius: 0.08, fill: { color: C.PRIMARY },
        }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

        // Step 2 box
        addTextOnShape(slide, "Step 2:  $10.00  −  (Step 1)  =  ?", {
          x: lg.rightX + 0.20, y: lg.panelTopPadded + 1.15, w: lg.rightW - 0.40, h: 0.55,
          rectRadius: 0.08, fill: { color: C.SECONDARY },
        }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

        slide.addText("Estimate the change before solving.", {
          x: lg.rightX + 0.20, y: lg.panelTopPadded + 1.85, w: lg.rightW - 0.40, h: 0.30,
          fontSize: 13, fontFace: FONT_B, color: C.MUTED, italic: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "Step 1: $4.80 + $2.50 = $7.30   |   Step 2: $10.00 − $7.30 = $2.70", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Hinge + reveal: 2 - 0.4 = 1.4 (WRONG, correct is 1.6)
  withReveal(
    () => cfuSlide(pres, "CFU", "Is this answer right?", "Thumbs Up or Thumbs Down",
      "Problem: A bottle holds 2 L. We pour out 0.4 L.\nHow much is left?\n\nStudent answer:   2 − 0.4 = 1.4 L\n\nThumbs UP if correct.   Thumbs DOWN if not.",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "NOT correct.  2.0 − 0.4: trade 1 one → 10 tenths, then 10 − 4 = 6 tenths.   Real answer: 1.6 L.", {
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
      { text: "Sections: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "1 — one step   |   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "2 — two-step (add then subtract)   |   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "3 — measurement.", options: { fontSize: 18, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.45;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    s.addText("Every problem", {
      x: 0.7, y: panelY + 0.15, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });

    addTextOnShape(s, "1.  DRAW a bar model    2.  WRITE the calculation    3.  CHECK with an estimate", {
      x: 0.7, y: panelY + 0.65, w: 8.6, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });

    addTextOnShape(s, "Cue words: total / altogether / combined = ADD   |   left / change / difference = SUBTRACT", {
      x: 0.6, y: panelY + 1.20, w: 8.8, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Exit ticket
  exitTicketSlide(pres,
    [
      "A water bottle holds 1.25 L. You drink 0.4 L. How much is left?",
      "A sandwich costs $6.50 and a drink costs $3.75. How much change from $15?",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what is one thing you can do now that you could not do at the start of the week?",
      scItems: [
        "I can decide if a word problem needs adding or subtracting.",
        "I can draw a bar model, write the calculation and solve a money or measurement problem.",
        "I can solve a two-step problem (add then subtract) and check the answer is reasonable.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecAdd_Lesson5_Decimal_Word_Problems.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // PDFs
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Use add and subtract decimals to solve real-world money and measurement problems.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "Every problem: DRAW a bar model, WRITE the calculation, CHECK with an estimate. Total / altogether = ADD. Left / change / difference = SUBTRACT.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 — One-step problems", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1.  Ahmed buys a book for $14.60 and a notebook for $3.25. How much does he spend in total?", y);
    y = addWriteLine(doc, "Calculation: ______________________   Answer: $______", y);
    y = addBodyText(doc, "2.  Mia has $20.00. She buys a meal for $12.45. How much change does she get?", y);
    y = addWriteLine(doc, "Calculation: ______________________   Answer: $______", y);
    y = addBodyText(doc, "3.  Jamal runs 2.4 km, then 1.85 km. How far did he run in total?", y);
    y = addWriteLine(doc, "Calculation: ______________________   Answer: ______ km", y);

    y = addSectionHeading(doc, "Section 2 — Two-step problems", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1.  Sam buys a pen for $1.45 and a notebook for $3.80. He pays with $10. How much change does he get?", y);
    y = addWriteLine(doc, "Step 1: ____________________   Step 2: ____________________   Answer: $______", y);
    y = addBodyText(doc, "2.  A jug holds 2 L. Ms Lee pours 0.45 L into one glass and 0.7 L into another. How much is left in the jug?", y);
    y = addWriteLine(doc, "Step 1: ____________________   Step 2: ____________________   Answer: ______ L", y);

    y = addSectionHeading(doc, "Section 3 — Measurement problems", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1.  Maya's plant was 1.2 m tall. It grew 0.45 m. How tall is it now?", y);
    y = addWriteLine(doc, "Calculation: ______________________   Answer: ______ m", y);
    y = addBodyText(doc, "2.  A bag weighs 4.5 kg. Lucas takes out 0.8 kg of books. How much does the bag weigh now?", y);
    y = addWriteLine(doc, "Calculation: ______________________   Answer: ______ kg", y);

    y = addSectionHeading(doc, "Section 4 — Extension (design your own)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Write your own two-step money problem. Then solve it. Swap with a partner.", y);
    y = addWriteLine(doc, "My problem: __________________________________________________", y);
    y = addWriteLine(doc, "My solution: __________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Decimal Word Problems | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the decimal word problems practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 — One-step problems", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1.  14.60 + 3.25 = $17.85.", y);
    y = addBodyText(doc, "2.  20.00 - 12.45 = $7.55.", y);
    y = addBodyText(doc, "3.  2.40 + 1.85 = 4.25 km.", y);

    y = addSectionHeading(doc, "Section 2 — Two-step problems", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1.  Step 1: 1.45 + 3.80 = $5.25.    Step 2: 10.00 - 5.25 = $4.75 change.", y);
    y = addBodyText(doc, "2.  Step 1: 0.45 + 0.70 = 1.15 L poured.    Step 2: 2.00 - 1.15 = 0.85 L left.", y);

    y = addSectionHeading(doc, "Section 3 — Measurement", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1.  1.20 + 0.45 = 1.65 m.", y);
    y = addBodyText(doc, "2.  4.50 - 0.80 = 3.70 kg.", y);

    y = addSectionHeading(doc, "Section 4 — Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "Accept any reasonable two-step problem that uses add and subtract decimals. Look for a written calculation and a check.", y);

    y = addTipBox(doc,
      "Watch for: students who add when they should subtract (and vice versa) - prompt them to draw the bar model first; students who lose track of the steps in two-step problems - encourage labelling Step 1 and Step 2 clearly.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 5 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
