"use strict";

// Decimals & Fractions Unit (Year 5/6 Numeracy) — Lesson 6: Subtracting fractions / LCD
// VC2M6N03 — solve problems involving addition AND subtraction of fractions using
//            knowledge of equivalent fractions; determining the lowest common
//            denominator using prime and composite numbers.
// Daily Review: Adding and subtracting fractions and decimals.
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
  addFractionStripSet,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 6;
const TOTAL = 6;
const UNIT_TITLE = "Decimals and Fractions";
const FOOTER = `Decimals & Fractions | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/DecFrac_Lesson6_Subtracting_Fractions_LCD";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 6 Subtracting Fractions Practice",
  "Subtracting fractions using the lowest common denominator.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 6 Answer Key",
  "Worked answers for the subtracting fractions sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Final lesson of the unit.
- Today we subtract fractions and use the lowest common denominator — the LCD.
- The big idea: make the parts the same SIZE, then subtract.

DO:
- Have whiteboards, markers, and fraction strips ready.

TEACHER NOTES:
Lesson 6 of 6. Subtraction works the same way as Lesson 5's addition: same-size parts first. The new layer is finding the LCD when denominators are NOT related (one is not a multiple of the other), which connects to factors and multiples from Lesson 5's Daily Review.

WATCH FOR:
- Students still adding/subtracting denominators - reteach the strip.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The subtracting fractions practice sheet is for the You Do.

DO:
- Print the practice sheet, answer key, and fraction strip reference.
- Have whiteboards ready.

TEACHER NOTES:
One student resource (practice sheet) plus answer key. The fraction strip is the main visual.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Warm up. Adding and subtracting fractions and decimals.
- One of each. On your whiteboard.

DO:
- Display the three prompts.
- 90 seconds.

TEACHER NOTES:
Daily Review consolidates the whole week's work. Quick fire.

WATCH FOR:
- Students who can do all three - secure unit-wide.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check.
- 0.5 + 0.45 is 0.95.
- 4.2 minus 1.75 is 2.45.
- 1/4 + 3/4 is 4/4 or 1 whole.

DO:
- Tick and fix.

TEACHER NOTES:
This consolidates the whole unit. Any persistent error here is the focus for the next unit's review.

WATCH FOR:
- Students who got all three - secure.
- Students who got the fraction wrong - reteach equivalent fractions.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Adding and subtracting decimals.

DO:
- 60 seconds.

TEACHER NOTES:
Last fluency check on decimals for this unit.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check.
- 2.6 plus 1.4 is 4.0.
- 5.0 minus 2.35 is 2.65.
- 0.75 + 0.25 is 1.0.

DO:
- Tick and fix.

TEACHER NOTES:
Standard checkpoint - students should be confident.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention.
- We are learning to subtract fractions using the lowest common denominator.
- Now the success criteria.

DO:
- Choral read.

TEACHER NOTES:
SC1 floor - subtract same-denominator fractions. SC2 core target - subtract related-denominator fractions. SC3 depth - find the LCD when denominators are not related.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me subtract 5/8 minus 1/8.
- Same denominator. Same SIZE parts.
- I take 1 eighth away from 5 eighths.
- 5 minus 1 is 4. Answer: 4/8.
- Simplify: 4/8 = 1/2.

DO:
- Display the fraction strip showing 5/8 with 1/8 crossed out.
- Point to the remaining 4 eighths.

TEACHER NOTES:
Same-denominator subtraction. The strip makes the move visual. This sets up the harder cases.

MISCONCEPTIONS:
- Misconception: Subtract numerator and denominator.
  Why: Same trap as addition.
  Impact: 5/8 - 1/8 = 4/0 (nonsense) or 4.
  Quick correction: "Subtract numerators only. Keep the size of the part the same."

WATCH FOR:
- Students who say 4 eighths - secure.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now I subtract 2/3 minus 1/4. Denominators are 3 and 4 - not related.
- I need a common denominator that BOTH 3 and 4 divide into.
- Multiples of 3: 3, 6, 9, 12...
- Multiples of 4: 4, 8, 12...
- The lowest common multiple is 12. So 12 is the LCD.
- Rewrite. 2/3 = 8/12 (multiply top and bottom by 4). 1/4 = 3/12 (multiply by 3).
- Subtract. 8/12 - 3/12 = 5/12.

DO:
- List the multiples of 3 and 4 on the board.
- Circle 12.
- Show the rewriting move.

TEACHER NOTES:
This is the new layer. Finding the LCD by listing multiples is the most concrete method. Connects directly to yesterday's Daily Review on multiples and to today's factor work.

MISCONCEPTIONS:
- Misconception: Use ANY common denominator (e.g., 24 instead of 12).
  Why: Any common multiple works mathematically but creates harder calculations.
  Impact: Bigger numbers, more chance of error.
  Quick correction: "Find the LOWEST common multiple to keep numbers small."

WATCH FOR:
- Students who list multiples and find 12 - secure.
- Students who multiply 3 x 4 = 12 - also works for these denominators, but flag for harder cases.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. Subtract 7/10 minus 3/10.
- Same denominator. On your whiteboard.

DO:
- 30 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your answer on three, two, one, show."
- Scan for: 4/10 (or 2/5).
PROCEED: If 80% correct, click to reveal.
PIVOT: Most likely misconception - students subtracted denominators.
- Reteach: "Subtract numerators only. Keep the denominator."
- Re-check: "Try again."

TEACHER NOTES:
Same-denominator check. Quick.

WATCH FOR:
- Students who hold up 4/10 or 2/5 - secure.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Subtract 3/4 minus 1/6.
- Denominators are 4 and 6. Not related.
- Step 1: list multiples to find the LCD.
- Step 2: rewrite both fractions.
- Step 3: subtract.
- Use your whiteboard.

DO:
- Display the prompt.
- 2 minutes.
- Walk and listen for the multiples thinking.

TEACHER NOTES:
Different denominators, neither a multiple of the other. The LCD step is essential.

WATCH FOR:
- Pairs who list multiples first - secure.
- Pairs who try 4 x 6 = 24 - also works, but inefficient (LCD is 12).

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- Multiples of 4: 4, 8, 12, 16. Multiples of 6: 6, 12. The LCD is 12.
- 3/4 = 9/12 (multiply by 3). 1/6 = 2/12 (multiply by 2).
- 9/12 minus 2/12 = 7/12.

DO:
- Click to reveal.
- Run the LCD step once more.

TEACHER NOTES:
The reveal restates the LCD-find and rewrite moves. Students who got 7/12 are on target.

WATCH FOR:
- Students who self-correct - secure.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. Which is the LCD for these two fractions: 3/8 and 1/6?
- Choices: A) 14   B) 24   C) 48
- Thumbs UP for A. Make a 2 with your fingers for B. Make a 3 for C.

DO:
- Display the choices.
- 20 seconds.

CFU CHECKPOINT:
Technique: Choose A B or C
Script:
- Say: "Show me A, B, or C with your fingers."
- Scan for: B (24). Multiples of 8: 8, 16, 24. Multiples of 6: 6, 12, 18, 24. LCD = 24.
PROCEED: If 80% choose B, click to reveal.
PIVOT: Most likely misconception - students added denominators (14) or multiplied (48).
- Reteach: "List multiples. Find the LOWEST that appears in BOTH lists."
- Re-check: "What is the LCD for 4 and 6?"

TEACHER NOTES:
The hinge probes whether students find the LCD by listing multiples instead of guessing or just multiplying. 24 is correct; 48 (8 x 6) works but is not the LCD; 14 (8 + 6) is the classic add-denominators trap.

WATCH FOR:
- Confident B - they listed multiples.
- A (14) - the add-denominators trap.
- C (48) - they just multiplied.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own. Take the practice sheet.
- Section 1: same denominator.
- Section 2: related denominators - rewrite the smaller one.
- Section 3: find the LCD first - denominators NOT related.

DO:
- Distribute the practice sheet.
- Circulate.

TEACHER NOTES:
Section 3 is the new layer from today. Watch for students who list multiples vs. those who just multiply.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use only Sections 1 and 2 with the printed fraction strip.
- Extra Notes: Sit with these students. Find the LCD by listing on paper.
EXTENDING PROMPT:
- Task: Section 4 - word problem applying fraction subtraction.
- Extra Notes: Encourage them to draw their model.

WATCH FOR:
- Students who list multiples and find the LCD - secure on threshold.
- Students who multiply denominators - works but inefficient.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task.
- Question 1: 3/4 minus 5/8. Find the LCD and subtract.
- Question 2: A student wrote 2/3 minus 1/4 = 1/12. Use a fraction strip or LCD method to check. Is the answer correct?

DO:
- Display.
- 4 minutes.
- Collect whiteboards.

TEACHER NOTES:
Exit ticket assesses SC2 (related denominators) and SC3 (LCD when denominators are not related).

WATCH FOR:
- Students who got 1/8 with the LCD = 8 step - secure on SC2.
- Students who corrected Q2 to 5/12 - secure on SC3.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- We finish the unit today.
- Look at the success criteria. Thumbs for each.
- Turn and tell your partner: what is the lowest common denominator and why does it matter?

DO:
- Read each I can statement.
- Use thumbs.
- Celebrate the unit completion.

TEACHER NOTES:
Final closing of the unit. Students who can explain the LCD have crossed the threshold. The fraction-and-decimal addition/subtraction work is now ready for problem-solving applications.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, UNIT_TITLE, "Lesson 6: Subtracting fractions / LCD",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  withReveal(
    () => dailyReviewSlide(pres,
      "Daily Review: Add & subtract — fractions and decimals",
      ["0.5 + 0.45", "4.2 − 1.75", "1/4 + 3/4"],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "0.95     2.45     4/4 = 1", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  withReveal(
    () => fluencySlide(pres, "Fluency: Adding & Subtracting Decimals",
      ["2.6 + 1.4", "5.0 − 2.35", "0.75 + 0.25"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "4.0     2.65     1.0", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  liSlide(pres,
    "We are learning to subtract fractions using the lowest common denominator.",
    [
      "I can subtract two fractions when they have the same denominator.",
      "I can subtract fractions with related denominators by rewriting one fraction.",
      "I can find the lowest common denominator and use it to subtract fractions with unrelated denominators.",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do 1 — 5/8 - 1/8
  workedExSlide(pres, 2, "I Do", "Subtract 5/8 − 1/8 (same denominator)",
    [
      "Same denominator → same-size parts.",
      "",
      "Subtract the NUMERATORS only.",
      "Keep the denominator the same.",
      "",
      "5/8 − 1/8 = 4/8.",
      "",
      "Simplify: 4/8 = 1/2.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("5/8 then take 1/8 away", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      addFractionStripSet(slide,
        lg.rightX + 0.20, lg.panelTopPadded + 0.50, lg.rightW - 0.40, 1.80,
        [
          { denom: 8, shaded: 5, label: "5/8" },
          { denom: 8, shaded: 1, label: "−1/8" },
          { denom: 8, shaded: 4, label: "4/8 = 1/2" },
        ],
        { labelW: 1.10, labelFontSize: 12 });
      slide.addText("5/8 − 1/8 = 4/8 (= 1/2)", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 2.50, w: lg.rightW - 0.20, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.SUCCESS, bold: true,
        align: "center", margin: 0,
      });
    }
  );

  // I Do 2 — 2/3 - 1/4 with LCD
  workedExSlide(pres, 2, "I Do", "Subtract 2/3 − 1/4 (find the LCD)",
    [
      "Denominators 3 and 4 are NOT related.",
      "",
      "1.  List multiples to find the LCD:",
      "     Multiples of 3: 3, 6, 9, 12...",
      "     Multiples of 4: 4, 8, 12...",
      "     LCD = 12.",
      "",
      "2.  Rewrite with denominator 12:",
      "     2/3 = 8/12 (× 4 top and bottom).",
      "     1/4 = 3/12 (× 3 top and bottom).",
      "",
      "3.  Subtract: 8/12 − 3/12 = 5/12.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY });
      slide.addText("LCD method", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      // Multiples listing
      slide.addText([
        { text: "Multiples of 3: ", options: { fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bold: true } },
        { text: "3, 6, 9, ", options: { fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL } },
        { text: "12", options: { fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true } },
      ], {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.50, w: lg.rightW - 0.30, h: 0.30,
        margin: 0, valign: "middle",
      });
      slide.addText([
        { text: "Multiples of 4: ", options: { fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bold: true } },
        { text: "4, 8, ", options: { fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL } },
        { text: "12", options: { fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true } },
      ], {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.85, w: lg.rightW - 0.30, h: 0.30,
        margin: 0, valign: "middle",
      });

      // LCD banner
      addTextOnShape(slide, "LCD = 12", {
        x: lg.rightX + 0.50, y: lg.panelTopPadded + 1.25, w: lg.rightW - 1.0, h: 0.40, rectRadius: 0.08,
        fill: { color: C.ALERT },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });

      // Rewriting
      slide.addText("2/3 = 8/12        1/4 = 3/12", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 1.85, w: lg.rightW - 0.20, h: 0.36,
        fontSize: 17, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Answer
      slide.addText("8/12 − 3/12 = 5/12", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 2.40, w: lg.rightW - 0.20, h: 0.40,
        fontSize: 19, fontFace: FONT_H, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // CFU — 7/10 - 3/10
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Subtract 7/10 − 3/10", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 20, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "7/10 − 3/10 =", options: { fontSize: 30, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Same denominator.", options: { fontSize: 18, color: C.CHARCOAL, breakLine: true } },
        { text: "Subtract numerators only.", options: { fontSize: 18, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.3, CONTENT_TOP, 4.2, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      s.addText("Tenths strip", {
        x: 5.3, y: CONTENT_TOP + 0.10, w: 4.2, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      addFractionStripSet(s,
        5.45, CONTENT_TOP + 0.55, 3.90, 2.00,
        [
          { denom: 10, shaded: 7, label: "7/10" },
          { denom: 10, shaded: 3, label: "−3/10" },
          { denom: 10, shaded: 0, label: "?/10" },
        ],
        { labelW: 1.00, labelFontSize: 11 });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "7/10 − 3/10 = 4/10 (= 2/5)", {
        x: 1.0, y: 4.55, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // We Do — 3/4 - 1/6
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Subtract 3/4 − 1/6",
      [
        "With your partner.",
        "",
        "Denominators 4 and 6. Not related.",
        "",
        "1.  List multiples to find the LCD.",
        "     Multiples of 4: ___, ___, ___.",
        "     Multiples of 6: ___, ___.",
        "",
        "2.  Rewrite both with the LCD.",
        "",
        "3.  Subtract.",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY });
        slide.addText("Build the LCD", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.28,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        // Empty multiples scaffold
        slide.addText("Multiples of 4:", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.30, h: 0.30,
          fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
          valign: "middle", margin: 0,
        });
        slide.addShape("line", {
          x: lg.rightX + 1.50, y: lg.panelTopPadded + 0.85, w: lg.rightW - 1.65, h: 0,
          line: { color: C.MUTED, width: 1 },
        });
        slide.addText("Multiples of 6:", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.10, w: lg.rightW - 0.30, h: 0.30,
          fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
          valign: "middle", margin: 0,
        });
        slide.addShape("line", {
          x: lg.rightX + 1.50, y: lg.panelTopPadded + 1.40, w: lg.rightW - 1.65, h: 0,
          line: { color: C.MUTED, width: 1 },
        });

        // LCD blank
        addTextOnShape(slide, "LCD = ____", {
          x: lg.rightX + 0.50, y: lg.panelTopPadded + 1.75, w: lg.rightW - 1.0, h: 0.45, rectRadius: 0.08,
          fill: { color: C.ALERT },
        }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });

        slide.addText("3/4 = ___/___    1/6 = ___/___", {
          x: lg.rightX + 0.10, y: lg.panelTopPadded + 2.40, w: lg.rightW - 0.20, h: 0.36,
          fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "LCD = 12.   3/4 = 9/12.   1/6 = 2/12.   9/12 − 2/12 = 7/12.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // CFU Hinge — LCD for 8 and 6
  withReveal(
    () => cfuSlide(pres, "CFU", "Which is the LCD?", "Choose A, B or C",
      "For 3/8 and 1/6, which is the LCD?\n\nA)  14   (8 + 6)\nB)  24\nC)  48   (8 × 6)\n\nShow A, B, or C with your fingers.",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "B = 24.   Multiples of 8: 8, 16, 24. Multiples of 6: 6, 12, 18, 24.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
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
      { text: "Section 1 — same denominator.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 2 — related denominators.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 3 — find the LCD first.", options: { fontSize: 18, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.45;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    s.addText("Finding the LCD", {
      x: 0.7, y: panelY + 0.15, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });

    addTextOnShape(s, "List multiples of each denominator.   First match is the LCD.", {
      x: 0.8, y: panelY + 0.65, w: 8.4, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    addTextOnShape(s, "Rewrite each fraction → then subtract numerators only.", {
      x: 1.0, y: panelY + 1.20, w: 8.0, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  exitTicketSlide(pres,
    [
      "Calculate 3/4 − 5/8. Find the LCD, rewrite, and subtract. Show your working.",
      "A student wrote 2/3 − 1/4 = 1/12. Check this answer. If wrong, fix it and explain.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what is the lowest common denominator and why does it matter?",
      scItems: [
        "I can subtract two fractions when they have the same denominator.",
        "I can subtract fractions with related denominators by rewriting one fraction.",
        "I can find the lowest common denominator and use it to subtract fractions with unrelated denominators.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecFrac_Lesson6_Subtracting_Fractions_LCD.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Subtract fractions using the lowest common denominator.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "Rule: only subtract same-size parts. Different denominators → find the LCD by listing multiples, rewrite each fraction, then subtract the numerators.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 — Same denominator", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   5/8 − 1/8     =   _______", y);
    y = addWriteLine(doc, "b)   7/10 − 3/10   =   _______", y);
    y = addWriteLine(doc, "c)   9/12 − 4/12   =   _______", y);

    y = addSectionHeading(doc, "Section 2 — Related denominators (rewrite first)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   3/4 − 1/8     Rewrite 3/4 = ___/8.   Answer: _______", y);
    y = addWriteLine(doc, "b)   5/6 − 1/3     Rewrite 1/3 = ___/6.   Answer: _______", y);
    y = addWriteLine(doc, "c)   7/10 − 2/5    Rewrite 2/5 = ___/10.  Answer: _______", y);

    y = addSectionHeading(doc, "Section 3 — Find the LCD first (unrelated denominators)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   2/3 − 1/4     LCD: ___   Rewrite:  2/3 = ___/___, 1/4 = ___/___.   Answer: _______", y);
    y = addWriteLine(doc, "b)   3/4 − 1/6     LCD: ___   Rewrite:  3/4 = ___/___, 1/6 = ___/___.   Answer: _______", y);
    y = addWriteLine(doc, "c)   5/6 − 3/8     LCD: ___   Rewrite:  5/6 = ___/___, 3/8 = ___/___.   Answer: _______", y);

    y = addSectionHeading(doc, "Section 4 — Word problem (optional extension)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Ava had 3/4 of a litre of paint. She used 1/3 of a litre painting her bedroom door.", y);
    y = addWriteLine(doc, "How much paint does Ava have left? Answer: _______ L", y);
    y = addWriteLine(doc, "Show your working: __________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Subtracting Fractions Practice | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the subtracting fractions practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 — Same denominator", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  4/8 (= 1/2)     b)  4/10 (= 2/5)     c)  5/12", y);

    y = addSectionHeading(doc, "Section 2 — Related denominators", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3/4 = 6/8, so 6/8 − 1/8 = 5/8.", y);
    y = addBodyText(doc, "b)  1/3 = 2/6, so 5/6 − 2/6 = 3/6 (= 1/2).", y);
    y = addBodyText(doc, "c)  2/5 = 4/10, so 7/10 − 4/10 = 3/10.", y);

    y = addSectionHeading(doc, "Section 3 — Find the LCD", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  LCD = 12.   2/3 = 8/12.   1/4 = 3/12.   8/12 − 3/12 = 5/12.", y);
    y = addBodyText(doc, "b)  LCD = 12.   3/4 = 9/12.   1/6 = 2/12.   9/12 − 2/12 = 7/12.", y);
    y = addBodyText(doc, "c)  LCD = 24.   5/6 = 20/24.  3/8 = 9/24.   20/24 − 9/24 = 11/24.", y);

    y = addSectionHeading(doc, "Section 4 — Word problem", y, { color: C.ACCENT });
    y = addBodyText(doc, "LCD of 4 and 3 = 12. 3/4 = 9/12. 1/3 = 4/12. 9/12 − 4/12 = 5/12 L.", y);

    y = addTipBox(doc,
      "Watch for: students who subtract denominators (e.g., 5/8 − 1/8 = 4/0) - reteach with the strip. Students who guess the LCD (or multiply denominators automatically) - prompt them to list multiples. Multiplying always works for the common denominator, but is not always the LOWEST.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 6 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
