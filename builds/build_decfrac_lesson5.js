"use strict";

// Decimals & Fractions Unit (Year 5/6 Numeracy) — Lesson 5: Adding fractions with related denominators
// VC2M6N03 — solve problems involving addition of fractions using equivalent fractions.
// Daily Review: Number patterns with factors and multiples.
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

const SESSION = 5;
const TOTAL = 6;
const UNIT_TITLE = "Decimals and Fractions";
const FOOTER = `Decimals & Fractions | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/DecFrac_Lesson5_Adding_Fractions";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 5 Adding Fractions Practice",
  "Adding fractions with related denominators using equivalent fractions.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 5 Answer Key",
  "Worked answers for the adding fractions practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Today we switch to fractions.
- We are learning to add fractions when the denominators are related — one is a multiple of the other.
- The key idea is making the parts the same SIZE.

DO:
- Have whiteboards, markers, and printed fraction strips ready.

TEACHER NOTES:
Lesson 5 of 6. Adding fractions with related denominators is the bridge between adding-same-denominator (prior learning) and adding any fractions (Lesson 6). The big idea: change one fraction into an equivalent fraction so the parts are the same size.

WATCH FOR:
- Students who try to add the denominators - the classic trap. We address it in I Do.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- Fraction strips for the I Do and We Do. Practice sheet for the You Do.

DO:
- Print the practice sheet, answer key, and a fraction strip reference.
- Have whiteboards ready.

TEACHER NOTES:
The fraction strip is the main visual model today. Print one reference strip per pair for the You Do.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Warm up. Number patterns with factors and multiples.
- Whisper the next number, then write it on your board.

DO:
- Display the three patterns.
- 90 seconds.

TEACHER NOTES:
Daily Review activates factor and multiple thinking - directly useful for today's equivalent fractions work.

WATCH FOR:
- Students who can identify the rule - they will pick up equivalent fractions fast.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check.
- 2, 4, 6, 8, 10, ... the rule is +2. Multiples of 2.
- 3, 6, 9, 12, 15, ... the rule is +3. Multiples of 3.
- 4, 8, 12, 16, 20, ... the rule is +4. Multiples of 4.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Multiples are the backbone of equivalent fractions. 2/4 = 4/8 = 6/12 because the denominators are multiples.

WATCH FOR:
- Students who name multiples confidently - secure.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency continues with decimals.
- Line up the decimal point.

DO:
- 60 seconds.

TEACHER NOTES:
Fluency continues - we keep the daily decimal workout even on a fraction day.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check.
- 1.4 plus 0.6 is 2.0 (or 2).
- 3.2 minus 1.8 is 1.4.
- 2.45 + 0.5 is 2.95.

DO:
- Tick and fix.

TEACHER NOTES:
Standard checkpoint.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention.
- We are learning to add fractions with related denominators by making the parts the same size.
- Now the success criteria.

DO:
- Choral read.
- Hold up two fraction strips: one thirds, one sixths.

TEACHER NOTES:
SC1 floor - add fractions with the SAME denominator. SC2 core target - add fractions when one denominator is a multiple of the other. SC3 stretches to explaining with a model.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me add 1/4 and 1/4.
- These are SAME denominator. The parts are the same size.
- I add the parts: 1 quarter plus 1 quarter is 2 quarters.
- 2/4 of a whole.
- Look at the fraction strip. Two quarters shaded.
- I can simplify 2/4 to 1/2. Both show the same shaded amount.

DO:
- Display the fraction strip showing 1/4 + 1/4.
- Point to the shaded parts as you add.

TEACHER NOTES:
First model uses same-denominator addition - already familiar. The strip makes "same size parts" visible. This sets up the contrast with different-denominator addition next.

MISCONCEPTIONS:
- Misconception: Add the numerators AND the denominators.
  Why: Students treat fractions like whole numbers.
  Impact: 1/4 + 1/4 becomes 2/8.
  Quick correction: "Add the parts only. Keep the SIZE of the part the same."

WATCH FOR:
- Students who say 2 quarters - secure.
- Students who say 2/8 - reteach with the strip.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now I add 1/2 and 1/4.
- The denominators are DIFFERENT. Halves and quarters are not the same SIZE part.
- I cannot add them as they are.
- Step 1: I find a common denominator. Quarters work for both - 1/2 is the same as 2/4.
- Step 2: rewrite. 1/2 = 2/4.
- Step 3: add. 2/4 + 1/4 = 3/4.
- Look at the fraction strips. Two quarters plus one quarter equals three quarters.

DO:
- Display the strips: 1/2 strip on top, 1/4 strip below, then 1/2 rewritten as 2/4.
- Point and trace.

TEACHER NOTES:
This is the threshold concept. To add fractions, the parts must be the same SIZE. The denominator names the size of the part. The trick: change the one that does not match.

MISCONCEPTIONS:
- Misconception: Just add 1/2 + 1/4 = 2/6 or 1 over 6.
  Why: Add numerator + numerator and denominator + denominator.
  Impact: Wildly wrong; smaller than 1/2.
  Quick correction: "The denominator names the SIZE of the part. We can only add same-size parts."

WATCH FOR:
- Students who rewrite 1/2 as 2/4 - secure on equivalent fractions.
- Students who add denominators - the classic trap. Reteach with the strip.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. Add 2/5 and 1/5.
- Same denominator. On your whiteboard.

DO:
- 30 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your answer on three, two, one, show."
- Scan for: 3/5.
PROCEED: If 80% have 3/5, click to reveal.
PIVOT: Most likely misconception - students added denominators to get 3/10.
- Reteach: "When the denominator is the same, just add the numerators."
- Re-check: "Try again."

TEACHER NOTES:
Same-denominator check before moving to harder.

WATCH FOR:
- Students who hold up 3/5 - secure.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Add 1/3 + 2/6.
- The denominators are different. But sixths is a multiple of thirds.
- Step 1: rewrite 1/3 with a denominator of 6.
- 1/3 = how many sixths? Use your fraction strips.
- Step 2: add.

DO:
- Display the strips.
- 90 seconds.
- Walk and listen for "1/3 equals 2/6."

TEACHER NOTES:
Related denominators where the LCM is the larger one. Students need to rewrite 1/3 as 2/6.

WATCH FOR:
- Pairs who write 1/3 = 2/6 - secure on equivalent fractions.
- Pairs who try to add 1+2 over 3+6 - reteach with the strip.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- 1/3 = 2/6 (we multiply top and bottom by 2).
- Now both fractions have sixths.
- 2/6 + 2/6 = 4/6.
- We can simplify 4/6 to 2/3 (divide both by 2).

DO:
- Click to reveal.
- Run the rewrite move once more.

TEACHER NOTES:
The reveal restates the rewrite and the simplification. Students who got 4/6 are on target. Students who simplified to 2/3 are stretching.

WATCH FOR:
- Students who self-correct - secure.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. Which of these is correct?
- A student wrote: 1/4 + 3/8 = 4/12.
- Use the fraction strips. Is 4/12 right?
- Thumbs UP for yes. DOWN for no.

DO:
- Display the work.
- 20 seconds.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Thumbs UP if correct. DOWN if not."
- Scan for: thumbs DOWN. The student added denominators.
PROCEED: If 80% thumbs down, click to reveal.
PIVOT: Most likely misconception - students agree with the student.
- Reteach: "Denominators name the SIZE. We cannot add the sizes."
- Re-check: "What is 1/4 written as eighths?"

TEACHER NOTES:
The hinge probes whether students see the denominator-adding trap. Correct: 1/4 = 2/8, so 2/8 + 3/8 = 5/8.

WATCH FOR:
- Confident thumbs down - they see the trap.
- Students unsure - reteach with the strip.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own. Take the practice sheet.
- Section 1: same denominator.
- Section 2: related denominators - rewrite first.
- Section 3: word problem.

DO:
- Distribute the practice sheet.
- Circulate.

TEACHER NOTES:
Section 2 is the core target. Watch students who consistently use the rewrite move.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use only Section 1 with the printed fraction strip reference.
- Extra Notes: Sit with these students. Use the strips for the first one.
EXTENDING PROMPT:
- Task: Section 4 - find the lowest common denominator before adding.
- Extra Notes: Connect to today's Daily Review (multiples).

WATCH FOR:
- Students who use the rewrite move - secure.
- Students who add denominators - reteach.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task.
- Question 1: 2/8 + 3/8.
- Question 2: 1/2 + 3/8. Rewrite first.

DO:
- Display.
- 3 minutes.

TEACHER NOTES:
Exit ticket assesses SC2 - adding with related denominators.

WATCH FOR:
- Students who got 5/8 and 7/8 - secure.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria. Thumbs for each.
- Turn and tell your partner: why do we rewrite fractions before adding?

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
The "same SIZE parts" idea is the threshold. Students who can say it are ready for Lesson 6.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, UNIT_TITLE, "Lesson 5: Adding fractions with related denominators",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  withReveal(
    () => dailyReviewSlide(pres,
      "Daily Review: Multiples — what comes next?",
      ["2, 4, 6, 8, ___", "3, 6, 9, 12, ___", "4, 8, 12, 16, ___"],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "10  (multiples of 2)   |   15  (multiples of 3)   |   20  (multiples of 4)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  withReveal(
    () => fluencySlide(pres, "Fluency: Adding & Subtracting Decimals",
      ["1.4 + 0.6", "3.2 − 1.8", "2.45 + 0.5"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "2.0     1.4     2.95", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  liSlide(pres,
    "We are learning to add fractions with related denominators by making the parts the same size.",
    [
      "I can add two fractions when they have the same denominator.",
      "I can add fractions when one denominator is a multiple of the other (related denominators).",
      "I can use a fraction strip or number line to explain why my answer is correct.",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do 1 — 1/4 + 1/4 (same denominator)
  workedExSlide(pres, 2, "I Do", "Add 1/4 + 1/4 (same denominator)",
    [
      "Same denominator → parts are the same SIZE.",
      "",
      "Add the NUMERATORS.",
      "Keep the denominator the same.",
      "",
      "1/4 + 1/4 = 2/4.",
      "",
      "Simplify: 2/4 = 1/2.",
      "",
      "Rule: only add same-size parts.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("Fraction strips", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      addFractionStripSet(slide,
        lg.rightX + 0.20, lg.panelTopPadded + 0.50, lg.rightW - 0.40, 1.80,
        [
          { denom: 4, shaded: 1, label: "1/4" },
          { denom: 4, shaded: 1, label: "1/4" },
          { denom: 4, shaded: 2, label: "2/4 = 1/2" },
        ],
        { labelW: 0.95, labelFontSize: 13 });
      slide.addText("1/4 + 1/4 = 2/4 (= 1/2)", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 2.50, w: lg.rightW - 0.20, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.SUCCESS, bold: true,
        align: "center", margin: 0,
      });
    }
  );

  // I Do 2 — 1/2 + 1/4 (related denominators)
  workedExSlide(pres, 2, "I Do", "Add 1/2 + 1/4 (related denominators)",
    [
      "Halves and quarters are DIFFERENT sizes.",
      "",
      "1.  Rewrite 1/2 so it has the same",
      "     denominator as 1/4.",
      "     1/2 = 2/4.",
      "",
      "2.  Now add. Same denominator.",
      "     2/4 + 1/4 = 3/4.",
      "",
      "Rule: change the denominators",
      "to match, THEN add.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY });
      slide.addText("1/2 rewritten as 2/4", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      addFractionStripSet(slide,
        lg.rightX + 0.20, lg.panelTopPadded + 0.50, lg.rightW - 0.40, 1.80,
        [
          { denom: 2, shaded: 1, label: "1/2" },
          { denom: 4, shaded: 2, label: "2/4" },
          { denom: 4, shaded: 3, label: "3/4" },
        ],
        { labelW: 0.95, labelFontSize: 13 });
      slide.addText("1/2 + 1/4 = 2/4 + 1/4 = 3/4", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 2.50, w: lg.rightW - 0.20, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.SUCCESS, bold: true,
        align: "center", margin: 0,
      });
    }
  );

  // CFU — 2/5 + 1/5
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Add 2/5 + 1/5", { color: C.ALERT });

      addTextOnShape(s, "✓ CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE },
        line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ALERT });
      s.addText([
        { text: "On your whiteboard:", options: { fontSize: 20, color: C.ALERT, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "2/5 + 1/5 =", options: { fontSize: 32, color: C.CHARCOAL, bold: true, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Same denominator.", options: { fontSize: 18, color: C.CHARCOAL, breakLine: true } },
        { text: "Add the numerators only.", options: { fontSize: 18, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.20, w: 4.0, h: SAFE_BOTTOM - CONTENT_TOP - 0.40,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.3, CONTENT_TOP, 4.2, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });
      s.addText("Fraction strip — fifths", {
        x: 5.3, y: CONTENT_TOP + 0.10, w: 4.2, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      addFractionStripSet(s,
        5.45, CONTENT_TOP + 0.55, 3.90, 2.00,
        [
          { denom: 5, shaded: 2, label: "2/5" },
          { denom: 5, shaded: 1, label: "1/5" },
          { denom: 5, shaded: 0, label: "? /5" },
        ],
        { labelW: 0.85, labelFontSize: 13 });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "2/5 + 1/5 = 3/5", {
        x: 1.0, y: 4.55, w: 8.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // We Do — 1/3 + 2/6
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Add 1/3 + 2/6",
      [
        "With your partner.",
        "",
        "Denominators 3 and 6.",
        "6 is a multiple of 3.",
        "",
        "1.  Rewrite 1/3 with a denominator of 6.",
        "     1/3 = ? /6.",
        "",
        "2.  Add.",
        "",
        "Use the fraction strip to check.",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY });
        slide.addText("1/3 and 2/6", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.28,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        addFractionStripSet(slide,
          lg.rightX + 0.20, lg.panelTopPadded + 0.50, lg.rightW - 0.40, 1.80,
          [
            { denom: 3, shaded: 1, label: "1/3" },
            { denom: 6, shaded: 2, label: "2/6" },
            { denom: 6, shaded: 0, label: "? /6" },
          ],
          { labelW: 0.95, labelFontSize: 13 });
        slide.addText("1/3 = ? /6", {
          x: lg.rightX + 0.10, y: lg.panelTopPadded + 2.50, w: lg.rightW - 0.20, h: 0.30,
          fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
      }),
    (slide) => {
      addTextOnShape(slide, "1/3 = 2/6, so 2/6 + 2/6 = 4/6  (= 2/3)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Hinge — 1/4 + 3/8 = 4/12 ?
  withReveal(
    () => cfuSlide(pres, "CFU", "Is this correct?", "Thumbs Up or Thumbs Down",
      "A student wrote:\n\n1/4 + 3/8 = 4/12\n\nUse the fraction strip to check.\nIs this correct?",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "NOT correct.   1/4 = 2/8, so 2/8 + 3/8 = 5/8 (not 4/12).", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.ALERT },
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
      { text: "Section 2 — related denominators (rewrite first).   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "Section 3 — word problem.", options: { fontSize: 18, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.45;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    s.addText("The key idea", {
      x: 0.7, y: panelY + 0.15, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });

    addTextOnShape(s, "Only add SAME-SIZE parts. Make the denominators match first.", {
      x: 0.8, y: panelY + 0.65, w: 8.4, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    addTextOnShape(s, "Related denominators: rewrite using equivalent fractions.", {
      x: 1.0, y: panelY + 1.20, w: 8.0, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  exitTicketSlide(pres,
    [
      "Add 2/8 + 3/8. Simplify if you can.",
      "Add 1/2 + 3/8. Rewrite 1/2 first, then add. Show your working.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: why do we rewrite fractions before adding?",
      scItems: [
        "I can add two fractions when they have the same denominator.",
        "I can add fractions when one denominator is a multiple of the other (related denominators).",
        "I can use a fraction strip or number line to explain why my answer is correct.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecFrac_Lesson5_Adding_Fractions.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Add fractions using equivalent fractions for related denominators.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "Rule: only add SAME-size parts. If the denominators are different, rewrite one fraction so both have the same denominator. Then add the numerators.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 — Same denominator", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   2/5 + 1/5    =   _______", y);
    y = addWriteLine(doc, "b)   3/8 + 2/8    =   _______", y);
    y = addWriteLine(doc, "c)   4/10 + 3/10  =   _______", y);
    y = addWriteLine(doc, "d)   5/12 + 4/12  =   _______", y);

    y = addSectionHeading(doc, "Section 2 — Related denominators (rewrite first)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   1/2 + 1/4   Rewrite 1/2 = ___/4.   Answer: _______", y);
    y = addWriteLine(doc, "b)   1/3 + 2/6   Rewrite 1/3 = ___/6.   Answer: _______", y);
    y = addWriteLine(doc, "c)   2/5 + 3/10  Rewrite 2/5 = ___/10.  Answer: _______", y);
    y = addWriteLine(doc, "d)   3/4 + 1/8   Rewrite 3/4 = ___/8.   Answer: _______", y);

    y = addSectionHeading(doc, "Section 3 — Word problem", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Liam used 1/4 of a cup of milk in his pancake recipe and another 3/8 of a cup for the topping.", y);
    y = addWriteLine(doc, "How much milk did Liam use in total? Answer: _______ cup", y);
    y = addWriteLine(doc, "Show your working: __________________________________________________", y);

    y = addSectionHeading(doc, "Section 4 — Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Find the lowest common denominator before adding. Show your steps.", y);
    y = addWriteLine(doc, "a)   1/4 + 5/12    LCD: ___    Answer: _______", y);
    y = addWriteLine(doc, "b)   2/3 + 1/9     LCD: ___    Answer: _______", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Adding Fractions Practice | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the adding fractions practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 — Same denominator", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3/5     b)  5/8     c)  7/10     d)  9/12 (= 3/4)", y);

    y = addSectionHeading(doc, "Section 2 — Related denominators", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  1/2 = 2/4, so 2/4 + 1/4 = 3/4.", y);
    y = addBodyText(doc, "b)  1/3 = 2/6, so 2/6 + 2/6 = 4/6 (= 2/3).", y);
    y = addBodyText(doc, "c)  2/5 = 4/10, so 4/10 + 3/10 = 7/10.", y);
    y = addBodyText(doc, "d)  3/4 = 6/8, so 6/8 + 1/8 = 7/8.", y);

    y = addSectionHeading(doc, "Section 3 — Word problem", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1/4 = 2/8. Total: 2/8 + 3/8 = 5/8 cup.", y);

    y = addSectionHeading(doc, "Section 4 — Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "a)  LCD = 12.   1/4 = 3/12, so 3/12 + 5/12 = 8/12 (= 2/3).", y);
    y = addBodyText(doc, "b)  LCD = 9.    2/3 = 6/9, so 6/9 + 1/9 = 7/9.", y);

    y = addTipBox(doc,
      "Watch for: students who add numerator and denominator (e.g., 1/2 + 1/4 = 2/6) - this is the classic trap. The denominator names the SIZE of the part. Reteach with fraction strips.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 5 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
