"use strict";

// Measurement Unit (Year 5/6 Numeracy) — Lesson 3: Converting between metric units of length.
// Year 6 content. Builds on Lessons 1-2: now we convert, not just choose.
// Daily Review: subtracting fractions.
// Fluency: subtraction.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(5));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  dailyReviewSlide, fluencySlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 3;
const TOTAL = 7;
const UNIT_TITLE = "Measurement: Units, Time and Timetables";
const FOOTER = `Measurement | Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/Meas_Lesson3_Converting_Length_Units";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Lesson 3 Length Conversions",
  "Convert between mm, cm, m and km. Use decimals where needed.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Lesson 3 Answer Key",
  "Worked answers for the Length Conversions sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome back. The last two days we picked the best unit.
- Today we convert one unit into another. Same length, different way of writing it.

DO:
- Settle students.
- Have rulers and the conversion staircase visible.

TEACHER NOTES:
Lesson 3 of 7. Year 6 content. The lesson teaches the conversion staircase (x10, x100, x1000) and decimal representations like 1.25 m = 125 cm.

WATCH FOR:
- Students who recall the small whole number rule from Lessons 1-2.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Today's materials.
- You will need your whiteboard and the Length Conversions sheet.

DO:
- Print one sheet per student.
- Print one answer key for yourself.
- Have rulers on desks.

TEACHER NOTES:
One student resource plus answer key. The conversion staircase visual is on the slides.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Subtracting fractions with the same denominator.
- Solve each on your whiteboard.

DO:
- Display the three prompts.
- 90 seconds.

TEACHER NOTES:
Subtraction mirrors yesterday's addition review. Same denominator means subtract the numerators.

WATCH FOR:
- Students who line up correctly - secure.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your work.
- 3/4 - 1/4 = 2/4 = 1/2.
- 7/8 - 3/8 = 4/8 = 1/2.
- 5/6 - 2/6 = 3/6 = 1/2.
- Tick or fix.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
Bonus: every answer simplifies to 1/2 today - a nice reveal for the class.

WATCH FOR:
- Students who simplify - secure.
- Students who stop at 2/4 etc - that is fine for SC1; prompt simplification as extension.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Subtraction facts.
- Quick whisper-answer then write on your board.

DO:
- Display the three prompts.
- 30 seconds.

TEACHER NOTES:
Subtraction fluency. Watch for students who count back one by one.

WATCH FOR:
- Students who use known facts (think addition) - secure.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 14 - 6 = 8.
- 15 - 7 = 8.
- 13 - 8 = 5.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Reinforce "think addition" for subtraction: 14 - 6 - what plus 6 makes 14?

WATCH FOR:
- Students who self correct - secure.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_VOCAB = `SAY:
- A quick reminder. Today we use the same four units, but we convert between them.
- "Convert" means write the same length in a different unit.
- For example: 200 cm and 2 m are the same length, just written two ways.

DO:
- Hold up the metre stick. Point to "200 cm".
- Point at the door and say "this is about 2 m".

TEACHER NOTES:
Hammer the idea that the length does not change. Only the unit (and so the number) changes.

WATCH FOR:
- Students who can say "the length is the same, only the number changes" - tracking.

[General: Key Vocabulary | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to convert between metric units of length.
- Now the success criteria.

DO:
- Choral read.

TEACHER NOTES:
SC1 is the bigger to smaller direction (multiply). SC2 is the core target - both directions. SC3 stretches to decimal answers (eg. 125 cm = 1.25 m).

WATCH FOR:
- Students who can repeat the SC - tracking.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. This is the conversion staircase.
- mm to cm: divide by 10. cm to mm: multiply by 10.
- cm to m: divide by 100. m to cm: multiply by 100.
- m to km: divide by 1000. km to m: multiply by 1000.
- Going DOWN the stairs (small unit to big unit): divide.
- Going UP the stairs (big unit to small unit): multiply.

DO:
- Point at each arrow.
- Repeat "down = divide" and "up = multiply" three times.
- Trace your hand down the stairs and up the stairs.

TEACHER NOTES:
The staircase is the core visual for this lesson. Going from a big unit to a small unit, we expect MORE of the small ones - so multiply. Going from a small unit to a big unit, we expect FEWER of the big ones - so divide. Say this language explicitly.

MISCONCEPTIONS:
- Misconception: Students multiply when they should divide (or vice versa).
  Why: They guess from the unit letters.
  Impact: Wildly wrong answers (eg. "3 km = 0.003 m").
  Quick correction: "Does the new number make sense? Should there be MORE or FEWER?"

WATCH FOR:
- Students who nod at "more or fewer" - tracking.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Watch me convert 1.25 m into cm.
- I am going from a BIG unit (m) to a SMALL unit (cm).
- Going up the stairs - multiply.
- The jump from m to cm is x100.
- 1.25 x 100 = 125. So 1.25 m = 125 cm.
- I check: does it make sense? Yes - we expect more centimetres than metres, and 125 > 1.25.

DO:
- Write "1.25 m" on the board.
- Trace the arrow up from m to cm with your finger.
- Write "x 100".
- Write "= 125 cm".
- Move the decimal point two places to the right as you say "x 100".

TEACHER NOTES:
This is the second core move. Decimal representations are exactly the same conversion - the rule never changes. Many students stall at 1.25 m because of the decimal; the trick is to move the decimal point. Show the trick AND the rule.

MISCONCEPTIONS:
- Misconception: Students drop the decimal point (eg. write 1.25 m = 125000 cm).
  Why: They confuse units and place value.
  Impact: Answers are 100x too big.
  Quick correction: "Multiply by 100 moves the decimal point TWO places. Not many places."

WATCH FOR:
- Students who can move the decimal point two places - tracking.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check.
- Convert 3.5 m into cm.
- On your whiteboard. Write your answer.
- Then check: more or fewer? Does it make sense?

DO:
- Display the prompt.
- 45 seconds.
- Walk and scan.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your boards on three, two, one, show."
- Scan for: 350 cm.
PROCEED: If 80% have 350, click to reveal and move to We Do.
PIVOT: Most likely misconception - students drop or misplace the decimal.
- Reteach: "Multiply by 100 moves the decimal TWO places to the right. 3.50 becomes 350."
- Re-check: "Convert 2.4 m to cm."

TEACHER NOTES:
Probe whether students apply the rule cleanly with a decimal start. 3.5 m = 350 cm.

WATCH FOR:
- Students who write 350 - secure.
- Students who write 35 or 3500 - dropped or misplaced the decimal.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me.
- Four conversions. Two going up the stairs, two going down.
- Partner talk for each: which direction (multiply or divide) and what number?

DO:
- Display the four conversions.
- 3 minutes total.
- Cold call after each.

TEACHER NOTES:
A balanced mix of directions and decimal answers.

WATCH FOR:
- Pairs who say "going up - multiply" - secure.
- Pairs who guess without saying the direction - prompt.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check together.
- 4 km = 4000 m (up, x1000).
- 250 cm = 2.5 m (down, divide by 100).
- 7 m = 700 cm (up, x100).
- 65 mm = 6.5 cm (down, divide by 10).

DO:
- Click to reveal.
- Run direction + operation for each.

TEACHER NOTES:
Reveal confirms direction + operation.

WATCH FOR:
- Students who got three of four - secure.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. Two students disagree about 4500 mm in metres.
- Ben says 4500 mm = 4.5 m. He divided by 1000.
- Mia says 4500 mm = 0.45 m. She divided by 10000.
- Thumbs UP for Ben. Thumbs DOWN for Mia.

DO:
- Display the disagreement.
- 15 seconds.

CFU CHECKPOINT:
Technique: Thumbs Up Down
Script:
- Say: "Thumbs UP for Ben. DOWN for Mia."
- Scan for: thumbs UP. Ben is right.
PROCEED: If 80% agree with Ben, click to reveal and confirm.
PIVOT: Most likely misconception - students multiply two jumps incorrectly.
- Reteach: "mm to cm is divide by 10. cm to m is divide by 100. Total: divide by 1000."
- Re-check: "Convert 3500 mm to m."

TEACHER NOTES:
The hinge probes two-step conversions (mm to m skips cm). The rule: combine the jumps.

WATCH FOR:
- Confident thumbs UP for Ben - they see the rule.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own.
- Take the Length Conversions sheet.
- Use the staircase. Check: more or fewer? Does it make sense?

DO:
- Distribute the sheet.
- Circulate. Listen for direction language.
- Cold call 1-2 students.

TEACHER NOTES:
Different numbers from the We Do. Same staircase.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use Part A only. Sit with these students. Do the first one together using the staircase.
- Extra Notes: Keep the staircase visible. Talk through direction first, then the multiplier.
EXTENDING PROMPT:
- Task: Part D (two-step conversions: mm to m, km to cm). Show the multiplier for each step.
- Extra Notes: Encourage explicit thinking about which direction.

WATCH FOR:
- Students who write the conversion and check direction - secure.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task. On your whiteboard.
- Convert 2.6 m into cm.
- Convert 350 cm into m.

DO:
- Display the prompt.
- 3 minutes.

TEACHER NOTES:
Exit ticket assesses SC2. Look for 260 cm and 3.5 m, with direction language if possible.

WATCH FOR:
- Students who get both right - secure.
- Students who get one - move them between groups tomorrow.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look at the success criteria again.
- Show me thumbs up, sideways, or down for each one.
- Turn and tell your partner: how do you know whether to multiply or divide?

DO:
- Read each I can statement.
- Use thumbs.

TEACHER NOTES:
The threshold idea is the staircase: up = multiply, down = divide. Tomorrow we apply the same rule to mass and capacity.

WATCH FOR:
- Strong thumbs up across all three - move at pace tomorrow.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Helpers ────────────────────────────────────────────────────────────────

// Draw the length conversion staircase.
function drawConversionStaircase(slide, x, y, w, h) {
  // Four steps from bottom-left (mm) to top-right (km). Each step is a rectangle.
  const units = [
    { code: "mm", color: C.PRIMARY },
    { code: "cm", color: C.SECONDARY },
    { code: "m",  color: C.ACCENT },
    { code: "km", color: C.ALERT },
  ];
  const stepCount = units.length;
  const stepW = w / stepCount;
  const stepH = h / stepCount;
  // Stairs go bottom-left -> top-right. Each step sits on top of the previous one.
  units.forEach((u, i) => {
    const sx = x + i * stepW;
    const sy = y + (stepCount - 1 - i) * stepH;
    // step rectangle
    slide.addShape("rect", {
      x: sx, y: sy, w: stepW, h: h - (stepCount - 1 - i) * stepH,
      fill: { color: u.color }, line: { color: C.CHARCOAL, width: 1 },
    });
    // step label
    slide.addText(u.code, {
      x: sx, y: sy + 0.06, w: stepW, h: 0.40,
      fontSize: 22, fontFace: FONT_H, color: C.WHITE,
      bold: true, align: "center", margin: 0,
    });
  });

  // Multipliers between steps — placed INSIDE each step (top half) so they
  // never collide with elements sitting above or below the staircase.
  const jumps = ["x10", "x100", "x1000"];
  jumps.forEach((j, i) => {
    const jx = x + (i + 0.5) * stepW;
    const jy = y + (stepCount - 1 - i) * stepH + 0.48;
    slide.addText(j, {
      x: jx - 0.40, y: jy, w: 0.80, h: 0.24,
      fontSize: 11, fontFace: FONT_B, color: C.WHITE,
      bold: true, align: "center", margin: 0,
    });
  });
}

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 3: Converting between length units",
    `Year 5/6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal — subtracting fractions
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Subtracting fractions",
      ["3/4 - 1/4 =", "7/8 - 3/8 =", "5/6 - 2/6 ="],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "2/4 = 1/2     4/8 = 1/2     3/6 = 1/2", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal — subtraction
  withReveal(
    () => fluencySlide(pres, "Fluency: Subtraction",
      ["14 - 6", "15 - 7", "13 - 8"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "8     8     5", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Key vocabulary — converting means writing the same length in a different unit
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Vocabulary", { color: C.PRIMARY });
    addTitle(s, "Convert = same length, different unit");

    addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: C.PRIMARY });

    // Big examples
    addTextOnShape(s, "200 cm  =  2 m", {
      x: 1.0, y: CONTENT_TOP + 0.30, w: 8.0, h: 0.80, rectRadius: 0.10,
      fill: { color: C.PRIMARY },
    }, { fontSize: 34, fontFace: FONT_H, color: C.WHITE, bold: true });

    addTextOnShape(s, "1.25 m  =  125 cm", {
      x: 1.0, y: CONTENT_TOP + 1.25, w: 8.0, h: 0.80, rectRadius: 0.10,
      fill: { color: C.SECONDARY },
    }, { fontSize: 34, fontFace: FONT_H, color: C.WHITE, bold: true });

    addTextOnShape(s, "3 km  =  3000 m", {
      x: 1.0, y: CONTENT_TOP + 2.20, w: 8.0, h: 0.80, rectRadius: 0.10,
      fill: { color: C.ACCENT },
    }, { fontSize: 34, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to convert between metric units of length.",
    [
      "I can multiply by 10, 100 or 1000 to convert from a big unit to a smaller unit.",
      "I can divide by 10, 100 or 1000 to convert from a small unit to a bigger unit.",
      "I can convert lengths written as decimals (e.g. 1.25 m = 125 cm) and check the answer makes sense.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: I Do (1) — the staircase
  workedExSlide(pres, 2, "I Do", "The conversion staircase",
    [
      "Going UP the stairs:",
      "big unit -> small unit",
      "We expect MORE.",
      "So we MULTIPLY.",
      "",
      "Going DOWN the stairs:",
      "small unit -> big unit",
      "We expect FEWER.",
      "So we DIVIDE.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.PRIMARY });
      slide.addText("Length staircase", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });

      drawConversionStaircase(slide,
        lg.rightX + 0.30, lg.panelTopPadded + 0.65,
        lg.rightW - 0.60, 1.90);
    }
  );

  // Slide 10: I Do (2) — worked example 1.25 m to cm
  workedExSlide(pres, 2, "I Do", "Convert 1.25 m into cm",
    [
      "Direction: big to small (m to cm).",
      "Going up the stairs - MULTIPLY.",
      "",
      "The jump from m to cm is x100.",
      "1.25 x 100 = 125.",
      "",
      "1.25 m = 125 cm.",
      "",
      "Check: more cm than m? Yes. Makes sense.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.SECONDARY });
      slide.addText("1.25 m  =  125 cm", {
        x: lg.rightX + 0.10, y: lg.panelTopPadded + 0.20, w: lg.rightW - 0.20, h: 0.55,
        fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      // Show the decimal point movement
      const moveY = lg.panelTopPadded + 0.95;
      slide.addText("1 . 2 5", {
        x: lg.rightX + 0.10, y: moveY, w: lg.rightW - 0.20, h: 0.45,
        fontSize: 26, fontFace: FONT_H, color: C.CHARCOAL,
        bold: true, align: "center", margin: 0,
      });
      slide.addText("x 100  ->  move the decimal", {
        x: lg.rightX + 0.10, y: moveY + 0.50, w: lg.rightW - 0.20, h: 0.30,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED,
        italic: true, align: "center", margin: 0,
      });
      slide.addText("1 2 5 .", {
        x: lg.rightX + 0.10, y: moveY + 0.90, w: lg.rightW - 0.20, h: 0.45,
        fontSize: 26, fontFace: FONT_H, color: C.SECONDARY,
        bold: true, align: "center", margin: 0,
      });

      addTextOnShape(slide, "Decimal moves 2 places right.", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.55, w: lg.rightW - 0.30, h: 0.45, rectRadius: 0.08,
        fill: { color: C.ALERT },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 11-12: CFU + reveal — convert 3.5 m to cm
  withReveal(
    () => cfuSlide(pres, "CFU", "Convert 3.5 m into cm", "Show Me Boards",
      "On your whiteboard:\n\nWrite the conversion of 3.5 m into cm.\n\nWhich direction (up or down)?  Multiply or divide?  By how much?",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "3.5 m = 350 cm   (UP, x100, decimal moves 2 right)", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 13-14: We Do + reveal — four conversions
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Convert each one with your partner", { color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9.0, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });

      const items = [
        { from: "4 km", arrow: "->", to: "? m" },
        { from: "250 cm", arrow: "->", to: "? m" },
        { from: "7 m", arrow: "->", to: "? cm" },
        { from: "65 mm", arrow: "->", to: "? cm" },
      ];

      const cellW = 2.0;
      const gap = 0.15;
      const totalW = cellW * 4 + gap * 3;
      const startX = (10 - totalW) / 2;
      const cardY = CONTENT_TOP + 0.30;
      const cardH = 1.50;

      items.forEach((it, i) => {
        const cx = startX + i * (cellW + gap);
        addCard(s, cx, cardY, cellW, cardH, { strip: C.SECONDARY, fill: C.WHITE });
        s.addText(it.from, {
          x: cx, y: cardY + 0.18, w: cellW, h: 0.40,
          fontSize: 20, fontFace: FONT_H, color: C.PRIMARY,
          bold: true, align: "center", margin: 0,
        });
        s.addText(it.arrow, {
          x: cx, y: cardY + 0.62, w: cellW, h: 0.30,
          fontSize: 16, fontFace: FONT_B, color: C.MUTED,
          align: "center", margin: 0,
        });
        s.addText(it.to, {
          x: cx, y: cardY + 0.95, w: cellW, h: 0.40,
          fontSize: 20, fontFace: FONT_H, color: C.SECONDARY,
          bold: true, align: "center", margin: 0,
        });
      });

      s.addText("Partner talk: which direction? What multiplier? Then write the answer.", {
        x: 0.7, y: SAFE_BOTTOM - 0.55, w: 8.6, h: 0.30,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, align: "center", margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "4 km = 4000 m   |   250 cm = 2.5 m   |   7 m = 700 cm   |   65 mm = 6.5 cm", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slides 15-16: CFU hinge + reveal — Ben vs Mia
  withReveal(
    () => cfuSlide(pres, "CFU", "Who is right?", "Thumbs Up or Thumbs Down",
      "Convert 4500 mm into m.\n\nBen says 4500 mm = 4.5 m. He divided by 1000.\n\nMia says 4500 mm = 0.45 m. She divided by 10000.\n\nThumbs UP for Ben.   Thumbs DOWN for Mia.",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Ben is right. mm to cm is ÷10. cm to m is ÷100. Total: ÷1000.", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 17: You Do — Length Conversions sheet
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: Length Conversions sheet", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.30, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "use the staircase to choose multiply or divide.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "do the conversion.   ", options: { fontSize: 18, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 20, color: C.ALERT, bold: true } },
      { text: "check: more or fewer?", options: { fontSize: 18, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.18, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.45;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });

    s.addText("UP the staircase = MULTIPLY     |     DOWN the staircase = DIVIDE", {
      x: 0.7, y: panelY + 0.10, w: 8.6, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });

    drawConversionStaircase(s, 2.5, panelY + 0.60, 5.0, 1.50);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // Slide 18: Exit Ticket
  exitTicketSlide(pres,
    [
      "Convert 2.6 m into cm. Show your direction and the multiplier.",
      "Convert 350 cm into m. Show your direction and the divider.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 19: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: how do you know whether to multiply or divide?",
      scItems: [
        "I can multiply by 10, 100 or 1000 to convert from a big unit to a smaller unit.",
        "I can divide by 10, 100 or 1000 to convert from a small unit to a bigger unit.",
        "I can convert lengths written as decimals (e.g. 1.25 m = 125 cm) and check the answer makes sense.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Meas_Lesson3_Converting_Length_Units.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Use the staircase. Up = multiply. Down = divide.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });
    y = addTipBox(doc,
      "mm to cm = ÷10. cm to m = ÷100. m to km = ÷1000. Going the other way, multiply. Check: more or fewer? Does it make sense?",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Part A — Multiply (big to small)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  3 m = ___ cm", y);
    y = addWriteLine(doc, "b)  5 km = ___ m", y);
    y = addWriteLine(doc, "c)  4 cm = ___ mm", y);
    y = addWriteLine(doc, "d)  6.5 m = ___ cm", y);

    y = addSectionHeading(doc, "Part B — Divide (small to big)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  400 cm = ___ m", y);
    y = addWriteLine(doc, "b)  8000 m = ___ km", y);
    y = addWriteLine(doc, "c)  90 mm = ___ cm", y);
    y = addWriteLine(doc, "d)  175 cm = ___ m", y);

    y = addSectionHeading(doc, "Part C — Mix (multiply or divide?)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  2.4 km = ___ m", y);
    y = addWriteLine(doc, "b)  120 mm = ___ cm", y);
    y = addWriteLine(doc, "c)  0.75 m = ___ cm", y);
    y = addWriteLine(doc, "d)  450 cm = ___ m", y);

    y = addSectionHeading(doc, "Part D — Extension: two-step conversions", y, { color: C.ACCENT });
    y = addBodyText(doc, "Show your two steps and the total multiplier or divider.", y);
    y = addWriteLine(doc, "a)  2 km = ___ cm   (Step 1: km to m. Step 2: m to cm.)", y);
    y = addWriteLine(doc, "b)  4500 mm = ___ m   (Step 1: mm to cm. Step 2: cm to m.)", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Length Conversions | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Length Conversions sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`,
    });

    y = addSectionHeading(doc, "Part A — Multiply (big to small)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  3 m = 300 cm    (x100)", y);
    y = addBodyText(doc, "b)  5 km = 5000 m    (x1000)", y);
    y = addBodyText(doc, "c)  4 cm = 40 mm    (x10)", y);
    y = addBodyText(doc, "d)  6.5 m = 650 cm    (x100; decimal moves 2 right)", y);

    y = addSectionHeading(doc, "Part B — Divide (small to big)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  400 cm = 4 m    (÷100)", y);
    y = addBodyText(doc, "b)  8000 m = 8 km    (÷1000)", y);
    y = addBodyText(doc, "c)  90 mm = 9 cm    (÷10)", y);
    y = addBodyText(doc, "d)  175 cm = 1.75 m    (÷100; decimal moves 2 left)", y);

    y = addSectionHeading(doc, "Part C — Mix", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  2.4 km = 2400 m    (x1000)", y);
    y = addBodyText(doc, "b)  120 mm = 12 cm    (÷10)", y);
    y = addBodyText(doc, "c)  0.75 m = 75 cm    (x100; decimal moves 2 right)", y);
    y = addBodyText(doc, "d)  450 cm = 4.5 m    (÷100)", y);

    y = addSectionHeading(doc, "Part D — Two-step conversions", y, { color: C.ACCENT });
    y = addBodyText(doc, "a)  2 km = 200 000 cm    (x1000 then x100; total x100 000).", y);
    y = addBodyText(doc, "b)  4500 mm = 4.5 m    (÷10 then ÷100; total ÷1000).", y);

    y = addTipBox(doc,
      "Watch for: students who multiply when they should divide (check 'more or fewer'); students who move the decimal too many places (x100 = 2 places, not 3); students who skip the check and end up off by a factor of 10.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 5/6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson 3 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
