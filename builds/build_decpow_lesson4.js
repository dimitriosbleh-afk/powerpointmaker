"use strict";

// Decimal Operations & Place Value Mastery (Year 6 Numeracy) - Lesson 4 of 4.
// Estimation in multiplicative situations: a decimal greater than 1 multiplied
// by a two- or three-digit number. Round each to a friendly number (a whole
// number and a multiple of 10 or 100), multiply mentally, and judge whether the
// estimate is a bit over or under. VC2M6N06.
// Daily Review: Shapes (prior). Fluency: lattice multiplication.
// Unit variant fixed (variant 2) for cohesion.
// Catch-up: launch re-activates rounding to friendly numbers; the worksheet has
// an enabling rounding scaffold. No session assumes the one before it.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(3)); // variant 2, fixed for the unit
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  fluencySlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 4;
const TOTAL = 4;
const UNIT_TITLE = "Decimal Operations and Place Value";
const FOOTER = `Decimal Operations | Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/DecPow_Lesson4_Estimation_Strategies";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Estimate decimal multiplications by rounding to friendly numbers.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Lesson 4 practice sheet.");
const EXT_RES = makeSessionResource(SESSION,
  "Year 8 Extension",
  "Refining estimates, over and under, and percentage error.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXT_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to the last lesson in our decimal operations unit.
- We have learned to multiply and divide decimals by powers of 10 and by multiples of them.
- Today we use rounding to estimate when a situation only needs a sensible close answer, not an exact one.

DO:
- Have whiteboards and markers ready.
- Settle the room and focus attention before you click past the title.

TEACHER NOTES:
Lesson 4 of 4. It draws the unit together. The launch rebuilds rounding, so a student who missed earlier lessons can still access estimation today.

WATCH FOR:
- Students who already round confidently - strong start.
- Students who look unsure - reassure, we build it together.

[General: Title | Element: Planning and Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The practice sheet is for the You Do near the end. The Year 8 extension is there for anyone ready for more.

DO:
- Print one practice sheet and one answer key per student.
- Have whiteboards and markers ready.
- Keep the Year 8 extension on hand for early finishers.

TEACHER NOTES:
One student practice sheet, an answer key, and a Year 8 extension. Most teaching happens on whiteboards.

CATCH-UP NOTE:
A student who missed earlier sessions can still do today. Estimation only needs rounding and an easy multiplication, both rebuilt in the launch. The worksheet has a rounding scaffold so a returner can start straight away.

[General: Resources | Element: Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. We are revisiting shapes from earlier work.
- Find the area and perimeter of the rectangle, the area of the triangle, and the side length of the square.
- Work each on your whiteboard.

DO:
- Display the three shapes with their measurements.
- Give about 2 minutes.
- Walk and listen for correct formulas and units.

TEACHER NOTES:
Daily Review is prior learning - shapes - not today's estimation. Area of a triangle is half base times height. Square side is perimeter divided by four.

WATCH FOR:
- Students who use square centimetres for area and centimetres for length - secure.
- Students who forget the half for the triangle - prompt the formula.

[Stage 1: Daily Review | Element: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- The rectangle: area is 8 times 5, which is 40 square centimetres. Perimeter is 8 plus 5 plus 8 plus 5, which is 26 centimetres.
- The triangle: half of 10 times 6, which is 30 square centimetres.
- The square: perimeter 24 divided by 4 is 6 centimetres a side.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The half on the triangle and the units are the usual slips. Note students who give area without square units.

WATCH FOR:
- Students who self-correct quickly - secure.

[Stage 1: Daily Review Answers | Element: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Lattice multiplication again.
- Draw your lattice grid, multiply each pair of digits, then add along the diagonals.
- Carry where you need to.

DO:
- Display the three prompts.
- Give about 2 to 3 minutes.
- Scan for neat lattice grids and correct diagonal sums.

TEACHER NOTES:
Same lattice method as last lesson. These products connect to today's estimation - once you have an estimate, the lattice gives the exact answer to compare against.

WATCH FOR:
- Students who set out a clean lattice - secure.
- Students who miss a diagonal carry - prompt at the ten boundary.

[Stage 1: Fluency | Element: Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 39 times 21 is 819.
- 48 times 32 is 1536.
- 19 times 43 is 817.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
Notice 39 times 21 is about 40 times 20, which is 800 - close to the exact 819. That is exactly the estimation idea we teach today.

WATCH FOR:
- Students who self-correct - secure.
- Students whose diagonals do not line up - small group focus.

[Stage 1: Fluency Answers | Element: Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Some of you may remember rounding to a friendly number.
- 4.2 is about 4. 19 is about 20.
- Friendly numbers are easy to multiply in your head: 4 times 20 is 80.
- Today we use rounding to estimate decimal multiplications when we only need a close answer.

DO:
- Point to each rounding, then the easy multiplication.
- Have students chorus 'round, round, then multiply'.
- Bridge: 'an estimate is a sensible close answer, not the exact one'.

TEACHER NOTES:
This launch re-activates rounding and easy multiplication. It is the catch-up bridge - estimation needs only rounding and a friendly product, both modelled here.

WATCH FOR:
- Students who round both numbers sensibly - strong prior knowledge.
- Students unsure - reassure, we build it together this lesson.

[Stage: Launch | Element: Knowledge and Memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to estimate answers to decimal multiplications by rounding to friendly numbers.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.

TEACHER NOTES:
The first criterion is reachable for everyone - round a decimal and a whole number. The second is the core target the exit ticket checks. The third stretches to judging whether the estimate is over or under.

WATCH FOR:
- Students who can repeat the language - tracking.

[General: LI/SC | Element: Planning]`;

const NOTES_VOCAB = `SAY:
- Three words for today.
- An estimate is a sensible close answer, not the exact one.
- To round means to change a number to a friendly number nearby.
- A friendly number is a whole number, or a multiple of 10 or 100, that is easy to multiply.

DO:
- Point to each word as you say it.
- Have students round 5.9 and 41 to friendly numbers out loud.

TEACHER NOTES:
Vocabulary comes after the learning intention. The key idea is that estimating is a choice we make when a situation only needs a close answer.

WATCH FOR:
- Students who round both numbers - secure.
- Students who think estimate means guess - clarify it is a reasoned close answer.

[General: Key Vocabulary | Element: Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us look at this one together. About how much is 4.2 times 19?
- The situation only needs a close answer, so I estimate.
- I round 4.2 down to 4, and 19 up to 20.
- Now I multiply the friendly numbers: 4 times 20 is 80.
- So 4.2 times 19 is about 80. I rounded one number down and one up, so my estimate is very close.

DO:
- Round each number out loud, then multiply.
- Say 'about 80', not 'equals 80'.
- Note that one number went down and one went up.

TEACHER NOTES:
This is the core estimation move. Stress that an estimate uses 'about'. Rounding one number up and one down keeps the estimate close.

MISCONCEPTIONS:
- Misconception: students round only the decimal and not the multiplier.
  Why: they focus on the decimal as the tricky part.
  Impact: 4 times 19 is harder than 4 times 20 and not really an estimate.
  Quick correction: round BOTH numbers to friendly numbers.

WATCH FOR:
- Students who round both and say 'about' - secure.
- Students who try to be exact - remind them the situation only needs a close answer.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now a bigger one. About how much is 3.8 times 312?
- I round 3.8 up to 4, and 312 down to 300, a friendly multiple of 100.
- Now I multiply: 4 times 300 is 1200.
- So 3.8 times 312 is about 1200.
- If I work out the exact answer with the lattice later, I expect something close to 1200 - that is how I check it is sensible.

DO:
- Round each number, then multiply using the times 100 move.
- Connect to fluency: the estimate checks the exact lattice answer.
- Say 'about 1200'.

TEACHER NOTES:
With a three-digit number, round to a multiple of 100. The estimate is a sense-check for the exact product. The exact answer here is 1185.6, close to 1200.

MISCONCEPTIONS:
- Misconception: students round 312 to 310 and lose the easy multiplication.
  Why: they round to the nearest ten out of habit.
  Impact: 4 times 310 is not a friendly mental product.
  Quick correction: for a three-digit number, round to the nearest hundred so the product is easy.

WATCH FOR:
- Students who round to 300 and multiply easily - secure.
- Students who keep too much detail - steer to the friendly hundred.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Estimate 5.1 times 39.
- Show me your two rounded numbers and your estimate.

DO:
- Display the prompt.
- Give 60 seconds.
- Walk and scan for both roundings.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me on three, two, one, show.
- Scan for: 5.1 rounds to 5, 39 rounds to 40, 5 times 40 is 200. About 200.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students round only 5.1 and try 5 times 39.
- Reteach: round BOTH numbers. 5.1 to 5, 39 to 40. Then 5 times 40 is 200.
- Re-check: what did you round each number to, and what is the estimate?

TEACHER NOTES:
The trap is rounding only one number. Make students show both rounded numbers before the product.

WATCH FOR:
- Students who show 5 and 40 and write about 200 - secure.
- Students who keep 39 - prompt them to round the multiplier too.

[Stage 2: CFU | Element: Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Estimate 6.1 times 48.
- Round each number to a friendly number, then multiply.
- Whisper to your partner: what did you round each one to?

DO:
- Display 6.1 times 48 above the estimate frame.
- Give 75 seconds.
- Listen for '6 times 50 is 300'.

TEACHER NOTES:
Same move as the I Do with new numbers. 6.1 rounds to 6, 48 rounds to 50.

WATCH FOR:
- Pairs who estimate about 300 - secure.
- Pairs who round 48 to 40 - prompt them to round to the nearer ten, 50.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- 6.1 rounds to 6. 48 rounds to 50.
- 6 times 50 is 300.
- So 6.1 times 48 is about 300. Both numbers rounded up a little, so the exact answer is a bit less.

DO:
- Click to reveal.
- Note that both rounded up, so the estimate is a little high.

TEACHER NOTES:
The exact answer is 292.8, just under 300, because both numbers rounded up. That is a good 'is it over or under' discussion.

WATCH FOR:
- Students who self-correct - secure.
- Students who rounded 48 down to 40 - remind them 48 is nearer 50.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, a money problem.
- A ticket costs $5.80. About how much for 195 tickets?
- The word 'about' tells us an estimate is fine. Round and multiply.

DO:
- Display the problem above the estimate frame.
- Give 90 seconds.
- Listen for '6 times 200 is 1200, about 1200 dollars'.

TEACHER NOTES:
This is the three-digit, real-context case. Round 5.80 to 6 and 195 to 200, a friendly multiple of 100. The estimate is about 1200 dollars.

WATCH FOR:
- Students who estimate about $1200 - secure.
- Students who round 195 to 190 - prompt them to round to the friendly 200.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- 5.80 rounds to 6. 195 rounds to 200.
- 6 times 200 is 1200.
- So 195 tickets cost about 1200 dollars. Both numbers rounded up, so the real cost is a bit less.

DO:
- Click to reveal.
- Note both rounded up, so the estimate is a little high.

TEACHER NOTES:
The exact cost is 1131 dollars, under the estimate because both rounded up. An estimate of about 1200 dollars is plenty to plan a budget.

WATCH FOR:
- Students who estimate about $1200 and reason over or under - ready for independent work.
- Students who try to find the exact cost - remind them the question only needs an estimate.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- Section 1 is straight estimates. Section 2 has worded situations.
- Section 3 asks you to decide if each estimate is a bit over or under. If you finish, try the extension box.

DO:
- Distribute the practice sheet.
- Circulate and check both numbers are rounded.
- Cold call one or two students to justify an over or under.

TEACHER NOTES:
Different numbers from the We Do, same move: round both to friendly numbers, multiply, say 'about'.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the rounding boxes at the top of the sheet. Round each number first, then multiply the friendly numbers.
- Extra Notes: Sit with these students and round the first pair together. This is also the rebuild for any returning student.
EXTENDING PROMPT:
- Task: The extension box asks students to decide whether an estimate is over or under before checking, and to estimate a three-step shopping total.
- Extra Notes: Students who are ready can move on to the Year 8 extension sheet.

WATCH FOR:
- Students who round both and use 'about' - secure.
- Students who round only one number - prompt for the other.

[Stage 4: You Do | Element: Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard, on your own.
- Estimate 7.9 times 41 by rounding to friendly numbers.
- Show your two rounded numbers and your estimate.

DO:
- Display the prompt.
- Give 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core success criterion - estimate a decimal times a two-digit number using friendly numbers. Look for 8 times 40 is 320, about 320.

WATCH FOR:
- Students who show 8 and 40 and write about 320 - secure.
- Students who keep 7.9 or 41 - they did not round, revisit rounding in the next unit.

[Stage 5: Exit Ticket | Element: Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: when is an estimate good enough instead of an exact answer?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is that estimating means rounding both numbers to friendly numbers and multiplying, and that we choose it when a situation only needs a close answer. This closes the unit: powers of 10, multiples of them, and now estimation.

WATCH FOR:
- Strong thumbs up across all three - the unit has landed.
- Sideways or down on the core criterion - small group revision in the next unit.

[General: Closing | Element: Retention and Recall]`;

// --- Helpers -----------------------------------------------------------------

// Estimation panel used on the I Do slides: round each factor, then multiply.
function addEstimatePanel(slide, lg, cfg) {
  const color = cfg.color || C.PRIMARY;
  const cardH = cfg.cardH || 3.35;
  const top = lg.panelTopPadded;
  const cx = lg.rightX + 0.25;
  const cw = lg.rightW - 0.50;
  addCard(slide, lg.rightX, top, lg.rightW, cardH, { strip: color });

  slide.addText("Round to friendly numbers", {
    x: lg.rightX + 0.15, y: top + 0.10, w: lg.rightW - 0.30, h: 0.32,
    fontSize: 14.5, fontFace: FONT_H, color: color, bold: true,
    align: "center", margin: 0,
  });

  const halfW = (cw - 0.20) / 2;
  addTextOnShape(slide, cfg.round1, {
    x: cx, y: top + 0.50, w: halfW, h: 0.58, rectRadius: 0.08,
    fill: { color: C.PRIMARY },
  }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
  addTextOnShape(slide, cfg.round2, {
    x: cx + halfW + 0.20, y: top + 0.50, w: halfW, h: 0.58, rectRadius: 0.08,
    fill: { color: C.PRIMARY },
  }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

  slide.addText("now multiply the friendly numbers", {
    x: cx, y: top + 1.16, w: cw, h: 0.28,
    fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
    align: "center", margin: 0,
  });

  addTextOnShape(slide, cfg.friendly, {
    x: cx, y: top + 1.46, w: cw, h: 0.60, rectRadius: 0.08,
    fill: { color: C.SECONDARY },
  }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });

  addTextOnShape(slide, cfg.result, {
    x: cx, y: top + 2.22, w: cw, h: 0.62, rectRadius: 0.08,
    fill: { color: C.SUCCESS },
  }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });

  if (cfg.note) {
    slide.addText(cfg.note, {
      x: lg.rightX + 0.15, y: top + 2.90, w: lg.rightW - 0.30, h: 0.32,
      fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
      align: "center", margin: 0,
    });
  }
}

// Blank estimate frame for the We Do slides (students fill the rounding).
function addEstimateBlank(slide, lg, headline, subline) {
  const top = lg.panelTopPadded;
  const cx = lg.rightX + 0.25;
  const cw = lg.rightW - 0.50;
  addCard(slide, lg.rightX, top, lg.rightW, 2.8, { strip: C.SECONDARY });
  slide.addText(headline, {
    x: lg.rightX + 0.2, y: top + 0.12, w: lg.rightW - 0.4, h: 0.42,
    fontSize: 26, fontFace: FONT_H, color: C.SECONDARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
  slide.addText(subline, {
    x: lg.rightX + 0.2, y: top + 0.62, w: lg.rightW - 0.4, h: 0.28,
    fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
    align: "center", margin: 0,
  });
  const halfW = (cw - 0.20) / 2;
  addTextOnShape(slide, "round -> ____", {
    x: cx, y: top + 1.02, w: halfW, h: 0.56, rectRadius: 0.08,
    fill: { color: C.PRIMARY },
  }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
  addTextOnShape(slide, "round -> ____", {
    x: cx + halfW + 0.20, y: top + 1.02, w: halfW, h: 0.56, rectRadius: 0.08,
    fill: { color: C.PRIMARY },
  }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
  addTextOnShape(slide, "____ x ____ = ____", {
    x: cx, y: top + 1.70, w: cw, h: 0.56, rectRadius: 0.08,
    fill: { color: C.PRIMARY },
  }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
}

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 4: Estimating decimal multiplications",
    `Year 6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - Shapes (custom visual)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Daily Review: Shapes", { color: STAGE_COLORS["1"] });

      const cardTop = CONTENT_TOP;
      const cardH = 3.05;
      addCard(s, 0.5, cardTop, 9.0, cardH, { strip: STAGE_COLORS["1"] });

      const shapeY = cardTop + 0.30;
      // Rectangle
      s.addShape("rect", {
        x: 1.15, y: shapeY, w: 1.75, h: 1.05,
        fill: { color: C.PRIMARY }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("8 cm by 5 cm", {
        x: 0.85, y: shapeY + 1.15, w: 2.35, h: 0.32,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });
      // Triangle
      s.addShape("triangle", {
        x: 4.25, y: shapeY, w: 1.5, h: 1.05,
        fill: { color: C.SECONDARY }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("base 10 cm, height 6 cm", {
        x: 3.65, y: shapeY + 1.15, w: 2.7, h: 0.32,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });
      // Square
      s.addShape("rect", {
        x: 7.35, y: shapeY, w: 1.05, h: 1.05,
        fill: { color: C.ACCENT }, line: { color: C.CHARCOAL, width: 1 },
      });
      s.addText("perimeter 24 cm", {
        x: 6.7, y: shapeY + 1.15, w: 2.4, h: 0.32,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });

      addTextOnShape(s, "Find:  rectangle area + perimeter    triangle area    square side", {
        x: 0.85, y: cardTop + cardH - 0.62, w: 8.3, h: 0.46, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, { fontSize: 14.5, fontFace: FONT_H, color: C.WHITE, bold: true });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      runSlideDiagnostics(s, pres);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Rectangle 40 cm² and 26 cm     Triangle 30 cm²     Square 6 cm", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - lattice multiplication
  withReveal(
    () => fluencySlide(pres, "Fluency: Lattice multiplication",
      ["39 x 21", "48 x 32", "19 x 43"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "819        1536        817", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - rounding to friendly numbers (catch-up bridge)
  contentSlide(pres, "Launch", C.ACCENT, "Friendly numbers are easy to multiply",
    [
      "Some of you may remember rounding.",
      "4.2 is about 4.   19 is about 20.",
      "",
      "4 x 20 = 80 is easy in your head.",
      "Today: estimate decimal multiplications.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.ACCENT });
      slide.addText("Round, round, multiply", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });
      const chips = [["4.2 -> 4", "round down"], ["19 -> 20", "round up"]];
      const cy0 = lg.panelTopPadded + 0.58;
      chips.forEach((r, i) => {
        const cy = cy0 + i * 0.62;
        addTextOnShape(slide, r[0], {
          x: lg.rightX + 0.25, y: cy, w: 2.0, h: 0.50, rectRadius: 0.07,
          fill: { color: C.PRIMARY },
        }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(r[1], {
          x: lg.rightX + 2.35, y: cy, w: lg.rightW - 2.55, h: 0.50,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", valign: "middle", margin: 0,
        });
      });
      addTextOnShape(slide, "4 x 20 = 80   (about 80)", {
        x: lg.rightX + 0.25, y: cy0 + 1.28, w: lg.rightW - 0.50, h: 0.52, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to estimate answers to decimal multiplications by rounding to friendly numbers.",
    [
      "I can round a decimal and a whole number to friendly numbers.",
      "I can estimate a decimal times a 2 or 3 digit number using a multiple of 10 or 100.",
      "I can explain whether my estimate is a bit over or a bit under, and why.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Words for today",
    [
      "Estimate = a sensible close answer, not exact",
      "Round = change to a friendly number nearby",
      "Friendly number = a whole number or a multiple of 10 or 100",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.SECONDARY });
      slide.addText("Round to friendly", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      const rows = [["4.2", "4"], ["19", "20"], ["312", "300"]];
      const ry0 = lg.panelTopPadded + 0.58;
      rows.forEach((r, i) => {
        const ry = ry0 + i * 0.56;
        addTextOnShape(slide, r[0], {
          x: lg.rightX + 0.5, y: ry, w: 1.3, h: 0.46, rectRadius: 0.06,
          fill: { color: C.PRIMARY },
        }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText("->", {
          x: lg.rightX + 1.85, y: ry, w: 0.5, h: 0.46,
          fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
        addTextOnShape(slide, r[1], {
          x: lg.rightX + 2.4, y: ry, w: 1.3, h: 0.46, rectRadius: 0.06,
          fill: { color: C.SECONDARY },
        }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
    }
  );

  // Slide 10: I Do #1 - estimate 4.2 x 19
  workedExSlide(pres, 2, "I Do", "About how much is 4.2 x 19?",
    [
      "We only need a close answer, so estimate.",
      "Round 4.2 down to 4.",
      "Round 19 up to 20.",
      "Multiply the friendly numbers:",
      "4 x 20 = 80.",
      "",
      "4.2 x 19 is about 80.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addEstimatePanel(slide, lg, {
        color: C.PRIMARY,
        round1: "4.2 -> 4",
        round2: "19 -> 20",
        friendly: "4 x 20 = 80",
        result: "about 80",
        note: "One rounded down, one up - very close.",
      });
    }
  );

  // Slide 11: I Do #2 - estimate 3.8 x 312
  workedExSlide(pres, 2, "I Do", "About how much is 3.8 x 312?",
    [
      "A three-digit number - round it to a",
      "friendly multiple of 100.",
      "Round 3.8 up to 4.",
      "Round 312 down to 300.",
      "4 x 300 = 1200.",
      "",
      "3.8 x 312 is about 1200.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addEstimatePanel(slide, lg, {
        color: C.PRIMARY,
        round1: "3.8 -> 4",
        round2: "312 -> 300",
        friendly: "4 x 300 = 1200",
        result: "about 1200",
        note: "Use the x 100 move for the friendly product.",
      });
    }
  );

  // Slides 12-13: CFU + reveal - estimate 5.1 x 39
  withReveal(
    () => cfuSlide(pres, "CFU", "Estimate 5.1 x 39",
      { technique: "Show Me Boards",
        question: "On your whiteboard: estimate 5.1 x 39.\n\nShow me your two rounded numbers and your estimate." },
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "5.1 -> 5,  39 -> 40,  5 x 40 = 200   so 5.1 x 39 is about 200", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15.5, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do #1 + reveal - estimate 6.1 x 48
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Estimate 6.1 x 48 together",
      [
        "With your partner.",
        "",
        "1.  Round 6.1 to a friendly number.",
        "2.  Round 48 to a friendly number.",
        "3.  Multiply and say 'about'.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addEstimateBlank(slide, lg, "6.1 x 48", "Round each to a friendly number.");
      }),
    (slide) => {
      addTextOnShape(slide, "6.1 -> 6,  48 -> 50,  6 x 50 = 300   so 6.1 x 48 is about 300", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15.5, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - money context 5.80 x 195
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "About how much for 195 tickets?",
      [
        "A ticket costs $5.80.",
        "About how much for 195 tickets?",
        "",
        "'About' means an estimate is fine.",
        "Round each number, then multiply.",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        addEstimateBlank(slide, lg, "$5.80 x 195", "Round to a friendly multiple of 100.");
      }),
    (slide) => {
      addTextOnShape(slide, "$5.80 -> $6,  195 -> 200,  6 x 200 = 1200   so about $1200", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15.5, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO2_A);
    }
  );

  // Slide 18: You Do
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: practice sheet", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.25, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "estimate the calculations.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "estimate the worded situations.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "decide if each estimate is over or under.", options: { fontSize: 17, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.16, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.4;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("Remember", {
      x: 0.7, y: panelY + 0.13, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });
    addTextOnShape(s, "Round BOTH numbers to friendly numbers, then multiply. Say 'about'.", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    const ex = [["4.2 x 19", "4 x 20 = 80"], ["3.8 x 312", "4 x 300 = 1200"]];
    const bx0 = 1.7;
    const bw = 2.9;
    ex.forEach((r, i) => {
      addTextOnShape(s, r[0] + "  ~  " + r[1], {
        x: bx0 + i * (bw + 0.3), y: panelY + 1.20, w: bw, h: 0.55, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.SECONDARY, width: 1.4 },
      }, { fontSize: 13.5, fontFace: FONT_H, color: C.SECONDARY, bold: true });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "Estimate 7.9 x 41 by rounding to friendly numbers.",
      "Show your two rounded numbers and your estimate.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: when is an estimate good enough instead of an exact answer?",
      scItems: [
        "I can round a decimal and a whole number to friendly numbers.",
        "I can estimate a decimal times a 2 or 3 digit number using a multiple of 10 or 100.",
        "I can explain whether my estimate is a bit over or a bit under, and why.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecPow_Lesson4_Estimation_Strategies.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Estimate decimal multiplications by rounding to friendly numbers.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "To estimate, round BOTH numbers to friendly numbers (a whole number, and a multiple of 10 or 100), then multiply. Use 'about'. Rounding boxes: round each number first, then multiply the friendly numbers.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Estimate 6.1 x 48: round 6.1 to 6 and 48 to 50. 6 x 50 = 300. So 6.1 x 48 is about 300. Both rounded up, so the exact answer is a little less.",
      y);

    y = addSectionHeading(doc, "Section 1 - Estimate (round both, then multiply)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  4.2 x 19   round to ____ x ____ = ____   (about ____)", y);
    y = addWriteLine(doc, "b)  5.1 x 39   round to ____ x ____ = ____   (about ____)", y);
    y = addWriteLine(doc, "c)  7.8 x 22   round to ____ x ____ = ____   (about ____)", y);
    y = addWriteLine(doc, "d)  3.9 x 288  round to ____ x ____ = ____   (about ____)", y);

    y = addSectionHeading(doc, "Section 2 - Worded situations", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  A drink costs $4.20. About how much for 31 drinks? __________", y);
    y = addWriteLine(doc, "b)  A book weighs 1.9 kg. About how much do 195 books weigh? __________", y);

    y = addSectionHeading(doc, "Section 3 - Over or under?", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "For 6.1 x 48 you estimated about 300. Is the exact answer a bit over or a bit under 300?", y);
    y = addWriteLine(doc, "Answer: __________   Why? _________________________________________", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Estimate a shopping total: 3 shirts at $19.90, 2 caps at $12.10. About how much altogether? __________", y);
    y = addWriteLine(doc, "Is your estimate a bit over or under the real total? Explain:", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Estimating Decimal Multiplications | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Lesson 4 practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Estimates", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  4 x 20 = 80      b)  5 x 40 = 200      c)  8 x 20 = 160      d)  4 x 300 = 1200", y);
    y = addBodyText(doc, "Exact values (for comparison): a) 79.8  b) 198.9  c) 171.6  d) 1123.2.", y, { fontSize: 10, color: C.MUTED });

    y = addSectionHeading(doc, "Section 2 - Worded situations", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  about $120 (4 x 30). Exact $130.20.      b)  about 400 kg (2 x 200). Exact 370.5 kg.", y);

    y = addSectionHeading(doc, "Section 3 - Over or under", y, { color: C.PRIMARY });
    y = addBodyText(doc, "A bit UNDER 300. Both 6.1 and 48 were rounded up to get 6 and 50, so the friendly product is a little larger than the exact answer (292.8).", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "Estimate: 3 x $20 = $60, plus 2 x $12 = $24, about $84. Exact is $83.90, so the estimate is a touch over. Reasonable rounding either way is fine.", y);

    y = addTipBox(doc,
      "Watch for: students who round only one number; students who round a 3-digit number to the nearest ten instead of hundred; students who write 'equals' instead of 'about'.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Year 8 Extension
  await (async () => {
    const doc = createPdf({ title: EXT_RES.name });
    let y = addPdfHeader(doc, EXT_RES.name, {
      subtitle: "Refining estimates, over and under, and percentage error.",
      color: C.SECONDARY,
      lessonInfo: `Lesson ${SESSION} | Year 8 challenge | extends Year 6 VC2M6N06`,
    });
    y = addTipBox(doc,
      "An estimate is a tool for judgement. Rounding one number up and one down keeps an estimate close. Rounding both the same way makes the estimate clearly over or under. You can also measure how far off you are.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Estimate 5.8 x 31 as 6 x 30 = 180. Exact value 179.8. Difference 0.2. Very close because the roundings nearly cancel.",
      y);

    y = addSectionHeading(doc, "Section 1 - Estimate, then check", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "For each: write your estimate, the exact product, and the difference.", y);
    y = addWriteLine(doc, "a)  5.8 x 31   estimate ____   exact ____   difference ____", y);
    y = addWriteLine(doc, "b)  6.1 x 290  estimate ____   exact ____   difference ____", y);

    y = addSectionHeading(doc, "Section 2 - Over or under, and why", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  3.9 x 51 estimated as 4 x 50 = 200. Over or under? Explain.", y);
    y = addWriteLine(doc, "b)  7.2 x 18 estimated as 7 x 20 = 140. Over or under? Explain.", y);

    y = addSectionHeading(doc, "Section 3 - Reasoning in context", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "A caterer needs about how much for 285 meals at $12.40 each? Estimate with friendly", y);
    y = addWriteLine(doc, "numbers. Is a budget of $4000 enough? Explain your reasoning.", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addTipBox(doc,
      "Teacher answers: S1  a) est 180, exact 179.8, diff 0.2.  b) est 1800 (6 x 300), exact 1769, diff 31.   S2  a) OVER: both 3.9 and 51 round to give 200, exact is 198.9.  b) OVER: 7.2 rounds down but 18 rounds up to 20, exact 129.6, estimate 140 is over.   S3  About 12 x 300 = 3600 (or 12.4 x 285 ~ 3534). Yes, $4000 is enough, with a few hundred to spare.",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, `Lesson ${SESSION} | Year 8 Extension | Estimation`);
    await writePdf(doc, path.join(OUT_DIR, EXT_RES.fileName));
    console.log("PDF written: " + EXT_RES.fileName);
  })();

  console.log("Lesson 4 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
