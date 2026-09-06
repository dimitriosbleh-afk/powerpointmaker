"use strict";

// Decimal Operations & Place Value Mastery (Year 6 Numeracy) - Lesson 3 of 4.
// Multiply and divide decimals by MULTIPLES of powers of 10 (x20, x300, /40).
// Strategy: split the multiple into a known fact and a power of 10, do the
// known fact first, then move the digits. VC2M6N06.
// Daily Review: Practical Problem Solving and Financial Reasoning (prior).
// Fluency: lattice multiplication.
// Unit variant fixed (variant 2) for cohesion.
// Catch-up: launch re-activates known facts and the times-10 move; the
// worksheet has an enabling start. No session assumes the one before it.

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
  fluencySlide, dailyReviewSlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 3;
const TOTAL = 4;
const UNIT_TITLE = "Decimal Operations and Place Value";
const FOOTER = `Decimal Operations | Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/DecPow_Lesson3_Multiples_Of_Powers_Of_10";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Multiply and divide decimals by multiples of 10, 100 and 1000.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Lesson 3 practice sheet.");
const EXT_RES = makeSessionResource(SESSION,
  "Year 8 Extension",
  "Splitting bigger multiples and working with index form.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXT_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back to our decimal operations unit.
- So far we have multiplied and divided decimals by 10, 100 and 1000.
- Today we use those moves to handle bigger jobs, like times 40 or divide by 20, by splitting them into two easy steps.

DO:
- Have whiteboards and markers ready.
- Settle the room and focus attention before you click past the title.

TEACHER NOTES:
Lesson 3 of 4. It builds on the times and divide moves, but the launch rebuilds them, so a student who missed earlier lessons can still access today.

WATCH FOR:
- Students who already see 40 as 4 times 10 - strong start.
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
A student who missed earlier sessions can still do today. The launch rebuilds the times-10 move and known facts, and Section 1 of the worksheet is started for them. The two-step split only needs one known fact and one place value move, both of which we model today.

[General: Resources | Element: Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. More money and problem solving from earlier work.
- Read each one carefully and work it on your whiteboard.
- Show your thinking, not just the answer.

DO:
- Display the three prompts.
- Give about 2 minutes.
- Walk and listen for how students reason about money and fractions of an amount.

TEACHER NOTES:
Daily Review is prior learning - practical problem solving and financial reasoning - not today's new content.

WATCH FOR:
- Students who simplify the fraction to a quarter - secure.
- Students who divide the total by the number of tickets cleanly - secure.

[Stage 1: Daily Review | Element: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- Saving 6 dollars off 24 dollars is 6 out of 24, which simplifies to one quarter.
- Three tickets at 37 dollars 50 total is 12 dollars 50 each.
- Nine dollars an hour for six hours is 54 dollars.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The fraction question is the key one - 6 out of 24 simplifies to one quarter. Note students who leave it as six twenty-fourths.

WATCH FOR:
- Students who self-correct quickly - secure.

[Stage 1: Daily Review Answers | Element: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Today we use the lattice method for multiplication.
- Draw your lattice grid, multiply each pair of digits, then add along the diagonals.
- Carry into the next diagonal where you need to.

DO:
- Display the three prompts.
- Give about 2 to 3 minutes.
- Scan for neat lattice grids and correct diagonal sums.

TEACHER NOTES:
We switch from vertical multiplication to lattice this lesson and next. Lattice keeps place value tidy along the diagonals, which is the same place value thinking we use with decimals. Have grid templates ready if students need them.

WATCH FOR:
- Students who set out a clean lattice - secure.
- Students who add diagonals without carrying - prompt them at the ten boundary.

[Stage 1: Fluency | Element: Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 27 times 34 is 918.
- 46 times 23 is 1058.
- 58 times 16 is 928.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
The lattice gives the same product as the vertical method. If a total is off by a multiple of ten, look for a missed diagonal carry.

WATCH FOR:
- Students who self-correct - secure.
- Students whose diagonals do not line up - small group focus.

[Stage 1: Fluency Answers | Element: Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Some of you may remember that 6 times 4 is 24.
- So what is 6 times 40? It is 6 times 4, which is 24, and then times 10, which makes 240.
- A multiple like 40 is just a known fact times a power of 10.
- Today we use that trick with decimals.

DO:
- Point to the known fact, then the times 10 step.
- Have students chorus 'known fact, then times ten'.
- Bridge: 'the same split works when the number is a decimal'.

TEACHER NOTES:
This launch re-activates known facts and the times-10 move from Lesson 1. It is the catch-up bridge - a returning student only needs one times table fact and one place value move.

WATCH FOR:
- Students who split 40 into 4 times 10 - strong prior knowledge.
- Students unsure - reassure, we build it together this lesson.

[Stage: Launch | Element: Knowledge and Memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to multiply and divide decimals by multiples of 10, 100 and 1000.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.
- Keep the place value chart on display for reference.

TEACHER NOTES:
The first criterion is reachable for everyone - split a multiple like 40 into 4 times 10. The second is the core target the exit ticket checks. The third stretches to explaining why the split works.

WATCH FOR:
- Students who can repeat the language - tracking.

[General: LI/SC | Element: Planning]`;

const NOTES_VOCAB = `SAY:
- Three words for today.
- A multiple of a power of 10 is a number like 20, 40 or 300 - a single digit times 10, 100 or 1000.
- To split means to break the multiple into a known fact and a power of 10.
- A known fact is a times table answer you already know.

DO:
- Point to each word as you say it.
- Have students split 60 into 6 times 10 out loud.

TEACHER NOTES:
Vocabulary comes after the learning intention. The split is the heart of the lesson: known fact first, then move the digits.

WATCH FOR:
- Students who can split a multiple - secure.
- Students who try to do it all in one step - steer them to two steps.

[General: Key Vocabulary | Element: Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us look at this one together. Watch how I work out 3.2 times 40.
- 40 is the same as 4 times 10, so I split the job into two steps.
- Step one, the known fact: 3.2 times 4 is 12.8.
- Step two, the power of 10: 12.8 times 10 moves every digit one place left, giving 128.
- So 3.2 times 40 is 128. Known fact first, then the times 10 move.

DO:
- Point to the split, then to each step in turn.
- Name the place value move in step two, do not say 'add a zero'.
- Say the two steps once more together.

TEACHER NOTES:
This is the core strategy. Keep the two steps separate and visible. Step two is exactly the move from Lesson 1.

MISCONCEPTIONS:
- Misconception: students do only the known fact and stop at 12.8.
  Why: they forget the times 10 step.
  Impact: their answer is ten times too small.
  Quick correction: ask 'have you done both steps - the fact and the power of 10?'

WATCH FOR:
- Students who complete both steps - secure.
- Students who stop after the known fact - prompt for step two.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now a division. Watch how I work out 8.4 divided by 20.
- 20 is 2 times 10, so again I split into two steps.
- Step one, the known fact: 8.4 divided by 2 is 4.2.
- Step two, the power of 10: 4.2 divided by 10 moves every digit one place right, giving 0.42.
- So 8.4 divided by 20 is 0.42. Same idea: known fact first, then the place value move.

DO:
- Point to the split, then to each step.
- Name the place value move in step two.
- Connect back: 'this is the divide move from last lesson'.

TEACHER NOTES:
Division by a multiple uses the same split. Order does not change the answer, but doing the known fact first keeps the numbers friendly.

MISCONCEPTIONS:
- Misconception: students divide by 2 and forget to divide by 10.
  Why: they stop after the known fact.
  Impact: their answer is ten times too big.
  Quick correction: ask 'have you done both steps?'

WATCH FOR:
- Students who complete both steps - secure.
- Students who stop after dividing by 2 - prompt for step two.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards.
- Work out 0.5 times 60.
- Show me your two steps, not just the answer.

DO:
- Display the prompt.
- Give 60 seconds.
- Walk and scan for both steps.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me on three, two, one, show.
- Scan for: 0.5 times 6 is 3, then 3 times 10 is 30. Answer 30.
PROCEED: If about 80 percent are correct, click to reveal and move to We Do.
PIVOT: Most likely misconception - students write 3 because they forgot the times 10 step.
- Reteach: 60 is 6 times 10. Do the known fact, 0.5 times 6 is 3, then the times 10 move, 3 times 10 is 30.
- Re-check: what are the two steps, and what is the final answer?

TEACHER NOTES:
The trap is doing only the known fact. A student who writes 3 has skipped the power of 10. Make them show both steps.

WATCH FOR:
- Students who show both steps and write 30 - secure.
- Students who write 3 - they skipped step two, reteach the split.

[Stage 2: CFU | Element: Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Work out 2.4 times 30.
- Split the 30 into 3 times 10.
- Whisper to your partner: what is step one, and what is step two?

DO:
- Display 2.4 times 30 above the two-step frame.
- Give 75 seconds.
- Listen for '2.4 times 3 is 7.2, then times 10 is 72'.

TEACHER NOTES:
Same strategy as the I Do with new numbers. Keep both steps visible.

WATCH FOR:
- Pairs who complete both steps - secure.
- Pairs who stop at 7.2 - prompt for the times 10 move.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- Step one: 2.4 times 3 is 7.2.
- Step two: 7.2 times 10 is 72.
- So 2.4 times 30 is 72.

DO:
- Click to reveal.
- Run the two steps once more together.

TEACHER NOTES:
Reveal restates the two steps. The times 10 move is the place value move from Lesson 1.

WATCH FOR:
- Students who self-correct - secure.
- Students who stop at 7.2 - remind them of step two.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together, a division.
- Work out 9.6 divided by 40.
- Split the 40 into 4 times 10. What is step one, and what is step two?

DO:
- Display 9.6 divided by 40 above the two-step frame.
- Give 90 seconds.
- Watch for students who divide by 4 but forget to divide by 10.

TEACHER NOTES:
Division by a multiple. Step one is the known fact 9.6 divided by 4 is 2.4. Step two is 2.4 divided by 10 is 0.24.

WATCH FOR:
- Students who write 0.24 - secure.
- Students who write 2.4 - they forgot step two, reteach the split.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- Step one: 9.6 divided by 4 is 2.4.
- Step two: 2.4 divided by 10 is 0.24.
- So 9.6 divided by 40 is 0.24.

DO:
- Click to reveal.
- Point to the divide-by-10 move in step two.

TEACHER NOTES:
If many stopped at 2.4, do one more divide-by-a-multiple example before releasing to the You Do.

WATCH FOR:
- Students who complete both steps - ready for independent work.
- Students who stop at 2.4 - enabling group for You Do.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- Section 1 is multiply by a multiple. Section 2 is divide by a multiple.
- Section 3 asks you to show the two steps and explain why the split works. If you finish, try the extension box.

DO:
- Distribute the practice sheet.
- Circulate and check both steps are shown.
- Cold call one or two students to explain a split.

TEACHER NOTES:
Different numbers from the We Do, same strategy: split the multiple, known fact first, then move the digits.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Section 1 question a is fully worked. Copy the two-step frame for the next ones, splitting the multiple first.
- Extra Notes: Sit with these students and split the first multiple together. This is also the rebuild for any returning student.
EXTENDING PROMPT:
- Task: The extension box asks students to reach a target answer in two different ways and to split a larger multiple like 400.
- Extra Notes: Students who are ready can move on to the Year 8 extension sheet.

WATCH FOR:
- Students who show both steps clearly - secure.
- Students who do one step only - prompt for the missing step.

[Stage 4: You Do | Element: Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard, on your own.
- Work out 4.5 times 20.
- Show both steps: the known fact, then the power of 10.

DO:
- Display the prompt.
- Give 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core success criterion - multiply a decimal by a multiple of a power of 10 in two steps. Look for 4.5 times 2 is 9, then 9 times 10 is 90.

WATCH FOR:
- Students who show both steps and write 90 - secure.
- Students who write 9 - they skipped the power of 10, revisit at the start of Lesson 4.

[Stage 5: Exit Ticket | Element: Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: how do you split a multiple like 300 into two easy steps?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is that a multiple of a power of 10 splits into a known fact and a power of 10. Students who can split and do both steps are ready to estimate next lesson.

WATCH FOR:
- Strong thumbs up across all three - move at pace next lesson.
- Sideways or down on the core criterion - small group revision at the start of Lesson 4.

[General: Closing | Element: Retention and Recall]`;

// --- Helpers -----------------------------------------------------------------

// Two-step "split" panel used on the I Do and We Do slides.
function addTwoStepPanel(slide, lg, cfg) {
  const color = cfg.color || C.PRIMARY;
  const cardH = cfg.cardH || 3.35;
  const top = lg.panelTopPadded;
  const cx = lg.rightX + 0.25;
  const cw = lg.rightW - 0.50;
  addCard(slide, lg.rightX, top, lg.rightW, cardH, { strip: color });

  slide.addText(cfg.splitText, {
    x: lg.rightX + 0.15, y: top + 0.10, w: lg.rightW - 0.30, h: 0.32,
    fontSize: 15, fontFace: FONT_H, color: color, bold: true,
    align: "center", margin: 0,
  });

  addTextOnShape(slide, cfg.step1, {
    x: cx, y: top + 0.52, w: cw, h: 0.60, rectRadius: 0.08,
    fill: { color: C.PRIMARY },
  }, { fontSize: 14.5, fontFace: FONT_H, color: C.WHITE, bold: true });

  slide.addText(cfg.bridge || "then move the digits", {
    x: cx, y: top + 1.16, w: cw, h: 0.28,
    fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
    align: "center", margin: 0,
  });

  addTextOnShape(slide, cfg.step2, {
    x: cx, y: top + 1.46, w: cw, h: 0.60, rectRadius: 0.08,
    fill: { color: C.SECONDARY },
  }, { fontSize: 14.5, fontFace: FONT_H, color: C.WHITE, bold: true });

  addTextOnShape(slide, cfg.result, {
    x: cx, y: top + 2.22, w: cw, h: 0.62, rectRadius: 0.08,
    fill: { color: C.SUCCESS },
  }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });

  if (cfg.note) {
    slide.addText(cfg.note, {
      x: lg.rightX + 0.15, y: top + 2.90, w: lg.rightW - 0.30, h: 0.32,
      fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
      align: "center", margin: 0,
    });
  }
}

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 3: Multiplying and dividing by multiples of 10, 100 and 1000",
    `Year 6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - financial reasoning
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Money problems",
      [
        "A $24 shirt is reduced by $6. What fraction do you save?",
        "Three tickets cost $37.50 in total. How much is one ticket?",
        "You earn $9 an hour for 6 hours. How much do you earn?",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "1/4 (six 24ths)          $12.50 each          $54", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - lattice multiplication
  withReveal(
    () => fluencySlide(pres, "Fluency: Lattice multiplication",
      ["27 x 34", "46 x 23", "58 x 16"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "918        1058        928", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - from whole numbers (catch-up bridge)
  contentSlide(pres, "Launch", C.ACCENT, "Bigger jobs, two easy steps",
    [
      "Some of you may remember: 6 x 4 = 24.",
      "So 6 x 40 = 240.",
      "",
      "That is 6 x 4, then x 10.",
      "Today: the same split, with decimals.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.ACCENT });
      slide.addText("Split the multiple", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });
      const chips = [["6 x 4 = 24", "known fact"], ["x 10 -> 240", "power of 10"]];
      const cy0 = lg.panelTopPadded + 0.58;
      chips.forEach((r, i) => {
        const cy = cy0 + i * 0.64;
        addTextOnShape(slide, r[0], {
          x: lg.rightX + 0.25, y: cy, w: 2.35, h: 0.50, rectRadius: 0.07,
          fill: { color: i === 0 ? C.PRIMARY : C.SECONDARY },
        }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(r[1], {
          x: lg.rightX + 2.7, y: cy, w: lg.rightW - 2.9, h: 0.50,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", valign: "middle", margin: 0,
        });
      });
      slide.addText("40 = 4 x 10. Known fact, then power of 10.", {
        x: lg.rightX + 0.15, y: cy0 + 1.34, w: lg.rightW - 0.30, h: 0.34,
        fontSize: 12.5, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to multiply and divide decimals by multiples of 10, 100 and 1000.",
    [
      "I can split a multiple like 40 into a known fact and a power of 10.",
      "I can multiply or divide a decimal by a multiple of a power of 10 using two steps.",
      "I can explain why splitting into a known fact and a power of 10 works.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Words for today",
    [
      "Multiple of a power of 10 = 20, 40, 300 ...",
      "Split = break it into a known fact x a power of 10",
      "Known fact = a times table you already know",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.SECONDARY });
      slide.addText("Split a multiple", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      const rows = [["40", "4 x 10"], ["300", "3 x 100"], ["2000", "2 x 1000"]];
      const ry0 = lg.panelTopPadded + 0.55;
      rows.forEach((r, i) => {
        const ry = ry0 + i * 0.56;
        addTextOnShape(slide, r[0], {
          x: lg.rightX + 0.35, y: ry, w: 1.5, h: 0.46, rectRadius: 0.06,
          fill: { color: C.PRIMARY },
        }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText("= " + r[1], {
          x: lg.rightX + 2.05, y: ry, w: lg.rightW - 2.3, h: 0.46,
          fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      });
    }
  );

  // Slide 10: I Do #1 - 3.2 x 40 (two-step)
  workedExSlide(pres, 2, "I Do", "Multiply 3.2 x 40",
    [
      "40 is the same as 4 x 10.",
      "So split the job into two steps.",
      "",
      "Step 1 (known fact): 3.2 x 4 = 12.8",
      "Step 2 (power of 10): 12.8 x 10 = 128",
      "",
      "3.2 x 40 = 128",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addTwoStepPanel(slide, lg, {
        color: C.PRIMARY,
        splitText: "Split: 40 = 4 x 10",
        step1: "Step 1   x 4:   3.2 x 4 = 12.8",
        step2: "Step 2   x 10:   12.8 x 10 = 128",
        result: "3.2 x 40 = 128",
        note: "Known fact first, then the power of 10.",
      });
    }
  );

  // Slide 11: I Do #2 - 8.4 / 20 (two-step divide)
  workedExSlide(pres, 2, "I Do", "Divide 8.4 ÷ 20",
    [
      "20 is the same as 2 x 10.",
      "So split the job into two steps.",
      "",
      "Step 1 (known fact): 8.4 ÷ 2 = 4.2",
      "Step 2 (power of 10): 4.2 ÷ 10 = 0.42",
      "",
      "8.4 ÷ 20 = 0.42",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addTwoStepPanel(slide, lg, {
        color: C.PRIMARY,
        splitText: "Split: 20 = 2 x 10",
        step1: "Step 1   ÷ 2:   8.4 ÷ 2 = 4.2",
        step2: "Step 2   ÷ 10:   4.2 ÷ 10 = 0.42",
        result: "8.4 ÷ 20 = 0.42",
        note: "Same idea: known fact, then the place value move.",
      });
    }
  );

  // Slides 12-13: CFU + reveal - 0.5 x 60
  withReveal(
    () => cfuSlide(pres, "CFU", "Work out 0.5 x 60",
      { technique: "Show Me Boards",
        question: "On your whiteboard: 0.5 x 60 = ?\n\nShow me BOTH steps, not just the answer." },
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "0.5 x 6 = 3, then 3 x 10 = 30   so   0.5 x 60 = 30", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do #1 + reveal - 2.4 x 30
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Build 2.4 x 30 together",
      [
        "With your partner.",
        "",
        "1.  Split 30 into 3 x 10.",
        "2.  Step 1: the known fact.",
        "3.  Step 2: the x 10 move.",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.8, { strip: C.SECONDARY });
        slide.addText("2.4 x 30", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.42,
          fontSize: 28, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        slide.addText("30 = 3 x 10", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.60, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
        addTextOnShape(slide, "Step 1   x 3:   ________", {
          x: lg.rightX + 0.25, y: lg.panelTopPadded + 1.02, w: lg.rightW - 0.50, h: 0.56, rectRadius: 0.08,
          fill: { color: C.PRIMARY },
        }, { fontSize: 14.5, fontFace: FONT_H, color: C.WHITE, bold: true });
        addTextOnShape(slide, "Step 2   x 10:   ________", {
          x: lg.rightX + 0.25, y: lg.panelTopPadded + 1.66, w: lg.rightW - 0.50, h: 0.56, rectRadius: 0.08,
          fill: { color: C.PRIMARY },
        }, { fontSize: 14.5, fontFace: FONT_H, color: C.WHITE, bold: true });
      }),
    (slide) => {
      addTextOnShape(slide, "Step 1: 2.4 x 3 = 7.2.  Step 2: 7.2 x 10 = 72.  So 2.4 x 30 = 72", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15.5, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - 9.6 / 40
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "A division: 9.6 ÷ 40",
      [
        "With your partner.",
        "",
        "1.  Split 40 into 4 x 10.",
        "2.  Step 1: 9.6 ÷ 4.",
        "3.  Step 2: the ÷ 10 move.",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.8, { strip: C.SECONDARY });
        slide.addText("9.6 ÷ 40", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.42,
          fontSize: 28, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        slide.addText("40 = 4 x 10", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.60, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "center", margin: 0,
        });
        addTextOnShape(slide, "Step 1   ÷ 4:   ________", {
          x: lg.rightX + 0.25, y: lg.panelTopPadded + 1.02, w: lg.rightW - 0.50, h: 0.56, rectRadius: 0.08,
          fill: { color: C.PRIMARY },
        }, { fontSize: 14.5, fontFace: FONT_H, color: C.WHITE, bold: true });
        addTextOnShape(slide, "Step 2   ÷ 10:   ________", {
          x: lg.rightX + 0.25, y: lg.panelTopPadded + 1.66, w: lg.rightW - 0.50, h: 0.56, rectRadius: 0.08,
          fill: { color: C.PRIMARY },
        }, { fontSize: 14.5, fontFace: FONT_H, color: C.WHITE, bold: true });
      }),
    (slide) => {
      addTextOnShape(slide, "Step 1: 9.6 ÷ 4 = 2.4.  Step 2: 2.4 ÷ 10 = 0.24.  So 9.6 ÷ 40 = 0.24", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
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
      { text: "multiply by a multiple.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "divide by a multiple.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "show both steps and explain the split.", options: { fontSize: 17, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.16, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.4;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("Remember the two steps", {
      x: 0.7, y: panelY + 0.13, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });
    addTextOnShape(s, "Step 1: known fact.    Step 2: move the digits (the power of 10).", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    const splits = [["x 40 = x 4 then x 10"], ["x 300 = x 3 then x 100"], ["÷ 20 = ÷ 2 then ÷ 10"]];
    const bx0 = 1.15;
    const bw = 2.45;
    splits.forEach((r, i) => {
      addTextOnShape(s, r[0], {
        x: bx0 + i * (bw + 0.15), y: panelY + 1.20, w: bw, h: 0.55, rectRadius: 0.08,
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
      "Work out 4.5 x 20.",
      "Show both steps: the known fact, then the power of 10.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: how do you split a multiple like 300 into two easy steps?",
      scItems: [
        "I can split a multiple like 40 into a known fact and a power of 10.",
        "I can multiply or divide a decimal by a multiple of a power of 10 using two steps.",
        "I can explain why splitting into a known fact and a power of 10 works.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "DecPow_Lesson3_Multiples_Of_Powers_Of_10.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Multiply and divide decimals by multiples of 10, 100 and 1000.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "To multiply or divide by a multiple of a power of 10, split it into a known fact and a power of 10. Do the known fact first, then move the digits. Example: x40 = x4 then x10; div40 = div4 then div10.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "4.2 x 30: split 30 into 3 x 10. Step 1, 4.2 x 3 = 12.6. Step 2, 12.6 x 10 = 126. So 4.2 x 30 = 126.",
      y);

    y = addSectionHeading(doc, "Section 1 - Multiply by a multiple (first one done)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  3.2 x 40 = 3.2 x 4 x 10 = 12.8 x 10 = 128", y);
    y = addWriteLine(doc, "b)  2.4 x 30 = _____ x 10 = _______", y);
    y = addWriteLine(doc, "c)  0.6 x 50 = _______", y);
    y = addWriteLine(doc, "d)  1.5 x 200 = _______", y);

    y = addSectionHeading(doc, "Section 2 - Divide by a multiple", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  8.4 ÷ 20 = _____ ÷ 10 = _______", y);
    y = addWriteLine(doc, "b)  9.6 ÷ 40 = _______", y);
    y = addWriteLine(doc, "c)  360 ÷ 300 = _______", y);
    y = addWriteLine(doc, "d)  7.2 ÷ 60 = _______", y);

    y = addSectionHeading(doc, "Section 3 - Show your two steps", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "4.5 x 20 = _______.  Step 1: __________   Step 2: __________", y);
    y = addWriteLine(doc, "Explain why splitting into a known fact and a power of 10 works:", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addBodyText(doc, "Split a bigger multiple: 2.5 x 400 = __________.  (Hint: 400 = 4 x 100.)", y);
    y = addWriteLine(doc, "Find TWO different decimal x multiple calculations that both give 240:", y);
    y = addWriteLine(doc, "__________________________   and   __________________________", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Multiples of Powers of 10 | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Lesson 3 practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Multiply by a multiple", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  128        b)  7.2 x 10 = 72        c)  3 x 10 = 30        d)  3 x 100 = 300", y);

    y = addSectionHeading(doc, "Section 2 - Divide by a multiple", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  4.2 ÷ 10 = 0.42        b)  2.4 ÷ 10 = 0.24        c)  120 ÷ 100 = 1.2        d)  1.2 ÷ 10 = 0.12", y);

    y = addSectionHeading(doc, "Section 3 - Two steps", y, { color: C.PRIMARY });
    y = addBodyText(doc, "4.5 x 20 = 90. Step 1: 4.5 x 2 = 9. Step 2: 9 x 10 = 90. It works because 20 = 2 x 10, so multiplying by 2 then by 10 is the same as multiplying by 20.", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "2.5 x 400 = 1000 (2.5 x 4 = 10, then x 100). Many pairs give 240, for example 0.6 x 400, 1.2 x 200, 2.4 x 100, 4 x 60, 8 x 30.", y);

    y = addTipBox(doc,
      "Watch for: students who do only the known fact and stop (answer ten times out); students who do only the power of 10; students who split the multiple incorrectly (40 is 4 x 10, not 4 x 100).",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Year 8 Extension
  await (async () => {
    const doc = createPdf({ title: EXT_RES.name });
    let y = addPdfHeader(doc, EXT_RES.name, {
      subtitle: "Splitting bigger multiples and working with index form.",
      color: C.SECONDARY,
      lessonInfo: `Lesson ${SESSION} | Year 8 challenge | extends Year 6 VC2M6N06`,
    });
    y = addTipBox(doc,
      "The same split scales up. Any multiple of a power of 10 is a known fact times a power of 10: 1200 = 12 x 100, 2000 = 2 x 1000. In index form, 100 = 10 squared and 1000 = 10 cubed.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "0.45 x 1200: split 1200 into 12 x 100. Step 1, 0.45 x 12 = 5.4. Step 2, 5.4 x 100 = 540. So 0.45 x 1200 = 540.",
      y);

    y = addSectionHeading(doc, "Section 1 - Split bigger multiples", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  3.5 x 2000 = _______", y);
    y = addWriteLine(doc, "b)  0.45 x 1200 = _______", y);
    y = addWriteLine(doc, "c)  84 ÷ 400 = _______", y);

    y = addSectionHeading(doc, "Section 2 - Index form", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  6.2 x (3 x 10 squared) = _______", y);
    y = addWriteLine(doc, "b)  9 ÷ (6 x 10) = _______", y);
    y = addWriteLine(doc, "c)  Write 7.5 x 800 as a known fact times a power of 10, then evaluate.", y);
    y = addWriteLine(doc, "    _____________________________________   = _______", y);

    y = addSectionHeading(doc, "Section 3 - Reasoning", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "Explain why 0.45 x 1200 can be done as 0.45 x 12 x 100. Which property are you using?", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addTipBox(doc,
      "Teacher answers: S1  a) 7000 (3.5 x 2 = 7, x 1000)  b) 540  c) 0.21 (84 div 4 = 21, div 100).   S2  a) 1860  b) 0.15  c) 7.5 x 8 x 100 = 60 x 100 = 6000.   S3  The associative property of multiplication: 1200 = 12 x 100, so 0.45 x 1200 = 0.45 x (12 x 100) = (0.45 x 12) x 100.",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, `Lesson ${SESSION} | Year 8 Extension | Multiples of Powers of 10`);
    await writePdf(doc, path.join(OUT_DIR, EXT_RES.fileName));
    console.log("PDF written: " + EXT_RES.fileName);
  })();

  console.log("Lesson 3 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
