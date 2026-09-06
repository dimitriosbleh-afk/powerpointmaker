"use strict";

// Four Processes (Year 6 Numeracy) - Lesson 5 of 5.
// Modelling real problems: draw a model, CHOOSE the operation and justify it,
// decide whether an exact answer or an estimate is needed, and build a simple
// budget (income, costs, money left over). VC2M5N09, VC2M6N09.
// Daily Review: Decimal Operations and Place Value Mastery (prior).
// Fluency: decimal x 1-digit algorithm.
// Unit variant fixed (variant 3, Ocean Logic) across all 5 lessons for cohesion.
// Catch-up: the launch rebuilds 'choose the operation' from a $12 pizza everyone
// can model, and worksheet Section 1 is an enabling rebuild. No session assumes
// the one before; this synthesis lesson re-teaches the modelling routine.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", 3); // variant 3 (Ocean Logic), fixed for the unit
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  fluencySlide, dailyReviewSlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 5;
const TOTAL = 5;
const UNIT_TITLE = "Four Processes";
const FOOTER = `Four Processes | Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FourP_Lesson5_Modelling_And_Budgets";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Model word problems, choose the operation, and build a simple budget.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Lesson 5 practice sheet.");
const EXT_RES = makeSessionResource(SESSION,
  "Year 8 Extension",
  "Budgeting with unit rates and savings plans - for students ready for more.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXT_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Helpers -----------------------------------------------------------------
function drawBarModel(slide, x, y, w, h, parts, shaded, opts) {
  const o = opts || {};
  const fill = o.fill || C.SECONDARY;
  const empty = o.empty || C.WHITE;
  const seg = w / parts;
  for (let i = 0; i < parts; i++) {
    slide.addShape("rect", {
      x: x + i * seg, y, w: seg, h,
      fill: { color: i < shaded ? fill : empty },
      line: { color: C.PRIMARY, width: 1 },
    });
    if (o.segLabel) {
      slide.addText(String(o.segLabel), {
        x: x + i * seg, y, w: seg, h,
        fontSize: o.segFont || 12, fontFace: FONT_B,
        color: i < shaded ? C.WHITE : C.MUTED, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  }
  return { seg };
}

// Operation chooser: four symbol chips. The chosen one is solid; the others
// are ghosted outlines so the choice stands out (and contrast stays clean).
function drawOperationChooser(slide, x, y, w, chosen) {
  const ops = ["+", "-", "x", "÷"];
  const chipW = 0.72;
  const gap = (w - 4 * chipW) / 3;
  ops.forEach((op, i) => {
    const cx = x + i * (chipW + gap);
    if (op === chosen) {
      addTextOnShape(slide, op, {
        x: cx, y, w: chipW, h: 0.6, rectRadius: 0.08, fill: { color: C.ALERT },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
    } else {
      slide.addShape("roundRect", {
        x: cx, y, w: chipW, h: 0.6, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.MUTED, width: 1.5 },
      });
      slide.addText(op, {
        x: cx, y, w: chipW, h: 0.6,
        fontSize: 24, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  });
}

// Simple budget table: rows of [label, amount, colour].
function drawBudgetTable(slide, x, y, w, rows) {
  const rowH = 0.56;
  rows.forEach((r, i) => {
    const ry = y + i * (rowH + 0.14);
    slide.addText(r[0], {
      x: x, y: ry, w: w * 0.52, h: rowH,
      fontSize: 15, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      valign: "middle", margin: 0,
    });
    addTextOnShape(slide, r[1], {
      x: x + w * 0.52, y: ry, w: w * 0.48, h: rowH, rectRadius: 0.07,
      fill: { color: r[2] },
    }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
  });
}

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to the last lesson of Four Processes. We have linked the operations, used the properties, worked with decimals, and found fractions and percentages.
- Today we put it all together on real problems. The skill is choosing the right operation and deciding when we need an exact answer or just a good estimate.
- We will finish by building a simple budget, just like planning a class fundraiser.

DO:
- Have whiteboards, markers and a calculator on hand for checking.
- Settle the room and focus attention before you click past the title.

TEACHER NOTES:
Lesson 5 of 5. This is the synthesis lesson. The routine is model it, choose the operation, decide exact or estimate.

WATCH FOR:
- Students who look unsure - that is expected. Reassure them: if this feels new, that is okay, we build it together.

[General: Title | Element: Planning and Enabling Learning]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The practice sheet is for the You Do near the end. The Year 8 extension is there for anyone who is flying.

DO:
- Print one practice sheet and one answer key per student.
- Have whiteboards, markers and a calculator on hand for checking exact answers.
- Keep the Year 8 extension on hand for early finishers.

TEACHER NOTES:
One student practice sheet, an answer key, and a Year 8 extension. Most teaching happens on whiteboards with bar models and a simple budget table.

CATCH-UP NOTE:
A student who missed earlier sessions can still access today. The launch rebuilds 'choose the operation' from a simple pizza-sharing problem, and Section 1 of the worksheet rebuilds it. A returner only needs a whiteboard and one minute with you to start. No session in this unit assumes the student saw the one before.

[General: Resources | Element: Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. We are warming up our decimal operations and place value from earlier work.
- Read each one carefully and work it on your whiteboard.
- Show your thinking, not just the answer.

DO:
- Display the three prompts.
- Give about 2 minutes.
- Walk and listen for clean place value and decimal alignment.

TEACHER NOTES:
Daily Review is prior learning, the decimal place value work from earlier this term. It feeds straight into today's money calculations.

WATCH FOR:
- Students who line up the decimal points when adding - secure.
- Students who add 3.6 and 2.45 without aligning places - reteach in the reveal.

[Stage 1: Daily Review | Element: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- 4.7 times 100 is 470. Each digit moves two places left.
- 3.6 plus 2.45 is 6.05. Line up the points: 3.60 plus 2.45.
- 8.62 rounded to the nearest whole number is 9, because the tenths digit is 6.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
The addition item checks lining up place value: 3.60 plus 2.45. Add a zero to 3.6 so the columns line up.

WATCH FOR:
- Students who self-correct quickly - secure.
- Students who add 3.6 and 2.45 as 3.6 plus 2.45 misaligned - reteach lining up the points.

[Stage 1: Daily Review Answers | Element: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Multiplying a decimal by a single digit using the algorithm.
- Set each one out vertically on your whiteboard and multiply.
- Keep your columns lined up and remember where the decimal point sits.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for clean column alignment and the point kept in line.

TEACHER NOTES:
Fluency this unit is the multiplication algorithm with decimals. Today's money problems use exactly this skill.

WATCH FOR:
- Students who keep the decimal point lined up - secure.
- Students who lose the point - prompt: one decimal place in, one decimal place out.

[Stage 1: Fluency | Element: Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 6.5 times 8 is 52.0, which we write as 52.
- 4.7 times 9 is 42.3.
- 3.8 times 7 is 26.6.
- Tick and fix.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
6.5 times 8 lands on 52.0; accept 52. Keep it brisk.

WATCH FOR:
- Students who self-correct - secure.
- Students whose point drifts - small group focus.

[Stage 1: Fluency Answers | Element: Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Some of you may remember sharing food fairly. Three friends share a 12 dollar pizza equally. How much does each pay?
- I can picture a bar of 12 dollars split into 3 equal parts. Sharing equally tells me to divide: 12 divided by 3 is 4.
- The words 'shared equally' told me which operation to use. That is the real skill today.
- All week we have practised the operations. Today we decide WHICH one a problem needs, and whether we need an exact answer or just a close estimate.

DO:
- Draw the bar of 12 split into 3.
- Ask: which operation, and which words told you? Expected: divide, because 'shared equally'.
- Bridge: 'read the problem, picture it, choose the operation'.

TEACHER NOTES:
This launch starts from a simple sharing problem everyone can model, then names today's routine. It is also the catch-up bridge for any returning student.

WATCH FOR:
- Students who connect 'shared equally' with dividing - strong prior knowledge.
- Students unsure - reassure, we build it together this lesson.

[Stage: Launch | Element: Knowledge and Memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are learning to model real problems, choose the right operation, and decide whether we need an exact answer or an estimate.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.
- Hold up a bar model and a simple budget drawn on the board.

TEACHER NOTES:
The first criterion is reachable for everyone - choose an operation and say why. The second is the core target the exit ticket checks - model and solve, deciding exact or estimate. The third stretches to building a budget.

WATCH FOR:
- Students who can repeat the language - tracking.

[General: LI/SC | Element: Planning]`;

const NOTES_VOCAB = `SAY:
- Three words for today.
- To model a problem means to show it with a bar or a diagram so the maths is clear.
- An operation is one of the four processes: add, subtract, multiply or divide.
- A budget is a plan of money coming in and money going out.

DO:
- Point to each word as you say it.
- Have students say 'model it, then choose the operation' once together.

TEACHER NOTES:
Vocabulary comes after the learning intention. 'Budget' may be new for some; connect it to planning a party or a fundraiser.

WATCH FOR:
- Students who can name the four operations - secure.
- Students unsure what a budget is - give the fundraiser example.

[General: Key Vocabulary | Element: Knowledge and Memory]`;

const NOTES_IDO1 = `SAY:
- Let us look at this one together. A fun run raises 240 dollars, shared equally between 6 classes. How much does each class get?
- First I model it: a bar of 240 dollars split into 6 equal parts.
- 'Shared equally' tells me to divide. 240 divided by 6 is 40. Each class gets 40 dollars.
- Then a second step: each class adds 10 dollars of their own. 'Adds' tells me to add: 40 plus 10 is 50.
- So each class ends with 50 dollars. The words in the problem told me which operation to use each time.

DO:
- Draw the bar of 240 split into 6.
- Circle the operation as you choose it: divide, then add.
- Have students chorus 'the words tell me the operation'.

TEACHER NOTES:
This is the core move: read the structure, choose the operation, justify it from the words. Model first so the operation is obvious.

MISCONCEPTIONS:
- Misconception: students grab the first operation they think of without reading the structure.
  Why: they rush to calculate.
  Impact: they multiply or add when the situation is sharing.
  Quick correction: ask 'are we putting together, taking away, making equal groups, or sharing?'

WATCH FOR:
- Students who justify divide from 'shared equally' - secure.
- Students who guess the operation - prompt them to model it first.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now a key decision: do I need an exact answer, or is an estimate enough?
- I have 50 dollars. A shirt is 19 dollars 95 and shorts are 24 dollars 95. Two questions.
- First: do I have enough? I do not need the exact total. Round: about 20 plus about 25 is about 45, which is under 50. Yes, I have enough. An estimate answered it.
- Second: exactly how much change? Now I need to be exact. 19.95 plus 24.95 is 44.90. 50 take away 44.90 is 5 dollars 10.
- Same shopping trip, two questions: one needs an estimate, one needs an exact answer. Pick the tool the question asks for.

DO:
- Show the estimate path and the exact path side by side.
- Point to the question word: 'enough' invites an estimate, 'exactly' demands exact.
- Have students chorus 'estimate to decide, exact to pay'.

TEACHER NOTES:
The big idea is matching the calculation to the question. Estimates are fast and fine for 'is it enough'. Exact answers are needed for change and bills.

MISCONCEPTIONS:
- Misconception: students think every problem needs an exact answer.
  Why: school maths often rewards exactness.
  Impact: they waste time on exact work when an estimate would do, or miss that 'about' signals an estimate.
  Quick correction: ask 'does the question need the exact number, or just close enough to decide?'

WATCH FOR:
- Students who estimate for 'enough' and calculate for 'change' - secure.
- Students who always calculate exactly - prompt them to read what the question needs.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_CFU_Q = `SAY:
- Quick check on your whiteboards. 6 boxes each hold 24 books. How many books altogether?
- Write which operation you would use, and an estimate of the answer.

DO:
- Display the prompt.
- Give 60 seconds.
- Walk and scan the whiteboards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me on three, two, one, show.
- Scan for: multiply, with an estimate of about 6 x 25 = 150 (exact 144).
PROCEED: If about 80 percent choose multiply with a sensible estimate, click to reveal and move to We Do.
PIVOT: Most likely misconception - students add 6 and 24, or are unsure which operation.
- Reteach: equal groups - 6 groups of 24 - means multiply. Picture 6 boxes each with 24.
- Re-check: are these equal groups, and what operation makes equal groups?

TEACHER NOTES:
The trap is not reading 'equal groups'. Six boxes of 24 is a multiplication structure.

WATCH FOR:
- Students who choose multiply and estimate about 150 - secure.
- Students who add - reteach the equal-groups structure.

[Stage 2: CFU | Element: Supported Application]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me, a two-step problem. A class sells 30 cupcakes at 2 dollars each, and they spent 18 dollars on ingredients. How much profit?
- Step 1: what is the income? Choose the operation.
- Step 2: how do you turn income into profit? Choose the operation.

DO:
- Display the cupcake problem with an income bar and a cost bar.
- Give 90 seconds.
- Listen for 'multiply for income, subtract the cost'.

TEACHER NOTES:
Same model-and-choose move as the I Do, now two steps. Income is 30 times 2 = 60. Profit is 60 minus 18 = 42.

WATCH FOR:
- Pairs who multiply then subtract - secure.
- Pairs who add the 18 - prompt: is the 18 money in or money out?

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO1_A = `SAY:
- Let us check.
- Income: 30 cupcakes at 2 dollars is 30 times 2, which is 60 dollars.
- Profit: take away what they spent: 60 minus 18 is 42 dollars.
- Two steps, two operations: multiply for income, subtract for profit.

DO:
- Click to reveal.
- Point to income, then cost, then profit on the bars.

TEACHER NOTES:
Reveal restates the two-step structure. Profit means income minus costs.

WATCH FOR:
- Students who self-correct - secure.
- Students who stop at 60 - remind them profit subtracts the cost.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_WEDO2_Q = `SAY:
- One more together: a real budget. Our cake stall raised 120 dollars. We spent 30 dollars on ingredients and 15 dollars on decorations.
- How much is left for the excursion?
- Fill in the budget: income, total costs, money left over.

DO:
- Display the budget table with blanks.
- Give 90 seconds.
- Watch for students who forget to total BOTH costs first.

TEACHER NOTES:
This is the budgeting elaboration. Total the costs first: 30 plus 15 is 45. Then money left is 120 minus 45, which is 75.

WATCH FOR:
- Students who total costs then subtract - secure.
- Students who subtract only one cost - prompt: add all the costs first.

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO2_A = `SAY:
- Check yours.
- Income is 120 dollars.
- Total costs: 30 plus 15 is 45 dollars.
- Money left over: 120 minus 45 is 75 dollars.
- That 75 dollars goes to the excursion. That is a budget: money in, money out, money left.

DO:
- Click to reveal.
- Read the budget down: income, costs, left over.

TEACHER NOTES:
The big takeaway: a budget adds the costs, then subtracts from income. The leftover is what you can spend.

WATCH FOR:
- Students who build the whole budget - ready for independent work.
- Students who miss a cost - enabling group for You Do.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the practice sheet.
- Section 1 warms up by choosing the operation for each problem.
- Section 2 decides exact or estimate, then solves.
- Section 3 builds a fundraiser budget. If you finish, try the extension box.

DO:
- Distribute the practice sheet.
- Circulate and listen for 'the words tell me the operation' and 'do I need exact or about'.
- Cold call one or two students to justify an operation choice.

TEACHER NOTES:
Different numbers from the We Do, same routine: model it, choose the operation, decide exact or estimate, and total a budget.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Section 1 has the bar drawn and the operation words underlined. Match the words to the operation.
- Extra Notes: Sit with these students and read the key words together. This is also the rebuild for any returning student.
EXTENDING PROMPT:
- Task: The extension box asks students to design their own class fundraiser budget with an income, at least two costs, and the money left over.
- Extra Notes: Students who are ready can move on to the Year 8 extension sheet on budgeting with unit rates and savings plans.

WATCH FOR:
- Students who justify each operation - secure.
- Students who guess operations - prompt them to model the problem first.

[Stage 4: You Do | Element: Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task of the unit, on your whiteboard, on your own.
- A cafe orders 6 trays of muffins. Each tray costs 7 dollars 50.
- Estimate the total, then work out the exact cost. Which one would the cafe need to pay the bill?

DO:
- Display the prompt.
- Give 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core success criterion - model and solve, choosing exact or estimate. Look for an estimate of 6 times 8 = 48, an exact answer of 6 times 7.50 = 45, and 'exact' to pay the bill.

WATCH FOR:
- Students who estimate, calculate exactly, and choose exact to pay - secure.
- Students who cannot decide exact or estimate - revisit at the start of the next unit.

[Stage 5: Exit Ticket | Element: Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again, and at the whole unit.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: what helps you decide which operation a problem needs?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is read the problem, model it, choose the operation, and decide exact or estimate. This wraps up the Four Processes unit.

WATCH FOR:
- Strong thumbs up across all three - the unit has landed.
- Sideways or down - note who to revisit in the next unit's daily reviews.

[General: Closing | Element: Retention and Recall]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Lesson 5: Modelling real problems and budgets",
    `Year 6 Numeracy | Lesson ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - decimal operations & place value
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Decimal operations & place value",
      [
        "4.7 x 100 = ?",
        "What is 3.6 + 2.45?",
        "Round 8.62 to the nearest whole number.",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "470          6.05          9", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - decimal x 1-digit
  withReveal(
    () => fluencySlide(pres, "Fluency: Decimal x single digit",
      ["6.5 x 8", "4.7 x 9", "3.8 x 7"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "52.0        42.3        26.6", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - choose the operation (catch-up bridge)
  contentSlide(pres, "Launch", C.ACCENT, "The words tell you the operation",
    [
      "3 friends share a $12 pizza equally.",
      "How much each?",
      "",
      "'Shared equally' means divide.",
      "12 ÷ 3 = $4 each.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.ACCENT });
      slide.addText("$12 shared into 3", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });
      drawBarModel(slide, lg.rightX + 0.35, lg.panelTopPadded + 0.55, lg.rightW - 0.70, 0.60, 3, 3,
        { fill: C.SECONDARY, segLabel: "$4", segFont: 17 });
      slide.addText("Which operation? Circle it:", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.30, w: lg.rightW - 0.30, h: 0.26,
        fontSize: 12.5, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });
      drawOperationChooser(slide, lg.rightX + 0.45, lg.panelTopPadded + 1.62, lg.rightW - 0.90, "÷");
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are learning to model real problems, choose the right operation, and decide whether we need an exact answer or an estimate.",
    [
      "I can choose the right operation for a word problem and say why.",
      "I can model and solve a problem, deciding if I need an exact answer or an estimate.",
      "I can build a simple budget and work out the money left over.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Words for today",
    [
      "Model = show the problem with a bar or diagram",
      "Operation = +, -, x or ÷",
      "Budget = a plan of money in and money out",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.SECONDARY });
      slide.addText("Match the words to the operation", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.32,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      const rows = [
        ["altogether, total", "+", C.PRIMARY],
        ["how many left", "-", C.SECONDARY],
        ["groups of, each", "x", C.ACCENT],
        ["shared equally", "÷", C.ALERT],
      ];
      const ry0 = lg.panelTopPadded + 0.56;
      rows.forEach((r, i) => {
        const ry = ry0 + i * 0.46;
        slide.addText(r[0], {
          x: lg.rightX + 0.30, y: ry, w: lg.rightW - 1.25, h: 0.40,
          fontSize: 13.5, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
          valign: "middle", margin: 0,
        });
        addTextOnShape(slide, r[1], {
          x: lg.rightX + lg.rightW - 0.85, y: ry + 0.02, w: 0.55, h: 0.36, rectRadius: 0.06,
          fill: { color: r[2] },
        }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
    }
  );

  // Slide 10: I Do #1 - model and choose the operation
  workedExSlide(pres, 2, "I Do", "Model it, then choose the operation",
    [
      "A fun run raises $240, shared equally between 6 classes.",
      "Model: a bar of $240 in 6 parts.",
      "'Shared equally' -> divide: 240 ÷ 6 = $40.",
      "Each class adds $10 of their own.",
      "'Adds' -> add: 40 + 10 = $50.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("$240 shared into 6 classes", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      drawBarModel(slide, lg.rightX + 0.30, lg.panelTopPadded + 0.52, lg.rightW - 0.60, 0.55, 6, 6,
        { fill: C.SECONDARY, segLabel: "$40", segFont: 12 });
      slide.addText("Step 1: shared equally -> divide", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.18, w: lg.rightW - 0.30, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });
      addTextOnShape(slide, "240 ÷ 6 = $40", {
        x: lg.rightX + 0.85, y: lg.panelTopPadded + 1.50, w: lg.rightW - 1.7, h: 0.46, rectRadius: 0.07,
        fill: { color: C.SECONDARY },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addText("Step 2: adds $10 -> add", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.04, w: lg.rightW - 0.30, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", margin: 0,
      });
      addTextOnShape(slide, "40 + 10 = $50 each", {
        x: lg.rightX + 0.85, y: lg.panelTopPadded + 2.36, w: lg.rightW - 1.7, h: 0.46, rectRadius: 0.07,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 11: I Do #2 - exact vs estimate
  workedExSlide(pres, 2, "I Do", "Exact or estimate? Match the question",
    [
      "$50 to spend. Shirt $19.95, shorts $24.95.",
      "'Enough?' -> ESTIMATE.",
      "About $20 + $25 = $45. Under $50, so yes.",
      "'Exact change?' -> EXACT.",
      "50 - 44.90 = $5.10.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      // Estimate card
      slide.addText("Do I have enough?", {
        x: lg.rightX + 0.25, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.5, h: 0.26,
        fontSize: 13.5, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });
      addTextOnShape(slide, "ESTIMATE: $20 + $25 = $45 < $50  YES", {
        x: lg.rightX + 0.30, y: lg.panelTopPadded + 0.44, w: lg.rightW - 0.60, h: 0.58, rectRadius: 0.07,
        fill: { color: C.SECONDARY },
      }, { fontSize: 13.5, fontFace: FONT_H, color: C.WHITE, bold: true });
      // Divider note
      slide.addText("same trip, different question", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.12, w: lg.rightW - 0.4, h: 0.24,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      });
      // Exact card
      slide.addText("Exactly how much change?", {
        x: lg.rightX + 0.25, y: lg.panelTopPadded + 1.42, w: lg.rightW - 0.5, h: 0.26,
        fontSize: 13.5, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", margin: 0,
      });
      addTextOnShape(slide, "EXACT: 19.95 + 24.95 = 44.90", {
        x: lg.rightX + 0.30, y: lg.panelTopPadded + 1.72, w: lg.rightW - 0.60, h: 0.50, rectRadius: 0.07,
        fill: { color: C.ALERT },
      }, { fontSize: 13.5, fontFace: FONT_H, color: C.WHITE, bold: true });
      addTextOnShape(slide, "50 - 44.90 = $5.10 change", {
        x: lg.rightX + 0.30, y: lg.panelTopPadded + 2.30, w: lg.rightW - 0.60, h: 0.50, rectRadius: 0.07,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 12-13: CFU + reveal - choose operation + estimate (6 boxes of 24)
  withReveal(
    () => cfuSlide(pres, "CFU", "6 boxes each hold 24 books",
      { technique: "Show Me Boards",
        question: "On your whiteboard: which operation finds the total number of books?\n\nWrite the operation and an estimate of the answer." },
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Multiply. Estimate 6 x 25 = 150 (exact 144).", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 14-15: We Do #1 + reveal - cupcake profit (two-step)
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Two steps: cupcake profit",
      [
        "With your partner.",
        "",
        "30 cupcakes sold at $2 each.",
        "$18 spent on ingredients.",
        "How much profit?",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.6, { strip: C.SECONDARY });
        slide.addText("Income and cost", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        slide.addText("Step 1: income = 30 x $2", {
          x: lg.rightX + 0.30, y: lg.panelTopPadded + 0.58, w: lg.rightW - 0.6, h: 0.30,
          fontSize: 13.5, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0,
        });
        slide.addShape("roundRect", {
          x: lg.rightX + 0.30, y: lg.panelTopPadded + 0.90, w: lg.rightW - 0.6, h: 0.42,
          fill: { color: C.WHITE }, line: { color: C.SECONDARY, width: 1.5 },
        });
        slide.addText("Step 2: profit = income - $18", {
          x: lg.rightX + 0.30, y: lg.panelTopPadded + 1.42, w: lg.rightW - 0.6, h: 0.30,
          fontSize: 13.5, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0,
        });
        slide.addShape("roundRect", {
          x: lg.rightX + 0.30, y: lg.panelTopPadded + 1.74, w: lg.rightW - 0.6, h: 0.42,
          fill: { color: C.WHITE }, line: { color: C.SECONDARY, width: 1.5 },
        });
      }),
    (slide) => {
      addTextOnShape(slide, "Income 30 x $2 = $60. Profit 60 - 18 = $42.", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slides 16-17: We Do #2 + reveal - build a budget
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Build the budget: cake stall",
      [
        "With your partner.",
        "",
        "Raised $120 at the cake stall.",
        "Spent $30 ingredients, $15 decorations.",
        "How much is left for the excursion?",
      ],
      NOTES_WEDO2_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.55, { strip: C.SECONDARY });
        slide.addText("Fill in the budget", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.30,
          fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "center", margin: 0,
        });
        drawBudgetTable(slide, lg.rightX + 0.35, lg.panelTopPadded + 0.58, lg.rightW - 0.70, [
          ["Income", "$120", C.SUCCESS],
          ["Total costs", "$30 + $15 = ?", C.ALERT],
          ["Left over", "?", C.PRIMARY],
        ]);
      }),
    (slide) => {
      addTextOnShape(slide, "Costs $30 + $15 = $45. Left over 120 - 45 = $75.", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
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
      { text: "choose the operation for each problem.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "decide exact or estimate, then solve.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "build the fundraiser budget.", options: { fontSize: 17, color: C.CHARCOAL } },
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
    addTextOnShape(s, "Model it. Let the words choose the operation. Ask: exact, or close enough?", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    // Reference key-word chips so the page is anchored, not text-only.
    const chips = [["altogether  +", C.PRIMARY], ["how many left  -", C.SECONDARY],
      ["groups of  x", C.ACCENT], ["shared equally  ÷", C.ALERT]];
    const cw = 2.05, gap = 0.18, startX = (10 - (4 * cw + 3 * gap)) / 2;
    chips.forEach((ch, i) => {
      addTextOnShape(s, ch[0], {
        x: startX + i * (cw + gap), y: panelY + 1.30, w: cw, h: 0.6, rectRadius: 0.08,
        fill: { color: ch[1] },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 19: Exit Ticket
  exitTicketSlide(pres,
    [
      "A cafe orders 6 trays of muffins at $7.50 each.",
      "Estimate the total, then work out the exact cost. Which would the cafe need to pay the bill?",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 20: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: what helps you decide which operation a problem needs?",
      scItems: [
        "I can choose the right operation for a word problem and say why.",
        "I can model and solve a problem, deciding if I need an exact answer or an estimate.",
        "I can build a simple budget and work out the money left over.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FourP_Lesson5_Modelling_And_Budgets.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Model word problems, choose the operation, and build a simple budget.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Read the problem, picture it, and let the words choose the operation: altogether and total mean add; how many left means subtract; groups of and each mean multiply; shared equally means divide. Then ask: do I need an exact answer, or is an estimate enough? A budget adds the costs, then subtracts from the income.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "8 friends share a $72 bill equally. 'Shared equally' means divide: 72 div 8 = $9 each.",
      y);

    y = addSectionHeading(doc, "Section 1 - Choose the operation (started for you)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Circle the operation, then solve.", y);
    y = addWriteLine(doc, "a)  5 trays of 12 cupcakes. How many cupcakes?   ( +  -  x  div )   Answer: ______", y);
    y = addWriteLine(doc, "b)  $35 saved, $12 spent. How much left?   ( +  -  x  div )   Answer: ______", y);
    y = addWriteLine(doc, "c)  $60 shared between 4 friends.   ( +  -  x  div )   Answer: ______", y);

    y = addSectionHeading(doc, "Section 2 - Exact or estimate?", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  You have $20. A book is $8.95 and a pen is $3.50. Do you have enough? Exact or estimate? ______", y);
    y = addWriteLine(doc, "b)  Exactly how much change from $20?   Answer: ______", y);

    y = addSectionHeading(doc, "Section 3 - Build the budget", y, { color: C.PRIMARY });
    y = addBodyText(doc, "A class raises $150 at a sausage sizzle. They spent $40 on sausages and $20 on bread.", y);
    y = addWriteLine(doc, "Income: ______      Total costs: ______      Left over: ______", y);

    y = addSectionHeading(doc, "Extension (optional)", y, { color: C.ACCENT });
    y = addWriteLine(doc, "Design your own fundraiser budget: one income, at least two costs, and the money left over.", y);
    y = addWriteLine(doc, "Income: ______   Costs: ______ , ______   Left over: ______", y);

    addPdfFooter(doc, `Lesson ${SESSION} | Modelling real problems and budgets | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Lesson 5 practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Lesson ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Choose the operation", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  multiply, 60 cupcakes (5 x 12).        b)  subtract, $23 ($35 - $12).        c)  divide, $15 each ($60 div 4).", y);

    y = addSectionHeading(doc, "Section 2 - Exact or estimate?", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  Estimate is enough: about $9 + $4 = $13, which is under $20, so yes.", y);
    y = addBodyText(doc, "b)  Exact needed: $8.95 + $3.50 = $12.45, change = $20 - $12.45 = $7.55.", y);

    y = addSectionHeading(doc, "Section 3 - Build the budget", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Income $150. Total costs $40 + $20 = $60. Left over $150 - $60 = $90.", y);

    y = addSectionHeading(doc, "Extension", y, { color: C.ACCENT });
    y = addBodyText(doc, "Budgets vary. Check: left over = income minus the total of all costs, and the leftover is not negative.", y);

    y = addTipBox(doc,
      "Watch for: students who grab an operation without reading the structure; students who calculate exactly when an estimate would do; students who subtract only one cost in a budget instead of totalling the costs first.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Lesson ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Year 8 Extension
  await (async () => {
    const doc = createPdf({ title: EXT_RES.name });
    let y = addPdfHeader(doc, EXT_RES.name, {
      subtitle: "Budgeting with unit rates and savings plans.",
      color: C.SECONDARY,
      lessonInfo: `Lesson ${SESSION} | Year 8 challenge | extends Year 6 VC2M5N09, VC2M6N09`,
    });
    y = addTipBox(doc,
      "A unit rate is the cost of ONE item - divide the price by the quantity to compare best buys. A savings plan divides the goal by the amount saved each week, then rounds UP to a whole number of weeks.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "A $200 tablet, saving $15 a week. Weeks = 200 div 15 = 13.3, so it takes 14 weeks (you cannot save for part of a week to reach the goal).",
      y);

    y = addSectionHeading(doc, "Section 1 - Best buy (unit rate)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  6 muffins for $9, or 8 muffins for $11.20. Which is the better buy? Show the unit price.", y);
    y = addWriteLine(doc, "    6 for $9 = $______ each.   8 for $11.20 = $______ each.   Better buy: ______", y);

    y = addSectionHeading(doc, "Section 2 - Savings plan", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  A $450 bike, saving $25 a week. How many weeks? ______", y);
    y = addWriteLine(doc, "b)  How much sooner if you save $30 a week instead? ______", y);

    y = addSectionHeading(doc, "Section 3 - Plan a class excursion budget", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "28 students. Bus $238. Entry $8.50 each. Work out the total cost and the cost per student.", y);
    y = addWriteLine(doc, "Total: ______      Per student: ______", y);

    y = addTipBox(doc,
      "Teacher answers: S1  6 for $9 = $1.50 each; 8 for $11.20 = $1.40 each; the 8-pack is the better buy.   S2  a) 450 div 25 = 18 weeks  b) 450 div 30 = 15 weeks, which is 3 weeks sooner.   S3  Entry 28 x 8.50 = $238; total 238 + 238 = $476; per student 476 div 28 = $17.",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, `Lesson ${SESSION} | Year 8 Extension | Budgeting and rates`);
    await writePdf(doc, path.join(OUT_DIR, EXT_RES.fileName));
    console.log("PDF written: " + EXT_RES.fileName);
  })();

  console.log("Lesson 5 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
