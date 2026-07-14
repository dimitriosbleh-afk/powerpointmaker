"use strict";

// Term 3 Maths Review (Year 6 Numeracy) - Session 4 of 4.
// PROBLEM SOLVING & MEASUREMENT. Pulls together Week 2 (VC2M5N08 + VC2M5N09),
// Week 4 (VC2M5M01 + VC2M6M01) and Week 7 (VC2M5N09 + VC2M6N09).
//   Block A - estimate to check, and choose the operation from a word problem.
//   Block B - convert between metric units (the x10 / division ladder).
//   Block C - solve a multi-step money problem (budget).
// Daily Review: Understanding & Operating with Fractions (prior, Week 8).
// Fluency: subtraction vertical algorithm with decimals (unit-wide focus).
// Unit variant fixed (variant 2) across all 4 sessions for cohesion.
// Catch-up: the launch starts from a rounding estimate everyone can do, and
// worksheet Section 1 re-grounds choosing the operation. No session assumes
// the one before.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", 2); // variant 2, fixed for the unit
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

const SESSION = 4;
const TOTAL = 4;
const UNIT_TITLE = "Term 3 Maths Review";
const FOOTER = `Term 3 Maths Review | Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/TermRev_Lesson4_Problem_Solving_And_Measurement";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Estimate to check, choose the operation, convert metric units, and solve money problems.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the Session 4 review sheet.");
const EXT_RES = makeSessionResource(SESSION,
  "Year 8 Extension",
  "Unit price and best buy (rates) - for students ready for more.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXT_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Small reusable visual ---------------------------------------------------

// A simple metric "ladder": chips for the units stacked top (bigger) to bottom
// (smaller), with the x10 / division rule shown between each step.
function drawLadder(slide, cx, topY, labels) {
  const chipW = 1.2, chipH = 0.48, gap = 0.5;
  labels.forEach((lab, i) => {
    const y = topY + i * (chipH + gap);
    addTextOnShape(slide, lab, {
      x: cx - chipW / 2, y, w: chipW, h: chipH, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    if (i < labels.length - 1) {
      slide.addText("x 10 down       ÷ 10 up", {
        x: cx - 1.7, y: y + chipH + 0.05, w: 3.4, h: 0.32,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  });
}

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to the last session of our Term 3 review. Today we use everything to solve real problems.
- We estimate to check our answers, choose the right operation for a problem, convert between units like metres and centimetres, and plan a budget.
- This is where all our number work pays off - turning a problem in words into maths we can solve.

DO:
- Have whiteboards, markers and the review sheet ready.
- Settle the room and focus attention before you click past the title.

TEACHER NOTES:
Session 4 of 4. The applied capstone. A review, so retrieve and firm up. Keep tying each problem back to a clear plan.

WATCH FOR:
- Students who rush to calculate before reading - slow them down to plan first.

[General: Title | Element: Planning]`;

const NOTES_RESOURCES = `SAY:
- These are today's materials.
- The review sheet is for the You Do near the end. The Year 8 extension is there for anyone who is flying.

DO:
- Print one review sheet and one answer key per student.
- Have whiteboards and markers ready for every check.
- Keep the Year 8 extension on hand for early finishers.

TEACHER NOTES:
One student review sheet, an answer key, and a Year 8 extension on unit price and best buys. Most of the lesson runs on whiteboards.

CATCH-UP NOTE:
A student who missed earlier work can still access today. The launch starts from a rounding estimate anyone can do, and Section 1 of the review sheet rebuilds choosing the operation from scratch. A returner needs only a whiteboard and one minute with you to start. No session in this unit assumes the student saw the one before.

[General: Resources | Element: Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Warming up our fraction work - comparing, adding and subtracting.
- Read each one and work it on your whiteboard.
- Show your thinking, not just the answer.

DO:
- Display the three prompts.
- Give about 2 minutes.
- Walk and listen for renaming to a common denominator.

TEACHER NOTES:
Daily Review is prior learning. Fractions keep ticking over from the last two sessions so they stay fresh past the unit.

WATCH FOR:
- Students who rename to compare and combine - secure.
- Students who add the bottoms - address it in the reveal.

[Stage 1: Daily Review | Element: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- 2/3 is 8/12 and 3/4 is 9/12, so 3/4 is larger.
- One half plus one quarter: rename to 2/4 plus 1/4, which is 3/4.
- Five sixths take away one half: rename 1/2 to 3/6, so 5/6 take away 3/6 is 2/6, which is 1/3.
- Tick or fix on your whiteboard.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
All three reward the same habit - make the pieces the same size before you compare or combine.

WATCH FOR:
- Students who self-correct quickly - secure.
- Students who add or subtract denominators - quick small group reminder.

[Stage 1: Daily Review Answers | Element: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Subtraction set out vertically, with decimals.
- Set each one out in columns on your whiteboard and subtract.
- Line up the decimal points so the places sit under each other.

DO:
- Display the three prompts.
- Give about 90 seconds.
- Scan for lined-up columns and the point kept in line.

TEACHER NOTES:
Fluency this whole unit is the subtraction algorithm with decimals. We use exactly this skill in the money problems later today.

WATCH FOR:
- Students who line up the points and fill gaps with zeros - secure.
- Students who right-align the digits instead of the points - fix it in the reveal.

[Stage 1: Fluency | Element: Retention and Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check yours.
- 8.1 take away 3.6 is 4.5.
- 12.00 take away 5.45 is 6.55.
- 9.3 take away 4.85 is 4.45. Writing 9.3 as 9.30 helps line it up.
- Tick and fix.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
The 9.3 minus 4.85 item is the key one. Fill the empty place with a zero to keep the columns honest. This is the same skill as the change in a money problem.

WATCH FOR:
- Students who self-correct - secure.
- Students whose columns drift - small group focus.

[Stage 1: Fluency Answers | Element: Retention and Recall]`;

const NOTES_LAUNCH = `SAY:
- Some of you may remember rounding to estimate. You are at the shop with three things in your basket: $4.95, $3.10 and $1.98. About how much altogether?
- We do not need the exact total to know if we have enough. We round to friendly numbers.
- $4.95 is about $5. $3.10 is about $3. $1.98 is about $2.
- Five plus three plus two is about ten dollars. So I will need about ten dollars. Estimating gives a quick, sensible answer - that drives today.

DO:
- Point to each price and its rounded value.
- Ask: about how much altogether? Thumbs up when you have an estimate.
- Bridge: 'estimating gives a close answer fast, and checks our exact answers are sensible'.

TEACHER NOTES:
This launch starts from a task anyone can do, then connects it to estimating and checking. It is also the catch-up bridge for any returning student.

WATCH FOR:
- Students who round and add confidently - strong prior knowledge.
- Students who try to add the exact cents - that is fine, but show how rounding is faster for a check.

[Stage: Launch | Element: Knowledge and Memory]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me.
- We are reviewing how to estimate, choose the right operation, convert between units, and solve money problems.
- Now the success criteria. Read each I can statement with me.

DO:
- Choral read the learning intention and the three I can statements.
- Point out that the first one is something everyone can do today.

TEACHER NOTES:
The first criterion is reachable for everyone - estimate to check. The second is the core target the exit ticket checks. The third stretches to a multi-step money plan.

WATCH FOR:
- Students who can repeat the language - tracking.

[General: LI/SC | Element: Planning]`;

const NOTES_VOCAB = `SAY:
- Three words for today.
- An estimate is a close, sensible answer - not the exact one. We use it to check.
- Convert means change to another unit, like 3 metres to 300 centimetres.
- A budget is a plan for spending money - what comes in and what goes out.

DO:
- Point to each word as you say it.
- Have students say 'an estimate is close, not exact' once together.

TEACHER NOTES:
Vocabulary comes after the learning intention. Estimate and convert are the two anchor words for today.

WATCH FOR:
- Students who can give an example of converting units - secure.
- Students who think an estimate must be exact - clarify it is a sensible close answer.

[General: Key Vocabulary | Element: Knowledge and Memory]`;

const NOTES_IDO_A = `SAY:
- A problem in words needs a plan before any maths. Listen: a school orders 6 boxes of markers, and each box has 48 markers. How many markers in total?
- I look for the structure. Equal groups of the same size - that is multiplication. Six groups of forty eight.
- So my number sentence is 6 times 48, which is 288 markers.
- Now I estimate to check it is sensible. Six times fifty is three hundred. My answer, 288, is close to that, so it looks right.

DO:
- Underline 'equal groups' and 'how many in total' as you read.
- Write the number sentence, then the estimate beside it.
- Have students chorus 'plan first, then calculate, then check'.

TEACHER NOTES:
This is VC2M5N09 plus VC2M5N08. The move is words to structure to operation, then estimate to check. Equal-groups language signals multiplication.

MISCONCEPTIONS:
- Misconception: students add 6 and 48 because both numbers are there.
  Why: they grab the numbers without finding the structure.
  Impact: a far too small answer that does not match an estimate.
  Quick correction: equal groups of the same size means multiply; check against 6 x 50.

WATCH FOR:
- Students who name the operation from the structure - secure.
- Students who guess the operation - prompt: is this equal groups, sharing, combining or taking away?

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_CFU_A_Q = `SAY:
- Quick check on your whiteboards.
- A baker makes 96 muffins and packs them into trays of 8. How many trays?
- Choose the operation, write the number sentence, then answer.

DO:
- Display the problem.
- Give 60 seconds.
- Walk and scan the whiteboards.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: show me on three, two, one, show.
- Scan for: 96 ÷ 8 = 12 trays.
PROCEED: If about 80 percent have it right, click to reveal and move to Block B.
PIVOT: Most likely misconception - students multiply 96 by 8.
- Reteach: making equal groups FROM a total is sharing, which is division. How many 8s fit into 96?
- Re-check: which operation makes equal groups from a total?

TEACHER NOTES:
The trap is choosing the wrong operation. Packing a total into equal trays is division.

WATCH FOR:
- Students who divide and get 12 - secure.
- Students who multiply - reteach sharing a total into equal groups.

[Stage 2: CFU | Element: Supported Application]`;

const NOTES_IDO_B = `SAY:
- Now we convert between metric units. The units step like a ladder. Each step down to a smaller unit, you multiply by ten; each step up to a bigger unit, you divide by ten.
- Metres, centimetres, millimetres are each ten apart. Going from metres to centimetres is two steps down... no, one family at a time. Metres to centimetres is times one hundred, because a metre is a hundred centimetres.
- Watch: 3.5 metres to centimetres. Metres to centimetres is times one hundred, so 3.5 times 100 is 350 centimetres.
- That is the same digits-move idea from Session 1. Bigger unit to smaller unit means more of them, so we multiply.

DO:
- Point up and down the ladder as you say multiply or divide.
- Write 3.5 m = 350 cm and say times one hundred.
- Have students whisper 'bigger to smaller, multiply' to a partner.

TEACHER NOTES:
This is VC2M5M01 plus VC2M6M01. Key equivalences: 1 km = 1000 m, 1 m = 100 cm, 1 cm = 10 mm, 1 kg = 1000 g, 1 L = 1000 mL. Tie the multiply or divide choice to whether the unit gets smaller or bigger.

MISCONCEPTIONS:
- Misconception: students divide when changing metres to centimetres because the number 'should get simpler'.
  Why: they do not link the operation to the size of the unit.
  Impact: 3.5 m becomes 0.035 cm.
  Quick correction: centimetres are smaller, so you need MORE of them - multiply.

WATCH FOR:
- Students who multiply going to a smaller unit - secure.
- Students who always multiply or always divide - back to the ladder, check the unit size.

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_WEDO_B_Q = `SAY:
- Your turn with a partner on your whiteboards.
- One: convert 2.5 metres to centimetres.
- Two: convert 3000 grams to kilograms.

DO:
- Display the two prompts.
- Give 75 seconds.
- Listen for the multiply or divide choice based on unit size.

TEACHER NOTES:
Metres to centimetres is times 100 (smaller unit, more of them). Grams to kilograms is divide by 1000 (bigger unit, fewer of them).

WATCH FOR:
- Pairs who pick the operation from the unit size - secure.
- Pairs unsure - prompt: is the new unit bigger or smaller?

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO_B_A = `SAY:
- Let us check.
- 2.5 metres to centimetres: centimetres are smaller, so multiply by 100. 2.5 times 100 is 250 centimetres.
- 3000 grams to kilograms: kilograms are bigger, so divide by 1000. 3000 divided by 1000 is 3 kilograms.
- Smaller unit, more of them; bigger unit, fewer of them.

DO:
- Click to reveal.
- Re-say the multiply or divide reason for each.

TEACHER NOTES:
If pairs went the wrong way, point at the ladder and ask whether the new unit is bigger or smaller before choosing.

WATCH FOR:
- Students who self-correct - secure.
- Students who reverse the operation - small group focus.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_IDO_C = `SAY:
- Now a money problem that takes more than one step. You have fifty dollars to plan a class party. You spend eighteen dollars on food and twelve dollars fifty on decorations. How much is left?
- I make a plan. First I add what I spent: eighteen dollars plus twelve dollars fifty is thirty dollars fifty.
- Then I take that from my budget: fifty dollars take away thirty dollars fifty is nineteen dollars fifty.
- So nineteen dollars fifty is left. A budget is just in and out - I add the spending, then subtract from what I have.

DO:
- Point to each line of the budget as you build it.
- Write the two steps clearly: total spent, then what is left.
- Have students whisper the plan: add the spending, then subtract.

TEACHER NOTES:
This is VC2M5N09 plus VC2M6N09. Multi-step: combine the costs, then find the change. The subtraction is the same decimal skill as the fluency warm-up.

MISCONCEPTIONS:
- Misconception: students subtract only one cost and forget the other.
  Why: they do one step and stop.
  Impact: too much money appears left over.
  Quick correction: add ALL the spending first, then subtract once.

WATCH FOR:
- Students who add the costs then subtract - secure.
- Students who do one step - prompt: have you counted every cost?

[Stage 2: I Do | Element: Explicit Teaching]`;

const NOTES_WEDO_C_Q = `SAY:
- Your turn with a partner. You have thirty dollars. You buy 3 books that cost seven dollars each. How much is left?
- Plan it: how much do the books cost altogether, then how much is left?
- Two steps. Show both.

DO:
- Display the problem.
- Give about 90 seconds.
- Listen for multiply then subtract.

TEACHER NOTES:
Two steps: 3 times 7 is 21 dollars for the books, then 30 take away 21 is 9 dollars left.

WATCH FOR:
- Pairs who multiply then subtract - secure.
- Pairs who subtract 7 once - prompt: how many books, and what do they cost altogether?

[Stage 3: We Do | Element: Supported Application]`;

const NOTES_WEDO_C_A = `SAY:
- Let us check.
- Three books at seven dollars each is 3 times 7, which is twenty one dollars.
- Thirty dollars take away twenty one dollars is nine dollars.
- So nine dollars is left. Multiply for the total cost, then subtract from what you have.

DO:
- Click to reveal.
- Point out the two steps: multiply, then subtract.

TEACHER NOTES:
If many did one step, do one more two-step money problem before the You Do.

WATCH FOR:
- Students who self-correct - ready for independent work.
- Students who only subtract once - enabling group for the You Do.

[Stage 3: We Do Reveal | Element: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own with the review sheet.
- Section 1 warms up by choosing the operation and estimating - that is for everyone.
- Section 2 converts metric units. Section 3 solves money problems.
- If you finish, try the challenge box, then the Year 8 extension.

DO:
- Distribute the review sheet.
- Circulate and listen for plan-first, then calculate, then check.
- Cold call one or two students to explain a plan.

TEACHER NOTES:
Different numbers from the We Do, same moves. Section 1 is the rebuild for any returning student.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Section 1 gives the operation choices and one worked plan. Students match each problem to an operation and estimate.
- Extra Notes: Sit with these students and ask: equal groups, sharing, combining, or taking away? This is also the rebuild for a returner.
EXTENDING PROMPT:
- Task: The challenge box asks students to plan a small party budget and stay under a set amount.
- Extra Notes: Students who are ready move on to the Year 8 extension on unit price and best buys.

WATCH FOR:
- Students who plan before calculating - secure.
- Students who grab numbers without a plan - prompt them to name the structure first.

[Stage 4: You Do | Element: Mastery and Application]`;

const NOTES_EXIT = `SAY:
- Last task, on your whiteboard, on your own.
- One: convert 2.5 kilometres to metres.
- Two: a class of 28 students each need 2 pencils. How many pencils altogether?

DO:
- Display the prompt.
- Give 3 minutes.
- Collect whiteboards or photograph them.

TEACHER NOTES:
Exit ticket assesses the core success criterion - convert a unit and choose the operation. Look for 2500 m and 56 pencils.

WATCH FOR:
- Students who convert and choose multiply - secure.
- Students who divide for the km, or add 28 and 2 - revisit at the start of next term.

[Stage 5: Exit Ticket | Element: Mastery and Application]`;

const NOTES_CLOSING = `SAY:
- Let us look at the success criteria again, and at our whole term of review.
- Show me thumbs up, sideways or down for each one.
- Turn and tell your partner: when you meet a word problem, what are the first two things you do before you calculate?

DO:
- Read each I can statement.
- Use thumbs.
- Cold call one or two students for the partner share.

TEACHER NOTES:
The threshold idea is plan first - find the structure and estimate - then calculate and check. That habit ties together every topic in this term's review. Celebrate how much they have brought back over the four sessions.

WATCH FOR:
- Strong thumbs up across all three - a strong end to the review.
- Sideways or down on the core criterion - note who to support early next term.

[General: Closing | Element: Retention and Recall]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Session 4: Problem solving and measurement",
    `Year 6 Numeracy | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slides 3-4: Daily Review + reveal - operating with fractions
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Operating with fractions",
      [
        "Which is larger: 2/3 or 3/4?",
        "Work out 1/2 + 1/4.",
        "Work out 5/6 - 1/2.",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "3/4 is larger       1/2 + 1/4 = 3/4       5/6 - 1/2 = 2/6 = 1/3", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 5-6: Fluency + reveal - subtraction with decimals
  withReveal(
    () => fluencySlide(pres, "Fluency: Subtraction with decimals",
      ["8.1 - 3.6", "12.00 - 5.45", "9.3 - 4.85"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "4.5        6.55        4.45", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 7: Launch - estimate by rounding (catch-up bridge)
  contentSlide(pres, "Launch", C.ACCENT, "About how much altogether?",
    [
      "Three things in the basket:",
      "$4.95,  $3.10,  $1.98.",
      "",
      "Round to friendly numbers.",
      "About $5 + $3 + $2.",
      "About $10 in total.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.ACCENT });
      slide.addText("Round, then add", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.14, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      const rows = [["$4.95", "$5"], ["$3.10", "$3"], ["$1.98", "$2"]];
      const ry0 = lg.panelTopPadded + 0.58;
      rows.forEach((r, i) => {
        slide.addText(r[0] + "   ->   " + r[1], {
          x: lg.rightX + 0.4, y: ry0 + i * 0.42, w: lg.rightW - 0.8, h: 0.36,
          fontSize: 17, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
        });
      });
      addTextOnShape(slide, "About $10", {
        x: lg.rightX + 0.9, y: ry0 + 1.4, w: lg.rightW - 1.8, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SECONDARY },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 8: LI/SC
  liSlide(pres,
    "We are reviewing how to estimate, choose the right operation, convert between units, and solve money problems.",
    [
      "I can estimate to check whether an answer is sensible.",
      "I can choose the operation and convert between metric units.",
      "I can solve a multi-step money problem and explain my plan.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 9: Key Vocabulary
  contentSlide(pres, "Key Vocabulary", C.SECONDARY, "Words for today",
    [
      "Estimate = a close, sensible answer (not exact)",
      "Convert = change to another unit (3 m = 300 cm)",
      "Budget = a plan for spending money",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.1, { strip: C.SECONDARY });
      slide.addText("Plan before you calculate", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.16, w: lg.rightW - 0.4, h: 0.34,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
      });
      const steps = ["1. Read and picture it", "2. Choose the operation", "3. Calculate, then check"];
      const sy0 = lg.panelTopPadded + 0.66;
      steps.forEach((st, i) => {
        addTextOnShape(slide, st, {
          x: lg.rightX + 0.35, y: sy0 + i * 0.46, w: lg.rightW - 0.7, h: 0.38, rectRadius: 0.06,
          fill: { color: C.PRIMARY },
        }, { fontSize: 13.5, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
    }
  );

  // Slide 10: I Do A - choose the operation + estimate to check
  workedExSlide(pres, 2, "I Do", "From words to a number sentence",
    [
      "6 boxes of markers, 48 in each box.",
      "How many markers in total?",
      "Equal groups of the same size -> multiply.",
      "6 x 48 = 288 markers.",
      "",
      "Check: 6 x 50 = 300, so 288 is sensible.",
    ],
    NOTES_IDO_A, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.PRIMARY });
      slide.addText("Plan, then calculate, then check", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 12.5, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      const chips = [
        ["Equal groups -> multiply", C.SECONDARY],
        ["6 x 48 = 288", C.PRIMARY],
        ["Estimate: 6 x 50 = 300", C.ALERT],
      ];
      const cy0 = lg.panelTopPadded + 0.6;
      chips.forEach((ch, i) => {
        addTextOnShape(slide, ch[0], {
          x: lg.rightX + 0.4, y: cy0 + i * 0.62, w: lg.rightW - 0.8, h: 0.48, rectRadius: 0.07,
          fill: { color: ch[1] },
        }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
    }
  );

  // Slides 11-12: CFU A + reveal - choose the operation
  withReveal(
    () => cfuSlide(pres, "CFU", "Choose the operation",
      { technique: "Show Me Boards",
        question: "A baker makes 96 muffins and packs them into trays of 8.\n\nHow many trays? Write the number sentence and answer." },
      NOTES_CFU_A_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "96 ÷ 8 = 12 trays", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 13: I Do B - convert metric units (ladder)
  workedExSlide(pres, 2, "I Do", "Convert between metric units",
    [
      "Bigger unit -> smaller unit: multiply.",
      "Smaller unit -> bigger unit: divide.",
      "",
      "1 m = 100 cm, so m -> cm is x 100.",
      "3.5 m = 3.5 x 100 = 350 cm.",
      "Same digits-move idea as Session 1.",
    ],
    NOTES_IDO_B, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.1, { strip: C.PRIMARY });
      slide.addText("The metric ladder", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      drawLadder(slide, lg.rightX + lg.rightW / 2, lg.panelTopPadded + 0.5, ["m", "cm", "mm"]);
    }
  );

  // Slides 14-15: We Do B + reveal - convert units
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Your turn: convert",
      [
        "With your partner.",
        "",
        "1.  2.5 m = ? cm",
        "2.  3000 g = ? kg",
        "",
        "Is the new unit bigger or smaller?",
      ],
      NOTES_WEDO_B_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.SECONDARY });
        slide.addText("Smaller unit -> more of them\nBigger unit -> fewer of them", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.2, w: lg.rightW - 0.4, h: 0.9,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
        });
        addTextOnShape(slide, "1 kg = 1000 g\n1 m = 100 cm", {
          x: lg.rightX + 0.5, y: lg.panelTopPadded + 1.25, w: lg.rightW - 1.0, h: 0.9, rectRadius: 0.08,
          fill: { color: C.PRIMARY },
        }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      }),
    (slide) => {
      addTextOnShape(slide, "2.5 m = 250 cm        3000 g = 3 kg", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_B_A);
    }
  );

  // Slide 16: I Do C - multi-step money problem (budget)
  workedExSlide(pres, 2, "I Do", "A class party budget",
    [
      "You have $50 to plan a party.",
      "Food costs $18. Decorations cost $12.50.",
      "Step 1 - add the spending: $18 + $12.50 = $30.50.",
      "Step 2 - take it from the budget: $50 - $30.50.",
      "",
      "$19.50 is left.",
    ],
    NOTES_IDO_C, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.9, { strip: C.PRIMARY });
      slide.addText("Add what you spend, then subtract", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 12.5, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      const rows = [
        ["Budget", "$50.00", C.SECONDARY],
        ["Food", "$18.00", C.PRIMARY],
        ["Decorations", "$12.50", C.PRIMARY],
        ["Left", "$19.50", C.ALERT],
      ];
      const ry0 = lg.panelTopPadded + 0.58;
      rows.forEach((r, i) => {
        const col = r[2];
        addTextOnShape(slide, r[0], {
          x: lg.rightX + 0.3, y: ry0 + i * 0.55, w: 2.1, h: 0.46, rectRadius: 0.06,
          fill: { color: col },
        }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
        addTextOnShape(slide, r[1], {
          x: lg.rightX + 2.5, y: ry0 + i * 0.55, w: 1.3, h: 0.46, rectRadius: 0.06,
          fill: { color: C.WHITE }, line: { color: col, width: 1.5 },
        }, { fontSize: 14, fontFace: FONT_H, color: col, bold: true });
      });
    }
  );

  // Slides 17-18: We Do C + reveal - two-step money problem
  withReveal(
    () => cfuSlide(pres, "We Do", "Two-step money problem",
      { technique: "Partner whiteboards",
        question: "You have $30. You buy 3 books that cost $7 each.\n\nHow much is left? Show both steps." },
      NOTES_WEDO_C_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "3 x $7 = $21, then $30 - $21 = $9 left", {
        x: 0.5, y: 4.5, w: 9.0, h: 0.55, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      }, { fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_C_A);
    }
  );

  // Slide 19: You Do
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: review sheet", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.25, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "choose the operation and estimate.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "convert the metric units.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "solve the money problems.", options: { fontSize: 17, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.16, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.4;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("Remember", {
      x: 0.7, y: panelY + 0.13, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
    });
    addTextOnShape(s, "Plan first: find the structure, choose the operation, then calculate and check.", {
      x: 1.0, y: panelY + 0.52, w: 8.0, h: 0.45, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });

    s.addText("Unit reminders:   1 km = 1000 m    1 m = 100 cm    1 cm = 10 mm    1 kg = 1000 g    1 L = 1000 mL", {
      x: 0.7, y: panelY + 1.18, w: 8.6, h: 0.34,
      fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 20: Exit Ticket
  exitTicketSlide(pres,
    [
      "Convert 2.5 km to metres.",
      "A class of 28 students each need 2 pencils. How many pencils altogether?",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // Slide 21: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: when you meet a word problem, what are the first two things you do before you calculate?",
      scItems: [
        "I can estimate to check whether an answer is sensible.",
        "I can choose the operation and convert between metric units.",
        "I can solve a multi-step money problem and explain my plan.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "TermRev_Lesson4_Problem_Solving_And_Measurement.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Estimate, choose the operation, convert metric units, and solve money problems.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });
    y = addTipBox(doc,
      "Plan before you calculate. Find the structure (equal groups, sharing, combining, taking away) to choose the operation, then estimate to check your answer is sensible. To convert units: a smaller unit needs MORE, so multiply; a bigger unit needs fewer, so divide.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "6 boxes of markers, 48 in each. Equal groups, so multiply: 6 x 48 = 288. Check: 6 x 50 = 300, so 288 is sensible.",
      y);

    y = addSectionHeading(doc, "Section 1 - Choose the operation and estimate (for everyone)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  72 stickers shared equally between 8 children.  Operation: ____  Answer: ____", y);
    y = addWriteLine(doc, "b)  5 packs of cards, 12 in each pack.  Operation: ____  Answer: ____", y);
    y = addWriteLine(doc, "c)  Estimate 3.9 + 6.1 + 2.2 by rounding first.  About ____", y);

    y = addSectionHeading(doc, "Section 2 - Convert metric units", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  4 m = ______ cm", y);
    y = addWriteLine(doc, "b)  2.5 km = ______ m", y);
    y = addWriteLine(doc, "c)  2000 mL = ______ L", y);
    y = addWriteLine(doc, "d)  500 g = ______ kg", y);

    y = addSectionHeading(doc, "Section 3 - Money problems", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  You have $40. You spend $15 and $8.50. How much is left?  ______", y);
    y = addWriteLine(doc, "b)  4 tickets cost $9 each. What is the total?  ______", y);

    y = addSectionHeading(doc, "Challenge (optional)", y, { color: C.ACCENT });
    y = addWriteLine(doc, "Plan a party with $25: list two things to buy and stay under budget.  ______", y);

    addPdfFooter(doc, `Session ${SESSION} | Problem solving and measurement | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the Session 4 review sheet.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 - Choose the operation and estimate", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  Division, 72 ÷ 8 = 9.   b)  Multiplication, 5 x 12 = 60.   c)  About 12 (4 + 6 + 2).", y);

    y = addSectionHeading(doc, "Section 2 - Convert metric units", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  400 cm        b)  2500 m        c)  2 L        d)  0.5 kg", y);

    y = addSectionHeading(doc, "Section 3 - Money problems", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  Spent $15 + $8.50 = $23.50; $40 - $23.50 = $16.50 left.   b)  4 x $9 = $36.", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.ACCENT });
    y = addBodyText(doc, "Answers vary. Check the two costs add to $25 or less, e.g. $14 food + $9 decorations = $23, under $25.", y);

    y = addTipBox(doc,
      "Watch for: students who grab numbers without finding the structure; students who multiply when they should divide (or the reverse) in conversions; students who do only one step of a multi-step money problem.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Year 8 Extension
  await (async () => {
    const doc = createPdf({ title: EXT_RES.name });
    let y = addPdfHeader(doc, EXT_RES.name, {
      subtitle: "Unit price and best buy (rates).",
      color: C.SECONDARY,
      lessonInfo: `Session ${SESSION} | Year 8 challenge | extends Year 6 VC2M6N09`,
    });
    y = addTipBox(doc,
      "A rate compares two quantities, like dollars per item or cents per gram. To find the best buy, work out the price for the SAME amount of each option - usually the price per single unit - then compare.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Worked example", y, { color: C.SECONDARY });
    y = addBodyText(doc,
      "Pack A: 6 drinks for $9. Pack B: 4 drinks for $5. Price per drink: A = 9 / 6 = $1.50; B = 5 / 4 = $1.25. Pack B is the better buy.",
      y);

    y = addSectionHeading(doc, "Section 1 - Unit price", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  5 apples for $4. Price per apple = ______", y);
    y = addWriteLine(doc, "b)  3 m of ribbon for $7.50. Price per metre = ______", y);
    y = addWriteLine(doc, "c)  500 g of rice for $2. Price per 100 g = ______", y);

    y = addSectionHeading(doc, "Section 2 - Best buy", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)  8 pens for $6, or 5 pens for $4. Which is the better buy? ______", y);
    y = addWriteLine(doc, "b)  2 L of juice for $3.60, or 1.5 L for $2.85. Which is the better buy? ______", y);

    y = addSectionHeading(doc, "Section 3 - Reasoning", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "Why is the bigger pack not always the better buy? Explain with an example.", y);
    y = addWriteLine(doc, "_____________________________________________________________________", y);

    y = addTipBox(doc,
      "Teacher answers: S1  a) $0.80  b) $2.50  c) $0.40.   S2  a) 8 for $6 = $0.75 each beats 5 for $4 = $0.80 each.  b) 2 L = $1.80/L beats 1.5 L = $1.90/L.   S3  A bigger pack can have a higher price per unit; you must compare per unit, not per pack.",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, `Session ${SESSION} | Year 8 Extension | Rates and best buy`);
    await writePdf(doc, path.join(OUT_DIR, EXT_RES.fileName));
    console.log("PDF written: " + EXT_RES.fileName);
  })();

  console.log("Session 4 build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
