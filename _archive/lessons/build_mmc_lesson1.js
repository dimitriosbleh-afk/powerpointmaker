"use strict";

// Metric Measurement & Unit Conversion (Year 6 Numeracy)
// Session 1 of 5: Choosing the right metric unit (length, mass & capacity).
// VC2M5M01 — choosing appropriate units; ordering units largest to smallest;
//            recognising the significance of the metric prefixes.
// Daily Review: Decimal Operations and Place Value Mastery (prior learning).
// Fluency: Division bracket (long-division layout). NOT the bus stop method.
//
// Unit routine (consistent across all 5 sessions):
//   SET THE STAIRS  ->  COUNT THE STEPS  ->  MULTIPLY OR DIVIDE.
// Each session is deliberately self-contained: a student who misses one or
// two sessions can still access the next because every conversion session
// re-teaches the staircase method from scratch with fresh worked examples.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

// Whole-unit palette: every session of this unit uses the SAME variant for
// cohesion (megaprompt theme-cohesion rule). Variant 0 chosen because its
// ACCENT passes WCAG AA on white (the v4/v5 olive ACCENT does not).
const UNIT_VARIANT = 0;
const T = createTheme("numeracy", "grade56", UNIT_VARIANT);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  dailyReviewSlide, fluencySlide,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  withReveal, addRevealAnswerBar,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const TOTAL = 5;
const UNIT_TITLE = "Metric Measurement & Conversion";
const FOOTER = `Metric Measurement | Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/MMC_Session1_Choosing_Metric_Units";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Choosing Metric Units Practice",
  "Name units, choose the best unit for each job, and order units largest to smallest.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the choosing-units practice sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome, mathematicians. This week we become measurement experts.
- Today our job is to pick the right unit for the job, the way a builder, a nurse or a chef would.
- By the end you will be able to look at something and say straight away which metric unit fits it best.

DO:
- Have whiteboards and markers ready on desks.
- Keep the printed practice sheet face down until the You Do section.

TEACHER NOTES:
Session 1 of 5 in the Metric Measurement and Conversion unit. Today is about choosing and ordering units, not heavy conversion arithmetic. That comes next session.

WATCH FOR:
- Students who only know centimetres and kilograms. We widen the picture to length, mass and capacity.

[General: Title | VTLM 2.0: Attention, focus and regulation]`;

const NOTES_RESOURCES = `SAY:
- Here is what we are using today.
- You will need a whiteboard for most of the lesson, and the practice sheet near the end.

DO:
- Print one practice sheet and one answer key per student who needs paper.
- Have whiteboards, markers, and a few classroom objects to measure (pencil, water bottle, ruler).

TEACHER NOTES:
One student resource today plus an answer key. The choosing and ordering work is done mostly on whiteboards and out loud, so very little printing is needed.

[General: Resources | VTLM 2.0: Planning]`;

const NOTES_OVERVIEW = `SAY:
- This slide is for me, not the class. I can skip past it once I have read it.

DO:
- Read before teaching. Hide or skip during the lesson if projecting.

TEACHER NOTES:
Five-session arc. Session 1: choose and order units. Session 2: convert length. Session 3: convert mass. Session 4: convert capacity and decimal equivalence. Session 5: choose, convert and compare in real problems. Every conversion session uses the same routine - Set the stairs, Count the steps, Multiply or divide - so the language stays familiar week to week. Each session is deliberately self-contained: if a student misses one or two sessions they can still access the next, because every session re-teaches its core move with fresh worked examples. Nothing in a later session assumes the exact problems from an earlier one. The Year 8 Extension Challenge sits with Session 5 for students who are ready to push further at any point.

[General: Teacher Overview | VTLM 2.0: Planning]`;

const NOTES_DR_Q = `SAY:
- Quick warm up on decimals we already know.
- Read each one, then write your answer on your whiteboard.
- These are from earlier learning, not today's new work.

DO:
- Display the three prompts.
- Allow about 90 seconds.
- Scan for decimal place value confusion.

TEACHER NOTES:
Daily Review revisits Decimal Operations and Place Value Mastery. This is prior learning and retrieval. Strong decimal place value helps later this week when 1.25 metres becomes 125 centimetres.

WATCH FOR:
- Students who line up the decimal points before adding - secure.
- Students who say the 6 in 12.36 is worth 6 - they have missed the hundredths place.

[Stage 1: Daily Review | VTLM 2.0: Retention and recall]`;

const NOTES_DR_A = `SAY:
- Let us check together.
- 4.7 plus 3.85 is 8.55. Line up the points and add.
- The 6 in 12.36 is worth six hundredths, or 0.06.
- Smallest to largest is 0.06, then 0.6, then 0.66.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
Place value of decimals is the bridge to this week. Keep the language of tenths and hundredths alive.

WATCH FOR:
- Confident self-correction - secure.
- Students who ordered by digits not value - revisit comparing decimals.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency now. Division using the division bracket.
- Set each one out with the division bracket and work from the left-most digit.
- Whisper your answer, then write it.

DO:
- Display the three divisions.
- Allow about 80 seconds.
- Watch the carrying of remainders to the next digit.

TEACHER NOTES:
Fluency builds division using the division bracket, the same focus across all five sessions. This is the long-division layout. Please call it the division bracket. Brisk retrieval, not new teaching.

WATCH FOR:
- Students who carry the remainder into the next digit cleanly - secure.
- Students who drop a remainder - reteach the carry step with the division bracket.

[Stage 1: Fluency | VTLM 2.0: Retention and recall]`;

const NOTES_FLUENCY_A = `SAY:
- Check your division brackets.
- 138 divided by 6 is 23.
- 192 divided by 8 is 24.
- 245 divided by 5 is 49.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
138 divided by 6 needs a carry: 13 divided by 6 is 2 remainder 1, carry the 1 to make 18, 18 divided by 6 is 3.

WATCH FOR:
- Clean carrying of remainders - secure.
- Answers like 21 for 138 divided by 6 - the remainder was dropped.

[Stage 1: Fluency Answers | VTLM 2.0: Retention and recall]`;

const NOTES_LAUNCH = `SAY:
- Picture three jobs. You measure the length of your pencil. You measure the distance to the next town. You measure how much water fills your drink bottle.
- Would you use the same unit for all three? Of course not.
- Turn and tell your partner which unit you would choose for each, and why.

DO:
- Display the three jobs.
- Give 40 seconds of partner talk.
- Take two or three responses. Listen for the words centimetre, kilometre, millilitre.

TEACHER NOTES:
The launch activates what students already know about units from earlier years and connects it to today's idea of choosing the best unit for the job. Do not correct yet. Collect their thinking; we sharpen it in the I Do.

WATCH FOR:
- Students who match a small object to a small unit - exactly the instinct we build today.
- Students who say they would measure a road in centimetres - we will return to why that is awkward.

[Stage 1: Launch | VTLM 2.0: Knowledge and memory]`;

const NOTES_LI_SC = `SAY:
- Here is what we are learning today.
- We are learning to choose and order the right metric units for length, mass and capacity.
- Read the success criteria with me.

DO:
- Choral read the learning intention.
- Read each I can statement.

TEACHER NOTES:
SC1 is reachable for everyone today - naming units. SC2 is the core target the exit ticket assesses - choosing a sensible unit. SC3 stretches into ordering units and explaining the choice. Tier labels stay off the slide.

WATCH FOR:
- Students repeating the language back - tracking with us.

[General: LI/SC | VTLM 2.0: Planning]`;

const NOTES_VOCAB = `SAY:
- Before we choose units, let us meet the prefixes that build their names.
- Kilo means a thousand. A kilometre is a thousand metres. A kilogram is a thousand grams.
- Centi means one hundredth. A centimetre is one hundredth of a metre.
- Milli means one thousandth. A millimetre is one thousandth of a metre, and a millilitre is one thousandth of a litre.

DO:
- Point to each prefix as you say it.
- Have students say kilo, centi, milli back to you.

TEACHER NOTES:
This addresses the elaboration about recognising the significance of the prefixes. The same prefix means the same size jump no matter the attribute. Keep it concrete: kilo big, milli tiny.

WATCH FOR:
- Students who connect kilo to a thousand in other words like kilobyte - secure.
- Students who mix up centi and milli - anchor centi to the 100 centimetres on a metre ruler.

[General: Key Vocabulary | VTLM 2.0: Knowledge and memory]`;

const NOTES_IDO1 = `SAY:
- Let us look at the family of length units together and put them in order.
- The biggest is the kilometre, for long distances like roads.
- Then the metre, about the height of a door handle to the floor.
- Then the centimetre, about the width of your fingernail.
- The smallest is the millimetre, the thickness of a fingernail or a coin edge.
- Notice the staircase. We always set them out largest at the top, smallest at the bottom. That order never changes, and it is the first move we make all week.

DO:
- Point to each step from top to bottom as you name it.
- Stress the words largest at the top, smallest at the bottom.

TEACHER NOTES:
This is the Set the stairs move - ordering units largest to smallest. It addresses the elaboration on ordering metric units. The same staircase reappears every session, so build the habit now.

MISCONCEPTIONS:
- Misconception: a millimetre is bigger than a centimetre because milli sounds like a lot.
  Why: the word million is large, so milli feels large.
  Impact: students order units the wrong way and convert in the wrong direction later.
  Quick correction: milli means one thousandth - the smallest of the four. Point to the tiny millimetre marks on a ruler.

WATCH FOR:
- Students who order km, m, cm, mm without help - secure and ready for depth.
- Students unsure where cm and mm sit - use a real ruler.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_IDO2 = `SAY:
- Now I choose the best unit for three different jobs. Watch how I decide.
- For the length of a pencil, a kilometre is silly and a metre is too big. Centimetres fit nicely, so I choose centimetres.
- For the distance between two towns, centimetres would take forever to count. Kilometres are made for long distances, so I choose kilometres.
- For the water in a drink bottle, length units do not help at all. That is capacity, so I choose millilitres or litres.
- The trick is to ask: what am I measuring, and which unit gives a sensible number that is not tiny and not enormous?

DO:
- Point to each job and name the chosen unit.
- Hold up a real pencil and bottle if you have them.

TEACHER NOTES:
This is the heart of the lesson and the exit ticket target. The decision is about a sensible-sized number for the context. It also addresses choosing the attribute - length, mass or capacity - before the unit.

MISCONCEPTIONS:
- Misconception: you can use any unit as long as you measure carefully.
  Why: technically true, but it ignores practicality.
  Impact: students measure a road in centimetres and produce an unwieldy number.
  Quick correction: ask whether the number is sensible. 400000 centimetres for a road is a clue to choose kilometres.

WATCH FOR:
- Students who name the attribute first, then the unit - secure.
- Students who jump straight to a unit without thinking about the job - slow them down with the what am I measuring question.

[Stage 2: I Do | VTLM 2.0: Explicit Teaching]`;

const NOTES_CFU1_Q = `SAY:
- Quick check on your whiteboards.
- Which unit would you choose to measure the length of your classroom?
- Show me your unit on three, two, one.

DO:
- Display the prompt.
- Allow about 30 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your unit on three, two, one, show."
- Scan for: metres. A classroom is a few metres across.
PROCEED: If 80 percent show metres, click to reveal and move on.
PIVOT: Most likely misconception - students choose centimetres because that is the unit they use most.
- Reteach: "Centimetres would give a number in the hundreds. Metres give a sensible number like 8. Which is easier to use?"
- Re-check: "Now choose a unit for the length of the basketball court."

TEACHER NOTES:
Checks the choose-a-sensible-unit habit before guided practice.

WATCH FOR:
- Metres with a reason - secure.
- Centimetres - ask what the number would look like.

[Stage 2: CFU | VTLM 2.0: Supported Application]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me now. Same thinking, four everyday jobs.
- For each one, decide the attribute first - length, mass or capacity - then choose a sensible unit.
- Talk it through with your partner and write your unit beside each job.

DO:
- Keep the four jobs visible.
- Walk and scan. Prompt with "what are you measuring first" where needed.
- Take 60 seconds before revealing.

TEACHER NOTES:
Guided practice. Students name the attribute, then the unit. The reveal confirms a sensible choice, not the only possible choice. Accept litres or millilitres for the bottle if justified.

WATCH FOR:
- Pairs who name the attribute before the unit - secure.
- Pairs who guess a unit with no reason - redirect to the attribute question.

[Stage 3: We Do | VTLM 2.0: Supported Application]`;

const NOTES_WEDO_A = `SAY:
- Let us check our choices.
- The height of a door - metres. The mass of an apple - grams. The water in a bath - litres. The length of an ant - millimetres.
- Each one gives a sensible number that is easy to use.

DO:
- Click to reveal.
- Ask one pair to justify a choice.

TEACHER NOTES:
Reasonable alternatives are fine if justified, for example centimetres for the door. The point is a sensible number for the job.

WATCH FOR:
- Choices with a clear reason - secure.
- An ant measured in centimetres - fine, but millimetres give a friendlier number.

[Stage 3: We Do Reveal | VTLM 2.0: Supported Application]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. A student says they will measure the distance they walk to school in millimetres.
- Do not just say no. Decide: is that a sensible choice?
- Thumbs up if sensible, thumbs down if not.

DO:
- Display the claim.
- Give 15 seconds of thinking time.

CFU CHECKPOINT:
Technique: Thumbs Up or Down
Script:
- Say: "Thumbs up if sensible, thumbs down if not."
- Scan for: thumbs down. Millimetres would give a number in the millions.
PROCEED: If 80 percent thumbs down with a reason, click to reveal.
PIVOT: Most likely misconception - students think any unit works so it must be sensible.
- Reteach: "A kilometre is for long distances. Millimetres for a walk would be millions of marks. Choose the unit that gives a sensible number."
- Re-check: "What unit should they use, and roughly what number would they get?"

TEACHER NOTES:
This checks the threshold idea: choosing a unit is about a sensible-sized number, not just about being able to measure. Distance to school is best in kilometres or metres.

WATCH FOR:
- Confident thumbs down naming kilometres - secure.
- Thumbs up - revisit the sensible-number idea with the road example.

[Stage 2: CFU Hinge | VTLM 2.0: Supported Application]`;

const NOTES_YOUDO = `SAY:
- Time to work on your own. Take the practice sheet.
- Section one, name a unit for each attribute. Section two, choose the best unit for each job. Section three, order a set of units from largest to smallest.
- Always ask: what am I measuring, and is my number sensible?

DO:
- Distribute the practice sheet.
- Circulate. Sit with the enabling group first.
- Watch for students naming the attribute before the unit.

TEACHER NOTES:
The sheet sequences naming, choosing and ordering. Section two is the threshold - choosing a sensible unit for the job.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the worked first item and the unit-ladder box at the top of the sheet. Match each job to one of just three units offered.
- Extra Notes: Do the first choosing item together with this group before they continue.
EXTENDING PROMPT:
- Task: The challenge box - for two jobs, name a second unit that would also work, and explain when each unit would be the better choice.
- Extra Notes: Push for a reason that names the context and the size of the number.

WATCH FOR:
- Students who name the attribute before the unit - secure.
- Students who pick a unit with no reason - redirect to the what am I measuring question.

[Stage 4: You Do | VTLM 2.0: Mastery and application]`;

const NOTES_EXIT = `SAY:
- Last task, on your own, on your whiteboard.
- One, choose the best unit to measure the mass of a school bag, and explain your choice in a few words.
- Two, order these units from largest to smallest: centimetre, kilometre, millimetre, metre.

DO:
- Display the prompt.
- Allow about 3 minutes.
- Collect whiteboards or take a photo for your records.

TEACHER NOTES:
Exit ticket assesses SC2 (choosing a sensible unit) and reaches into SC3 (ordering units). Internal target is SC2. Do not display any SC label to students.

WATCH FOR:
- Kilograms for the bag with a reason, and km, m, cm, mm in order - secure.
- A school bag measured in tonnes or milligrams - revisit sensible numbers in Session 2.

[Stage 5: Exit Ticket | VTLM 2.0: Mastery and application]`;

const NOTES_CLOSING = `SAY:
- Look back at our success criteria.
- Show me thumbs for each one - got it, getting there, or need more practice.
- Turn and tell your partner: how do you decide which unit to use?

DO:
- Read each I can statement.
- Use thumbs to self-assess.
- Note who is still unsure for tomorrow.

TEACHER NOTES:
The big idea is choosing the unit that gives a sensible number for the job. Students who can explain that have the understanding we want before we start converting next session.

WATCH FOR:
- Students who explain the sensible-number idea - secure.
- Students who only name units - revisit choosing in Session 2's launch.

[General: Closing | VTLM 2.0: Mastery and application]`;

// ─── Visual helpers (build-script local) ─────────────────────────────────────

// The metric staircase: units ordered largest (top) to smallest (bottom),
// with the multiply/divide factor between each step. This is the unit's hero
// visual and reappears in every conversion session. Drawn in a right column.
function metricStaircase(slide, lg, opts) {
  const o = opts || {};
  const x = lg.rightX;
  const w = lg.rightW;
  const y0 = lg.panelTopPadded;
  const bottom = o.bottom != null ? o.bottom : lg.safeBottom;
  const color = o.color || C.PRIMARY;
  const steps = o.steps; // [{u}, {u, f}, ...] largest -> smallest
  const cardH = bottom - y0;
  addCard(slide, x, y0, w, cardH, { strip: color });

  slide.addText(o.title || "Metric staircase", {
    x: x + 0.18, y: y0 + 0.10, w: w - 0.36, h: 0.28,
    fontSize: 14.5, fontFace: FONT_H, color, bold: true, align: "center", margin: 0,
  });

  const innerTop = y0 + 0.44;
  const keyH = 0.66;
  const region = bottom - innerTop - keyH - 0.06;
  const n = steps.length;
  const pillH = 0.38;
  const gap = (region - n * pillH) / (n - 1);
  const pillW = w - 1.1;
  const pillX = x + (w - pillW) / 2;

  let cy = innerTop;
  steps.forEach((st, i) => {
    if (i > 0) {
      slide.addText([
        { text: "× " + st.f, options: { color: C.SUCCESS, bold: true } },
        { text: "      ÷ " + st.f, options: { color: C.ALERT, bold: true } },
      ], {
        x: pillX, y: cy - gap, w: pillW, h: gap,
        fontSize: 13, fontFace: FONT_B, align: "center", valign: "middle", margin: 0,
      });
    }
    addTextOnShape(slide, st.u, {
      x: pillX, y: cy, w: pillW, h: pillH, rectRadius: 0.08,
      fill: { color: i === 0 ? color : C.SECONDARY },
    }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    cy += pillH + gap;
  });

  // Direction key
  const keyY = bottom - keyH + 0.02;
  addCard(slide, x + 0.14, keyY, w - 0.28, keyH - 0.08, { fill: C.BG_LIGHT });
  slide.addText([
    { text: "Down to a smaller unit:  × multiply", options: { color: C.SUCCESS, bold: true, breakLine: true } },
    { text: "Up to a bigger unit:  ÷ divide", options: { color: C.ALERT, bold: true } },
  ], {
    x: x + 0.24, y: keyY + 0.04, w: w - 0.48, h: keyH - 0.16,
    fontSize: 12, fontFace: FONT_B, valign: "middle", margin: 0,
  });
}

// A stacked "choice panel": each row is a job + the sensible unit chosen.
function unitChoicePanel(slide, lg, rows, opts) {
  const o = opts || {};
  const x = lg.rightX;
  const w = lg.rightW;
  const y = lg.panelTopPadded;
  const bottom = o.bottom != null ? o.bottom : lg.safeBottom;
  const gap = 0.14;
  const cardH = (bottom - y - gap * (rows.length - 1)) / rows.length;
  rows.forEach((r, i) => {
    const cy = y + i * (cardH + gap);
    addCard(slide, x, cy, w, cardH, { strip: r.color || C.SECONDARY });
    slide.addText(r.job, {
      x: x + 0.20, y: cy + 0.08, w: w - 1.5, h: cardH - 0.16,
      fontSize: 13.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "middle", margin: 0,
    });
    addTextOnShape(slide, r.unit, {
      x: x + w - 1.28, y: cy + (cardH - 0.42) / 2, w: 1.1, h: 0.42, rectRadius: 0.08,
      fill: { color: r.color || C.SECONDARY },
    }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
  });
}

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  titleSlide(pres, UNIT_TITLE, "Session 1: Choosing the right unit",
    `Year 6 Numeracy | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3. Unit overview (teacher-facing)
  contentSlide(pres, "For the teacher", C.MUTED,
    "Unit overview — 5 sessions",
    [
      "Routine every session: Set the stairs → Count the steps → Multiply or divide.",
      "S1: choose & order units (today).   S2: convert length.   S3: convert mass.",
      "S4: convert capacity & equivalence.   S5: choose, convert & compare.",
      "Each session re-teaches its core move — a student who misses one can still access the next.",
      "Year 8 Extension Challenge ships with Session 5 for early finishers any week.",
    ],
    NOTES_OVERVIEW, FOOTER);

  // 4-5. Daily Review + reveal
  withReveal(
    () => dailyReviewSlide(pres,
      "Daily Review: Decimals",
      [
        "Calculate 4.7 + 3.85",
        "What is the value of the 6 in 12.36?",
        "Order from smallest: 0.6, 0.06, 0.66",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide,
        ["8.55", "0.06 (six hundredths)", "0.06, 0.6, 0.66"],
        { color: C.SUCCESS });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // 6-7. Fluency + reveal (division bracket)
  withReveal(
    () => fluencySlide(pres, "Fluency: Division bracket",
      ["138 ÷ 6", "192 ÷ 8", "245 ÷ 5"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, ["23", "24", "49"], { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // 8. Launch
  contentSlide(pres, "Launch", C.SECONDARY,
    "Same unit for all three?",
    [
      "The length of your pencil.",
      "The distance to the next town.",
      "The water in your drink bottle.",
      "Which unit fits each job? Tell your partner.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      unitChoicePanel(slide, lg, [
        { job: "Pencil length", unit: "?", color: C.PRIMARY },
        { job: "Distance to town", unit: "?", color: C.SECONDARY },
        { job: "Water in a bottle", unit: "?", color: C.ALERT },
      ]);
    });

  // 9. LI / SC
  liSlide(pres,
    "We are learning to choose and order the right metric units for length, mass and capacity.",
    [
      "I can name metric units for length, mass and capacity.",
      "I can choose a sensible unit to measure a given object.",
      "I can order metric units from largest to smallest and explain my choice.",
    ],
    NOTES_LI_SC, FOOTER);

  // 10. Key Vocabulary — the prefixes
  contentSlide(pres, "Key Vocabulary", C.PRIMARY,
    "The prefixes build the names",
    [
      "kilo = a thousand",
      "centi = one hundredth",
      "milli = one thousandth",
      "",
      "Same prefix, same size jump — for length, mass or capacity.",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      unitChoicePanel(slide, lg, [
        { job: "kilometre = 1000 metres", unit: "kilo", color: C.PRIMARY },
        { job: "centimetre = 1/100 m", unit: "centi", color: C.SECONDARY },
        { job: "millimetre = 1/1000 m", unit: "milli", color: C.ALERT },
      ]);
    });

  // 11. I Do 1 — order the length units (the staircase)
  workedExSlide(pres, 2, "I Do", "Set the stairs: order the units",
    [
      "First move every week: set the stairs.",
      "Largest unit at the top, smallest at the bottom.",
      "",
      "Length: km, then m, then cm, then mm.",
      "",
      "This order never changes.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      metricStaircase(slide, lg, {
        title: "Length staircase",
        color: C.PRIMARY,
        steps: [{ u: "km" }, { u: "m", f: 1000 }, { u: "cm", f: 100 }, { u: "mm", f: 10 }],
      });
    });

  // 12. I Do 2 — choose the best unit for the job
  workedExSlide(pres, 2, "I Do", "Choose the unit that fits the job",
    [
      "Ask: what am I measuring?",
      "Then: which unit gives a sensible number?",
      "",
      "Pencil → centimetres.",
      "Distance between towns → kilometres.",
      "Water in a bottle → millilitres or litres.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      unitChoicePanel(slide, lg, [
        { job: "Length of a pencil", unit: "cm", color: C.PRIMARY },
        { job: "Distance between towns", unit: "km", color: C.SECONDARY },
        { job: "Water in a bottle", unit: "mL / L", color: C.ALERT },
      ]);
    });

  // 13. CFU 1 + reveal
  withReveal(
    () => cfuSlide(pres, "CFU", "Which unit?", "Show Me Boards",
      "Which unit would you choose to measure the length of your classroom?",
      NOTES_CFU1_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "Metres — a classroom is a few metres across (a sensible number like 8).",
        { color: C.SUCCESS, label: "Answer", fontSize: 20 });
    }
  );

  // 14-15. We Do + reveal
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Choose a unit for each job",
      [
        "With your partner.",
        "",
        "Step 1: what are you measuring — length, mass or capacity?",
        "Step 2: choose a sensible unit.",
        "",
        "Write your unit beside each job.",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        // Stop the panel above the reveal answer bar (clears y 4.3) so the
        // revealed bar on the duplicate slide does not cover the 4th card.
        unitChoicePanel(slide, lg, [
          { job: "Height of a door", unit: "?", color: C.PRIMARY },
          { job: "Mass of an apple", unit: "?", color: C.SECONDARY },
          { job: "Water in a bath", unit: "?", color: C.ALERT },
          { job: "Length of an ant", unit: "?", color: C.PRIMARY },
        ], { bottom: 4.15 });
      }),
    (slide) => {
      addRevealAnswerBar(slide,
        "Door → m   |   Apple → g   |   Bath → L   |   Ant → mm",
        { color: C.SUCCESS, label: "Answer", fontSize: 18 });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // 16-17. CFU hinge + reveal
  withReveal(
    () => cfuSlide(pres, "CFU", "Sensible or not?", "Thumbs Up or Thumbs Down",
      "A student says they will measure how far they walk to school in millimetres.\n\nIs that a sensible choice?",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide,
        "NOT sensible. That would be millions of millimetres. Use kilometres (or metres).",
        { color: C.ALERT, label: "Check", fontSize: 18 });
    }
  );

  // 18. You Do — practice sheet
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: practice sheet", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.25, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "Section 1 — name a unit for each attribute.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "Section 2 — choose the best unit for each job.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "Section 3 — order the units largest to smallest.", options: { fontSize: 17, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.16, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.40;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("Always ask yourself", {
      x: 0.7, y: panelY + 0.14, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });
    addTextOnShape(s, "1.  What am I measuring — length, mass or capacity?", {
      x: 1.0, y: panelY + 0.54, w: 8.0, h: 0.42, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(s, "2.  Which unit gives a sensible number?", {
      x: 1.0, y: panelY + 1.06, w: 8.0, h: 0.42, rectRadius: 0.08,
      fill: { color: C.SECONDARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(s, "3.  Order from largest to smallest, like a staircase.", {
      x: 1.0, y: panelY + 1.58, w: 8.0, h: 0.42, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // 19. Exit Ticket
  exitTicketSlide(pres,
    [
      "Choose the best unit to measure the mass of a school bag. Explain your choice in a few words.",
      "Order these units from largest to smallest: centimetre, kilometre, millimetre, metre.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // 20. Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: how do you decide which unit to use?",
      scItems: [
        "I can name metric units for length, mass and capacity.",
        "I can choose a sensible unit to measure a given object.",
        "I can order metric units from largest to smallest and explain my choice.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "MMC_Session1_Choosing_Metric_Units.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  await generatePdfs();
  console.log("Session 1 build complete.");
}

// Defensive contrast pick for the ACCENT pill (some palettes have light accents).
function getContrastColorSafe(bg) {
  try { return T.getContrastColor(bg); } catch (e) { return C.WHITE; }
}

// ─── PDFs ─────────────────────────────────────────────────────────────────

async function generatePdfs() {
  // ── Worksheet ──────────────────────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Name units, choose the best unit for each job, and order units largest to smallest.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addTipBox(doc,
      "Unit ladder (largest to smallest):  LENGTH km, m, cm, mm.  MASS tonne, kg, g, mg.  CAPACITY ML, kL, L, mL.  Always ask: what am I measuring, and is my number sensible?",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 — Name a unit for each attribute", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   A length unit:   ______________   (example: cm)", y);
    y = addWriteLine(doc, "b)   A mass unit:      ______________", y);
    y = addWriteLine(doc, "c)   A capacity unit:  ______________", y);

    y = addSectionHeading(doc, "Section 2 — Choose the best unit for the job", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   Length of a soccer field   →   ______________   (worked: metres)", y);
    y = addWriteLine(doc, "b)   Mass of a feather              →   ______________", y);
    y = addWriteLine(doc, "c)   Petrol in a car's tank         →   ______________", y);
    y = addWriteLine(doc, "d)   Length of a fingernail        →   ______________", y);
    y = addWriteLine(doc, "e)   Distance Melbourne to Sydney →   ______________", y);

    y = addSectionHeading(doc, "Section 3 — Order the units from largest to smallest", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   mm, km, cm, m    →   ______, ______, ______, ______", y);
    y = addWriteLine(doc, "b)   g, tonne, mg, kg  →   ______, ______, ______, ______", y);
    y = addWriteLine(doc, "c)   L, mL, ML, kL      →   ______, ______, ______, ______", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.SECONDARY });
    y = addBodyText(doc, "Choosing a unit is a decision. For TWO of the jobs in Section 2, name a second unit that would also work, and write one sentence about when each unit would be the better choice.", y);
    y = addWriteLine(doc, "Job: ____________   Unit A: ______   Unit B: ______", y);
    y = addWriteLine(doc, "When each is better: ___________________________________________", y);

    addPdfFooter(doc, `Session ${SESSION} | Choosing Metric Units | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // ── Answer Key ─────────────────────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the choosing-units practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 — Name a unit (sample answers)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  any of mm, cm, m, km.   b)  any of mg, g, kg, tonne.   c)  any of mL, L, kL, ML.", y);

    y = addSectionHeading(doc, "Section 2 — Best unit for the job", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  metres (worked).   b)  milligrams or grams.   c)  litres.   d)  millimetres.   e)  kilometres.", y);
    y = addBodyText(doc, "Accept a reasonable alternative if the student gives a sensible number for the job.", y);

    y = addSectionHeading(doc, "Section 3 — Largest to smallest", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  km, m, cm, mm.   b)  tonne, kg, g, mg.   c)  ML, kL, L, mL.", y);

    y = addSectionHeading(doc, "Challenge — sample answer", y, { color: C.SECONDARY });
    y = addBodyText(doc, "Field: metres for the whole length, or centimetres for a precise line marking. Metres give a sensible number; centimetres give more accuracy for a small marking.", y);

    y = addTipBox(doc,
      "Watch for: students who pick the unit they use most (often centimetres) instead of the unit that gives a sensible number. Redirect with: what would the number look like?",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
