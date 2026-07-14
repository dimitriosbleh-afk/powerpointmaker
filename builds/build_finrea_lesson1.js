"use strict";

// Practical Problem Solving & Financial Reasoning (Year 6 Numeracy)
// Session 1 of 3: Estimate to check — rounding money & reasonableness.
// VC2M5N08 — estimation strategies; recognising the effect of rounding; choosing
//            when to round up, round down, or not round (cash vs digital).
// Daily Review: Advanced Number Concepts and Representation (prior learning).
// Fluency: Subtraction vertical algorithm.
//
// Unit routine (consistent across all 3 sessions): ESTIMATE -> MODEL -> SOLVE -> CHECK.
// Sessions are deliberately self-contained so a student who misses one can still
// access the next (each re-teaches its core move with worked examples).

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(3));
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
const TOTAL = 3;
const UNIT_TITLE = "Practical Problem Solving";
const FOOTER = `Financial Reasoning | Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FinRea_Session1_Estimate_To_Check";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Estimation and Reasonableness Practice",
  "Round money, estimate totals, and decide whether an answer is reasonable.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the estimation and reasonableness sheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome, mathematicians. Today our job is to be sensible with money, not just fast.
- We are going to estimate before we calculate, and then use that estimate to check whether an answer actually makes sense.
- By the end you will be able to look at a money answer and say straight away, that is about right, or that cannot be right.

DO:
- Have whiteboards and markers ready on desks.
- Keep the printed practice sheet face down until the You Do section.

TEACHER NOTES:
Session 1 of 3 in the Financial Reasoning unit. The recurring unit routine is Estimate, Model, Solve, Check. Today's focus is the Estimate and Check moves.

WATCH FOR:
- Students who think estimating is just guessing. We build the habit of rounding first, then adding.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Here is what we are using today.
- You will need a whiteboard for most of the lesson, and the practice sheet near the end.

DO:
- Print one practice sheet and one answer key per teacher.
- Have whiteboards and markers out.

TEACHER NOTES:
One student resource today (practice sheet) plus an answer key. Estimation is done mentally and on whiteboards, so very little printing is needed. The shopping-list visuals on the slides are the main models.

[General: Resources]`;

const NOTES_OVERVIEW = `SAY:
- This slide is for me, not the class. I can skip past it once I have read it.

DO:
- Read before teaching. Hide or skip during the lesson if projecting.

TEACHER NOTES:
Three-session arc. Session 1: estimate and check money answers. Session 2: model word problems with bar models and choose the operation. Session 3: build a simple fundraiser budget and find profit. Every session uses the same four-move routine - Estimate, Model, Solve, Check - so the language stays familiar week to week. Each session is deliberately self-contained: if a student misses one, they can still access the next because each session re-teaches its core move with fresh worked examples. Nothing in Session 3 assumes the exact problems from Sessions 1 or 2.

[General: Teacher Overview | VTLM 2.0: Planning]`;

const NOTES_DR_Q = `SAY:
- Quick warm up on numbers we already know.
- Read each one, then write your answer on your whiteboard.
- These are from earlier learning, not today's new work.

DO:
- Display the three prompts.
- Allow about 90 seconds.
- Scan for place value confusion.

TEACHER NOTES:
Daily Review revisits Advanced Number Concepts and Representation - place value, expanded form, and comparing large numbers. This is prior learning and retrieval, not today's estimation focus.

WATCH FOR:
- Students who read 40 250 correctly - secure with place value.
- Students who muddle the thousands - small group reminder this week.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check together.
- 40 250 in expanded form is 40 000 plus 200 plus 50.
- The larger number is 38 906, because 9 hundreds beats 8 hundreds.
- 6 740 rounded to the nearest thousand is 7 000.

DO:
- Click to reveal.
- Pause for tick and fix.

TEACHER NOTES:
Rounding 6740 to the nearest thousand bridges nicely into today - same idea of nearest, now with money.

WATCH FOR:
- Strong self-correction - secure.
- Students who rounded 6740 down to 6000 - they did not check the hundreds digit.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency now. Subtraction using the vertical algorithm.
- Set each one out in columns on your whiteboard and work top to bottom.
- Regroup when you need to. Whisper your answer, then write it.

DO:
- Display the three subtractions.
- Allow about 75 seconds.
- Watch for the regrouping step.

TEACHER NOTES:
Fluency builds the subtraction vertical algorithm, the same focus across all three sessions. Brisk retrieval, not new teaching.

WATCH FOR:
- Students who line up the columns and regroup cleanly - secure.
- Students who subtract the smaller digit from the larger out of habit - reteach regrouping.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check your columns.
- 6000 minus 2437 is 3563.
- 8052 minus 1675 is 6377.
- 7300 minus 4628 is 2672.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
The first one needs regrouping across zeros. Students who got 4437 subtracted the wrong way in the ones column.

WATCH FOR:
- Clean regrouping across the zeros in 6000 - secure.
- Answers larger than the top number - the columns were swapped.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LAUNCH = `SAY:
- Picture this. You walk into the shop with one twenty dollar note.
- You pick up milk, bread and eggs. The prices are on the screen.
- I am not asking for the exact total yet. I want a sensible guess. Roughly, will twenty dollars be enough?
- Turn and tell your partner what you think, and how you worked it out.

DO:
- Display the shopping prices.
- Give 30 seconds of partner talk.
- Take two or three responses. Listen for rounding language.

TEACHER NOTES:
The launch connects prior rounding knowledge to today's money context. We want students rounding each price to a friendly number and adding roughly. Do not calculate the exact answer here - that is the point. The estimate is the tool.

WATCH FOR:
- Students who round to the nearest dollar and add - exactly the move we teach today.
- Students who try to add the cents exactly - gently redirect to a rough estimate.

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Here is what we are learning today.
- We are learning to use rounding and estimation to check whether a money answer is reasonable.
- Now read the success criteria with me.

DO:
- Choral read the learning intention.
- Read each I can statement.

TEACHER NOTES:
SC1 is reachable for everyone today - rounding one amount to a friendly number. SC2 is the core target the exit ticket assesses - estimating a total. SC3 stretches into judging and explaining reasonableness.

WATCH FOR:
- Students repeating the language back - tracking with us.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Let us work through this one together. Watch how I estimate the total of a shopping list.
- I am not going to add the exact cents. Instead I round each price to the nearest dollar first.
- Three dollars eighty is closer to four than to three, so I round it to four dollars.
- Two dollars ten is closer to two, so it stays at two dollars.
- Five dollars forty rounds down to five dollars.
- Now I add my friendly numbers. Four plus two plus five is eleven.
- So my estimate is about eleven dollars. Quick, sensible, and good enough to plan with.

DO:
- Point to each price as you round it.
- Underline the cents to show what you are deciding about.
- Write the rounded numbers beside the originals.

TEACHER NOTES:
The receipt is the visual model. Rounding each item to the nearest dollar is the core skill. Keep the think-aloud about nearest, not chopping.

MISCONCEPTIONS:
- Misconception: estimating means adding the exact prices then rounding the answer.
  Why: students trust the exact calculation more than the estimate.
  Impact: it is slow and defeats the purpose of a quick check.
  Quick correction: round each number first, then add the friendly numbers.

WATCH FOR:
- Students who round each price before adding - secure.
- Students who reach for an exact total - redirect to round first.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now I want to show you that rounding is a decision, not always the same move.
- If I am buying one item with cash, I round up so I am sure to bring enough. Seven dollars twenty, I take eight dollars to be safe.
- If I am estimating a whole shop, I round each item to the nearest dollar to get a quick total.
- But if I am paying by card, the machine takes the exact amount, so there is no need to round at all. I just pay seven dollars twenty.
- Same numbers, different sensible choice, depending on the situation.

DO:
- Point to each situation on the panel.
- Stress the words up, nearest, and exact.

TEACHER NOTES:
This is the VC2M5N08 rounding-context idea. The choice of rounding depends on the situation - cash versus digital, one item versus a whole shop. Students often think rounding is one fixed rule; here it is a judgement.

MISCONCEPTIONS:
- Misconception: you always round to the nearest dollar.
  Why: that is the rule they practised most.
  Impact: they bring too little cash, or round a digital payment that does not need it.
  Quick correction: ask what the situation needs - enough cash, a quick total, or an exact amount.

WATCH FOR:
- Students who can say why each choice fits - secure and ready for depth.
- Students who round everything the same way - revisit the situation question.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1_Q = `SAY:
- Quick check on your whiteboards.
- Round each price to the nearest dollar, then estimate the total.
- Three eighty, four ninety-five, two ten. Show me your estimate on three, two, one.

DO:
- Display the prompt.
- Allow about 40 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your estimate on three, two, one, show."
- Scan for: about 11 dollars (4 plus 5 plus 2).
PROCEED: If 80 percent show about 11, click to reveal and move on.
PIVOT: Most likely misconception - students rounded 4.95 down to 4 instead of up to 5.
- Reteach: "4.95 is almost 5 dollars. The next place is 9, so it rounds up to 5."
- Re-check: "What does 4.95 round to? Then add again."
- Use the worked receipt model again if needed.

TEACHER NOTES:
Checks the round-then-add habit before guided practice.

WATCH FOR:
- Estimates of about 11 - secure.
- Estimates near 8 or 9 - they rounded 4.95 the wrong way.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me now. Same routine, a lunch order this time.
- Step one, round each price to the nearest dollar.
- Step two, add your friendly numbers.
- Step three, write your estimate. Work it on your whiteboard with your partner.

DO:
- Keep the steps visible.
- Walk and scan. Prompt with "round first" where needed.
- Take 60 seconds before revealing.

TEACHER NOTES:
Same structure as the I Do, new numbers. Students do the rounding and adding; the steps stay on screen as a self-cue.

WATCH FOR:
- Pairs who round each item first - secure.
- Pairs who try to add exact cents - redirect to rounding.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check our estimate.
- Twelve thirty rounds to twelve. Seven eighty rounds to eight. Four ninety-five rounds to five.
- Twelve plus eight plus five is twenty-five.
- So the estimate is about twenty-five dollars.

DO:
- Click to reveal.
- Run the rounding once more aloud.

TEACHER NOTES:
The reveal restates each rounding and the total. Look for students who wrote about 23 - they probably rounded 7.80 down.

WATCH FOR:
- Self-correction on the reveal - secure.
- Estimates near 23 - 7.80 was rounded down by mistake.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. A student added two prices and wrote this answer.
- Eight ninety-five plus eleven twenty equals two hundred and one dollars fifty.
- Do not calculate it exactly. Use estimation to decide. Is that answer reasonable?
- Thumbs up if reasonable, thumbs down if not.

DO:
- Display the calculation.
- Give 15 seconds of thinking time.

CFU CHECKPOINT:
Technique: Thumbs Up or Down
Script:
- Say: "Thumbs up if reasonable, thumbs down if not."
- Scan for: thumbs down. About 9 plus 11 is 20, not 200.
PROCEED: If 80 percent thumbs down, click to reveal.
PIVOT: Most likely misconception - students try to work out the exact answer and stall.
- Reteach: "Estimation is the quick check. Round first. 9 plus 11 is about 20. Is 201 close to 20?"
- Re-check: "Use your estimate. Reasonable or not?"

TEACHER NOTES:
The point is using estimation as a reasonableness check. The wrong answer here comes from a slipped decimal point - a very common real error. Students who try to calculate exactly are not yet using estimation as a tool.

WATCH FOR:
- Confident thumbs down with a reason - they used estimation.
- Students recalculating exactly - reteach the estimate-first check.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Time to work on your own. Take the practice sheet.
- Section one, round each amount to the nearest dollar. Section two, estimate the totals. Section three, decide if each answer is reasonable and explain why.
- Estimate first every time. Show your rounded numbers.

DO:
- Distribute the practice sheet.
- Circulate. Sit with the enabling group first.
- Watch for students using estimation as a check, not a guess.

TEACHER NOTES:
The sheet sequences the three skills: round, estimate a total, then judge reasonableness. Section three is the threshold - using the estimate to decide.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the modelled first item and the rounding rule box at the top of the sheet. Round each amount, then add the friendly numbers.
- Extra Notes: Do the first total in Section 2 together with this group before they continue.
EXTENDING PROMPT:
- Task: The challenge box - decide when you should round up, round to nearest, or not round at all, and write a sentence explaining your choice for a cash payment versus a card payment.
- Extra Notes: Push for a reason that names the situation.

WATCH FOR:
- Students who round before adding - secure.
- Students who add exactly then round the answer - redirect to round first.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task, on your own, on your whiteboard.
- One, estimate eight dollars ninety-five plus eleven dollars ten by rounding to the nearest dollar.
- Two, a student wrote that answer as ninety-one dollars five. Use estimation to decide if that is reasonable, and explain why.

DO:
- Display the prompt.
- Allow about 3 minutes.
- Collect whiteboards or take a photo for your records.

TEACHER NOTES:
Exit ticket assesses SC2 (estimating a total) and reaches into SC3 (judging reasonableness with an explanation). Internal target is SC2. Do not display any SC label to students.

WATCH FOR:
- Estimate of about 20 with a clear not-reasonable judgement - secure.
- Students who cannot judge without the exact answer - revisit estimate-first in Session 2.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look back at our success criteria.
- Show me thumbs for each one - got it, getting there, or need more practice.
- Turn and tell your partner: why is it useful to estimate before you calculate?

DO:
- Read each I can statement.
- Use thumbs to self-assess.
- Note who is still unsure for tomorrow.

TEACHER NOTES:
The big idea is estimation as a reasonableness check. Students who can say "to know roughly what to expect, so I can spot a silly answer" have the relational understanding we want.

WATCH FOR:
- Students who explain the why - secure.
- Students who only restate the steps - revisit the check idea in Session 2.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Visual helpers (build-script local) ─────────────────────────────────────

// A small itemised "receipt" card: each row shows item and a price string
// (often "$3.80 -> $4"), with an optional coloured estimate-total bar.
function priceCard(slide, lg, opts) {
  const o = opts;
  const x = lg.rightX;
  const w = lg.rightW;
  const y = lg.panelTopPadded;
  const rows = o.rows;
  const color = o.color || C.PRIMARY;
  const rowH = 0.40;
  const headH = 0.46;
  const totalH = o.total ? 0.58 : 0;
  const cardH = Math.min(lg.safeBottom - y,
    headH + rows.length * rowH + totalH + 0.30);
  addCard(slide, x, y, w, cardH, { strip: color });
  slide.addText(o.title, {
    x: x + 0.2, y: y + 0.10, w: w - 0.4, h: 0.30,
    fontSize: 15, fontFace: FONT_H, color, bold: true,
    align: "center", margin: 0,
  });
  let ry = y + headH;
  const colSplit = (w - 0.5) * 0.42;
  rows.forEach((r) => {
    slide.addText(r.item, {
      x: x + 0.25, y: ry, w: colSplit, h: rowH,
      fontSize: 13.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "middle", margin: 0,
    });
    slide.addText(r.right, {
      x: x + 0.25 + colSplit, y: ry, w: (w - 0.5) - colSplit, h: rowH,
      fontSize: 13.5, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "right", valign: "middle", margin: 0,
    });
    ry += rowH;
  });
  if (o.total) {
    addTextOnShape(slide, o.total, {
      x: x + 0.25, y: ry + 0.06, w: w - 0.5, h: 0.48, rectRadius: 0.08,
      fill: { color: C.SUCCESS },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
  }
}

// A stacked "decision panel": each row is a situation + the sensible action.
function decisionPanel(slide, lg, rows) {
  const x = lg.rightX;
  const w = lg.rightW;
  const y = lg.panelTopPadded;
  const gap = 0.14;
  const cardH = (lg.safeBottom - y - gap * (rows.length - 1)) / rows.length;
  rows.forEach((r, i) => {
    const cy = y + i * (cardH + gap);
    addCard(slide, x, cy, w, cardH, { strip: r.color });
    slide.addText(r.situation, {
      x: x + 0.22, y: cy + 0.10, w: w - 0.44, h: 0.34,
      fontSize: 13.5, fontFace: FONT_H, color: r.color, bold: true,
      valign: "middle", margin: 0,
    });
    slide.addText(r.action, {
      x: x + 0.22, y: cy + 0.44, w: w - 0.44, h: cardH - 0.52,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0,
    });
  });
}

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  titleSlide(pres, UNIT_TITLE, "Session 1: Estimate to check — rounding money",
    `Year 6 Numeracy | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3. Unit overview (teacher-facing)
  contentSlide(pres, "For the teacher", C.MUTED,
    "Unit overview — 3 sessions",
    [
      "Routine every session: Estimate → Model → Solve → Check.",
      "Session 1: estimate and check money answers (today).",
      "Session 2: model word problems with bar models, choose the operation.",
      "Session 3: build a fundraiser budget and find profit.",
      "Each session is self-contained — a student who misses one can still access the next.",
    ],
    NOTES_OVERVIEW, FOOTER);

  // 4-5. Daily Review + reveal
  withReveal(
    () => dailyReviewSlide(pres,
      "Daily Review: Numbers we know",
      [
        "Write 40 250 in expanded form.",
        "Which is larger: 38 906 or 38 890?",
        "Round 6 740 to the nearest thousand.",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide,
        ["40 000 + 200 + 50", "38 906", "7 000"],
        { color: C.SUCCESS });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // 6-7. Fluency + reveal (subtraction vertical algorithm)
  withReveal(
    () => fluencySlide(pres, "Fluency: Subtraction",
      ["6 000 − 2 437", "8 052 − 1 675", "7 300 − 4 628"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, ["3 563", "6 377", "2 672"], { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // 8. Launch
  contentSlide(pres, "Launch", C.SECONDARY,
    "Will $20 be enough?",
    [
      "You have one $20 note.",
      "You grab milk, bread and eggs.",
      "Don't add it exactly.",
      "Roughly — is $20 enough?",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      priceCard(slide, lg, {
        title: "On the shelf",
        color: C.SECONDARY,
        rows: [
          { item: "Milk", right: "$4.20" },
          { item: "Bread", right: "$3.90" },
          { item: "Eggs", right: "$5.80" },
        ],
        total: "Estimate first, then decide",
      });
    });

  // 9. LI / SC
  liSlide(pres,
    "We are learning to use rounding and estimation to check whether a money answer is reasonable.",
    [
      "I can round an amount of money to the nearest dollar.",
      "I can estimate the total of a money problem before I calculate.",
      "I can decide whether an answer is reasonable and explain why.",
    ],
    NOTES_LI_SC, FOOTER);

  // 10. I Do 1 — estimate a shopping total
  workedExSlide(pres, 2, "I Do", "Estimate the total — round to the nearest $",
    [
      "Round each price to the nearest dollar.",
      "$3.80 → $4.",
      "$2.10 → $2.",
      "$5.40 → $5.",
      "",
      "Add the friendly numbers.",
      "$4 + $2 + $5 = $11.",
      "",
      "Estimate: about $11.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      priceCard(slide, lg, {
        title: "Shopping list",
        color: C.PRIMARY,
        rows: [
          { item: "Bread", right: "$3.80 → $4" },
          { item: "Drink", right: "$2.10 → $2" },
          { item: "Cheese", right: "$5.40 → $5" },
        ],
        total: "Estimate ≈ $11",
      });
    });

  // 11. I Do 2 — rounding is a decision
  workedExSlide(pres, 2, "I Do", "Round up, nearest, or not at all?",
    [
      "Rounding is a choice, not one fixed rule.",
      "",
      "Ask: what does the situation need?",
      "",
      "Bring enough cash → round UP.",
      "Quick total for a shop → nearest $.",
      "Paying by card → exact, don't round.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      decisionPanel(slide, lg, [
        { situation: "One item, paying cash", action: "Round UP. $7.20 → take $8 so you bring enough.", color: C.PRIMARY },
        { situation: "Estimating a whole shop", action: "Round each item to the nearest dollar for a quick total.", color: C.SECONDARY },
        { situation: "Paying by card (digital)", action: "Don't round. The exact amount comes out: $7.20.", color: C.ALERT },
      ]);
    });

  // 12-13. CFU 1 + reveal
  withReveal(
    () => cfuSlide(pres, "CFU", "Estimate the total", "Show Me Boards",
      "Round each price to the nearest dollar, then estimate the total:\n\n$3.80  +  $4.95  +  $2.10",
      NOTES_CFU1_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "About $11   ($4 + $5 + $2)",
        { color: C.SUCCESS, label: "Answer" });
    }
  );

  // 14-15. We Do + reveal
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Estimate the lunch order",
      [
        "With your partner.",
        "",
        "Step 1: round each price to the nearest $.",
        "Step 2: add the friendly numbers.",
        "Step 3: write your estimate.",
        "",
        "Estimate ≈ __________",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        priceCard(slide, lg, {
          title: "Lunch order",
          color: C.SECONDARY,
          rows: [
            { item: "Wrap", right: "$12.30" },
            { item: "Juice", right: "$7.80" },
            { item: "Muffin", right: "$4.95" },
          ],
          total: "Round each first",
        });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "About $25   ($12 + $8 + $5)",
        { color: C.SUCCESS, label: "Answer" });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // 16-17. CFU hinge + reveal
  withReveal(
    () => cfuSlide(pres, "CFU", "Reasonable or not?", "Thumbs Up or Thumbs Down",
      "A student wrote:\n\n$8.95  +  $11.20  =  $201.50\n\nUse estimation. Round first.\nIs this answer reasonable?",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide,
        "NOT reasonable. $9 + $11 ≈ $20, not $200. The decimal point slipped.",
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
      { text: "Section 1 — round each amount to the nearest dollar.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "Section 2 — estimate each total.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "Section 3 — is the answer reasonable? Explain.", options: { fontSize: 17, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.16, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.40;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("Remember the routine", {
      x: 0.7, y: panelY + 0.14, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });
    addTextOnShape(s, "1.  Round each amount to the nearest dollar FIRST.", {
      x: 1.0, y: panelY + 0.54, w: 8.0, h: 0.42, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(s, "2.  Add the friendly numbers to estimate the total.", {
      x: 1.0, y: panelY + 1.06, w: 8.0, h: 0.42, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    }, { fontSize: 16, fontFace: FONT_H, color: getContrastColorSafe(C.ACCENT), bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // 19. Exit Ticket
  exitTicketSlide(pres,
    [
      "Estimate $8.95 + $11.10 by rounding each amount to the nearest dollar.",
      "A student wrote $8.95 + $11.10 = $91.05. Use estimation to decide if this is reasonable. Explain why.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // 20. Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: why is it useful to estimate before you calculate?",
      scItems: [
        "I can round an amount of money to the nearest dollar.",
        "I can estimate the total of a money problem before I calculate.",
        "I can decide whether an answer is reasonable and explain why.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FinRea_Session1_Estimate_To_Check.pptx");
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
      subtitle: "Round money to the nearest dollar, estimate totals, and check if answers are reasonable.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addTipBox(doc,
      "Round to the nearest dollar: 50 cents or more rounds UP, less than 50 cents rounds down. Estimate FIRST, then add your friendly numbers.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section 1 — Round each amount to the nearest dollar", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   $3.80   →   $4      (modelled for you)", y);
    y = addWriteLine(doc, "b)   $6.20   →   _______", y);
    y = addWriteLine(doc, "c)   $2.55   →   _______", y);
    y = addWriteLine(doc, "d)   $9.49   →   _______", y);
    y = addWriteLine(doc, "e)   $14.90  →   _______", y);

    y = addSectionHeading(doc, "Section 2 — Estimate each total (round first, then add)", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   $4.80 + $3.10 + $2.95          Estimate:  _______", y);
    y = addWriteLine(doc, "b)   $11.20 + $6.75                      Estimate:  _______", y);
    y = addWriteLine(doc, "c)   $7.40 + $8.60 + $5.20          Estimate:  _______", y);

    y = addSectionHeading(doc, "Section 3 — Is the answer reasonable? Estimate, decide, explain", y, { color: C.PRIMARY });
    y = addWriteLine(doc, "a)   $5.90 + $4.20 = $101.    Estimate: ___   Reasonable? YES / NO", y);
    y = addWriteLine(doc, "      Why? ______________________________________________________", y);
    y = addWriteLine(doc, "b)   $12.50 + $6.30 = $18.80.   Estimate: ___   Reasonable? YES / NO", y);
    y = addWriteLine(doc, "      Why? ______________________________________________________", y);
    y = addWriteLine(doc, "c)   $9.95 + $10.10 = $20.05.   Estimate: ___   Reasonable? YES / NO", y);
    y = addWriteLine(doc, "      Why? ______________________________________________________", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.SECONDARY });
    y = addBodyText(doc, "Rounding is a choice. Write one sentence about how you would round when paying with cash, and one sentence about paying by card.", y);
    y = addWriteLine(doc, "Cash:  _________________________________________________________", y);
    y = addWriteLine(doc, "Card:  _________________________________________________________", y);

    addPdfFooter(doc, `Session ${SESSION} | Estimation & Reasonableness | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // ── Answer Key ─────────────────────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the estimation and reasonableness practice sheet.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Section 1 — Nearest dollar", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  $4   (modelled)     b)  $6     c)  $3     d)  $9     e)  $15", y);

    y = addSectionHeading(doc, "Section 2 — Estimated totals", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  about $11   ($5 + $3 + $3).", y);
    y = addBodyText(doc, "b)  about $18   ($11 + $7).", y);
    y = addBodyText(doc, "c)  about $21   ($7 + $9 + $5).", y);

    y = addSectionHeading(doc, "Section 3 — Reasonableness", y, { color: C.PRIMARY });
    y = addBodyText(doc, "a)  Estimate about $10 ($6 + $4). NOT reasonable - $101 means the decimal point slipped.", y);
    y = addBodyText(doc, "b)  Estimate about $19 ($13 + $6). Reasonable - $18.80 is close to the estimate.", y);
    y = addBodyText(doc, "c)  Estimate about $20 ($10 + $10). Reasonable - $20.05 is close to the estimate.", y);

    y = addSectionHeading(doc, "Challenge — sample answers", y, { color: C.SECONDARY });
    y = addBodyText(doc, "Cash: I round the price UP so I bring enough money. For example, $7.20 means I take $8.", y);
    y = addBodyText(doc, "Card: I do not round. The exact amount is taken from the card, so I pay $7.20.", y);

    y = addTipBox(doc,
      "Watch for: students who add the exact amounts and then round the answer. Redirect them to round each number FIRST. In Section 3, a student who cannot judge without calculating exactly is not yet using estimation as a check.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
