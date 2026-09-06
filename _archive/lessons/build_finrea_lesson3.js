"use strict";

// Practical Problem Solving & Financial Reasoning (Year 6 Numeracy)
// Session 3 of 3: Make a plan — a fundraiser budget.
// VC2M5N09 — model a financial situation by creating a financial plan;
//            combine operations to find total cost, income and profit;
//            decide if the plan works and justify a change.
// Daily Review: Interpreting Graphs and Critiquing Data (prior learning).
// Fluency: Subtraction vertical algorithm.
//
// Unit routine (consistent across all 3 sessions): ESTIMATE -> MODEL -> SOLVE -> CHECK.
// Self-contained: budgeting is taught from scratch with worked examples, so a
// student who missed Sessions 1 or 2 can still access today's lesson.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine, addLinedArea,
  addProblem, addTwoColumnOrganiser,
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

const SESSION = 3;
const TOTAL = 3;
const UNIT_TITLE = "Practical Problem Solving";
const FOOTER = `Financial Reasoning | Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FinRea_Session3_Fundraiser_Budget";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const PLANNER_RES = makeSessionResource(SESSION,
  "Fundraiser Budget Planner",
  "Work out total cost, income and profit, then suggest one improvement.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the fundraiser budget planner.");
const EXT_RES = makeSessionResource(SESSION,
  "Year 8 Extension Task",
  "Stretch task: GST, discounts, break-even and profit margin as a percentage.");
const EXT_KEY_RES = makeSessionResource(SESSION,
  "Year 8 Extension Answer Key",
  "Worked answers for the Year 8 extension task.");
const RESOURCE_ITEMS = [PLANNER_RES, ANSWER_KEY_RES, EXT_RES, EXT_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome back, mathematicians. This is the session where it all comes together.
- A budget is a plan for money - what we spend, what we earn, and what is left over.
- By the end you will be able to work out the total cost, the income, and the profit of a fundraiser, and decide whether the plan is worth running.

DO:
- Whiteboards and markers ready.
- Keep the budget planner and extension task face down until the You Do section.

TEACHER NOTES:
Session 3 of 3, the final session. The unit routine is Estimate, Model, Solve, Check. Today combines all of them inside a budget. Self-contained: budgeting is taught from the start.

WATCH FOR:
- Students who confuse income with profit. Profit is what is LEFT after costs.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Here is what we are using today.
- The budget planner is for the You Do. There is also a Year 8 extension task for anyone ready for a bigger challenge.

DO:
- Print the budget planner and answer key.
- Print the Year 8 extension task and its answer key for your extending students.
- Whiteboards and markers out.

TEACHER NOTES:
Unit note: each session is self-contained, so missed sessions are not a barrier. Today: a budget planner, an answer key, and a Year 8 extension task with its own answer key. The budget tables on the slides are the main visual model.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Quick warm up. Look at the graph of cakes sold each day.
- Read each question and write your answer on your whiteboard.
- This is from our earlier data work.

DO:
- Display the bar graph and the three prompts.
- Allow about 90 seconds.

TEACHER NOTES:
Daily Review revisits Interpreting Graphs and Critiquing Data - prior learning. The third prompt is a data-critique question about why a zero baseline matters for a fair comparison.

WATCH FOR:
- Students who read the bars accurately - secure.
- Students who guess the difference - prompt them to read both bar heights first.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- Friday sold the most cakes - 35.
- Friday minus Monday is 35 minus 10, which is 25 more.
- The graph starts at zero, and that matters because it lets us compare the bar heights fairly. If it did not start at zero, the differences would look bigger than they really are.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
The zero-baseline idea is the data-critique point. Keep it short but name why it matters.

WATCH FOR:
- Students who explain the fair-comparison idea - strong.
- Students unsure - revisit with a squashed-axis example another day.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Subtraction with the vertical algorithm one more time.
- Columns lined up, top to bottom, regroup when you need to.
- Whisper, then write.

DO:
- Display the three subtractions.
- Allow about 75 seconds.

TEACHER NOTES:
Same subtraction focus across the unit. Brisk retrieval. Subtraction also matters today because profit is income minus costs.

WATCH FOR:
- Clean regrouping - secure.
- Smaller-from-larger habit - reteach regrouping.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check your columns.
- 4000 minus 2165 is 1835.
- 7206 minus 3489 is 3717.
- 8500 minus 5743 is 2757.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
The regrouping across zeros in the first and third ones is the tricky part. Link to today: we will subtract costs from income to find profit.

WATCH FOR:
- Correct regrouping across zeros - secure.
- Off answers - the regroup was not carried through.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LAUNCH = `SAY:
- Imagine our class wants to run a cake stall to raise money.
- Before we earn a cent, we have to spend some. So two questions.
- What will we spend money ON? And what will we earn money FROM?
- Turn and tell your partner one thing for each side - money out and money in.

DO:
- Display the money out and money in columns.
- Give 30 seconds of partner talk.
- Collect a few ideas onto each side.

TEACHER NOTES:
The launch surfaces the two sides of a budget - costs and income - before we formalise them. This sets up profit as the difference. Keep it quick and concrete.

WATCH FOR:
- Students who name both costs and income - ready for the budget.
- Students who only think about earning - prompt "but what do we buy first?"

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Here is today's learning.
- We are learning to create a simple budget that combines operations to plan a fundraising event.
- Read the success criteria with me.

DO:
- Choral read the learning intention.
- Read each I can statement.

TEACHER NOTES:
SC1 is reachable for everyone - reading a budget to find costs and income. SC2 is the core target the exit ticket assesses - combining operations to find total cost and profit. SC3 reaches into deciding if a plan works and suggesting a change.

WATCH FOR:
- Students repeating the language - tracking with us.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Let us build the first half of a budget together - the costs.
- We need 3 packets of flour at 4 dollars each. That is equal groups, so I multiply. 3 times 4 is 12 dollars.
- We also need icing, which is 6 dollars.
- Now I add the costs together. 12 plus 6 is 18 dollars.
- So the total cost of running the stall is 18 dollars. Notice I used multiplication AND addition in one budget.

DO:
- Point to each line of the cost table.
- Say "multiply for the equal groups, then add the costs".
- Underline the total cost.

TEACHER NOTES:
A budget combines operations. Multiply for equal groups, then add to find the total cost. Keep the table visible as the model.

MISCONCEPTIONS:
- Misconception: every line in a budget is just added.
  Why: students think a budget is one big addition.
  Impact: they miss the 3 times 4 and write 3 plus 4.
  Quick correction: ask "is this an equal group? Then multiply that line first."

WATCH FOR:
- Students who multiply the flour line then add - secure.
- Students who add every number - point to the equal groups.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now the second half - income and profit.
- We plan to sell 40 cakes at 2 dollars each. Equal groups again, so I multiply. 40 times 2 is 80 dollars of income.
- Here is the key idea. Profit is what is LEFT after we pay our costs.
- So profit is income minus costs. 80 minus 18 is 62 dollars.
- Our plan makes a profit of 62 dollars. Quick check: about 80 take away about 20 is about 60, so 62 is reasonable.

DO:
- Point to the income line, then the profit line.
- Stress "profit is income MINUS costs".
- Run the estimate check aloud.

TEACHER NOTES:
Profit equals income minus costs is the threshold idea today. The estimate-to-check from Session 1 returns here.

MISCONCEPTIONS:
- Misconception: profit is the same as income.
  Why: the income number is big and looks like the result.
  Impact: students report 80 dollars profit and forget the costs.
  Quick correction: "Did we pay for anything? Then take the costs off first."

WATCH FOR:
- Students who subtract costs from income - secure.
- Students who stop at income - remind them about costs.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1_Q = `SAY:
- Quick check on your whiteboards.
- A stall earns 95 dollars of income. Its costs are 32 dollars.
- What is the profit? Write the number sentence and solve it.

DO:
- Display the prompt.
- Allow about 40 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your number sentence and answer on three, two, one, show."
- Scan for: 95 minus 32 equals 63.
PROCEED: If 80 percent have 95 minus 32 equals 63, click to reveal and move on.
PIVOT: Most likely misconception - students add income and costs, or report income as profit.
- Reteach: "Profit is what is LEFT. We take the costs off the income. 95 take away 32."
- Re-check: "Income take away costs. What is the profit?"

TEACHER NOTES:
Checks the profit equals income minus costs idea before guided practice.

WATCH FOR:
- 95 minus 32 equals 63 - secure.
- 95 plus 32, or 95 alone - reteach profit as what is left.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. Let us budget a lemonade stall.
- Costs: lemons and sugar are 8 dollars, and cups are 4 dollars.
- Income: we sell 50 cups at 1 dollar each.
- Step 1: total cost. Step 2: income. Step 3: profit. Work it on your whiteboard with your partner.

DO:
- Keep the budget table visible.
- Walk and scan.
- Take about 75 seconds before revealing.

TEACHER NOTES:
A full budget with three calculations: add the costs, multiply for the income, subtract for the profit. Different context from the I Do on purpose.

WATCH FOR:
- Pairs who find all three totals - secure.
- Pairs who report income as profit - point to the costs.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check the lemonade budget.
- Total cost: 8 plus 4 is 12 dollars.
- Income: 50 times 1 is 50 dollars.
- Profit: 50 minus 12 is 38 dollars.
- Quick check: about 50 take away about 12 is about 38. Reasonable.

DO:
- Click to reveal each total.
- Run the three steps aloud.

TEACHER NOTES:
The reveal restates total cost, income, and profit, then the estimate check. Look for students who wrote 50 as the profit.

WATCH FOR:
- All three totals correct - secure.
- Profit reported as 50 - they forgot to subtract the costs.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. Look at this plan.
- Income is 40 dollars. Costs are 55 dollars.
- Is this plan worth running? Thumbs up for yes, thumbs down for no.
- Be ready to say why.

DO:
- Display the plan.
- Give 15 seconds of thinking time.

CFU CHECKPOINT:
Technique: Thumbs Up or Down
Script:
- Say: "Thumbs up if it is worth running, thumbs down if it is not."
- Scan for: thumbs down. Costs are higher than income, so it makes a loss.
PROCEED: If 80 percent thumbs down with a reason, click to reveal.
PIVOT: Most likely misconception - students see 40 dollars income and think any income is good.
- Reteach: "Profit is income minus costs. 40 minus 55 is negative 15. The costs are higher, so it loses money."
- Re-check: "Does the plan make money or lose money?"

TEACHER NOTES:
This is a decision and reasonableness check. The plan makes a loss because costs exceed income. Strong students will suggest a change - sell more, charge more, or spend less.

WATCH FOR:
- Thumbs down with "it makes a loss" - secure and ready to suggest a change.
- Thumbs up - reteach profit as income minus costs.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Time to work on your own. Take the budget planner.
- Work out the total cost, then the income, then the profit. Last, suggest one change that would make a bigger profit.
- Remember: profit is income minus costs.

DO:
- Distribute the budget planner. Hand the Year 8 extension task to students ready for a bigger challenge.
- Circulate. Sit with the enabling group first.
- Watch for students subtracting costs from income for the profit.

TEACHER NOTES:
The planner gives a partly completed budget so students focus on combining the operations and on the profit. The final question asks for a justified improvement - that is SC3.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the budget planner. Some numbers are filled in. Find the totals and the profit using the labels as a guide.
- Extra Notes: Do the cost total together with this group, then let them try income and profit.
EXTENDING PROMPT:
- Task: The Year 8 Extension Task - GST, a bulk discount, break-even, and profit as a percentage.
- Extra Notes: This introduces percentages, so hand it only to students ready to stretch. It has its own answer key.

WATCH FOR:
- Students who find profit by subtracting costs - secure.
- Students who report income as profit - point back to the costs row.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task, on your own, on your whiteboard.
- A class fun run budget: income is 120 dollars, and costs are 45 dollars.
- One, work out the profit. Two, is this plan worth running? Explain why.

DO:
- Display the prompt.
- Allow about 3 minutes.
- Collect whiteboards or photograph for your records.

TEACHER NOTES:
Exit ticket assesses SC2 (combine operations to find profit) and SC3 (decide if the plan works and justify). Profit is 120 minus 45, which is 75 dollars, so yes it is worth running. Internal target is SC2. Do not display any SC label to students.

WATCH FOR:
- Profit of 75 with a clear yes-it-makes-a-profit reason - secure.
- Students who add, or report 120 - revisit profit as what is left.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look back at our success criteria, and at the whole unit.
- Show me thumbs for each one.
- Turn and tell your partner: across these three sessions, what does a good problem solver do BEFORE they trust an answer?

DO:
- Read each I can statement.
- Use thumbs to self-assess.
- Celebrate the unit - estimate, model, solve, check.

TEACHER NOTES:
The closing reflects on the whole unit routine - Estimate, Model, Solve, Check. Students who name "estimate or check that it is reasonable" have carried the big idea right through.

WATCH FOR:
- Students who mention checking or estimating - secure across the unit.
- Students unsure - a quick small-group recap of the routine next week.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Visual helpers ──────────────────────────────────────────────────────────

// A simple column (bar) graph drawn inside a panel: axes, scaled bars, value
// labels above and category labels below. Used for the data Daily Review.
function miniBarGraph(slide, x, y, w, h, data, color) {
  const maxV = Math.max(...data.map((d) => d.value));
  const baseY = y + h - 0.30;
  const axisX = x + 0.10;
  slide.addShape("line", { x: axisX, y: y + 0.05, w: 0, h: baseY - (y + 0.05), line: { color: C.CHARCOAL, width: 1.5 } });
  slide.addShape("line", { x: axisX, y: baseY, w: w - 0.2, h: 0, line: { color: C.CHARCOAL, width: 1.5 } });
  const n = data.length;
  const slot = (w - 0.3) / n;
  const barW = slot * 0.5;
  data.forEach((d, i) => {
    const bh = (d.value / maxV) * (baseY - (y + 0.28));
    const bx = axisX + 0.12 + i * slot + (slot - barW) / 2;
    const by = baseY - bh;
    slide.addShape("rect", { x: bx, y: by, w: barW, h: bh, fill: { color: d.fill || color }, line: { color, width: 0.5 } });
    slide.addText(String(d.value), {
      x: bx - 0.15, y: by - 0.26, w: barW + 0.3, h: 0.22,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
    });
    slide.addText(String(d.label), {
      x: bx - 0.2, y: baseY + 0.04, w: barW + 0.4, h: 0.22,
      fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
    });
  });
}

// A budget table: label/amount rows, with highlighted total and profit rows.
function budgetTable(slide, lg, opts) {
  const o = opts;
  const x = lg.rightX;
  const w = lg.rightW;
  const y = lg.panelTopPadded;
  const color = o.color || C.PRIMARY;
  const rows = o.rows;
  const headH = 0.46;
  const rowH = 0.50;
  const cardH = Math.min(lg.safeBottom - y, headH + rows.length * rowH + 0.26);
  addCard(slide, x, y, w, cardH, { strip: color });
  slide.addText(o.title, {
    x: x + 0.2, y: y + 0.10, w: w - 0.4, h: 0.30,
    fontSize: 15, fontFace: FONT_H, color, bold: true, align: "center", margin: 0,
  });
  let ry = y + headH;
  const labelW = (w - 0.5) * 0.62;
  rows.forEach((r) => {
    if (r.kind === "total" || r.kind === "profit") {
      const fill = r.kind === "profit" ? C.SUCCESS : color;
      addTextOnShape(slide, r.label + "    " + r.amount, {
        x: x + 0.22, y: ry + 0.04, w: w - 0.44, h: rowH - 0.10, rectRadius: 0.06,
        fill: { color: fill },
      }, { fontSize: 13.5, fontFace: FONT_H, color: C.WHITE, bold: true });
    } else {
      slide.addText(r.label, {
        x: x + 0.25, y: ry, w: labelW, h: rowH,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
      });
      slide.addText(r.amount, {
        x: x + 0.25 + labelW, y: ry, w: (w - 0.5) - labelW, h: rowH,
        fontSize: 13.5, fontFace: FONT_H, color: r.amountColor || C.CHARCOAL, bold: true,
        align: "right", valign: "middle", margin: 0,
      });
    }
    ry += rowH;
  });
}

// Two-column "money out / money in" card for the launch.
function moneyCard(slide, lg) {
  const x = lg.rightX;
  const w = lg.rightW;
  const y = lg.panelTopPadded;
  const cardH = Math.min(lg.safeBottom - y, 2.9);
  const colW = (w - 0.3) / 2;
  // Money out
  addCard(slide, x, y, colW, cardH, { strip: C.ALERT });
  slide.addText("Money OUT", {
    x: x + 0.1, y: y + 0.12, w: colW - 0.2, h: 0.3,
    fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", margin: 0,
  });
  slide.addText([
    { text: "Ingredients", options: { breakLine: true } },
    { text: "Cups & bags", options: { breakLine: true } },
    { text: "Posters", options: {} },
  ], {
    x: x + 0.15, y: y + 0.55, w: colW - 0.3, h: cardH - 0.7,
    fontSize: 13.5, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
    paraSpaceAfter: 6,
  });
  // Money in
  const x2 = x + colW + 0.3;
  addCard(slide, x2, y, colW, cardH, { strip: C.SUCCESS });
  slide.addText("Money IN", {
    x: x2 + 0.1, y: y + 0.12, w: colW - 0.2, h: 0.3,
    fontSize: 14, fontFace: FONT_H, color: C.SUCCESS, bold: true, align: "center", margin: 0,
  });
  slide.addText([
    { text: "Cake sales", options: { breakLine: true } },
    { text: "Drink sales", options: { breakLine: true } },
    { text: "Raffle", options: {} },
  ], {
    x: x2 + 0.15, y: y + 0.55, w: colW - 0.3, h: cardH - 0.7,
    fontSize: 13.5, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
    paraSpaceAfter: 6,
  });
}

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  titleSlide(pres, UNIT_TITLE, "Session 3: Make a plan — a fundraiser budget",
    `Year 6 Numeracy | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3-4. Daily Review (graphs) + reveal
  withReveal(
    () => dailyReviewSlide(pres,
      "Daily Review: Reading a graph",
      [
        "Which day sold the most cakes?",
        "How many MORE on Friday than Monday?",
        "Why does it matter that the graph starts at 0?",
      ],
      NOTES_DR_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, lg.safeBottom - lg.panelTopPadded, { strip: C.SECONDARY });
        slide.addText("Cakes sold per day", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.10, w: lg.rightW - 0.4, h: 0.28,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
        });
        miniBarGraph(slide,
          lg.rightX + 0.25, lg.panelTopPadded + 0.50, lg.rightW - 0.5, (lg.safeBottom - lg.panelTopPadded) - 0.70,
          [
            { label: "Mon", value: 10 },
            { label: "Wed", value: 25 },
            { label: "Fri", value: 35, fill: C.ACCENT },
          ],
          C.SECONDARY);
      }),
    (slide) => {
      addRevealAnswerBar(slide,
        ["Friday (35)", "25 more", "fair comparison"],
        { color: C.SUCCESS, fontSize: 18 });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // 5-6. Fluency + reveal (subtraction vertical algorithm)
  withReveal(
    () => fluencySlide(pres, "Fluency: Subtraction",
      ["4 000 − 2 165", "7 206 − 3 489", "8 500 − 5 743"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, ["1 835", "3 717", "2 757"], { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // 7. Launch
  contentSlide(pres, "Launch", C.SECONDARY,
    "Money out, money in",
    [
      "Our class runs a cake stall.",
      "First we SPEND, then we EARN.",
      "What goes on each side?",
      "Tell your partner one of each.",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => { moneyCard(slide, lg); });

  // 8. LI / SC
  liSlide(pres,
    "We are learning to create a simple budget that combines operations to plan a fundraising event.",
    [
      "I can read a budget to find the costs and the income.",
      "I can combine operations to work out total cost and profit.",
      "I can decide if a plan makes a profit and suggest one change.",
    ],
    NOTES_LI_SC, FOOTER);

  // 9. I Do 1 — total cost (× then +)
  workedExSlide(pres, 2, "I Do", "Total cost — multiply, then add",
    [
      "A budget combines operations.",
      "",
      "Flour: 3 packs × $4 each.",
      "Equal groups → 3 × $4 = $12.",
      "",
      "Icing: $6.",
      "",
      "Add the costs: $12 + $6 = $18.",
      "Total cost = $18.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      budgetTable(slide, lg, {
        title: "Cake stall — costs",
        color: C.PRIMARY,
        rows: [
          { label: "Flour  3 × $4", amount: "$12" },
          { label: "Icing", amount: "$6" },
          { label: "TOTAL COST", amount: "$18", kind: "total" },
        ],
      });
    });

  // 10. I Do 2 — income & profit (× then −)
  workedExSlide(pres, 2, "I Do", "Income & profit — what is LEFT",
    [
      "Income: sell 40 cakes × $2 each.",
      "Equal groups → 40 × $2 = $80.",
      "",
      "Profit is what is LEFT after costs.",
      "",
      "Profit = income − costs.",
      "$80 − $18 = $62.",
      "",
      "Check: ≈ $80 − $20 = $60. Reasonable.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      budgetTable(slide, lg, {
        title: "Cake stall — profit",
        color: C.PRIMARY,
        rows: [
          { label: "Income  40 × $2", amount: "$80" },
          { label: "Costs", amount: "$18" },
          { label: "PROFIT  $80 − $18", amount: "$62", kind: "profit" },
        ],
      });
    });

  // 11-12. CFU 1 + reveal — find the profit
  withReveal(
    () => cfuSlide(pres, "CFU", "Find the profit", "Show Me Boards",
      "A stall earns $95 income.\nIts costs are $32.\n\nWhat is the profit?\nWrite the number sentence and solve.",
      NOTES_CFU1_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, "$95 − $32 = $63 profit", { color: C.SUCCESS, label: "Answer" });
    }
  );

  // 13-14. We Do + reveal — full budget
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Budget the lemonade stall",
      [
        "With your partner.",
        "",
        "Costs: lemons & sugar $8, cups $4.",
        "Income: sell 50 cups × $1 each.",
        "",
        "Step 1: total cost.",
        "Step 2: income.",
        "Step 3: profit = income − costs.",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        budgetTable(slide, lg, {
          title: "Lemonade stall",
          color: C.SECONDARY,
          rows: [
            { label: "Lemons & sugar", amount: "$8" },
            { label: "Cups", amount: "$4" },
            { label: "Income  50 × $1", amount: "$50" },
            { label: "PROFIT = ?", amount: "?", kind: "profit" },
          ],
        });
      }),
    (slide) => {
      addRevealAnswerBar(slide,
        "Cost $12   •   Income $50   •   Profit $50 − $12 = $38",
        { color: C.SUCCESS, label: "Answer", fontSize: 18 });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // 15-16. CFU hinge + reveal — worth running?
  withReveal(
    () => cfuSlide(pres, "CFU", "Worth running?", "Thumbs Up or Thumbs Down",
      "A plan:\n\nIncome  =  $40\nCosts   =  $55\n\nIs this plan worth running?\nThumbs up = yes.  Thumbs down = no.",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide,
        "No — costs ($55) are higher than income ($40). It makes a $15 LOSS.",
        { color: C.ALERT, label: "Check", fontSize: 18 });
    }
  );

  // 17. You Do — budget planner
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: budget planner", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.25, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "work out the total cost.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "work out the income, then the profit.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "suggest one change to make more profit.", options: { fontSize: 17, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.16, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.40;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("The one rule to remember", {
      x: 0.7, y: panelY + 0.16, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
    });
    addTextOnShape(s, "PROFIT  =  INCOME  −  COSTS", {
      x: 1.6, y: panelY + 0.62, w: 6.8, h: 0.66, rectRadius: 0.08,
      fill: { color: C.SUCCESS },
    }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    s.addText("Costs higher than income? The plan makes a loss.", {
      x: 0.7, y: panelY + 1.45, w: 8.6, h: 0.34,
      fontSize: 14.5, fontFace: FONT_B, color: C.MUTED, italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // 18. Exit Ticket
  exitTicketSlide(pres,
    [
      "A fun run budget: income $120, costs $45. Work out the profit.",
      "Is this plan worth running? Explain why.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // 19. Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: across the unit, what do good problem solvers do before they trust an answer?",
      scItems: [
        "I can read a budget to find the costs and the income.",
        "I can combine operations to work out total cost and profit.",
        "I can decide if a plan makes a profit and suggest one change.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FinRea_Session3_Fundraiser_Budget.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  await generatePdfs();
  console.log("Session 3 build complete.");
}

// ─── PDFs ─────────────────────────────────────────────────────────────────

async function generatePdfs() {
  // ── Budget Planner (worksheet, with built-in enabling structure) ───────
  await (async () => {
    const doc = createPdf({ title: PLANNER_RES.name });
    let y = addPdfHeader(doc, PLANNER_RES.name, {
      subtitle: "Work out the total cost, the income and the profit. Then suggest one improvement.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addTipBox(doc,
      "PROFIT = INCOME - COSTS.  Multiply for equal groups, add the costs, then subtract to find the profit.",
      y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Plan A — Popcorn stall", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Costs:", y, { fontSize: 11, color: "2D3142" });
    y = addWriteLine(doc, "Kernels (given): $9", y);
    y = addWriteLine(doc, "Bags:   5 packs × $2  =  __________", y);
    y = addWriteLine(doc, "TOTAL COST  =  __________", y);
    y = addBodyText(doc, "Income:", y, { fontSize: 11, color: "2D3142" });
    y = addWriteLine(doc, "Sell 60 bags × $1 each  =  __________", y);
    y = addBodyText(doc, "Profit:", y, { fontSize: 11, color: "2D3142" });
    y = addWriteLine(doc, "Profit  =  income − total cost  =  __________", y);

    y = addSectionHeading(doc, "Plan B — Badge stall", y, { color: C.SECONDARY });
    y = addWriteLine(doc, "Badges:  4 packs × $5  =  __________   (this is the total cost)", y);
    y = addWriteLine(doc, "Income:  sell 30 badges × $2  =  __________", y);
    y = addWriteLine(doc, "Profit  =  income − cost  =  __________", y);

    y = addSectionHeading(doc, "Decide & improve", y, { color: C.ACCENT });
    y = addWriteLine(doc, "Which plan makes MORE profit? __________   By how much? __________", y);
    y = addWriteLine(doc, "Suggest ONE change to make even more profit: ______________________________", y);
    y = addWriteLine(doc, "Why would your change help? ______________________________________________", y);

    addPdfFooter(doc, `Session ${SESSION} | Fundraiser Budget Planner | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, PLANNER_RES.fileName));
    console.log("PDF written: " + PLANNER_RES.fileName);
  })();

  // ── Answer Key ─────────────────────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the fundraiser budget planner.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Plan A — Popcorn stall", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Bags: 5 x $2 = $10.   Total cost: $9 + $10 = $19.", y);
    y = addBodyText(doc, "Income: 60 x $1 = $60.   Profit: $60 - $19 = $41.", y);

    y = addSectionHeading(doc, "Plan B — Badge stall", y, { color: C.SECONDARY });
    y = addBodyText(doc, "Cost: 4 x $5 = $20.   Income: 30 x $2 = $60.   Profit: $60 - $20 = $40.", y);

    y = addSectionHeading(doc, "Decide & improve", y, { color: C.ACCENT });
    y = addBodyText(doc, "Plan A makes more profit ($41 vs $40), by $1.", y);
    y = addBodyText(doc, "Sample improvements (accept any reasonable, justified change): charge a little more per item (raises income), buy supplies in bulk for a lower price (lowers costs), or sell more items. Look for a reason linked to income or costs.", y);

    y = addTipBox(doc,
      "Watch for: students who write the income as the profit. Profit is what is LEFT after costs. A student who reports $60 profit for Plan A has not subtracted the $19 of costs.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // ── Year 8 Extension Task ──────────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: EXT_RES.name });
    let y = addPdfHeader(doc, EXT_RES.name, {
      subtitle: "Stretch your financial reasoning with percentages. Show your working.",
      color: C.SECONDARY,
      lessonInfo: `Session ${SESSION} | Aimed at Year 8 | Extension`,
    });

    y = addTipBox(doc,
      "Worked example - adding GST (10%): a $50 cost plus 10% GST. 10% of $50 is $5. Total = $50 + $5 = $55. A percentage is 'out of 100', so 10% means 10 for every 100.",
      y, { color: C.SECONDARY });

    y = addSectionHeading(doc, "Stretch problems", y, { color: C.SECONDARY });
    y = addProblem(doc, 1, "A supplier charges $80 for ingredients, plus 10% GST. What is the total cost?", y,
      { color: C.SECONDARY, writeLines: [{ label: "Working & answer:" }] });
    y = addProblem(doc, 2, "Cups normally cost $40. A 15% bulk discount applies. What is the discounted price?", y,
      { color: C.SECONDARY, writeLines: [{ label: "Working & answer:" }] });
    y = addProblem(doc, 3, "A stall's costs are $54. Each item sells for $4. How many items must they sell to break even (cover the costs)?", y,
      { color: C.SECONDARY, writeLines: [{ label: "Working & answer:" }] });
    y = addProblem(doc, 4, "A stall makes $150 income with $90 of costs. Find the profit, then write the profit as a percentage of the income.", y,
      { color: C.SECONDARY, writeLines: [{ label: "Working & answer:" }] });
    y = addProblem(doc, 5, "A cake stall sells cakes for $2 each. Fixed costs are $18. Find the profit if they sell 100 cakes. Why do the costs NOT change?", y,
      { color: C.SECONDARY, writeLines: [{ label: "Working & answer:" }] });

    addPdfFooter(doc, `Session ${SESSION} | Year 8 Extension Task | Financial Reasoning`);
    await writePdf(doc, path.join(OUT_DIR, EXT_RES.fileName));
    console.log("PDF written: " + EXT_RES.fileName);
  })();

  // ── Year 8 Extension Answer Key ────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: EXT_KEY_RES.name });
    let y = addPdfHeader(doc, EXT_KEY_RES.name, {
      subtitle: "Worked answers for the Year 8 extension task.",
      color: C.SECONDARY,
      lessonInfo: `Session ${SESSION} | Aimed at Year 8 | Extension`,
    });

    y = addSectionHeading(doc, "Worked answers", y, { color: C.SECONDARY });
    y = addBodyText(doc, "1.  GST: 10% of $80 = $8.   Total = $80 + $8 = $88.", y);
    y = addBodyText(doc, "2.  Discount: 15% of $40 = $6.   Discounted price = $40 - $6 = $34.", y);
    y = addBodyText(doc, "3.  Break-even: $54 / $4 = 13.5, so they must sell 14 items (you cannot sell half an item, so round up).", y);
    y = addBodyText(doc, "4.  Profit = $150 - $90 = $60.   As a percentage of income: 60 / 150 = 0.4 = 40%.", y);
    y = addBodyText(doc, "5.  Income = 100 x $2 = $200.   Profit = $200 - $18 = $182.   The costs are FIXED ($18), so selling more cakes does not change them - only the income grows.", y);

    y = addTipBox(doc,
      "Discussion point for Q5: in real life some costs are fixed (one-off, like a stall fee) and some are variable (grow with sales, like ingredients). Strong students may notice that ingredient costs would actually rise with 100 cakes.",
      y, { color: C.SECONDARY });

    addPdfFooter(doc, `Session ${SESSION} | Year 8 Extension Answer Key | Financial Reasoning`);
    await writePdf(doc, path.join(OUT_DIR, EXT_KEY_RES.fileName));
    console.log("PDF written: " + EXT_KEY_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
