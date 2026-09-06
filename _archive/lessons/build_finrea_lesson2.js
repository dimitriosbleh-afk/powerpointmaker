"use strict";

// Practical Problem Solving & Financial Reasoning (Year 6 Numeracy)
// Session 2 of 3: Model it with a bar model — choose the operation.
// VC2M5N09 — model an everyday situation with materials/diagrams/bar models;
//            decide which operations solve it; write it as a number sentence;
//            justify the choice of operation (additive & multiplicative).
// Daily Review: Tessellations (prior learning).
// Fluency: Subtraction vertical algorithm.
//
// Unit routine (consistent across all 3 sessions): ESTIMATE -> MODEL -> SOLVE -> CHECK.
// Self-contained: re-teaches bar-model modelling from scratch with worked examples,
// so a student who missed Session 1 can still access today's lesson.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine, addLinedArea, addProblem,
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

const SESSION = 2;
const TOTAL = 3;
const UNIT_TITLE = "Practical Problem Solving";
const FOOTER = `Financial Reasoning | Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`;
const OUT_DIR = "output/FinRea_Session2_Bar_Models";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Bar Model Problem Solving",
  "Draw a bar model, choose the operation, write a number sentence and solve.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers for the bar model problem solving sheet.");
const SCAFFOLD_RES = makeSessionResource(SESSION,
  "Bar Model Scaffold",
  "Pre-drawn bar models for students who need a starting structure.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, SCAFFOLD_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome back, mathematicians. Yesterday we estimated and checked. Today we model.
- When a word problem looks tricky, a bar model turns the words into a picture so we can see what to do.
- By the end you will be able to draw a bar model, choose whether to add, subtract or multiply, and solve it.

DO:
- Whiteboards and markers ready.
- Keep the practice sheet and scaffold face down until the You Do section.

TEACHER NOTES:
Session 2 of 3. The unit routine is Estimate, Model, Solve, Check. Today's focus is the Model and Solve moves, with an estimate-to-check at the end. This session stands alone - it re-teaches bar models from the start.

WATCH FOR:
- Students who jump to calculating before drawing. The picture comes first today.

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- Here is what we are using today.
- The practice sheet is for the You Do. The bar model scaffold is there for anyone who wants the bars already drawn.

DO:
- Print the practice sheet and answer key. Print the bar model scaffold for the students who need it.
- Whiteboards and markers out.

TEACHER NOTES:
Unit note: each session is self-contained so missed sessions are not a barrier. Today: one core worksheet, an answer key, and an enabling bar model scaffold. The bar models on the slides are the main visual model.

[General: Resources]`;

const NOTES_DR_Q = `SAY:
- Quick warm up from our geometry work.
- Read each one and write your answer on your whiteboard.
- This is earlier learning, not today's money work.

DO:
- Display the three prompts.
- Allow about 90 seconds.

TEACHER NOTES:
Daily Review revisits Tessellations - prior learning and retrieval. Keeps shape and angle knowledge warm while today's lesson moves to problem solving.

WATCH FOR:
- Students who recall that angles around a point make 360 - secure.
- Students unsure which shapes tessellate - quick reminder with pattern blocks.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let us check.
- Squares, equilateral triangles and regular hexagons all tessellate on their own.
- A regular pentagon does not tessellate on its own - false.
- The angles around a point add to 360 degrees.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
The 360 fact explains why some shapes tessellate and others leave gaps. Connect briefly if time allows.

WATCH FOR:
- Students who can explain the gap a pentagon leaves - strong.
- Students who said a pentagon tessellates - revisit with pattern blocks later.

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency. Subtraction with the vertical algorithm again.
- Columns lined up, work top to bottom, regroup when you need to.
- Whisper, then write.

DO:
- Display the three subtractions.
- Allow about 75 seconds.

TEACHER NOTES:
Same subtraction focus across the unit. Brisk retrieval, not new teaching.

WATCH FOR:
- Clean regrouping - secure.
- Smaller-from-larger habit in a column - reteach regrouping.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check your columns.
- 5000 minus 1836 is 3164.
- 6401 minus 2758 is 3643.
- 9070 minus 3485 is 5585.

DO:
- Click to reveal.
- Tick and fix.

TEACHER NOTES:
The middle one needs regrouping in more than one column. Watch for students who forget to reduce the next column after regrouping.

WATCH FOR:
- Correct multi-column regrouping - secure.
- Off-by-a-thousand errors - the regroup was not carried through.

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LAUNCH = `SAY:
- Here is a quick problem. A stall sold 7 cupcakes for 3 dollars each.
- I am not asking for the answer yet. I am asking how you would work it out.
- Would you add, or would you multiply? Turn and tell your partner why.

DO:
- Display the 7 cupcake groups, each 3 dollars.
- Give 30 seconds of partner talk.
- Take two or three responses. Listen for "equal groups" language.

TEACHER NOTES:
The launch surfaces the idea that 7 equal groups of 3 can be added or, more efficiently, multiplied. This sets up choosing the operation, which is today's focus. The picture shows the equal groups so the multiply idea is visible.

WATCH FOR:
- Students who see equal groups and say multiply - ready.
- Students who only see "add them all" - valid, but nudge towards the quicker multiply.

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Here is today's learning.
- We are learning to model a money word problem with a bar model and choose the operation to solve it.
- Read the success criteria with me.

DO:
- Choral read the learning intention.
- Read each I can statement.

TEACHER NOTES:
SC1 is reachable for everyone - drawing a bar model that matches. SC2 is the core target the exit ticket assesses - choosing the operation and writing the number sentence. SC3 reaches into solving and checking reasonableness.

WATCH FOR:
- Students repeating the language - tracking with us.

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Let us work through this together. A stall made 45 dollars on Saturday and 38 dollars on Sunday. How much altogether?
- Watch how I draw the bar model first. I make one box for Saturday's 45 dollars and another box for Sunday's 38 dollars.
- The two boxes joined together are the whole amount, and that is what I am looking for. I mark the whole with a question mark.
- Both parts are known and I want the whole, so this is an addition. I write 45 plus 38.
- 45 plus 38 is 83. So the stall made 83 dollars altogether.
- Quick check: 45 and 38 are about 40 and 40, so about 80. My answer of 83 is reasonable.

DO:
- Draw the two part boxes, then the whole bracket with a question mark.
- Point to the parts as you name them.
- Say the number sentence aloud as you write it.

TEACHER NOTES:
Parts known, whole unknown means addition. The bar model makes that visible. End with the estimate-to-check from Session 1.

MISCONCEPTIONS:
- Misconception: pick the operation from a key word like "altogether".
  Why: students are taught to hunt for trigger words.
  Impact: key words mislead on multi-step or comparison problems.
  Quick correction: draw the bar model and ask what is known and what is missing.

WATCH FOR:
- Students who draw parts then a whole - secure.
- Students who guess the operation without drawing - bring them back to the model.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now a different shape of problem. A class sold 6 raffle books for 4 dollars each. How much did they raise?
- Watch the bar model. This time I draw 6 equal boxes, because there are 6 equal groups, and each box is worth 4 dollars.
- Equal groups means I can multiply. I write 6 times 4.
- 6 times 4 is 24. So they raised 24 dollars.
- Check: 6 groups of about 4 dollars is about 24 dollars. That matches, so it is reasonable.

DO:
- Draw 6 equal boxes, each labelled 4 dollars.
- Run your finger along the boxes as you count the equal groups.
- Write 6 times 4 equals 24.

TEACHER NOTES:
Equal groups means multiplication. The 6 equal boxes are the key difference from the first model. Keep linking the picture to the number sentence.

MISCONCEPTIONS:
- Misconception: add the two numbers because they are both there.
  Why: addition is the most familiar operation.
  Impact: 6 plus 4 gives 10, which does not match equal groups.
  Quick correction: are the groups equal? If yes, multiply.

WATCH FOR:
- Students who draw equal boxes and multiply - secure.
- Students who write 6 plus 4 - point back to the equal groups.

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1_Q = `SAY:
- Quick check on your whiteboards.
- Look at the bar model. The whole is 90 dollars. One part is 55 dollars. The other part is missing.
- Which number sentence finds the missing part? Write it and solve it.

DO:
- Display the bar model.
- Allow about 45 seconds.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Show me your number sentence and answer on three, two, one, show."
- Scan for: 90 minus 55 equals 35.
PROCEED: If 80 percent have 90 minus 55 equals 35, click to reveal and move on.
PIVOT: Most likely misconception - students add 90 and 55 because both numbers are shown.
- Reteach: "The whole is already there - 90. We do not want a bigger number. We take away the known part."
- Re-check: "Whole take away known part. What is the sentence?"

TEACHER NOTES:
Whole known, one part known, other part missing means subtraction. The bar model shows we are finding a missing part inside the whole.

WATCH FOR:
- 90 minus 55 equals 35 - secure.
- 90 plus 55 - they did not notice the whole was already known.

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. The cake stall raised 72 dollars. The drink stall raised 48 dollars. How much MORE did the cake stall raise?
- Look at the comparison bars. The top bar is the cake stall. The bottom bar is the drink stall plus the extra part we are looking for.
- Step 1: what are we comparing? Step 2: which operation finds the difference? Step 3: write it and solve it on your whiteboard.

DO:
- Keep the comparison bars visible.
- Walk and scan.
- Take 60 seconds before revealing.

TEACHER NOTES:
A comparison problem. "How much more" means find the difference, which is subtraction. The two bars side by side make the gap visible. Different structure from the I Do examples on purpose.

WATCH FOR:
- Pairs who write 72 minus 48 - secure.
- Pairs who add - point to the gap between the bars.

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Let us check.
- We are finding the difference, so we subtract. 72 minus 48.
- 72 minus 48 is 24. The cake stall raised 24 dollars more.
- Quick check: 72 and 48 are about 70 and 50, a gap of about 20. Our answer of 24 is reasonable.

DO:
- Click to reveal.
- Run the subtraction once more aloud.

TEACHER NOTES:
The reveal restates the operation and the answer, then the estimate check. Look for students who wrote 120 - they added instead of finding the difference.

WATCH FOR:
- Self-correction on the reveal - secure.
- Answers near 120 - they added the two bars.

[Stage 3: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge question. Read the problem on the screen.
- A team bought 5 tickets that cost 6 dollars each. What is the total cost?
- Which number sentence matches the problem? Hold up 1, 2 or 3 fingers for your choice.

DO:
- Display the three options.
- Give 15 seconds of thinking time.

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "One, two or three. Show me your fingers."
- Scan for: option 2, which is 5 times 6.
PROCEED: If 80 percent choose 5 times 6, click to reveal.
PIVOT: Most likely misconception - students pick 5 plus 6 because both numbers appear.
- Reteach: "5 equal tickets at 6 dollars each is 5 equal groups. Equal groups means multiply."
- Re-check: "Equal groups - which operation?"

TEACHER NOTES:
Each wrong option maps to a misconception. Option 1 (5 plus 6) is the add-the-numbers error. Option 3 (6 minus 5) is the grab-an-operation error. Option 2 is correct.

WATCH FOR:
- Confident 2 with reasoning - secure.
- 1 or 3 - reteach the equal-groups idea.

[Stage 2: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Time to work on your own. Take the practice sheet.
- For each problem: draw the bar model first, choose the operation, write the number sentence, then solve. Finish by checking your answer is reasonable.
- The bar model comes first every time.

DO:
- Distribute the practice sheet. Hand the bar model scaffold to anyone who wants the bars already drawn.
- Circulate. Sit with the enabling group first.
- Watch for students drawing before calculating.

TEACHER NOTES:
The sheet mixes additive and multiplicative problems so students have to choose, not just repeat one operation. Question content is different from the We Do.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Bar Model Scaffold. The boxes are already drawn - fill in the numbers, then write the number sentence.
- Extra Notes: Do the first problem together with this group, naming what is known and what is missing.
EXTENDING PROMPT:
- Task: The two-step challenge - a problem where you multiply first, then add. Draw both bars and write both number sentences.
- Extra Notes: This previews Session 3 budgeting. Push for a clear two-step bar model.

WATCH FOR:
- Students who draw the model before choosing the operation - secure.
- Students who grab an operation from key words - redirect to the bar model.

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Last task, on your own, on your whiteboard.
- A school sold 8 wristbands for 5 dollars each.
- One, draw a bar model and write the number sentence. Two, solve it and check your answer is reasonable.

DO:
- Display the prompt.
- Allow about 3 minutes.
- Collect whiteboards or photograph for your records.

TEACHER NOTES:
Exit ticket assesses SC2 (choose the operation and write the number sentence) and reaches SC3 (solve and check). This is an equal-groups problem: 8 times 5 equals 40. Internal target is SC2. Do not display any SC label to students.

WATCH FOR:
- A bar of 8 equal boxes and 8 times 5 equals 40 - secure.
- 8 plus 5 - the equal-groups idea needs another model in Session 3.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Look back at our success criteria.
- Show me thumbs for each one.
- Turn and tell your partner: how does drawing a bar model help you choose the operation?

DO:
- Read each I can statement.
- Use thumbs to self-assess.
- Note who is still unsure for tomorrow.

TEACHER NOTES:
The big idea is that the model reveals the operation - we read the picture, not the key words. Students who can say "the picture shows what is known and what is missing" have the relational understanding.

WATCH FOR:
- Students who explain the model-first idea - secure.
- Students who still hunt for key words - revisit in Session 3.

[General: Closing | VTLM 2.0: Reflection]`;

// ─── Visual helper: bar models ───────────────────────────────────────────────

// Draws a titled card containing one or more horizontal bars. Each bar is a
// row of labelled segments (rect + centred label). Auto-sizes the card to the
// content. Used for part-part-whole, equal-groups, and comparison models.
function barPanel(slide, lg, opts) {
  const o = opts;
  const x = lg.rightX;
  const w = lg.rightW;
  const y = lg.panelTopPadded;
  const color = o.color || C.PRIMARY;
  const bars = o.bars;
  const titleH = 0.46;
  let needed = titleH;
  bars.forEach((b) => {
    needed += (b.rowLabel ? 0.30 : 0) + (b.h || 0.62) + (b.gapAfter != null ? b.gapAfter : 0.18);
  });
  if (o.caption) needed += 0.46;
  needed += 0.14;
  const cardH = Math.min(lg.safeBottom - y, needed);
  addCard(slide, x, y, w, cardH, { strip: color });
  slide.addText(o.title, {
    x: x + 0.2, y: y + 0.10, w: w - 0.4, h: 0.30,
    fontSize: 15, fontFace: FONT_H, color, bold: true,
    align: "center", margin: 0,
  });
  const innerX = x + 0.3;
  const innerW = w - 0.6;
  let cy = y + titleH + 0.04;
  bars.forEach((bar) => {
    if (bar.rowLabel) {
      slide.addText(bar.rowLabel, {
        x: innerX, y: cy, w: innerW, h: 0.26,
        fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
        valign: "middle", margin: 0,
      });
      cy += 0.30;
    }
    const bh = bar.h || 0.62;
    const totalWeight = bar.segments.reduce((s, g) => s + (g.weight || 1), 0);
    const barW = innerW * (bar.widthRatio != null ? bar.widthRatio : 1);
    let sx = innerX;
    bar.segments.forEach((g) => {
      const sw = barW * (g.weight || 1) / totalWeight;
      slide.addShape("rect", {
        x: sx, y: cy, w: sw, h: bh,
        fill: { color: g.fill || C.WHITE },
        line: { color, width: 1.5 },
      });
      slide.addText(String(g.label), {
        x: sx, y: cy, w: sw, h: bh,
        fontSize: g.fontSize || 14, fontFace: FONT_H,
        color: g.textColor || C.CHARCOAL, bold: true,
        align: "center", valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });
      sx += sw;
    });
    cy += bh + (bar.gapAfter != null ? bar.gapAfter : 0.18);
  });
  if (o.caption) {
    slide.addText(o.caption, {
      x: innerX, y: cy, w: innerW, h: 0.4,
      fontSize: 12.5, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "top", margin: 0,
    });
  }
}

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // 1. Title
  titleSlide(pres, UNIT_TITLE, "Session 2: Model it with a bar model",
    `Year 6 Numeracy | Session ${SESSION} of ${TOTAL}`, NOTES_TITLE);

  // 2. Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // 3-4. Daily Review (Tessellations) + reveal
  withReveal(
    () => dailyReviewSlide(pres,
      "Daily Review: Tessellations",
      [
        "Name one shape that tessellates on its own.",
        "True or false: a regular pentagon tessellates on its own.",
        "Angles around a point add to ____ degrees.",
      ],
      NOTES_DR_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide,
        ["square / triangle / hexagon", "False", "360°"],
        { color: C.SUCCESS, fontSize: 20 });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // 5-6. Fluency + reveal (subtraction vertical algorithm)
  withReveal(
    () => fluencySlide(pres, "Fluency: Subtraction",
      ["5 000 − 1 836", "6 401 − 2 758", "9 070 − 3 485"],
      NOTES_FLUENCY_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide, ["3 164", "3 643", "5 585"], { color: C.SUCCESS });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // 7. Launch
  contentSlide(pres, "Launch", C.SECONDARY,
    "Add or multiply?",
    [
      "A stall sold 7 cupcakes.",
      "Each cupcake was $3.",
      "How would you work out the total?",
      "Add — or multiply? Why?",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      barPanel(slide, lg, {
        title: "7 cupcakes, $3 each",
        color: C.SECONDARY,
        bars: [
          {
            rowLabel: "Equal groups:",
            segments: [
              { label: "$3" }, { label: "$3" }, { label: "$3" }, { label: "$3" },
              { label: "$3" }, { label: "$3" }, { label: "$3" },
            ],
            h: 0.7,
          },
        ],
        caption: "7 equal groups of $3 — what does that tell you?",
      });
    });

  // 8. LI / SC
  liSlide(pres,
    "We are learning to model a money word problem with a bar model and choose the operation to solve it.",
    [
      "I can draw a bar model that matches a word problem.",
      "I can choose the operation and write a number sentence.",
      "I can solve the problem and check my answer is reasonable.",
    ],
    NOTES_LI_SC, FOOTER);

  // 9. I Do 1 — additive (total unknown)
  workedExSlide(pres, 2, "I Do", "Parts known, whole missing → add",
    [
      "Saturday: $45.   Sunday: $38.",
      "How much altogether?",
      "",
      "Draw a box for each part.",
      "The whole is what we want — mark it ?",
      "",
      "Both parts known, whole missing → ADD.",
      "$45 + $38 = $83.",
      "",
      "Check: ≈ $40 + $40 = $80. Reasonable.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      barPanel(slide, lg, {
        title: "Bar model — total",
        color: C.PRIMARY,
        bars: [
          {
            rowLabel: "Whole = ?",
            segments: [
              { label: "Sat $45", weight: 45 },
              { label: "Sun $38", weight: 38, fill: C.BG_LIGHT },
            ],
            h: 0.78,
          },
        ],
        caption: "Two parts join to make the whole → $45 + $38",
      });
    });

  // 10. I Do 2 — multiplicative (equal groups)
  workedExSlide(pres, 2, "I Do", "Equal groups → multiply",
    [
      "6 raffle books, $4 each.",
      "How much raised?",
      "",
      "Draw 6 EQUAL boxes, each $4.",
      "",
      "Equal groups → MULTIPLY.",
      "6 × $4 = $24.",
      "",
      "Check: 6 groups of ~$4 ≈ $24. Reasonable.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      barPanel(slide, lg, {
        title: "Bar model — equal groups",
        color: C.SECONDARY,
        bars: [
          {
            rowLabel: "6 groups of $4 = ?",
            segments: [
              { label: "$4" }, { label: "$4" }, { label: "$4" },
              { label: "$4" }, { label: "$4" }, { label: "$4" },
            ],
            h: 0.78,
          },
        ],
        caption: "6 equal boxes → 6 × $4",
      });
    });

  // 11-12. CFU 1 + reveal — missing part (subtraction)
  withReveal(
    () => workedExSlide(pres, 2, "CFU", "Find the missing part",
      [
        "On your whiteboard.",
        "",
        "The whole is $90.",
        "One part is $55.",
        "",
        "Which number sentence finds",
        "the missing part? Solve it.",
      ],
      NOTES_CFU1_Q, FOOTER,
      (slide, lg) => {
        barPanel(slide, lg, {
          title: "Bar model",
          color: C.ALERT,
          bars: [
            {
              rowLabel: "Whole = $90",
              segments: [
                { label: "$55", weight: 55 },
                { label: "?", weight: 35, fill: C.ALERT, textColor: C.WHITE },
              ],
              h: 0.82,
            },
          ],
          caption: "Whole known, one part known → find the other part",
        });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "$90 − $55 = $35", { color: C.SUCCESS, label: "Answer" });
    }
  );

  // 13-14. We Do + reveal — comparison (difference)
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "How much more? → find the difference",
      [
        "With your partner.",
        "",
        "Cake stall raised $72.",
        "Drink stall raised $48.",
        "",
        "How much MORE did the cake stall raise?",
        "",
        "Write the number sentence and solve.",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        barPanel(slide, lg, {
          title: "Comparison bars",
          color: C.SECONDARY,
          bars: [
            {
              rowLabel: "Cake stall",
              segments: [{ label: "$72", weight: 1 }],
              h: 0.6, gapAfter: 0.16,
            },
            {
              rowLabel: "Drink stall",
              segments: [
                { label: "$48", weight: 48 },
                { label: "? more", weight: 24, fill: C.ALERT, textColor: C.WHITE },
              ],
              h: 0.6,
            },
          ],
          caption: "The gap between the bars is the difference",
        });
      }),
    (slide) => {
      addRevealAnswerBar(slide, "$72 − $48 = $24 more", { color: C.SUCCESS, label: "Answer" });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // 15-16. CFU hinge + reveal — choose the number sentence
  withReveal(
    () => cfuSlide(pres, "CFU", "Which number sentence?", "Finger Voting",
      "A team bought 5 tickets that cost $6 each.\nWhat is the total cost?\n\n1.   5 + 6\n2.   5 × 6\n3.   6 − 5",
      NOTES_HINGE_Q, FOOTER),
    (slide) => {
      addRevealAnswerBar(slide,
        "Option 2:  5 × $6 = $30   (5 equal groups → multiply)",
        { color: C.SUCCESS, label: "Answer", fontSize: 18 });
    }
  );

  // 17. You Do — practice sheet
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Your turn: model & solve", { color: STAGE_COLORS["4"] });

    addCard(s, 0.5, CONTENT_TOP, 9.0, 1.25, { strip: STAGE_COLORS["4"] });
    s.addText([
      { text: "First: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "draw the bar model.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Next: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "choose the operation and write the number sentence.   ", options: { fontSize: 17, color: C.CHARCOAL } },
      { text: "Then: ", options: { fontSize: 19, color: C.ALERT, bold: true } },
      { text: "solve and check it is reasonable.", options: { fontSize: 17, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: CONTENT_TOP + 0.16, w: 8.5, h: 0.95,
      fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const panelY = CONTENT_TOP + 1.40;
    addCard(s, 0.5, panelY, 9.0, SAFE_BOTTOM - panelY, { strip: C.SECONDARY });
    s.addText("Read the bar model, not the key words", {
      x: 0.7, y: panelY + 0.14, w: 8.6, h: 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true,
      align: "center", margin: 0,
    });
    addTextOnShape(s, "Parts known, whole missing  →  ADD", {
      x: 1.0, y: panelY + 0.54, w: 8.0, h: 0.40, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, { fontSize: 15.5, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(s, "Whole known, one part missing  →  SUBTRACT", {
      x: 1.0, y: panelY + 1.02, w: 8.0, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, { fontSize: 15.5, fontFace: FONT_H, color: C.WHITE, bold: true });
    addTextOnShape(s, "Equal groups  →  MULTIPLY", {
      x: 1.0, y: panelY + 1.50, w: 8.0, h: 0.40, rectRadius: 0.08,
      fill: { color: C.SECONDARY },
    }, { fontSize: 15.5, fontFace: FONT_H, color: C.WHITE, bold: true });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  })();

  // 18. Exit Ticket
  exitTicketSlide(pres,
    [
      "A school sold 8 wristbands for $5 each. Draw a bar model and write the number sentence.",
      "Solve it, then check your answer is reasonable.",
    ],
    NOTES_EXIT, FOOTER, { assessesSc: 2 });

  // 19. Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: how does a bar model help you choose the operation?",
      scItems: [
        "I can draw a bar model that matches a word problem.",
        "I can choose the operation and write a number sentence.",
        "I can solve the problem and check my answer is reasonable.",
      ],
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "FinRea_Session2_Bar_Models.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  await generatePdfs();
  console.log("Session 2 build complete.");
}

// ─── PDFs ─────────────────────────────────────────────────────────────────

// Draw a labelled bar (row of segments) spanning the content width in a PDF.
// Returns the y below the bar. Boxes are outlined (writable); known values
// are printed, unknowns shown as a light "?".
function pdfBarRow(doc, y, rowLabel, segments, accent) {
  const left = doc.page.margins.left;
  const w = doc.page.width - doc.page.margins.left - doc.page.margins.right;
  const stroke = "#" + accent;
  if (rowLabel) {
    doc.fontSize(10).font("Sans-Bold").fillColor("#2D3142");
    doc.text(rowLabel, left, y);
    y = doc.y + 2;
  }
  const h = 34;
  const totalWeight = segments.reduce((s, g) => s + (g.weight || 1), 0);
  let sx = left;
  segments.forEach((g) => {
    const sw = w * (g.weight || 1) / totalWeight;
    doc.save();
    doc.roundedRect(sx, y, sw, h, 2).lineWidth(1).strokeColor(stroke).stroke();
    doc.restore();
    doc.fontSize(11).font("Sans-Bold").fillColor(g.muted ? "#9CA3AF" : "#2D3142");
    doc.text(String(g.label || ""), sx, y + h / 2 - 6, { width: sw, align: "center" });
    sx += sw;
  });
  return y + h + 10;
}

async function generatePdfs() {
  // ── Worksheet ──────────────────────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Draw a bar model, choose the operation, write a number sentence and solve.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addTipBox(doc,
      "Draw the bar model FIRST. Parts known, whole missing = add. Whole known, one part missing = subtract. Equal groups = multiply. Then check your answer is reasonable.",
      y, { color: C.ACCENT });

    const problems = [
      "A cake stall made $36 on Friday and $29 on Saturday. How much altogether?",
      "A class sold 7 raffle tickets for $5 each. How much did they raise?",
      "A class needs $80 for an excursion. They have raised $52. How much more do they need?",
      "The senior team raised $96. The junior team raised $74. How much MORE did the seniors raise?",
    ];
    problems.forEach((p, i) => {
      y = addProblem(doc, i + 1, p, y, { color: C.PRIMARY });
      y = addBodyText(doc, "Draw your bar model:", y, { fontSize: 10, color: "6B7280" });
      y = addLinedArea(doc, y, 2, { lineSpacing: 22 });
      y = addWriteLine(doc, "Number sentence: ______________________     Answer: __________", y);
      y += 4;
    });

    // Keep the challenge heading with its body — break to a fresh page if the
    // whole block (heading + prompt + drawing space) would not fit.
    if (y > doc.page.height - doc.page.margins.bottom - 150) {
      doc.addPage();
      y = doc.page.margins.top;
    }
    y = addSectionHeading(doc, "Challenge — two steps", y, { color: C.SECONDARY });
    y = addBodyText(doc, "A stall sold 9 muffins at $3 each AND 6 juices at $2 each. How much did it raise in total? Draw both bars and write both number sentences.", y);
    y = addLinedArea(doc, y, 3, { lineSpacing: 22 });

    addPdfFooter(doc, `Session ${SESSION} | Bar Model Problem Solving | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // ── Answer Key ─────────────────────────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Worked answers for the bar model problem solving sheet.",
      color: C.PRIMARY,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addSectionHeading(doc, "Problems", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1.  Parts known, whole missing -> ADD.   $36 + $29 = $65.", y);
    y = addBodyText(doc, "2.  Equal groups -> MULTIPLY.   7 x $5 = $35.", y);
    y = addBodyText(doc, "3.  Whole known, one part missing -> SUBTRACT.   $80 - $52 = $28 more needed.", y);
    y = addBodyText(doc, "4.  Comparison, find the difference -> SUBTRACT.   $96 - $74 = $22 more.", y);

    y = addSectionHeading(doc, "Challenge — two steps", y, { color: C.SECONDARY });
    y = addBodyText(doc, "Muffins: 9 x $3 = $27.   Juices: 6 x $2 = $12.   Total: $27 + $12 = $39.", y);

    y = addTipBox(doc,
      "Watch for: students who pick an operation from a key word instead of the bar model. A student who writes 7 + 5 for problem 2 has not noticed the equal groups. Redirect to the model: equal boxes means multiply.",
      y, { color: C.ACCENT });

    addPdfFooter(doc, `Session ${SESSION} | Answer Key | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // ── Bar Model Scaffold (enabling) ──────────────────────────────────────
  await (async () => {
    const doc = createPdf({ title: SCAFFOLD_RES.name });
    let y = addPdfHeader(doc, SCAFFOLD_RES.name, {
      subtitle: "The bar models are drawn for you. Fill in the numbers, then write the number sentence.",
      color: C.SUCCESS,
      lessonInfo: `Session ${SESSION} of ${TOTAL} | Year 6 Numeracy`,
    });

    y = addTipBox(doc,
      "Look at the bars. What is known? What is missing? That tells you whether to add, subtract or multiply.",
      y, { color: C.SUCCESS });

    // Problem 1 — additive total
    y = addSectionHeading(doc, "Problem 1 — $36 on Friday and $29 on Saturday. Altogether?", y, { color: C.PRIMARY });
    y = pdfBarRow(doc, y, "Whole = ?  (this is what we want)", [
      { label: "$36", weight: 36 },
      { label: "$29", weight: 29 },
    ], C.PRIMARY);
    y = addBodyText(doc, "Both parts are known, the whole is missing -> ADD.", y, { fontSize: 10, color: "6B7280" });
    y = addWriteLine(doc, "Number sentence: ______________________     Answer: __________", y);
    y += 6;

    // Problem 2 — equal groups
    y = addSectionHeading(doc, "Problem 2 — 7 raffle tickets at $5 each. How much raised?", y, { color: C.SECONDARY });
    y = pdfBarRow(doc, y, "7 equal groups of $5  (total = ?)", [
      { label: "$5", weight: 1 }, { label: "$5", weight: 1 }, { label: "$5", weight: 1 },
      { label: "$5", weight: 1 }, { label: "$5", weight: 1 }, { label: "$5", weight: 1 },
      { label: "$5", weight: 1 },
    ], C.SECONDARY);
    y = addBodyText(doc, "Equal boxes -> MULTIPLY.", y, { fontSize: 10, color: "6B7280" });
    y = addWriteLine(doc, "Number sentence: ______________________     Answer: __________", y);
    y += 6;

    // Problem 3 — missing part
    y = addSectionHeading(doc, "Problem 3 — Need $80. Raised $52. How much more?", y, { color: C.ALERT });
    y = pdfBarRow(doc, y, "Whole = $80", [
      { label: "$52", weight: 52 },
      { label: "? more", weight: 28, muted: true },
    ], C.ALERT);
    y = addBodyText(doc, "Whole known, one part known -> SUBTRACT to find the missing part.", y, { fontSize: 10, color: "6B7280" });
    y = addWriteLine(doc, "Number sentence: ______________________     Answer: __________", y);

    addPdfFooter(doc, `Session ${SESSION} | Bar Model Scaffold | Year 6 Numeracy`);
    await writePdf(doc, path.join(OUT_DIR, SCAFFOLD_RES.fileName));
    console.log("PDF written: " + SCAFFOLD_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
