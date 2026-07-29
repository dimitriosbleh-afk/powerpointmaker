"use strict";

// Multiplication Facts Unit — Lesson 4: x9 + Consolidation
// Session 4, Grade 3/4 Numeracy, Variant 0
// x9 = x10 - one group
// Consolidation of x2, x3, x4, x5, x6, x8, x9, x10
// DR: Addition Algorithm (Vertical, 4- & 5-digits)
// Fluency: Subtraction Algorithm (Vertical, 3- & 4-digits)

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade34", 0);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar,
  withReveal, STAGE_COLORS,
} = T;

const SESSION = 4;
const FOOTER = "Multiplication Facts | Session 4 of 4 | Grade 3/4 Numeracy";
const OUT_DIR = "output/MF_Session4_x9_Consolidation";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Mixed practice - all multiplication tables 2-10.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with worked answers.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Strategy toolkit reference card with pre-filled examples for each table.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Self-contained investigation into multiplication patterns in the hundred chart.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ───────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Last session of our multiplication facts unit
- We have one more table to learn: the 9s. And there is a brilliant pattern to help you
- Then we consolidate everything: x2, x3, x4, x5, x6, x8, x9, x10
- By the end of today, you will have a strategy for every single table

DO:
- Display title slide
- Have whiteboards and worksheets ready

TEACHER NOTES:
Session 4 of 4. x9 uses the x10 - one group strategy. Then consolidation mixes all tables, requiring strategy selection. This is the capstone session.

WATCH FOR:
- Students who have gaps in earlier tables -- they will need the enabling scaffold (strategy toolkit reference)

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Last daily review for this unit -- addition algorithm
- Four problems, whiteboards, vertical algorithm

DO:
- Display problems. 90 seconds each.

TEACHER NOTES:
Final DR in the unit. Problems include 5-digit with multiple renamings. By now, most students should be fluent.

WATCH FOR:
- Compare accuracy with Session 1 -- growth should be visible

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check answers. Tick and fix.
- This is our last DR for this unit. How are you going compared to Session 1?

DO:
- Reveal answers. Students self-check.
- Brief acknowledgement of progress since Session 1

TEACHER NOTES:
Closing the loop on DR. Noting progress reinforces growth mindset.

WATCH FOR:
- Students who have shown consistent improvement -- acknowledge privately

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Last fluency session for this unit too. Subtraction algorithm, 2 minutes.

DO:
- Display problems. Time 2 minutes.

TEACHER NOTES:
Final fluency session. Problems are the most challenging yet -- renaming across multiple zeros and multi-step operations.

WATCH FOR:
- Overall improvement in speed and accuracy compared to Session 1

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention together
- Read from slide: "We are learning to use the x10 minus one group strategy for x9 facts and to choose the best strategy for any multiplication fact"
- Our success criteria today build on everything from the whole unit

DO:
- Choral read LI and SCs

TEACHER NOTES:
SC1 is x9 specifically. SC2 is strategy selection across all tables. SC3 extends to explaining the network of strategies and why they connect. This is the culminating assessment point.

WATCH FOR:
- Students who are confident about x9 but unsure about strategy selection for other tables -- the consolidation practice will address this

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO_X9 = `SAY:
- The 9s. These look hard but they have a beautiful pattern.
- Strategy: x9 = x10 minus one group.
- Watch: 9 x 7
- Step 1: What is 10 x 7? Easy -- 70. (Put a zero on the end.)
- Step 2: Take away one group of 7: 70 - 7 = 63.
- So 9 x 7 = 63.
- Why does this work? Because 9 groups is one less than 10 groups.

DO:
- Write each step clearly
- Point to the visual
- Show the connection: 10 groups minus 1 group = 9 groups

TEACHER NOTES:
x9 = x10 - one group is the most efficient strategy for x9. It leverages x10 (the easiest table) and requires only one subtraction. The digit sum pattern (digits always sum to 9) is also useful as a check.

MISCONCEPTIONS:
- Misconception: Students subtract the multiplier instead of one group (e.g., 9 x 7: 70 - 9 = 61 instead of 70 - 7 = 63)
  Why: Confusion about which number represents the group size
  Impact: Systematic errors across all x9 facts
  Quick correction: "70 is ten groups of 7. You want nine groups of 7. Remove ONE group. What is one group? 7. So 70 - 7 = 63."

WATCH FOR:
- Students who confuse which number to subtract -- always subtract the group size, not 9
- Students who already know x9 facts by rote -- still teach the strategy so they have a backup and can verify

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_PATTERN = `SAY:
- Let me show you a bonus pattern for the 9s
- Look at all the x9 answers: 9, 18, 27, 36, 45, 54, 63, 72, 81, 90
- Notice anything? The tens digit goes up by 1: 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
- The ones digit goes DOWN by 1: 9, 8, 7, 6, 5, 4, 3, 2, 1, 0
- And here is the magic: the digits ALWAYS add up to 9. 1+8=9. 2+7=9. 3+6=9.
- This is a great way to CHECK your x9 answers

DO:
- Display the digit pattern visually
- Have students verify by adding digits of several x9 products
- Emphasise: this is a checking tool, the x10-minus-one-group strategy is the primary method

TEACHER NOTES:
The digit sum = 9 pattern is a powerful verification tool. It is not a primary calculation strategy because it requires students to already know the tens digit. But it is excellent for self-checking.

WATCH FOR:
- Students who are fascinated by the pattern -- channel that into explaining WHY the digits sum to 9 (extending)
- Readiness signal: students spontaneously checking their x9 answers by adding digits

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. Think-pair-share.
- 9 x 8. Think for 5 seconds. Share with your partner. Be ready.

DO:
- 5 seconds think time
- 15 seconds pair sharing
- Cold call 3 pairs

CFU CHECKPOINT:
Technique: Think-Pair-Share
Script:
- Say: "9 x 8. Think... now share with your partner how you would work it out."
- After 15 seconds: "Pair at the back, what did you get?" [72] "How?" [10 x 8 = 80, minus 8 = 72]
- Verify: "Add the digits: 7 + 2 = 9. Checks out."
PROCEED: If 80%+ of pairs share 72 with correct strategy, move to consolidation.
PIVOT: Most likely misconception -- students subtract 9 instead of 8 (getting 71). Reteach: "How many groups do you want? 9. How many do you have? 10 groups of 8 = 80. Remove how many groups? Just 1 group. What is one group of 8? 8. So 80 - 8 = 72."

TEACHER NOTES:
Think-pair-share gives all students processing time before responding. Cold calling multiple pairs gives a representative sample.

WATCH FOR:
- The 80 - 9 = 71 error (subtracting the multiplier) -- this is the key diagnostic for x9

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_CONSOLIDATION = `SAY:
- Now we have strategies for EVERY table. Let me show you the complete toolkit.
- x2: double. x3: x2 + one more group. x4: double-double.
- x5: skip count (ends in 0 or 5). x10: put a zero.
- x6: double x3. x8: double x4.
- x9: x10 minus one group.
- The key skill now is CHOOSING the right strategy for any fact you see.

DO:
- Display the full strategy toolkit on screen
- Go through each one briefly with a quick example
- This is a review consolidation slide, not new teaching

TEACHER NOTES:
This consolidation slide maps the entire strategy network. Students should see that all facts connect back to x2 and x10 as anchor facts. The web of strategies means no fact needs to be memorised in isolation.

WATCH FOR:
- Students who can see the connections between strategies -- they have grasped the big idea of the unit
- Students who still see each table as separate -- they may need more time to internalise the connections

[Stage 2: Consolidation | VTLM 2.0: Making Connections]`;

const NOTES_WEDO_Q = `SAY:
- Partner challenge. Mixed table problems.
- Solve each one AND write which strategy you used.
- These are from ALL the tables we have learned.
- 1) 9 x 6.   2) 8 x 7.   3) 6 x 4.   4) 3 x 12.

DO:
- Partners work on whiteboards. 2 minutes.
- Circulate: check for correct strategy selection and naming

TEACHER NOTES:
We Do mixes all tables. The challenge is strategy selection, not just calculation. Different numbers from any previous session to ensure transfer.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the strategy toolkit reference card (Session 4 Enabling Scaffold). Look up which strategy matches each table, then apply it step by step.
- Extra Notes: Distribute Session 4 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: For each problem, find a SECOND strategy that also works. Compare: which was faster? Why?
- Extra Notes: Develops flexible strategic thinking.

WATCH FOR:
- Students who default to skip counting for everything instead of using derived-fact strategies -- prompt: "Which strategy from our toolkit would be fastest?"
- Partners who split up problems rather than working together -- redirect: "Both solve each one and compare"

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Answers and strategies.
- 9 x 6 = 54. Strategy: 10x6=60, minus 6 = 54.
- 8 x 7 = 56. Strategy: 4x7=28, double = 56.
- 6 x 4 = 24. Strategy: 3x4=12, double = 24.
- 3 x 12 = 36. Strategy: 2x12=24, plus 12 = 36.
- Who used a different valid strategy for any of these?

DO:
- Reveal answers. Cold call for strategies.
- Invite alternative strategies -- celebrate multiple valid approaches

TEACHER NOTES:
Multiple valid strategies for the same fact is a sign of number sense. Celebrating this reinforces flexible thinking.

WATCH FOR:
- Students who used a valid but less efficient strategy (e.g., skip counting 8s) -- acknowledge validity but ask: "Is there a faster way?"

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Final hinge check for the unit. Whiteboard.
- I will show you a fact. Write the answer AND the strategy name.
- Here it is: 6 x 9.

DO:
- Students write on whiteboards. 20 seconds.
- Show boards.

CFU CHECKPOINT:
Technique: Show Me Boards with Justification
Script:
- Say: "6 x 9. Answer AND strategy name. 20 seconds, boards up."
- Scan for: 54. Valid strategies include: "double x3" (3x9=27, dbl=54) or "x10-1 group" (10x6=60-6=54) or "x10-1 group" (10x9=90... no, that gives 9x10). Best: 6x9 via double-x3: 3x9=27, dbl=54.
PROCEED: If 80%+ show 54 with a named strategy, proceed to You Do.
PIVOT: Most likely misconception -- students write 56 (confusing 6x9 with 8x7) or 48 (confusing 6x9 with 6x8). Reteach: "6 x 9. Start with the table: x6. x6 strategy is double x3. So 3 x 9 = 27. Double 27 = 54."

TEACHER NOTES:
6 x 9 is a good hinge because students can approach it multiple ways. It tests whether they can identify and execute a strategy for a mixed-table problem.

WATCH FOR:
- Students who get 54 but cannot name a strategy -- they may be recalling from rote rather than using the toolkit
- Students who name the strategy but get the wrong answer -- execution error, not strategy error

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Final independent practice for the unit
- Your worksheet has problems from ALL tables: x2 through x10
- First: look at the table. Next: choose your strategy. Then: solve and write the strategy.
- 10 minutes. Show what you have learned this week.

DO:
- Distribute Session 4 Worksheet
- Enabling students get the Session 4 Enabling Scaffold (strategy reference card)
- Circulate, conference with students who have shown growth

TEACHER NOTES:
Culminating You Do covers all tables. Problems are deliberately mixed and include two-digit multipliers to test transfer. Strategy naming remains mandatory to reinforce metacognition.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the strategy toolkit reference card. For each problem, look up the table, find the strategy, and follow the steps. The card shows a worked example for each table.
- Extra Notes: Distribute Session 4 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Investigate multiplication patterns on a hundred chart. Colour the multiples of 3, 4, 6, and 9 in different colours. Where do they overlap? What patterns emerge?
- Extra Notes: Distribute Session 4 Extension PDF. Self-contained investigation with a printable hundred chart.

WATCH FOR:
- Students who can rapidly solve across all tables with named strategies -- they have achieved the unit goal
- Students who are strong on anchor facts but struggle with derived facts -- note for future fluency focus

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Final exit ticket. Five questions covering the whole unit.
- On your own. 4 minutes.

DO:
- Display exit ticket. Students work in workbooks.
- This is the summative assessment for the unit. Sort into: secure, developing, beginning across all tables.

TEACHER NOTES:
Q1 tests x9 (today). Q2 tests x6. Q3 tests x8. Q4 tests x3/x4 (Session 2). Q5 requires strategy explanation. Results inform future fluency planning.

WATCH FOR:
- Overall performance compared to Session 1 exit ticket -- growth data for reporting

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Final success criteria check for the unit.
- SC1: I can use x10 minus one group for x9 facts. Thumbs?
- SC2: I can choose the best strategy for any fact from x2 to x10. Thumbs?
- SC3: I can explain why all the strategies connect back to doubling and x10. Thumbs?
- Turn and talk: At the start of the week, how did you feel about times tables? How do you feel now? What changed?

DO:
- Display SCs. Thumbs for each.
- Turn and Talk 45 seconds -- this is the unit reflection, give it a bit more time
- Cold call 3-4 students for reflections
- Celebrate progress: "You went from skip counting everything to having a strategy for every table"

TEACHER NOTES:
Unit closing. The reflection question prompts metacognition about learning growth, not just content knowledge. This is powerful for student self-efficacy and builds a positive relationship with multiplication.

WATCH FOR:
- Students who express increased confidence -- this is the most important outcome
- Students who identify specific tables they still need to work on -- excellent self-assessment

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- All resources for Session 4 are linked here

DO:
- Point out each resource

TEACHER NOTES:
Resource slide for Session 4 companion PDFs.

WATCH FOR:
- N/A

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ───────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, "Multiplication Facts", "Session 4: The Nines & Full Consolidation", "Grade 3/4 Numeracy | Session 4 of 4", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2-3: Daily Review (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      T.addTitle(s, "Addition Algorithm Review", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const problems = [
        "1.  4 297 + 6 835 =",
        "2.  8 569 + 3 754 =",
        "3.  42 876 + 19 458 =",
        "4.  67 394 + 28 607 =",
      ];
      s.addText(problems.map((p, i) => ({
        text: p,
        options: { fontSize: 15, color: C.CHARCOAL, breakLine: i < problems.length - 1, paraSpaceAfter: 12 },
      })), {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboards:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "Vertical algorithm. Rename where needed.", options: { fontSize: 14, color: C.CHARCOAL } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: 1.0,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Answers:  1) 11 132    2) 12 323    3) 62 334    4) 96 001", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency
  const sFluency = pres.addSlide();
  addTopBar(sFluency, STAGE_COLORS["1"]);
  addStageBadge(sFluency, 1, "Fluency");
  T.addTitle(sFluency, "Subtraction Algorithm Practice", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

  addCard(sFluency, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  const subProblems = [
    "1.  758 - 369 =",
    "2.  1 000 - 456 =",
    "3.  3 207 - 1 489 =",
    "4.  6 003 - 2 778 =",
    "5.  8 400 - 5 621 =",
    "6.  9 105 - 6 789 =",
  ];
  sFluency.addText(subProblems.map((p, i) => ({
    text: p,
    options: { fontSize: 14, color: C.CHARCOAL, breakLine: i < subProblems.length - 1, paraSpaceAfter: 8 },
  })), {
    x: 0.75, y: CONTENT_TOP + 0.12, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.24,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addCard(sFluency, 5.2, CONTENT_TOP, 4.3, 1.6, { strip: C.SECONDARY });
  sFluency.addText([
    { text: "Whiteboards - 2 minutes", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
    { text: "Last fluency for this unit. Show what you can do!", options: { fontSize: 13, color: C.ALERT, bold: true } },
  ], {
    x: 5.45, y: CONTENT_TOP + 0.12, w: 3.8, h: 1.36,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    ["We are learning to use the x10 minus one group strategy for x9 facts and to choose the best strategy for any multiplication fact"],
    [
      "I can use x10 minus one group to work out x9 facts",
      "I can choose the best strategy for any fact from x2 to x10",
      "I can explain why all the strategies connect back to doubling and x10",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — x9 strategy
  workedExSlide(pres, 2, "I Do", "x9 Facts: x10 Minus One Group",
    [
      "9 is one less than 10",
      "So x9 = x10 minus one group",
      "",
      "Example: 9 x 7",
      "  Step 1: 10 x 7 = 70  (put a zero -- easy!)",
      "  Step 2: 70 - 7 = 63  (take away one group of 7)",
      "  Answer: 9 x 7 = 63",
    ],
    NOTES_IDO_X9, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.6, { strip: C.PRIMARY });
      slide.addText("x10 - One Group", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      // Visual: 10 rows of 7, with last row crossed out
      const dotR = 0.11;
      const dotGap = 0.19;
      const arrX = lg.rightX + 0.4;
      const arrY = lg.panelTopPadded + 0.5;
      for (let row = 0; row < 10; row++) {
        for (let col = 0; col < 7; col++) {
          slide.addShape("roundRect", {
            x: arrX + col * dotGap, y: arrY + row * dotGap,
            w: dotR, h: dotR, rectRadius: dotR / 2,
            fill: { color: row < 9 ? C.PRIMARY : C.MUTED },
          });
        }
      }
      // Cross out last row
      slide.addShape("line", {
        x: arrX - 0.05, y: arrY + 9 * dotGap + dotR / 2,
        w: 7 * dotGap, h: 0,
        line: { color: C.ALERT, width: 2 },
      });
      slide.addText("10 x 7 = 70", {
        x: lg.rightX + 0.15, y: arrY + 2.0, w: 1.8, h: 0.2,
        fontSize: 10, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
      });
      slide.addText("- 1 x 7 = 7", {
        x: lg.rightX + 0.15, y: arrY + 2.2, w: 1.8, h: 0.2,
        fontSize: 10, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
      });
    }
  );

  // Slide 7: I Do — x9 digit pattern
  contentSlide(pres, "Pattern", C.SECONDARY, "x9 Digit Pattern — Check Your Answers",
    [
      "All x9 answers have digits that add to 9:",
      "9 x 2 = 18   (1 + 8 = 9)",
      "9 x 3 = 27   (2 + 7 = 9)",
      "9 x 4 = 36   (3 + 6 = 9)",
      "9 x 5 = 45   (4 + 5 = 9)",
      "9 x 6 = 54   (5 + 4 = 9)",
      "9 x 7 = 63   (6 + 3 = 9)",
      "Use this to CHECK your x9 answers!",
    ],
    NOTES_IDO_PATTERN, FOOTER);

  // Slide 8-9: CFU 1 (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check", "Think-Pair-Share",
      "9 x 8 = ?\n\nThink... then share with your partner.\nUse the x10 minus one group strategy.",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "9 x 8 = 72   (10 x 8 = 80, minus 8 = 72)   Check: 7 + 2 = 9", {
        x: 0.5, y: 4.0, w: 9, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10: Strategy Toolkit Consolidation
  contentSlide(pres, "Toolkit", C.PRIMARY, "Your Complete Strategy Toolkit",
    [
      "x2:  Double it",
      "x3:  x2 + one more group",
      "x4:  Double-double",
      "x5:  Skip count (ends in 0 or 5)",
      "x6:  Double your x3",
      "x8:  Double your x4",
      "x9:  x10 minus one group",
      "x10: Put a zero on the end",
    ],
    NOTES_CONSOLIDATION, FOOTER);

  // Slide 11-12: We Do (withReveal)
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Mixed Table Challenge",
      [
        "With your partner, solve AND name the strategy:",
        "",
        "1.   9 x 6 = ___     Strategy: ___",
        "2.   8 x 7 = ___     Strategy: ___",
        "3.   6 x 4 = ___     Strategy: ___",
        "4.   3 x 12 = ___    Strategy: ___",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.SECONDARY });
        slide.addText("Which strategy?", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        const strats = [
          { table: "x9", strat: "x10 - 1 group", c: C.ALERT },
          { table: "x8", strat: "double x4", c: C.ACCENT },
          { table: "x6", strat: "double x3", c: C.PRIMARY },
          { table: "x3", strat: "x2 + 1 group", c: C.SECONDARY },
        ];
        strats.forEach((s, i) => {
          const sy = lg.panelTopPadded + 0.5 + i * 0.45;
          addTextOnShape(slide, s.table, {
            x: lg.rightX + 0.2, y: sy, w: 0.6, h: 0.32, rectRadius: 0.05,
            fill: { color: s.c },
          }, { fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true });
          slide.addText(s.strat, {
            x: lg.rightX + 0.9, y: sy, w: lg.rightW - 1.1, h: 0.32,
            fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "middle",
          });
        });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "1) 54 (x10-1grp)  2) 56 (dbl x4)  3) 24 (dbl x3)  4) 36 (x2+1grp)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 13-14: CFU 2 Hinge (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Final Hinge Check", "Show Me Boards",
      "6 x 9 = ___\n\nWrite the answer AND the strategy name on your whiteboard.",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "6 x 9 = 54   (double x3: 3x9=27, double 27=54)", {
        x: 1.0, y: 4.0, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 15: You Do
  workedExSlide(pres, 4, "You Do", "You Do — All Tables Mixed",
    [
      "First: Look at the table (x2? x6? x9?).",
      "Next: Choose the strategy from your toolkit.",
      "Then: Solve and write the strategy name.",
      "",
      "Problems from ALL tables: x2 through x10.",
      "10 minutes. Show everything you have learned!",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.ALERT });
      slide.addText("Strategy Toolkit", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.3, h: 0.25,
        fontSize: 13, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      const toolkit = [
        "x2 = double",
        "x3 = x2 + 1 group",
        "x4 = double-double",
        "x5 = skip count",
        "x6 = double x3",
        "x8 = double x4",
        "x9 = x10 - 1 group",
        "x10 = put a zero",
      ];
      slide.addText(toolkit.map((t, i) => ({
        text: t,
        options: { fontSize: 11, color: C.CHARCOAL, breakLine: i < toolkit.length - 1, paraSpaceAfter: 2 },
      })), {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.35, w: lg.rightW - 0.4, h: 1.9,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Slide 16: Exit Ticket
  exitTicketSlide(pres,
    [
      "9 x 6 = ___",
      "6 x 7 = ___",
      "8 x 9 = ___",
      "4 x 8 = ___",
      "Pick one fact above. Explain the strategy you used and why it works.",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 17: Closing
  closingSlide(pres,
    "At the start of the week, how did you feel about times tables? How do you feel now? What changed?",
    [
      "I can use x10 minus one group for x9 facts",
      "I can choose the best strategy for any fact x2 to x10",
      "I can explain why all strategies connect to doubling and x10",
    ],
    NOTES_CLOSING);

  // Slide 18: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "MF_Session4_x9_Consolidation.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ──────────────────────────────────────────────────────────────────

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "All Multiplication Tables — x2 to x10",
      color: C.NAVY, lessonInfo: "Session 4 of 4 | Grade 3/4 Numeracy",
    });
    y = addTipBox(doc, "For each problem: write the answer AND the strategy you used.", y, { color: C.TEAL });
    y = addSectionHeading(doc, "x9 Facts (x10 - One Group)", y, { color: C.NAVY });
    y = addProblem(doc, 1, "9 x 4 = ___    Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 2, "9 x 7 = ___    Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 3, "9 x 11 = ___   Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 4, "9 x 12 = ___   Strategy: ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "Mixed Tables", y, { color: C.NAVY });
    y = addProblem(doc, 5, "6 x 7 = ___    Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 6, "8 x 9 = ___    Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 7, "4 x 8 = ___    Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 8, "3 x 11 = ___   Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 9, "5 x 12 = ___   Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 10, "10 x 9 = ___  Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 11, "2 x 15 = ___  Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 12, "9 x 8 = ___   Strategy: ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addProblem(doc, 13, "Find THREE different strategies for 6 x 8. Show all three.", y, { color: C.NAVY, writeLines: [{ label: "" }, { label: "" }, { label: "" }] });
    addPdfFooter(doc, "Session 4 | Multiplication Facts | Grade 3/4");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "All Multiplication Tables", color: C.NAVY,
      lessonInfo: "Session 4 of 4 | Grade 3/4 Numeracy",
    });
    y = addSectionHeading(doc, "x9", y, { color: C.NAVY });
    y = addBodyText(doc, "1. 9x4=36 (10x4=40, -4=36)    2. 9x7=63 (10x7=70, -7=63)", y);
    y = addBodyText(doc, "3. 9x11=99 (10x11=110, -11=99)  4. 9x12=108 (10x12=120, -12=108)", y);
    y = addSectionHeading(doc, "Mixed", y, { color: C.NAVY });
    y = addBodyText(doc, "5. 6x7=42 (dbl x3: 3x7=21, dbl=42)", y);
    y = addBodyText(doc, "6. 8x9=72 (dbl x4: 4x9=36, dbl=72)", y);
    y = addBodyText(doc, "7. 4x8=32 (dbl-dbl: dbl 8=16, dbl 16=32)", y);
    y = addBodyText(doc, "8. 3x11=33 (x2+1grp: 2x11=22, +11=33)", y);
    y = addBodyText(doc, "9. 5x12=60 (skip count or 5x10+5x2=50+10=60)", y);
    y = addBodyText(doc, "10. 10x9=90 (put a zero)", y);
    y = addBodyText(doc, "11. 2x15=30 (double 15)", y);
    y = addBodyText(doc, "12. 9x8=72 (10x8=80, -8=72)", y);
    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "13. Three ways to get 6x8=48:", y);
    y = addBodyText(doc, "  a) Double x3: 3x8=24, dbl=48", y);
    y = addBodyText(doc, "  b) Double x4 (as 8x6): 4x6=24, dbl=48", y);
    y = addBodyText(doc, "  c) x10-1grp (as 6x8): 6x10=60, -6x2=12... or skip count", y);
    addPdfFooter(doc, "Session 4 | Multiplication Facts | Grade 3/4");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold — Strategy Toolkit Reference Card
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Strategy Toolkit Reference Card", color: C.TEAL,
      lessonInfo: "Session 4 of 4 | Grade 3/4 Numeracy",
    });
    y = addTipBox(doc, "Keep this card beside you. For each problem, find the table and follow the strategy.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "x2: Double", y, { color: C.NAVY });
    y = addBodyText(doc, "Example: 2 x 7 = double 7 = 14", y);

    y = addSectionHeading(doc, "x3: x2 + One More Group", y, { color: C.NAVY });
    y = addBodyText(doc, "Example: 3 x 6. Step 1: 2x6=12. Step 2: 12+6=18", y);

    y = addSectionHeading(doc, "x4: Double-Double", y, { color: C.NAVY });
    y = addBodyText(doc, "Example: 4 x 7. Step 1: double 7=14. Step 2: double 14=28", y);

    y = addSectionHeading(doc, "x5: Skip Count (ends in 0 or 5)", y, { color: C.NAVY });
    y = addBodyText(doc, "Example: 5 x 8 = 40 (count: 5, 10, 15, 20, 25, 30, 35, 40)", y);

    y = addSectionHeading(doc, "x6: Double Your x3", y, { color: C.NAVY });
    y = addBodyText(doc, "Example: 6 x 7. Step 1: 3x7=21. Step 2: double 21=42", y);

    y = addSectionHeading(doc, "x8: Double Your x4", y, { color: C.NAVY });
    y = addBodyText(doc, "Example: 8 x 6. Step 1: 4x6=24. Step 2: double 24=48", y);

    y = addSectionHeading(doc, "x9: x10 Minus One Group", y, { color: C.NAVY });
    y = addBodyText(doc, "Example: 9 x 7. Step 1: 10x7=70. Step 2: 70-7=63", y);
    y = addBodyText(doc, "Check: digits add to 9 (6+3=9)", y);

    y = addSectionHeading(doc, "x10: Put a Zero", y, { color: C.NAVY });
    y = addBodyText(doc, "Example: 10 x 6 = 60", y);

    addPdfFooter(doc, "Session 4 | Enabling Scaffold | Grade 3/4");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Extension — Hundred Chart Patterns
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Multiplication Patterns on the Hundred Chart", color: C.NAVY,
      lessonInfo: "Session 4 of 4 | Grade 3/4 Numeracy",
    });

    y = addSectionHeading(doc, "Your Investigation", y, { color: C.NAVY });
    y = addBodyText(doc, "Use the hundred chart below to investigate multiplication patterns.", y);
    y = addBodyText(doc, "1. Circle all multiples of 3 in BLUE.", y);
    y = addBodyText(doc, "2. Put a triangle around all multiples of 4 in RED.", y);
    y = addBodyText(doc, "3. Underline all multiples of 6 in GREEN.", y);
    y = addBodyText(doc, "4. Put a star next to all multiples of 9 in PURPLE.", y);

    // Draw a hundred chart
    y += 5;
    const cellW = 38;
    const cellH = 20;
    const chartX = 50;
    for (let row = 0; row < 10; row++) {
      for (let col = 0; col < 10; col++) {
        const num = row * 10 + col + 1;
        const cx = chartX + col * cellW;
        doc.rect(cx, y + row * cellH, cellW, cellH).stroke("#999999");
        doc.fillColor("#333333").fontSize(9).text(String(num), cx + 2, y + row * cellH + 5, { width: cellW - 4, align: "center" });
      }
    }
    y += 10 * cellH + 15;

    y = addSectionHeading(doc, "Questions", y, { color: C.NAVY });
    y = addBodyText(doc, "a) Which numbers are multiples of BOTH 3 and 4? What table do they belong to?", y);
    y = addWriteLine(doc, "Answer:", y);
    y = addBodyText(doc, "b) Which numbers are multiples of BOTH 3 and 6? Why does this happen?", y);
    y = addWriteLine(doc, "Answer:", y);
    y = addBodyText(doc, "c) What pattern do the multiples of 9 make on the chart? Can you explain why?", y);
    y = addWriteLine(doc, "Answer:", y);

    addPdfFooter(doc, "Session 4 | Extension | Grade 3/4");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 4 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
