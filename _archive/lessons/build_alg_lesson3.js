"use strict";

// Algebra Unit — Session 3: Distributive Property
// Week 1 Session 3, Grade 5/6 Numeracy, Variant 0
// DR: Number Patterns with Factors and Multiples
// Fluency: Mixed multiplication/division speed drill
// VC2M5A02 — using materials, diagrams or arrays to recognise the distributive property

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", 0);
const {
  C, FONT_H, FONT_B, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle,
  withReveal, STAGE_COLORS,
} = T;

const SESSION = 3;
const FOOTER = "Algebra | Session 3 of 10 | Grade 5/6 Numeracy";
const OUT_DIR = "output/ALG_Session3_Distributive_Property";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Independent practice - distributive property with area models.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with worked answers.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Pre-drawn area model templates for distributive property.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Self-contained investigation into using distributive property for mental maths.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ───────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Session 3 of our algebra unit
- So far we have the commutative and associative properties in our toolkit
- Today we add a third: the distributive property -- this one is a game changer for mental maths

DO:
- Display title slide
- Have whiteboards ready

TEACHER NOTES:
Session 3 of 10. The distributive property (a x (b + c) = a x b + a x c) is the most practically useful property for students. It underpins mental multiplication strategies and is essential for algebraic manipulation in secondary school.

WATCH FOR:
- Students who remember commutative and associative from Session 2 -- build on this foundation

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Daily review. Today we are working with factors and multiples
- Find all the factor pairs for each number
- Write them on your whiteboard

DO:
- Display the problems
- Allow 2 minutes, circulate
- Watch for systematic listing vs random guessing

TEACHER NOTES:
Daily Review focuses on factors and multiples, which directly supports today's distributive property work. Students who can identify factors fluently will find splitting numbers easier.

WATCH FOR:
- Students who list factors randomly rather than systematically (1 x n, 2 x ?, 3 x ?, etc.)
- Students who miss factor pairs -- they may not be checking systematically

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your factor pairs -- tick the ones you found, add any you missed
- Ask: How many factor pairs does 24 have? [4 pairs: 1x24, 2x12, 3x8, 4x6]
- Numbers with lots of factor pairs are very useful in maths

DO:
- Reveal answers
- Students self-check
- Highlight that knowing factors helps with the distributive property

TEACHER NOTES:
Connecting factor fluency to today's distributive property work. Students who know factors well will find it easier to split numbers strategically.

WATCH FOR:
- Students who found all pairs for 24 and 36 -- strong factor knowledge
- Students who missed several pairs -- factor fluency needs continued work

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time -- mixed facts, multiplication and division
- Write each answer as fast as you can
- Go!

DO:
- Display 10 mixed facts
- Time 2 minutes
- Self-check against answers

TEACHER NOTES:
Mixed multiplication and division fluency. Both operations feature in distributive property work. Students need fast recall to focus on the property rather than calculation.

WATCH FOR:
- Students who are fast on multiplication but slow on division -- division may be the bottleneck
- Students who are automatic on both -- fluency is strong

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me
- Read from slide: "We are learning to use the distributive property to break apart multiplication into easier parts"
- Distributive means you can distribute -- or spread -- the multiplication across an addition
- Read our success criteria together

DO:
- Choral read LI and SC
- Quick example: "What is 4 x 13? Hard to do in your head. But what if I told you there is a trick to make it easy?"

TEACHER NOTES:
The hook is practical: the distributive property makes hard multiplication easy. Students will see that 4 x 13 = 4 x 10 + 4 x 3 = 40 + 12 = 52 -- much easier than trying to multiply 4 x 13 directly. This motivates the property.

WATCH FOR:
- Students who already use this mental strategy intuitively -- name it: "You are already using the distributive property!"

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. I want to work out 4 x 13
- 13 is not a nice number to multiply. But I can split it into 10 + 3
- So 4 x 13 = 4 x (10 + 3)
- Now I distribute the 4 across both parts: 4 x 10 + 4 x 3
- 4 x 10 = 40. 4 x 3 = 12. Add them: 40 + 12 = 52
- I just turned a hard problem into two easy ones!
- Look at the area model on screen. The rectangle is 4 by 13. I split the 13 into 10 and 3. Two smaller rectangles. Same total area.

DO:
- Write each step clearly
- Point to the area model as you explain
- Emphasise: "splitting the number, then multiplying each part"

TEACHER NOTES:
The area model makes the distributive property visual: a 4 x 13 rectangle splits into 4 x 10 and 4 x 3 rectangles. The total area stays the same. This is the key visual anchor.

MISCONCEPTIONS:
- Misconception: You can only split the second number, not the first
  Why: All examples shown in this format (a x (b + c)) suggest the first number stays whole
  Impact: Minor -- either number can be split, but splitting the larger or more awkward number is strategically better
  Quick correction: "You could split either number. 4 x 13 = (2 + 2) x 13 = 2 x 13 + 2 x 13 works too. But splitting 13 into 10 + 3 gives easier parts."

WATCH FOR:
- Students who recognise this as a mental maths strategy they already use
- Students who struggle with the area model -- they may need to see it physically with grid paper

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Another example. 7 x 15.
- I will split 15 into 10 + 5
- 7 x 15 = 7 x (10 + 5) = 7 x 10 + 7 x 5 = 70 + 35 = 105
- Now let me try a different split: 15 = 8 + 7
- 7 x 15 = 7 x (8 + 7) = 7 x 8 + 7 x 7 = 56 + 49 = 105
- Same answer! You can split any way you like -- the distributive property always works
- But notice: splitting into 10 + 5 was easier. Choose your split wisely!

DO:
- Show both splits side by side
- Circle the fact that both give 105
- Emphasise the strategic choice of split

TEACHER NOTES:
Showing two different valid splits for the same problem demonstrates that the property is flexible. But the pedagogical message is clear: split strategically (usually around 10s) for the easiest calculation. This builds number sense alongside algebraic understanding.

WATCH FOR:
- Students who see that splitting by 10 is strategically better -- good number sense
- Students who want to try their own splits -- encourage this exploration

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. On your whiteboards, use the distributive property to work out 6 x 14
- Split 14 into parts. Show your working. 15 seconds.

DO:
- Students write on whiteboards
- Scan for: 6 x (10 + 4) = 60 + 24 = 84

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Use the distributive property. Split 14 into easier parts. Show me 6 x 14 with your working."
- Scan for: 84, with working showing 6 x 10 + 6 x 4.
PROCEED: If 80%+ show correct working with 84, move to We Do.
PIVOT: Most likely misconception -- students split 14 into 1 and 4 (treating digits separately) giving 6 x 1 + 6 x 4 = 6 + 24 = 30. Reteach: "14 is not 1 and 4. It is 10 and 4. The 1 in 14 represents 10. So we split 14 into 10 + 4, not 1 + 4."

TEACHER NOTES:
First CFU tests the core skill: splitting a two-digit number and distributing. The place-value error (splitting digits instead of values) is the most common misconception.

WATCH FOR:
- Students who split 14 into 7 + 7 instead of 10 + 4 -- valid but less efficient
- Students who get 84 but cannot show the distributive steps

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- With your partner, use the distributive property for each problem
- Split the number, multiply each part, add the results
- Show all your working

DO:
- Display four problems
- Partners work on whiteboards
- Circulate and listen for strategic splitting
- Allow 2 minutes

TEACHER NOTES:
We Do problems increase in complexity. Problems 1-2 have obvious 10-based splits. Problems 3-4 require more strategic thinking about how to split.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 3 Enabling Scaffold with pre-drawn area model grids. Each grid shows the total rectangle already split into two parts. Students count the squares in each part and write the distributive equation.
- Extra Notes: Distribute the Session 3 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Use the distributive property as a mental maths power tool. Calculate 8 x 25, 12 x 15, 7 x 99, and 6 x 48 using the distributive property. Then create your own "impossible-looking" multiplication that becomes easy with the right split.
- Extra Notes: Distribute the Session 3 Extension PDF.

WATCH FOR:
- Students who always split at 10 -- valid, but encourage flexibility
- Readiness signal: partners explaining their split choice to each other

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check. 3 x 16: split 16 into 10 + 6. 3 x 10 = 30, 3 x 6 = 18. 30 + 18 = 48.
- 5 x 23: split 23 into 20 + 3. 5 x 20 = 100, 5 x 3 = 15. 100 + 15 = 115.
- 8 x 12: split 12 into 10 + 2. 8 x 10 = 80, 8 x 2 = 16. 80 + 16 = 96.
- 4 x 17: split 17 into 10 + 7. 4 x 10 = 40, 4 x 7 = 28. 40 + 28 = 68.

DO:
- Reveal one at a time
- Cold call for each
- Ask: "Did anyone use a different split?" Validate alternatives

TEACHER NOTES:
Validating alternative splits reinforces that the property works regardless of how you split. The strategic choice is about ease of calculation, not correctness.

WATCH FOR:
- Students who found alternative splits and got the same answer -- celebrate this
- Students who made arithmetic errors in the partial products -- the strategy is correct, the calculation needs practice

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. Finger vote.
- Which split is BEST for 9 x 18?
- 1 finger: 9 x (10 + 8)
- 2 fingers: 9 x (9 + 9)
- 3 fingers: 9 x (20 - 2)

DO:
- Students show fingers
- All three are valid! But scan for reasoning
- Follow up: "They all work. Which gives the easiest calculation?"

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "Three possible splits for 9 x 18. Which is BEST? 1 = split into 10+8, 2 = split into 9+9, 3 = split into 20-2. Show me."
- Scan for: Accept any (all valid). 1 finger (10+8) gives 90+72=162. 3 fingers (20-2) gives 180-18=162 (subtraction variant). All correct.
PROCEED: If 80%+ select a valid option and can justify, move to You Do.
PIVOT: Most likely misconception -- students think only one split is "correct." Reteach: "All three splits work -- watch: 9x10+9x8=90+72=162. 9x9+9x9=81+81=162. 9x20-9x2=180-18=162. The distributive property does not care how you split -- it always gives the same answer."

TEACHER NOTES:
This hinge introduces the subtraction variant (9 x (20-2)) which previews compensation strategies. All options are valid -- the pedagogical point is strategic flexibility.

WATCH FOR:
- Students who only picked one option -- they may not yet see that multiple splits work
- Students who chose option 3 -- they may have strong number sense

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice
- Your worksheet has problems where you use the distributive property
- First: Choose your split. Next: Multiply each part. Then: Add the results.
- Show all working. 8 minutes.

DO:
- Distribute Session 3 Worksheet
- Students work independently
- Circulate: support students with area model scaffold if needed
- Distribute enabling/extending resources as appropriate

TEACHER NOTES:
You Do uses different numbers from We Do. Section B requires students to identify the distributive property in given equations and Section C asks students to write their own distributive splits.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the pre-drawn area model grids. Count squares to find each partial product, then add.
- Extra Notes: Distribute Session 3 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Mental maths power tool investigation. Use distributive property for challenging multiplications and create "impossible-looking" problems.
- Extra Notes: Distribute Session 3 Extension PDF.

WATCH FOR:
- Students who can split and calculate but do not show distributive notation -- redirect to write a x (b + c) = a x b + a x c
- Readiness signal: 80%+ completing Section A correctly within 5 minutes

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Three questions. 3 minutes. Work alone.

DO:
- Display questions
- Students work silently
- Collect to sort

TEACHER NOTES:
Q1 tests basic application (SC1). Q2 tests recognising the property in an equation (SC2). Q3 tests writing a distributive equation for a given area model (SC3).

WATCH FOR:
- Students who calculate correctly but do not show distributive steps
- Students who get all three instantly -- ready for unknown-finding in Session 4

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Success criteria check.
- SC1: I can split a number and use the distributive property to multiply. Thumbs?
- SC2: I can explain why the distributive property works using an area model. Thumbs?
- SC3: I can choose the best split for a given multiplication. Thumbs?
- Turn and talk: How would you use the distributive property to work out 7 x 99? Tell your partner.

DO:
- Run thumbs for each SC
- 30 seconds Turn and Talk
- Cold call -- listen for 7 x (100-1) = 700-7 = 693

TEACHER NOTES:
The 7 x 99 prompt naturally leads to the subtraction variant (split as 100-1). Students who see this have strong number sense and algebraic flexibility.

WATCH FOR:
- Students who try 7 x (90+9) -- valid but harder. Guide toward 100-1
- Students who use the property confidently -- ready for Session 4

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Resources for today are linked here

DO:
- Point out each resource

TEACHER NOTES:
Resource slide with clickable links to companion PDFs.

WATCH FOR:
- N/A

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ───────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, "Algebra: Finding Unknown Values", "Session 3: The Distributive Property", "Grade 5/6 Numeracy | Session 3 of 10 | Week 1", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2-3: Daily Review (withReveal) — Number Patterns with Factors and Multiples
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Factors and Multiples", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const problems = [
        "1.  List all factor pairs of 24",
        "2.  List all factor pairs of 36",
        "3.  What are the first 5 multiples of 7?",
        "4.  Find a number that is a multiple of both 4 and 6",
      ];
      s.addText(problems.map((p, i) => ({
        text: p,
        options: { fontSize: 14, color: C.CHARCOAL, breakLine: i < problems.length - 1, paraSpaceAfter: 14 },
      })), {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboards:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "List factor pairs systematically.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "Start: 1 x ?, 2 x ?, 3 x ?, ...", options: { fontSize: 14, color: C.CHARCOAL } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) 1x24,2x12,3x8,4x6  2) 1x36,2x18,3x12,4x9,6x6  3) 7,14,21,28,35  4) 12,24...", {
        x: 0.3, y: 4.55, w: 9.4, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 11.5, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency
  const sFl = pres.addSlide();
  addTopBar(sFl, STAGE_COLORS["1"]);
  addStageBadge(sFl, 1, "Fluency");
  addTitle(sFl, "Mixed Facts Speed Round", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

  addCard(sFl, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  const facts = [
    "1.  8 x 7 =",     "2.  54 / 9 =",
    "3.  12 x 6 =",    "4.  48 / 8 =",
    "5.  9 x 9 =",     "6.  72 / 12 =",
    "7.  7 x 6 =",     "8.  63 / 7 =",
    "9.  11 x 8 =",    "10. 96 / 8 =",
  ];
  sFl.addText(facts.map((p, i) => ({
    text: p,
    options: { fontSize: 14, color: C.CHARCOAL, breakLine: i < facts.length - 1, paraSpaceAfter: 4 },
  })), {
    x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addCard(sFl, 5.2, CONTENT_TOP, 4.3, 1.6, { strip: C.SECONDARY });
  sFl.addText([
    { text: "Whiteboards - 2 minutes", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
    { text: "Speed and accuracy!", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Answers: 56, 6, 72, 6, 81, 6, 42, 9, 88, 12", options: { fontSize: 11, color: C.MUTED } },
  ], {
    x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: 1.4,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addFooter(sFl, FOOTER);
  sFl.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    ["We are learning to use the distributive property to break apart multiplication into easier parts"],
    [
      "I can split a number and use the distributive property to calculate a product",
      "I can explain why the distributive property works using an area model",
      "I can choose the best way to split a number for easier calculation",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — Distributive Property with Area Model
  workedExSlide(pres, 2, "I Do", "Distributive Property: Break It Apart",
    [
      "4 x 13 is hard to do in your head",
      "Split 13 into 10 + 3:",
      "",
      "  4 x 13 = 4 x (10 + 3)",
      "         = 4 x 10 + 4 x 3",
      "         = 40 + 12",
      "         = 52",
      "",
      "The distributive property:",
      "  a x (b + c) = a x b + a x c",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.PRIMARY });
      slide.addText("Area Model: 4 x 13", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.3, h: 0.28,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      // Area model: two adjacent rectangles
      const amX = lg.rightX + 0.35;
      const amY = lg.panelTopPadded + 0.7;
      const amH = 1.1;
      const am10W = 2.0;
      const am3W = 0.8;

      // Labels above the rectangles
      slide.addText("10", {
        x: amX, y: amY - 0.28, w: am10W, h: 0.24,
        fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, align: "center", margin: 0, bold: true,
      });
      slide.addText("3", {
        x: amX + am10W, y: amY - 0.28, w: am3W, h: 0.24,
        fontSize: 12, fontFace: FONT_B, color: C.SECONDARY, align: "center", margin: 0, bold: true,
      });

      // 4 label to the left
      slide.addText("4", {
        x: amX - 0.3, y: amY, w: 0.25, h: amH,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0, bold: true,
      });

      // 4 x 10 rectangle
      slide.addShape("rect", {
        x: amX, y: amY, w: am10W, h: amH,
        fill: { color: C.PRIMARY }, line: { color: C.CHARCOAL, width: 1 },
      });
      slide.addText("4 x 10 = 40", {
        x: amX, y: amY, w: am10W, h: amH,
        fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // 4 x 3 rectangle
      slide.addShape("rect", {
        x: amX + am10W, y: amY, w: am3W, h: amH,
        fill: { color: C.SECONDARY }, line: { color: C.CHARCOAL, width: 1 },
      });
      slide.addText("4 x 3\n= 12", {
        x: amX + am10W, y: amY, w: am3W, h: amH,
        fontSize: 11, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Total
      addTextOnShape(slide, "40 + 12 = 52", {
        x: lg.rightX + 0.5, y: amY + amH + 0.15, w: lg.rightW - 1.0, h: 0.4, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

      // General rule
      addTextOnShape(slide, "a x (b + c) = a x b + a x c", {
        x: lg.rightX + 0.3, y: amY + amH + 0.7, w: lg.rightW - 0.6, h: 0.4, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 7: I Do — Multiple splits
  workedExSlide(pres, 2, "I Do", "Different Splits, Same Answer",
    [
      "7 x 15: try two different splits",
      "",
      "Split A: 15 = 10 + 5",
      "  7 x (10 + 5) = 70 + 35 = 105",
      "",
      "Split B: 15 = 8 + 7",
      "  7 x (8 + 7) = 56 + 49 = 105",
      "",
      "Both give 105!",
      "But Split A was easier to calculate.",
      "Choose your split wisely.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.4, { strip: C.SUCCESS });
      slide.addText("Split A: 10 + 5", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 13, fontFace: FONT_H, color: C.SUCCESS, bold: true, margin: 0,
      });
      slide.addText("7 x 10 = 70\n7 x 5 = 35\nTotal = 105", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.35, w: lg.rightW - 0.6, h: 0.8,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });

      addCard(slide, lg.rightX, lg.panelTopPadded + 1.6, lg.rightW, 1.4, { strip: C.ACCENT });
      slide.addText("Split B: 8 + 7", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.68, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 13, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
      });
      slide.addText("7 x 8 = 56\n7 x 7 = 49\nTotal = 105", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.95, w: lg.rightW - 0.6, h: 0.8,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });
    }
  );

  // Slide 8-9: CFU 1 (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check", "Show Me Boards",
      "Use the distributive property:\n\n6 x 14 = ?\n\nSplit 14, show your working.",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "6 x (10 + 4) = 6 x 10 + 6 x 4 = 60 + 24 = 84", {
        x: 0.8, y: 4.0, w: 8.4, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10-11: We Do (withReveal)
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Your Turn — Distributive Property",
      [
        "Use the distributive property for each.",
        "Show: split, multiply parts, add.",
        "",
        "1.   3 x 16 =",
        "2.   5 x 23 =",
        "3.   8 x 12 =",
        "4.   4 x 17 =",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.2, { strip: C.SECONDARY });
        slide.addText("Method:", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.25,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        slide.addText([
          { text: "1. Split the bigger number", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "2. Multiply each part", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "3. Add the results", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
        ], {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.4, w: lg.rightW - 0.5, h: 0.8,
          fontFace: FONT_B, margin: 0, valign: "top",
        });
        addTextOnShape(slide, "a x (b + c) = a x b + a x c", {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.5, w: lg.rightW - 0.6, h: 0.35, rectRadius: 0.06,
          fill: { color: C.PRIMARY },
        }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "1) 48  2) 115  3) 96  4) 68", {
        x: 1.5, y: 4.55, w: 7, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 12-13: CFU 2 Hinge (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Check — Best Split", "Finger Voting",
      "Which split is BEST for 9 x 18?\n\n1 finger: 9 x (10 + 8)\n2 fingers: 9 x (9 + 9)\n3 fingers: 9 x (20 - 2)",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "All valid!  10+8: 90+72=162   9+9: 81+81=162   20-2: 180-18=162", {
        x: 0.5, y: 4.0, w: 9, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do
  workedExSlide(pres, 4, "You Do", "Independent Practice",
    [
      "First: Read each problem on your worksheet.",
      "Next: Choose your split and apply the distributive property.",
      "Then: Show your full working.",
      "",
      "Write the distributive equation for each.",
      "You have 8 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.8, { strip: C.ALERT });
      slide.addText("Distributive Property:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      addTextOnShape(slide, "a x (b + c) = a x b + a x c", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.5, w: lg.rightW - 0.6, h: 0.4, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addText("Split at 10 for easiest parts!", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.05, w: lg.rightW - 0.6, h: 0.3,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, align: "center",
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "Use the distributive property to calculate 9 x 14. Show your working.",
      "TRUE or FALSE: 5 x 16 = 5 x 10 + 5 x 6. Explain why.",
      "Draw an area model for 3 x 12, showing how you split 12.",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 16: Closing
  closingSlide(pres,
    "How would you use the distributive property to work out 7 x 99? Tell your partner.",
    [
      "I can split a number and use the distributive property to calculate",
      "I can explain why it works using an area model",
      "I can choose the best split for easier calculation",
    ],
    NOTES_CLOSING);

  // Slide 17: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "ALG_Session3_Distributive_Property.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Generate PDFs ─────────────────────────────────────────────────────────

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "The Distributive Property",
      color: C.NAVY, lessonInfo: "Session 3 of 10 | Grade 5/6 Numeracy",
    });
    y = addTipBox(doc, "For each problem: split the number, multiply each part, then add the results. Show the distributive equation: a x (b + c) = a x b + a x c", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A: Apply the Distributive Property", y, { color: C.NAVY });
    y = addProblem(doc, 1, "7 x 16 = 7 x (___ + ___) = ___ + ___ = ___", y, { color: C.NAVY });
    y = addProblem(doc, 2, "9 x 13 = 9 x (___ + ___) = ___ + ___ = ___", y, { color: C.NAVY });
    y = addProblem(doc, 3, "5 x 24 = 5 x (___ + ___) = ___ + ___ = ___", y, { color: C.NAVY });
    y = addProblem(doc, 4, "8 x 15 = 8 x (___ + ___) = ___ + ___ = ___", y, { color: C.NAVY });
    y = addProblem(doc, 5, "6 x 19 = 6 x (___ + ___) = ___ + ___ = ___", y, { color: C.NAVY });
    y = addProblem(doc, 6, "4 x 22 = 4 x (___ + ___) = ___ + ___ = ___", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Section B: TRUE or FALSE?", y, { color: C.NAVY });
    y = addProblem(doc, 7, "3 x 18 = 3 x 10 + 3 x 8          TRUE / FALSE", y, { color: C.NAVY });
    y = addProblem(doc, 8, "6 x 14 = 6 x 10 + 6 x 14         TRUE / FALSE", y, { color: C.NAVY });
    y = addProblem(doc, 9, "5 x 17 = 5 x 10 + 5 x 7          TRUE / FALSE", y, { color: C.NAVY });
    y = addProblem(doc, 10, "8 x 25 = 8 x 20 + 8 x 5         TRUE / FALSE", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Section C: Write Your Own", y, { color: C.NAVY });
    y = addProblem(doc, 11, "Write a distributive equation for 12 x 7:", y, { color: C.NAVY });
    y = addWriteLine(doc, "", y);
    y = addProblem(doc, 12, "Write a distributive equation for 9 x 25:", y, { color: C.NAVY });
    y = addWriteLine(doc, "", y);

    addPdfFooter(doc, "Session 3 | Algebra: Finding Unknown Values | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "The Distributive Property",
      color: C.NAVY, lessonInfo: "Session 3 of 10 | Grade 5/6 Numeracy",
    });
    y = addSectionHeading(doc, "Section A", y, { color: C.NAVY });
    y = addBodyText(doc, "1. 7 x 16 = 7 x (10+6) = 70+42 = 112", y);
    y = addBodyText(doc, "2. 9 x 13 = 9 x (10+3) = 90+27 = 117", y);
    y = addBodyText(doc, "3. 5 x 24 = 5 x (20+4) = 100+20 = 120", y);
    y = addBodyText(doc, "4. 8 x 15 = 8 x (10+5) = 80+40 = 120", y);
    y = addBodyText(doc, "5. 6 x 19 = 6 x (10+9) = 60+54 = 114", y);
    y = addBodyText(doc, "6. 4 x 22 = 4 x (20+2) = 80+8 = 88", y);
    y = addSectionHeading(doc, "Section B", y, { color: C.NAVY });
    y = addBodyText(doc, "7. TRUE (3x10+3x8 = 30+24 = 54 = 3x18)", y);
    y = addBodyText(doc, "8. FALSE (6x10+6x14 = 60+84 = 144, but 6x14 = 84)", y);
    y = addBodyText(doc, "9. TRUE (5x10+5x7 = 50+35 = 85 = 5x17)", y);
    y = addBodyText(doc, "10. TRUE (8x20+8x5 = 160+40 = 200 = 8x25)", y);
    y = addSectionHeading(doc, "Section C", y, { color: C.NAVY });
    y = addBodyText(doc, "11. 12 x 7 = 12 x (5+2) = 60+24 = 84 (or other valid splits)", y);
    y = addBodyText(doc, "12. 9 x 25 = 9 x (20+5) = 180+45 = 225 (or other valid splits)", y);
    addPdfFooter(doc, "Session 3 | Answer Key | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold — Area model templates
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Area Model Templates",
      color: C.TEAL, lessonInfo: "Session 3 of 10 | Grade 5/6 Numeracy",
    });
    y = addTipBox(doc, "Each area model shows a rectangle split into two parts. Write the multiplication for each part, then add them together.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Example: 4 x 13 = 4 x (10 + 3)", y, { color: C.NAVY });
    y = addBodyText(doc, "Part 1: 4 x 10 = 40", y);
    y = addBodyText(doc, "Part 2: 4 x 3 = 12", y);
    y = addBodyText(doc, "Total: 40 + 12 = 52", y);
    y += 10;

    y = addSectionHeading(doc, "Now you try:", y, { color: C.NAVY });
    const probs = [
      { desc: "3 x 16 = 3 x (10 + 6)", p1: "3 x 10 = ___", p2: "3 x 6 = ___" },
      { desc: "5 x 14 = 5 x (10 + 4)", p1: "5 x 10 = ___", p2: "5 x 4 = ___" },
      { desc: "7 x 12 = 7 x (10 + 2)", p1: "7 x 10 = ___", p2: "7 x 2 = ___" },
      { desc: "8 x 15 = 8 x (10 + 5)", p1: "8 x 10 = ___", p2: "8 x 5 = ___" },
    ];
    probs.forEach((p, i) => {
      y = addBodyText(doc, (i + 1) + ". " + p.desc, y);
      y = addBodyText(doc, "   Part 1: " + p.p1, y);
      y = addBodyText(doc, "   Part 2: " + p.p2, y);
      y = addBodyText(doc, "   Total: ___ + ___ = ___", y);
      y += 10;
    });

    addPdfFooter(doc, "Session 3 | Enabling Scaffold | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Extension — Mental Maths Power Tool
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Mental Maths Power Tool",
      color: C.NAVY, lessonInfo: "Session 3 of 10 | Grade 5/6 Numeracy",
    });

    y = addSectionHeading(doc, "The Distributive Property as a Mental Maths Strategy", y, { color: C.NAVY });
    y = addBodyText(doc, "The distributive property is not just a maths rule -- it is a mental maths superpower!", y);
    y = addBodyText(doc, "By choosing clever splits, you can turn hard multiplications into easy ones.", y);

    y = addSectionHeading(doc, "Worked Example: 7 x 99", y, { color: C.NAVY });
    y = addBodyText(doc, "99 is close to 100. Split it: 99 = 100 - 1", y);
    y = addBodyText(doc, "7 x 99 = 7 x (100 - 1) = 7 x 100 - 7 x 1 = 700 - 7 = 693", y);
    y = addBodyText(doc, "That is the subtraction variant of the distributive property!", y);

    y = addSectionHeading(doc, "Your Challenges", y, { color: C.NAVY });
    y = addProblem(doc, 1, "8 x 25 = 8 x (___) = ___ = ___", y, { color: C.NAVY });
    y = addProblem(doc, 2, "12 x 15 = 12 x (___) = ___ = ___", y, { color: C.NAVY });
    y = addProblem(doc, 3, "7 x 99 = 7 x (___) = ___ = ___", y, { color: C.NAVY });
    y = addProblem(doc, 4, "6 x 48 = 6 x (___) = ___ = ___", y, { color: C.NAVY });
    y = addProblem(doc, 5, "9 x 101 = 9 x (___) = ___ = ___", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Create Your Own", y, { color: C.NAVY });
    y = addBodyText(doc, "Write a multiplication that LOOKS impossible but is easy with the right split:", y);
    y = addWriteLine(doc, "The problem:", y);
    y = addWriteLine(doc, "The split:", y);
    y = addWriteLine(doc, "The answer:", y);

    y = addTipBox(doc, "Pro Tip: Numbers close to 10, 20, 50, 100, or 1000 are great candidates for the subtraction variant. 6 x 98 = 6 x (100-2) = 600-12 = 588. Try to impress your classmates!", y, { color: C.TEAL });

    addPdfFooter(doc, "Session 3 | Extension Investigation | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 3 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
