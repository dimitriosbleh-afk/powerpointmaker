"use strict";

// BODMAS Unit - Session 1: What Is BODMAS?
// Year 5/6 Numeracy, Week 9 Term 2 (variant 2)
// VC2M6A02 - find unknown values in numerical equations involving brackets
// and combinations of arithmetic operations
// Daily Review: Solving Equations with Multiplication, Division, and Operations
// Number Fluency: mixed addition, subtraction, multiplication, division

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(9));
const {
  C, FONT_H, FONT_B,
  SLIDE_W, SLIDE_H, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addBadge, addTitle,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const UNIT_TITLE = "BODMAS: Order of Operations";
const FOOTER = "BODMAS | Session 1 of 10 | Year 5/6 Numeracy";
const OUT_DIR = "output/BOD_Session1_What_Is_BODMAS";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "BODMAS Reference Sheet", "Print and stick in maths book - the BODMAS rule and the steps.");
const PRACTICE_RES = makeSessionResource(SESSION, "Practice Sheet", "Light practice applying BODMAS to simple multi-step equations.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with worked answers.");
const RESOURCE_ITEMS = [WORKSHEET_RES, PRACTICE_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today we are starting a brand new unit called BODMAS
- BODMAS is a set of rules that mathematicians all over the world use
- Some of you may have heard about it. If this feels new, that's okay - we will build it together over the next two weeks

DO:
- Display title slide as students settle
- Have whiteboards and markers ready
- Have the reference sheet ready to hand out at the end of the lesson

TEACHER NOTES:
Session 1 of 10. This is the gateway lesson - the goal is for students to know what BODMAS stands for and to understand WHY we need a rule. The viral problem hook is the key teaching moment - different students will get different answers, and that disagreement is the bridge into the rule.

WATCH FOR:
- Students who claim they "already know" BODMAS - good, they will be helpful in pair-share, but check whether they apply the M and D step correctly later

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Let's warm up. Look at the number 465,102
- Answer the questions in your book - quick responses
- Use what you know about place value and rounding

DO:
- Display the number-of-the-day questions
- Allow about 3 minutes - students respond in their books
- Circulate and check working

TEACHER NOTES:
Daily Review uses 'Number of the Day' to retrieve place value, rounding and number naming. This is prior learning, not new content. Worth doing because students will need fluent place value to handle the larger answers in the BODMAS practice ahead.

WATCH FOR:
- Students who confuse "nearest hundred" with "nearest thousand" - prompt them to underline the place they are rounding to
- Students who write the expanded form with extra zeros - check the structure

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers - tick what you got right, fix what you didn't
- Fixing it is the learning
- Ask: What place are you rounding to for the nearest thousand? [The thousands column]

DO:
- Click to reveal answers
- Students tick and fix in their books

TEACHER NOTES:
Tick-and-fix gives immediate feedback. Note any students who struggle with rounding - they will need extra support during fluency.

WATCH FOR:
- Students who got all correct - place value is solid
- Students who confused expanded form - reteach later in a small group

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time. Six equations on the board
- For each, write the equation in your book and work it out left to right
- We will check together in a moment

DO:
- Display the six fluency equations
- Allow about 4 minutes - students work in books
- Circulate, check working
- Note: these are LEFT TO RIGHT for now - we are not yet using BODMAS

TEACHER NOTES:
Fluency builds automaticity with addition, subtraction, multiplication and division. These are deliberately simple chains so students can focus on accuracy. After today's lesson, students will see these chains differently - the order matters.

WATCH FOR:
- Students who add before subtracting in the chain "15 + 8 - 3 + 2" - good for now, but mark this for the multi-step lesson on Monday
- Students who get tangled in 5 x 5 x 2 / 5 - prompt them to work strictly left to right

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Check your answers - tick or fix
- These are easier when you go step by step

DO:
- Reveal answers
- Note any students who got the multiplication chains wrong

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me
- Read from slide: "I am learning what BODMAS is and why we use it"
- Now our success criteria. Read each one together
- Ask: What do you think BODMAS could stand for? [Hands down - we will find out today]

DO:
- Choral read the LI then each SC
- Brief think-pair-share on what BODMAS might mean

TEACHER NOTES:
This is the gateway lesson. BODMAS is the threshold concept for the whole unit. SC1 says "name the letters" - achievable for everyone. SC2 says "follow the order on a simple two-step equation" - the core target. SC3 says "explain why a rule is needed" - the depth criterion.

WATCH FOR:
- Students who think there is no point to a rule because "you just go left to right" - this is the productive misconception we want to surface in the launch

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_LAUNCH = `SAY:
- Look at this equation: 6 + 2 x 3
- On your whiteboard, work it out. Don't talk yet
- Show me. [Pause - some will write 24, some will write 12]
- Different answers. Why?
- Turn to your partner and explain what YOU did first
- [After 90 seconds] Hands up if you got 24. Hands up if you got 12. Hands up if you got something else
- Both 24 and 12 look correct depending on what you did first. We need a rule

DO:
- Display the equation 6 + 2 x 3
- Wait for whiteboards before moving on
- Tally answers on the board
- Make the disagreement visible - this is the hook

TEACHER NOTES:
This is the engagement hook AND the rationale for BODMAS. The class will split. Some students will compute 6+2 first (=8), then x3 (=24). Others will compute 2x3 first (=6), then +6 (=12). The correct BODMAS answer is 12, but DO NOT reveal this yet - the disagreement IS the teaching. Lead them to "we need a rule everyone agrees on" before naming BODMAS.

MISCONCEPTIONS:
- Misconception: Maths equations are always solved left to right
  Why: Years of simple chains like 4 + 3 + 2 reinforce a left-to-right habit
  Impact: Students get wrong answers on multi-step equations and lose marks across all of upper primary maths
  Quick correction: "When operations are mixed, the order matters. We need a rule. That rule is BODMAS"

WATCH FOR:
- Students who confidently announce "12" - they have prior knowledge of BODMAS, get them to explain WHY
- Students who say "but everyone got different answers, what's the point?" - exactly, that is why we need a rule

[Stage 2: Launch | VTLM 2.0: Engagement]`;

const NOTES_BODMAS_DEFINE = `SAY:
- BODMAS is an acronym - each letter stands for a step
- B is Brackets - solve anything in brackets first
- O is Orders - that means powers, like 2 squared, or roots
- D is Division - same level as multiplication
- M is Multiplication - same level as division
- A is Addition - same level as subtraction
- S is Subtraction - same level as addition
- We work from the top of the list to the bottom
- D and M are the SAME step - work left to right
- A and S are the SAME step - work left to right

DO:
- Point to each letter as you say it
- Pause on B and O - students may not know what brackets and orders look like
- Use the colour-coded reference on screen

TEACHER NOTES:
This is the explicit definition. Students need to hear it, see it, and have a reference. The reference sheet at the end of the lesson is for sticking in their books - they will need it for every lesson in this unit. Keep this slide clean - the acronym is the hero.

WATCH FOR:
- Confusion about "Orders" - clarify: it means a number raised to a power, like 5 squared (5x5) or 2 cubed (2x2x2)
- Students who try to memorise the acronym before they understand the steps - they will get it wrong

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_VIRAL = `SAY:
- Now let's go back to our viral problem - 6 + 2 x 3
- Watch me think this through using BODMAS
- B - Brackets? No, there are no brackets
- O - Orders? No, there are no powers
- D and M - Division or Multiplication? Yes - 2 x 3
- I do that first. 2 x 3 = 6. Now my equation is 6 + 6
- A and S - Addition or Subtraction? Yes - 6 + 6 = 12
- The answer is 12, not 24. The multiplication came first because of BODMAS

DO:
- Write each step on the board as you say it
- Cross out the part you have just solved and replace it
- Use a different colour for the step you are doing
- Emphasise the order - B, O, D/M, A/S

TEACHER NOTES:
This think-aloud demonstrates the BODMAS routine students will use for the entire unit. The exact steps matter - go through B and O even though they don't apply here. Students need to see the WHOLE checklist run, not just the parts that fire. This builds the habit.

WATCH FOR:
- Students who say "but I got 24, that's wrong" - reframe: "Your maths was correct, but the order wasn't. Now you have a rule to follow"

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. On your whiteboards
- 10 - 6 / 2 - what do you do FIRST?
- Just write the operation, not the answer
- Show me

DO:
- Students write on whiteboards
- Scan for "Division" or "/" - that is correct
- Pivot if more than 20% write "Subtraction"

TEACHER NOTES:
This is a quick CFU on the order, not the answer. Students need to identify division as the first step before they can solve anything correctly. If many get this wrong, reteach with the BODMAS reference open and visible.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Write the FIRST step you would do
- Just the operation, not the whole answer
- Scan for: "Division" or "/" or "6 / 2"
PROCEED:
- 80% or more identify division - move on
PIVOT:
- Most likely misconception: students go left to right and pick subtraction first
- Reteach using the BODMAS reference: "B, O, D and M, A and S - which letters appear in this equation?"
- Re-check with: 8 + 4 x 2 - what is the first step? [Multiplication]

WATCH FOR:
- Students who pick subtraction - clear left-to-right thinking, needs the BODMAS habit
- Students who write "10" - they are answering, not identifying the step

[Stage 2: CFU | VTLM 2.0: Active Checking]`;

const NOTES_CFU1_A = `SAY:
- The first step is division: 6 / 2 = 3
- Then subtraction: 10 - 3 = 7
- Final answer: 7

DO:
- Reveal answer with full working
- Note who needs more support

[Stage 2: CFU Answer | VTLM 2.0: Active Checking]`;

const NOTES_WEDO_Q = `SAY:
- Now we will do one together
- 20 - 4 x 3 + 2
- B - any brackets? No
- O - any orders? No
- D and M - any division or multiplication? [4 x 3] - which one comes first?
- 4 x 3 is 12. Rewrite the equation: 20 - 12 + 2
- Now A and S, left to right. 20 - 12 = 8. Then 8 + 2 = 10
- Final answer: 10

DO:
- Walk through each step on the board with class input
- Ask "what comes next?" at each stage
- Pause for whiteboard answers at each step
- Use cold call after thinking time

TEACHER NOTES:
We Do is the guided practice - the class is helping you build the answer. Pause at each step and check thinking. The trap here is students doing 12 + 2 = 14 first because they want to keep going left to right. Catch that misconception.

WATCH FOR:
- Students who try to do A and S in the wrong order (e.g. 12 + 2 first) - this is the most common error
- Students who say "I would just do 20 - 4 first" - help them see that x is in a higher group

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Final answer: 10
- Tick if you got it. Fix it if you didn't

DO:
- Reveal answer
- Quick tally of correct answers

[Stage 3: We Do Answer | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge question. On your whiteboards
- 12 + 8 / 4
- Work through BODMAS in your head and write the FINAL answer
- 30 seconds. Show me

DO:
- Students compute and show whiteboards
- Scan for "14"
- 14 means correct (8/4=2, 12+2=14)
- 5 means they did 12+8 first (=20, /4 = 5) - left to right error

TEACHER NOTES:
This is a hinge - it tells you whether to proceed to You Do or pivot. The two wrong answers map to two specific misconceptions: 5 means left-to-right thinking, 20.5 means they did /4 of the wrong number. Use the result to decide.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Final answer only
- Show me
- Scan for: 14
PROCEED:
- 80% or more show 14 - go to You Do
PIVOT:
- Most likely misconception: students did 12+8 first because of left-to-right habit
- Reteach: "Look at the equation. Which letters in BODMAS appear? Division. Then Addition. Division comes BEFORE addition"
- Re-check: 6 + 10 / 2 [answer: 11]

WATCH FOR:
- Students showing 5 - left-to-right error, very common
- Students showing 14 confidently - ready for independent practice

[Stage 3: Hinge CFU | VTLM 2.0: Active Checking]`;

const NOTES_CFU2_A = `SAY:
- Final answer: 14
- 8 divided by 4 is 2
- 12 plus 2 is 14
- If you got 5, you went left to right - that's the BODMAS trap

DO:
- Reveal full working
- Acknowledge the trap

[Stage 3: Hinge Answer | VTLM 2.0: Active Checking]`;

const NOTES_YOUDO = `SAY:
- Time to fly solo
- Open your maths book
- Solve each equation using BODMAS
- Show your steps - I want to see what you did first, second, third
- If you finish, try the challenge at the bottom

DO:
- Hand out practice sheet OR write equations on board
- Allow 8 to 10 minutes
- Circulate - check students are showing steps, not jumping straight to an answer
- Pull a small group of any students who struggled with the hinge CFU

TEACHER NOTES:
You Do is independent application. The practice sheet has 5 equations that all use just two operations - we are not yet using brackets or orders. That is the next lesson. The challenge at the bottom uses three operations and is for students who finish early.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Pull a small group. Use the BODMAS reference. Together, IDENTIFY the operations in each equation BEFORE solving. Write B, O, D, M, A, S down the side and tick which appear
- Extra Notes: Once they can identify the order, let them solve the equations independently
EXTENDING PROMPT:
- Task: Challenge equation: 30 - 4 x 5 + 2 x 3
- Extra Notes: Two multiplications and two additions/subtractions. Order matters. Answer: 16

WATCH FOR:
- Students who get the answer right but don't show working - prompt: "I want to see WHICH step you did first"
- Students who skip the BODMAS check at the start - get them to write B-O-D-M-A-S down the side

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Three quick questions
- Solve each one using BODMAS
- I am collecting these to see how you went today

DO:
- Hand out exit ticket OR have students do on a sticky note
- Allow 4 minutes
- Collect at the door

TEACHER NOTES:
Exit ticket assesses SC2 - "I can use BODMAS to solve a simple two-step equation". Q1 is one-operation jump (everyone), Q2 is the core BODMAS check, Q3 is the depth question.

WATCH FOR:
- Students who get Q2 wrong - they need more practice on M before A
- Students who write "BODMAS" but don't show the steps - they are surface-learning

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Let's review our success criteria
- Read each one
- Self-check - thumbs up, sideways, or down
- Tell your partner one thing you learned about BODMAS today

DO:
- Display closing slide
- Pause for self-check
- Brief turn-and-talk

TEACHER NOTES:
Closing reviews the three success criteria and asks for self-assessment. Use the data to plan tomorrow - if many students show "needs more practice", revisit the basic order before introducing brackets.

WATCH FOR:
- Students who confidently self-assess as "got it" but bombed the exit ticket - check their evidence
- Students who say "I'm confused" - flag for tomorrow's small group

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- Here are the resources for today
- Print and stick the BODMAS reference sheet in your maths book - we will use it every lesson

DO:
- Hand out the reference sheet
- Have practice sheets available for those who want extra practice

[General: Resources | VTLM 2.0: Resource Awareness]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres,
    "BODMAS: Order of Operations",
    "What is it? Why do we need it?",
    "Session 1 of 10  |  Year 5/6 Numeracy  |  VC2M6A02",
    NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2: Daily Review
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Number of the Day: 465,102", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      // Hero number card on left
      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      addTextOnShape(s, "465,102", {
        x: 0.7, y: CONTENT_TOP + 0.6, w: 4.1, h: 1.4, rectRadius: 0.1,
        fill: { color: STAGE_COLORS["1"] },
      }, { fontSize: 48, fontFace: FONT_H, color: C.WHITE, bold: true });

      s.addText([
        { text: "Odd or even?", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "What is the number BEFORE?", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "What is the number AFTER?", options: { bullet: true, fontSize: 14, color: C.CHARCOAL } },
      ], {
        x: 0.7, y: CONTENT_TOP + 2.2, w: 4.1, h: 1.5,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "In your books:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Round to the nearest TEN", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Round to the nearest HUNDRED", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Round to the nearest THOUSAND", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Write in EXPANDED FORM", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Write the number IN WORDS", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Even  |  Before: 465,101  |  After: 465,103  |  Ten: 465,100  |  Hundred: 465,100  |  Thousand: 465,000", {
        x: 0.5, y: 4.45, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 11, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Fluency");
      addTitle(s, "Mixed Operations Warm-Up", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const fluency = [
        "1.  15 + 8 - 3 + 2",
        "2.  24 - 4 + 2 - 10",
        "3.  66 - 11 + 81 - 9 - 4",
        "4.  5 x 5 x 2 / 5",
        "5.  7 x 9 x 2 / 8",
        "6.  9 x 6 x 3 / 9",
      ];
      s.addText(fluency.map((p, i) => ({
        text: p,
        options: { fontSize: 16, color: C.CHARCOAL, breakLine: i < fluency.length - 1, paraSpaceAfter: 8 },
      })), {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 4.1, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, 2.0, { strip: C.SECONDARY });
      s.addText([
        { text: "Write each in your book", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Solve LEFT to RIGHT", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "Show every step", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "4 minutes", options: { bullet: true, fontSize: 14, color: C.ALERT, bold: true } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: 1.7,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_FLUENCY);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) 22    2) 12    3) 123    4) 10    5) 15.75    6) 18", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 6: LI/SC
  liSlide(pres,
    ["I am learning what BODMAS is and why we use it"],
    [
      "I can name what each letter in BODMAS stands for",
      "I can use BODMAS to solve a simple two-step equation",
      "I can explain why we need a rule for the order of operations",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 7: Launch - the viral problem
  contentSlide(pres, "Launch", C.PRIMARY, "What does this equation equal?",
    [
      "Work it out on your whiteboard",
      "Don't talk yet - just solve it",
      "Show me when you're ready",
    ],
    NOTES_LAUNCH, FOOTER,
    (slide, lg) => {
      // Hero equation panel
      addTextOnShape(slide, "6 + 2 x 3 = ?", {
        x: lg.rightX, y: lg.panelTopPadded, w: lg.rightW, h: 1.6, rectRadius: 0.12,
        fill: { color: C.PRIMARY },
      }, { fontSize: 44, fontFace: FONT_H, color: C.WHITE, bold: true });

      // Two answer boxes - showing the disagreement
      addTextOnShape(slide, "Some say:\n24", {
        x: lg.rightX, y: lg.panelTopPadded + 1.85, w: 2.0, h: 1.3, rectRadius: 0.1,
        fill: { color: C.SECONDARY },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });

      addTextOnShape(slide, "Some say:\n12", {
        x: lg.rightX + 2.2, y: lg.panelTopPadded + 1.85, w: 2.0, h: 1.3, rectRadius: 0.1,
        fill: { color: C.ACCENT },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });

      slide.addText("Why do we get different answers?", {
        x: lg.rightX, y: lg.panelTopPadded + 3.25, w: lg.rightW, h: 0.4,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });
    }
  );

  // Slide 8: BODMAS defined - the acronym
  contentSlide(pres, "I Do", C.PRIMARY, "BODMAS - the Order of Operations",
    [
      "We need a rule everyone agrees on",
      "BODMAS gives us that rule",
      "Work from the TOP of the list to the BOTTOM",
    ],
    NOTES_BODMAS_DEFINE, FOOTER,
    (slide, lg) => {
      const letters = [
        { L: "B", word: "Brackets",         color: C.PRIMARY,   note: "Solve inside ( ) first" },
        { L: "O", word: "Orders",           color: C.SECONDARY, note: "Powers (e.g. 2²)" },
        { L: "D", word: "Division",         color: C.ACCENT,    note: "Same step as M" },
        { L: "M", word: "Multiplication",   color: C.ACCENT,    note: "Left to right" },
        { L: "A", word: "Addition",         color: C.ALERT,     note: "Same step as S" },
        { L: "S", word: "Subtraction",      color: C.ALERT,     note: "Left to right" },
      ];
      const rowH = 0.50;
      const startY = lg.panelTopPadded;
      letters.forEach((l, i) => {
        const y = startY + i * (rowH + 0.05);
        // Letter box
        addTextOnShape(slide, l.L, {
          x: lg.rightX, y, w: 0.55, h: rowH, rectRadius: 0.06,
          fill: { color: l.color },
        }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
        // Word
        slide.addText(l.word, {
          x: lg.rightX + 0.65, y, w: 1.7, h: rowH,
          fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bold: true, valign: "middle", margin: 0,
        });
        // Note
        slide.addText(l.note, {
          x: lg.rightX + 2.4, y, w: lg.rightW - 2.4, h: rowH,
          fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true, valign: "middle", margin: 0,
        });
      });
    }
  );

  // Slide 9: I Do - solve the viral problem with BODMAS
  workedExSlide(pres, 2, "I Do", "Now Let's Use BODMAS",
    [
      "Equation: 6 + 2 x 3",
      "",
      "Step 1 - Brackets? NO",
      "Step 2 - Orders? NO",
      "Step 3 - Division or Multiplication?",
      "  YES - 2 x 3 = 6",
      "  Equation now: 6 + 6",
      "",
      "Step 4 - Addition or Subtraction?",
      "  YES - 6 + 6 = 12",
      "",
      "Final answer: 12",
    ],
    NOTES_IDO_VIRAL, FOOTER,
    (slide, lg) => {
      // BODMAS checklist visual
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.85, { strip: C.PRIMARY });
      slide.addText("BODMAS Checklist", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.30,
        fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      const steps = [
        { L: "B", status: "skip",  text: "No brackets" },
        { L: "O", status: "skip",  text: "No orders" },
        { L: "D/M", status: "do",  text: "2 x 3 = 6" },
        { L: "A/S", status: "do",  text: "6 + 6 = 12" },
      ];
      steps.forEach((st, i) => {
        const y = lg.panelTopPadded + 0.45 + i * 0.55;
        const fill = st.status === "do" ? C.SUCCESS : C.MUTED;
        addTextOnShape(slide, st.L, {
          x: lg.rightX + 0.25, y, w: 0.7, h: 0.42, rectRadius: 0.06,
          fill: { color: fill },
        }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(st.text, {
          x: lg.rightX + 1.05, y, w: lg.rightW - 1.25, h: 0.42,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });

      // Final answer banner
      addTextOnShape(slide, "ANSWER: 12", {
        x: lg.rightX, y: lg.panelTopPadded + 2.95, w: lg.rightW, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10-11: CFU 1 with reveal
  withReveal(
    () => cfuSlide(pres, "CFU", "What Comes FIRST?", "Show Me Boards",
      "10 - 6 / 2\n\nWrite the FIRST operation you would do.\n(Just the operation, not the answer)",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "FIRST: Division (6 / 2 = 3)    THEN: Subtraction (10 - 3 = 7)    Answer: 7", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU1_A);
    }
  );

  // Slide 12-13: We Do
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Let's Try Together",
      [
        "Equation: 20 - 4 x 3 + 2",
        "",
        "What comes first?",
        "B? O? D/M? A/S?",
        "",
        "Step by step on your whiteboard",
        "Show every step, like the I Do",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.SECONDARY });
        slide.addText("With your partner:", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.36,
          fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        slide.addText([
          { text: "Identify the operations", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
          { text: "Tick the BODMAS letters that appear", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
          { text: "Solve in BODMAS order", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
          { text: "Show every step", options: { bullet: true, fontSize: 14, color: C.CHARCOAL } },
        ], {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.6,
          fontFace: FONT_B, margin: 0, valign: "top",
        });

        // Reminder pill
        addTextOnShape(slide, "Trap: don't go left to right!", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.5, w: lg.rightW - 0.4, h: 0.5, rectRadius: 0.08,
          fill: { color: C.ALERT },
        }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "Step 1: 4 x 3 = 12   |   Step 2: 20 - 12 = 8   |   Step 3: 8 + 2 = 10", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 14-15: Hinge CFU
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Question", "Show Me Boards",
      "12 + 8 / 4\n\nWrite the FINAL answer on your whiteboard.",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Answer: 14   (8 / 4 = 2, then 12 + 2 = 14)   Trap: 5 = left-to-right error", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU2_A);
    }
  );

  // Slide 16: You Do
  workedExSlide(pres, 4, "You Do", "Independent Practice",
    [
      "Solve each equation using BODMAS",
      "Show your steps - WHICH operation came first?",
      "",
      "1.  14 + 7 - 10 - 4",
      "2.  3 + 3 x 3 - 5",
      "3.  20 - 6 / 2 + 1",
      "4.  4 x 5 + 6 - 3",
      "5.  18 / 6 + 7 x 2",
      "",
      "Challenge: 30 - 4 x 5 + 2 x 3",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.ALERT });
      slide.addText("Strategy reminder:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.36,
        fontSize: 15, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "1. Brackets first", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "2. Orders (powers)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "3. Division/Multiplication, left to right", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "4. Addition/Subtraction, left to right", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.8,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addTextOnShape(slide, "8 minutes", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.45, w: lg.rightW - 0.4, h: 0.4, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "Solve:  10 + 2 x 4    (show your steps)",
      "Solve:  20 - 12 / 3    (show your steps)",
      "Tell your teacher: WHY do we need BODMAS?",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 18: Closing
  closingSlide(pres, {
    reflectionPrompt: "Tell your partner one thing you learned about BODMAS today.",
    scItems: [
      "I can name what each letter in BODMAS stands for",
      "I can use BODMAS to solve a simple two-step equation",
      "I can explain why we need a rule for the order of operations",
    ],
    selfAssessment: { prompt: "Self-check", options: ["Got it", "Getting there", "Need more practice"] },
  }, NOTES_CLOSING);

  // Slide 19: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "BOD_Session1_What_Is_BODMAS.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  // Reference Sheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "BODMAS - Order of Operations",
      color: C.NAVY,
      lessonInfo: "Session 1 of 10 | Year 5/6 Numeracy",
    });
    y = addTipBox(doc, "Stick this in your maths book. Look at it every time you solve an equation with mixed operations.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "BODMAS - the Order of Operations", y, { color: C.NAVY });
    y = addBodyText(doc, "B - Brackets   solve inside ( ) first", y);
    y = addBodyText(doc, "O - Orders   powers, like 2 squared", y);
    y = addBodyText(doc, "D - Division   same step as M", y);
    y = addBodyText(doc, "M - Multiplication   same step as D", y);
    y = addBodyText(doc, "A - Addition   same step as S", y);
    y = addBodyText(doc, "S - Subtraction   same step as A", y);
    y += 6;
    y = addBodyText(doc, "When two letters are at the SAME LEVEL (D and M, A and S), work LEFT TO RIGHT.", y);

    y = addSectionHeading(doc, "Worked Example", y, { color: C.NAVY });
    y = addBodyText(doc, "Equation:   6 + 2 x 3", y);
    y = addBodyText(doc, "Step 1:   Brackets? NO", y);
    y = addBodyText(doc, "Step 2:   Orders? NO", y);
    y = addBodyText(doc, "Step 3:   D or M? YES   2 x 3 = 6   --> 6 + 6", y);
    y = addBodyText(doc, "Step 4:   A or S? YES   6 + 6 = 12", y);
    y = addBodyText(doc, "Answer:   12", y);

    y = addSectionHeading(doc, "Common Trap", y, { color: C.NAVY });
    y = addBodyText(doc, "Going LEFT TO RIGHT and ignoring BODMAS gets you the WRONG answer.", y);
    y = addBodyText(doc, "Example: 12 + 8 / 4", y);
    y = addBodyText(doc, "  WRONG (left to right):   12 + 8 = 20, 20 / 4 = 5", y);
    y = addBodyText(doc, "  RIGHT (BODMAS):   8 / 4 = 2, 12 + 2 = 14", y);

    addPdfFooter(doc, "Session 1 | BODMAS Reference | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Practice Sheet
  await (async () => {
    const doc = createPdf({ title: PRACTICE_RES.name });
    let y = addPdfHeader(doc, PRACTICE_RES.name, {
      subtitle: "Two-Step BODMAS Practice",
      color: C.NAVY,
      lessonInfo: "Session 1 of 10 | Year 5/6 Numeracy",
    });
    y = addTipBox(doc, "For each equation: write down the FIRST operation you do, work through BODMAS, and show every step.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A - Two-step equations", y, { color: C.NAVY });
    y = addProblem(doc, 1, "10 + 2 x 4 =                          First step: __________", y, { color: C.NAVY });
    y = addProblem(doc, 2, "20 - 6 / 3 =                          First step: __________", y, { color: C.NAVY });
    y = addProblem(doc, 3, "15 + 12 / 4 =                         First step: __________", y, { color: C.NAVY });
    y = addProblem(doc, 4, "25 - 5 x 3 =                          First step: __________", y, { color: C.NAVY });
    y = addProblem(doc, 5, "8 x 3 + 6 =                           First step: __________", y, { color: C.NAVY });
    y = addProblem(doc, 6, "30 / 6 + 7 =                          First step: __________", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Section B - Three-step equations", y, { color: C.NAVY });
    y = addProblem(doc, 7, "10 + 4 x 2 - 3 =                      Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 8, "24 / 4 + 5 - 2 =                      Show every step", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addProblem(doc, 9, "30 - 4 x 5 + 2 x 3 =                  Show every step", y, { color: C.NAVY });
    y += 8;
    y = addBodyText(doc, "Reflect: Why is BODMAS important?", y);
    y = addWriteLine(doc, "", y);
    y = addWriteLine(doc, "", y);

    addPdfFooter(doc, "Session 1 | Practice Sheet | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, PRACTICE_RES.fileName));
    console.log("PDF written: " + PRACTICE_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "BODMAS Practice Answer Key",
      color: C.NAVY,
      lessonInfo: "Session 1 of 10 | Year 5/6 Numeracy",
    });

    y = addSectionHeading(doc, "Section A - Two-step", y, { color: C.NAVY });
    y = addBodyText(doc, "1.  10 + 2 x 4   --> 2 x 4 = 8, 10 + 8 = 18   First: Multiplication", y);
    y = addBodyText(doc, "2.  20 - 6 / 3   --> 6 / 3 = 2, 20 - 2 = 18   First: Division", y);
    y = addBodyText(doc, "3.  15 + 12 / 4  --> 12 / 4 = 3, 15 + 3 = 18  First: Division", y);
    y = addBodyText(doc, "4.  25 - 5 x 3   --> 5 x 3 = 15, 25 - 15 = 10  First: Multiplication", y);
    y = addBodyText(doc, "5.  8 x 3 + 6    --> 8 x 3 = 24, 24 + 6 = 30  First: Multiplication", y);
    y = addBodyText(doc, "6.  30 / 6 + 7   --> 30 / 6 = 5, 5 + 7 = 12   First: Division", y);

    y = addSectionHeading(doc, "Section B - Three-step", y, { color: C.NAVY });
    y = addBodyText(doc, "7.  10 + 4 x 2 - 3  -->  4 x 2 = 8  -->  10 + 8 = 18  -->  18 - 3 = 15", y);
    y = addBodyText(doc, "8.  24 / 4 + 5 - 2  -->  24 / 4 = 6  -->  6 + 5 = 11  -->  11 - 2 = 9", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "9.  30 - 4 x 5 + 2 x 3", y);
    y = addBodyText(doc, "    Step 1: 4 x 5 = 20  -->  30 - 20 + 2 x 3", y);
    y = addBodyText(doc, "    Step 2: 2 x 3 = 6   -->  30 - 20 + 6", y);
    y = addBodyText(doc, "    Step 3: 30 - 20 = 10  -->  10 + 6", y);
    y = addBodyText(doc, "    Step 4: 10 + 6 = 16   ANSWER: 16", y);

    addPdfFooter(doc, "Session 1 | Answer Key | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Session 1 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
