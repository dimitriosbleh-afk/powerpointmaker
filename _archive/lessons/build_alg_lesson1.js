"use strict";

// Algebra Unit — Session 1: Equivalent Number Sentences
// Week 1 Session 1, Grade 5/6 Numeracy, Variant 0
// DR: Exploring Mathematical Patterns and Algorithms
// Fluency: Multiplication facts speed drill
// VC2M5A02 — find unknown values in numerical equations involving multiplication and division

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  formatSessionResourceFileName,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", 0);
const {
  C, FONT_H, FONT_B,
  SLIDE_W, SLIDE_H, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addBadge, addTitle,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const UNIT_TITLE = "Algebra: Finding Unknown Values";
const FOOTER = "Algebra | Session 1 of 10 | Grade 5/6 Numeracy";
const OUT_DIR = "output/ALG_Session1_Equivalent_Number_Sentences";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Independent practice - identifying and building equivalent number sentences.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with worked answers.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Pre-filled balance model scaffold for equivalent sentences.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Self-contained investigation into chains of equivalent expressions.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ───────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- We are starting a new maths unit today -- algebra
- Algebra is about finding missing values and understanding how numbers relate to each other
- Over the next two weeks you will learn to crack equations and find unknown values

DO:
- Display title slide as students settle
- Have whiteboards and markers ready

TEACHER NOTES:
Session 1 of 10. This lesson establishes the foundational concept of equivalence -- that both sides of an equals sign must have the same value. Without this understanding, students cannot meaningfully solve for unknowns in later sessions.

WATCH FOR:
- Students who seem anxious about the word "algebra" -- frame it positively: "We are building detective skills with numbers"

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Let's warm up with some pattern work from last week
- Look at each sequence and find the missing number
- Use what you know about how patterns grow or shrink

DO:
- Display the four pattern problems
- Allow 90 seconds total, students work on whiteboards
- Circulate to check reasoning, not just answers

TEACHER NOTES:
Daily Review retrieves prior learning on mathematical patterns and algorithms. Students should identify the rule (operation) driving each pattern before finding the missing value. This is spaced retrieval, not new content.

WATCH FOR:
- Students who guess without identifying the pattern rule -- prompt: "What is happening each step? Is it adding, multiplying, or something else?"
- Students who identify the rule but apply it inconsistently

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers -- tick if you got it right, fix it if you didn't
- Fixing it is the learning
- Ask: What rule did you use for pattern 2? [Multiply by 2 each time]

DO:
- Click to reveal answers
- Students tick correct responses and fix errors
- Scan for common mistakes

TEACHER NOTES:
Tick-and-fix gives immediate feedback on pattern recognition. Note students who struggle -- pattern thinking underpins the algebraic reasoning in this unit.

WATCH FOR:
- Students who got all four correct quickly -- pattern recognition is strong
- Students who could not identify the rule -- they may need extra support with multiplicative thinking

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time -- multiplication facts speed round
- I will show you 10 facts. Write each answer on your whiteboard as fast as you can
- Ready? Go.

DO:
- Display the 10 multiplication facts
- Time 2 minutes. Students work on whiteboards
- After time, reveal answers for self-check
- Note students who are not yet automatic -- plan future fluency around their gaps

TEACHER NOTES:
Fluency builds automaticity with multiplication facts. Strong multiplication recall is essential for this algebra unit -- students who are slow on facts will have working memory consumed by calculation rather than algebraic reasoning.

WATCH FOR:
- Students who count on fingers for basic facts -- they need continued fluency support
- Students who finish quickly and accurately -- multiplication recall is automated

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me
- Read from slide: "We are learning to understand equivalent number sentences so we can find unknown values"
- Now our success criteria. Read each one together.
- Ask: What do you think "equivalent" means? [Same value, equal, balanced]

DO:
- Choral read the LI then each SC
- Brief think-pair-share on what equivalent might mean

TEACHER NOTES:
This is the gateway lesson. Equivalence is the foundational concept -- if students do not understand that both sides of = must have the same value, they cannot solve equations. SC1 is the floor (recognising equal values), SC2 is the target (building equivalent sentences), SC3 extends to justification.

WATCH FOR:
- Students who think the equals sign means "the answer is" rather than "is the same value as" -- this is the critical misconception to address in I Do

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. I am going to think about what the equals sign really means
- Most people think equals means "the answer is." But it actually means "is the same value as"
- Think of it like a balance scale. Both sides must balance
- Look: 3 x 5 = 15. The left side, 3 x 5, equals 15. The right side is 15. They balance.
- Now watch this: 3 x 5 = 30 / 2. Is this true?
- Left side: 3 x 5 = 15. Right side: 30 / 2 = 15. Both sides equal 15. So yes, they are equivalent.

DO:
- Point to the balance visual on screen
- Write 3 x 5 = 15 and 30 / 2 = 15 clearly
- Draw arrows showing both sides equal 15
- Emphasise the balance metaphor throughout

TEACHER NOTES:
The equals-sign-as-balance is the threshold concept. Research shows most primary students hold an operational view (= means "write the answer") rather than a relational view (= means "same value"). This I Do must shift that understanding. The balance visual makes the relational view concrete.

MISCONCEPTIONS:
- Misconception: The equals sign means "the answer comes next"
  Why: Years of arithmetic practice in the form a + b = ___ trains students to see = as a signal to calculate and write an answer on the right
  Impact: Students cannot make sense of equations like 15 = 3 x 5 or 3 x 5 = 30 / 2 because no "answer" appears after the equals sign
  Quick correction: "The equals sign is like a balance. Both sides must weigh the same. 15 = 3 x 5 is perfectly valid because both sides equal 15."

WATCH FOR:
- Students who look confused when you write 3 x 5 = 30 / 2 -- they may expect a single number after the equals sign
- Readiness signal: students nodding when you explain the balance metaphor

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now I will show you how to check if two expressions are equivalent
- Here is my method. I work out each side separately, then compare
- Look: Is 4 x 6 = 8 x 3?
- Left side: 4 x 6 = 24. Right side: 8 x 3 = 24. Same value. So yes, equivalent.
- Another one: Is 5 x 8 = 10 x 3?
- Left side: 5 x 8 = 40. Right side: 10 x 3 = 30. Different values. Not equivalent.
- My check: I always calculate both sides and compare. If the values match, the equation is true.

DO:
- Write each example step by step
- Circle the final values on each side to compare
- Use a tick for equivalent, a cross for not equivalent

TEACHER NOTES:
This I Do models the verification strategy: evaluate each side, compare. This strategy is the workhorse for the entire unit. Students who automate "work out each side, compare" have a reliable method for checking any equation.

WATCH FOR:
- Students who try to calculate across the equals sign (e.g. 4 x 6 = 8 x 3 becoming 24 = 8 x 3 = 72) -- they are chaining rather than comparing
- Readiness signal: students saying "same" or "different" before you announce the result

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. On your whiteboards, write TRUE or FALSE
- Is this equation true: 6 x 4 = 12 x 2?
- Work out each side. 10 seconds. Show me.

DO:
- Students write on whiteboards
- Scan for TRUE (both sides = 24)

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "On your whiteboard, write TRUE or FALSE. Is 6 x 4 = 12 x 2? Work out each side first. You have 10 seconds. Show me on my signal."
- Scan for: TRUE. Students who show TRUE quickly have grasped the evaluate-and-compare strategy.
PROCEED: If 80%+ show TRUE, move to the next I Do example.
PIVOT: Most likely misconception -- students write FALSE because they think both sides need to look the same (same numbers, same operation). Reteach: "Equivalent does not mean identical. 6 x 4 looks different from 12 x 2, but both equal 24. The VALUE matters, not the appearance."

TEACHER NOTES:
First CFU checks whether students can apply the evaluate-and-compare strategy. Both sides equal 24, so the equation is true.

WATCH FOR:
- Students who write FALSE -- they may still hold the operational view of equals
- Students who write TRUE instantly without showing working -- confirm they actually evaluated both sides

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with a partner. I will show you four equations
- For each one, work out the left side, work out the right side, then write TRUE or FALSE
- Show your working, not just the answer

DO:
- Display the four equations
- Partners work together on whiteboards
- Circulate -- look for the evaluate-and-compare method
- Allow 2 minutes, then cold call for each

TEACHER NOTES:
We Do checks all four operations in equivalence contexts. Requiring working (not just T/F) ensures students are using the strategy, not guessing. The mix of true and false equations prevents pattern-matching.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 1 Enabling Scaffold which provides a balance model template. Students write each side's calculation in the balance pans, find the total, then compare. Scaffold has the first equation pre-filled as a model.
- Extra Notes: Distribute the Session 1 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: For each TRUE equation, write a THIRD expression that is also equivalent. For example, if 4 x 6 = 12 x 2, find another way to make 24 (e.g., 24 x 1, 48 / 2).
- Extra Notes: This builds toward the chain-of-equivalence concept in the extension investigation.

WATCH FOR:
- Students who evaluate one side but forget to evaluate the other -- prompt: "You need both sides to compare"
- Readiness signal: partners finishing all four within 90 seconds with correct working shown

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check. First equation: 7 x 4 = 14 x 2. Left side? [28] Right side? [28] True or false? [TRUE]
- Second: 9 x 3 = 6 x 5. Left? [27] Right? [30] [FALSE]
- Third: 36 / 4 = 18 / 2. Left? [9] Right? [9] [TRUE]
- Fourth: 5 x 8 = 20 x 3. Left? [40] Right? [60] [FALSE]

DO:
- Reveal answers one at a time
- Cold call different students for each
- For incorrect equations, ask: "What would we need to change to make it true?"

TEACHER NOTES:
The follow-up question "What would we need to change?" previews the unknown-finding work in Sessions 2-4. It plants the seed without making it the focus yet.

WATCH FOR:
- Students who got the TRUE ones right but struggled with FALSE -- they may be assuming all equations shown are true
- Students who made calculation errors but used the right strategy -- affirm the method, correct the arithmetic

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. Thumbs up if the equation is true, thumbs down if false
- Ready? 15 / 3 = 10 / 2
- Think... show me now.

DO:
- Students show thumbs
- Scan: both sides equal 5, so TRUE (thumbs up)
- Follow up: "How do you know?" Cold call 2 students

CFU CHECKPOINT:
Technique: Thumbs Up/Down
Script:
- Say: "Thumbs up for true, thumbs down for false. 15 / 3 = 10 / 2. Think... show me."
- Scan for: thumbs up (both sides = 5).
PROCEED: If 80%+ show thumbs up with confidence, move to You Do.
PIVOT: Most likely misconception -- students show thumbs down because 15/3 and 10/2 use different numbers and they have not evaluated both sides. Reteach: "Let's slow down. Left side: 15 divided by 3. What is that? [5]. Right side: 10 divided by 2. What is that? [5]. Same value? [Yes]. So the equation is true."

TEACHER NOTES:
This hinge uses division to check transfer from the multiplication examples. If students can evaluate both sides with division, the strategy has generalised.

WATCH FOR:
- Students who hesitate -- they may be unsure about division facts
- Readiness signal: confident thumbs up with no hesitation

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice time
- Open your worksheet. You have 12 equations to check -- TRUE or FALSE
- First: Work out the left side. Next: Work out the right side. Then: Write TRUE or FALSE
- Show your working for every equation
- You have 8 minutes. Go.

DO:
- Distribute Session 1 Worksheet
- Students work independently
- Circulate: prioritise students who struggled during We Do
- For enabling students, distribute the Session 1 Enabling Scaffold
- For extending students, distribute the Session 1 Extension

TEACHER NOTES:
You Do uses different numbers and mixes all four operations. Problems progress from single-operation equations to mixed-operation equations (e.g., 3 x 8 = 48 / 2). The challenge section introduces writing your own equivalent sentences.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Complete the enabling scaffold with balance model templates. Each equation is laid out with two balance pans. Students calculate each side in the pan, then compare. The first two equations are pre-filled as models.
- Extra Notes: Distribute the Session 1 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Investigate chains of equivalent expressions. Given that 24 can be written as 4 x 6, 8 x 3, 12 x 2, 24 x 1, 48 / 2, etc., find the longest chain of equivalent expressions you can build for the numbers 36 and 60.
- Extra Notes: Distribute the Session 1 Extension PDF. Self-contained investigation.

WATCH FOR:
- Students who skip the working and just write T/F -- redirect: "Show me HOW you know"
- Students who finish all 12 quickly and correctly -- direct to extension
- Readiness signal: 80%+ completing 8+ problems correctly within 6 minutes

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket time. Three quick questions in your workbook
- Work on your own. You have 3 minutes.

DO:
- Display the exit ticket questions
- Students work silently in workbooks
- Collect or scan responses to sort into secure, developing, beginning

TEACHER NOTES:
Exit ticket assesses SC1 (Q1 -- recognise equivalence) and SC2 (Q2-Q3 -- construct equivalent sentences). Sort results to plan support for Session 2.

WATCH FOR:
- Students who get Q1 correct but struggle with Q2-Q3 -- they can verify but not yet construct
- Students who complete all three instantly -- they are ready for unknown-finding in Session 2

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Let's check our success criteria. Look at the slide.
- SC1: I can decide whether two number sentences have the same value. Thumbs up, sideways, or down.
- SC2: I can build my own equivalent number sentence for a given value. Thumbs?
- SC3: I can explain why two expressions are equivalent using the balance model. Thumbs?
- Turn and talk: What does the equals sign REALLY mean? Tell your partner in your own words.

DO:
- Display success criteria on screen
- Run thumbs check for each SC
- Allow 30 seconds for Turn and Talk
- Cold call 2-3 students to share

TEACHER NOTES:
Closing brings the lesson full circle. The Turn and Talk targets the relational understanding of equals -- the threshold concept for this unit. Students who can articulate "same value on both sides" have shifted from operational to relational thinking.

WATCH FOR:
- Students who still say "equals means the answer" -- they need more work on the balance model in Session 2
- Students who use the word "balance" or "same value" -- the relational view is forming

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- These are the resources for today's lesson
- The worksheet and answer key are linked here if you need to reprint

DO:
- Point out each resource and its purpose

TEACHER NOTES:
Resource slide provides clickable links to all companion PDFs for this session.

WATCH FOR:
- N/A -- resource slide is teacher reference only

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ───────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Session 1: Equivalent Number Sentences", "Grade 5/6 Numeracy | Session 1 of 10 | Week 1", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2-3: Daily Review (withReveal) — Exploring Mathematical Patterns and Algorithms
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Patterns and Algorithms", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const problems = [
        "1.  2, 6, 18, 54, ___  (rule: ?)",
        "2.  5, 10, 20, 40, ___  (rule: ?)",
        "3.  96, 48, 24, 12, ___  (rule: ?)",
        "4.  3, 12, 48, 192, ___  (rule: ?)",
      ];
      s.addText(problems.map((p, i) => ({
        text: p,
        options: { fontSize: 14.5, color: C.CHARCOAL, breakLine: i < problems.length - 1, paraSpaceAfter: 12 },
      })), {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboards:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "Find the missing number in each pattern.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Steps:", options: { fontSize: 14, bold: true, color: C.ALERT, breakLine: true } },
        { text: "Identify the rule (what operation?)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Apply the rule to find the next number", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Check: does your answer fit the pattern?", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Answers:  1) 162 (x3)    2) 80 (x2)    3) 6 (/2)    4) 768 (x4)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Multiplication facts speed drill
  const sFluency = pres.addSlide();
  addTopBar(sFluency, STAGE_COLORS["1"]);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Multiplication Facts Speed Round", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

  addCard(sFluency, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  const fluencyFacts = [
    "1.  7 x 8 =",    "2.  6 x 9 =",
    "3.  12 x 4 =",   "4.  8 x 5 =",
    "5.  9 x 7 =",    "6.  11 x 6 =",
    "7.  3 x 12 =",   "8.  8 x 8 =",
    "9.  6 x 7 =",    "10. 9 x 4 =",
  ];
  sFluency.addText(fluencyFacts.map((p, i) => ({
    text: p,
    options: { fontSize: 14, color: C.CHARCOAL, breakLine: i < fluencyFacts.length - 1, paraSpaceAfter: 4 },
  })), {
    x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addCard(sFluency, 5.2, CONTENT_TOP, 4.3, 1.8, { strip: C.SECONDARY });
  sFluency.addText([
    { text: "Whiteboards - 2 minutes", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
    { text: "Write each answer as fast as you can.", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
    { text: "Speed AND accuracy matter.", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Answers: 56, 54, 48, 40, 63, 66, 36, 64, 42, 36", options: { fontSize: 11, color: C.MUTED } },
  ], {
    x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: 1.6,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    ["We are learning to understand equivalent number sentences so we can find unknown values"],
    [
      "I can decide whether two number sentences have the same value",
      "I can build my own equivalent number sentence for a given value",
      "I can explain why two expressions are equivalent using the balance model",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — What does equals REALLY mean?
  workedExSlide(pres, 2, "I Do", "What Does Equals Really Mean?",
    [
      "The equals sign means 'is the same value as'",
      "Think of it as a balance scale",
      "",
      "Example 1:",
      "  3 x 5 = 15      (left = 15, right = 15)",
      "",
      "Example 2:",
      "  3 x 5 = 30 / 2",
      "  Left: 3 x 5 = 15",
      "  Right: 30 / 2 = 15",
      "  Same value -- equivalent!",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      // Balance scale visual
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.PRIMARY });
      slide.addText("Balance Model", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      // Draw balance beam
      const beamY = lg.panelTopPadded + 1.6;
      const beamCenterX = lg.rightX + lg.rightW / 2;
      // Triangle base
      slide.addShape("isoscelesTriangle", {
        x: beamCenterX - 0.25, y: beamY + 0.02, w: 0.5, h: 0.4,
        fill: { color: C.CHARCOAL },
      });
      // Beam line
      slide.addShape("line", {
        x: lg.rightX + 0.4, y: beamY, w: lg.rightW - 0.8, h: 0,
        line: { color: C.CHARCOAL, width: 3 },
      });

      // Left pan
      addTextOnShape(slide, "3 x 5", {
        x: lg.rightX + 0.4, y: beamY - 0.7, w: 1.4, h: 0.55, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addText("= 15", {
        x: lg.rightX + 0.4, y: beamY - 0.15, w: 1.4, h: 0.25,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
      });

      // Right pan
      addTextOnShape(slide, "30 / 2", {
        x: lg.rightX + lg.rightW - 1.8, y: beamY - 0.7, w: 1.4, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SECONDARY },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addText("= 15", {
        x: lg.rightX + lg.rightW - 1.8, y: beamY - 0.15, w: 1.4, h: 0.25,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
      });

      // Equals / balanced label
      addTextOnShape(slide, "BALANCED!", {
        x: beamCenterX - 0.8, y: beamY + 0.5, w: 1.6, h: 0.35, rectRadius: 0.06,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 7: I Do — How to check equivalence
  workedExSlide(pres, 2, "I Do", "Checking Equivalence",
    [
      "Method: Evaluate each side, then compare",
      "",
      "Is 4 x 6 = 8 x 3 ?",
      "  Left:  4 x 6 = 24",
      "  Right: 8 x 3 = 24",
      "  24 = 24  --> TRUE (equivalent)",
      "",
      "Is 5 x 8 = 10 x 3 ?",
      "  Left:  5 x 8 = 40",
      "  Right: 10 x 3 = 30",
      "  40 =/= 30  --> FALSE (not equivalent)",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.4, { strip: C.SUCCESS });
      slide.addText("Strategy", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 15, fontFace: FONT_H, color: C.SUCCESS, bold: true, margin: 0,
      });
      const stratSteps = [
        "1. Work out the LEFT side",
        "2. Work out the RIGHT side",
        "3. Compare the values",
        "4. Same = TRUE, Different = FALSE",
      ];
      slide.addText(stratSteps.map((st, i) => ({
        text: st,
        options: { fontSize: 13, color: C.CHARCOAL, breakLine: i < stratSteps.length - 1, paraSpaceAfter: 6 },
      })), {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.5, h: 0.9,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Result cards
      addTextOnShape(slide, "4 x 6 = 8 x 3    24 = 24    TRUE", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.65, w: lg.rightW - 0.3, h: 0.4, rectRadius: 0.06,
        fill: { color: C.SUCCESS },
      }, { fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true });

      addTextOnShape(slide, "5 x 8 = 10 x 3    40 =/= 30    FALSE", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.15, w: lg.rightW - 0.3, h: 0.4, rectRadius: 0.06,
        fill: { color: C.ALERT },
      }, { fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 8-9: CFU 1 (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check", "Show Me Boards",
      "TRUE or FALSE?\n\n6 x 4 = 12 x 2\n\nWork out each side first.",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "TRUE:  6 x 4 = 24    12 x 2 = 24    Same value!", {
        x: 1.0, y: 4.0, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10-11: We Do (withReveal)
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Your Turn — TRUE or FALSE?",
      [
        "Work out each side. Write TRUE or FALSE.",
        "Show your working!",
        "",
        "1.   7 x 4 = 14 x 2",
        "2.   9 x 3 = 6 x 5",
        "3.   36 / 4 = 18 / 2",
        "4.   5 x 8 = 20 x 3",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.SECONDARY });
        slide.addText("With your partner:", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
          fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        slide.addText([
          { text: "Work out the LEFT side", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Work out the RIGHT side", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Compare: same or different?", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Write TRUE or FALSE", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
        ], {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.45, w: lg.rightW - 0.5, h: 1.2,
          fontFace: FONT_B, margin: 0, valign: "top",
        });

        // Balance icon reminder
        addTextOnShape(slide, "Remember: = means BALANCE", {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.8, w: lg.rightW - 0.6, h: 0.4, rectRadius: 0.06,
          fill: { color: C.PRIMARY },
        }, { fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "1) TRUE (28=28)   2) FALSE (27/30)   3) TRUE (9=9)   4) FALSE (40/60)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 12-13: CFU 2 Hinge (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Check", "Thumbs Up/Down",
      "TRUE or FALSE?\n\n15 / 3 = 10 / 2",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "TRUE:  15 / 3 = 5    10 / 2 = 5    Both equal 5!", {
        x: 1.0, y: 4.0, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do
  workedExSlide(pres, 4, "You Do", "Independent Practice",
    [
      "First: Read each equation on your worksheet.",
      "Next: Work out the left side and the right side.",
      "Then: Write TRUE or FALSE with your working.",
      "",
      "Show your working for every equation.",
      "If you finish, try the challenge section.",
      "",
      "You have 8 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.0, { strip: C.ALERT });
      slide.addText("Strategy Reminder:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "1. Left side = ?", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "2. Right side = ?", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "3. Same value = TRUE", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "4. Different value = FALSE", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.45, w: lg.rightW - 0.5, h: 1.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "TRUE or FALSE:  8 x 5 = 20 x 2   (show working)",
      "Write a number sentence that is equivalent to 6 x 6 = 36.",
      "Explain WHY 4 x 9 = 12 x 3 is true, using the balance model.",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 16: Closing
  closingSlide(pres,
    "What does the equals sign REALLY mean? Tell your partner in your own words.",
    [
      "I can decide whether two number sentences have the same value",
      "I can build my own equivalent number sentence for a given value",
      "I can explain why two expressions are equivalent using the balance model",
    ],
    NOTES_CLOSING);

  // Slide 17: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "ALG_Session1_Equivalent_Number_Sentences.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Generate PDFs ─────────────────────────────────────────────────────────

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Equivalent Number Sentences",
      color: C.NAVY,
      lessonInfo: "Session 1 of 10 | Grade 5/6 Numeracy",
    });
    y = addTipBox(doc, "For each equation: work out the LEFT side, work out the RIGHT side, then write TRUE or FALSE. Show your working!", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A: Multiplication Equations", y, { color: C.NAVY });
    y = addProblem(doc, 1, "3 x 9 = 9 x 3          Left: ___   Right: ___   TRUE / FALSE", y, { color: C.NAVY });
    y = addProblem(doc, 2, "7 x 5 = 5 x 8          Left: ___   Right: ___   TRUE / FALSE", y, { color: C.NAVY });
    y = addProblem(doc, 3, "6 x 8 = 16 x 3         Left: ___   Right: ___   TRUE / FALSE", y, { color: C.NAVY });
    y = addProblem(doc, 4, "4 x 12 = 8 x 6         Left: ___   Right: ___   TRUE / FALSE", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Section B: Division Equations", y, { color: C.NAVY });
    y = addProblem(doc, 5, "24 / 6 = 20 / 5        Left: ___   Right: ___   TRUE / FALSE", y, { color: C.NAVY });
    y = addProblem(doc, 6, "42 / 7 = 36 / 6        Left: ___   Right: ___   TRUE / FALSE", y, { color: C.NAVY });
    y = addProblem(doc, 7, "56 / 8 = 49 / 7        Left: ___   Right: ___   TRUE / FALSE", y, { color: C.NAVY });
    y = addProblem(doc, 8, "30 / 5 = 24 / 3        Left: ___   Right: ___   TRUE / FALSE", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Section C: Mixed Operations", y, { color: C.NAVY });
    y = addProblem(doc, 9, "3 x 8 = 48 / 2         Left: ___   Right: ___   TRUE / FALSE", y, { color: C.NAVY });
    y = addProblem(doc, 10, "5 x 7 = 70 / 2        Left: ___   Right: ___   TRUE / FALSE", y, { color: C.NAVY });
    y = addProblem(doc, 11, "9 x 4 = 72 / 2        Left: ___   Right: ___   TRUE / FALSE", y, { color: C.NAVY });
    y = addProblem(doc, 12, "6 x 11 = 132 / 2      Left: ___   Right: ___   TRUE / FALSE", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Challenge: Build Your Own", y, { color: C.NAVY });
    y = addProblem(doc, 13, "Write TWO different equivalent number sentences for the value 48.", y, { color: C.NAVY, writeLines: [{ label: "" }, { label: "" }] });

    addPdfFooter(doc, "Session 1 | Algebra: Finding Unknown Values | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Equivalent Number Sentences",
      color: C.NAVY,
      lessonInfo: "Session 1 of 10 | Grade 5/6 Numeracy",
    });
    y = addSectionHeading(doc, "Section A: Multiplication", y, { color: C.NAVY });
    y = addBodyText(doc, "1. 3 x 9 = 27, 9 x 3 = 27 -> TRUE (commutative property)", y);
    y = addBodyText(doc, "2. 7 x 5 = 35, 5 x 8 = 40 -> FALSE (35 =/= 40)", y);
    y = addBodyText(doc, "3. 6 x 8 = 48, 16 x 3 = 48 -> TRUE", y);
    y = addBodyText(doc, "4. 4 x 12 = 48, 8 x 6 = 48 -> TRUE", y);
    y = addSectionHeading(doc, "Section B: Division", y, { color: C.NAVY });
    y = addBodyText(doc, "5. 24 / 6 = 4, 20 / 5 = 4 -> TRUE", y);
    y = addBodyText(doc, "6. 42 / 7 = 6, 36 / 6 = 6 -> TRUE", y);
    y = addBodyText(doc, "7. 56 / 8 = 7, 49 / 7 = 7 -> TRUE", y);
    y = addBodyText(doc, "8. 30 / 5 = 6, 24 / 3 = 8 -> FALSE (6 =/= 8)", y);
    y = addSectionHeading(doc, "Section C: Mixed Operations", y, { color: C.NAVY });
    y = addBodyText(doc, "9. 3 x 8 = 24, 48 / 2 = 24 -> TRUE", y);
    y = addBodyText(doc, "10. 5 x 7 = 35, 70 / 2 = 35 -> TRUE", y);
    y = addBodyText(doc, "11. 9 x 4 = 36, 72 / 2 = 36 -> TRUE", y);
    y = addBodyText(doc, "12. 6 x 11 = 66, 132 / 2 = 66 -> TRUE", y);
    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "13. Accept any two correct equivalent expressions for 48.", y);
    y = addBodyText(doc, "    Examples: 6 x 8 = 48, 4 x 12 = 48, 96 / 2 = 48, 16 x 3 = 48", y);
    addPdfFooter(doc, "Session 1 | Answer Key | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold — Balance model template
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Balance Model for Equivalent Number Sentences",
      color: C.TEAL,
      lessonInfo: "Session 1 of 10 | Grade 5/6 Numeracy",
    });
    y = addTipBox(doc, "For each equation: write the calculation in each balance pan. Find the value. If both sides are the same, the equation is TRUE.", y, { color: C.TEAL });

    // Pre-filled example
    y = addSectionHeading(doc, "Example (done for you)", y, { color: C.NAVY });
    y = addBodyText(doc, "Equation: 4 x 6 = 8 x 3", y);
    y += 5;
    // Draw simple balance visual
    const balanceX = 80;
    doc.rect(balanceX, y, 180, 40).stroke("#" + C.NAVY);
    doc.fontSize(11).fillColor("#333333").text("LEFT: 4 x 6 = 24", balanceX + 10, y + 12, { width: 160 });
    doc.rect(balanceX + 220, y, 180, 40).stroke("#" + C.NAVY);
    doc.text("RIGHT: 8 x 3 = 24", balanceX + 230, y + 12, { width: 160 });
    y += 48;
    doc.fontSize(12).fillColor("#" + (C.SUCCESS || "2E7D32")).text("24 = 24 -> TRUE (balanced!)", balanceX, y, { width: 400 });
    y += 25;

    y = addSectionHeading(doc, "Now you try:", y, { color: C.NAVY });

    const scaffoldProblems = [
      "7 x 4 = 14 x 2",
      "9 x 3 = 6 x 5",
      "36 / 4 = 18 / 2",
      "5 x 8 = 20 x 3",
      "3 x 9 = 9 x 3",
      "24 / 6 = 20 / 5",
    ];
    scaffoldProblems.forEach((prob, i) => {
      y = addBodyText(doc, (i + 1) + ". Equation: " + prob, y);
      y += 3;
      const bx = 80;
      doc.rect(bx, y, 180, 35).stroke("#" + C.NAVY);
      doc.fontSize(9).fillColor("#999999").text("LEFT: ___ = ___", bx + 10, y + 10, { width: 160 });
      doc.rect(bx + 220, y, 180, 35).stroke("#" + C.NAVY);
      doc.text("RIGHT: ___ = ___", bx + 230, y + 10, { width: 160 });
      y += 42;
      doc.fillColor("#333333");
      y = addBodyText(doc, "   ___ = ___    TRUE / FALSE", y);
      y += 8;
      if (y > 700) { doc.addPage(); y = 50; }
    });

    addPdfFooter(doc, "Session 1 | Enabling Scaffold | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Extension — Chains of Equivalent Expressions
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Chains of Equivalent Expressions",
      color: C.NAVY,
      lessonInfo: "Session 1 of 10 | Grade 5/6 Numeracy",
    });

    y = addSectionHeading(doc, "What Is an Equivalence Chain?", y, { color: C.NAVY });
    y = addBodyText(doc, "An equivalence chain is a series of expressions that all have the same value.", y);
    y = addBodyText(doc, "For example, 24 can be written in MANY ways:", y);
    y = addBodyText(doc, "   4 x 6 = 8 x 3 = 12 x 2 = 24 x 1 = 48 / 2 = 72 / 3", y);
    y = addBodyText(doc, "All of these expressions equal 24, so they form an equivalence chain.", y);

    y = addSectionHeading(doc, "Worked Example: Building a Chain for 20", y, { color: C.NAVY });
    y = addBodyText(doc, "Start: 20", y);
    y = addBodyText(doc, "   4 x 5 = 20    (multiplication)", y);
    y = addBodyText(doc, "   2 x 10 = 20   (multiplication)", y);
    y = addBodyText(doc, "   1 x 20 = 20   (multiplication)", y);
    y = addBodyText(doc, "   40 / 2 = 20   (division)", y);
    y = addBodyText(doc, "   60 / 3 = 20   (division)", y);
    y = addBodyText(doc, "   100 / 5 = 20  (division)", y);
    y = addBodyText(doc, "Chain length: 6 expressions! Can you beat this?", y);

    y = addSectionHeading(doc, "Your Investigation", y, { color: C.NAVY });
    y = addBodyText(doc, "Build the LONGEST equivalence chain you can for each number.", y);
    y = addBodyText(doc, "You can use multiplication, division, or both.", y);
    y += 5;

    y = addBodyText(doc, "Target Number: 36", y);
    for (let i = 0; i < 8; i++) {
      y = addWriteLine(doc, "___ x ___ = 36  OR  ___ / ___ = 36", y);
    }
    y += 5;
    y = addBodyText(doc, "Target Number: 60", y);
    for (let i = 0; i < 8; i++) {
      y = addWriteLine(doc, "___ x ___ = 60  OR  ___ / ___ = 60", y);
    }

    y = addSectionHeading(doc, "Extension Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "Which numbers have the LONGEST chains? Why?", y);
    y = addBodyText(doc, "Hint: Think about how many factors a number has.", y);
    y = addWriteLine(doc, "", y);
    y = addWriteLine(doc, "", y);

    y = addTipBox(doc, "Did You Know? Numbers with lots of factors (like 12, 24, 36, 48, 60) are called 'highly composite numbers.' They are useful in everyday life -- that is why we have 12 hours on a clock, 60 minutes in an hour, and 360 degrees in a circle!", y, { color: C.TEAL });

    addPdfFooter(doc, "Session 1 | Extension Investigation | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 1 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
