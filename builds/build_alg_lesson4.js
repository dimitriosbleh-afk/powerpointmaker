"use strict";

// Algebra Unit — Session 4: Finding Unknown Values (Multiplication & Division)
// Week 1 Session 4, Grade 5/6 Numeracy, Variant 0
// DR: Number Patterns with Factors and Multiples
// Fluency: Factor pairs speed drill
// VC2M5A02 — finding unknown values in numerical equations using properties

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

const SESSION = 4;
const FOOTER = "Algebra | Session 4 of 10 | Grade 5/6 Numeracy";
const OUT_DIR = "output/ALG_Session4_Finding_Unknowns";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Independent practice - finding unknown values in equations.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with worked answers.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Step-by-step guided scaffold for finding unknowns.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Self-contained investigation into creating equations with given solutions.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

const NOTES_TITLE = `SAY:
- Session 4. Today we put our properties to work
- We have learned commutative, associative, and distributive. Now we use them to find unknown values
- This is where algebra gets exciting -- you become number detectives

DO:
- Display title slide
- Have whiteboards ready

TEACHER NOTES:
Session 4 of 10. This is the payoff lesson -- students apply the properties learned in Sessions 1-3 to find unknown values in equations. The curriculum elaboration (3 x 5 = 30 / ?, 3 x 4 = ? x 2) is directly addressed.

WATCH FOR:
- Students who remember the properties from Sessions 2-3 -- build confidence

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Daily review. Factors and multiples
- Find the common factors of each pair of numbers
- Write them on your whiteboards

DO:
- Display problems
- Allow 2 minutes
- Watch for systematic listing

TEACHER NOTES:
DR reviews factors and common factors. Finding common factors is directly useful for today's work -- when students see 3 x 4 = ? x 2, recognising that 4 = 2 x 2 (common factor 2) is the key insight.

WATCH FOR:
- Students who list all factors of both numbers then circle the common ones -- systematic approach
- Students who only find one or two common factors -- may miss some

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers
- Ask: What are the common factors of 12 and 18? [1, 2, 3, 6]
- Knowing common factors helps us rewrite expressions -- you will see why today

DO:
- Reveal answers, students self-check

TEACHER NOTES:
Link to today's lesson: common factors enable us to decompose and recompose expressions to find unknowns.

WATCH FOR:
- Students who found all common factors -- strong number sense

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency -- factor pairs. Write all factor pairs for each number as fast as you can
- Go!

DO:
- Display 5 numbers for factor pair listing
- Time 2 minutes
- Self-check

TEACHER NOTES:
Factor pair fluency is essential for today. Students who can quickly identify that 12 = 3x4 = 2x6 = 1x12 will find unknown-solving much easier.

WATCH FOR:
- Students who list systematically vs randomly
- Students who miss pairs -- factor knowledge needs continued fluency work

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me
- Read from slide: "We are learning to find unknown values in multiplication and division equations"
- Today we use everything we have learned so far to find the missing number
- Read our success criteria

DO:
- Choral read LI and SC

TEACHER NOTES:
This lesson targets the core curriculum outcome: finding unknowns in equations like 3 x 5 = 30 / ? and 3 x 4 = ? x 2. The three properties from Sessions 1-3 are the tools.

WATCH FOR:
- Students who are eager to try -- the detective metaphor builds engagement

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me find the unknown. 3 x 5 = 30 / ?
- First, I work out the left side: 3 x 5 = 15
- So the equation becomes: 15 = 30 / ?
- Now I ask: what number divides into 30 to give 15?
- 30 / ? = 15. I know 30 / 2 = 15. So ? = 2.
- Check: 3 x 5 = 15, and 30 / 2 = 15. Both sides equal 15. It works!

DO:
- Write each step clearly
- Circle the left side result (15) and show how it becomes the target for the right side
- Emphasise the "evaluate one side, then solve the other" strategy

TEACHER NOTES:
The strategy is: evaluate the known side, then use the result to find the unknown on the other side. This is the first systematic approach to solving equations.

MISCONCEPTIONS:
- Misconception: Students try to "do something" to both sides simultaneously without first evaluating
  Why: They have seen cross-multiplication or balancing in older contexts and try to apply half-remembered procedures
  Impact: Leads to procedural errors and confusion
  Quick correction: "Start simple. Work out the side you CAN calculate. That tells you what the other side must equal. Then find the missing piece."

WATCH FOR:
- Students who see the answer immediately -- celebrate but insist on showing the method
- Students who get stuck at "what divides 30 to give 15" -- factor knowledge gap

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Second type. 3 x 4 = ? x 2
- Left side: 3 x 4 = 12
- So ? x 2 = 12. What times 2 equals 12? 6.
- But WHY does this work? Let me show you using the associative property
- 3 x 4 = 3 x (2 x 2) = (3 x 2) x 2 = 6 x 2
- So the unknown is 6 because we can split 4 into 2 x 2, then regroup
- This is the associative property in action!

DO:
- Write the evaluation method first (quick solve)
- Then show the property-based reasoning (deep understanding)
- Both give the same answer -- connect them

TEACHER NOTES:
Two paths to the answer: (1) evaluate-and-solve (procedural), (2) property-based reasoning (conceptual). Students need both. The procedural path is fast, the conceptual path builds algebraic understanding for secondary school.

WATCH FOR:
- Students who can do the procedural solve but do not see the property connection -- guide them through the decomposition
- Students who see both paths -- strong algebraic thinking

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. Find the unknown.
- 5 x 6 = ? x 10
- Work out the left side first. Then find what times 10 gives that value.
- Whiteboards. 15 seconds.

DO:
- Students write on whiteboards
- Scan for ? = 3

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "5 x 6 = ? x 10. Work out the left side, then find the unknown. Show me."
- Scan for: 3 (since 5 x 6 = 30, and 30 / 10 = 3, so ? = 3).
PROCEED: If 80%+ show 3, move to We Do.
PIVOT: Most likely misconception -- students write 30 (they evaluated the left side but forgot to divide by 10 to find the unknown). Reteach: "Left side = 30. So ? x 10 = 30. What times 10 equals 30? Not 30 -- that would be 30 x 10 = 300. Think: ? x 10 = 30. So ? = 3."

TEACHER NOTES:
Tests the core skill: evaluate one side, then solve for the unknown on the other side.

WATCH FOR:
- Students who write 30 instead of 3 -- they stopped at evaluation and did not solve
- Students who get 3 instantly -- the strategy is clicking

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- With your partner, find the unknown in each equation
- Show your working: evaluate one side, then solve the other

DO:
- Display four problems
- Partners work on whiteboards
- Circulate, listen for strategy talk
- Allow 3 minutes

TEACHER NOTES:
We Do problems mix multiplication and division unknowns. Problem 4 connects to the associative property explicitly.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 4 Enabling Scaffold which guides students through each step: (1) work out the left side, (2) write "the right side must equal ___", (3) find the missing number. First two problems are completed as models.
- Extra Notes: Distribute the Session 4 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Create equations with given solutions. Given that ? = 8, write three different equations where 8 is the unknown. Use at least two different operations. Then challenge a partner to solve them.
- Extra Notes: Distribute the Session 4 Extension PDF.

WATCH FOR:
- Students who evaluate correctly but make errors finding the unknown -- focus on the "what times/divided by gives" reasoning
- Readiness signal: partners completing all four within 2 minutes

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check. 4 x 8 = ? x 4. Left = 32. ? x 4 = 32. ? = 8.
- 7 x 6 = 84 / ?. Left = 42. 84 / ? = 42. ? = 2.
- 9 x ? = 36. ? = 4 (since 9 x 4 = 36).
- 3 x 4 = ? x 2. Left = 12. ? x 2 = 12. ? = 6. Or: 3 x (2 x 2) = (3 x 2) x 2 = 6 x 2.

DO:
- Reveal one at a time
- For problem 4, show both the procedural and associative approaches
- Cold call for each

TEACHER NOTES:
Problem 4 mirrors the curriculum elaboration exactly. Showing both paths reinforces the connection between procedural solving and algebraic properties.

WATCH FOR:
- Students who can solve procedurally but struggle with the property explanation -- they may need more practice with associative decomposition
- Students who see both paths -- strong conceptual understanding

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. Cold call -- I will pick someone to answer
- Find the unknown: 15 / 3 = 30 / ?
- Think... what does the left side equal? What must the right side equal?

DO:
- Wait 10 seconds for thinking
- Cold call a student
- Follow up with a different student for justification

CFU CHECKPOINT:
Technique: Cold Call
Script:
- Say: "Everyone think. 15 / 3 = 30 / ?. What is the left side? [5] So 30 / ? must equal? [5]. What is ?... [Name], what did you get?"
- Scan for: ? = 6 (since 30 / 6 = 5).
PROCEED: If the cold-called student and 2 follow-ups get 6, move to You Do.
PIVOT: Most likely misconception -- students say ? = 10 (dividing 30 by 3 instead of recognising the left side equals 5). Reteach: "Left side first: 15 / 3 = 5. Not 3 -- five. So we need 30 / ? = 5. What number goes into 30 five times? Count by 5s: 5, 10, 15, 20, 25, 30 -- that is 6 counts. So ? = 6."

TEACHER NOTES:
Division-based unknowns are harder because students must think about what divides into the dividend to give the quotient. This is inverse thinking.

WATCH FOR:
- Students who confuse the divisor with the quotient
- Students who get 6 confidently -- inverse thinking is developing

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice. Your worksheet has equations with unknowns
- First: Work out the known side. Next: Write what the other side must equal. Then: Find the unknown.
- Show all working. 8 minutes.

DO:
- Distribute Session 4 Worksheet
- Students work independently
- Distribute enabling/extending as needed

TEACHER NOTES:
You Do uses different numbers from We Do. Problems progress from single-operation to mixed-operation equations. The challenge section asks students to write their own equations.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the step-by-step scaffold: evaluate, target, solve. Pre-filled models for first two problems.
- Extra Notes: Distribute Session 4 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Create equations with given solutions. Design at least 3 equations for each target value (8 and 15). Use multiple operations.
- Extra Notes: Distribute Session 4 Extension PDF.

WATCH FOR:
- Students who can solve multiplication unknowns but struggle with division -- division inverse thinking needs more support
- Readiness signal: 80%+ completing 8+ problems correctly

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Three questions. 3 minutes. Work alone.

DO:
- Display questions, students work silently, collect to sort

TEACHER NOTES:
Q1 tests multiplication unknown (SC1). Q2 tests division unknown (SC2). Q3 tests justification with properties (SC3).

WATCH FOR:
- Students who solve Q1-Q2 but struggle with Q3 -- they can solve procedurally but not yet explain with properties

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Success criteria check.
- SC1: I can find an unknown in a multiplication equation. Thumbs?
- SC2: I can find an unknown in a division equation. Thumbs?
- SC3: I can explain my reasoning using number properties. Thumbs?
- Turn and talk: What is your strategy for finding an unknown? Explain it step by step to your partner.

DO:
- Thumbs for each SC
- 30 seconds Turn and Talk
- Cold call 2 students

TEACHER NOTES:
The strategy summary (evaluate, target, solve) is the key takeaway. Students who can articulate this process have a transferable problem-solving approach.

WATCH FOR:
- Students who describe the three-step process clearly -- ready for consolidation in Session 5
- Students who are still unsure -- they need the scaffold again in Session 5

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Resources linked here

DO:
- Point out resources

TEACHER NOTES:
Resource slide with companion PDF links.

WATCH FOR:
- N/A

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ───────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, "Algebra: Finding Unknown Values", "Session 4: Finding the Missing Number", "Grade 5/6 Numeracy | Session 4 of 10 | Week 1", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // DR — Number Patterns with Factors and Multiples
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Factors and Common Factors", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const probs = [
        "1.  List common factors of 12 and 18",
        "2.  List common factors of 20 and 30",
        "3.  What is the HCF of 16 and 24?",
        "4.  Is 7 a factor of 42? How do you know?",
      ];
      s.addText(probs.map((p, i) => ({
        text: p, options: { fontSize: 14, color: C.CHARCOAL, breakLine: i < probs.length - 1, paraSpaceAfter: 14 },
      })), { x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3, fontFace: FONT_B, margin: 0, valign: "top" });

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboards:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "List factors of BOTH numbers.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "Circle the ones they share.", options: { fontSize: 14, color: C.CHARCOAL } },
      ], { x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3, fontFace: FONT_B, margin: 0, valign: "top" });

      addFooter(s, FOOTER); s.addNotes(NOTES_DR_Q); return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) 1,2,3,6  2) 1,2,5,10  3) HCF=8  4) Yes (42/7=6)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Fluency — Factor pairs
  const sFl = pres.addSlide();
  addTopBar(sFl, STAGE_COLORS["1"]);
  addStageBadge(sFl, 1, "Fluency");
  addTitle(sFl, "Factor Pairs Speed Round", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

  addCard(sFl, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  sFl.addText([
    "List ALL factor pairs for:",
    "",
    "1.  18",
    "2.  24",
    "3.  30",
    "4.  40",
    "5.  48",
  ].map((p, i) => ({
    text: p, options: { fontSize: 15, color: C.CHARCOAL, breakLine: i < 6, paraSpaceAfter: 6 },
  })), { x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2, fontFace: FONT_B, margin: 0, valign: "top" });

  addCard(sFl, 5.2, CONTENT_TOP, 4.3, 2.2, { strip: C.SECONDARY });
  sFl.addText([
    { text: "Whiteboards - 2 minutes", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
    { text: "Start: 1 x ?, 2 x ?, 3 x ?...", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
    { text: "Stop when factors repeat.", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "18: 1x18,2x9,3x6", options: { fontSize: 10, color: C.MUTED, breakLine: true } },
    { text: "24: 1x24,2x12,3x8,4x6", options: { fontSize: 10, color: C.MUTED, breakLine: true } },
    { text: "30: 1x30,2x15,3x10,5x6", options: { fontSize: 10, color: C.MUTED, breakLine: true } },
    { text: "40: 1x40,2x20,4x10,5x8", options: { fontSize: 10, color: C.MUTED, breakLine: true } },
    { text: "48: 1x48,2x24,3x16,4x12,6x8", options: { fontSize: 10, color: C.MUTED } },
  ], { x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: 2.0, fontFace: FONT_B, margin: 0, valign: "top" });

  addFooter(sFl, FOOTER); sFl.addNotes(NOTES_FLUENCY);

  // LI/SC
  liSlide(pres,
    ["We are learning to find unknown values in multiplication and division equations"],
    [
      "I can find an unknown value in a multiplication equation",
      "I can find an unknown value in a division equation",
      "I can explain my reasoning using number properties",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do 1 — Finding unknown in 3 x 5 = 30 / ?
  workedExSlide(pres, 2, "I Do", "Finding the Unknown: Evaluate and Solve",
    [
      "Find ?: 3 x 5 = 30 / ?",
      "",
      "Step 1: Evaluate the known side",
      "  Left: 3 x 5 = 15",
      "",
      "Step 2: Set up the other side",
      "  15 = 30 / ?",
      "",
      "Step 3: Solve",
      "  What divides into 30 to give 15?",
      "  30 / 2 = 15, so ? = 2",
      "",
      "Check: 3 x 5 = 15  and  30 / 2 = 15  (balanced!)",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.8, { strip: C.SUCCESS });
      slide.addText("The 3-Step Strategy", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 15, fontFace: FONT_H, color: C.SUCCESS, bold: true, margin: 0, align: "center",
      });
      const steps = [
        { num: "1", label: "EVALUATE", desc: "the known side" },
        { num: "2", label: "TARGET", desc: "what the other side must equal" },
        { num: "3", label: "SOLVE", desc: "find the missing number" },
      ];
      steps.forEach((st, i) => {
        const sy = lg.panelTopPadded + 0.42 + i * 0.42;
        addTextOnShape(slide, st.num, {
          x: lg.rightX + 0.2, y: sy, w: 0.35, h: 0.32, rectRadius: 0.06,
          fill: { color: C.PRIMARY },
        }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(st.label + ": " + st.desc, {
          x: lg.rightX + 0.65, y: sy, w: lg.rightW - 0.85, h: 0.32,
          fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "middle",
        });
      });
    }
  );

  // I Do 2 — Finding unknown using associative property
  workedExSlide(pres, 2, "I Do", "Using Properties to Find Unknowns",
    [
      "Find ?: 3 x 4 = ? x 2",
      "",
      "Quick solve:",
      "  Left: 3 x 4 = 12",
      "  ? x 2 = 12, so ? = 6",
      "",
      "Why it works (associative property):",
      "  3 x 4 = 3 x (2 x 2)",
      "        = (3 x 2) x 2",
      "        = 6 x 2",
      "  So ? = 6",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.5, { strip: C.PRIMARY });
      slide.addText("Procedural Path", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
      });
      slide.addText("3 x 4 = 12\n? x 2 = 12\n? = 12 / 2 = 6", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.35, w: lg.rightW - 0.6, h: 0.85,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });

      addCard(slide, lg.rightX, lg.panelTopPadded + 1.7, lg.rightW, 1.5, { strip: C.SECONDARY });
      slide.addText("Property Path (Associative)", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.76, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
      });
      slide.addText("3 x 4 = 3 x (2 x 2)\n      = (3 x 2) x 2\n      = 6 x 2\nSo ? = 6", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 2.05, w: lg.rightW - 0.6, h: 0.85,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });
    }
  );

  // CFU 1
  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check", "Show Me Boards",
      "Find the unknown:\n\n5 x 6 = ? x 10",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "5 x 6 = 30    ? x 10 = 30    ? = 3", {
        x: 1.0, y: 4.0, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // We Do
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Find the Unknown — With a Partner",
      [
        "Find the unknown in each equation.",
        "Show: evaluate, target, solve.",
        "",
        "1.   4 x 8 = ? x 4",
        "2.   7 x 6 = 84 / ?",
        "3.   9 x ? = 36",
        "4.   3 x 4 = ? x 2",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.0, { strip: C.SECONDARY });
        slide.addText("Strategy:", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        slide.addText([
          { text: "1. EVALUATE the known side", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "2. TARGET: what must the other side equal?", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "3. SOLVE for the unknown", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "4. CHECK: do both sides balance?", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
        ], { x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.38, w: lg.rightW - 0.5, h: 1.2, fontFace: FONT_B, margin: 0, valign: "top" });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "1) ?=8   2) ?=2   3) ?=4   4) ?=6", {
        x: 1.5, y: 4.55, w: 7, h: 0.5, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // CFU 2 Hinge
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Check", "Cold Call",
      "Find the unknown:\n\n15 / 3 = 30 / ?",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "15 / 3 = 5    30 / ? = 5    ? = 6", {
        x: 1.0, y: 4.0, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // You Do
  workedExSlide(pres, 4, "You Do", "Independent Practice",
    [
      "First: Read each equation.",
      "Next: Evaluate the known side.",
      "Then: Find the unknown value.",
      "",
      "Show all working.",
      "Challenge: write your own equations!",
      "",
      "You have 8 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.6, { strip: C.ALERT });
      slide.addText("Remember:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "EVALUATE -> TARGET -> SOLVE", options: { fontSize: 14, color: C.PRIMARY, bold: true, breakLine: true } },
        { text: "Always CHECK both sides balance!", options: { fontSize: 13, color: C.CHARCOAL } },
      ], { x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.4, w: lg.rightW - 0.5, h: 0.8, fontFace: FONT_B, margin: 0, valign: "top" });
    }
  );

  exitTicketSlide(pres,
    [
      "Find ?: 6 x 7 = ? x 3  (show working)",
      "Find ?: 48 / 6 = 24 / ?  (show working)",
      "Explain WHY 5 x 4 = 10 x 2 using a number property. Name the property.",
    ],
    NOTES_EXIT, FOOTER);

  closingSlide(pres,
    "What are the three steps for finding an unknown? Explain them to your partner.",
    [
      "I can find an unknown in a multiplication equation",
      "I can find an unknown in a division equation",
      "I can explain my reasoning using number properties",
    ],
    NOTES_CLOSING);



  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "ALG_Session4_Finding_Unknowns.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // PDFs
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, { subtitle: "Finding Unknown Values", color: C.NAVY, lessonInfo: "Session 4 of 10 | Grade 5/6 Numeracy" });
    y = addTipBox(doc, "Strategy: EVALUATE the known side, set your TARGET, SOLVE for the unknown, CHECK both sides balance.", y, { color: C.TEAL });
    y = addSectionHeading(doc, "Section A: Multiplication Unknowns", y, { color: C.NAVY });
    y = addProblem(doc, 1, "8 x 3 = ? x 6          ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 2, "5 x 8 = ? x 4          ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 3, "6 x 9 = ? x 3          ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 4, "7 x 4 = ? x 2          ? = ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "Section B: Division Unknowns", y, { color: C.NAVY });
    y = addProblem(doc, 5, "4 x 5 = 40 / ?         ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 6, "6 x 3 = 36 / ?         ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 7, "24 / 4 = 30 / ?        ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 8, "8 x 5 = 80 / ?         ? = ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "Section C: Mixed", y, { color: C.NAVY });
    y = addProblem(doc, 9, "9 x ? = 72             ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 10, "? x 7 = 56            ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 11, "5 x 6 = ? x 3         ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 12, "48 / ? = 24 / 3        ? = ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "Challenge: Create Your Own", y, { color: C.NAVY });
    y = addBodyText(doc, "13. Write an equation where the unknown value is 9:", y);
    y = addWriteLine(doc, "", y);
    addPdfFooter(doc, "Session 4 | Algebra | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, { subtitle: "Finding Unknown Values", color: C.NAVY, lessonInfo: "Session 4 of 10 | Grade 5/6" });
    y = addSectionHeading(doc, "Section A", y, { color: C.NAVY });
    y = addBodyText(doc, "1. 8x3=24, ?x6=24, ?=4", y);
    y = addBodyText(doc, "2. 5x8=40, ?x4=40, ?=10", y);
    y = addBodyText(doc, "3. 6x9=54, ?x3=54, ?=18", y);
    y = addBodyText(doc, "4. 7x4=28, ?x2=28, ?=14", y);
    y = addSectionHeading(doc, "Section B", y, { color: C.NAVY });
    y = addBodyText(doc, "5. 4x5=20, 40/?=20, ?=2", y);
    y = addBodyText(doc, "6. 6x3=18, 36/?=18, ?=2", y);
    y = addBodyText(doc, "7. 24/4=6, 30/?=6, ?=5", y);
    y = addBodyText(doc, "8. 8x5=40, 80/?=40, ?=2", y);
    y = addSectionHeading(doc, "Section C", y, { color: C.NAVY });
    y = addBodyText(doc, "9. 9x?=72, ?=8", y);
    y = addBodyText(doc, "10. ?x7=56, ?=8", y);
    y = addBodyText(doc, "11. 5x6=30, ?x3=30, ?=10", y);
    y = addBodyText(doc, "12. 24/?=8 (since 24/3=8), 48/?=8, ?=6", y);
    addPdfFooter(doc, "Session 4 | Answer Key | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF: " + ANSWER_KEY_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, { subtitle: "Step-by-Step Unknown Finder", color: C.TEAL, lessonInfo: "Session 4 of 10 | Grade 5/6" });
    y = addTipBox(doc, "Follow the three steps for each equation: EVALUATE the known side, write the TARGET, SOLVE for the unknown.", y, { color: C.TEAL });
    y = addSectionHeading(doc, "Example (done for you)", y, { color: C.NAVY });
    y = addBodyText(doc, "Equation: 4 x 5 = ? x 10", y);
    y = addBodyText(doc, "Step 1 EVALUATE: 4 x 5 = 20", y);
    y = addBodyText(doc, "Step 2 TARGET: ? x 10 must equal 20", y);
    y = addBodyText(doc, "Step 3 SOLVE: What times 10 = 20? Answer: ? = 2", y);
    y = addBodyText(doc, "CHECK: 4 x 5 = 20 and 2 x 10 = 20. Balanced!", y);
    y += 10;
    y = addSectionHeading(doc, "Example 2 (done for you)", y, { color: C.NAVY });
    y = addBodyText(doc, "Equation: 3 x 6 = 36 / ?", y);
    y = addBodyText(doc, "Step 1 EVALUATE: 3 x 6 = 18", y);
    y = addBodyText(doc, "Step 2 TARGET: 36 / ? must equal 18", y);
    y = addBodyText(doc, "Step 3 SOLVE: 36 divided by what = 18? Answer: ? = 2", y);
    y += 10;
    y = addSectionHeading(doc, "Now you try:", y, { color: C.NAVY });
    const probs = ["8 x 3 = ? x 6", "5 x 4 = 40 / ?", "6 x 5 = ? x 3", "7 x 6 = 84 / ?"];
    probs.forEach((p, i) => {
      y = addBodyText(doc, (i + 1) + ". Equation: " + p, y);
      y = addBodyText(doc, "   Step 1 EVALUATE: ___", y);
      y = addBodyText(doc, "   Step 2 TARGET: ___ must equal ___", y);
      y = addBodyText(doc, "   Step 3 SOLVE: ? = ___", y);
      y += 8;
    });
    addPdfFooter(doc, "Session 4 | Enabling Scaffold | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF: " + ENABLING_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, { subtitle: "Equation Creator Challenge", color: C.NAVY, lessonInfo: "Session 4 of 10 | Grade 5/6" });
    y = addSectionHeading(doc, "The Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "Instead of SOLVING equations, you will CREATE them!", y);
    y = addBodyText(doc, "You know the answer. Your job is to build equations that have that answer.", y);
    y = addSectionHeading(doc, "Worked Example: Answer = 5", y, { color: C.NAVY });
    y = addBodyText(doc, "Equation 1: 3 x 5 = 30 / ? (because 30/5=6... wait, that gives ?=6)", y);
    y = addBodyText(doc, "Let me try: 2 x 5 = ? x 2. Left=10, ?x2=10, ?=5. YES!", y);
    y = addBodyText(doc, "Equation 2: 20 / ? = 4. Since 20/5=4. YES!", y);
    y = addBodyText(doc, "Equation 3: ? x 8 = 40. Since 5x8=40. YES!", y);
    y = addSectionHeading(doc, "Your Turn", y, { color: C.NAVY });
    y = addBodyText(doc, "Target answer: ? = 8. Create 3 equations:", y);
    y = addWriteLine(doc, "Equation 1:", y);
    y = addWriteLine(doc, "Equation 2:", y);
    y = addWriteLine(doc, "Equation 3:", y);
    y += 5;
    y = addBodyText(doc, "Target answer: ? = 15. Create 3 equations:", y);
    y = addWriteLine(doc, "Equation 1:", y);
    y = addWriteLine(doc, "Equation 2:", y);
    y = addWriteLine(doc, "Equation 3:", y);
    y += 5;
    y = addSectionHeading(doc, "Ultimate Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "Create an equation that uses ALL FOUR operations (+, -, x, /) where ? = 12.", y);
    y = addWriteLine(doc, "", y);
    y = addTipBox(doc, "Tip: Start with a multiplication fact you know (e.g., 8 x 12 = 96), then rearrange it into an equation with a ? in place of 12.", y, { color: C.TEAL });
    addPdfFooter(doc, "Session 4 | Extension | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 4 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
