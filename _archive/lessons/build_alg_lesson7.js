"use strict";

// Algebra Unit — Session 7: Equivalent Sentences with Brackets
// Week 2 Session 2, Grade 5/6 Numeracy, Variant 0
// DR: Solving Equations with Multiplication, Division, and Operations
// Fluency: Brackets simplification
// VC2M6A02 — constructing equivalent number sentences involving brackets

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
const { C, FONT_H, FONT_B, SAFE_BOTTOM, CONTENT_TOP, titleSlide, liSlide, cfuSlide, closingSlide, workedExSlide, exitTicketSlide, addStageBadge, addTextOnShape, addCard, addFooter, addTopBar, addTitle, withReveal, STAGE_COLORS } = T;

const SESSION = 7;
const FOOTER = "Algebra | Session 7 of 10 | Grade 5/6 Numeracy";
const OUT_DIR = "output/ALG_Session7_Equivalent_with_Brackets";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Practice constructing equivalent sentences with brackets.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Guided equivalence checker with brackets template.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Investigation into nested brackets.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

const NOTES_TITLE = `SAY:
- Session 7. Yesterday we learned the order of operations and brackets
- Today we combine that with equivalence. We will build equivalent number sentences that USE brackets
- This is where Week 1 and Week 2 skills come together

DO:
- Display title slide

TEACHER NOTES:
Session 7 of 10. Students construct equivalent number sentences that include brackets and multiple operations. This combines Week 1 equivalence skills with Week 2 order of operations understanding.

WATCH FOR:
- Students who are confident with order of operations from Session 6

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Daily review. Find the unknowns -- these use multiplication, division, and operations from last week
- Evaluate, target, solve

DO:
- Display problems, 90 seconds

TEACHER NOTES:
DR retrieves equation-solving with single operations, preparing for multi-operation equations.

WATCH FOR:
- Speed and confidence with the evaluate-target-solve strategy

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers

DO:
- Reveal, students self-check

TEACHER NOTES:
Quick tick-and-fix.

WATCH FOR:
- Any persistent errors from Week 1

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency -- simplify each expression. Watch the brackets!
- Go!

DO:
- Display 8 bracket expressions
- Time 2 minutes

TEACHER NOTES:
Fluency reinforces brackets-first from Session 6.

WATCH FOR:
- Students who process brackets confidently

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention
- Read from slide: "We are learning to construct equivalent number sentences that include brackets and multiple operations"
- This means building equations where both sides have the same value, even though they look different
- Read our success criteria

DO:
- Choral read

TEACHER NOTES:
Constructing equivalent sentences (not just verifying them) is the core Level 6 skill. Students must produce, not just check.

WATCH FOR:
- Students who remember equivalence from Session 1

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me build equivalent sentences with brackets
- Start with: 3 x (4 + 6) = ?
- Brackets first: 4 + 6 = 10. Then 3 x 10 = 30.
- Now I need another expression that equals 30.
- How about: 6 x (2 + 3) = 6 x 5 = 30? Yes!
- So 3 x (4 + 6) = 6 x (2 + 3). Both equal 30. Equivalent!
- I can also write: 3 x (4 + 6) = 5 x 6. No brackets needed on the right if it equals 30.

DO:
- Write each step
- Show how to construct the second expression by finding a target value then building toward it

TEACHER NOTES:
The construction strategy: evaluate one side to find the target value, then build a different expression that reaches the same value. This reverses the verification process from Session 1.

WATCH FOR:
- Students who can suggest alternative expressions for 30
- Readiness signal: students offering expressions before you write them

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Another way to construct equivalence: use the distributive property WITH brackets
- 4 x 13 = 4 x (10 + 3). These are equivalent because distributive property.
- I can also write: 4 x 13 = 2 x (10 + 16). Check: 2 x 26 = 52. And 4 x 13 = 52. TRUE!
- Or: 4 x 13 = (8 + 5) x 4. Brackets: 8 + 5 = 13. Then 13 x 4 = 52. TRUE!
- Multiple valid equivalent sentences exist for every value. Your job is to construct them.

DO:
- Show multiple equivalent constructions
- Emphasise that there are many correct answers

TEACHER NOTES:
This I Do demonstrates the creative nature of constructing equivalences. Students need to see that there are infinite valid answers, not just one "right" one.

WATCH FOR:
- Students who offer creative constructions -- celebrate variety

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. On your whiteboard, write an expression with brackets that equals 24
- It must use at least two different operations. 15 seconds.

DO:
- Students write
- Scan for valid expressions (e.g., 4 x (3 + 3), (12 + 12), 3 x (6 + 2), (20 + 4))

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Write an expression WITH brackets that equals 24. Use at least two operations. Show me."
- Scan for: any valid expression equalling 24 with brackets (e.g., 4 x (3+3)=24, (4+8) x 2=24).
PROCEED: If 80%+ show a valid expression, continue.
PIVOT: Most likely error -- students write just "24" or "4 x 6" (no brackets). Reteach: "I need brackets in your expression. Think: what addition or subtraction could I put inside brackets? Like (3 + 3) = 6, and then 4 x (3 + 3) = 4 x 6 = 24."

TEACHER NOTES:
This checks construction ability, not just evaluation. Students must produce, not just verify.

WATCH FOR:
- Students who create valid but simple expressions -- push for variety
- Students who create complex expressions -- strong algebraic thinking

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- With your partner, construct equivalent sentences
- For each value, write TWO different expressions using brackets

DO:
- Display target values
- Partners work on whiteboards
- 3 minutes

TEACHER NOTES:
We Do requires construction of multiple equivalent expressions. This is harder than verification.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 7 Enabling Scaffold which provides one expression for each target value. Students evaluate it to confirm the value, then build ONE more equivalent expression using the guided template.
- Extra Notes: Distribute Session 7 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Investigate nested brackets -- expressions with brackets inside brackets, like 2 x ((3 + 4) x 2). Evaluate given nested expressions and create your own.
- Extra Notes: Distribute Session 7 Extension PDF.

WATCH FOR:
- Students who create valid expressions but all use the same pattern -- push for variety
- Readiness signal: partners creating 2 different expressions for each target

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's share. Target value 40: who has an expression?
- Accept multiple valid answers. Examples: 5 x (4+4), (10+10) x 2, 8 x (3+2), 4 x (7+3)
- Target value 36: examples: 6 x (3+3), 4 x (5+4), (12+6) x 2, 9 x (2+2)

DO:
- Cold call for each target
- Validate all correct answers
- Ask students to verify a classmate's expression

TEACHER NOTES:
Sharing and verifying classmates' expressions builds mathematical communication skills.

WATCH FOR:
- Creative and unusual expressions -- celebrate these

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. TRUE or FALSE?
- (5 + 3) x 6 = 5 x 6 + 3 x 6
- Thumbs up for true, down for false.

DO:
- Students show thumbs
- Both = 48. TRUE! This is the distributive property with brackets.

CFU CHECKPOINT:
Technique: Thumbs Up/Down
Script:
- Say: "(5 + 3) x 6 = 5 x 6 + 3 x 6. True or false? Thumbs."
- Scan for: thumbs up (both sides = 48).
PROCEED: If 80%+ thumbs up, move to You Do.
PIVOT: Students show thumbs down because the expressions look different. Reteach: "Left: (5+3)x6 = 8x6 = 48. Right: 5x6+3x6 = 30+18 = 48. Same value! This is actually the distributive property -- you saw this in Session 3."

TEACHER NOTES:
This hinge connects brackets-based equivalence to the distributive property from Week 1. If students see the connection, their understanding is cohesive.

WATCH FOR:
- Students who recognise the distributive property -- they are connecting Week 1 and Week 2

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice
- Your worksheet has three sections: verify equivalence, construct equivalence, and place brackets
- Show all working. 8 minutes.

DO:
- Distribute worksheet, circulate

TEACHER NOTES:
You Do uses different values from We Do.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Guided template for constructing equivalences.
- Extra Notes: Distribute Session 7 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Nested brackets investigation.
- Extra Notes: Distribute Session 7 Extension PDF.

WATCH FOR:
- Students who verify but struggle to construct -- construction is harder
- Readiness signal: completing Section B with valid constructions

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. 3 questions. 3 minutes.

DO:
- Display questions, collect

TEACHER NOTES:
Q1: verify equivalence with brackets. Q2: construct an equivalent. Q3: explain using properties.

WATCH FOR:
- Students who can construct but not explain the property

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- SC check. Thumbs for each.
- SC1: I can verify if two expressions with brackets are equivalent. Thumbs?
- SC2: I can construct my own equivalent expression using brackets. Thumbs?
- SC3: I can use the distributive property to explain equivalence. Thumbs?
- Turn and talk: How many different ways can you write an expression that equals 48?

DO:
- Thumbs, Turn and Talk, cold call

TEACHER NOTES:
The Turn and Talk connects to the infinite nature of equivalent expressions. Every value has infinitely many representations.

WATCH FOR:
- Students who list many expressions -- strong algebraic fluency

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Resources linked

DO:
- Point out resources

TEACHER NOTES:
Resource slide.

WATCH FOR:
- N/A

[General: Resources | VTLM 2.0: Planning]`;

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, "Algebra: Finding Unknown Values", "Session 7: Equivalent Sentences with Brackets", "Grade 5/6 Numeracy | Session 7 of 10 | Week 2", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]); addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Equation Solving Review", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      s.addText(["1.  3 x 8 = ? x 6", "2.  45 / ? = 9", "3.  ? x 7 = 63", "4.  (3 + 5) x ? = 40"].map((p, i) => ({
        text: p, options: { fontSize: 15, color: C.CHARCOAL, breakLine: i < 3, paraSpaceAfter: 14 },
      })), { x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3, fontFace: FONT_B, margin: 0, valign: "top" });
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "Evaluate - Target - Solve", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "Remember brackets first!", options: { fontSize: 14, color: C.CHARCOAL } },
      ], { x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3, fontFace: FONT_B, margin: 0, valign: "top" });
      addFooter(s, FOOTER); s.addNotes(NOTES_DR_Q); return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) ?=4  2) ?=5  3) ?=9  4) ?=5", {
        x: 1.5, y: 4.55, w: 7, h: 0.5, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Fluency
  const sFl = pres.addSlide();
  addTopBar(sFl, STAGE_COLORS["1"]); addStageBadge(sFl, 1, "Fluency");
  addTitle(sFl, "Brackets Simplification", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });
  addCard(sFl, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  sFl.addText(["1.  (3 + 7) x 4 =", "2.  5 x (8 - 3) =", "3.  (12 + 6) / 3 =", "4.  (9 - 4) x 8 =", "5.  7 x (2 + 5) =", "6.  (15 - 5) / 2 =", "7.  3 x (11 - 4) =", "8.  (6 + 6) x 3 ="].map((p, i) => ({
    text: p, options: { fontSize: 14, color: C.CHARCOAL, breakLine: i < 7, paraSpaceAfter: 4 },
  })), { x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2, fontFace: FONT_B, margin: 0, valign: "top" });
  addCard(sFl, 5.2, CONTENT_TOP, 4.3, 1.6, { strip: C.SECONDARY });
  sFl.addText([
    { text: "Whiteboards - 2 minutes", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
    { text: "Brackets FIRST, always.", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "40, 25, 6, 40, 49, 5, 21, 36", options: { fontSize: 11, color: C.MUTED } },
  ], { x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: 1.4, fontFace: FONT_B, margin: 0, valign: "top" });
  addFooter(sFl, FOOTER); sFl.addNotes(NOTES_FLUENCY);

  liSlide(pres,
    ["We are learning to construct equivalent number sentences that include brackets and multiple operations"],
    [
      "I can verify if two expressions with brackets are equivalent",
      "I can construct my own equivalent expression using brackets",
      "I can use the distributive property to explain bracket equivalence",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do 1
  workedExSlide(pres, 2, "I Do", "Building Equivalence with Brackets",
    [
      "3 x (4 + 6) = ?",
      "  Brackets: 4 + 6 = 10",
      "  3 x 10 = 30. Target value: 30.",
      "",
      "Now build another expression = 30:",
      "  6 x (2 + 3) = 6 x 5 = 30",
      "",
      "So: 3 x (4 + 6) = 6 x (2 + 3)",
      "Both equal 30. Equivalent!",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.8, { strip: C.PRIMARY });
      slide.addText("Construction Strategy", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      const steps = [
        { num: "1", label: "Evaluate one side to find the target value" },
        { num: "2", label: "Build a NEW expression with brackets" },
        { num: "3", label: "Check it equals the target" },
      ];
      steps.forEach((st, i) => {
        const sy = lg.panelTopPadded + 0.42 + i * 0.45;
        addTextOnShape(slide, st.num, {
          x: lg.rightX + 0.2, y: sy, w: 0.35, h: 0.32, rectRadius: 0.06, fill: { color: C.PRIMARY },
        }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(st.label, {
          x: lg.rightX + 0.65, y: sy, w: lg.rightW - 0.85, h: 0.32,
          fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "middle",
        });
      });
      // Examples
      addTextOnShape(slide, "30 = 5 x (3+3) = (10+5) x 2 = 6 x (2+3)", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.9, w: lg.rightW - 0.3, h: 0.4, rectRadius: 0.06, fill: { color: C.SUCCESS },
      }, { fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // I Do 2
  workedExSlide(pres, 2, "I Do", "Distributive Property with Brackets",
    [
      "The distributive property creates equivalence:",
      "",
      "4 x 13 = 4 x (10 + 3)",
      "       = 4 x 10 + 4 x 3",
      "       = 40 + 12 = 52",
      "",
      "So: 4 x 13 = 4 x (10 + 3) = 4 x 10 + 4 x 3",
      "All equivalent!",
      "",
      "Multiple valid constructions exist for every value.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.0, { strip: C.SECONDARY });
      slide.addText("Distributive: Creates Equivalence", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      slide.addText("a x (b + c) = a x b + a x c\n\n4 x (10 + 3) = 4 x 10 + 4 x 3\n             = 40 + 12 = 52", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.6, h: 1.2,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });
    }
  );

  // CFU 1
  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check — Construct", "Show Me Boards",
      "Write an expression WITH brackets that equals 24.\n\nUse at least two different operations.",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Examples: 4 x (3+3)   (12+12)   3 x (6+2)   (4+8) x 2", {
        x: 0.5, y: 4.0, w: 9, h: 0.55, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // We Do
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Construct Equivalent Expressions",
      [
        "For each target value, write TWO different",
        "expressions using brackets.",
        "",
        "Target: 40",
        "  Expression 1: ___",
        "  Expression 2: ___",
        "",
        "Target: 36",
        "  Expression 1: ___",
        "  Expression 2: ___",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.8, { strip: C.SECONDARY });
        slide.addText("Tips:", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        slide.addText([
          { text: "Find factors of the target", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Split one factor using brackets", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Check: does it equal the target?", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
        ], { x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.38, w: lg.rightW - 0.5, h: 0.8, fontFace: FONT_B, margin: 0, valign: "top" });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "40: 5x(4+4), 8x(3+2), 4x(7+3)  |  36: 6x(3+3), 4x(5+4), 9x(2+2)", {
        x: 0.3, y: 4.55, w: 9.4, h: 0.5, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 12, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // CFU 2 Hinge
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge — Distributive Connection", "Thumbs Up/Down",
      "TRUE or FALSE?\n\n(5 + 3) x 6 = 5 x 6 + 3 x 6",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "TRUE: (5+3)x6 = 8x6 = 48    5x6+3x6 = 30+18 = 48", {
        x: 0.5, y: 4.0, w: 9, h: 0.55, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // You Do
  workedExSlide(pres, 4, "You Do", "Independent Practice",
    [
      "First: Verify the given equivalences (Section A).",
      "Next: Construct your own equivalences (Section B).",
      "Then: Place brackets to make equations true (Section C).",
      "",
      "Show all working. 8 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.4, { strip: C.ALERT });
      slide.addText("Remember:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Evaluate each side separately", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Brackets first, always", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Many correct answers exist!", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], { x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.38, w: lg.rightW - 0.5, h: 0.8, fontFace: FONT_B, margin: 0, valign: "top" });
    }
  );

  exitTicketSlide(pres,
    [
      "TRUE or FALSE: 4 x (5 + 3) = 4 x 5 + 4 x 3. Show working.",
      "Write an expression with brackets that equals 60.",
      "Explain why (2 + 3) x 4 = 2 x 4 + 3 x 4 using the distributive property.",
    ],
    NOTES_EXIT, FOOTER);

  closingSlide(pres,
    "How many different ways can you write an expression that equals 48? Share with your partner.",
    [
      "I can verify bracket equivalence",
      "I can construct equivalent expressions with brackets",
      "I can explain bracket equivalence using the distributive property",
    ],
    NOTES_CLOSING);



  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "ALG_Session7_Equivalent_with_Brackets.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // PDFs
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, { subtitle: "Equivalent Sentences with Brackets", color: C.NAVY, lessonInfo: "Session 7 of 10 | Grade 5/6" });
    y = addTipBox(doc, "Section A: verify. Section B: construct. Section C: place brackets.", y, { color: C.TEAL });
    y = addSectionHeading(doc, "A: TRUE or FALSE?", y, { color: C.NAVY });
    y = addProblem(doc, 1, "3 x (5 + 7) = 3 x 5 + 3 x 7            TRUE / FALSE", y, { color: C.NAVY });
    y = addProblem(doc, 2, "(4 + 6) x 5 = 4 x 5 + 6 x 5            TRUE / FALSE", y, { color: C.NAVY });
    y = addProblem(doc, 3, "8 x (3 + 2) = 8 x 3 + 2                TRUE / FALSE", y, { color: C.NAVY });
    y = addProblem(doc, 4, "7 x (10 - 3) = 7 x 10 - 7 x 3          TRUE / FALSE", y, { color: C.NAVY });
    y = addSectionHeading(doc, "B: Construct Equivalences", y, { color: C.NAVY });
    y = addProblem(doc, 5, "Write TWO expressions with brackets that equal 48:", y, { color: C.NAVY });
    y = addWriteLine(doc, "Expression 1:", y);
    y = addWriteLine(doc, "Expression 2:", y);
    y = addProblem(doc, 6, "Write TWO expressions with brackets that equal 72:", y, { color: C.NAVY });
    y = addWriteLine(doc, "Expression 1:", y);
    y = addWriteLine(doc, "Expression 2:", y);
    y = addSectionHeading(doc, "C: Place the Brackets", y, { color: C.NAVY });
    y = addProblem(doc, 7, "3 + 5 x 4 = 32  (add brackets to make true)", y, { color: C.NAVY });
    y = addProblem(doc, 8, "6 x 2 + 3 = 30  (add brackets to make true)", y, { color: C.NAVY });
    y = addProblem(doc, 9, "10 - 2 x 4 = 32 (add brackets to make true)", y, { color: C.NAVY });
    y = addProblem(doc, 10, "Write your own: create an expression where brackets change the answer", y, { color: C.NAVY });
    y = addWriteLine(doc, "Without brackets:", y);
    y = addWriteLine(doc, "With brackets:", y);
    addPdfFooter(doc, "Session 7 | Algebra | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, { subtitle: "Equivalent Sentences with Brackets", color: C.NAVY, lessonInfo: "Session 7" });
    y = addSectionHeading(doc, "A", y, { color: C.NAVY });
    y = addBodyText(doc, "1. TRUE: 3x12=36, 15+21=36", y);
    y = addBodyText(doc, "2. TRUE: 10x5=50, 20+30=50", y);
    y = addBodyText(doc, "3. FALSE: 8x5=40, but 24+2=26", y);
    y = addBodyText(doc, "4. TRUE: 7x7=49, 70-21=49", y);
    y = addSectionHeading(doc, "B", y, { color: C.NAVY });
    y = addBodyText(doc, "5. 48: e.g. 6x(4+4)=48, 8x(3+3)=48", y);
    y = addBodyText(doc, "6. 72: e.g. 8x(5+4)=72, 9x(4+4)=72", y);
    y = addSectionHeading(doc, "C", y, { color: C.NAVY });
    y = addBodyText(doc, "7. (3+5)x4 = 8x4 = 32", y);
    y = addBodyText(doc, "8. 6x(2+3) = 6x5 = 30", y);
    y = addBodyText(doc, "9. (10-2)x4 = 8x4 = 32", y);
    addPdfFooter(doc, "Session 7 | Answer Key | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF: " + ANSWER_KEY_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, { subtitle: "Guided Equivalence Builder", color: C.TEAL, lessonInfo: "Session 7" });
    y = addTipBox(doc, "For each target: check the given expression equals the target, then build your own.", y, { color: C.TEAL });
    y = addSectionHeading(doc, "Example: Target = 30", y, { color: C.NAVY });
    y = addBodyText(doc, "Given: 5 x (3 + 3). Brackets: 3+3=6. Then 5x6=30. YES!", y);
    y = addBodyText(doc, "My expression: ___ x (___ + ___) = ___. Does it equal 30? ___", y);
    y += 10;
    y = addSectionHeading(doc, "Target = 40", y, { color: C.NAVY });
    y = addBodyText(doc, "Given: 8 x (2 + 3). Check: brackets=___, then 8x___=___. Equals 40? ___", y);
    y = addBodyText(doc, "My expression: ___ x (___ + ___) = ___", y);
    y += 10;
    y = addSectionHeading(doc, "Target = 36", y, { color: C.NAVY });
    y = addBodyText(doc, "Given: 4 x (5 + 4). Check: brackets=___, then 4x___=___. Equals 36? ___", y);
    y = addBodyText(doc, "My expression: ___ x (___ + ___) = ___", y);
    y += 10;
    y = addSectionHeading(doc, "Target = 48", y, { color: C.NAVY });
    y = addBodyText(doc, "Given: 6 x (4 + 4). Check: brackets=___, then 6x___=___. Equals 48? ___", y);
    y = addBodyText(doc, "My expression: ___ x (___ + ___) = ___", y);
    addPdfFooter(doc, "Session 7 | Enabling Scaffold | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF: " + ENABLING_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, { subtitle: "Nested Brackets Investigation", color: C.NAVY, lessonInfo: "Session 7" });
    y = addSectionHeading(doc, "What Are Nested Brackets?", y, { color: C.NAVY });
    y = addBodyText(doc, "Nested brackets are brackets inside brackets: 2 x ((3 + 4) x 2)", y);
    y = addBodyText(doc, "Rule: work from the INSIDE out. Solve the innermost brackets first.", y);
    y = addSectionHeading(doc, "Worked Example", y, { color: C.NAVY });
    y = addBodyText(doc, "2 x ((3 + 4) x 2)", y);
    y = addBodyText(doc, "Inner brackets: (3+4) = 7", y);
    y = addBodyText(doc, "Outer brackets: (7 x 2) = 14", y);
    y = addBodyText(doc, "Finally: 2 x 14 = 28", y);
    y = addSectionHeading(doc, "Evaluate These", y, { color: C.NAVY });
    y = addProblem(doc, 1, "3 x ((2 + 5) x 2) =", y, { color: C.NAVY });
    y = addProblem(doc, 2, "((8 - 3) x 4) + 10 =", y, { color: C.NAVY });
    y = addProblem(doc, 3, "((6 + 6) / 3) x 5 =", y, { color: C.NAVY });
    y = addProblem(doc, 4, "100 / ((3 + 2) x 4) =", y, { color: C.NAVY });
    y = addSectionHeading(doc, "Create Your Own", y, { color: C.NAVY });
    y = addBodyText(doc, "Write a nested bracket expression that equals 100:", y);
    y = addWriteLine(doc, "", y);
    y = addBodyText(doc, "Write a nested bracket expression that equals 1:", y);
    y = addWriteLine(doc, "", y);
    addPdfFooter(doc, "Session 7 | Extension | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 7 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
