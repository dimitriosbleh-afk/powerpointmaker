"use strict";

// Algebra Unit — Session 2: Commutative & Associative Properties
// Week 1 Session 2, Grade 5/6 Numeracy, Variant 0
// DR: Exploring Mathematical Patterns and Algorithms
// Fluency: Division facts speed drill
// VC2M5A02 — using arrays to demonstrate multiplication is commutative and associative

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
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

const SESSION = 2;
const FOOTER = "Algebra | Session 2 of 10 | Grade 5/6 Numeracy";
const OUT_DIR = "output/ALG_Session2_Commutative_Associative";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Independent practice - commutative and associative properties with arrays.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with worked answers.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Pre-drawn arrays for commutative property exploration.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Self-contained investigation into why division is not commutative.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ───────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome back to our algebra unit
- Yesterday we learned that the equals sign means "same value" -- like a balance
- Today we are going to discover two powerful properties that help us build equivalent number sentences

DO:
- Display title slide as students settle
- Have whiteboards and markers ready

TEACHER NOTES:
Session 2 of 10. This lesson introduces the commutative and associative properties of multiplication. These properties are the tools students will use to construct equivalent number sentences and find unknowns in later sessions.

WATCH FOR:
- Students who remember the balance model from Session 1 -- build on this

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Daily review time. More patterns and algorithms from last week
- Find the missing number and identify the rule
- Write your answers on your whiteboards

DO:
- Display the four pattern problems
- Allow 90 seconds, students work on whiteboards
- Circulate to check reasoning

TEACHER NOTES:
Daily Review retrieves pattern and algorithm work. These patterns use multiplication and division rules, connecting to today's property focus.

WATCH FOR:
- Students who apply additive thinking to multiplicative patterns
- Students who identify the rule but make calculation errors

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers -- tick if correct, fix if not
- Fixing it is the learning
- Ask: What rule did pattern 3 use? [Divide by 3 each time]

DO:
- Click to reveal answers
- Students self-check and fix
- Note students who struggle with division patterns

TEACHER NOTES:
Tick-and-fix for DR. Division pattern recognition will support understanding why division is NOT commutative (addressed in extension).

WATCH FOR:
- Students who got multiplicative patterns right but division wrong -- division fluency may be a gap

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time -- division facts
- Write each answer on your whiteboard as fast as you can
- Go!

DO:
- Display 10 division facts
- Time 2 minutes
- After time, reveal answers for self-check

TEACHER NOTES:
Division fluency supports today's exploration of whether division is commutative (it is not). Students need fast division recall to evaluate equations efficiently.

WATCH FOR:
- Students who reverse dividend and divisor (e.g., 6 / 42 instead of 42 / 6)
- Students who are automatic with division -- fluency is strong

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me
- Read from slide: "We are learning to use the commutative and associative properties to build equivalent number sentences"
- Those are big words. Commutative means you can swap the order. Associative means you can regroup.
- Let's read our success criteria together.

DO:
- Choral read the LI and each SC
- Quick thumbs: "Has anyone heard the word commutative before?"

TEACHER NOTES:
The commutative property (a x b = b x a) and associative property ((a x b) x c = a x (b x c)) are the first algebraic tools students will use to construct equivalent sentences. Naming them builds mathematical vocabulary.

WATCH FOR:
- Students confused by the terminology -- the visual demonstrations in I Do will make it concrete

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. I will show you why 2 x 3 equals 3 x 2 using an array
- Here is 2 x 3: 2 rows of 3. Count them: 6
- Now I rotate the array. Same dots, same total. But now it is 3 rows of 2
- The order changed but the value stayed the same. This is the commutative property.
- Commutative means you can swap the order of multiplication and the answer stays the same
- Does this work for any numbers? Let's check: 4 x 7 = 28. And 7 x 4 = 28. Yes!
- So a x b = b x a. Always.

DO:
- Point to the array on screen
- Physically gesture "rotating" the array
- Write a x b = b x a clearly

TEACHER NOTES:
The array rotation makes the commutative property visual and concrete. Students can see that the same objects are just rearranged. This is not a rule to memorise -- it is a structure to see.

MISCONCEPTIONS:
- Misconception: The commutative property works for division too
  Why: Students overgeneralise from multiplication
  Impact: Students will write 6 / 3 = 3 / 6 and believe both equal 2
  Quick correction: "Try it: 6 / 3 = 2. But 3 / 6 = 0.5. Different answers! Division does NOT let you swap."

WATCH FOR:
- Students who nod and seem to already know this -- the commutative property may be intuitive for some
- Readiness signal: students who can predict the rotated array before you show it

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now a more powerful property: the associative property
- Watch: 2 x 2 x 3. How do I work this out?
- I could do (2 x 2) x 3 = 4 x 3 = 12
- OR I could do 2 x (2 x 3) = 2 x 6 = 12
- Same numbers, different grouping, same answer
- This is the associative property: you can regroup the factors and the product stays the same
- Why is this useful? Because 3 x 4 = 3 x (2 x 2) = (3 x 2) x 2 = 6 x 2
- I just turned 3 x 4 into 6 x 2! Same value, different expression. Equivalent!

DO:
- Write both groupings clearly on the board
- Circle the grouping brackets to show what changes
- Show the factor-splitting example step by step

TEACHER NOTES:
The associative property is the key algebraic tool for this unit. It allows students to decompose and recompose factors to create equivalent expressions. The example 3 x 4 = (3 x 2) x 2 = 6 x 2 directly mirrors the curriculum elaboration.

WATCH FOR:
- Students who understand commutative but find associative harder -- the grouping concept is more abstract
- Readiness signal: students who can suggest their own regrouping

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. On your whiteboards, use the commutative property to write an equivalent sentence for 8 x 5
- You have 5 seconds. Show me.

DO:
- Students write on whiteboards
- Scan for: 5 x 8

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Use the commutative property. Write an equivalent sentence for 8 x 5. Show me."
- Scan for: 5 x 8 (or 5 x 8 = 40). Students who write this instantly have grasped commutative.
PROCEED: If 80%+ show 5 x 8, move to associative practice.
PIVOT: Most likely misconception -- students write 8 + 5 or 40 (they gave the answer instead of an equivalent expression). Reteach: "Commutative means SWAP the order. 8 x 5 becomes 5 x 8. Not the answer -- a different way to write the same multiplication."

TEACHER NOTES:
This checks whether students can apply commutativity, not just recognise it. Writing the equivalent expression (not just the answer) is the target skill.

WATCH FOR:
- Students who write 40 instead of 5 x 8 -- they are answering rather than rewriting
- Students who instantly write 5 x 8 -- commutative is solid

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with a partner
- For each multiplication, use a property to write an equivalent expression
- Tell your partner WHICH property you used -- commutative or associative

DO:
- Display the four problems
- Partners work on whiteboards
- Circulate and listen for property naming
- Allow 2 minutes

TEACHER NOTES:
We Do requires students to choose the appropriate property and apply it. Problems 1-2 are commutative, problems 3-4 are associative. The mix forces selection, not just application.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 2 Enabling Scaffold with pre-drawn arrays. For each array, count the rows and columns, then write the two commutative forms (e.g., 3 x 4 and 4 x 3). The first two are completed as models.
- Extra Notes: Distribute the Session 2 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Investigate whether division is commutative. Test at least 5 division pairs (e.g., 12 / 3 vs 3 / 12). Record results and write an explanation of why division behaves differently from multiplication.
- Extra Notes: Distribute the Session 2 Extension PDF. Self-contained investigation.

WATCH FOR:
- Students who mix up commutative and associative -- prompt: "Commutative is SWAP. Associative is REGROUP."
- Readiness signal: partners naming the correct property without hesitation

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check. Problem 1: 9 x 7. Commutative gives us? [7 x 9] Both equal? [63]
- Problem 2: 6 x 12. Commutative: [12 x 6]. Both equal? [72]
- Problem 3: 2 x 3 x 5. Associative: [(2 x 3) x 5 = 6 x 5 = 30] or [2 x (3 x 5) = 2 x 15 = 30]
- Problem 4: 4 x 5 x 3. Associative: [(4 x 5) x 3 = 20 x 3 = 60] or [4 x (5 x 3) = 4 x 15 = 60]

DO:
- Reveal answers one at a time
- Cold call different students for each
- Celebrate when students name the property correctly

TEACHER NOTES:
The associative problems show that regrouping factors creates different-looking but equivalent expressions. This is the mechanism students will use to find unknowns in Sessions 3-4.

WATCH FOR:
- Students who got commutative right but struggled with associative -- they may need more modelling of regrouping
- Students who found multiple valid regroupings for the associative problems -- strong algebraic thinking

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. Use the associative property to rewrite 3 x 4 as ___ x 2
- Hint: you need to split one of the factors
- Write your answer on your whiteboard. Show me.

DO:
- Students think and write
- Scan for 6 x 2 (since 3 x 4 = 3 x (2 x 2) = (3 x 2) x 2 = 6 x 2)

CFU CHECKPOINT:
Technique: Show Me Boards with Justification
Script:
- Say: "Rewrite 3 x 4 as something times 2. Use the associative property. Show me your whiteboard AND your working."
- Scan for: 6 x 2. Students who show 3 x 4 = 3 x (2 x 2) = (3 x 2) x 2 = 6 x 2 have the full reasoning.
PROCEED: If 80%+ show 6 x 2, move to You Do.
PIVOT: Most likely misconception -- students write 12 x 2 = 24 (they calculated 3 x 4 = 12 and then wrote 12 x 2, which equals 24, not 12). Reteach: "We need the same VALUE. 3 x 4 = 12. So ___ x 2 must also equal 12. What times 2 equals 12? [6]. Now, how did we get 6? We split the 4 into 2 x 2, then regrouped: 3 x 2 = 6."

TEACHER NOTES:
This hinge directly tests the curriculum elaboration: 3 x 4 = (3 x 2) x 2 = 6 x 2. If students can do this, they can construct equivalent expressions by decomposing factors.

WATCH FOR:
- Students who cannot split 4 into 2 x 2 -- factor knowledge is a prerequisite gap
- Students who get 6 x 2 but cannot show the working -- press for the reasoning

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice time
- Your worksheet has three sections: commutative, associative, and mixed
- First: Read each problem. Next: Apply the property. Then: Write the equivalent expression.
- Show your working. You have 8 minutes.

DO:
- Distribute Session 2 Worksheet
- Students work independently
- Circulate: support students who struggled with associative
- For enabling students, distribute Session 2 Enabling Scaffold
- For extending students, distribute Session 2 Extension

TEACHER NOTES:
You Do uses different numbers from We Do. Section C (mixed) requires students to choose the property, not just apply a given one. This is the transfer step.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Complete the enabling scaffold with pre-drawn arrays. Match each array to its commutative partner by drawing the rotated version and writing both multiplication sentences.
- Extra Notes: Distribute the Session 2 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Investigate why division is NOT commutative. Test division pairs, record results, and write an explanation of why swapping dividend and divisor changes the value.
- Extra Notes: Distribute the Session 2 Extension PDF.

WATCH FOR:
- Students who can do commutative instantly but stall on associative -- check their factor knowledge
- Readiness signal: 80%+ completing Sections A and B correctly within 5 minutes

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket time. Three questions in your workbook.
- Work on your own. 3 minutes.

DO:
- Display exit ticket questions
- Students work silently
- Collect to sort into secure, developing, beginning

TEACHER NOTES:
Q1 tests commutative (SC1). Q2 tests associative (SC2). Q3 tests reasoning about which property was used (SC3). Sort to plan Session 3 support.

WATCH FOR:
- Students who get Q1 and Q2 but struggle with Q3 -- they can apply but not yet name/explain
- Students who finish all three instantly -- ready for distributive property in Session 3

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. Look at the slide.
- SC1: I can use the commutative property to rewrite a multiplication. Thumbs?
- SC2: I can use the associative property to regroup factors. Thumbs?
- SC3: I can explain which property makes two expressions equivalent. Thumbs?
- Turn and talk: What is the difference between commutative and associative? Explain to your partner.

DO:
- Run thumbs check for each SC
- 30 seconds Turn and Talk
- Cold call 2 students

TEACHER NOTES:
The Turn and Talk targets the distinction between the two properties. Students who can articulate "commutative is swapping order, associative is changing grouping" have the conceptual understanding needed for Session 3.

WATCH FOR:
- Students who conflate the two properties -- they may need visual anchor cards
- Students who use the terms correctly and confidently -- strong mathematical vocabulary

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Resources for today's lesson are linked here

DO:
- Point out each resource

TEACHER NOTES:
Resource slide provides clickable links to all companion PDFs.

WATCH FOR:
- N/A

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ───────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, "Algebra: Finding Unknown Values", "Session 2: Commutative & Associative Properties", "Grade 5/6 Numeracy | Session 2 of 10 | Week 1", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2-3: Daily Review (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Patterns and Algorithms", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const problems = [
        "1.  4, 12, 36, 108, ___  (rule: ?)",
        "2.  7, 14, 28, 56, ___  (rule: ?)",
        "3.  729, 243, 81, 27, ___  (rule: ?)",
        "4.  1, 5, 25, 125, ___  (rule: ?)",
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
        { text: "Find the missing number.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "Name the rule.", options: { fontSize: 14, color: C.CHARCOAL } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Answers:  1) 324 (x3)    2) 112 (x2)    3) 9 (/3)    4) 625 (x5)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Division facts
  const sFl = pres.addSlide();
  addTopBar(sFl, STAGE_COLORS["1"]);
  addStageBadge(sFl, 1, "Fluency");
  addTitle(sFl, "Division Facts Speed Round", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

  addCard(sFl, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  const divFacts = [
    "1.  56 / 8 =",   "2.  63 / 7 =",
    "3.  48 / 6 =",   "4.  72 / 9 =",
    "5.  45 / 5 =",   "6.  84 / 12 =",
    "7.  36 / 4 =",   "8.  54 / 6 =",
    "9.  81 / 9 =",   "10. 96 / 8 =",
  ];
  sFl.addText(divFacts.map((p, i) => ({
    text: p,
    options: { fontSize: 14, color: C.CHARCOAL, breakLine: i < divFacts.length - 1, paraSpaceAfter: 4 },
  })), {
    x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addCard(sFl, 5.2, CONTENT_TOP, 4.3, 1.6, { strip: C.SECONDARY });
  sFl.addText([
    { text: "Whiteboards - 2 minutes", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
    { text: "Write each answer as fast as you can.", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Answers: 7, 9, 8, 8, 9, 7, 9, 9, 9, 12", options: { fontSize: 11, color: C.MUTED } },
  ], {
    x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: 1.4,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addFooter(sFl, FOOTER);
  sFl.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    ["We are learning to use the commutative and associative properties to build equivalent number sentences"],
    [
      "I can use the commutative property to rewrite a multiplication in a different order",
      "I can use the associative property to regroup factors and create an equivalent expression",
      "I can explain which property makes two expressions equivalent",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — Commutative Property with Arrays
  workedExSlide(pres, 2, "I Do", "Commutative Property: Swap the Order",
    [
      "Commutative means: swap the order",
      "a x b = b x a  (always true for multiplication)",
      "",
      "Example: 2 x 3 = 3 x 2",
      "  2 rows of 3 = 6",
      "  3 rows of 2 = 6",
      "  Same array, just rotated!",
      "",
      "Check: 4 x 7 = 7 x 4",
      "  28 = 28   (equivalent)",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.PRIMARY });
      slide.addText("Array: 2 x 3", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: 1.8, h: 0.28,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
      });
      // 2 rows x 3 cols array
      const dotR = 0.22;
      const dotGap = 0.36;
      const sx1 = lg.rightX + 0.3;
      const sy1 = lg.panelTopPadded + 0.45;
      for (let r = 0; r < 2; r++) {
        for (let c = 0; c < 3; c++) {
          slide.addShape("roundRect", {
            x: sx1 + c * dotGap, y: sy1 + r * dotGap,
            w: dotR, h: dotR, rectRadius: dotR / 2,
            fill: { color: C.PRIMARY },
          });
        }
      }

      // Arrow
      slide.addText("Rotate!", {
        x: lg.rightX + 1.6, y: lg.panelTopPadded + 0.55, w: 1.0, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0, align: "center",
      });

      slide.addText("Array: 3 x 2", {
        x: lg.rightX + 2.5, y: lg.panelTopPadded + 0.08, w: 1.8, h: 0.28,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
      });
      // 3 rows x 2 cols array
      const sx2 = lg.rightX + 2.7;
      for (let r = 0; r < 3; r++) {
        for (let c = 0; c < 2; c++) {
          slide.addShape("roundRect", {
            x: sx2 + c * dotGap, y: sy1 + r * dotGap,
            w: dotR, h: dotR, rectRadius: dotR / 2,
            fill: { color: C.SECONDARY },
          });
        }
      }

      // Both = 6
      addTextOnShape(slide, "Both = 6    Equivalent!", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.65, w: lg.rightW - 0.6, h: 0.35, rectRadius: 0.06,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });

      // General rule
      addTextOnShape(slide, "a x b = b x a", {
        x: lg.rightX + 0.6, y: lg.panelTopPadded + 2.2, w: lg.rightW - 1.2, h: 0.45, rectRadius: 0.08,
        fill: { color: C.PRIMARY },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });

      slide.addText("The COMMUTATIVE property", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 2.75, w: lg.rightW - 0.6, h: 0.3,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0,
      });
    }
  );

  // Slide 7: I Do — Associative Property
  workedExSlide(pres, 2, "I Do", "Associative Property: Regroup the Factors",
    [
      "Associative means: regroup the factors",
      "",
      "2 x 2 x 3 = ?",
      "  Group A: (2 x 2) x 3 = 4 x 3 = 12",
      "  Group B: 2 x (2 x 3) = 2 x 6 = 12",
      "  Same answer either way!",
      "",
      "Power move: rewrite 3 x 4 as ___ x 2",
      "  3 x 4 = 3 x (2 x 2)",
      "        = (3 x 2) x 2",
      "        = 6 x 2     (equivalent!)",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.5, { strip: C.SECONDARY });
      slide.addText("Associative Property", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      slide.addText("(a x b) x c = a x (b x c)", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.4, w: lg.rightW - 0.6, h: 0.35,
        fontSize: 17, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      slide.addText("Regroup the factors -- same product!", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.8, w: lg.rightW - 0.6, h: 0.3,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, align: "center",
      });

      // Power move box
      addCard(slide, lg.rightX, lg.panelTopPadded + 1.7, lg.rightW, 1.5, { strip: C.ACCENT });
      slide.addText("The Power Move", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.76, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "3 x 4 = 12", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Split 4 into 2 x 2", options: { fontSize: 13, color: C.ALERT, bold: true, breakLine: true } },
        { text: "3 x (2 x 2) = (3 x 2) x 2 = 6 x 2", options: { fontSize: 13, color: C.PRIMARY, bold: true, breakLine: true } },
        { text: "So 3 x 4 = 6 x 2", options: { fontSize: 14, color: C.CHARCOAL, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 2.1, w: lg.rightW - 0.5, h: 1.0,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Slide 8-9: CFU 1 (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check — Commutative", "Show Me Boards",
      "Use the commutative property:\n\nWrite an equivalent expression for 8 x 5",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "8 x 5 = 5 x 8 = 40    Just swap the order!", {
        x: 1.0, y: 4.0, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10-11: We Do (withReveal)
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Apply the Properties",
      [
        "Use a property to write an equivalent expression.",
        "Name the property you used.",
        "",
        "1.   9 x 7 = ___  (property: ___)",
        "2.   6 x 12 = ___  (property: ___)",
        "3.   2 x 3 x 5 = ___ x 5  (property: ___)",
        "4.   4 x 5 x 3 = ___ x 3  (property: ___)",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.SECONDARY });
        slide.addText("Property Reference:", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.28,
          fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        const props = [
          { label: "Commutative", desc: "Swap the order: a x b = b x a", color: C.PRIMARY },
          { label: "Associative", desc: "Regroup: (a x b) x c = a x (b x c)", color: C.SECONDARY },
        ];
        props.forEach((prop, i) => {
          const py = lg.panelTopPadded + 0.45 + i * 0.85;
          addTextOnShape(slide, prop.label, {
            x: lg.rightX + 0.2, y: py, w: 1.8, h: 0.35, rectRadius: 0.06,
            fill: { color: prop.color },
          }, { fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true });
          slide.addText(prop.desc, {
            x: lg.rightX + 0.25, y: py + 0.4, w: lg.rightW - 0.5, h: 0.3,
            fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
          });
        });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "1) 7x9 (comm.)  2) 12x6 (comm.)  3) 6x5 (assoc.)  4) 20x3 (assoc.)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 12-13: CFU 2 Hinge (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Check — Associative", "Show Me Boards with Justification",
      "Use the associative property:\n\nRewrite 3 x 4 as ___ x 2\n\nShow your working.",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "3 x 4 = 3 x (2 x 2) = (3 x 2) x 2 = 6 x 2", {
        x: 0.8, y: 4.0, w: 8.4, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do
  workedExSlide(pres, 4, "You Do", "Independent Practice",
    [
      "First: Read each problem on your worksheet.",
      "Next: Apply the correct property.",
      "Then: Write the equivalent expression.",
      "",
      "Name the property for each one.",
      "Show your working.",
      "",
      "You have 8 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.0, { strip: C.ALERT });
      slide.addText("Remember:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Commutative = SWAP the order", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Associative = REGROUP the factors", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Both give EQUIVALENT expressions", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.5, h: 1.0,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "Write the commutative partner for 11 x 8.",
      "Rewrite 5 x 6 as ___ x 3 using the associative property. Show working.",
      "Tom says 4 x 9 = 9 x 4 because of the associative property. Is he correct? Explain.",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 16: Closing
  closingSlide(pres,
    "What is the difference between commutative and associative? Explain to your partner with an example.",
    [
      "I can use the commutative property to rewrite a multiplication",
      "I can use the associative property to regroup factors",
      "I can explain which property makes two expressions equivalent",
    ],
    NOTES_CLOSING);

  // Slide 17: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "ALG_Session2_Commutative_Associative.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Generate PDFs ─────────────────────────────────────────────────────────

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Commutative & Associative Properties",
      color: C.NAVY, lessonInfo: "Session 2 of 10 | Grade 5/6 Numeracy",
    });
    y = addTipBox(doc, "For each problem: apply the named property and write an equivalent expression. Show your working!", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A: Commutative Property (swap the order)", y, { color: C.NAVY });
    y = addProblem(doc, 1, "7 x 9 = ___ x ___", y, { color: C.NAVY });
    y = addProblem(doc, 2, "12 x 5 = ___ x ___", y, { color: C.NAVY });
    y = addProblem(doc, 3, "15 x 4 = ___ x ___", y, { color: C.NAVY });
    y = addProblem(doc, 4, "8 x 11 = ___ x ___", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Section B: Associative Property (regroup the factors)", y, { color: C.NAVY });
    y = addProblem(doc, 5, "2 x 4 x 3 = (2 x 4) x 3 = ___ x 3 = ___", y, { color: C.NAVY });
    y = addProblem(doc, 6, "5 x 2 x 6 = 5 x (2 x 6) = 5 x ___ = ___", y, { color: C.NAVY });
    y = addProblem(doc, 7, "Rewrite 4 x 6 as ___ x 3 (split 6 into 2 x 3, regroup)", y, { color: C.NAVY });
    y = addProblem(doc, 8, "Rewrite 5 x 8 as ___ x 4 (split 8 into 2 x 4, regroup)", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Section C: Mixed — Which Property?", y, { color: C.NAVY });
    y = addProblem(doc, 9, "6 x 7 = 7 x 6        Property used: ___", y, { color: C.NAVY });
    y = addProblem(doc, 10, "3 x 4 x 2 = 12 x 2   Property used: ___", y, { color: C.NAVY });
    y = addProblem(doc, 11, "9 x 8 = 8 x 9        Property used: ___", y, { color: C.NAVY });
    y = addProblem(doc, 12, "Rewrite 6 x 4 as ___ x 2 using the associative property. Show working:", y, { color: C.NAVY });
    y = addWriteLine(doc, "", y);

    addPdfFooter(doc, "Session 2 | Algebra: Finding Unknown Values | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Commutative & Associative Properties",
      color: C.NAVY, lessonInfo: "Session 2 of 10 | Grade 5/6 Numeracy",
    });
    y = addSectionHeading(doc, "Section A: Commutative", y, { color: C.NAVY });
    y = addBodyText(doc, "1. 7 x 9 = 9 x 7 = 63", y);
    y = addBodyText(doc, "2. 12 x 5 = 5 x 12 = 60", y);
    y = addBodyText(doc, "3. 15 x 4 = 4 x 15 = 60", y);
    y = addBodyText(doc, "4. 8 x 11 = 11 x 8 = 88", y);
    y = addSectionHeading(doc, "Section B: Associative", y, { color: C.NAVY });
    y = addBodyText(doc, "5. (2 x 4) x 3 = 8 x 3 = 24", y);
    y = addBodyText(doc, "6. 5 x (2 x 6) = 5 x 12 = 60", y);
    y = addBodyText(doc, "7. 4 x 6 = 4 x (2 x 3) = (4 x 2) x 3 = 8 x 3 = 24", y);
    y = addBodyText(doc, "8. 5 x 8 = 5 x (2 x 4) = (5 x 2) x 4 = 10 x 4 = 40", y);
    y = addSectionHeading(doc, "Section C: Mixed", y, { color: C.NAVY });
    y = addBodyText(doc, "9. Commutative (order swapped)", y);
    y = addBodyText(doc, "10. Associative (3 x 4 regrouped to 12)", y);
    y = addBodyText(doc, "11. Commutative (order swapped)", y);
    y = addBodyText(doc, "12. 6 x 4 = 6 x (2 x 2) = (6 x 2) x 2 = 12 x 2 = 24", y);
    addPdfFooter(doc, "Session 2 | Answer Key | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold — Pre-drawn arrays
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Array Models for Commutative Property",
      color: C.TEAL, lessonInfo: "Session 2 of 10 | Grade 5/6 Numeracy",
    });
    y = addTipBox(doc, "For each array: count the rows and columns. Write the multiplication sentence. Then draw the rotated array and write its multiplication sentence.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Example (done for you)", y, { color: C.NAVY });
    y = addBodyText(doc, "Array: 3 rows of 4 = 3 x 4 = 12", y);
    y = addBodyText(doc, "Rotated: 4 rows of 3 = 4 x 3 = 12", y);
    y = addBodyText(doc, "Same value! Commutative property: 3 x 4 = 4 x 3", y);
    y += 10;

    y = addSectionHeading(doc, "Now you try:", y, { color: C.NAVY });
    const arrayProbs = [
      { rows: 2, cols: 5 },
      { rows: 4, cols: 3 },
      { rows: 6, cols: 2 },
      { rows: 3, cols: 7 },
    ];
    arrayProbs.forEach((ap, i) => {
      y = addBodyText(doc, (i + 1) + ". Array: " + ap.rows + " rows of " + ap.cols, y);
      y = addBodyText(doc, "   ___ x ___ = ___", y);
      y = addBodyText(doc, "   Rotated: ___ rows of ___", y);
      y = addBodyText(doc, "   ___ x ___ = ___", y);
      y = addBodyText(doc, "   Commutative: ___ x ___ = ___ x ___", y);
      y += 10;
    });

    addPdfFooter(doc, "Session 2 | Enabling Scaffold | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Extension — Why Division is NOT Commutative
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Is Division Commutative?",
      color: C.NAVY, lessonInfo: "Session 2 of 10 | Grade 5/6 Numeracy",
    });

    y = addSectionHeading(doc, "The Big Question", y, { color: C.NAVY });
    y = addBodyText(doc, "We know multiplication is commutative: a x b = b x a (always true).", y);
    y = addBodyText(doc, "But what about division? Is a / b = b / a?", y);
    y = addBodyText(doc, "Let's investigate!", y);

    y = addSectionHeading(doc, "Worked Example", y, { color: C.NAVY });
    y = addBodyText(doc, "Test: Is 12 / 3 = 3 / 12?", y);
    y = addBodyText(doc, "   12 / 3 = 4", y);
    y = addBodyText(doc, "   3 / 12 = 0.25", y);
    y = addBodyText(doc, "   4 =/= 0.25 -> NOT equal! Division is NOT commutative here.", y);

    y = addSectionHeading(doc, "Your Investigation", y, { color: C.NAVY });
    y = addBodyText(doc, "Test at least 5 division pairs. Record your results:", y);
    y += 5;

    const pairs = ["20 / 4 vs 4 / 20", "6 / 2 vs 2 / 6", "15 / 5 vs 5 / 15", "8 / 4 vs 4 / 8", "24 / 6 vs 6 / 24"];
    pairs.forEach((pair, i) => {
      y = addProblem(doc, i + 1, pair + "     Left: ___   Right: ___   Equal? ___", y, { color: C.NAVY });
    });

    y = addSectionHeading(doc, "Your Conclusion", y, { color: C.NAVY });
    y = addBodyText(doc, "Is division commutative? Explain WHY or WHY NOT in your own words:", y);
    y = addWriteLine(doc, "", y);
    y = addWriteLine(doc, "", y);
    y = addWriteLine(doc, "", y);

    y = addSectionHeading(doc, "Challenge Question", y, { color: C.NAVY });
    y = addBodyText(doc, "Is there ANY pair of numbers where a / b = b / a?", y);
    y = addBodyText(doc, "Hint: Think about what happens when both numbers are the same.", y);
    y = addWriteLine(doc, "", y);

    y = addTipBox(doc, "Did You Know? In mathematics, an operation that works both ways (like multiplication) is called commutative. Division is called non-commutative because the order matters. This is why we have to be extra careful with division in equations!", y, { color: C.TEAL });

    addPdfFooter(doc, "Session 2 | Extension Investigation | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 2 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
