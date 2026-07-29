"use strict";

// Algebra Unit — Session 10: Consolidation & Application
// Week 2 Session 5, Grade 5/6 Numeracy, Variant 0
// DR: Number Patterns with Factors and Multiples
// VC2M5A02/VC2M6A02 — consolidation of all algebra skills

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

const SESSION = 10;
const FOOTER = "Algebra | Session 10 of 10 | Grade 5/6 Numeracy";
const OUT_DIR = "output/ALG_Session10_Consolidation_Final";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Final consolidation - all algebra skills mixed.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Property and strategy reference card with guided practice.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Open-ended algebra puzzle challenge.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

const NOTES_TITLE = `SAY:
- Session 10 -- the final lesson in our algebra unit!
- Over two weeks you have built a serious toolkit: equivalence, properties, order of operations, brackets, unknowns, and pairs
- Today we put it ALL together in a mixed challenge

DO:
- Display title slide

TEACHER NOTES:
Session 10 of 10. Final consolidation. No new content. Students apply all skills from both weeks to mixed problems. This lesson serves as both practice and summative assessment preparation.

WATCH FOR:
- Overall confidence level -- this is the culmination of the unit

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Last daily review of the unit -- factors and multiples
- Find the HCF and LCM of each pair

DO:
- Display problems, 90 seconds

TEACHER NOTES:
DR on HCF/LCM connects to factor work that underpins many of the algebraic strategies used throughout the unit.

WATCH FOR:
- Students who find HCF and LCM confidently

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers

DO:
- Reveal, self-check

TEACHER NOTES:
Quick tick-and-fix for final DR.

WATCH FOR:
- Factor fluency

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Final fluency -- mixed operations speed round
- Everything: brackets, order of ops, multiplication, division
- Go!

DO:
- Display 10 expressions, 2 minutes

TEACHER NOTES:
Comprehensive fluency covering all skills from the unit.

WATCH FOR:
- Students applying order of operations automatically

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention
- Read from slide: "We are learning to apply all our algebra skills to solve mixed problems confidently"
- Today you choose the strategy. There is no single approach -- you use whatever tool fits.
- Read our success criteria

DO:
- Choral read

TEACHER NOTES:
Metacognitive demand is highest today -- students must select from their full toolkit. The success criteria span both weeks.

WATCH FOR:
- Confidence in strategy selection

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me work through a mixed problem set. I will think aloud about which tool to use.
- Problem 1: Is 4 x (3 + 5) = 8 x 4 true?
- Left: brackets first, 3+5=8, then 4x8=32. Right: 8x4=32. TRUE.
- I could also prove it: 4x(3+5)=4x8 and 8x4=32. Commutative: 4x8=8x4.
- Problem 2: Find ? in 6 x ? - 10 = 32
- Work backwards: undo -10 first, 32+10=42. Then 6x?=42, ?=7.
- Check: 6x7-10=42-10=32. Correct.

DO:
- Show strategy selection process explicitly
- Name the tool used each time

TEACHER NOTES:
I Do models the metacognitive process: read the problem, select the strategy, apply it, check. This is the expert approach students should emulate.

WATCH FOR:
- Students who can predict the strategy before you name it

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Problem 3: Find all natural number pairs for triangle + 2 x square = 16
- This has two unknowns. Use a systematic table.
- If square=1: triangle+2=16, triangle=14. If square=2: triangle+4=16, triangle=12.
- Continue: square=3, tri=10. square=4, tri=8. ... square=7, tri=2.
- That gives 7 pairs.
- Problem 4: Use the distributive property to calculate 9 x 25.
- 9x(20+5)=180+45=225. Quick mental maths!

DO:
- Show both problem types with strategy selection
- Emphasise: you know WHEN to use WHICH tool

TEACHER NOTES:
The final I Do covers paired unknowns and distributive calculation -- spanning both weeks. Students see the full breadth of the unit in two examples.

WATCH FOR:
- Students who are comfortable with both problem types

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. Which strategy would you use for: 5 x (? + 4) = 45?
- A: Commutative  B: Distributive  C: Backwards (undo)  D: Systematic table
- Whisper to your partner.

DO:
- Partners whisper, then cold call

CFU CHECKPOINT:
Technique: Think-Pair-Share
Script:
- Say: "Which strategy for 5 x (? + 4) = 45? Whisper..."
- Scan for: C -- backwards/undo. (Undo x5: bracket=9. Then ?+4=9, ?=5.)
PROCEED: If student answers C correctly, continue.
PIVOT: Students say B (distributive) -- distributive EXPANDS brackets, it does not solve for unknowns inside them. Reteach: "Distributive is for CALCULATING, like 5 x 14. When you need to FIND a missing number, you work BACKWARDS -- undo the outer operation."

TEACHER NOTES:
Strategy identification is the metacognitive test. Students who can select the right tool before solving have internalized the toolkit.

WATCH FOR:
- Students who confuse strategy selection

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Mixed challenge with your partner
- Six problems, six different types. Choose your strategy for each.

DO:
- Display problems, 4 minutes

TEACHER NOTES:
We Do covers: equivalence check, property ID, single unknown, brackets unknown, distributive calc, and paired unknowns. One of each type.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use Session 10 Enabling Scaffold with the full strategy reference card. Match each problem to a strategy first, then solve.
- Extra Notes: Distribute Session 10 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Open-ended algebra puzzles. Create equations that have exactly 3 solutions, exactly 1 solution, and no natural number solutions. Explain your reasoning.
- Extra Notes: Distribute Session 10 Extension PDF.

WATCH FOR:
- Students selecting correct strategies without prompting
- Readiness signal: completing all six correctly

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check.
- 1: TRUE/FALSE: 7x9 = 9x7. TRUE (commutative).
- 2: Name property: 5x(10+4) = 50+20. Distributive.
- 3: Find ?: 8x?=56. ?=7.
- 4: Find ?: 3x(?+5)=27. Bracket=9, ?=4.
- 5: Distributive: 7x16 = 7x(10+6) = 70+42 = 112.
- 6: Pairs for axb=12: (1,12)(2,6)(3,4)(4,3)(6,2)(12,1) = 6 pairs.

DO:
- Reveal, cold call for strategy names

TEACHER NOTES:
Celebrating correct strategy selection is as important as correct answers.

WATCH FOR:
- Students who got strategies AND answers correct

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Final hinge of the unit. Thumbs up or down.
- Is this true: (3 + 5) x (2 + 4) = 8 x 6 ?

DO:
- Students show thumbs

CFU CHECKPOINT:
Technique: Thumbs Up/Down
Script:
- Say: "(3+5) x (2+4) = 8 x 6. True or false? Thumbs."
- Scan for: thumbs up. Left: 8 x 6 = 48. Right: 8 x 6 = 48. TRUE.
PROCEED: If 80%+ thumbs up, move to You Do.
PIVOT: Students show thumbs down (they may have evaluated brackets incorrectly). Reteach: "Bracket 1: 3+5=8. Bracket 2: 2+4=6. So (3+5)x(2+4) = 8x6. The right side is also 8x6. Same expression! Of course it is true."

TEACHER NOTES:
Students who see that the left side simplifies to exactly the right side have strong bracket processing.

WATCH FOR:
- Students who evaluate both sides correctly and see the equivalence

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Final independent practice of the unit
- Your worksheet is mixed -- every type of problem we have done
- Choose your strategy. Show your working.
- 10 minutes.

DO:
- Distribute worksheet, circulate

TEACHER NOTES:
Extended You Do (10 minutes) for the final session. This serves as informal assessment of all unit skills.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Strategy reference card with matched practice problems.
- Extra Notes: Distribute Session 10 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Open-ended algebra puzzles.
- Extra Notes: Distribute Session 10 Extension PDF.

WATCH FOR:
- Overall mastery level across the class
- Which problem types are most commonly incorrect -- plan future review

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Final exit ticket. 4 questions covering the whole unit.
- Work on your own. 4 minutes.

DO:
- Display questions, collect

TEACHER NOTES:
Summative-style exit ticket. Q1: equivalence check. Q2: single unknown with brackets. Q3: order of operations. Q4: paired unknowns.

WATCH FOR:
- This data informs future review and follow-up

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Unit wrap-up! Let's review ALL the success criteria from this unit.
- Over 10 sessions you learned: equivalence, commutative, associative, distributive properties, order of operations, brackets, finding single unknowns, finding pairs of unknowns.
- SC1: I can apply properties to build and verify equivalent equations. Thumbs?
- SC2: I can find unknown values using backwards strategies. Thumbs?
- SC3: I can solve problems with brackets and order of operations. Thumbs?
- Turn and talk: What was the most important thing you learned in this unit?

DO:
- Thumbs for each SC
- Extended Turn and Talk (1 minute)
- Cold call 3-4 students
- Celebrate the class's growth over the unit

TEACHER NOTES:
Final reflection. Students who can name a specific strategy or concept as "most important" have metacognitive awareness. Celebrate growth, not just mastery.

WATCH FOR:
- The range of reflections -- this tells you what stuck
- Students who name specific strategies -- strong metacognition

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- All resources for Session 10 are linked here

DO:
- Point out resources

TEACHER NOTES:
Final resource slide.

WATCH FOR:
- N/A

[General: Resources | VTLM 2.0: Planning]`;

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, "Algebra: Finding Unknown Values", "Session 10: Final Consolidation", "Grade 5/6 Numeracy | Session 10 of 10 | Week 2", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // DR — Number Patterns with Factors and Multiples
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]); addStageBadge(s, 1, "Daily Review");
      addTitle(s, "HCF and LCM", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      s.addText(["1.  Find the HCF of 12 and 18", "2.  Find the LCM of 4 and 6", "3.  Find the HCF of 20 and 30", "4.  Find the LCM of 8 and 12"].map((p, i) => ({
        text: p, options: { fontSize: 14.5, color: C.CHARCOAL, breakLine: i < 3, paraSpaceAfter: 14 },
      })), { x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3, fontFace: FONT_B, margin: 0, valign: "top" });
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboards:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "HCF = largest factor they share", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "LCM = smallest multiple they share", options: { fontSize: 14, color: C.CHARCOAL } },
      ], { x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3, fontFace: FONT_B, margin: 0, valign: "top" });
      addFooter(s, FOOTER); s.addNotes(NOTES_DR_Q); return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) HCF=6  2) LCM=12  3) HCF=10  4) LCM=24", {
        x: 1.0, y: 4.55, w: 8, h: 0.5, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Fluency
  const sFl = pres.addSlide();
  addTopBar(sFl, STAGE_COLORS["1"]); addStageBadge(sFl, 1, "Fluency");
  addTitle(sFl, "Mixed Operations Final Round", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });
  addCard(sFl, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  sFl.addText(["1.  (4 + 5) x 3 =", "2.  7 x 8 - 6 =", "3.  48 / (12 - 4) =", "4.  5 x 6 + 4 x 3 =", "5.  (9 - 3) x (2 + 5) =", "6.  100 - 8 x 12 =", "7.  3 x (15 - 7) + 4 =", "8.  60 / 5 + 8 x 2 =", "9.  (3 + 7) x (4 + 1) =", "10. 24 / (2 + 4) x 3 ="].map((p, i) => ({
    text: p, options: { fontSize: 13, color: C.CHARCOAL, breakLine: i < 9, paraSpaceAfter: 2 },
  })), { x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2, fontFace: FONT_B, margin: 0, valign: "top" });
  addCard(sFl, 5.2, CONTENT_TOP, 4.3, 2.0, { strip: C.SECONDARY });
  sFl.addText([
    { text: "Whiteboards - 2 minutes", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
    { text: "Brackets first, then x /, then + -", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "27, 50, 6, 42, 42, 4, 28, 28, 50, 12", options: { fontSize: 10, color: C.MUTED } },
  ], { x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: 1.8, fontFace: FONT_B, margin: 0, valign: "top" });
  addFooter(sFl, FOOTER); sFl.addNotes(NOTES_FLUENCY);

  liSlide(pres,
    ["We are learning to apply all our algebra skills to solve mixed problems confidently"],
    [
      "I can apply properties to build and verify equivalent equations",
      "I can find unknown values using backwards strategies with brackets and operations",
      "I can solve problems involving order of operations, brackets, and paired unknowns",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do 1 — Mixed problems with strategy selection
  workedExSlide(pres, 2, "I Do", "Choose Your Tool",
    [
      "Problem 1: TRUE or FALSE?",
      "  4 x (3 + 5) = 8 x 4",
      "  Left: 4 x 8 = 32. Right: 8 x 4 = 32. TRUE.",
      "  Tool: evaluate and compare + commutative",
      "",
      "Problem 2: Find ? in 6 x ? - 10 = 32",
      "  Undo -10: 6 x ? = 42. Then ? = 7.",
      "  Check: 6 x 7 - 10 = 42 - 10 = 32.",
      "  Tool: backwards strategy",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("Your Full Toolkit", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      const tools = [
        "Evaluate & Compare (equivalence)",
        "Commutative (swap order)",
        "Associative (regroup factors)",
        "Distributive (split across +)",
        "Order of Operations",
        "Backwards / Undo (find unknowns)",
        "Systematic Tables (paired unknowns)",
      ];
      slide.addText(tools.map((t, i) => ({
        text: t, options: { bullet: true, fontSize: 11, color: C.CHARCOAL, breakLine: i < tools.length - 1 },
      })), { x: lg.rightX + 0.25, y: lg.panelTopPadded + 0.38, w: lg.rightW - 0.45, h: 2.4, fontFace: FONT_B, margin: 0, valign: "top" });
    }
  );

  // I Do 2
  workedExSlide(pres, 2, "I Do", "Paired Unknowns and Distributive",
    [
      "Problem 3: Find pairs for tri + 2 x sq = 16",
      "  sq=1: tri+2=16, tri=14",
      "  sq=2: tri+4=16, tri=12",
      "  sq=3: tri+6=16, tri=10",
      "  ... sq=7: tri+14=16, tri=2  (7 pairs)",
      "  Tool: systematic table",
      "",
      "Problem 4: 9 x 25 using distributive",
      "  9 x (20 + 5) = 180 + 45 = 225",
      "  Tool: distributive property",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.4, { strip: C.SECONDARY });
      slide.addText("Paired Unknowns", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
      });
      slide.addText("Fix one unknown, calculate the other.\nUse a systematic table.\nStop when result <= 0.", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.35, w: lg.rightW - 0.6, h: 0.8,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });

      addCard(slide, lg.rightX, lg.panelTopPadded + 1.6, lg.rightW, 1.2, { strip: C.ACCENT });
      slide.addText("Distributive for Mental Maths", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.66, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 13, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
      });
      slide.addText("Split into friendly parts.\n9 x 25 = 9 x (20+5) = 180+45 = 225", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.95, w: lg.rightW - 0.6, h: 0.6,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });
    }
  );

  // CFU 1
  withReveal(
    () => cfuSlide(pres, "CFU", "Strategy Selection", "Think-Pair-Share",
      "Which strategy for:\n\n5 x (? + 4) = 45\n\nA: Commutative  B: Distributive\nC: Backwards  D: Systematic table",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "C: Backwards. Bracket = 45/5 = 9. ? + 4 = 9. ? = 5.", {
        x: 0.5, y: 4.0, w: 9, h: 0.55, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // We Do
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Mixed Challenge",
      [
        "Choose your strategy for each:",
        "",
        "1.  TRUE/FALSE: 7 x 9 = 9 x 7",
        "2.  Name property: 5 x (10+4) = 50+20",
        "3.  Find ?: 8 x ? = 56",
        "4.  Find ?: 3 x (? + 5) = 27",
        "5.  Distributive: 7 x 16 = ?",
        "6.  Pairs for a x b = 12: list all",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.8, { strip: C.SECONDARY });
        slide.addText("Match the problem to the tool!", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.3,
          fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
        });
        slide.addText([
          { text: "Equivalence check", options: { bullet: true, fontSize: 12, color: C.CHARCOAL, breakLine: true } },
          { text: "Property identification", options: { bullet: true, fontSize: 12, color: C.CHARCOAL, breakLine: true } },
          { text: "Single unknown (evaluate/backwards)", options: { bullet: true, fontSize: 12, color: C.CHARCOAL, breakLine: true } },
          { text: "Distributive calculation", options: { bullet: true, fontSize: 12, color: C.CHARCOAL, breakLine: true } },
          { text: "Paired unknowns (table/factor pairs)", options: { bullet: true, fontSize: 12, color: C.CHARCOAL } },
        ], { x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.4, w: lg.rightW - 0.5, h: 1.2, fontFace: FONT_B, margin: 0, valign: "top" });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "1)TRUE  2)Distrib.  3)?=7  4)?=4  5)112  6) 6 pairs", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // CFU 2 Hinge
  withReveal(
    () => cfuSlide(pres, "CFU", "Final Hinge", "Thumbs Up/Down",
      "TRUE or FALSE?\n\n(3 + 5) x (2 + 4) = 8 x 6",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "TRUE: (3+5)=8, (2+4)=6, so 8 x 6 = 8 x 6 = 48", {
        x: 0.5, y: 4.0, w: 9, h: 0.55, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // You Do
  workedExSlide(pres, 4, "You Do", "Final Independent Challenge",
    [
      "First: Read each problem.",
      "Next: Choose your strategy.",
      "Then: Solve and show working.",
      "",
      "This is your final challenge.",
      "Use everything you have learned!",
      "",
      "You have 10 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.6, { strip: C.ALERT });
      slide.addText("Unit Toolkit:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      const tools = ["Evaluate & Compare", "Commutative (swap)", "Associative (regroup)", "Distributive (split)", "Order of Operations", "Backwards (undo)", "Systematic Tables (pairs)"];
      slide.addText(tools.map((t, i) => ({
        text: t, options: { bullet: true, fontSize: 12, color: C.CHARCOAL, breakLine: i < tools.length - 1 },
      })), { x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.38, w: lg.rightW - 0.5, h: 2.0, fontFace: FONT_B, margin: 0, valign: "top" });
    }
  );

  exitTicketSlide(pres,
    [
      "TRUE or FALSE: 5 x (4 + 6) = 10 x 5. Show working.",
      "Find ?: 4 x (? + 3) = 36. Show all steps.",
      "Calculate: 15 + 6 x 3 - 4 = ?  (order of operations)",
      "Find 2 natural number pairs for: a + 3 x b = 19",
    ],
    NOTES_EXIT, FOOTER);

  closingSlide(pres,
    "What was the most important thing you learned in this algebra unit? Tell your partner.",
    [
      "I can apply properties to build and verify equivalent equations",
      "I can find unknown values using backwards strategies",
      "I can solve problems with brackets, order of ops, and paired unknowns",
    ],
    NOTES_CLOSING);



  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "ALG_Session10_Consolidation_Final.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // PDFs
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, { subtitle: "Final Consolidation", color: C.NAVY, lessonInfo: "Session 10 of 10 | Grade 5/6" });
    y = addTipBox(doc, "Mixed practice. Choose your strategy for each problem. Show all working!", y, { color: C.TEAL });
    y = addSectionHeading(doc, "A: Equivalence and Properties", y, { color: C.NAVY });
    y = addProblem(doc, 1, "TRUE/FALSE: 6 x 12 = 12 x 6. Name the property.", y, { color: C.NAVY });
    y = addProblem(doc, 2, "TRUE/FALSE: 4 x (3+5) = 4x3 + 4x5. Name the property.", y, { color: C.NAVY });
    y = addProblem(doc, 3, "TRUE/FALSE: 3x8 = 6x5. Show working.", y, { color: C.NAVY });
    y = addSectionHeading(doc, "B: Find the Unknown", y, { color: C.NAVY });
    y = addProblem(doc, 4, "7 x ? = 63                        ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 5, "5 x (? + 2) = 35                  ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 6, "4 x ? - 8 = 20                    ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 7, "60 / (? + 5) = 6                  ? = ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "C: Order of Operations", y, { color: C.NAVY });
    y = addProblem(doc, 8, "8 + 3 x 5 = ___", y, { color: C.NAVY });
    y = addProblem(doc, 9, "(8 + 3) x 5 = ___", y, { color: C.NAVY });
    y = addProblem(doc, 10, "24 / 6 + 4 x 3 = ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "D: Distributive Calculation", y, { color: C.NAVY });
    y = addProblem(doc, 11, "8 x 23 = 8 x (___+___) = ___+___ = ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "E: Paired Unknowns", y, { color: C.NAVY });
    y = addProblem(doc, 12, "Find ALL natural number pairs: a x b = 20", y, { color: C.NAVY });
    y = addWriteLine(doc, "Pairs:", y);
    y = addProblem(doc, 13, "Find 3 pairs: 2 x a + b = 18", y, { color: C.NAVY });
    y = addWriteLine(doc, "Pairs:", y);
    addPdfFooter(doc, "Session 10 | Algebra | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName)); console.log("PDF: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, { subtitle: "Final Consolidation", color: C.NAVY, lessonInfo: "Session 10" });
    y = addSectionHeading(doc, "A", y, { color: C.NAVY });
    y = addBodyText(doc, "1. TRUE (commutative)  2. TRUE (distributive)  3. FALSE (24 =/= 30)", y);
    y = addSectionHeading(doc, "B", y, { color: C.NAVY });
    y = addBodyText(doc, "4. ?=9  5. Bracket=7, ?+2=7, ?=5  6. 4x?=28, ?=7  7. ?+5=10, ?=5", y);
    y = addSectionHeading(doc, "C", y, { color: C.NAVY });
    y = addBodyText(doc, "8. 8+15=23  9. 11x5=55  10. 4+12=16", y);
    y = addSectionHeading(doc, "D", y, { color: C.NAVY });
    y = addBodyText(doc, "11. 8x(20+3)=160+24=184", y);
    y = addSectionHeading(doc, "E", y, { color: C.NAVY });
    y = addBodyText(doc, "12. (1,20)(2,10)(4,5)(5,4)(10,2)(20,1) = 6 pairs", y);
    y = addBodyText(doc, "13. (1,16)(2,14)(3,12)(4,10)(5,8)(6,6)(7,4)(8,2) -- any 3", y);
    addPdfFooter(doc, "Session 10 | Answer Key | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName)); console.log("PDF: " + ANSWER_KEY_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, { subtitle: "Full Strategy Reference", color: C.TEAL, lessonInfo: "Session 10" });
    y = addSectionHeading(doc, "Your Algebra Toolkit", y, { color: C.NAVY });
    y = addBodyText(doc, "EQUIVALENCE: Evaluate each side. Same value = equivalent.", y);
    y = addBodyText(doc, "COMMUTATIVE: a x b = b x a (swap the order).", y);
    y = addBodyText(doc, "ASSOCIATIVE: (a x b) x c = a x (b x c) (regroup factors).", y);
    y = addBodyText(doc, "DISTRIBUTIVE: a x (b+c) = a x b + a x c (split across +).", y);
    y = addBodyText(doc, "ORDER OF OPS: 1. Brackets  2. x /  3. + -", y);
    y = addBodyText(doc, "BACKWARDS: Total -> undo outer op -> solve inner.", y);
    y = addBodyText(doc, "PAIRS: Fix one unknown, calculate the other. Use a table.", y);
    y += 10;
    y = addSectionHeading(doc, "Guided Practice", y, { color: C.NAVY });
    y = addProblem(doc, 1, "TRUE/FALSE: 6x12 = 12x6. Strategy: ___. Answer: ___", y, { color: C.NAVY });
    y = addProblem(doc, 2, "Find ?: 5x(?+2) = 35. Strategy: ___. Answer: ___", y, { color: C.NAVY });
    y = addProblem(doc, 3, "Calculate: 8+3x5 = ___. Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 4, "8x23 using distributive = ___", y, { color: C.NAVY });
    y = addProblem(doc, 5, "Find 2 pairs for axb=20: ___", y, { color: C.NAVY });
    addPdfFooter(doc, "Session 10 | Enabling Scaffold | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName)); console.log("PDF: " + ENABLING_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, { subtitle: "Algebra Puzzle Challenge", color: C.NAVY, lessonInfo: "Session 10" });
    y = addSectionHeading(doc, "Puzzle 1: How Many Solutions?", y, { color: C.NAVY });
    y = addBodyText(doc, "Create an equation that has EXACTLY 3 natural number solutions.", y);
    y = addBodyText(doc, "Hint: Think about a multiplication equation with a number that has exactly 3 factor pairs.", y);
    y = addWriteLine(doc, "Your equation:", y);
    y = addWriteLine(doc, "The 3 solutions:", y);
    y = addSectionHeading(doc, "Puzzle 2: Impossible Equations", y, { color: C.NAVY });
    y = addBodyText(doc, "Create an equation that has NO natural number solutions.", y);
    y = addBodyText(doc, "Hint: What happens when the equation requires a fraction or decimal?", y);
    y = addWriteLine(doc, "Your equation:", y);
    y = addWriteLine(doc, "Why no solutions:", y);
    y = addSectionHeading(doc, "Puzzle 3: The Biggest Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "Find all natural number pairs (a, b, c) where:", y);
    y = addBodyText(doc, "a x b + c = 20  AND  a + b + c = 10", y);
    y = addBodyText(doc, "You need to satisfy BOTH equations with the SAME values!", y);
    y = addWriteLine(doc, "", y); y = addWriteLine(doc, "", y); y = addWriteLine(doc, "", y);
    y = addTipBox(doc, "This is a system of equations -- you will study these in high school! For now, use trial and improvement: try values that satisfy one equation, then check if they work in the other.", y, { color: C.TEAL });
    addPdfFooter(doc, "Session 10 | Extension | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName)); console.log("PDF: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 10 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
