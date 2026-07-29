"use strict";

// Algebra Unit — Session 5: Consolidation — Applying Properties to Solve Unknowns
// Week 1 Session 5, Grade 5/6 Numeracy, Variant 0
// DR: Compare, order and represent fractions
// Fluency: Doubling and halving speed drill
// VC2M5A02 — consolidation of all Level 5 algebra properties

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

const SESSION = 5;
const FOOTER = "Algebra | Session 5 of 10 | Grade 5/6 Numeracy";
const OUT_DIR = "output/ALG_Session5_Consolidation_Week1";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Mixed practice - all properties and unknown-finding strategies.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with worked answers.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Property reference card with guided steps.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Self-contained investigation into multi-step equations.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

const NOTES_TITLE = `SAY:
- Last session of Week 1. Today we bring it all together
- You have three properties in your toolkit: commutative, associative, and distributive
- You have a strategy for finding unknowns: evaluate, target, solve
- Today you use all of these on mixed problems

DO:
- Display title slide

TEACHER NOTES:
Session 5 of 10. Consolidation lesson -- no new content. Students apply all three properties and the evaluate-target-solve strategy to mixed problems. This builds fluency and confidence before Week 2 introduces brackets and order of operations.

WATCH FOR:
- Students who are confident and ready for more challenge
- Students who still need support with specific properties

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Daily review -- fractions today
- Compare these fractions using < > or =
- Use what you know about equivalent fractions and benchmarks

DO:
- Display fraction comparison problems
- Allow 90 seconds
- Circulate

TEACHER NOTES:
DR on fractions is a different strand -- genuine spaced retrieval. Comparing and ordering fractions uses equivalence thinking (same value, different form) which parallels the algebraic equivalence work.

WATCH FOR:
- Students who convert to common denominators vs those who use benchmark reasoning (e.g., both close to 1/2)

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers
- Ask: How did you compare 3/4 and 5/8? [Convert to 6/8 and 5/8, so 3/4 > 5/8]

DO:
- Reveal answers, students self-check

TEACHER NOTES:
Fraction equivalence connects to algebraic equivalence -- both involve recognising that different forms can represent the same value.

WATCH FOR:
- Students who used benchmarks effectively vs those who needed common denominators

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency -- doubling and halving
- I say a number, you write the double AND the half
- Ready? Go.

DO:
- Display 8 numbers for doubling and halving
- Time 2 minutes

TEACHER NOTES:
Doubling and halving connects to the commutative and associative properties (x2, /2 are key decomposition operations). Fast doubling/halving supports factor manipulation.

WATCH FOR:
- Students who double easily but struggle with halving odd numbers
- Students who are automatic on both

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me
- Read from slide: "We are learning to choose and apply the right property to solve unknown-value problems"
- Today there is no single strategy -- you choose which property to use
- Read our success criteria

DO:
- Choral read LI and SC

TEACHER NOTES:
The metacognitive demand increases -- students must select the appropriate property, not just apply a given one. This is the consolidation challenge.

WATCH FOR:
- Students who are confident with the vocabulary (commutative, associative, distributive)

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me solve a mixed problem. I need to choose the right tool.
- 6 x 8 = ? x 4
- Step 1: Left side = 48. Step 2: ? x 4 = 48. Step 3: ? = 12.
- But which property explains this? I can split 8 into 2 x 4:
- 6 x 8 = 6 x (2 x 4) = (6 x 2) x 4 = 12 x 4. Associative property!
- Now a different one: 7 x 14. I use distributive to make it easier.
- 7 x (10 + 4) = 70 + 28 = 98.
- The key is choosing the right tool for the job.

DO:
- Show both examples, naming the property each time
- Emphasise the decision-making process

TEACHER NOTES:
I Do models the metacognitive process of property selection. Thinking aloud about WHY you choose a particular property is the key teaching move.

WATCH FOR:
- Students who recognise which property before you name it
- Students who are still unclear on the differences between properties

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Harder one. Is this true or false? 4 x 15 = 12 x 5
- I could evaluate both sides: 4 x 15 = 60, 12 x 5 = 60. TRUE.
- But I can also PROVE it with properties:
- 4 x 15 = 4 x (3 x 5) = (4 x 3) x 5 = 12 x 5
- I used the associative property to show they are equivalent.
- Knowing the properties lets you prove equivalence without calculating both sides.

DO:
- Show both approaches: evaluate-and-compare vs property-based proof
- Highlight that properties give you a deeper understanding

TEACHER NOTES:
Moving from evaluation to proof is a significant conceptual step. Students who can prove equivalence using properties have genuine algebraic understanding, not just arithmetic skill.

WATCH FOR:
- Students who prefer the evaluate approach -- that is fine, but nudge toward property reasoning
- Students who attempt the property proof -- celebrate this

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. Which property would you use?
- 9 x 7 = 7 x 9. Name the property. Whisper it to your partner.

DO:
- Partners whisper, then you cold call
- Listen for: commutative

CFU CHECKPOINT:
Technique: Think-Pair-Share
Script:
- Say: "9 x 7 = 7 x 9. Which property? Whisper to your partner... now [Name], which property?"
- Scan for: commutative (swap the order).
PROCEED: If student answers correctly, move on.
PIVOT: If student says associative or distributive, reteach: "Look at what changed. The order swapped: 9 x 7 became 7 x 9. Same numbers, different order. That is commutative. Associative would be regrouping three or more factors."

TEACHER NOTES:
Quick identification check -- can students name properties by sight?

WATCH FOR:
- Students who confuse commutative and associative -- the distinction matters

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Mixed challenge with your partner
- For each problem: solve it AND name the property you used
- Some are unknowns, some are true/false checks

DO:
- Display problems
- Partners work on whiteboards
- 3 minutes

TEACHER NOTES:
We Do mixes problem types -- unknown-finding, equivalence checking, and property identification. This is the consolidation challenge.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 5 Enabling Scaffold with a property reference card. The card lists all three properties with examples. Students match each problem to the correct property before solving.
- Extra Notes: Distribute Session 5 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Multi-step equation investigation. Find the value of ? in chains like: 3 x 4 = ? x 2, then use that ? to solve ? x 5 = ?? x 10. Build a three-step chain.
- Extra Notes: Distribute Session 5 Extension PDF.

WATCH FOR:
- Students who can solve but struggle to name the property -- property vocabulary needs reinforcement
- Readiness signal: completing all four with correct property names

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check.
- 1: 5 x 12 = ? x 6. Left=60, ?x6=60, ?=10. Associative: 5x(2x6)=(5x2)x6=10x6.
- 2: TRUE or FALSE: 3 x 16 = 6 x 8. Left=48, Right=48. TRUE. Associative: 3x(2x8)=(3x2)x8=6x8.
- 3: 8 x ? = 56. ?=7. (Factor knowledge.)
- 4: Use distributive to calculate 6 x 23. 6x(20+3)=120+18=138.

DO:
- Reveal, cold call, discuss properties

TEACHER NOTES:
Problem 2 shows that what looks like a simple true/false check can be explained through associative reasoning. Problem 4 applies distributive as a calculation tool.

WATCH FOR:
- Students who found alternative property explanations -- validate these

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. Finger vote.
- Which property is used here? 4 x 13 = 4 x 10 + 4 x 3
- 1 = commutative, 2 = associative, 3 = distributive

DO:
- Students vote with fingers
- Scan for 3 (distributive)

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "1=commutative, 2=associative, 3=distributive. Which property: 4 x 13 = 4 x 10 + 4 x 3? Show me."
- Scan for: 3 fingers (distributive -- the multiplication is distributed across the addition).
PROCEED: If 80%+ show 3 fingers, move to You Do.
PIVOT: Most likely misconception -- students show 2 (associative) because they see the number being split. Reteach: "Associative is regrouping factors: (a x b) x c. Distributive is splitting across addition: a x (b+c) = a x b + a x c. Here 13 was split into 10+3, and the 4 was multiplied by EACH part. That is distributive."

TEACHER NOTES:
The distinction between associative and distributive is the key conceptual hurdle. This hinge tests it directly.

WATCH FOR:
- Students who confuse associative and distributive -- this is the most common error at this stage

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent mixed practice
- Your worksheet covers all three properties and unknown-finding
- First: Read each problem. Next: Choose your property/strategy. Then: Solve and name the property.
- 8 minutes.

DO:
- Distribute Session 5 Worksheet
- Circulate, support as needed

TEACHER NOTES:
You Do is a mixed assessment of all Week 1 skills. It uses different content from We Do.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the property reference card. Match problems to properties first, then solve.
- Extra Notes: Distribute Session 5 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Multi-step equation chains.
- Extra Notes: Distribute Session 5 Extension PDF.

WATCH FOR:
- Students who solve correctly but cannot name properties -- vocabulary gap
- Readiness signal: completing the mixed section with confidence

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Three questions. 3 minutes.

DO:
- Display questions, students work alone, collect

TEACHER NOTES:
Exit ticket assesses all three SC across property identification, application, and unknown-finding.

WATCH FOR:
- Students who master everything -- ready for Week 2
- Students who struggle -- plan targeted support before Session 6

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Week 1 wrap-up! Look at our success criteria.
- SC1: I can identify which property makes an equation true. Thumbs?
- SC2: I can find unknowns using evaluate-target-solve. Thumbs?
- SC3: I can justify my answer using a named property. Thumbs?
- Turn and talk: Which property do you find most useful? Why?

DO:
- Thumbs for each SC
- Turn and Talk
- Cold call 2-3 students
- Preview: "Next week we add brackets and order of operations!"

TEACHER NOTES:
End of Week 1 consolidation. The Turn and Talk targets metacognitive reflection on property utility. Preview of Week 2 builds anticipation.

WATCH FOR:
- Students who name distributive as most useful -- it has the strongest practical application
- Overall confidence level across the class

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Resources linked here

DO:
- Point out resources

TEACHER NOTES:
Resource slide.

WATCH FOR:
- N/A

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ───────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, "Algebra: Finding Unknown Values", "Session 5: Consolidation — All Properties", "Grade 5/6 Numeracy | Session 5 of 10 | Week 1", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // DR — Fractions
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Compare and Order Fractions", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const probs = [
        "1.  3/4  ___  5/8     (< > or =)",
        "2.  2/3  ___  4/6     (< > or =)",
        "3.  Order from smallest: 1/2, 3/8, 5/6",
        "4.  Which is closer to 1: 7/8 or 4/5?",
      ];
      s.addText(probs.map((p, i) => ({
        text: p, options: { fontSize: 14, color: C.CHARCOAL, breakLine: i < probs.length - 1, paraSpaceAfter: 14 },
      })), { x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3, fontFace: FONT_B, margin: 0, valign: "top" });
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboards:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "Convert to common denominators or use benchmarks.", options: { fontSize: 14, color: C.CHARCOAL } },
      ], { x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3, fontFace: FONT_B, margin: 0, valign: "top" });
      addFooter(s, FOOTER); s.addNotes(NOTES_DR_Q); return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) >  2) =  3) 3/8, 1/2, 5/6  4) 7/8 (1/8 away vs 1/5 away)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Fluency — Doubling and halving
  const sFl = pres.addSlide();
  addTopBar(sFl, STAGE_COLORS["1"]);
  addStageBadge(sFl, 1, "Fluency");
  addTitle(sFl, "Doubling and Halving", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });
  addCard(sFl, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  const nums = ["1.  36:  Double = ___  Half = ___", "2.  48:  Double = ___  Half = ___", "3.  75:  Double = ___  Half = ___", "4.  64:  Double = ___  Half = ___", "5.  55:  Double = ___  Half = ___", "6.  120: Double = ___  Half = ___", "7.  96:  Double = ___  Half = ___", "8.  250: Double = ___  Half = ___"];
  sFl.addText(nums.map((p, i) => ({
    text: p, options: { fontSize: 13.5, color: C.CHARCOAL, breakLine: i < nums.length - 1, paraSpaceAfter: 4 },
  })), { x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2, fontFace: FONT_B, margin: 0, valign: "top" });
  addCard(sFl, 5.2, CONTENT_TOP, 4.3, 1.8, { strip: C.SECONDARY });
  sFl.addText([
    { text: "Whiteboards - 2 minutes", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
    { text: "Double AND half each number.", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "72/18, 96/24, 150/37.5, 128/32, 110/27.5, 240/60, 192/48, 500/125", options: { fontSize: 9, color: C.MUTED } },
  ], { x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: 1.6, fontFace: FONT_B, margin: 0, valign: "top" });
  addFooter(sFl, FOOTER); sFl.addNotes(NOTES_FLUENCY);

  // LI/SC
  liSlide(pres,
    ["We are learning to choose and apply the right property to solve unknown-value problems"],
    [
      "I can identify which property makes an equation true",
      "I can find unknown values using the evaluate-target-solve strategy",
      "I can justify my answer by naming and explaining the property used",
    ],
    NOTES_LI_SC, FOOTER);

  // I Do 1
  workedExSlide(pres, 2, "I Do", "Choosing the Right Tool",
    [
      "Problem A: 6 x 8 = ? x 4",
      "  Evaluate: 6 x 8 = 48",
      "  Target: ? x 4 = 48",
      "  Solve: ? = 12",
      "  Property: Associative (split 8 = 2x4, regroup)",
      "",
      "Problem B: 7 x 14",
      "  Use distributive: 7 x (10+4) = 70+28 = 98",
      "",
      "Choose the tool that fits the problem!",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("Your Toolkit", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      const tools = [
        { name: "Commutative", desc: "Swap the order: a x b = b x a", col: C.PRIMARY },
        { name: "Associative", desc: "Regroup factors: (axb)xc = ax(bxc)", col: C.SECONDARY },
        { name: "Distributive", desc: "Split across +: ax(b+c) = axb+axc", col: C.ACCENT },
        { name: "Evaluate-Target-Solve", desc: "Calculate, set target, find unknown", col: C.ALERT },
      ];
      tools.forEach((t, i) => {
        const ty = lg.panelTopPadded + 0.42 + i * 0.6;
        addTextOnShape(slide, t.name, {
          x: lg.rightX + 0.15, y: ty, w: 1.9, h: 0.32, rectRadius: 0.06,
          fill: { color: t.col },
        }, { fontSize: 10, fontFace: FONT_B, color: C.WHITE, bold: true });
        slide.addText(t.desc, {
          x: lg.rightX + 2.15, y: ty, w: lg.rightW - 2.3, h: 0.32,
          fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "middle",
        });
      });
    }
  );

  // I Do 2
  workedExSlide(pres, 2, "I Do", "Proving Equivalence with Properties",
    [
      "Is 4 x 15 = 12 x 5 true?",
      "",
      "Method 1 (evaluate):",
      "  Left: 4 x 15 = 60. Right: 12 x 5 = 60. TRUE.",
      "",
      "Method 2 (prove with property):",
      "  4 x 15 = 4 x (3 x 5)",
      "         = (4 x 3) x 5",
      "         = 12 x 5",
      "  Associative property proves it!",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.4, { strip: C.SUCCESS });
      slide.addText("Method 1: Calculate", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 13, fontFace: FONT_H, color: C.SUCCESS, bold: true, margin: 0,
      });
      slide.addText("4 x 15 = 60\n12 x 5 = 60\n60 = 60 -> TRUE", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.35, w: lg.rightW - 0.6, h: 0.75,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });

      addCard(slide, lg.rightX, lg.panelTopPadded + 1.6, lg.rightW, 1.6, { strip: C.SECONDARY });
      slide.addText("Method 2: Prove with Property", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.66, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
      });
      slide.addText("4 x 15\n= 4 x (3 x 5)\n= (4 x 3) x 5\n= 12 x 5\nAssociative property!", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.95, w: lg.rightW - 0.6, h: 0.95,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });
    }
  );

  // CFU 1
  withReveal(
    () => cfuSlide(pres, "CFU", "Property Identification", "Think-Pair-Share",
      "Which property?\n\n9 x 7 = 7 x 9",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "COMMUTATIVE — the order was swapped", {
        x: 1.5, y: 4.0, w: 7, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // We Do
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Mixed Challenge",
      [
        "Solve and name the property:",
        "",
        "1.   5 x 12 = ? x 6   (find ? and name property)",
        "2.   TRUE/FALSE: 3 x 16 = 6 x 8  (name property)",
        "3.   8 x ? = 56   (find ?)",
        "4.   Use distributive: 6 x 23 = ?",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.2, { strip: C.SECONDARY });
        slide.addText("Toolkit:", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
          fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        const tips = ["Commutative: swap order", "Associative: regroup factors", "Distributive: split across +", "Evaluate-Target-Solve: for unknowns"];
        slide.addText(tips.map((t, i) => ({
          text: t, options: { bullet: true, fontSize: 12, color: C.CHARCOAL, breakLine: i < tips.length - 1 },
        })), { x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.38, w: lg.rightW - 0.5, h: 1.4, fontFace: FONT_B, margin: 0, valign: "top" });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "1) ?=10 (assoc.)  2) TRUE (assoc.)  3) ?=7  4) 138 (distrib.)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08, fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // CFU 2 Hinge
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge — Which Property?", "Finger Voting",
      "Which property?\n\n4 x 13 = 4 x 10 + 4 x 3\n\n1 = commutative\n2 = associative\n3 = distributive",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "3 — DISTRIBUTIVE: 4 is distributed across the addition (10 + 3)", {
        x: 0.5, y: 4.0, w: 9, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // You Do
  workedExSlide(pres, 4, "You Do", "Independent Mixed Practice",
    [
      "First: Read each problem on your worksheet.",
      "Next: Choose your property or strategy.",
      "Then: Solve and name the property.",
      "",
      "This is mixed practice -- use everything!",
      "8 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.5, { strip: C.ALERT });
      slide.addText("Week 1 Toolkit:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.25,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Commutative (swap)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Associative (regroup)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Distributive (split)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Evaluate-Target-Solve", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], { x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.38, w: lg.rightW - 0.5, h: 0.9, fontFace: FONT_B, margin: 0, valign: "top" });
    }
  );

  exitTicketSlide(pres,
    [
      "Name the property: 8 x 9 = 9 x 8",
      "Find ?: 4 x 9 = ? x 3  (show working and name the property)",
      "Use the distributive property to calculate 5 x 18. Show full working.",
    ],
    NOTES_EXIT, FOOTER);

  closingSlide(pres,
    "Which property do you find most useful in real life? Tell your partner why.",
    [
      "I can identify which property makes an equation true",
      "I can find unknown values using evaluate-target-solve",
      "I can justify my answer with a named property",
    ],
    NOTES_CLOSING);



  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "ALG_Session5_Consolidation_Week1.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // PDFs
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, { subtitle: "Mixed Properties Practice", color: C.NAVY, lessonInfo: "Session 5 of 10 | Grade 5/6" });
    y = addTipBox(doc, "For each problem: solve it AND name the property you used. Use your evaluate-target-solve strategy for unknowns.", y, { color: C.TEAL });
    y = addSectionHeading(doc, "Section A: Name the Property", y, { color: C.NAVY });
    y = addProblem(doc, 1, "11 x 5 = 5 x 11                Property: ___", y, { color: C.NAVY });
    y = addProblem(doc, 2, "3 x 2 x 7 = 6 x 7              Property: ___", y, { color: C.NAVY });
    y = addProblem(doc, 3, "5 x 14 = 5 x 10 + 5 x 4        Property: ___", y, { color: C.NAVY });
    y = addProblem(doc, 4, "8 x 7 = 7 x 8                   Property: ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "Section B: Find the Unknown", y, { color: C.NAVY });
    y = addProblem(doc, 5, "4 x 9 = ? x 6              ? = ___  Property: ___", y, { color: C.NAVY });
    y = addProblem(doc, 6, "7 x 8 = 112 / ?            ? = ___", y, { color: C.NAVY });
    y = addProblem(doc, 7, "5 x 6 = ? x 10             ? = ___  Property: ___", y, { color: C.NAVY });
    y = addProblem(doc, 8, "? x 9 = 63                  ? = ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "Section C: Apply Distributive", y, { color: C.NAVY });
    y = addProblem(doc, 9, "8 x 17 = 8 x (___+___) = ___+___ = ___", y, { color: C.NAVY });
    y = addProblem(doc, 10, "6 x 25 = 6 x (___+___) = ___+___ = ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "Section D: TRUE or FALSE? Prove it.", y, { color: C.NAVY });
    y = addProblem(doc, 11, "3 x 14 = 7 x 6      TRUE / FALSE   Property: ___", y, { color: C.NAVY });
    y = addProblem(doc, 12, "5 x 8 = 4 x 10      TRUE / FALSE   Explain:", y, { color: C.NAVY });
    y = addWriteLine(doc, "", y);
    addPdfFooter(doc, "Session 5 | Algebra | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, { subtitle: "Mixed Properties Practice", color: C.NAVY, lessonInfo: "Session 5 of 10 | Grade 5/6" });
    y = addSectionHeading(doc, "Section A", y, { color: C.NAVY });
    y = addBodyText(doc, "1. Commutative  2. Associative  3. Distributive  4. Commutative", y);
    y = addSectionHeading(doc, "Section B", y, { color: C.NAVY });
    y = addBodyText(doc, "5. 4x9=36, ?x6=36, ?=6 (Associative: 4x(3x3)x... or eval-target-solve)", y);
    y = addBodyText(doc, "6. 7x8=56, 112/?=56, ?=2", y);
    y = addBodyText(doc, "7. 5x6=30, ?x10=30, ?=3 (Associative)", y);
    y = addBodyText(doc, "8. ?x9=63, ?=7", y);
    y = addSectionHeading(doc, "Section C", y, { color: C.NAVY });
    y = addBodyText(doc, "9. 8x(10+7)=80+56=136", y);
    y = addBodyText(doc, "10. 6x(20+5)=120+30=150", y);
    y = addSectionHeading(doc, "Section D", y, { color: C.NAVY });
    y = addBodyText(doc, "11. 3x14=42, 7x6=42. TRUE. Associative: 3x(2x7)=(3x2)x7=6x7=42", y);
    y = addBodyText(doc, "12. 5x8=40, 4x10=40. TRUE. Both equal 40.", y);
    addPdfFooter(doc, "Session 5 | Answer Key | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF: " + ANSWER_KEY_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, { subtitle: "Property Reference Card", color: C.TEAL, lessonInfo: "Session 5 of 10 | Grade 5/6" });
    y = addSectionHeading(doc, "Your Three Properties", y, { color: C.NAVY });
    y = addBodyText(doc, "COMMUTATIVE: Swap the order. a x b = b x a", y);
    y = addBodyText(doc, "Example: 3 x 7 = 7 x 3 = 21", y); y += 5;
    y = addBodyText(doc, "ASSOCIATIVE: Regroup the factors. (a x b) x c = a x (b x c)", y);
    y = addBodyText(doc, "Example: 2 x 3 x 5 = (2 x 3) x 5 = 6 x 5 = 30", y); y += 5;
    y = addBodyText(doc, "DISTRIBUTIVE: Split across addition. a x (b + c) = a x b + a x c", y);
    y = addBodyText(doc, "Example: 4 x 13 = 4 x (10 + 3) = 40 + 12 = 52", y); y += 10;
    y = addSectionHeading(doc, "Strategy for Finding Unknowns", y, { color: C.NAVY });
    y = addBodyText(doc, "1. EVALUATE the known side", y);
    y = addBodyText(doc, "2. TARGET: what must the other side equal?", y);
    y = addBodyText(doc, "3. SOLVE for the missing number", y);
    y = addBodyText(doc, "4. CHECK: do both sides balance?", y); y += 10;
    y = addSectionHeading(doc, "Practice with the Reference Card", y, { color: C.NAVY });
    const probs = ["6 x 8 = ? x 4", "9 x 3 = 3 x 9", "7 x 12 = 7 x (10 + 2) = ?", "4 x 5 = ? x 2"];
    probs.forEach((p, i) => {
      y = addProblem(doc, i + 1, p + "    Property: ___", y, { color: C.NAVY });
    });
    addPdfFooter(doc, "Session 5 | Enabling Scaffold | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF: " + ENABLING_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, { subtitle: "Multi-Step Equation Chains", color: C.NAVY, lessonInfo: "Session 5 of 10 | Grade 5/6" });
    y = addSectionHeading(doc, "What Is an Equation Chain?", y, { color: C.NAVY });
    y = addBodyText(doc, "An equation chain links multiple equations together. The answer from one equation becomes part of the next.", y);
    y = addSectionHeading(doc, "Worked Example", y, { color: C.NAVY });
    y = addBodyText(doc, "Step 1: 3 x 4 = ? x 2.  Left=12, ?x2=12, ?=6.", y);
    y = addBodyText(doc, "Step 2: Use that 6. Now: 6 x 5 = ? x 10.  Left=30, ?x10=30, ?=3.", y);
    y = addBodyText(doc, "Step 3: Use that 3. Now: 3 x 8 = ? x 4.  Left=24, ?x4=24, ?=6.", y);
    y = addBodyText(doc, "The chain: 3 x 4 -> 6 -> 6 x 5 -> 3 -> 3 x 8 -> 6.", y);
    y = addSectionHeading(doc, "Build Your Own Chain", y, { color: C.NAVY });
    y = addBodyText(doc, "Start with: 2 x 6 = ? x 3. Find ?.", y);
    y = addWriteLine(doc, "Step 1: ? = ___", y);
    y = addBodyText(doc, "Now use your answer. Write an equation where it appears: ___ x ___ = ? x ___", y);
    y = addWriteLine(doc, "Step 2: ? = ___", y);
    y = addBodyText(doc, "Use THAT answer for a third equation:", y);
    y = addWriteLine(doc, "Step 3: ? = ___", y);
    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "Can you build a chain of 5 steps where the final answer equals the starting number?", y);
    y = addBodyText(doc, "That would create a loop! Try it:", y);
    y = addWriteLine(doc, "", y); y = addWriteLine(doc, "", y); y = addWriteLine(doc, "", y);
    y = addTipBox(doc, "Hint: If you double and then halve (or triple and then divide by 3), you get back to where you started. Use this to close the loop.", y, { color: C.TEAL });
    addPdfFooter(doc, "Session 5 | Extension | Grade 5/6");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 5 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
