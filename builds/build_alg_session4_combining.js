"use strict";

// Algebra Unit — Session 4: Combining Properties to Find Unknowns
// Session 4 of 6, Grade 5/6 Numeracy, Variant 0
// DR: Rounding in Addition & Subtraction
// Fluency: Number Talk (24 x 5)
// VC2M5A02 elaboration 5: 3 x 4 = 6 x 2 via factor decomposition + associative property

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
  C, FONT_H, FONT_B,
  SLIDE_W, SLIDE_H, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addBadge, addTitle,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 4;
const UNIT_TITLE = "Algebra: Finding Unknown Values";
const FOOTER = "Session 4 of 6 | Algebra: Unknown Values | Year 5/6 Maths";
const OUT_DIR = "output/ALG6_Session4_Combining_Properties";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "6 problems using factor decomposition and combined properties.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Step-by-step framework provided for each problem.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Design your own equations where properties work cleanly.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ───────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- We now have three tools: equivalence, commutative/associative, and distributive. Today we combine them to crack harder equations.
- You will decompose factors and use the associative property to find unknowns that are not obvious.

DO:
- Display the title slide. Whiteboards ready.

TEACHER NOTES:
Lesson 4 synthesises Lessons 1-3. The curriculum reference is VC2M5A02 elaboration 5: "considering 3 x 4 = 12 and knowing 2 x 2 = 4, then 3 x 4 can be written as 3 x (2 x 2) and, using the associative property, (3 x 2) x 2 so 3 x 4 = 6 x 2 and so 6 is the solution to 3 x 4 = [] x 2." This is the most algebraically demanding lesson in the first half of the unit.

WATCH FOR:
- Students who struggled with L2 (associative) or L3 (distributive) exit tickets -- they may need pre-teaching.

[Maths: Planning | VTLM 2.0: Planning]`;

const NOTES_DR_Q = `SAY:
- Daily Review: Rounding in addition and subtraction.
- Problem 1: Round 347 + 582 to the nearest hundred, then calculate the estimate. [300 + 600 = 900]
- Problem 2: 8,256 - 3,471. Estimate first by rounding to thousands. [8,000 - 3,000 = 5,000]
- Problem 3: 23.7 + 46.8. Round to the nearest whole number and estimate. [24 + 47 = 71]

DO:
- Display the slide. Students work on whiteboards, 15 seconds per problem. "Show me!"
- Emphasise: "An estimate uses rounded numbers. It is NOT the exact answer -- it is a quick check."

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Round and estimate on your board. Show me!"
- Scan for: reasonable estimates on 80%+ of boards.
PROCEED: If 80%+ correct estimates, move to Fluency.
PIVOT: If students calculate exactly instead of rounding, reteach: "Rounding means replacing each number with its nearest hundred (or thousand). 347 rounds to 300. 582 rounds to 600. THEN add: 300 + 600 = 900. That is the estimate." Re-check: "Estimate 678 + 215 by rounding to hundreds."

TEACHER NOTES:
Daily Review targets: "Four Processes -- I can use rounding in addition and subtraction calculations." This is retrieval practice from prior learning and connects to estimation strategies students will use to verify algebraic solutions.

WATCH FOR:
- Students who calculate exactly instead of estimating -- they may not understand the purpose of rounding.
- Students who round incorrectly (347 -> 400 instead of 300) -- review the rounding rule.

[Maths: Daily Review | VTLM 2.0: Retention & Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers -- tick if correct, fix if not.
- 1) 300 + 600 = 900. 2) 8,000 - 3,000 = 5,000. 3) 24 + 47 = 71.

DO:
- Click to reveal answers. Students self-check.

TEACHER NOTES:
Quick tick-and-fix. Move to Fluency.

WATCH FOR:
- Students who got all three -- estimation skills are solid.

[Maths: Daily Review Answers | VTLM 2.0: Retention & Recall]`;

const NOTES_FLUENCY = `SAY:
- Number Talk: I will show you a multiplication. You solve it mentally and explain your strategy.
- 24 x 5. Think about it. No pencil -- mental only. When you have an answer and a strategy, show me a quiet thumbs up.
- Who wants to share? What did you get and HOW?
- Possible strategies: 24 x 5 = 12 x 10 = 120 (doubling/halving); 20 x 5 + 4 x 5 = 120 (distributive); 24 x 10 / 2 = 120.
- Multiple strategies, same answer. That is the power of number properties.

DO:
- Display the problem. Wait for thumbs up (no rushing).
- Take 3-4 strategies. Record them briefly on the board or verbally.
- Second problem if time: "36 x 5."

TEACHER NOTES:
Number Talks are a different fluency format from the sprint/chain activities in previous lessons. The mental strategies students share are exactly the properties they have been studying: doubling/halving uses the associative property (24 x 5 = 12 x 2 x 5 = 12 x 10), the break-apart strategy uses the distributive property. Naming these connections reinforces that "number properties" are not abstract rules -- they are the strategies students already use.

WATCH FOR:
- Students who can only use one strategy -- encourage: "Can you think of a different way?"
- Students who share a strategy that uses a property -- name it: "You just used the distributive property!"

[Maths: Fluency | VTLM 2.0: Retention & Recall]`;

const NOTES_LI_SC = `SAY:
- Read from slide: "We are learning to use factors and number properties to find unknown values in multiplication equations."
- SC1 is about decomposing -- breaking a factor into its own factors. SC2 is the big one -- using the associative property to rearrange and find []. SC3 asks you to explain your reasoning.

DO:
- Display the slide. Point to each SC.
- "SC2 is our main target."

TEACHER NOTES:
The LI addresses VC2M5A02 elaboration 5 directly. SC1 (factor decomposition) is the prerequisite skill from the NP unit -- students need to identify that 4 = 2 x 2 before they can rewrite 3 x 4 as 3 x (2 x 2). SC2 (finding unknowns using associativity) is the core new learning. SC3 (explaining reasoning) extends to mathematical communication.

WATCH FOR:
- Students who do not connect "decomposing factors" to their prior learning on factor pairs.

[Maths: Planning -- Curriculum Alignment | VTLM 2.0: Planning]`;

const NOTES_IDO1 = `SAY:
- Watch me solve: 3 x 4 = [] x 2. What is []?
- I could just calculate: 3 x 4 = 12, then 12 / 2 = 6, so [] = 6. But today I want to show you WHY [] = 6 using properties.
- I know that 4 = 2 x 2. So I can rewrite 3 x 4 as 3 x (2 x 2).
- Now I use the associative property to regroup: 3 x (2 x 2) = (3 x 2) x 2 = 6 x 2.
- So 3 x 4 = 6 x 2, which means [] = 6.
- Why did I decompose 4 into 2 x 2 specifically? Because the equation has x 2 on the right side. I needed a 2 in my decomposition to match.
- Verify: 3 x 4 = 12. 6 x 2 = 12. Both sides equal 12. Yes.

DO:
- Display the slide with step-by-step working.
- Emphasise the factor decomposition step (4 = 2 x 2) and the regrouping step.
- Circle the matching x 2 on both sides.

TEACHER NOTES:
This is the exact example from VC2M5A02 elaboration 5. The think-aloud makes two critical moves visible: (1) the strategic decomposition of 4 into 2 x 2 to match the x 2 in the target equation, and (2) the associative regrouping to isolate [].

MISCONCEPTIONS:
- Misconception: "I can always just divide to find the unknown -- I do not need properties."
  Why: Division IS valid for these equations. But properties explain WHY the unknown has a particular value and generalise to more complex equations where simple division does not work.
  Impact: Students who only rely on calculation will struggle with multi-step equations in Lessons 5-6.
  Quick correction: "Division gives you the answer. Properties explain the REASON. Both matter -- today we focus on the reason."

WATCH FOR:
- Students who immediately say "[] = 6" by dividing 12 / 2 -- validate but redirect to the property-based method.
- Students confused by the factor decomposition step -- "4 = 2 x 2" may seem like a strange move if they do not see the strategic purpose.

[Maths: Launch -- Explicit Instruction (I Do) | VTLM 2.0: Explicit Explanation & Modelling]`;

const NOTES_IDO2 = `SAY:
- Here is another: 5 x 6 = [] x 3. Find [] using properties.
- The right side has x 3. So I need to find a 3 inside the left side. What factor of 6 is 3? 6 = 3 x 2. So I decompose: 5 x 6 = 5 x (3 x 2).
- Regroup: 5 x (3 x 2). I need the 3 on the outside. Use commutativity first: 3 x 2 = 2 x 3. So 5 x (2 x 3).
- Now associative: (5 x 2) x 3 = 10 x 3.
- So 5 x 6 = 10 x 3. [] = 10.
- I used THREE properties here: decomposition (6 = 3 x 2), commutativity (3 x 2 = 2 x 3), and associativity (regroup to isolate x 3). They work together.
- Verify: 5 x 6 = 30. 10 x 3 = 30. Yes.

DO:
- Display step-by-step. Show each property labelled.
- Emphasise: "I looked at the right side (x 3) and worked BACKWARDS to create it."

TEACHER NOTES:
This second I Do demonstrates the full three-property strategy. The key insight is working backwards from the target equation structure -- if the right side has x 3, decompose the left side to produce a x 3 factor. The explicit labelling of each property shows students the toolkit they are using.

WATCH FOR:
- Students who look overwhelmed by three properties in one problem -- reassure: "Each step is simple. The power is using them together."
- Students who can follow but could not replicate -- the We Do provides guided practice.

[Maths: Launch -- Explicit Instruction (I Do) | VTLM 2.0: Explicit Explanation & Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. What factor decomposition would you use to solve: 4 x 6 = [] x 3?
- On your whiteboard, write how you would decompose 6 to match the x 3. You have 15 seconds.

DO:
- Display the question. 15 seconds think time. "Show me!"
- Scan for 6 = 3 x 2 or equivalent on boards.
- Cold call for the final answer.

CFU CHECKPOINT:
Technique: Show Me Boards + Cold Call
Script:
- "Write the factor decomposition on your board. Show me!"
- Scan for: 6 = 3 x 2 (or 2 x 3) on 80%+ of boards.
PROCEED: If 80%+ identify the correct decomposition, students understand the strategic step. Move to We Do.
PIVOT: If students decompose 6 as 1 x 6 or just write "6," they do not see the strategic purpose. Reteach: "Look at the right side: [] x 3. I need a 3 in my decomposition. What are the factors of 6 that include 3? 3 x 2 = 6. That is the decomposition I need." Re-check: "8 x 9 = [] x 3. How would you decompose 9?"

TEACHER NOTES:
This CFU isolates the strategic decomposition step -- the most challenging part of the combined-property method. If students can identify the right decomposition, they have the key to the entire procedure.

WATCH FOR:
- Students who write 6 = 1 x 6 -- valid decomposition but strategically useless for this equation.
- Students who correctly write 6 = 3 x 2 but cannot proceed to [] = 8 -- they need the regrouping modelled again.

[Maths: Monitor Progress | VTLM 2.0: Monitor Progress (CFU)]`;

const NOTES_CFU_A = `SAY:
- 6 = 3 x 2 or 6 = 2 x 3. Both work because we need a 3 to match x 3 on the right.
- Follow-up: So what would [] be? [8, because 4 x (2 x 3) = (4 x 2) x 3 = 8 x 3]

DO:
- Click to reveal. Show full working.

TEACHER NOTES:
The full solution path: 6 = 2 x 3 -> 4 x (2 x 3) -> (4 x 2) x 3 = 8 x 3 -> [] = 8.

WATCH FOR:
- Students who got the decomposition right but not the final answer -- they need more practice on regrouping.

[Maths: Monitor Progress | VTLM 2.0: Monitor Progress (CFU)]`;

const NOTES_WEDO1_Q = `SAY:
- Together: 2 x 8 = [] x 4. Find [] using properties.
- Ask: What is on the right side? [[] x 4] So I need a 4 inside the left side.
- Ask: How can we decompose 8 to include a 4? [8 = 4 x 2]
- So 2 x 8 = 2 x (4 x 2). Now regroup: (2 x 4) x 2? No, wait -- I need x 4 on the outside.
- Ask: Should I swap the factors inside first? [Yes -- 4 x 2 = 2 x 4, so 2 x (2 x 4)]
- Now associate: (2 x 2) x 4 = 4 x 4. So [] = 4.
- Verify: 2 x 8 = 16. 4 x 4 = 16. Yes.

DO:
- Display the equation. Use Cold Call for each step. Click to reveal.

CFU CHECKPOINT:
Technique: Cold Call
Script:
- "[Name], what factor decomposition do we use?" "[Name], what is the regrouped form?"
PROCEED: If students provide correct steps, move to PP2.
PIVOT: If students cannot decompose 8 to include 4, reteach: "What times what gives 8? 1x8, 2x4, 4x2. Which one has a 4? 4 x 2 or 2 x 4. Use that." Re-check: "6 x 10 = [] x 5. How do you decompose 10?"

TEACHER NOTES:
This problem pair mirrors I Do 1 in structure but with different numbers. The deliberate pause at "wait -- I need x 4 on the outside" models the metacognitive checking that students must do when deciding how to regroup.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide a partially completed table: "2 x 8 = [] x 4. Step 1: 8 = ___ x ___. Step 2: 2 x (___ x ___). Step 3: (___ x ___) x 4. Step 4: [] = ___." Students fill blanks.
EXTENDING PROMPT:
- Task: "Find []: 3 x 8 = [] x 6. Hint: you might need to decompose BOTH factors. Can you find more than one way?"
- Extra Notes: This requires more creative decomposition -- multiple solution paths exist.

WATCH FOR:
- Students who decompose 8 but then cannot regroup -- the associative step is the hurdle.
- Students who get [] = 4 and say "that is the same as 4 x 4 -- obvious!" -- they are seeing the pattern.

[Maths: Explore -- Guided Practice (We Do) | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO1_A = `SAY:
- 8 = 2 x 4 -> 2 x (2 x 4) -> (2 x 2) x 4 = 4 x 4 -> [] = 4.

DO:
- Click to reveal the solution.

TEACHER NOTES:
Confirm the answer and verification.

WATCH FOR:
- Students who see the answer -- ready for PP2.

[Maths: Explore -- Guided Practice (We Do) | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO2_Q = `SAY:
- Boards: 6 x 4 = [] x 8. Find [] using properties. You have 60 seconds.

DO:
- Display equation. 60 seconds on boards. Reveal.
- Discuss both approaches: property-based and calculate-then-decompose.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Write [] on your board. Show me!"
- Scan for: [] = 3 on 80%+.
PROCEED: If 80%+ correct, move to hinge question.
PIVOT: If students write [] = 48 (6 x 4 x 8 -- multiplying all three), reteach: "The equation says [] x 8 = 6 x 4 = 24. If [] x 8 = 24, then [] = 24 / 8 = 3." Re-check: "4 x 5 = [] x 10. What is []?"

TEACHER NOTES:
This problem is deliberately harder -- the decomposition path is not as clean as previous examples. The "calculate then decompose" approach (find the product, then express it as [] x target) is a valid and often faster strategy.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Students use the calculate-first approach: "6 x 4 = ___. ___ / 8 = ___. So [] = ___."
EXTENDING PROMPT:
- Task: "Find []: 5 x 12 = [] x 15. Can you use properties? When is the calculate-first method better?"
- Extra Notes: The Session 4 Extension PDF explores when each method is more efficient.

WATCH FOR:
- Students stuck because the property path is messy -- validate: "Sometimes calculating first IS the smart move."
- Students who solve both ways -- excellent mathematical reasoning.

[Maths: Explore -- Guided Practice (We Do) | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO2_A = `SAY:
- 6 x 4 = 24. 24 / 8 = 3. [] = 3. Verify: 3 x 8 = 24. Yes.
- Sometimes the fastest path is: calculate the product, then decompose into the target form.

DO:
- Click to reveal. Discuss.

TEACHER NOTES:
Both the property-based and calculate-first approaches give [] = 3. Building flexibility between methods is key.

WATCH FOR:
- Students who solved both ways -- excellent.

[Maths: Explore -- Guided Practice (We Do) | VTLM 2.0: Scaffold Practice]`;

const NOTES_HINGE_Q = `SAY:
- Gate check. What is [] in: 4 x 9 = [] x 6? Show me fingers: A) 36, B) 6, C) 4, D) 2.

DO:
- Display options. 15 seconds. "Show fingers." Scan. Reveal.

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- "What is []? A=36, B=6, C=4, D=2. Fingers NOW."
- Scan for: B on 80%+.
PROCEED: If 80%+ choose B, ready for You Do.
PIVOT: If A (students multiplied 4 x 9 but forgot to divide), reteach: "4 x 9 = 36. But the equation says [] x 6 = 36. So [] = 36 / 6 = 6." Re-check: "3 x 8 = [] x 4. What is []?"

TEACHER NOTES:
Distractors: A (product without dividing), C (4 appears in the equation), D (common small-number guess). This hinge tests whether students can complete the full process.

WATCH FOR:
- Students choosing A (36) -- they found the product but did not finish.
- Quick correct answers -- ready for independent work.

[Maths: Monitor Progress | VTLM 2.0: Monitor Progress (CFU)]`;

const NOTES_HINGE_A = `SAY:
- B -- [] = 6. Because 4 x 9 = 36 and 36 / 6 = 6. OR: using properties, 9 = 3 x 3 but the property path is complex here. The key is: 4 x 9 = 36 = 6 x 6. So [] = 6.

DO:
- Click to reveal answer and explanation.

TEACHER NOTES:
This problem works more cleanly with calculate-first. Validating both approaches builds flexibility.

WATCH FOR:
- Students who chose A and now understand -- progress.

[Maths: Monitor Progress | VTLM 2.0: Monitor Progress (CFU)]`;

const NOTES_YOUDO = `SAY:
- Read from slide: "First: Decide your approach (properties or calculate-first). Next: Find []. Then: Verify and explain which properties you used."
- Complete all problems on your worksheet.

DO:
- Distribute Session 4 Worksheet. Circulate. Focus on students' reasoning, not just answers.

TEACHER NOTES:
The worksheet has 6 graduated problems. Problems 1-2 have clean decomposition paths (matching the I Do). Problems 3-4 work better with calculate-first. Problems 5-6 challenge students to explain which approach they chose and why.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Session 4 Enabling Scaffold with the strategy steps pre-structured for each problem.
EXTENDING PROMPT:
- Task: Session 4 Extension investigation on creating equations: "Make up 3 equations of the form a x b = [] x c where the properties path works cleanly. What makes a 'nice' equation?"

WATCH FOR:
- Students who calculate answers but cannot explain using properties -- push for reasoning.

[Maths: Summarise -- Independent Practice (You Do) | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Two questions. 3 minutes.

DO:
- Display exit ticket. Students work silently. Collect after 3 minutes.

TEACHER NOTES:
Q1: Find [] using properties (assess SC2). Q2: Explain reasoning (assess SC3). Data informs Lesson 5 -- students who can combine properties will transition smoothly to order of operations.

WATCH FOR:
- Students who can find the answer but cannot name the properties used.

[Maths: Summarise -- Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Success criteria check.
- SC1: I can identify factors of a number to rewrite a multiplication. Thumbs.
- SC2: I can use the associative property to find an unknown value. Thumbs.
- SC3: I can explain my reasoning using mathematical properties. Thumbs.
- Turn and talk: When would you use properties vs calculate-first to find an unknown?
- Tomorrow we step up to equations with MULTIPLE operations and brackets -- order of operations.

DO:
- Display closing slide. Run thumbs for each SC.

TEACHER NOTES:
The forward look to Lesson 5 signals the shift from single-operation to multi-operation equations. Students who are solid on properties will find order of operations an extension of the same thinking.

WATCH FOR:
- Thumbs-down on SC2 -- these students need support in Lesson 5.

[Maths: Monitor Progress & Feedback | VTLM 2.0: Monitor Progress]`;

const NOTES_RESOURCES = `SAY:
- Printable resources for today.

DO:
- Print Session 4 Worksheet (class set), Session 4 Enabling Scaffold (enabling), Session 4 Extension (extending).

TEACHER NOTES:
Session 4 Worksheet has 6 problems. Session 4 Enabling Scaffold provides structured scaffolding. Session 4 Extension challenges students to design "nice" equations.

WATCH FOR:
- Ensure Session 4 Extension is available for early finishers.

[Maths: Planning -- Preparation | VTLM 2.0: Planning]`;

// ─── Build ───────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // ── Slide 1: Title ──
  titleSlide(pres, UNIT_TITLE, "Session 4: Combining Properties to Find Unknowns", "Factor Decomposition and the Associative Property | Session 4 of 6 | Year 5/6", NOTES_TITLE);

  // ── Slides 2-3: Daily Review (withReveal) — Rounding ──
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Rounding in Addition & Subtraction", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 9, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const problems = [
        { text: "1.  Estimate: 347 + 582 (round to nearest 100)", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true, bold: true, paraSpaceAfter: 4 } },
        { text: "     347 approx ? + 582 approx ?", options: { fontSize: 13, color: C.MUTED, breakLine: true, paraSpaceAfter: 14 } },
        { text: "2.  Estimate: 8,256 - 3,471 (round to nearest 1000)", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true, bold: true, paraSpaceAfter: 4 } },
        { text: "     8,256 approx ? - 3,471 approx ?", options: { fontSize: 13, color: C.MUTED, breakLine: true, paraSpaceAfter: 14 } },
        { text: "3.  Estimate: 23.7 + 46.8 (round to nearest whole)", options: { fontSize: 15, color: C.CHARCOAL, breakLine: true, bold: true, paraSpaceAfter: 4 } },
        { text: "     23.7 approx ? + 46.8 approx ?", options: { fontSize: 13, color: C.MUTED } },
      ];
      s.addText(problems, {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 8.5, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) 300+600=900    2) 8000-3000=5000    3) 24+47=71", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // ── Slide 4: Fluency — Number Talk ──
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "Number Talk", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

    // Big problem display
    addCard(s, 1.5, CONTENT_TOP, 7, 2.0, { strip: C.ACCENT });
    s.addText("24 x 5", {
      x: 1.5, y: CONTENT_TOP + 0.1, w: 7, h: 1.2,
      fontSize: 48, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", valign: "middle", margin: 0,
    });
    s.addText("Solve mentally. When you have an answer AND a strategy, show a quiet thumbs up.", {
      x: 1.8, y: CONTENT_TOP + 1.3, w: 6.4, h: 0.5,
      fontSize: 13, fontFace: FONT_B, color: C.MUTED, align: "center", margin: 0,
    });

    // Strategy prompt
    addCard(s, 1.5, CONTENT_TOP + 2.3, 7, 1.2, { strip: C.SECONDARY });
    s.addText("How many strategies can we find?", {
      x: 1.8, y: CONTENT_TOP + 2.45, w: 6.4, h: 0.9,
      fontSize: 18, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FLUENCY);
  })();

  // ── Slide 5: LI/SC ──
  liSlide(pres,
    ["We are learning to use factors and number properties to find unknown values in multiplication equations"],
    [
      "I can identify factors of a number to rewrite a multiplication",
      "I can use the associative property to find an unknown value in an equation",
      "I can explain my reasoning using mathematical properties",
    ],
    NOTES_LI_SC, FOOTER);

  // ── Slide 6: I Do 1 — 3 x 4 = [] x 2 ──
  workedExSlide(pres, 2, "I Do", "Combining Properties: 3 x 4 = [] x 2",
    [
      "Given: 3 x 4 = [] x 2",
      "",
      "Step 1 -- Decompose 4:",
      "  4 = 2 x 2",
      "",
      "Step 2 -- Rewrite:",
      "  3 x 4 = 3 x (2 x 2)",
      "",
      "Step 3 -- Associative regroup:",
      "  (3 x 2) x 2 = 6 x 2",
      "",
      "So [] = 6",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.8, { strip: C.PRIMARY });
      slide.addText("Property Chain", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      const steps = [
        { text: "3 x 4 = [] x 2", color: C.CHARCOAL },
        { text: "4 = 2 x 2", color: C.ACCENT },
        { text: "(3 x 2) x 2", color: C.SECONDARY },
        { text: "[] = 6", color: C.SUCCESS },
      ];
      steps.forEach((st, i) => {
        addTextOnShape(slide, st.text, {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.5 + i * 0.52, w: lg.rightW - 0.6, h: 0.38, rectRadius: 0.06,
          fill: { color: st.color },
        }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
      // Verify
      slide.addText("Verify: 3 x 4 = 12, 6 x 2 = 12", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 2.6, w: lg.rightW - 0.6, h: 0.25,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, align: "center", margin: 0,
      });
    }
  );

  // ── Slide 7: I Do 2 — 5 x 6 = [] x 3 ──
  workedExSlide(pres, 2, "I Do", "Three Properties Together: 5 x 6 = [] x 3",
    [
      "Given: 5 x 6 = [] x 3",
      "",
      "Step 1 -- Decompose 6 to include 3:",
      "  6 = 3 x 2 = 2 x 3 (commutative)",
      "",
      "Step 2 -- Rewrite:",
      "  5 x 6 = 5 x (2 x 3)",
      "",
      "Step 3 -- Associative:",
      "  (5 x 2) x 3 = 10 x 3",
      "",
      "So [] = 10",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.SECONDARY });
      slide.addText("Three Properties Used", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      const props = [
        { label: "Decomposition", detail: "6 = 3 x 2" },
        { label: "Commutativity", detail: "3 x 2 = 2 x 3" },
        { label: "Associativity", detail: "(5 x 2) x 3 = 10 x 3" },
      ];
      props.forEach((p, i) => {
        slide.addText(p.label, {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.5 + i * 0.55, w: 1.6, h: 0.35,
          fontSize: 11, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
        });
        slide.addText(p.detail, {
          x: lg.rightX + 1.8, y: lg.panelTopPadded + 0.5 + i * 0.55, w: lg.rightW - 2.0, h: 0.35,
          fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
        });
      });
      // Result
      addTextOnShape(slide, "[] = 10    Verify: 30 = 30", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 2.0, w: lg.rightW - 0.6, h: 0.35, rectRadius: 0.06,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // ── Slides 8-9: CFU (withReveal) ──
  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check: Factor Decomposition", "Show Me Boards + Cold Call",
      "4 x 6 = [] x 3\n\nOn your board, write how you would decompose 6\nto match the x 3 on the right side.",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "6 = 3 x 2  ->  4 x (2 x 3)  ->  (4 x 2) x 3 = 8 x 3  ->  [] = 8", {
        x: 0.5, y: 4.0, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU_A);
    }
  );

  // ── Slides 10-11: We Do PP1 (withReveal) — 2 x 8 = [] x 4 ──
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "2 x 8 = [] x 4",
      [
        "2 x 8 = [] x 4",
        "",
        "Step 1: Decompose 8 to include 4:",
        "  8 = ___ x ___",
        "",
        "Step 2: Rewrite:",
        "  2 x (___ x ___)",
        "",
        "Step 3: Regroup:",
        "  (___ x ___) x 4 = ___ x 4",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.8, { strip: C.SECONDARY });
        slide.addText("With your partner:", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
          fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        slide.addText([
          { text: "Find a 4 inside the left side", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Decompose and regroup", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Solve for [] and verify", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
        ], {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.45, w: lg.rightW - 0.5, h: 1.0,
          fontFace: FONT_B, margin: 0, valign: "top",
        });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "8 = 2x4  ->  2x(2x4)  ->  (2x2)x4 = 4x4  ->  [] = 4", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // ── Slides 12-13: We Do PP2 (withReveal) — 6 x 4 = [] x 8 ──
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "6 x 4 = [] x 8", { y: 0.65, fontSize: 22, color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9, 2.4, { strip: STAGE_COLORS["3"] });
      s.addText([
        { text: "6 x 4 = [] x 8", options: { fontSize: 22, color: C.CHARCOAL, breakLine: true, bold: true, paraSpaceAfter: 12 } },
        { text: "60 seconds on your whiteboard!", options: { fontSize: 15, color: C.ALERT, breakLine: true, bold: true, paraSpaceAfter: 8 } },
        { text: "Hint: Try calculate-first if the property path is tricky.", options: { fontSize: 13, color: C.MUTED } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.3, w: 8.5, h: 1.8,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "6 x 4 = 24.  24 / 8 = 3.  [] = 3.  Verify: 3 x 8 = 24", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO2_A);
    }
  );

  // ── Slides 14-15: Hinge CFU (withReveal) — Finger Voting ──
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Find []:  4 x 9 = [] x 6", { color: C.ALERT });

      s.addShape("roundRect", {
        x: 0.5, y: CONTENT_TOP, w: 2.4, h: 0.40, rectRadius: 0.08,
        fill: { color: C.ALERT },
      });
      s.addText("Finger Voting", {
        x: 0.5, y: CONTENT_TOP, w: 2.4, h: 0.40,
        fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      const qY = CONTENT_TOP + 0.56;
      addCard(s, 0.5, qY, 9, 2.8, { strip: C.ALERT, fill: C.WHITE });
      const options = [
        "A)  36",
        "B)  6",
        "C)  4",
        "D)  2",
      ];
      s.addText(options.map((opt, i) => ({
        text: opt,
        options: { fontSize: 18, color: C.CHARCOAL, breakLine: i < options.length - 1, paraSpaceAfter: 14 },
      })), {
        x: 0.9, y: qY + 0.2, w: 8.2, h: 1.8,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      s.addText("Hold up 1, 2, 3, or 4 fingers.", {
        x: 0.9, y: qY + 2.1, w: 8.2, h: 0.4,
        fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_HINGE_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "B)  [] = 6     (4 x 9 = 36, 36 / 6 = 6)", {
        x: 0.5, y: 4.0, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_HINGE_A);
    }
  );

  // ── Slide 16: You Do ──
  workedExSlide(pres, 4, "You Do", "Find the Unknown",
    [
      "First: Decide: properties or calculate-first?",
      "Next: Find [].",
      "Then: Verify and name the property used.",
      "",
      "1.  3 x 8 = [] x 4",
      "2.  5 x 6 = [] x 10",
      "3.  7 x 4 = [] x 2",
      "",
      "Complete all 6 problems on your worksheet.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.8, { strip: C.ALERT });
      slide.addText("Two Approaches:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Properties: Decompose, regroup, solve", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Calculate-first: Find product, then divide", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Choose the best approach for each problem!", options: { fontSize: 12, color: C.MUTED, italic: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.45, w: lg.rightW - 0.5, h: 1.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // ── Slide 17: Exit Ticket ──
  exitTicketSlide(pres,
    [
      "Find [] using properties: 4 x 6 = [] x 3. Show each step and name the properties you used.",
      "Explain in words: Why does 3 x 4 = 6 x 2? Use the words 'associative' or 'decompose' in your answer.",
    ],
    NOTES_EXIT, FOOTER);

  // ── Slide 18: Closing ──
  closingSlide(pres,
    "When would you use properties vs calculate-first to find an unknown?",
    [
      "I can identify factors of a number to rewrite a multiplication",
      "I can use the associative property to find an unknown value in an equation",
      "I can explain my reasoning using mathematical properties",
    ],
    NOTES_CLOSING);

  // ── Slide 19: Resources ──
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // ── Write PPTX ──
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "ALG6_Session4_Combining_Properties.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Generate PDFs ─────────────────────────────────────────────────────────

  // Session 4 Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Combining Properties to Find Unknowns",
      color: C.NAVY,
      lessonInfo: "Session 4 of 6 | Algebra: Unknown Values | Year 5/6",
    });
    y = addTipBox(doc, "For each equation: decide whether to use properties (decompose, regroup) or calculate-first (find product, divide). Show ALL working and name the properties you used.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A: Clean Decomposition Path", y, { color: C.NAVY });
    y = addProblem(doc, 1, "3 x 8 = [] x 4     (Hint: decompose 8 to include 4)", y, { color: C.NAVY, writeLines: [{ label: "Decomposition:" }, { label: "Working:" }, { label: "[] =" }] });
    y = addProblem(doc, 2, "5 x 6 = [] x 10    (Hint: decompose one factor)", y, { color: C.NAVY, writeLines: [{ label: "Decomposition:" }, { label: "Working:" }, { label: "[] =" }] });

    y = addSectionHeading(doc, "Section B: Calculate-First May Be Easier", y, { color: C.NAVY });
    y = addProblem(doc, 3, "7 x 4 = [] x 2", y, { color: C.NAVY, writeLines: [{ label: "Method:" }, { label: "Working:" }, { label: "[] =" }] });
    y = addProblem(doc, 4, "9 x 6 = [] x 4", y, { color: C.NAVY, writeLines: [{ label: "Method:" }, { label: "Working:" }, { label: "[] =" }] });

    y = addSectionHeading(doc, "Section C: Explain Your Reasoning", y, { color: C.NAVY });
    y = addProblem(doc, 5, "4 x 9 = [] x 6. Find [] and explain which approach you chose and why.", y, { color: C.NAVY, writeLines: [{ label: "[] =" }, { label: "Explanation:" }, { label: "" }] });
    y = addProblem(doc, 6, "8 x 3 = [] x 12. Find [] and name each property used.", y, { color: C.NAVY, writeLines: [{ label: "[] =" }, { label: "Properties used:" }, { label: "" }] });

    addPdfFooter(doc, "Session 4 | Algebra: Finding Unknown Values | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Session 4 Enabling Scaffold
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Step-by-Step Framework for Combining Properties",
      color: C.TEAL,
      lessonInfo: "Session 4 of 6 | Algebra: Unknown Values | Year 5/6",
    });
    y = addTipBox(doc, "Follow the steps for each problem. Fill in each blank. The first one is done as an example.", y, { color: C.TEAL });

    // Example
    y = addSectionHeading(doc, "Example (done for you)", y, { color: C.NAVY });
    y = addBodyText(doc, "3 x 4 = [] x 2", y, { fontSize: 12 });
    y = addBodyText(doc, "Step 1: What is on the right? x 2. So I need a 2 in my decomposition.", y);
    y = addBodyText(doc, "Step 2: Decompose: 4 = 2 x 2", y);
    y = addBodyText(doc, "Step 3: Rewrite: 3 x (2 x 2)", y);
    y = addBodyText(doc, "Step 4: Regroup: (3 x 2) x 2 = 6 x 2", y);
    y = addBodyText(doc, "Step 5: [] = 6. Check: 3 x 4 = 12, 6 x 2 = 12. Yes!", y);
    y += 8;

    y = addSectionHeading(doc, "Now you try:", y, { color: C.NAVY });

    const problems = [
      "3 x 8 = [] x 4",
      "5 x 6 = [] x 10",
      "7 x 4 = [] x 2",
      "9 x 6 = [] x 4",
    ];
    problems.forEach((prob, i) => {
      y = addBodyText(doc, (i + 1) + ". " + prob, y, { fontSize: 12 });
      y = addWriteLine(doc, "Right side has: x ___. Need a ___ in decomposition.", y);
      y = addWriteLine(doc, "Decompose: ___ = ___ x ___", y);
      y = addWriteLine(doc, "Rewrite: ___ x (___ x ___)", y);
      y = addWriteLine(doc, "Regroup: (___ x ___) x ___ = ___ x ___", y);
      y = addWriteLine(doc, "[] = ___. Check: ", y);
      y += 8;
      if (y > 700) { doc.addPage(); y = 50; }
    });

    addPdfFooter(doc, "Session 4 | Enabling Scaffold | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Session 4 Extension — Design Your Own Equations
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Create Equations Where Properties Work Cleanly",
      color: C.NAVY,
      lessonInfo: "Session 4 of 6 | Algebra: Unknown Values | Year 5/6",
    });

    y = addSectionHeading(doc, "What Makes a 'Nice' Equation?", y, { color: C.NAVY });
    y = addBodyText(doc, "Some equations like 3 x 4 = [] x 2 have clean decomposition paths -- you can decompose 4 into 2 x 2 and regroup easily.", y);
    y = addBodyText(doc, "Other equations like 6 x 4 = [] x 8 are trickier -- the property path is messy, so calculate-first works better.", y);
    y = addBodyText(doc, "A 'nice' equation is one where the target factor (the number after x on the right side) is a FACTOR of one of the numbers on the left side.", y);
    y += 5;

    y = addSectionHeading(doc, "Worked Examples", y, { color: C.NAVY });
    y = addBodyText(doc, "Nice: 5 x 8 = [] x 4. Why? Because 4 is a factor of 8 (8 = 4 x 2).", y);
    y = addBodyText(doc, "Not as nice: 5 x 7 = [] x 3. Why? Because 3 is NOT a factor of 5 or 7.", y);
    y += 5;

    y = addSectionHeading(doc, "Your Challenge", y, { color: C.NAVY });
    y = addTipBox(doc, "Create 3 'nice' equations of the form a x b = [] x c. Solve each one using properties AND verify.", y, { color: C.TEAL });

    y = addProblem(doc, 1, "Create your equation: ___ x ___ = [] x ___", y, { color: C.NAVY, writeLines: [{ label: "Decomposition:" }, { label: "Solution: [] =" }, { label: "Verify:" }] });
    y = addProblem(doc, 2, "Create your equation: ___ x ___ = [] x ___", y, { color: C.NAVY, writeLines: [{ label: "Decomposition:" }, { label: "Solution: [] =" }, { label: "Verify:" }] });
    y = addProblem(doc, 3, "Create your equation: ___ x ___ = [] x ___", y, { color: C.NAVY, writeLines: [{ label: "Decomposition:" }, { label: "Solution: [] =" }, { label: "Verify:" }] });

    y = addSectionHeading(doc, "Bonus: The Pattern", y, { color: C.NAVY });
    y = addBodyText(doc, "Look at your three equations. What do the 'nice' ones have in common? Write a rule for creating equations where properties work cleanly:", y);
    y = addWriteLine(doc, "My rule: ", y);
    y = addWriteLine(doc, "", y);

    addPdfFooter(doc, "Session 4 | Extension Investigation | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("\nSession 4 build complete!");
}

build().catch(err => { console.error(err); process.exit(1); });
