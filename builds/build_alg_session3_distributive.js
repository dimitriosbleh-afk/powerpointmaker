"use strict";

// Algebra Unit — Session 3: The Distributive Property
// Session 3 of 6, Grade 5/6 Numeracy, Variant 0
// DR: LCM and HCF
// Fluency: 90-second multiplication sprint
// VC2M5A02 elaboration 4: 4 x 13 = 4 x 10 + 4 x 3

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

const SESSION = 3;
const UNIT_TITLE = "Algebra: Finding Unknown Values";
const FOOTER = "Session 3 of 6 | Algebra: Unknown Values | Year 5/6 Maths";
const OUT_DIR = "output/ALG6_Session3_Distributive_Property";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "6 problems for independent practice using the distributive property.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Array-based scaffolding for students needing support.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Extending: does the distributive property work with subtraction?");
const RESOURCE_ITEMS = [WORKSHEET_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ───────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Sessions 1 and 2 gave us two tools -- equivalence and the commutative/associative properties. Today we add a third: the distributive property.
- This property lets you BREAK APART a hard multiplication into two easy ones. By the end, you will use it to find unknown values.

DO:
- Display the title slide. Whiteboards ready.

TEACHER NOTES:
Lesson 3 introduces the distributive property (VC2M5A02 elaboration 4: 4 x 13 = 4 x 10 + 4 x 3). This connects to students' mental computation strategies -- many already "break apart" numbers when multiplying without knowing the formal name. The lesson makes this implicit strategy explicit and shows how it creates equivalent expressions useful for finding unknowns. The array representation makes the property concrete and visual.

WATCH FOR:
- Students who used the associative property confidently in L2 -- they are ready.
- Readiness signal: settled with materials.

[Maths: Planning | VTLM 2.0: Planning]`;

const NOTES_DR_Q = `SAY:
- Today's review: Lowest Common Multiples and Highest Common Factors.
- First: What is the LCM of 6 and 9? Write it on your board.

DO:
- Display the slide. Run each problem with 15-second think time.
- Students write on whiteboards. "Show me!"
- Model the listing strategy for any errors.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Write the LCM of 6 and 9 on your board. Show me!"
- Scan for: 18 on 80%+ of boards.
PROCEED: If 80%+ correct on LCM and HCF, move to Fluency.
PIVOT: If errors on LCM, students may be finding the product (6 x 9 = 54) instead of the LOWEST common multiple. Reteach: "List the multiples of each number until you find a match. 6: 6, 12, 18... 9: 9, 18... They both hit 18 -- that is the LCM. It is NOT 6 x 9." Re-check: "LCM of 3 and 4?"

TEACHER NOTES:
Daily Review targets: "Number Properties and Algorithms -- I can identify lowest common multiples and highest common factors of pairs or triples of natural numbers (e.g. the LCM of {6, 9} is 18 and the HCF is 3)." LCM and HCF connect to today's lesson because the distributive property involves decomposing numbers into factors -- students who understand factor relationships will access distributive reasoning more easily.

WATCH FOR:
- Students who confuse LCM and HCF -- LCM is usually bigger, HCF is usually smaller.
- Students who multiply the numbers instead of listing multiples -- direct misconception.
- Readiness signal: correct answers with listing strategy shown.

[Maths: Daily Review | VTLM 2.0: Retention & Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers -- tick if you got it right, fix it if you didn't.
- Fixing it is the learning.
- After boards: LCM of 6 and 9 is 18. The multiples of 6 are 6, 12, 18, 24... The multiples of 9 are 9, 18, 27... The smallest they share is 18.
- HCF is 3. The factors of 6 are 1, 2, 3, 6. The factors of 9 are 1, 3, 9. The biggest they share is 3.
- One more: LCM of 4 and 6? [12]

DO:
- Click to reveal answers. Students tick correct responses and fix errors.
- Scan for common mistakes.

TEACHER NOTES:
Tick-and-fix gives immediate feedback on factor/multiple reasoning. Note students who struggle -- factor thinking underpins the distributive reasoning in today's lesson.

WATCH FOR:
- Students who got all three correct quickly -- factor reasoning is strong.
- Students who confused LCM and HCF -- they may need a mnemonic.

[Maths: Daily Review Answers | VTLM 2.0: Retention & Recall]`;

const NOTES_FLUENCY_Q = `SAY:
- Fluency: Repeated Practice. You have 90 seconds to answer as many as you can.
- These are ALL multiplications you know -- the challenge is speed. Go!

DO:
- Display the 15-problem grid (mixed x facts). Students work silently.
- Time 90 seconds exactly. "GO" and "STOP."

TEACHER NOTES:
Repeated practice builds automaticity with multiplication facts. Today's lesson requires students to decompose multiplications (e.g., 4 x 13 = 4 x 10 + 4 x 3), which demands fluent recall of partial products. If students cannot quickly compute 4 x 10 and 4 x 3, the distributive property lesson will stall at the arithmetic stage rather than reaching the algebraic insight.

WATCH FOR:
- Students who complete fewer than 8 in 90 seconds -- they need additional fluency support outside the lesson.
- Readiness signal: 12+ correct in 90 seconds.

[Maths: Fluency | VTLM 2.0: Retention & Recall]`;

const NOTES_FLUENCY_A = `SAY:
- Pens down. Let's check. How many did you get? Hands up for 12 or more.
- Read answers quickly for self-mark.

DO:
- Click to reveal answers. Students self-mark.
- Note students who completed fewer than 8 for future fluency planning.

TEACHER NOTES:
Self-marking builds metacognition. Students who can self-assess their fluency gaps will target their own practice.

WATCH FOR:
- Students who scored 12+ -- multiplication recall is automated.

[Maths: Fluency Answers | VTLM 2.0: Retention & Recall]`;

const NOTES_LI_SC = `SAY:
- Read from slide: "We are learning to use the distributive property to break apart multiplication problems and find unknown values."
- SC1: You can split a multiplication into two parts using an array. SC2: You can write the equation. SC3: You can use it to find an unknown.
- SC2 is our main target today.

DO:
- Display the slide. Point to each SC.

TEACHER NOTES:
The LI addresses VC2M5A02 elaboration 4 directly. The SC progress from concrete (array) to symbolic (equation) to algebraic (unknown). This CRA-like progression within a single lesson ensures accessibility -- students who only achieve SC1 still have a concrete understanding of the property, while SC3 extends to algebraic application.

WATCH FOR:
- Students who seem unsure about "distributive" -- the vocabulary slide will clarify.
- Readiness signal: students reading ahead.

[Maths: Planning -- Curriculum Alignment | VTLM 2.0: Planning]`;

const NOTES_VOCAB = `SAY:
- One new term today: distributive.
- Point to DISTRIBUTIVE: "The distributive property says you can 'distribute' a multiplication over addition. It means you can break a hard multiplication into two easier ones and add the results."
- Example: 4 x 13 is tricky. But 4 x 10 = 40, and 4 x 3 = 12. Add them: 40 + 12 = 52. So 4 x 13 = 52.
- You have probably done this in your head before -- this property gives it a name.
- Choral response: say it with me -- "distributive."

DO:
- Display the slide with the term and visual example.
- Point to the array diagram showing 4 x 13 split into 4 x 10 and 4 x 3.
- Run the choral response for pronunciation.

TEACHER NOTES:
Many students already use this strategy mentally (breaking 13 into 10 + 3) without knowing its name. This lesson elevates an intuitive strategy to explicit mathematical language. The connection to mental computation makes it immediately relevant -- students have been "distributing" without knowing it.

MISCONCEPTIONS:
- Misconception: "You can only break apart the second number, not the first."
  Why: All initial examples show a x (b + c), so students assume the first factor stays fixed.
  Impact: Students miss that (a + b) x c also works, limiting their flexibility.
  Quick correction: "You can break apart EITHER factor. 4 x 13 = 4 x (10 + 3) works. But you could also do (2 + 2) x 13 = 2 x 13 + 2 x 13. Both ways work!"

WATCH FOR:
- Students who say "I already do this" -- validate and name it: "You are using the distributive property."
- Readiness signal: students nodding at the mental computation connection.

[Maths: Launch -- Explicit Instruction | VTLM 2.0: Explicit Explanation]`;

const NOTES_IDO1 = `SAY:
- Watch me use an array to prove the distributive property.
- I want to calculate 4 x 13. That is a big array -- 4 rows of 13. But I can split the 13 into 10 and 3.
- Now I have TWO arrays side by side: 4 x 10 on the left, and 4 x 3 on the right. Together they make the original 4 x 13.
- Why did I split 13 into 10 + 3 and not 7 + 6? Because 10 and 3 are easy to multiply by 4. I chose a split that makes the maths simple.
- 4 x 10 = 40. 4 x 3 = 12. Add them: 40 + 12 = 52. So 4 x 13 = 52.
- In symbols: 4 x 13 = 4 x (10 + 3) = 4 x 10 + 4 x 3 = 40 + 12 = 52.
- Let me verify: 4 x 13 = 52. Does 4 x 10 + 4 x 3 = 52? 40 + 12 = 52. Yes.

DO:
- Display the slide showing the full 4 x 13 array, then the split array.
- Use a vertical line to show where the array splits (after column 10).
- Write the symbolic form underneath.

TEACHER NOTES:
This I Do uses the exact example from VC2M5A02 elaboration 4: 4 x 13 = 4 x 10 + 4 x 3. The array representation makes the property visual and concrete -- students can SEE that the two smaller arrays together form the original. The decision point about WHY we chose 10 + 3 (rather than other decompositions) models strategic thinking.

MISCONCEPTIONS:
- Misconception: "4 x 13 = 4 x 10 x 3 = 120."
  Why: Students confuse the distributive property (distribute over addition) with the associative property (regroup factors). They write x instead of +.
  Impact: This produces wildly wrong answers and conflates two different properties.
  Quick correction: "The distributive property uses ADDITION, not multiplication. You BREAK APART 13 into 10 + 3, then multiply EACH by 4 and ADD the results. 4 x 10 + 4 x 3 = 52, not 4 x 10 x 3 = 120."

WATCH FOR:
- Students writing x instead of + between the partial products -- the key error to prevent.
- Students who understand the array but struggle with the symbolic form -- they may need more concrete-to-abstract bridging.
- Readiness signal: students mouthing "40 + 12 = 52" as you work.

[Maths: Launch -- Explicit Instruction (I Do) | VTLM 2.0: Explicit Explanation & Modelling]`;

const NOTES_IDO2 = `SAY:
- Now let me show you how the distributive property helps find unknowns.
- Here is the equation: 5 x 14 = 5 x 10 + 5 x []. What is []?
- I can see that 14 has been broken into 10 + []. The distributive property says 5 x 14 = 5 x (10 + []) = 5 x 10 + 5 x [].
- If 14 = 10 + [], then [] = 14 - 10 = 4.
- Let me verify: 5 x 10 + 5 x 4 = 50 + 20 = 70. And 5 x 14 = 70. Yes.
- Notice -- finding the unknown here was just working out what 14 was broken into. The distributive property TELLS us the structure.

DO:
- Display the slide with the equation and step-by-step working.
- Emphasise the connection: 14 = 10 + [] -> [] = 4.
- Show the verification on the right side.

TEACHER NOTES:
This second I Do transitions from demonstrating the distributive property to USING it to find unknowns -- the core algebraic skill. The equation 5 x 14 = 5 x 10 + 5 x [] reveals the structure: the common factor (5) is distributed, and the unknown ([]) is the missing addend.

WATCH FOR:
- Students who try to solve by calculating 5 x 14 = 70, then 70 / 5 = 14 -- valid but redirect to the property-based method.
- Students who see "[] = 4" and understand immediately -- they are ready for We Do.
- Readiness signal: students nodding at the "14 = 10 + []" insight.

[Maths: Launch -- Explicit Instruction (I Do) | VTLM 2.0: Explicit Explanation & Modelling]`;

const NOTES_CFU_Q = `SAY:
- Quick check. On your whiteboard, complete this equation using the distributive property.
- 3 x 15 = 3 x ___ + 3 x ___. Fill in the two blanks. You have 15 seconds.

DO:
- Display the question. Students work on whiteboards.
- "Show me!" Scan for correct splits.

CFU CHECKPOINT:
Technique: Show Me Boards (open response)
Script:
- "Fill in the blanks: 3 x 15 = 3 x ___ + 3 x ___. Show me your boards!"
- Scan for: any valid decomposition where the two blanks add to 15.
PROCEED: If 80%+ produce a valid split, students understand the property. Move to We Do.
PIVOT: If students write 3 x 15 = 3 x 10 x 5 (using x instead of +), the key error from the MISCONCEPTIONS section has surfaced. Reteach: "The distributive property uses ADDITION. You break 15 into two parts that ADD to 15: 10 + 5 = 15. Then multiply EACH part by 3 and ADD: 3 x 10 + 3 x 5." Re-check: "4 x 12 = 4 x ___ + 4 x ___."

TEACHER NOTES:
This CFU tests SC1 and SC2 -- can students decompose and write the distributive form? The open-response format is deliberate: there are multiple correct answers (10+5, 12+3, 8+7, etc.), and accepting all of them reinforces that the property works for ANY decomposition of 15.

WATCH FOR:
- Students writing x between partial products -- catch and correct immediately.
- Students who only use the 10+remainder split -- valid but limited.
- Readiness signal: correct splits with + between partial products.

[Maths: Monitor Progress | VTLM 2.0: Monitor Progress (CFU)]`;

const NOTES_CFU_A = `SAY:
- The most common split is 3 x 10 + 3 x 5. Let's check: 30 + 15 = 45. And 3 x 15 = 45. Yes.
- Did anyone split it differently? 3 x 12 + 3 x 3? That works too! 36 + 9 = 45.

DO:
- Click to reveal. Celebrate multiple valid answers.

TEACHER NOTES:
Accepting multiple valid decompositions builds mathematical flexibility and reinforces that 15 can be split many ways.

WATCH FOR:
- Students who see only one valid answer -- encourage exploration of alternatives.

[Maths: Monitor Progress | VTLM 2.0: Monitor Progress (CFU)]`;

const NOTES_WEDO1_Q = `SAY:
- Together now. The equation is: 6 x 17 = 6 x [] + 6 x 7. Find [].
- Ask: What has 17 been broken into? [[] + 7]
- Ask: If 17 = [] + 7, what is []? [10]
- Let's verify: 6 x 10 + 6 x 7 = 60 + 42 = 102. And 6 x 17 = 102. Yes.
- The distributive property told us the structure: 17 = [] + 7. Simple subtraction gives the unknown.

DO:
- Display the equation with prompts.
- Cold Call for each step. Click to reveal the solution.

CFU CHECKPOINT:
Technique: Cold Call
Script:
- "[Name], what has 17 been split into?" "[Name], so what is []?"
- Scan for: correct identification of 17 = [] + 7 -> [] = 10.
PROCEED: If correct, move to PP2.
PIVOT: If students struggle to see 17 = [] + 7, reteach: "Look at the equation. On the left: 6 x 17. On the right: 6 x [] + 6 x 7. The 6 is common. The 17 has been split into [] and 7. So [] + 7 = 17." Re-check with: "8 x 15 = 8 x [] + 8 x 5. What is []?"

TEACHER NOTES:
Problem Pair 1 applies the distributive property to find an unknown. The structure is transparent: the common factor (6) appears in both partial products, and the unknown ([]) is the missing addend. Students who see the structure solve it instantly (17 = [] + 7 -> [] = 10).

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide students with a pre-drawn array for 6 x 17, split at column 10. Students count the columns in each section to identify [] = 10. The visual scaffold makes the decomposition concrete.
EXTENDING PROMPT:
- Task: "Find [] in: 4 x [] = 4 x 20 + 4 x 8. Then find [] in: [] x 15 = 3 x 15 + 7 x 15. What is different about the second one?"
- Extra Notes: The second problem distributes the OTHER factor -- extending to (a+b) x c form.

WATCH FOR:
- Students who try to calculate 6 x 17 first then work backwards -- valid but not using the property directly.
- Students who write [] = 10 confidently -- they see the structure.
- Readiness signal: immediate identification of 17 = [] + 7.

[Maths: Explore -- Guided Practice (We Do) | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO1_A = `SAY:
- 17 = [] + 7 so [] = 10. Check: 60 + 42 = 102 = 6 x 17. Yes.

DO:
- Click to reveal the solution. Celebrate correct answers.

TEACHER NOTES:
Reveal confirms the answer and verification. Move to PP2.

WATCH FOR:
- Students who got it right -- ready for the next problem.

[Maths: Explore -- Guided Practice (We Do) | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO2_Q = `SAY:
- Your turn on boards. Find [] in: 8 x 23 = 8 x 20 + 8 x []. You have 30 seconds.
- Now a harder one on boards: 7 x [] = 7 x 30 + 7 x 6. What is []? 30 seconds.

DO:
- Display the first equation. 30 seconds, then boards up. Reveal.
- Display the second equation. 30 seconds, then boards up. Reveal.

CFU CHECKPOINT:
Technique: Show Me Boards (open response)
Script:
- "Write [] on your board. Show me!" (for each problem)
- Scan for: [] = 3 on first, [] = 36 on second.
PROCEED: If 80%+ correct on both, move to hinge question.
PIVOT: If the second problem ([] = 36) causes errors, students may be adding 30 + 6 incorrectly or not recognising that [] = 30 + 6. Reteach: "The right side shows 7 x 30 + 7 x 6. The 7 is distributed across (30 + 6). So the left side is 7 x (30 + 6) = 7 x 36." Re-check: "5 x [] = 5 x 40 + 5 x 2."

TEACHER NOTES:
Two problems with increasing complexity. The first ([] = 3) is straightforward -- the unknown is the smaller addend. The second ([] = 36) requires combining the two addends -- students must recognise that [] IS the sum, not one of the parts.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Students work with single-digit decompositions only: "4 x 15 = 4 x 10 + 4 x []." Smaller numbers, same structure.
EXTENDING PROMPT:
- Task: "Use the distributive property to calculate 7 x 99 without a calculator. Hint: 99 = 100 - 1. Does the distributive property work with subtraction too?"
- Extra Notes: The Session 3 Extension PDF explores distribution over subtraction.

WATCH FOR:
- Students who get [] = 3 but not [] = 36 -- the second problem requires addition (30+6) not subtraction (23-20).
- Students who solve both quickly -- ready for You Do.
- Readiness signal: both answers correct with reasoning articulated.

[Maths: Explore -- Guided Practice (We Do) | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO2_A = `SAY:
- A) [] = 3. Because 23 = 20 + [], so [] = 3.
- Verify: 8 x 20 + 8 x 3 = 160 + 24 = 184. And 8 x 23 = 184. Yes.
- B) [] = 36. Because the distributive property says 7 x 36 = 7 x 30 + 7 x 6.

DO:
- Click to reveal answers. Discuss.

TEACHER NOTES:
The second answer ([] = 36) is the key insight -- when the unknown IS the whole number being distributed, students add the parts.

WATCH FOR:
- Students who solved both correctly -- ready for the hinge.

[Maths: Explore -- Guided Practice (We Do) | VTLM 2.0: Scaffold Practice]`;

const NOTES_HINGE_Q = `SAY:
- Gate check. Which equation correctly uses the distributive property? Show fingers: 1, 2, 3, or 4.

DO:
- Display options. 15 seconds. "Show fingers." Scan. Reveal.

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- "Which equation shows the distributive property correctly? 1, 2, 3, or 4. Show me NOW."
- Scan for: option B on 80%+.
PROCEED: If 80%+ choose B, students can identify the distributive form. Move to You Do.
PIVOT: If students choose A (5 x 10 x 8), the x vs + confusion is still present. Reteach: "Distributive means break apart with ADDITION, then multiply each part. 5 x 18 = 5 x (10 + 8) = 5 x 10 + 5 x 8. The PLUS is the key." Re-check: "Is 3 x 21 = 3 x 20 + 3 x 1 distributive? Thumbs."

TEACHER NOTES:
Each distractor maps to a specific error: A (using x instead of +, confusing distributive with associative), C (adding factors instead of distributing), D (incorrect partial products). Finger voting ensures whole-class participation and rapid interpretation. This is the gate between We Do and You Do.

WATCH FOR:
- Students choosing A -- the x vs + error is the #1 misconception in this lesson.
- Students choosing quickly and correctly -- ready for You Do.

[Maths: Monitor Progress | VTLM 2.0: Monitor Progress (CFU)]`;

const NOTES_HINGE_A = `SAY:
- The answer is B: 6 x 14 = 6 x 10 + 6 x 4. Check: 10 + 4 = 14, and the 6 is distributed across both.
- Why not A? 5 x 18 = 5 x 10 x 8 uses MULTIPLICATION, not addition. That is the associative property, not distributive.
- Why not C? 3 x 12 = 3 + 12 does not distribute at all.
- Why not D? 4 x 15 = 40 + 45 = 85, but 4 x 15 = 60. Wrong numbers.

DO:
- Click to reveal answer and explanations.

TEACHER NOTES:
Explaining each distractor publicly names the misconceptions so all students can self-check.

WATCH FOR:
- Students who chose A and now understand the error -- progress.

[Maths: Monitor Progress | VTLM 2.0: Monitor Progress (CFU)]`;

const NOTES_YOUDO = `SAY:
- Read from slide: "First: Identify the common factor. Next: Work out what the number was broken into. Then: Find the unknown and verify."
- Complete all 6 problems on your worksheet.

DO:
- Distribute Session 3 Worksheet. Display the slide. Circulate.
- Check: are students identifying the structure (a x n = a x [] + a x [])?
- After 12 minutes: "Two-minute warning."

TEACHER NOTES:
The worksheet has 6 graduated problems: 1-2 mirror We Do exactly (find missing addend), 3-4 find the whole number ([] = sum of addends), 5-6 reverse the direction (given the partial products, identify the original multiplication). This progression builds from familiar to novel within the independent phase.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Session 3 Enabling Scaffold with arrays drawn for each problem and the first step completed. Students identify the missing section of the array.
EXTENDING PROMPT:
- Task: Session 3 Extension investigation on distribution over subtraction: "Does 7 x 99 = 7 x 100 - 7 x 1? Test it. Does this always work?"

WATCH FOR:
- Students who get stuck on problems 3-4 (finding the whole) -- they may need the explicit prompt: "Add the two parts together to find []."
- Readiness signal: completing problem 4+ correctly.

[Maths: Summarise -- Independent Practice (You Do) | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Two questions. Show your working. 3 minutes.

DO:
- Display the exit ticket. Students work silently. Collect after 3 minutes.

TEACHER NOTES:
Q1 tests SC2 (write the distributive form): students must decompose a multiplication. Q2 tests SC3 (find unknown): students must use the distributive structure to identify []. Data informs Lesson 4 -- students who struggle here need reteaching before combining multiple properties.

WATCH FOR:
- Students using x instead of + -- the persistent misconception. Note for tomorrow's DR.
- Readiness signal: both questions completed with correct symbolic form.

[Maths: Summarise -- Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Success criteria check.
- SC1: I can break a multiplication into two parts using an array. Thumbs.
- SC2: I can write an equation showing the distributive property. Thumbs.
- SC3: I can use the distributive property to find an unknown value. Thumbs.
- Turn and talk: How does the distributive property help you multiply big numbers in your head?
- Tomorrow we combine ALL the properties -- commutative, associative, AND distributive -- to solve harder equations.

DO:
- Display the closing slide. Run thumbs for each SC.

TEACHER NOTES:
The forward look to Lesson 4 (combining properties) signals increasing complexity. Students who are solid on SC1-SC2 but shaky on SC3 will benefit from seeing how properties combine.

WATCH FOR:
- Thumbs-down on SC3 -- these students need the enabling approach in Lesson 4.

[Maths: Monitor Progress & Feedback | VTLM 2.0: Monitor Progress]`;

const NOTES_RESOURCES = `SAY:
- Printable resources for this lesson.

DO:
- Print Session 3 Worksheet (class set), Session 3 Enabling Scaffold (enabling), Session 3 Extension (extending) before the lesson.

TEACHER NOTES:
Session 3 Worksheet is the main worksheet. Session 3 Enabling Scaffold provides array-based scaffolding. Session 3 Extension investigates distribution over subtraction -- a self-contained investigation for extending students.

WATCH FOR:
- Ensure Session 3 Extension is available for early finishers.

[Maths: Planning -- Preparation | VTLM 2.0: Planning]`;

// ─── Build ───────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // ── Slide 1: Title ──
  titleSlide(pres, UNIT_TITLE, "Session 3: The Distributive Property", "Breaking Apart Multiplications to Find Unknowns | Session 3 of 6 | Year 5/6", NOTES_TITLE);

  // ── Slides 2-3: Daily Review (withReveal) — LCM and HCF ──
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "LCM and HCF", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const problems = [
        "1.  Find the LCM of 6 and 9",
        "     List multiples of each until",
        "     you find a match",
        "",
        "2.  Find the HCF of 6 and 9",
        "     List factors of each and find",
        "     the biggest common one",
        "",
        "3.  Find the LCM of 4 and 6",
        "     Multiples of 4: 4, 8, 12...",
        "     Multiples of 6: 6, 12...",
      ];
      s.addText(problems.map((p, i) => ({
        text: p,
        options: { fontSize: 13.5, color: C.CHARCOAL, breakLine: i < problems.length - 1, paraSpaceAfter: 2 },
      })), {
        x: 0.75, y: CONTENT_TOP + 0.12, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.25,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, 2.2, { strip: C.SECONDARY });
      s.addText([
        { text: "Common = the same", options: { fontSize: 15, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "LCM: Lowest Common Multiple", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true, paraSpaceAfter: 6 } },
        { text: "HCF: Highest Common Factor", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true, paraSpaceAfter: 6 } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "On your whiteboards!", options: { fontSize: 14, bold: true, color: C.ALERT } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.12, w: 3.8, h: 2.0,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "LCM(6,9) = 18     HCF(6,9) = 3     LCM(4,6) = 12", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // ── Slides 4-5: Fluency (withReveal) — 90-Second Sprint ──
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Fluency");
      addTitle(s, "Repeated Practice: 90-Second Sprint", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const facts = [
        "1.   6 x 7 =",    "2.   9 x 4 =",
        "3.   8 x 5 =",    "4.   7 x 3 =",
        "5.   12 x 4 =",   "6.   5 x 9 =",
        "7.   8 x 8 =",    "8.   11 x 6 =",
        "9.   4 x 7 =",    "10.  6 x 12 =",
        "11.  3 x 9 =",    "12.  7 x 7 =",
        "13.  8 x 6 =",    "14.  9 x 11 =",
        "15.  7 x 8 =",
      ];
      s.addText(facts.map((p, i) => ({
        text: p,
        options: { fontSize: 13, color: C.CHARCOAL, breakLine: i < facts.length - 1, paraSpaceAfter: 2 },
      })), {
        x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, 1.6, { strip: C.ALERT });
      s.addText([
        { text: "90 seconds -- GO!", options: { fontSize: 20, bold: true, color: C.ALERT, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Write answers on your whiteboard.", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Speed AND accuracy!", options: { fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: 1.4,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_FLUENCY_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "42, 36, 40, 21, 48, 45, 64, 66, 28, 72, 27, 49, 48, 99, 56", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // ── Slide 6: LI/SC ──
  liSlide(pres,
    ["We are learning to use the distributive property to break apart multiplication problems and find unknown values"],
    [
      "I can break a multiplication into two simpler parts using an array",
      "I can write an equation showing the distributive property",
      "I can use the distributive property to find an unknown value",
    ],
    NOTES_LI_SC, FOOTER);

  // ── Slide 7: Key Vocabulary — The Distributive Property ──
  contentSlide(pres, "Key Vocabulary", C.PRIMARY, "The Distributive Property",
    [
      "Distributive property means we can break a",
      "difficult multiplication into two easy ones",
      "and add the results.",
      "",
      "4 x 13 can be broken down like this:",
      "",
      "Step 1:  4 x 10 + 4 x 3",
      "Step 2:  40 + 12 = 52",
      "",
      "a x (b + c) = a x b + a x c",
    ],
    NOTES_VOCAB, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.8, { strip: C.SECONDARY });
      slide.addText("Visual Example", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      // 4 x 10 box
      addTextOnShape(slide, "4 x 10 = 40", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.55, w: 2.2, h: 0.8, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      // + sign
      slide.addText("+", {
        x: lg.rightX + 2.5, y: lg.panelTopPadded + 0.65, w: 0.4, h: 0.6,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
      });
      // 4 x 3 box
      addTextOnShape(slide, "4 x 3 = 12", {
        x: lg.rightX + 2.9, y: lg.panelTopPadded + 0.55, w: 1.1, h: 0.8, rectRadius: 0.06,
        fill: { color: C.ACCENT },
      }, { fontSize: 12, fontFace: FONT_H, color: C.WHITE, bold: true });
      // Result
      addTextOnShape(slide, "40 + 12 = 52", {
        x: lg.rightX + 0.4, y: lg.panelTopPadded + 1.65, w: lg.rightW - 0.8, h: 0.45, rectRadius: 0.06,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      // Label
      slide.addText("4 x 13 = 52", {
        x: lg.rightX + 0.4, y: lg.panelTopPadded + 2.25, w: lg.rightW - 0.8, h: 0.3,
        fontSize: 14, fontFace: FONT_B, color: C.MUTED, align: "center", margin: 0,
      });
    }
  );

  // ── Slide 8: I Do 1 — Distributive Property with Arrays ──
  workedExSlide(pres, 2, "I Do", "Distributive Property with Arrays",
    [
      "Calculate: 4 x 13",
      "",
      "Split 13 into 10 + 3",
      "",
      "4 x 13 = 4 x (10 + 3)",
      "       = 4 x 10 + 4 x 3",
      "       = 40 + 12",
      "       = 52",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      // Array visualisation on right side
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("Array: 4 x 13", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      // Left array block: 4 x 10
      slide.addShape("rect", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.5, w: 2.3, h: 1.2,
        fill: { color: C.PRIMARY }, line: { color: C.WHITE, width: 1 },
      });
      slide.addText("4 x 10", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.5, w: 2.3, h: 0.5,
        fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true, align: "center", valign: "middle", margin: 0,
      });
      slide.addText("= 40", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.0, w: 2.3, h: 0.4,
        fontSize: 14, fontFace: FONT_B, color: C.WHITE, align: "center", valign: "middle", margin: 0,
      });
      // Right array block: 4 x 3
      slide.addShape("rect", {
        x: lg.rightX + 2.6, y: lg.panelTopPadded + 0.5, w: 1.3, h: 1.2,
        fill: { color: C.ACCENT }, line: { color: C.WHITE, width: 1 },
      });
      slide.addText("4 x 3", {
        x: lg.rightX + 2.6, y: lg.panelTopPadded + 0.5, w: 1.3, h: 0.5,
        fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true, align: "center", valign: "middle", margin: 0,
      });
      slide.addText("= 12", {
        x: lg.rightX + 2.6, y: lg.panelTopPadded + 1.0, w: 1.3, h: 0.4,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, align: "center", valign: "middle", margin: 0,
      });
      // Result bar
      addTextOnShape(slide, "40 + 12 = 52", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.95, w: lg.rightW - 0.6, h: 0.45, rectRadius: 0.06,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      // Check mark
      slide.addText("4 x 13 = 52  Verified!", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 2.55, w: lg.rightW - 0.6, h: 0.3,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED, align: "center", margin: 0,
      });
    }
  );

  // ── Slide 9: I Do 2 — Using Distributive Property for Unknowns ──
  workedExSlide(pres, 2, "I Do", "Using the Distributive Property for Unknowns",
    [
      "Find [] in: 5 x 14 = 5 x 10 + 5 x []",
      "",
      "The distributive property says:",
      "5 x 14 = 5 x (10 + [])",
      "",
      "So 14 = 10 + []",
      "[] = 14 - 10 = 4",
      "",
      "Check: 5 x 10 + 5 x 4 = 50 + 20 = 70",
      "5 x 14 = 70  Verified!",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.SECONDARY });
      slide.addText("Key Insight", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      // Step arrows
      const steps = [
        "5 x 14 = 5 x 10 + 5 x []",
        "14 = 10 + []",
        "[] = 4",
      ];
      steps.forEach((st, i) => {
        addTextOnShape(slide, st, {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.5 + i * 0.55, w: lg.rightW - 0.6, h: 0.4, rectRadius: 0.06,
          fill: { color: i === 2 ? C.SUCCESS : C.PRIMARY },
        }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
    }
  );

  // ── Slides 10-11: CFU (withReveal) ──
  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check: Distributive Property", "Show Me Boards",
      "Complete this equation:\n\n3 x 15 = 3 x ___ + 3 x ___\n\nFill in the two blanks. Any correct split is great!",
      NOTES_CFU_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "3 x 15 = 3 x 10 + 3 x 5 = 30 + 15 = 45", {
        x: 0.5, y: 3.8, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addText("Other valid splits: 3x12 + 3x3 = 45   or   3x8 + 3x7 = 45", {
        x: 0.5, y: 4.4, w: 9, h: 0.3,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, align: "center", margin: 0,
      });
      slide.addNotes(NOTES_CFU_A);
    }
  );

  // ── Slides 12-13: We Do PP1 (withReveal) — 6 x 17 ──
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Find Unknown Values Using Distribution",
      [
        "6 x 17 = 6 x [] + 6 x 7",
        "",
        "17 has been broken into [] + 7",
        "",
        "If 17 = [] + 7, then [] = ___",
        "",
        "Check: 6 x ___ + 6 x 7 = ___ + ___ = ___",
      ],
      NOTES_WEDO1_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.8, { strip: C.SECONDARY });
        slide.addText("With your partner:", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
          fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        slide.addText([
          { text: "Identify the common factor", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Find what 17 was split into", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Solve for [] and verify", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
        ], {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.45, w: lg.rightW - 0.5, h: 1.0,
          fontFace: FONT_B, margin: 0, valign: "top",
        });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "17 = [] + 7  ->  [] = 10     Check: 60 + 42 = 102 = 6 x 17", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // ── Slides 14-15: We Do PP2 (withReveal) — Boards: Two Problems ──
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Boards: Two Problems", { y: 0.65, fontSize: 22, color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9, 2.8, { strip: STAGE_COLORS["3"] });
      s.addText([
        { text: "A)  8 x 23 = 8 x 20 + 8 x []", options: { fontSize: 17, color: C.CHARCOAL, breakLine: true, bold: true, paraSpaceAfter: 16 } },
        { text: "B)  7 x [] = 7 x 30 + 7 x 6", options: { fontSize: 17, color: C.CHARCOAL, breakLine: true, bold: true, paraSpaceAfter: 16 } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "60 seconds per problem -- work on your whiteboard!", options: { fontSize: 14, color: C.ALERT, bold: true } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.2, w: 8.5, h: 2.4,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "A) [] = 3 (23 = 20 + 3)     B) [] = 36 (30 + 6 = 36)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO2_A);
    }
  );

  // ── Slides 16-17: Hinge CFU (withReveal) — Finger Voting ──
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Which Is the Distributive Property?", { color: C.ALERT });

      // Technique pill
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
      addCard(s, 0.5, qY, 9, 3.0, { strip: C.ALERT, fill: C.WHITE });
      const options = [
        "A)  5 x 18 = 5 x 10 x 8",
        "B)  6 x 14 = 6 x 10 + 6 x 4",
        "C)  3 x 12 = 3 + 12",
        "D)  4 x 15 = 40 + 45",
      ];
      s.addText(options.map((opt, i) => ({
        text: opt,
        options: { fontSize: 16, color: C.CHARCOAL, breakLine: i < options.length - 1, paraSpaceAfter: 14, bold: i === 0 },
      })), {
        x: 0.9, y: qY + 0.2, w: 8.2, h: 2.0,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      s.addText("Hold up 1, 2, 3, or 4 fingers.", {
        x: 0.9, y: qY + 2.3, w: 8.2, h: 0.4,
        fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_HINGE_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "B)  6 x 14 = 6 x 10 + 6 x 4 = 60 + 24 = 84", {
        x: 0.5, y: 3.9, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addText("A: x not + (associative)   C: no distribution   D: wrong products (40+45=85, not 60)", {
        x: 0.5, y: 4.5, w: 9, h: 0.3,
        fontSize: 10, fontFace: FONT_B, color: C.MUTED, align: "center", margin: 0,
      });
      slide.addNotes(NOTES_HINGE_A);
    }
  );

  // ── Slide 18: You Do ──
  workedExSlide(pres, 4, "You Do", "Distributive Property",
    [
      "First: Identify the common factor.",
      "Next: Work out what the number was",
      "      broken into (find the missing addend).",
      "Then: Find [] and verify.",
      "",
      "1.  3 x 16 = 3 x 10 + 3 x []",
      "2.  9 x [] = 9 x 20 + 9 x 5",
      "3.  5 x 24 = 5 x [] + 5 x 4",
      "",
      "Complete all 6 problems on your worksheet.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.8, { strip: C.ALERT });
      slide.addText("Strategy Reminder:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "1. Find the common factor", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "2. What was the number split into?", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "3. Solve for []", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "4. Verify both sides equal", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.45, w: lg.rightW - 0.5, h: 1.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // ── Slide 19: Exit Ticket ──
  exitTicketSlide(pres,
    [
      "Use the distributive property to write 7 x 16 as the sum of two simpler multiplications.",
      "Find []:  4 x 25 = 4 x 20 + 4 x []. Show your working.",
    ],
    NOTES_EXIT, FOOTER);

  // ── Slide 20: Closing ──
  closingSlide(pres,
    "How does the distributive property help you multiply big numbers in your head?",
    [
      "I can break a multiplication into two simpler parts using an array",
      "I can write an equation showing the distributive property",
      "I can use the distributive property to find an unknown value",
    ],
    NOTES_CLOSING);

  // ── Slide 21: Resources ──
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // ── Write PPTX ──
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "ALG6_Session3_Distributive_Property.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Generate PDFs ─────────────────────────────────────────────────────────

  // Session 3 Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "The Distributive Property",
      color: C.NAVY,
      lessonInfo: "Session 3 of 6 | Algebra: Unknown Values | Year 5/6",
    });
    y = addTipBox(doc, "For each equation: identify the common factor, work out what the number was broken into, find the unknown, and verify by checking both sides are equal.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A: Find the Missing Addend", y, { color: C.NAVY });
    y = addProblem(doc, 1, "3 x 16 = 3 x 10 + 3 x []          [] = ___", y, { color: C.NAVY, writeLines: [{ label: "Working:" }, { label: "Verify:" }] });
    y = addProblem(doc, 2, "4 x 18 = 4 x 10 + 4 x []          [] = ___", y, { color: C.NAVY, writeLines: [{ label: "Working:" }, { label: "Verify:" }] });

    y = addSectionHeading(doc, "Section B: Find the Whole Number", y, { color: C.NAVY });
    y = addProblem(doc, 3, "9 x [] = 9 x 20 + 9 x 5           [] = ___", y, { color: C.NAVY, writeLines: [{ label: "Working:" }, { label: "Verify:" }] });
    y = addProblem(doc, 4, "6 x [] = 6 x 30 + 6 x 2           [] = ___", y, { color: C.NAVY, writeLines: [{ label: "Working:" }, { label: "Verify:" }] });

    y = addSectionHeading(doc, "Section C: Find the Missing Part", y, { color: C.NAVY });
    y = addProblem(doc, 5, "5 x 24 = 5 x [] + 5 x 4           [] = ___", y, { color: C.NAVY, writeLines: [{ label: "Working:" }, { label: "Verify:" }] });
    y = addProblem(doc, 6, "8 x 15 = 8 x [] + 8 x 5           [] = ___", y, { color: C.NAVY, writeLines: [{ label: "Working:" }, { label: "Verify:" }] });

    addPdfFooter(doc, "Session 3 | Algebra: Finding Unknown Values | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Session 3 Enabling Scaffold
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Array-Based Scaffolding for the Distributive Property",
      color: C.TEAL,
      lessonInfo: "Session 3 of 6 | Algebra: Unknown Values | Year 5/6",
    });
    y = addTipBox(doc, "Each problem has an array drawn for you. Count the columns in each section to find the missing value. The first one is done as an example.", y, { color: C.TEAL });

    // Example
    y = addSectionHeading(doc, "Example (done for you)", y, { color: C.NAVY });
    y = addBodyText(doc, "6 x 17 = 6 x [] + 6 x 7", y);
    y += 5;
    // Draw array
    const arrX = 80;
    doc.rect(arrX, y, 140, 50).fill("#" + C.PRIMARY).stroke("#" + C.PRIMARY);
    doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF").text("6 x 10 = 60", arrX + 10, y + 18, { width: 120 });
    doc.rect(arrX + 150, y, 60, 50).fill("#" + C.ACCENT).stroke("#" + C.ACCENT);
    doc.fillColor("#FFFFFF").text("6 x 7 = 42", arrX + 155, y + 18, { width: 50 });
    y += 58;
    doc.fontSize(11).font("Sans-Bold").fillColor("#" + C.SUCCESS).text("10 + 7 = 17, so [] = 10.  Check: 60 + 42 = 102 = 6 x 17", arrX, y, { width: 400 });
    y += 25;

    y = addSectionHeading(doc, "Now you try:", y, { color: C.NAVY });

    const scaffoldProblems = [
      { eq: "3 x 16 = 3 x 10 + 3 x []", leftW: 140, leftLabel: "3 x 10", rightW: 80, rightLabel: "3 x ?" },
      { eq: "9 x 25 = 9 x 20 + 9 x []", leftW: 140, leftLabel: "9 x 20", rightW: 70, rightLabel: "9 x ?" },
      { eq: "5 x 24 = 5 x [] + 5 x 4", leftW: 140, leftLabel: "5 x ?", rightW: 50, rightLabel: "5 x 4" },
      { eq: "4 x 18 = 4 x 10 + 4 x []", leftW: 140, leftLabel: "4 x 10", rightW: 100, rightLabel: "4 x ?" },
    ];

    scaffoldProblems.forEach((p, i) => {
      y = addBodyText(doc, (i + 1) + ". " + p.eq, y);
      y += 3;
      doc.rect(arrX, y, p.leftW, 40).lineWidth(1).stroke("#" + C.PRIMARY);
      doc.fontSize(9).font("Sans").fillColor("#999999").text(p.leftLabel, arrX + 5, y + 14, { width: p.leftW - 10 });
      doc.rect(arrX + p.leftW + 5, y, p.rightW, 40).lineWidth(1).stroke("#" + C.ACCENT);
      doc.text(p.rightLabel, arrX + p.leftW + 10, y + 14, { width: p.rightW - 10 });
      y += 48;
      doc.fillColor("#333333");
      y = addWriteLine(doc, "[] = ", y);
      y += 8;
      if (y > 720) { doc.addPage(); y = 50; }
    });

    addPdfFooter(doc, "Session 3 | Enabling Scaffold | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Session 3 Extension — Distribution Over Subtraction
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Does the Distributive Property Work with Subtraction?",
      color: C.NAVY,
      lessonInfo: "Session 3 of 6 | Algebra: Unknown Values | Year 5/6",
    });

    y = addSectionHeading(doc, "What You Already Know", y, { color: C.NAVY });
    y = addBodyText(doc, "The distributive property says you can break a multiplication apart using ADDITION:", y);
    y = addBodyText(doc, "4 x 13 = 4 x (10 + 3) = 4 x 10 + 4 x 3 = 40 + 12 = 52", y, { italic: true });
    y += 5;

    y = addSectionHeading(doc, "The Big Question", y, { color: C.NAVY });
    y = addBodyText(doc, "Does the distributive property also work with SUBTRACTION? Can you break a number apart using subtraction and still get the right answer?", y);
    y += 5;

    y = addSectionHeading(doc, "Worked Example", y, { color: C.NAVY });
    y = addBodyText(doc, "Let's test: 7 x 99", y);
    y = addBodyText(doc, "99 = 100 - 1, so: 7 x 99 = 7 x (100 - 1) = 7 x 100 - 7 x 1 = 700 - 7 = 693", y);
    y = addBodyText(doc, "Check: 7 x 99 = 693? Let's verify: 7 x 100 = 700, minus 7 = 693. Yes!", y);
    y += 5;

    y = addSectionHeading(doc, "Another Example", y, { color: C.NAVY });
    y = addBodyText(doc, "6 x 48 = 6 x (50 - 2) = 6 x 50 - 6 x 2 = 300 - 12 = 288", y);
    y = addBodyText(doc, "Check: 6 x 48 = 288? Yes!", y);
    y += 5;

    y = addSectionHeading(doc, "Your Investigation", y, { color: C.NAVY });
    y = addTipBox(doc, "Use the distributive property with subtraction to solve these. Show all your working.", y, { color: C.TEAL });
    y = addProblem(doc, 1, "Calculate 5 x 98 using: 98 = 100 - 2", y, { color: C.NAVY, writeLines: [{ label: "Working:" }, { label: "Answer:" }] });
    y = addProblem(doc, 2, "Calculate 8 x 47 using: 47 = 50 - 3", y, { color: C.NAVY, writeLines: [{ label: "Working:" }, { label: "Answer:" }] });
    y = addProblem(doc, 3, "Calculate 9 x 199 using subtraction. Choose your own split!", y, { color: C.NAVY, writeLines: [{ label: "Split:" }, { label: "Working:" }, { label: "Answer:" }] });

    y = addSectionHeading(doc, "Did You Know?", y, { color: C.NAVY });
    y = addBodyText(doc, "The full distributive property works with BOTH addition and subtraction:", y);
    y = addBodyText(doc, "a x (b + c) = a x b + a x c     AND     a x (b - c) = a x b - a x c", y, { italic: true });
    y = addBodyText(doc, "This is why you can do mental maths tricks like 7 x 99 = 700 - 7. The distributive property is the reason it works!", y);

    addPdfFooter(doc, "Session 3 | Extension Investigation | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("\nSession 3 build complete!");
}

build().catch(err => { console.error(err); process.exit(1); });
