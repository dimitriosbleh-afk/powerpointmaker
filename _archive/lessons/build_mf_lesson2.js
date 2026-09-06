"use strict";

// Multiplication Facts Unit — Lesson 2: x3, x4 (Building from Known Facts)
// Session 2, Grade 3/4 Numeracy, Variant 0
// DR: Addition Algorithm (Vertical, 4- & 5-digits, with/without renaming)
// Fluency: Subtraction Algorithm (Vertical, 3- & 4-digits, with/without renaming)

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade34", 0);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addBadge, addTitle,
  withReveal, STAGE_COLORS,
} = T;

const SESSION = 2;
const FOOTER = "Multiplication Facts | Session 2 of 4 | Grade 3/4 Numeracy";
const OUT_DIR = "output/MF_Session2_Derived_Facts_x3_x4";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Independent practice - x3 and x4 multiplication facts.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with worked answers.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Pre-filled doubling bridge for x3 and x4 facts.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Self-contained investigation into triangular numbers.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ───────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Yesterday we nailed our anchor facts: x2, x5, x10
- Today we build from those. We are using what we already know to figure out x3 and x4
- The big idea: you do not need to memorise every fact from scratch. You can BUILD new facts from old ones

DO:
- Display title slide
- Have whiteboards ready

TEACHER NOTES:
Session 2 of 4. This lesson introduces derived-fact strategies: x3 = x2 + one more group, x4 = double-double (x2 then x2 again). Both build from x2 which was consolidated in Session 1.

WATCH FOR:
- Students who are anxious about "harder" tables -- reassure: "You already know x2. That is your launchpad."

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Daily review -- addition algorithm
- Four problems on your whiteboards. Set out vertically, rename where needed.

DO:
- Display problems. Allow 90 seconds each
- These are different numbers from Session 1 DR to provide fresh retrieval practice
- Circulate, check column alignment and renaming

TEACHER NOTES:
DR continues addition algorithm practice. Problems include 5-digit numbers requiring multiple renamings. This is spaced retrieval of prior learning, not today's focus.

WATCH FOR:
- Students who made renaming errors yesterday -- check they are self-correcting today
- Students who rush and misalign columns

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers. Tick correct, fix errors.
- Fixing is the learning.

DO:
- Reveal answers. Students self-check.
- Quick scan -- note improvement from yesterday

TEACHER NOTES:
Tick-and-fix cycle. Compare error patterns with Session 1.

WATCH FOR:
- Improvement in renaming accuracy compared to Session 1

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time -- subtraction vertical algorithm
- Six problems, whiteboards, 2 minutes. Go.

DO:
- Display problems. Time 2 minutes.
- Cold call answers after

TEACHER NOTES:
Fluency drills subtraction algorithm. Problems include renaming across zeros (e.g., 4 000 - 1 256) which is the hardest subtraction pattern.

WATCH FOR:
- Students who freeze on subtraction with zeros -- prompt: "Rename from the column to the left"

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention together
- Read from slide: "We are learning to use doubling strategies to work out x3 and x4 multiplication facts"
- Success criteria -- read together. Ask: What does SC2 mean by 'double-double'? [Take responses -- we will explore this today]

DO:
- Choral read LI and SCs

TEACHER NOTES:
x3 and x4 are derived from x2. SC1 is the floor (connecting to x2), SC2 is the core (applying the strategies), SC3 extends to choosing between strategies flexibly.

WATCH FOR:
- Students who do not see the connection between x2 and x3/x4 yet -- that is expected, the I Do will make it explicit

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO_X4 = `SAY:
- Let me show you the x4 strategy first because it is the clearest
- x4 means 4 groups. But 4 is 2 x 2. So x4 is double-double.
- Watch: 4 x 7
- Step 1: Double 7. That is 14. (That is 2 x 7)
- Step 2: Double 14. That is 28. (That is 2 x 14, which is 4 x 7)
- So 4 x 7 = 28. I just doubled twice.

DO:
- Write each step clearly on screen
- Point to the visual showing the doubling bridge
- Emphasise: "Double it, then double it again"

TEACHER NOTES:
Double-double for x4 is the most intuitive derived-fact strategy. Starting with x4 (not x3) gives students the clearest first experience of building from x2. The array visual shows 4 rows split into two groups of 2 rows.

MISCONCEPTIONS:
- Misconception: 4 x 7 = 4 + 7 = 11 (addition instead of multiplication)
  Why: Students default to addition when uncertain about multiplication
  Impact: Fundamental operation confusion that undermines all multiplication work
  Quick correction: "4 x 7 means 4 GROUPS of 7. Not 4 plus 7. Draw 4 groups with 7 in each."

WATCH FOR:
- Students who double correctly the first time but struggle with the second double (doubling a teen number) -- they may need a bridging strategy: double 14 = double 10 + double 4 = 20 + 8 = 28

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_X3 = `SAY:
- Now x3. Three groups of something.
- Here is the strategy: x3 = x2 + one more group
- Watch: 3 x 6
- Step 1: I already know 2 x 6 = 12 (that is my x2 anchor fact)
- Step 2: Add one more group of 6: 12 + 6 = 18
- So 3 x 6 = 18
- Let me check with skip counting: 6, 12, 18. Yes.

DO:
- Write each step
- Circle the x2 fact that serves as the bridge
- Show the array: 2 rows of 6 plus 1 more row of 6

TEACHER NOTES:
x3 = x2 + one more group. This is slightly harder than double-double because the second step is addition of a different number (not doubling). The array makes the +1 group visible.

WATCH FOR:
- Students who add the wrong "one more group" (e.g., for 3 x 6, they add 3 instead of 6) -- prompt: "What is the group size? 6. So add one more group of 6."
- Readiness signal: students nodding along and mouthing the steps

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. Thumbs up if true, thumbs down if false.
- Statement: "4 x 5 is the same as doubling 5, then doubling again"

DO:
- Students show thumbs
- Scan for thumbs up (correct)

CFU CHECKPOINT:
Technique: Thumbs Up/Down
Script:
- Say: "Thumbs up if this is TRUE, thumbs down if FALSE: 4 x 5 is the same as doubling 5, then doubling again."
- Scan for: thumbs up. Double 5 = 10, double 10 = 20. 4 x 5 = 20. True.
PROCEED: If 80%+ show thumbs up, move to We Do.
PIVOT: Most likely misconception -- students think 4 x 5 = 9 (adding) or that double-double does not apply to x4. Reteach with a simpler number: "Try 4 x 2. Double 2 is 4. Double 4 is 8. Count: 2, 4, 6, 8. Yes -- 4 x 2 = 8."

TEACHER NOTES:
True/false format keeps the check fast and checks conceptual understanding of the double-double strategy rather than calculation.

WATCH FOR:
- Students who are unsure and wait to copy neighbours -- use "eyes closed, thumbs up" to get honest data

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Partner work. Two problems on your whiteboards.
- Problem 1: 4 x 8. Use the double-double strategy. Show both steps.
- Problem 2: 3 x 7. Use the x2 + one more group strategy. Show both steps.
- Write the strategy steps, not just the answer.

DO:
- Display both problems
- Partners work together on whiteboards
- Circulate: check for correct strategy steps, not just answers
- Allow 90 seconds

TEACHER NOTES:
We Do pairs one x4 and one x3 problem. Requiring written strategy steps makes the derived-fact process visible and checkable. Different numbers from the I Do examples ensure transfer.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the enabling scaffold which shows the doubling bridge pre-filled for the first step. Students only need to complete the second step. For 4 x 8: "Double 8 = 16" is pre-filled, students write "Double 16 = ___".
- Extra Notes: Distribute Session 2 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Find TWO different ways to work out 4 x 9 using known facts (e.g., double-double AND x5 - one group). Show both methods and check they give the same answer.
- Extra Notes: This develops flexible strategy use -- a key step toward number sense.

WATCH FOR:
- Students who get the answer right but skip writing the strategy steps -- redirect: "Show me the double-double steps"
- Students who confuse which number to double (doubling the 4 instead of the 8 for 4 x 8) -- prompt: "Which number are we making groups of? 8. So double 8 first."

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Check your working.
- 4 x 8: double 8 = 16, double 16 = 32. So 4 x 8 = 32.
- 3 x 7: I know 2 x 7 = 14. One more group of 7: 14 + 7 = 21. So 3 x 7 = 21.
- Who got both right? Who got the strategy steps written out?

DO:
- Reveal answers
- Cold call: "Show me your whiteboard" for 2-3 students
- Celebrate correct strategy steps, not just correct answers

TEACHER NOTES:
Confirming both the answer and the working reinforces the process. Students who got the right answer via a different valid method (e.g., skip counting for x3) should be acknowledged but encouraged to also try the derived-fact strategy.

WATCH FOR:
- Students who skip counted instead of using the derived strategy -- the answer is correct but the strategy practice is missed. Prompt: "Can you show me how you would do it with the doubling bridge?"

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check time. On your whiteboard, solve this and show your working.
- 3 x 9 = ___. Use the x2 + one more group strategy.

DO:
- Students write on whiteboards. Allow 30 seconds.
- Signal: hold up boards

CFU CHECKPOINT:
Technique: Show Me Boards with Justification
Script:
- Say: "3 x 9. Use x2 + one more group. Show both steps. 30 seconds, then boards up."
- Scan for: Step 1: 2 x 9 = 18. Step 2: 18 + 9 = 27. Answer: 27.
PROCEED: If 80%+ show correct working AND answer, move to You Do.
PIVOT: Most likely misconception -- students write 2 x 9 = 18 but then add 18 + 3 = 21 (adding the multiplier instead of one more group of 9). Reteach: "One more GROUP of 9. The group size is 9. You had 2 groups (18), now add a third group: 18 + 9 = 27. Not 18 + 3."

TEACHER NOTES:
This hinge checks execution of the x3 strategy, which is the harder of the two derived strategies because students must add the original group size (not double). The +3 error is the most diagnostic.

WATCH FOR:
- The +3 error (adding the multiplier instead of one more group) -- this is the threshold misconception
- Students who got 27 via skip counting -- correct answer but did not practise the strategy

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice time. Open your Session 2 Worksheet.
- First: identify whether the problem is x3 or x4.
- Next: choose the right strategy (x2 + one more group OR double-double).
- Then: show your working steps and write the answer.
- You have 8-10 minutes.

DO:
- Distribute worksheets
- For enabling students, distribute the Session 2 Enabling Scaffold
- Circulate, prioritise students who struggled with the hinge check

TEACHER NOTES:
You Do uses different numbers from the We Do and mixes x3 and x4 problems (unlike the We Do which separated them). This requires strategy selection as well as execution.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Complete the enabling scaffold with pre-filled first steps. For x4 problems, the first double is shown; students complete the second. For x3 problems, the x2 fact is shown; students add one more group.
- Extra Notes: Distribute Session 2 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Investigate triangular numbers -- numbers you make by adding 1+2+3+4+... (1, 3, 6, 10, 15...). Find the first 10 and look for connections to multiplication facts.
- Extra Notes: Distribute Session 2 Extension PDF.

WATCH FOR:
- Students who apply the wrong strategy to the wrong table (double-double for x3) -- prompt: "Is this x3 or x4? Which strategy goes with each?"
- Students finishing quickly with all correct -- direct to the extension investigation

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Four questions. Work on your own, 3 minutes.

DO:
- Display exit ticket. Students work in workbooks.
- Collect/sort into: secure (all 4 correct with working), developing (2-3 correct), beginning (0-1).

TEACHER NOTES:
Q1-Q2 assess x4 (double-double). Q3-Q4 assess x3 (x2 + one more group). Q4 requires strategy explanation (SC2). Sort results for Session 3 planning.

WATCH FOR:
- Students who get answers right without showing strategy steps -- note for conferencing
- Students who confuse x3 and x4 strategies

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Check our success criteria.
- SC1: I can use my x2 facts as a starting point for x3 and x4. Thumbs?
- SC2: I can explain the double-double strategy for x4 and the x2 + one more group strategy for x3. Thumbs?
- SC3: I can choose the right strategy when I see a x3 or x4 problem. Thumbs?
- Turn and talk: Which strategy did you find easier -- double-double or x2 + one more group? Why?

DO:
- Display SCs. Run thumbs for each.
- 30 seconds Turn and Talk, then cold call 2 pairs

TEACHER NOTES:
Most students find double-double (x4) easier because both steps are doubling. The x3 strategy requires adding a different number, which is cognitively harder. This is useful data for Session 3 planning.

WATCH FOR:
- Students who prefer skip counting and resist the derived strategies -- they may need more convincing that derived strategies are faster for larger numbers

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Resources for today's lesson are linked here

DO:
- Point out each resource

TEACHER NOTES:
Resource slide for Session 2 companion PDFs.

WATCH FOR:
- N/A

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ───────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, "Multiplication Facts", "Session 2: Building From Known Facts — x3, x4", "Grade 3/4 Numeracy | Session 2 of 4", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2-3: Daily Review (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Addition Algorithm Review", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const problems = [
        "1.  3 614 + 5 273 =",
        "2.  6 458 + 1 987 =",
        "3.  23 506 + 8 794 =",
        "4.  45 189 + 27 836 =",
      ];
      s.addText(problems.map((p, i) => ({
        text: p,
        options: { fontSize: 15, color: C.CHARCOAL, breakLine: i < problems.length - 1, paraSpaceAfter: 12 },
      })), {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboards:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "Solve using the vertical algorithm.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Remember:", options: { fontSize: 14, bold: true, color: C.ALERT, breakLine: true } },
        { text: "Line up place value columns", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Start from the ones", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Rename when a column totals 10+", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Answers:  1) 8 887    2) 8 445    3) 32 300    4) 73 025", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency
  const sFluency = pres.addSlide();
  addTopBar(sFluency, STAGE_COLORS["1"]);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Subtraction Algorithm Practice", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

  addCard(sFluency, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  const subProblems = [
    "1.  483 - 251 =",
    "2.  726 - 389 =",
    "3.  2 034 - 876 =",
    "4.  4 000 - 1 256 =",
    "5.  6 301 - 2 845 =",
    "6.  8 100 - 4 567 =",
  ];
  sFluency.addText(subProblems.map((p, i) => ({
    text: p,
    options: { fontSize: 14, color: C.CHARCOAL, breakLine: i < subProblems.length - 1, paraSpaceAfter: 8 },
  })), {
    x: 0.75, y: CONTENT_TOP + 0.12, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.24,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addCard(sFluency, 5.2, CONTENT_TOP, 4.3, 1.6, { strip: C.SECONDARY });
  sFluency.addText([
    { text: "Whiteboards - 2 minutes", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
    { text: "Set out vertically. Rename where needed.", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
    { text: "Watch for zeros -- rename carefully!", options: { fontSize: 13, color: C.ALERT, bold: true } },
  ], {
    x: 5.45, y: CONTENT_TOP + 0.12, w: 3.8, h: 1.36,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    ["We are learning to use doubling strategies to work out x3 and x4 multiplication facts"],
    [
      "I can use my x2 facts as a starting point for x3 and x4",
      "I can explain the double-double strategy for x4 and the x2 + one more group strategy for x3",
      "I can choose the right strategy when I see a x3 or x4 problem",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — x4 (double-double)
  workedExSlide(pres, 2, "I Do", "x4 Facts: The Double-Double Strategy",
    [
      "x4 means 4 groups",
      "4 = 2 x 2, so x4 = double, then double again",
      "",
      "Example: 4 x 7",
      "  Step 1: Double 7 = 14  (that is 2 x 7)",
      "  Step 2: Double 14 = 28  (that is 2 x 14)",
      "  Answer: 4 x 7 = 28",
    ],
    NOTES_IDO_X4, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.8, { strip: C.PRIMARY });
      slide.addText("The Doubling Bridge", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      // Visual: 7 -> 14 -> 28
      const boxW = 0.9;
      const boxH = 0.45;
      const startX2 = lg.rightX + 0.15;
      const midX = lg.rightX + 1.55;
      const endX = lg.rightX + 2.95;
      const boxY = lg.panelTopPadded + 0.55;

      addTextOnShape(slide, "7", { x: startX2, y: boxY, w: boxW, h: boxH, rectRadius: 0.08, fill: { color: C.SECONDARY } },
        { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      addTextOnShape(slide, "14", { x: midX, y: boxY, w: boxW, h: boxH, rectRadius: 0.08, fill: { color: C.PRIMARY } },
        { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      addTextOnShape(slide, "28", { x: endX, y: boxY, w: boxW, h: boxH, rectRadius: 0.08, fill: { color: C.ALERT } },
        { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });

      // Arrow labels between boxes
      slide.addText("x2", {
        x: startX2 + boxW + 0.02, y: boxY + 0.08, w: midX - startX2 - boxW - 0.04, h: 0.2,
        fontSize: 10, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, align: "center", bold: true,
      });
      slide.addText("x2", {
        x: midX + boxW + 0.02, y: boxY + 0.08, w: endX - midX - boxW - 0.04, h: 0.2,
        fontSize: 10, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, align: "center", bold: true,
      });

      // Array: 4 rows of 7
      const dotR = 0.14;
      const dotGap = 0.25;
      const arrStartX = lg.rightX + 0.3;
      const arrStartY = lg.panelTopPadded + 1.3;
      for (let row = 0; row < 4; row++) {
        const rowColor = row < 2 ? C.PRIMARY : C.SECONDARY;
        for (let col = 0; col < 7; col++) {
          slide.addShape("roundRect", {
            x: arrStartX + col * dotGap, y: arrStartY + row * dotGap,
            w: dotR, h: dotR, rectRadius: dotR / 2,
            fill: { color: rowColor },
          });
        }
      }
      slide.addText("2 x 7 = 14", {
        x: lg.rightX + 0.2, y: arrStartY + 1.1, w: 1.8, h: 0.22,
        fontSize: 10, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
      });
      slide.addText("+ 2 x 7 = 14", {
        x: lg.rightX + 0.2, y: arrStartY + 1.3, w: 1.8, h: 0.22,
        fontSize: 10, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
      });
      slide.addText("= 4 x 7 = 28", {
        x: lg.rightX + 0.2, y: arrStartY + 1.55, w: 1.8, h: 0.22,
        fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
      });
    }
  );

  // Slide 7: I Do — x3
  workedExSlide(pres, 2, "I Do", "x3 Facts: x2 + One More Group",
    [
      "x3 means 3 groups",
      "Strategy: work out x2 first, then add one more group",
      "",
      "Example: 3 x 6",
      "  Step 1: 2 x 6 = 12  (my x2 anchor fact)",
      "  Step 2: 12 + 6 = 18  (add one more group of 6)",
      "  Answer: 3 x 6 = 18",
      "",
      "Check: 6, 12, 18. Yes!",
    ],
    NOTES_IDO_X3, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.6, { strip: C.SECONDARY });
      slide.addText("x2 + One More Group", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      // Array: 3 rows of 6
      const dotR = 0.16;
      const dotGap = 0.3;
      const arrStartX = lg.rightX + 0.45;
      const arrStartY = lg.panelTopPadded + 0.55;
      for (let row = 0; row < 3; row++) {
        const rowColor = row < 2 ? C.PRIMARY : C.ACCENT;
        for (let col = 0; col < 6; col++) {
          slide.addShape("roundRect", {
            x: arrStartX + col * dotGap, y: arrStartY + row * dotGap,
            w: dotR, h: dotR, rectRadius: dotR / 2,
            fill: { color: rowColor },
          });
        }
      }
      // Bracket labels
      slide.addText("2 x 6 = 12", {
        x: lg.rightX + 0.2, y: arrStartY + 1.0, w: 1.8, h: 0.22,
        fontSize: 10, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
      });
      slide.addText("+ 1 x 6 = 6", {
        x: lg.rightX + 0.2, y: arrStartY + 1.22, w: 1.8, h: 0.22,
        fontSize: 10, fontFace: FONT_B, color: C.ACCENT, bold: true, margin: 0,
      });
      slide.addText("= 3 x 6 = 18", {
        x: lg.rightX + 0.2, y: arrStartY + 1.48, w: 1.8, h: 0.22,
        fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
      });
    }
  );

  // Slide 8-9: CFU 1 (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check", "Thumbs Up / Down",
      "True or false?\n\n\"4 x 5 is the same as doubling 5, then doubling again.\"",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "TRUE: Double 5 = 10. Double 10 = 20. 4 x 5 = 20.", {
        x: 1.5, y: 4.0, w: 7, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10-11: We Do (withReveal)
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Your Turn — With a Partner",
      [
        "Show your strategy steps on your whiteboard.",
        "",
        "Problem 1 (double-double):",
        "  4 x 8 = ___",
        "  Step 1: Double 8 = ___",
        "  Step 2: Double ___ = ___",
        "",
        "Problem 2 (x2 + one more group):",
        "  3 x 7 = ___",
        "  Step 1: 2 x 7 = ___",
        "  Step 2: ___ + 7 = ___",
      ],
      NOTES_WEDO_Q, FOOTER),
    (slide) => {
      addTextOnShape(slide, "1) 4 x 8 = 32 (double 8=16, double 16=32)   2) 3 x 7 = 21 (2x7=14, 14+7=21)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 12-13: CFU 2 Hinge (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Check — Show Your Working", "Show Me Boards",
      "3 x 9 = ___\n\nUse the x2 + one more group strategy.\nShow BOTH steps on your whiteboard.",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Step 1: 2 x 9 = 18.  Step 2: 18 + 9 = 27.  Answer: 27", {
        x: 1.0, y: 4.0, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do
  workedExSlide(pres, 4, "You Do", "You Do — x3 and x4 Facts",
    [
      "First: Check if the problem is x3 or x4.",
      "Next: Choose the right strategy.",
      "  x4 -> double-double",
      "  x3 -> x2 + one more group",
      "Then: Show your strategy steps and write the answer.",
      "",
      "Work on your own. 8-10 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.2, { strip: C.ALERT });
      slide.addText("Strategy Reminder", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      const reminders = [
        { label: "x4", desc: "Double, then double again", color: C.PRIMARY },
        { label: "x3", desc: "x2 + one more group", color: C.SECONDARY },
      ];
      reminders.forEach((r, i) => {
        const ry = lg.panelTopPadded + 0.5 + i * 0.7;
        addTextOnShape(slide, r.label, {
          x: lg.rightX + 0.2, y: ry, w: 0.8, h: 0.4, rectRadius: 0.06,
          fill: { color: r.color },
        }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(r.desc, {
          x: lg.rightX + 1.15, y: ry, w: lg.rightW - 1.35, h: 0.4,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "middle",
        });
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "4 x 6 = ___  (show your double-double steps)",
      "3 x 8 = ___  (show your x2 + one more group steps)",
      "4 x 9 = ___",
      "Explain: Why does double-double work for x4?",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 16: Closing
  closingSlide(pres,
    "Which strategy did you find easier — double-double for x4 or x2 + one more group for x3? Tell your partner why.",
    [
      "I can use my x2 facts as a starting point for x3 and x4",
      "I can explain double-double (x4) and x2 + one more group (x3)",
      "I can choose the right strategy for x3 or x4 problems",
    ],
    NOTES_CLOSING);

  // Slide 17: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "MF_Session2_Derived_Facts_x3_x4.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Generate PDFs ─────────────────────────────────────────────────────────

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Multiplication Facts — x3 and x4",
      color: C.NAVY,
      lessonInfo: "Session 2 of 4 | Grade 3/4 Numeracy",
    });
    y = addTipBox(doc, "Show your strategy steps for each problem. x4 = double-double. x3 = x2 + one more group.", y, { color: C.TEAL });
    y = addSectionHeading(doc, "x4 Facts (Double-Double)", y, { color: C.NAVY });
    y = addProblem(doc, 1, "4 x 3 = ___    Step 1: Double 3 = ___    Step 2: Double ___ = ___", y, { color: C.NAVY });
    y = addProblem(doc, 2, "4 x 5 = ___    Step 1: Double 5 = ___    Step 2: Double ___ = ___", y, { color: C.NAVY });
    y = addProblem(doc, 3, "4 x 8 = ___    Steps: ___", y, { color: C.NAVY });
    y = addProblem(doc, 4, "4 x 12 = ___   Steps: ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "x3 Facts (x2 + One More Group)", y, { color: C.NAVY });
    y = addProblem(doc, 5, "3 x 4 = ___    Step 1: 2 x 4 = ___    Step 2: ___ + 4 = ___", y, { color: C.NAVY });
    y = addProblem(doc, 6, "3 x 6 = ___    Step 1: 2 x 6 = ___    Step 2: ___ + 6 = ___", y, { color: C.NAVY });
    y = addProblem(doc, 7, "3 x 9 = ___    Steps: ___", y, { color: C.NAVY });
    y = addProblem(doc, 8, "3 x 11 = ___   Steps: ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "Mixed Practice", y, { color: C.NAVY });
    y = addProblem(doc, 9, "4 x 7 = ___", y, { color: C.NAVY });
    y = addProblem(doc, 10, "3 x 5 = ___", y, { color: C.NAVY });
    y = addProblem(doc, 11, "4 x 11 = ___", y, { color: C.NAVY });
    y = addProblem(doc, 12, "3 x 12 = ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addProblem(doc, 13, "Can you find TWO ways to work out 4 x 6? Show both methods.", y, { color: C.NAVY, writeLines: [{ label: "" }, { label: "" }] });
    addPdfFooter(doc, "Session 2 | Multiplication Facts | Grade 3/4");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Multiplication Facts — x3 and x4",
      color: C.NAVY,
      lessonInfo: "Session 2 of 4 | Grade 3/4 Numeracy",
    });
    y = addSectionHeading(doc, "x4 Facts", y, { color: C.NAVY });
    y = addBodyText(doc, "1. 4 x 3 = 12 (double 3=6, double 6=12)", y);
    y = addBodyText(doc, "2. 4 x 5 = 20 (double 5=10, double 10=20)", y);
    y = addBodyText(doc, "3. 4 x 8 = 32 (double 8=16, double 16=32)", y);
    y = addBodyText(doc, "4. 4 x 12 = 48 (double 12=24, double 24=48)", y);
    y = addSectionHeading(doc, "x3 Facts", y, { color: C.NAVY });
    y = addBodyText(doc, "5. 3 x 4 = 12 (2x4=8, 8+4=12)", y);
    y = addBodyText(doc, "6. 3 x 6 = 18 (2x6=12, 12+6=18)", y);
    y = addBodyText(doc, "7. 3 x 9 = 27 (2x9=18, 18+9=27)", y);
    y = addBodyText(doc, "8. 3 x 11 = 33 (2x11=22, 22+11=33)", y);
    y = addSectionHeading(doc, "Mixed", y, { color: C.NAVY });
    y = addBodyText(doc, "9. 4 x 7 = 28 (double 7=14, double 14=28)", y);
    y = addBodyText(doc, "10. 3 x 5 = 15 (2x5=10, 10+5=15)", y);
    y = addBodyText(doc, "11. 4 x 11 = 44 (double 11=22, double 22=44)", y);
    y = addBodyText(doc, "12. 3 x 12 = 36 (2x12=24, 24+12=36)", y);
    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "13. Method 1: double-double (double 6=12, double 12=24). Method 2: 3x6=18 then +6=24, or skip count 4,8,12,16,20,24. Both = 24.", y);
    addPdfFooter(doc, "Session 2 | Multiplication Facts | Grade 3/4");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Doubling Bridge Support for x3 and x4",
      color: C.TEAL,
      lessonInfo: "Session 2 of 4 | Grade 3/4 Numeracy",
    });
    y = addTipBox(doc, "The first step is done for you. Complete the second step to find the answer.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "x4 Facts: Double-Double", y, { color: C.NAVY });
    y = addBodyText(doc, "4 x 3:  Step 1: Double 3 = 6    Step 2: Double 6 = ___", y);
    y = addBodyText(doc, "4 x 5:  Step 1: Double 5 = 10   Step 2: Double 10 = ___", y);
    y = addBodyText(doc, "4 x 7:  Step 1: Double 7 = 14   Step 2: Double 14 = ___", y);
    y = addBodyText(doc, "4 x 8:  Step 1: Double 8 = 16   Step 2: Double 16 = ___", y);
    y = addBodyText(doc, "4 x 9:  Step 1: Double 9 = 18   Step 2: Double 18 = ___", y);

    y = addSectionHeading(doc, "x3 Facts: x2 + One More Group", y, { color: C.NAVY });
    y = addBodyText(doc, "3 x 4:  Step 1: 2 x 4 = 8      Step 2: 8 + 4 = ___", y);
    y = addBodyText(doc, "3 x 6:  Step 1: 2 x 6 = 12     Step 2: 12 + 6 = ___", y);
    y = addBodyText(doc, "3 x 7:  Step 1: 2 x 7 = 14     Step 2: 14 + 7 = ___", y);
    y = addBodyText(doc, "3 x 8:  Step 1: 2 x 8 = 16     Step 2: 16 + 8 = ___", y);
    y = addBodyText(doc, "3 x 9:  Step 1: 2 x 9 = 18     Step 2: 18 + 9 = ___", y);

    addPdfFooter(doc, "Session 2 | Enabling Scaffold | Grade 3/4");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Extension — Triangular Numbers
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Triangular Numbers Investigation",
      color: C.NAVY,
      lessonInfo: "Session 2 of 4 | Grade 3/4 Numeracy",
    });

    y = addSectionHeading(doc, "What Are Triangular Numbers?", y, { color: C.NAVY });
    y = addBodyText(doc, "A triangular number is what you get when you add up counting numbers starting from 1.", y);
    y = addBodyText(doc, "1 = 1 (first triangular number)", y);
    y = addBodyText(doc, "1 + 2 = 3 (second triangular number)", y);
    y = addBodyText(doc, "1 + 2 + 3 = 6 (third triangular number)", y);
    y = addBodyText(doc, "1 + 2 + 3 + 4 = 10 (fourth triangular number)", y);
    y = addBodyText(doc, "They are called triangular because you can arrange dots in a triangle shape.", y);

    y = addSectionHeading(doc, "Your Investigation", y, { color: C.NAVY });
    y = addBodyText(doc, "1. Find the first 10 triangular numbers.", y);
    y = addBodyText(doc, "2. Are any triangular numbers also in the x3 times table? Which ones?", y);
    y = addBodyText(doc, "3. Can you spot any other pattern?", y);
    y = addWriteLine(doc, "Triangular numbers:", y);
    y = addWriteLine(doc, "In the x3 table:", y);
    y = addWriteLine(doc, "Pattern I noticed:", y);
    y = addWriteLine(doc, "", y);

    addPdfFooter(doc, "Session 2 | Extension | Grade 3/4");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 2 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
