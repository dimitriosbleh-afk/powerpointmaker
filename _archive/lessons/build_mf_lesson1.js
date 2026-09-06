"use strict";

// Multiplication Facts Unit — Lesson 1: x2, x5, x10 (Anchor Facts)
// Session 1, Grade 3/4 Numeracy, Variant 0
// DR: Addition Algorithm (Vertical, 4- & 5-digits, with/without renaming)
// Fluency: Subtraction Algorithm (Vertical, 3- & 4-digits, with/without renaming)

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

const T = createTheme("numeracy", "grade34", 0);
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
const FOOTER = "Multiplication Facts | Session 1 of 4 | Grade 3/4 Numeracy";
const OUT_DIR = "output/MF_Session1_Anchor_Facts";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Independent practice - x2, x5, x10 multiplication facts.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with worked answers.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Pre-filled skip counting support for x2, x5, x10.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Self-contained investigation into square numbers.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ───────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Today we start our multiplication facts unit
- We are building our times tables toolkit, starting with the facts most of you can already skip count
- By the end of this unit you will have fast recall of your 2s, 3s, 4s, 5s, 6s, 8s, 9s, and 10s

DO:
- Display title slide as students settle
- Have whiteboards and markers ready for distribution

TEACHER NOTES:
Session 1 of 4. This lesson anchors on x2, x5, and x10 because skip counting patterns are accessible and build confidence before tackling derived-fact strategies in Sessions 2-4.

WATCH FOR:
- Students who seem anxious about times tables -- frame it positively: "We are building strategies, not testing memory"

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Let's warm up with some addition from last week
- On your whiteboards, solve each problem using the vertical algorithm
- Set out your working neatly -- digits lined up in columns

DO:
- Display the four addition problems
- Allow 90 seconds per problem, circulate to check column alignment
- Watch for renaming errors in the thousands column

TEACHER NOTES:
Daily Review revisits the addition algorithm with 4- and 5-digit numbers. Students should be setting these out vertically and renaming where needed. This is prior learning retrieval, not new content.

WATCH FOR:
- Students who forget to rename (carry) when a column totals 10 or more -- prompt: "What happens when we get a two-digit answer in one column?"
- Students who misalign place value columns -- prompt: "Start from the ones and line up each column"

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers -- tick if you got it right, fix it if you didn't
- Fixing it is the learning
- Ask: Which problem needed renaming? [Problems 2 and 4] Where did the renaming happen? [Various columns]

DO:
- Click to reveal answers
- Students tick correct responses and fix errors
- Scan for common mistakes -- note students who consistently miss renaming

TEACHER NOTES:
Tick-and-fix gives immediate feedback. Students who self-correct are consolidating. Students who cannot identify their error need targeted support.

WATCH FOR:
- Students who got all four correct quickly -- they are ready and fluency is strong
- Students who made the same renaming error on multiple problems -- brief side conference after DR

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time -- subtraction vertical algorithm
- Solve these on your whiteboards. Set out vertically, rename where you need to
- You have 2 minutes. Go.

DO:
- Display the six subtraction problems
- Time 2 minutes. Students work on whiteboards
- After time, cold call students to share answers and working for selected problems

TEACHER NOTES:
Fluency builds automaticity with the subtraction algorithm. This is separate from DR -- DR retrieves addition, fluency drills subtraction. Both are prior learning.

WATCH FOR:
- Students who subtract the smaller digit from the larger regardless of position (e.g., 402 - 158 writing 4-1=3, 5-0=5, 8-2=6) -- the classic subtraction direction error
- Students who rename but forget to reduce the next column

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me
- Read from slide: "We are learning to recall multiplication facts for 2s, 5s, and 10s quickly and accurately"
- Now our success criteria. Read each one together.
- Ask: Which success criterion do you think will be easiest? [Most will say SC1 -- skip counting is familiar]

DO:
- Choral read the LI then each SC
- Brief think-pair-share on which SC feels most achievable

TEACHER NOTES:
x2, x5, x10 are anchor facts because students can build them from skip counting. SC1 is the floor (skip counting), SC2 is the target (connecting skip counting to multiplication), SC3 extends to using facts flexibly.

WATCH FOR:
- Students who confuse multiplication with addition -- clarify early: "Multiplication is groups of, not adding on"

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO_X2 = `SAY:
- Watch me work through the x2 facts
- I know my 2s because x2 means double. 2 groups of something is the same as that number plus itself
- So 2 x 7... I think: double 7. That is 7 + 7 = 14
- Let me check: can I skip count to confirm? 2, 4, 6, 8, 10, 12, 14. Yes -- the 7th number in my count is 14

DO:
- Point to the array visual on screen as you model
- Write 2 x 7 = 14 clearly
- Trace the skip counting pattern on the number line or array

TEACHER NOTES:
The x2 strategy is doubling. Connecting doubling to skip counting gives students two access paths. The array makes the groups-of structure visible.

WATCH FOR:
- Students who mouth-count from 1 instead of skip counting by 2s -- they may need the number line scaffold

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_X5_X10 = `SAY:
- Now the 5s. Skip count with me: 5, 10, 15, 20, 25, 30...
- Notice the pattern? Every x5 answer ends in 0 or 5. Always.
- So 5 x 8... I skip count by 5s eight times: 5, 10, 15, 20, 25, 30, 35, 40
- Now the 10s -- these are the easiest. 10 x anything just puts a zero on the end
- 10 x 6 = 60. 10 x 9 = 90. The pattern is always: the number with a zero after it

DO:
- Use fingers to track skip counts visually
- Write each fact clearly, circling the pattern (ending in 0 or 5 for x5, appended zero for x10)

TEACHER NOTES:
x5 and x10 have highly visible patterns. Naming the patterns explicitly (ends in 0 or 5; append zero) converts skip counting into a recognition strategy that is faster than counting every time.

MISCONCEPTIONS:
- Misconception: 10 x 6 = 106 (students append the zero to the wrong side or treat it as concatenation)
  Why: Confusion between place value and digit joining
  Impact: Breaks place value understanding and produces errors in all x10 facts
  Quick correction: "10 x 6 means 10 groups of 6. Count by 10s six times: 10, 20, 30, 40, 50, 60. Not one-hundred-and-six."

WATCH FOR:
- Students who can skip count by 5s but lose track of how many counts they have done -- suggest finger tracking
- Readiness signal: students calling out x10 answers instantly

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. On your whiteboards, write the answer to 2 x 9
- 3, 2, 1, show me

DO:
- Students write on whiteboards
- Scan for 18

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "On your whiteboard, write the answer to 2 x 9. You have 10 seconds. Show me on my signal."
- Scan for: 18. Students who write 18 quickly have automated doubling.
PROCEED: If 80%+ show 18, move to We Do.
PIVOT: Most likely misconception -- students write 11 (added 2+9 instead of multiplying). Reteach: "2 x 9 means 2 groups of 9. Show me 2 groups of 9 on your fingers. How many altogether?"

TEACHER NOTES:
First CFU checks whether students connect x2 to doubling. If most get it, the anchor strategy is landing.

WATCH FOR:
- Students who write 11 (adding, not multiplying) -- this is the critical misconception to catch early
- Students who count on fingers from 1 -- they need the doubling shortcut reinforced

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn to try with a partner
- On your whiteboards, solve these three problems
- For each one, write the strategy you used: doubling, skip counting, or pattern
- First problem: 5 x 7. Second: 10 x 4. Third: 2 x 8.

DO:
- Display the three problems
- Partners work together on whiteboards
- Circulate -- listen for strategy talk, not just answers
- Allow 60 seconds, then cold call for each answer and strategy

TEACHER NOTES:
We Do checks all three anchor strategies together. Requiring students to name the strategy (not just give the answer) builds metacognition about which tool to use.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Skip count by 2s, 5s, and 10s to 50 using a number line with every multiple marked. Circle the answers for 5 x 7, 10 x 4, and 2 x 8 on the number lines.
- Extra Notes: The enabling scaffold PDF provides pre-drawn number lines with multiples highlighted.
EXTENDING PROMPT:
- Task: Find all the numbers between 1 and 50 that appear in BOTH the x2 and x5 skip counting patterns. What do you notice about these numbers? Can you explain why?
- Extra Notes: Students should discover that common multiples of 2 and 5 are multiples of 10.

WATCH FOR:
- Students who get the answer but cannot name a strategy -- prompt: "How did you figure that out?"
- Readiness signal: partners finishing all three within 30 seconds with correct strategies named

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check. 5 x 7? [35] What strategy? [Skip counting by 5s, or 5 x 7 ends in 5 because 7 is odd]
- 10 x 4? [40] Strategy? [Put a zero on the end of 4]
- 2 x 8? [16] Strategy? [Double 8]
- Great work. You are connecting the strategy to the answer, not just guessing

DO:
- Reveal answers one at a time
- Cold call different students for each answer AND strategy
- Celebrate strategy naming, not just correct answers

TEACHER NOTES:
Revealing answers confirms correctness. The strategy discussion is the higher-value part -- it builds the metacognitive toolkit students need for derived facts in Sessions 2-4.

WATCH FOR:
- Students who got wrong answers but used a valid strategy -- affirm the thinking, correct the execution
- Students who got right answers but said "I just knew it" -- prompt: "What would you do if you forgot? Which strategy would help?"

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. I am going to show you a multiplication fact. You show me which strategy you would use.
- Hold up 1 finger for doubling, 2 fingers for skip counting by 5s, 3 fingers for put-a-zero
- Ready? The fact is: 5 x 12

DO:
- Display 5 x 12
- Students hold up fingers
- Scan: most should show 2 fingers (skip counting by 5s)
- Follow up: "Why not doubling?" [Because it is x5, not x2]

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "1 finger = doubling, 2 fingers = skip count by 5s, 3 fingers = put-a-zero. Which strategy for 5 x 12? Show me."
- Scan for: 2 fingers (skip counting by 5s). Accept students who say "I know 5 x 12 = 60 because 5 x 10 = 50 plus 5 x 2 = 10" -- that is a derived fact strategy, which is advanced.
PROCEED: If 80%+ select the correct strategy, move to You Do.
PIVOT: Most likely misconception -- students hold up 3 fingers (confusing x5 with x10 because both involve 5 and 0). Reteach: "x10 means TEN groups. x5 means FIVE groups. 5 x 12 is five groups of 12, not ten groups of 12. Skip count by 5s: 5, 10, 15... that is x5."

TEACHER NOTES:
This hinge question tests strategy selection, not calculation. It reveals whether students can match the table (x2, x5, x10) to the correct shortcut. Strategy selection is the threshold skill for this lesson.

WATCH FOR:
- Students who cannot decide -- they may not yet see the connection between the multiplier and the strategy
- Readiness signal: fast, confident finger raises with correct strategy

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice time
- Open your worksheet. You have 12 problems covering x2, x5, and x10
- First: read each problem. Next: choose your strategy. Then: write the answer and name the strategy
- If you finish early, try the challenge problems at the bottom

DO:
- Distribute Session 1 Worksheet
- Students work independently for 8-10 minutes
- Circulate: prioritise students who struggled during We Do
- For enabling students, distribute the Session 1 Enabling Scaffold

TEACHER NOTES:
You Do uses different numbers from the We Do. Problems progress from single-digit multipliers to two-digit (e.g., 2 x 12, 5 x 11, 10 x 15) to build transfer. The strategy column is critical -- it reinforces metacognition.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Complete the enabling scaffold which provides pre-drawn skip counting number lines for x2, x5, and x10. Students circle the answer on the number line, then write the multiplication fact.
- Extra Notes: Distribute the Session 1 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Investigate square numbers -- numbers you get when you multiply a number by itself (e.g., 2x2=4, 3x3=9). Find all square numbers up to 100 and look for patterns.
- Extra Notes: Distribute the Session 1 Extension PDF. This is a self-contained investigation.

WATCH FOR:
- Students who skip the strategy column -- redirect: "The strategy is just as important as the answer"
- Students who are fast and accurate on all 12 -- direct them to the extension task
- Readiness signal: 80%+ of students completing 8+ problems correctly within 8 minutes

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket time. Four quick questions in your workbook
- Work on your own, no help from partners
- You have 3 minutes

DO:
- Display the exit ticket questions
- Students work silently in workbooks
- Collect or scan responses to sort into three piles: secure, developing, beginning

TEACHER NOTES:
Exit ticket assesses SC1 and SC2 directly. Q1-Q2 test recall (SC1 floor). Q3-Q4 test strategy connection (SC2 core target). Sort results to plan enabling support for Session 2.

WATCH FOR:
- Students who finish in under a minute with all correct -- they have automated these anchor facts
- Students who cannot answer Q1 or Q2 -- they need skip counting consolidation before Session 2

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Let's check our success criteria. Look at the slide.
- SC1: I can skip count by 2s, 5s, and 10s to find multiplication answers. Thumbs up, sideways, or down.
- SC2: I can explain why doubling works for x2 and why x10 puts a zero on the end. Thumbs?
- SC3: I can choose the best strategy for a given multiplication fact. Thumbs?
- Turn and talk: Which strategy was most useful for you today? Tell your partner.

DO:
- Display success criteria on screen
- Run thumbs check for each SC -- note the spread
- Allow 30 seconds for Turn and Talk
- Cold call 2-3 students to share

TEACHER NOTES:
Closing brings the lesson full circle. Self-assessment against SC builds metacognition. The Turn and Talk prompt targets strategy awareness, which is the foundation for derived-fact strategies in Sessions 2-4.

WATCH FOR:
- Students who show thumbs down on SC1 -- they need additional skip counting fluency work
- Students who show thumbs up on all three -- they are ready for the derived-fact strategies in Session 2

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- These are the resources for today's lesson
- The worksheet and answer key are linked here if you need to reprint

DO:
- Point out each resource and its purpose
- Remind students: enabling scaffold is for students who need skip counting support
- Extension is for students who finished early and want a challenge

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
  titleSlide(pres, "Multiplication Facts", "Session 1: Anchor Facts — x2, x5, x10", "Grade 3/4 Numeracy | Session 1 of 4", NOTES_TITLE);

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
        "1.  2 456 + 3 321 =",
        "2.  4 783 + 2 659 =",
        "3.  12 405 + 6 382 =",
        "4.  34 567 + 18 495 =",
      ];
      s.addText(problems.map((p, i) => ({
        text: p,
        options: { fontSize: 15, color: C.CHARCOAL, breakLine: i < problems.length - 1, paraSpaceAfter: 12 },
      })), {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right instruction card
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboards:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "Solve each problem using the vertical algorithm.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Remember:", options: { fontSize: 14, bold: true, color: C.ALERT, breakLine: true } },
        { text: "Line up your place value columns", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Start from the ones column", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
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
      addTextOnShape(slide, "Answers:  1) 5 777    2) 7 442    3) 18 787    4) 53 062", {
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
    "1.  567 - 234 =",
    "2.  842 - 379 =",
    "3.  1 205 - 463 =",
    "4.  3 000 - 1 847 =",
    "5.  5 432 - 2 678 =",
    "6.  7 001 - 3 456 =",
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
    { text: "Check: does your answer make sense?", options: { fontSize: 13, color: C.CHARCOAL } },
  ], {
    x: 5.45, y: CONTENT_TOP + 0.12, w: 3.8, h: 1.36,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    ["We are learning to recall multiplication facts for 2s, 5s, and 10s quickly and accurately"],
    [
      "I can skip count by 2s, 5s, and 10s to find multiplication answers",
      "I can explain why doubling works for x2 and why x10 puts a zero on the end",
      "I can choose the best strategy (doubling, skip counting, or pattern) for a given fact",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — x2 facts
  workedExSlide(pres, 2, "I Do", "x2 Facts: The Doubling Strategy",
    [
      "x2 means 2 groups of a number",
      "2 groups of something = double it",
      "Example: 2 x 7",
      "  Think: double 7 = 7 + 7 = 14",
      "  Check by skip counting: 2, 4, 6, 8, 10, 12, 14 (7th number)",
      "So 2 x 7 = 14",
    ],
    NOTES_IDO_X2, FOOTER,
    (slide, lg) => {
      // Array visual: 2 rows of 7 dots
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.8, { strip: C.PRIMARY });
      slide.addText("2 x 7 = 14", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.35,
        fontSize: 20, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      // Draw 2 rows of 7 circles
      const dotR = 0.18;
      const dotGap = 0.38;
      const startX = lg.rightX + 0.35;
      const startY = lg.panelTopPadded + 0.65;
      for (let row = 0; row < 2; row++) {
        for (let col = 0; col < 7; col++) {
          slide.addShape("roundRect", {
            x: startX + col * dotGap, y: startY + row * dotGap,
            w: dotR, h: dotR, rectRadius: dotR / 2,
            fill: { color: row === 0 ? C.PRIMARY : C.SECONDARY },
          });
        }
      }
      slide.addText("Row 1: 7", {
        x: lg.rightX + 0.15, y: startY + 0.85, w: 1.8, h: 0.25,
        fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });
      slide.addText("Row 2: 7", {
        x: lg.rightX + 0.15, y: startY + 1.1, w: 1.8, h: 0.25,
        fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });
      slide.addText("Total: 7 + 7 = 14", {
        x: lg.rightX + 0.15, y: startY + 1.4, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
      });
    }
  );

  // Slide 7: I Do — x5 and x10 facts
  workedExSlide(pres, 2, "I Do", "x5 and x10: Patterns You Can See",
    [
      "x5 facts: skip count by 5s",
      "  Pattern: answers always end in 0 or 5",
      "  Example: 5 x 8 = 40 (count: 5, 10, 15, 20, 25, 30, 35, 40)",
      "",
      "x10 facts: put a zero on the end",
      "  10 x 6 = 60 (6 with a zero)",
      "  10 x 9 = 90 (9 with a zero)",
      "  Why? 10 groups means we move up one place value",
    ],
    NOTES_IDO_X5_X10, FOOTER,
    (slide, lg) => {
      // x5 pattern display
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.3, { strip: C.SECONDARY });
      slide.addText("x5 Pattern", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.06, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
      });
      slide.addText("5, 10, 15, 20, 25, 30, 35, 40, 45, 50", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.38, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, align: "center",
      });
      slide.addText("Always ends in 0 or 5", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.72, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0, align: "center",
      });

      // x10 pattern display
      addCard(slide, lg.rightX, lg.panelTopPadded + 1.5, lg.rightW, 1.3, { strip: C.PRIMARY });
      slide.addText("x10 Pattern", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.56, w: lg.rightW - 0.4, h: 0.28,
        fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
      });
      slide.addText("10, 20, 30, 40, 50, 60, 70, 80, 90, 100", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.88, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, align: "center",
      });
      slide.addText("Just put a zero on the end!", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.22, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0, align: "center",
      });
    }
  );

  // Slide 8-9: CFU 1 (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check", "Show Me Boards",
      "On your whiteboard:\n\n2 x 9 = ?",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Answer: 2 x 9 = 18  (Double 9: 9 + 9 = 18)", {
        x: 1.5, y: 4.0, w: 7, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10-11: We Do (withReveal)
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Your Turn — With a Partner",
      [
        "Solve each problem on your whiteboard.",
        "Write the STRATEGY you used next to each answer.",
        "",
        "1.   5 x 7 = ___      Strategy: ___",
        "2.   10 x 4 = ___     Strategy: ___",
        "3.   2 x 8 = ___      Strategy: ___",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.6, { strip: C.SECONDARY });
        slide.addText("Strategy Toolkit", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.3,
          fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        const strategies = [
          { label: "Doubling", desc: "x2: double the number", color: C.PRIMARY },
          { label: "Skip Count by 5s", desc: "x5: count 5, 10, 15...", color: C.SECONDARY },
          { label: "Put a Zero", desc: "x10: append a zero", color: C.ACCENT },
        ];
        strategies.forEach((strat, i) => {
          const sy = lg.panelTopPadded + 0.5 + i * 0.65;
          addTextOnShape(slide, strat.label, {
            x: lg.rightX + 0.2, y: sy, w: 1.6, h: 0.35, rectRadius: 0.06,
            fill: { color: strat.color },
          }, { fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true });
          slide.addText(strat.desc, {
            x: lg.rightX + 1.9, y: sy, w: lg.rightW - 2.1, h: 0.35,
            fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "middle",
          });
        });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "1) 35 (skip count)   2) 40 (put a zero)   3) 16 (double 8)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 12-13: CFU 2 Hinge (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Check — Strategy Selection", "Finger Voting",
      "Which strategy would you use for 5 x 12?\n\n1 finger = Doubling\n2 fingers = Skip count by 5s\n3 fingers = Put a zero",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Answer: 2 fingers — Skip count by 5s.  5 x 12 = 60", {
        x: 1.0, y: 4.0, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do
  workedExSlide(pres, 4, "You Do", "You Do — Multiplication Facts",
    [
      "First: Read each problem on your worksheet.",
      "Next: Choose your strategy (doubling, skip count, or pattern).",
      "Then: Write the answer AND name your strategy.",
      "",
      "Work on your own. If you finish early, try the challenge section.",
      "",
      "You have 8-10 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.0, { strip: C.ALERT });
      slide.addText("Remember:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 15, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "x2 = double it", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "x5 = skip count (ends in 0 or 5)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "x10 = put a zero on the end", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.45, w: lg.rightW - 0.5, h: 1.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "2 x 6 = ___",
      "5 x 9 = ___",
      "10 x 7 = ___",
      "Which strategy would you use for 2 x 12? Why?",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 16: Closing
  closingSlide(pres,
    "Which multiplication strategy was most useful for you today — doubling, skip counting, or the zero pattern? Tell your partner why.",
    [
      "I can skip count by 2s, 5s, and 10s to find answers",
      "I can explain why doubling works for x2 and why x10 adds a zero",
      "I can choose the best strategy for a given fact",
    ],
    NOTES_CLOSING);

  // Slide 17: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "MF_Session1_Anchor_Facts.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Generate PDFs ─────────────────────────────────────────────────────────

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Multiplication Facts — x2, x5, x10",
      color: C.NAVY,
      lessonInfo: "Session 1 of 4 | Grade 3/4 Numeracy",
    });
    y = addTipBox(doc, "For each problem: write the answer AND the strategy you used (doubling, skip count by 5s, or put-a-zero).", y, { color: C.TEAL });
    y = addSectionHeading(doc, "x2 Facts (Doubling)", y, { color: C.NAVY });
    y = addProblem(doc, 1, "2 x 4 = ___    Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 2, "2 x 7 = ___    Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 3, "2 x 11 = ___   Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 4, "2 x 12 = ___   Strategy: ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "x5 Facts (Skip Counting)", y, { color: C.NAVY });
    y = addProblem(doc, 5, "5 x 3 = ___    Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 6, "5 x 6 = ___    Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 7, "5 x 9 = ___    Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 8, "5 x 11 = ___   Strategy: ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "x10 Facts (Pattern)", y, { color: C.NAVY });
    y = addProblem(doc, 9, "10 x 3 = ___   Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 10, "10 x 8 = ___  Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 11, "10 x 12 = ___ Strategy: ___", y, { color: C.NAVY });
    y = addProblem(doc, 12, "10 x 15 = ___ Strategy: ___", y, { color: C.NAVY });
    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addProblem(doc, 13, "Which is bigger: 5 x 8 or 2 x 19? Show how you know.", y, { color: C.NAVY, writeLines: [{ label: "" }] });
    addPdfFooter(doc, "Session 1 | Multiplication Facts | Grade 3/4");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Multiplication Facts — x2, x5, x10",
      color: C.NAVY,
      lessonInfo: "Session 1 of 4 | Grade 3/4 Numeracy",
    });
    y = addSectionHeading(doc, "x2 Facts", y, { color: C.NAVY });
    y = addBodyText(doc, "1. 2 x 4 = 8 (doubling: 4 + 4)", y);
    y = addBodyText(doc, "2. 2 x 7 = 14 (doubling: 7 + 7)", y);
    y = addBodyText(doc, "3. 2 x 11 = 22 (doubling: 11 + 11)", y);
    y = addBodyText(doc, "4. 2 x 12 = 24 (doubling: 12 + 12)", y);
    y = addSectionHeading(doc, "x5 Facts", y, { color: C.NAVY });
    y = addBodyText(doc, "5. 5 x 3 = 15 (skip count: 5, 10, 15)", y);
    y = addBodyText(doc, "6. 5 x 6 = 30 (skip count: 5, 10, 15, 20, 25, 30)", y);
    y = addBodyText(doc, "7. 5 x 9 = 45 (skip count or 5 x 10 - 5 = 45)", y);
    y = addBodyText(doc, "8. 5 x 11 = 55 (skip count: 5 x 10 + 5 = 55)", y);
    y = addSectionHeading(doc, "x10 Facts", y, { color: C.NAVY });
    y = addBodyText(doc, "9. 10 x 3 = 30 (put a zero: 3 -> 30)", y);
    y = addBodyText(doc, "10. 10 x 8 = 80 (put a zero: 8 -> 80)", y);
    y = addBodyText(doc, "11. 10 x 12 = 120 (put a zero: 12 -> 120)", y);
    y = addBodyText(doc, "12. 10 x 15 = 150 (put a zero: 15 -> 150)", y);
    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "13. 5 x 8 = 40, 2 x 19 = 38. So 5 x 8 is bigger by 2.", y);
    addPdfFooter(doc, "Session 1 | Multiplication Facts | Grade 3/4");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Skip Counting Support for x2, x5, x10",
      color: C.TEAL,
      lessonInfo: "Session 1 of 4 | Grade 3/4 Numeracy",
    });
    y = addTipBox(doc, "Use the number lines below. Skip count along, then circle the answer.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "x2 — Count by 2s", y, { color: C.NAVY });
    y = addBodyText(doc, "2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24", y);
    y = addProblem(doc, 1, "2 x 4 = ___  (circle the 4th number in the count)", y, { color: C.NAVY });
    y = addProblem(doc, 2, "2 x 7 = ___  (circle the 7th number in the count)", y, { color: C.NAVY });
    y = addProblem(doc, 3, "2 x 11 = ___ (circle the 11th number in the count)", y, { color: C.NAVY });

    y = addSectionHeading(doc, "x5 — Count by 5s", y, { color: C.NAVY });
    y = addBodyText(doc, "5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55", y);
    y = addProblem(doc, 4, "5 x 3 = ___  (circle the 3rd number)", y, { color: C.NAVY });
    y = addProblem(doc, 5, "5 x 6 = ___  (circle the 6th number)", y, { color: C.NAVY });
    y = addProblem(doc, 6, "5 x 9 = ___  (circle the 9th number)", y, { color: C.NAVY });

    y = addSectionHeading(doc, "x10 — Count by 10s", y, { color: C.NAVY });
    y = addBodyText(doc, "10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 140, 150", y);
    y = addProblem(doc, 7, "10 x 3 = ___  (circle the 3rd number)", y, { color: C.NAVY });
    y = addProblem(doc, 8, "10 x 8 = ___  (circle the 8th number)", y, { color: C.NAVY });
    y = addProblem(doc, 9, "10 x 12 = ___ (circle the 12th number)", y, { color: C.NAVY });

    addPdfFooter(doc, "Session 1 | Enabling Scaffold | Grade 3/4");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Extension — Square Numbers Investigation
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Square Numbers Investigation",
      color: C.NAVY,
      lessonInfo: "Session 1 of 4 | Grade 3/4 Numeracy",
    });

    y = addSectionHeading(doc, "What Are Square Numbers?", y, { color: C.NAVY });
    y = addBodyText(doc, "A square number is what you get when you multiply a number by itself.", y);
    y = addBodyText(doc, "For example: 3 x 3 = 9, so 9 is a square number.", y);
    y = addBodyText(doc, "We call them 'square' because you can arrange that many objects into a perfect square shape.", y);

    y = addSectionHeading(doc, "Worked Examples", y, { color: C.NAVY });
    y = addBodyText(doc, "1 x 1 = 1    (1 is a square number)", y);
    y = addBodyText(doc, "2 x 2 = 4    (4 is a square number — picture a 2x2 grid)", y);
    y = addBodyText(doc, "3 x 3 = 9    (9 is a square number — picture a 3x3 grid)", y);
    y = addBodyText(doc, "4 x 4 = 16   (16 is a square number — picture a 4x4 grid)", y);

    y = addSectionHeading(doc, "Your Investigation", y, { color: C.NAVY });
    y = addBodyText(doc, "Find ALL the square numbers up to 100. There are 10 of them.", y);
    y = addBodyText(doc, "Record them in the table below:", y);

    y += 5;
    // Draw a simple table
    const tableX = 60;
    const colW = 120;
    const rowH = 22;
    const headers = ["Number x Itself", "Square Number"];
    headers.forEach((h, i) => {
      doc.rect(tableX + i * colW, y, colW, rowH).fill(C.NAVY ? "#" + C.NAVY : "#1B3F94");
      doc.fillColor("white").fontSize(10).text(h, tableX + i * colW + 5, y + 5, { width: colW - 10, align: "center" });
    });
    y += rowH;
    doc.fillColor("#333333");
    for (let r = 0; r < 10; r++) {
      headers.forEach((_, i) => {
        doc.rect(tableX + i * colW, y, colW, rowH).stroke("#999999");
      });
      if (r < 4) {
        const vals = [[1,1],[2,4],[3,9],[4,16]];
        doc.fontSize(10).text(vals[r][0] + " x " + vals[r][0], tableX + 5, y + 5, { width: colW - 10, align: "center" });
        doc.text(String(vals[r][1]), tableX + colW + 5, y + 5, { width: colW - 10, align: "center" });
      }
      y += rowH;
    }
    y += 10;

    y = addSectionHeading(doc, "Think About It", y, { color: C.NAVY });
    y = addBodyText(doc, "Look at your list of square numbers. What patterns can you see?", y);
    y = addBodyText(doc, "Hint: look at the difference between each square number and the next one.", y);
    y = addWriteLine(doc, "I noticed:", y);
    y = addWriteLine(doc, "", y);

    addPdfFooter(doc, "Session 1 | Extension | Grade 3/4");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 1 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
