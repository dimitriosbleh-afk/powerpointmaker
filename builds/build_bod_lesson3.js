"use strict";

// BODMAS Unit - Session 3: Multiplication and Division (left to right)
// Year 5/6 Numeracy, Week 9 Term 2 (variant 2)
// VC2M6A02

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", weekToVariant(9));
const {
  C, FONT_H, FONT_B,
  SLIDE_W, SLIDE_H, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addBadge, addTitle,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 3;
const FOOTER = "BODMAS | Session 3 of 10 | Year 5/6 Numeracy";
const OUT_DIR = "output/BOD_Session3_Multiplication_Division";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const PRACTICE_RES = makeSessionResource(SESSION, "Practice Sheet", "Multiplication and division left-to-right practice.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference.");
const RESOURCE_ITEMS = [PRACTICE_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Day 3 of BODMAS
- Today is the trickiest letters - the M and the D
- Many students get this wrong because they think M always comes before D
- It doesn't. They're at the same level

DO:
- Display title slide
- Have BODMAS reference visible

TEACHER NOTES:
Session 3 of 10. The M and D rule is a common stumbling block. The order BODMAS suggests M then D, but they are actually at the SAME LEVEL - work left to right. This lesson teaches that explicitly with worked examples that catch the misconception.

WATCH FOR:
- Students who confidently say "M always comes before D" - this is the misconception we are teaching against today

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Brackets and Orders from yesterday
- Solve each one - show steps

DO:
- Display 4 brackets/orders review
- Allow 4 minutes
- Circulate

TEACHER NOTES:
Daily Review keeps brackets and orders fresh. Yesterday was the lesson - today we add the M and D nuance. Spaced practice is essential.

WATCH FOR:
- Students who get the orders wrong - quick reminder: 5 squared = 5 x 5

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Tick or fix
- Anyone get all four? Hand up

DO:
- Reveal answers

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency. Quick mixed multiplication and division facts
- Mini-whiteboards. Show me each answer

DO:
- Display 8 facts
- Quick rhythm - 6 seconds per fact

TEACHER NOTES:
Mixed M and D facts at speed. Students need fluent recall to handle today's lesson without working memory overload. If many students hesitate, schedule extra fact practice.

WATCH FOR:
- Students who count up for facts - they need more recall practice

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Tick or fix
- Note any you missed

DO:
- Reveal answers

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Today's intention: M and D are at the SAME level
- Read SC together
- Ask: True or false - we always do multiplication before division? [FALSE!]

DO:
- Choral read
- Quick true/false to surface the misconception

TEACHER NOTES:
Today's threshold concept: M and D are equal-priority. They work left-to-right. The acronym BODMAS is misleading because M comes before D, but they are at the same level. Same for A and S tomorrow.

WATCH FOR:
- Students who confidently say "TRUE - M before D" - perfect misconception to teach against

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_HOOK = `SAY:
- Look at this equation: 24 / 6 x 2
- On your whiteboards, work it out
- Show me. [Pause - some will write 8, some will write 2]
- Some of you got 8. Some got 2. Why?
- Turn and talk - what did YOU do first?
- [After 90 seconds] Hands up if you got 8. Hands up if you got 2
- The CORRECT answer is 8. Here's why...

DO:
- Display 24 / 6 x 2
- Wait for whiteboards
- Tally answers
- Don't reveal the rule yet - let the disagreement land

TEACHER NOTES:
The hook surfaces the misconception. Students who do x first get 24 / 12 = 2. Students who go left to right get 24/6 = 4, then 4 x 2 = 8. The correct BODMAS answer is 8 because D comes first when reading left to right. This is the most important reveal of the unit.

MISCONCEPTIONS:
- Misconception: BODMAS means M is always done before D
  Why: The acronym puts M first - students assume order
  Impact: Wrong answers on equations like 24 / 6 x 2 (correct: 8, wrong: 2)
  Quick correction: "M and D are at the SAME LEVEL. Look at the equation - which one comes FIRST as you read left to right?"

WATCH FOR:
- Students who got 2 confidently - they are reading BODMAS literally
- Students who got 8 - they may be doing left-to-right correctly OR by luck. Check their working

[Stage 2: Launch | VTLM 2.0: Engagement]`;

const NOTES_RULE = `SAY:
- Here's the rule: M and D are at the SAME level
- When you reach this step, work LEFT TO RIGHT
- Whichever comes FIRST in the equation is the one you do first
- 24 / 6 x 2 - which comes first reading left to right? Division
- So 24 / 6 = 4, then 4 x 2 = 8
- The answer is 8

DO:
- Write the rule on the board: M and D - SAME LEVEL - LEFT TO RIGHT
- Show the equation again with the order numbered

TEACHER NOTES:
The explicit rule reveal. The phrase "same level - left to right" should be repeated like a mantra. Students need to hear it, see it, and use it.

WATCH FOR:
- Students who try to memorise without understanding - check they can apply it to a new equation

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_EXAMPLES = `SAY:
- Two examples to make this stick
- Example 1: 36 / 4 x 3
- Reading left to right - division comes first. 36 / 4 = 9. Then 9 x 3 = 27
- Example 2: 5 x 12 / 3
- Reading left to right - multiplication comes first. 5 x 12 = 60. Then 60 / 3 = 20
- Watch the order - it's whatever comes first as you read

DO:
- Write each example with the BODMAS order numbered
- Highlight which symbol comes first
- Use a colour for left-to-right marking

TEACHER NOTES:
Two contrasting examples - one starts with division, one starts with multiplication. Students see that the rule applies regardless of which symbol comes first. The phrase "reading left to right" is the key.

WATCH FOR:
- Students who keep doing M before D out of habit - call them on it

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. Whiteboards
- 30 / 5 x 2
- Final answer
- 30 seconds

DO:
- Students compute and show
- Scan for "12"
- 12 means correct (30/5=6, 6x2=12)
- 3 means they did 5x2 first (30/10=3) - wrong, BODMAS-literal error

TEACHER NOTES:
This CFU directly tests today's rule. Wrong answers show the BODMAS-literal misconception clearly.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Final answer only
- Show me
- Scan for: 12
PROCEED:
- 80% or more show 12 - move on
PIVOT:
- Most likely misconception: students did x before / because M comes before D in BODMAS
- Reteach: "M and D are at the SAME level. Which one comes FIRST as you read?"
- Re-check with: 40 / 8 x 5 [answer: 25]

WATCH FOR:
- Students showing 3 - classic BODMAS-literal error
- Students showing 12 - rule is landing

[Stage 2: CFU | VTLM 2.0: Active Checking]`;

const NOTES_CFU1_A = `SAY:
- Answer: 12
- 30 divided by 5 is 6
- Then 6 times 2 is 12
- Division came first because it was first reading left to right

DO:
- Reveal full working

[Stage 2: CFU Answer | VTLM 2.0: Active Checking]`;

const NOTES_WEDO_Q = `SAY:
- Together. Three operations now
- 8 x 6 / 4 x 2
- All multiplication and division. Strict left to right
- What's first? 8 x 6 = 48. Equation now: 48 / 4 x 2
- Next? 48 / 4 = 12. Equation now: 12 x 2
- Then: 12 x 2 = 24
- Final: 24

DO:
- Walk through with the class
- Pause at each step
- Choral response - "what's next?"

TEACHER NOTES:
Three-operation chain to reinforce strict left-to-right working. Students must NOT skip ahead - work each operation in turn.

WATCH FOR:
- Students who try to do all the multiplications first then divisions - this is the wrong habit

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Final: 24
- Tick or fix

DO:
- Reveal answer

[Stage 3: We Do Answer | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge question. Whiteboards
- 100 / 5 x 4
- Final answer

DO:
- Students compute and show
- Scan for "80"
- 80 means correct (100/5=20, 20x4=80)
- 5 means they did 5x4=20, then 100/20=5 - BODMAS-literal error
- 25 means they did 100/5=20, then made a calc error

TEACHER NOTES:
Hinge tests the threshold understanding. The 80 vs 5 split tells you exactly who has the rule and who is still BODMAS-literal.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Final answer
- Show me
- Scan for: 80
PROCEED:
- 80% or more show 80 - go to You Do
PIVOT:
- Most likely misconception: students multiplied first because M is before D in BODMAS
- Reteach: "M and D - SAME level. Read LEFT TO RIGHT. The division 100 / 5 comes FIRST"
- Re-check: 60 / 6 x 3 [answer: 30]

WATCH FOR:
- Students showing 5 - classic error, needs intervention
- Students showing 80 confidently - solid

[Stage 3: Hinge CFU | VTLM 2.0: Active Checking]`;

const NOTES_CFU2_A = `SAY:
- Answer: 80
- 100 / 5 = 20 (division came first - reading left to right)
- 20 x 4 = 80
- If you got 5, you did multiplication first - that's the BODMAS trap

DO:
- Reveal full working

[Stage 3: Hinge Answer | VTLM 2.0: Active Checking]`;

const NOTES_YOUDO = `SAY:
- Independent practice
- Show every step
- Underline which operation comes first reading left to right

DO:
- Hand out practice sheet
- Allow 10 minutes
- Pull a small group of students who failed the hinge

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Pull a small group. Use the BODMAS reference. For each equation, mark the symbols 1, 2, 3 as you read left to right
- Extra Notes: Solve in numbered order. Build the habit before adding speed
EXTENDING PROMPT:
- Task: Challenge equation: 144 / 12 / 3 x 2 + 5 (hint: D and D, same level)
- Extra Notes: All same-level operations. Strict left to right. Answer: 13

WATCH FOR:
- Students who skip the left-to-right marking - the marking IS the strategy
- Students who get tangled in three or more operations - chunk into pairs

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Three exit ticket questions
- Show your steps - I want to see what came first

DO:
- Hand out exit ticket
- Collect

TEACHER NOTES:
Q1 simple D-then-M, Q2 M-then-D, Q3 explanation. Tests SC2.

WATCH FOR:
- Students who explain "M before D always" on Q3 - they have NOT got the rule
- Students who say "left to right" on Q3 - excellent

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Read SC together
- Self-check
- Tell your partner: what's the trick to remember about M and D?

DO:
- Closing slide
- Self-check
- Turn and talk

TEACHER NOTES:
Closing reinforces the M and D = same level rule. If many students still confused, plan extra small group time.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- Practice sheet for tonight or for early finishers

DO:
- Hand out practice sheets

[General: Resources | VTLM 2.0: Resource Awareness]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres,
    "M and D - Same Level",
    "Multiplication and Division - work left to right",
    "Session 3 of 10  |  Year 5/6 Numeracy",
    NOTES_TITLE);

  // Slide 2-3: Daily Review
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Brackets and Orders Recap", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const dr = [
        "1.  (10 - 4) x 3",
        "2.  20 - (5 + 4)",
        "3.  3 squared + 6",
        "4.  4 squared - (8 + 2)",
      ];
      s.addText(dr.map((p, i) => ({
        text: p,
        options: { fontSize: 20, color: C.CHARCOAL, breakLine: i < dr.length - 1, paraSpaceAfter: 12 },
      })), {
        x: 0.75, y: CONTENT_TOP + 0.2, w: 4.1, h: SAFE_BOTTOM - CONTENT_TOP - 0.4,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, 2.0, { strip: C.SECONDARY });
      s.addText([
        { text: "Brackets first, then orders", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Show every step", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "4 minutes", options: { bullet: true, fontSize: 14, color: C.ALERT, bold: true } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: 1.7,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) 18    2) 11    3) 15    4) 6", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4-5: Fluency
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Fluency");
      addTitle(s, "Mixed Multiplication and Division Facts", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 9, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const fluency = [
        "8 x 7", "56 / 7", "9 x 6",
        "72 / 8", "12 x 4", "60 / 5",
        "11 x 5", "84 / 12",
      ];
      // 4-col x 2-row grid
      fluency.forEach((q, i) => {
        const col = i % 4;
        const row = Math.floor(i / 4);
        const x = 0.7 + col * 2.16;
        const y = CONTENT_TOP + 0.4 + row * 1.55;
        addTextOnShape(s, q, {
          x, y, w: 1.95, h: 1.15, rectRadius: 0.1,
          fill: { color: STAGE_COLORS["1"] },
        }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_FLUENCY);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "56  |  8  |  54  |  9  |  48  |  12  |  55  |  7", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 6: LI/SC
  liSlide(pres,
    ["I am learning that Multiplication and Division are at the SAME LEVEL"],
    [
      "I can identify which operation comes first reading left to right",
      "I can solve an equation with both x and / using BODMAS",
      "I can explain why we don't always do x before /",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 7: Hook - the 24/6x2 problem
  contentSlide(pres, "Launch", C.PRIMARY, "What does this equal?",
    [
      "Work it out on your whiteboard",
      "Show me when you have an answer",
      "Don't talk yet",
    ],
    NOTES_HOOK, FOOTER,
    (slide, lg) => {
      addTextOnShape(slide, "24 / 6 x 2 = ?", {
        x: lg.rightX, y: lg.panelTopPadded, w: lg.rightW, h: 1.6, rectRadius: 0.12,
        fill: { color: C.PRIMARY },
      }, { fontSize: 40, fontFace: FONT_H, color: C.WHITE, bold: true });

      addTextOnShape(slide, "Some say:\n8", {
        x: lg.rightX, y: lg.panelTopPadded + 1.85, w: 2.0, h: 1.3, rectRadius: 0.1,
        fill: { color: C.SECONDARY },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });

      addTextOnShape(slide, "Some say:\n2", {
        x: lg.rightX + 2.2, y: lg.panelTopPadded + 1.85, w: 2.0, h: 1.3, rectRadius: 0.1,
        fill: { color: C.ACCENT },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });

      slide.addText("Both groups think BODMAS!", {
        x: lg.rightX, y: lg.panelTopPadded + 3.25, w: lg.rightW, h: 0.4,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });
    }
  );

  // Slide 8: The rule
  contentSlide(pres, "I Do", C.PRIMARY, "M and D are at the SAME LEVEL",
    [
      "BODMAS puts M before D - but they are EQUAL",
      "Work LEFT TO RIGHT when you reach this step",
      "Whichever symbol comes FIRST gets done FIRST",
      "",
      "24 / 6 x 2:",
      "  Division comes first (reading L to R)",
      "  24 / 6 = 4, then 4 x 2 = 8",
    ],
    NOTES_RULE, FOOTER,
    (slide, lg) => {
      // Big rule card
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.ALERT });
      slide.addText("THE RULE", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.4,
        fontSize: 16, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0, align: "center",
      });

      slide.addText([
        { text: "M and D = ", options: { fontSize: 22, color: C.CHARCOAL, bold: true, breakLine: false } },
        { text: "SAME LEVEL", options: { fontSize: 22, color: C.ALERT, bold: true, breakLine: true } },
      ], {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.65, w: lg.rightW - 0.4, h: 0.6,
        fontFace: FONT_H, margin: 0, align: "center",
      });

      // Arrow indicator
      slide.addText("--> Read LEFT to RIGHT -->", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.4, w: lg.rightW - 0.4, h: 0.5,
        fontSize: 18, fontFace: FONT_B, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });

      slide.addText("Whichever you see first --> do first", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.0, w: lg.rightW - 0.4, h: 0.4,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });

      addTextOnShape(slide, "24 / 6 x 2 = 8", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.6, w: lg.rightW - 0.4, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 9: I Do - two examples
  workedExSlide(pres, 2, "I Do", "Two More Examples",
    [
      "Example 1: 36 / 4 x 3",
      "  Division first (it's first L to R)",
      "  36 / 4 = 9",
      "  9 x 3 = 27",
      "",
      "Example 2: 5 x 12 / 3",
      "  Multiplication first (first L to R)",
      "  5 x 12 = 60",
      "  60 / 3 = 20",
    ],
    NOTES_IDO_EXAMPLES, FOOTER,
    (slide, lg) => {
      // Two example cards
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.5, { strip: C.PRIMARY });
      slide.addText("Example 1", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
      });
      slide.addText("36 / 4 x 3", {
        x: lg.rightX, y: lg.panelTopPadded + 0.42, w: lg.rightW, h: 0.45,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
      });
      addTextOnShape(slide, "= 27", {
        x: lg.rightX + 1.2, y: lg.panelTopPadded + 0.95, w: 1.8, h: 0.4, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

      addCard(slide, lg.rightX, lg.panelTopPadded + 1.7, lg.rightW, 1.5, { strip: C.SECONDARY });
      slide.addText("Example 2", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.78, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
      });
      slide.addText("5 x 12 / 3", {
        x: lg.rightX, y: lg.panelTopPadded + 2.12, w: lg.rightW, h: 0.45,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", margin: 0,
      });
      addTextOnShape(slide, "= 20", {
        x: lg.rightX + 1.2, y: lg.panelTopPadded + 2.65, w: 1.8, h: 0.4, rectRadius: 0.06,
        fill: { color: C.SECONDARY },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10-11: CFU 1
  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check", "Show Me Boards",
      "30 / 5 x 2\n\nFinal answer.",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Answer: 12   (30 / 5 = 6, then 6 x 2 = 12)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU1_A);
    }
  );

  // Slide 12-13: We Do
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Three Operations",
      [
        "Equation: 8 x 6 / 4 x 2",
        "",
        "All M and D - strict left to right",
        "",
        "Step 1: 8 x 6 = ?",
        "Step 2: ___ / 4 = ?",
        "Step 3: ___ x 2 = ?",
        "",
        "Show every step in your books",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.SECONDARY });
        slide.addText("Strict left to right", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.36,
          fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
        });

        // Sequence visual
        const seq = ["8 x 6", "/ 4", "x 2"];
        seq.forEach((s, i) => {
          const y = lg.panelTopPadded + 0.55 + i * 0.6;
          addTextOnShape(slide, String(i + 1), {
            x: lg.rightX + 0.25, y, w: 0.5, h: 0.45, rectRadius: 0.06,
            fill: { color: C.SECONDARY },
          }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
          slide.addText(s, {
            x: lg.rightX + 0.85, y, w: lg.rightW - 1.05, h: 0.45,
            fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, bold: true, valign: "middle", margin: 0,
          });
        });

        addTextOnShape(slide, "Same level - work in order", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.5, w: lg.rightW - 0.4, h: 0.5, rectRadius: 0.08,
          fill: { color: C.PRIMARY },
        }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "Step 1: 48   |   Step 2: 12   |   Step 3: 24", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 14-15: Hinge CFU
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Question", "Show Me Boards",
      "100 / 5 x 4\n\nFinal answer.",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Answer: 80   (100 / 5 = 20, then 20 x 4 = 80)   Trap: 5 = M before D", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU2_A);
    }
  );

  // Slide 16: You Do
  workedExSlide(pres, 4, "You Do", "Independent Practice",
    [
      "Solve - strict left to right for M and D",
      "Underline which one comes first reading L to R",
      "",
      "1.  40 / 8 x 3",
      "2.  6 x 10 / 4",
      "3.  72 / 9 x 5",
      "4.  3 x 16 / 6",
      "5.  60 / 4 / 3 x 2",
      "",
      "Challenge: 144 / 12 / 3 x 2 + 5",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.ALERT });
      slide.addText("Today's rule", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.36,
        fontSize: 15, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "M and D = same level", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Work LEFT to RIGHT", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Whichever comes first --> do first", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Do NOT always do x before /", options: { bullet: true, fontSize: 13, color: C.ALERT, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.8,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addTextOnShape(slide, "10 minutes", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.45, w: lg.rightW - 0.4, h: 0.4, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "Solve:  20 / 4 x 3   (show steps)",
      "Solve:  6 x 8 / 4   (show steps)",
      "Tell your teacher: WHY don't we always do x before / ?",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 18: Closing
  closingSlide(pres, {
    reflectionPrompt: "Tell your partner: what is the trick to remember about M and D?",
    scItems: [
      "I can identify which operation comes first reading left to right",
      "I can solve an equation with both x and / using BODMAS",
      "I can explain why we don't always do x before /",
    ],
    selfAssessment: { prompt: "Self-check", options: ["Got it", "Getting there", "Need more practice"] },
  }, NOTES_CLOSING);

  // Slide 19: Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "BOD_Session3_Multiplication_Division.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: PRACTICE_RES.name });
    let y = addPdfHeader(doc, PRACTICE_RES.name, {
      subtitle: "Multiplication and Division - Same Level",
      color: C.NAVY,
      lessonInfo: "Session 3 of 10 | Year 5/6 Numeracy",
    });
    y = addTipBox(doc, "M and D are at the same level. Work LEFT TO RIGHT - whichever comes first gets done first.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A - Two operations", y, { color: C.NAVY });
    y = addProblem(doc, 1, "40 / 8 x 3 =                          Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 2, "6 x 10 / 4 =                          Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 3, "72 / 9 x 5 =                          Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 4, "3 x 16 / 6 =                          Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 5, "100 / 4 x 2 =                         Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 6, "8 x 9 / 6 =                           Show every step", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Section B - Three operations", y, { color: C.NAVY });
    y = addProblem(doc, 7, "60 / 4 / 3 x 2 =                      Strict L to R", y, { color: C.NAVY });
    y = addProblem(doc, 8, "5 x 12 / 4 / 3 =                      Strict L to R", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addProblem(doc, 9, "144 / 12 / 3 x 2 + 5 =                All steps shown", y, { color: C.NAVY });
    y += 8;
    y = addBodyText(doc, "Reflect: Why don't we always do x before / ?", y);
    y = addWriteLine(doc, "", y);
    y = addWriteLine(doc, "", y);

    addPdfFooter(doc, "Session 3 | Practice Sheet | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, PRACTICE_RES.fileName));
    console.log("PDF written: " + PRACTICE_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "M and D Practice Answer Key",
      color: C.NAVY,
      lessonInfo: "Session 3 of 10 | Year 5/6 Numeracy",
    });

    y = addSectionHeading(doc, "Section A - Two operations", y, { color: C.NAVY });
    y = addBodyText(doc, "1.  40 / 8 x 3   -->  5 x 3 = 15", y);
    y = addBodyText(doc, "2.  6 x 10 / 4   -->  60 / 4 = 15", y);
    y = addBodyText(doc, "3.  72 / 9 x 5   -->  8 x 5 = 40", y);
    y = addBodyText(doc, "4.  3 x 16 / 6   -->  48 / 6 = 8", y);
    y = addBodyText(doc, "5.  100 / 4 x 2  -->  25 x 2 = 50", y);
    y = addBodyText(doc, "6.  8 x 9 / 6    -->  72 / 6 = 12", y);

    y = addSectionHeading(doc, "Section B - Three operations", y, { color: C.NAVY });
    y = addBodyText(doc, "7.  60 / 4 / 3 x 2   -->  15 / 3 = 5  -->  5 x 2 = 10", y);
    y = addBodyText(doc, "8.  5 x 12 / 4 / 3   -->  60 / 4 = 15  -->  15 / 3 = 5", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "9.  144 / 12 / 3 x 2 + 5", y);
    y = addBodyText(doc, "    144 / 12 = 12   -->   12 / 3 = 4   -->   4 x 2 = 8   -->   8 + 5 = 13", y);

    addPdfFooter(doc, "Session 3 | Answer Key | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Session 3 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
