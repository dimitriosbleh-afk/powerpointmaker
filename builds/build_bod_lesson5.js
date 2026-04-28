"use strict";

// BODMAS Unit - Session 5: Full BODMAS - Putting It All Together
// Year 5/6 Numeracy, Week 9 Term 2 (variant 2)

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

const SESSION = 5;
const FOOTER = "BODMAS | Session 5 of 10 | Year 5/6 Numeracy";
const OUT_DIR = "output/BOD_Session5_Full_BODMAS";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const PRACTICE_RES = makeSessionResource(SESSION, "Mixed Practice Sheet", "Full BODMAS practice across all four levels.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference.");
const RESOURCE_ITEMS = [PRACTICE_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Day 5 of BODMAS - end of week one
- Today we put it all together
- Full BODMAS equations, error analysis, and the viral problems

DO:
- Display title slide
- Have BODMAS reference visible

TEACHER NOTES:
Session 5 of 10. Today is consolidation. No new content - we apply everything from sessions 1-4 to longer, mixed equations. Error analysis sharpens the same understanding from a different angle.

WATCH FOR:
- Students who still apply BODMAS literally - a final chance to fix before next week's calendar project

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Mixed BODMAS from this week
- Show every step

DO:
- Display 4 mixed equations
- Allow 5 minutes
- Circulate

TEACHER NOTES:
Daily Review pulls from all four lessons this week. Use the data to spot which level is still wobbly.

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Tick or fix
- Hands up if you got all four

DO:
- Reveal answers

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency. Mixed +, -, x, / facts at speed
- Quick whiteboard responses

DO:
- Display 8 facts
- Brisk pace

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Tick or fix

DO:
- Reveal answers

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Today is consolidation
- Read SC together
- Ask: which BODMAS step have you found hardest this week?

DO:
- Choral read
- Quick turn-and-talk

TEACHER NOTES:
Today's success criteria are about confident, fluent application across the full BODMAS rule. SC1 is "I can apply BODMAS in order". SC2 is "I can solve full equations". SC3 is "I can find the mistake".

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VIRAL = `SAY:
- Remember our first viral problem? 6 + 2 x 3 = 12
- Here is another famous one. Some people argue about it online
- Work it out using BODMAS

DO:
- Display the equation
- Wait for whiteboards
- Tally answers

TEACHER NOTES:
The viral problem is the engagement hook AND the consolidation check. By session 5, students should solve this confidently with BODMAS. Use it to celebrate growth.

WATCH FOR:
- Students who solve it correctly and quickly - they have BODMAS
- Students who still hesitate - they need more practice

[Stage 2: Launch | VTLM 2.0: Engagement]`;

const NOTES_FULL_EXAMPLE = `SAY:
- Full BODMAS worked example
- 2 squared + (15 - 5) / 2 - 3
- Watch every step
- B: (15 - 5) = 10. Now: 2 squared + 10 / 2 - 3
- O: 2 squared = 4. Now: 4 + 10 / 2 - 3
- D/M: 10 / 2 = 5. Now: 4 + 5 - 3
- A/S left to right: 4 + 5 = 9, 9 - 3 = 6

DO:
- Write each step on the board
- Cross out and replace
- Use a different colour for each level

TEACHER NOTES:
Full worked example showing all four BODMAS levels firing. The pattern: solve B, then O, then sweep through D/M left to right, then A/S left to right.

WATCH FOR:
- Students who skip the BODMAS check at the start - prompt: "Which letters appear in this equation?"

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_ERROR_INTRO = `SAY:
- Now we change the game
- I'm going to show you three equations - each one has been SOLVED by a student
- Your job: spot the error
- This is what mathematicians do - check each other's work

DO:
- Display the error analysis slide
- Read each one out loud
- Wait before revealing the errors

TEACHER NOTES:
Error analysis is a high-yield strategy. Students must hold the rule and apply it diagnostically. This is harder than just solving - it requires deeper understanding.

WATCH FOR:
- Students who say "I don't see anything wrong" - they may have the same misconception
- Students who quickly find the error - their understanding is solid

[Stage 2: I Do | VTLM 2.0: Error Analysis]`;

const NOTES_CFU1 = `SAY:
- Quick check. On your whiteboards
- Look at this equation: 12 - 8 / 4 + 3
- A student wrote: 12 - 2 + 3 = 13
- Final answer: 13. Is the student RIGHT or WRONG? Just write R or W

DO:
- Students write R or W
- Show me

TEACHER NOTES:
The student worked correctly here. 8/4 = 2 (D first), then 12-2+3 = 13. This CFU tests whether students can recognise CORRECT working, not just spot errors.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- R or W
- Show me
- Scan for: R
PROCEED:
- 80% or more show R - move on
PIVOT:
- Most likely misconception: students mark it W because they expect to find an error
- Reteach: "Walk through it yourself. Did the student do D before A/S? Yes. Did they go left to right at the end? Yes"

WATCH FOR:
- Students showing W - check their working and reasoning

[Stage 2: CFU | VTLM 2.0: Active Checking]`;

const NOTES_CFU1_A = `SAY:
- The student is RIGHT
- 8 / 4 = 2 (division first - well done)
- 12 - 2 + 3 = 13 (left to right - well done)

DO:
- Reveal R

[Stage 2: CFU Answer | VTLM 2.0: Active Checking]`;

const NOTES_WEDO_Q = `SAY:
- Together. Find the error
- 20 - 6 / 2 + 4 = 11
- A student claims this. Walk through with me - is it right?
- B? No. O? No. D/M? Yes - 6/2 = 2  WAIT - is 6/2 = 2?
- No - 6/2 = 3. The student made an arithmetic error
- Correct: 20 - 3 + 4 = 21

DO:
- Walk through with class
- Pause for hands - who agrees so far?
- Reveal the arithmetic mistake

TEACHER NOTES:
Error analysis - this one is an arithmetic mistake (6/2 = 2 instead of 3), not a BODMAS rule error. Important: students need to spot BOTH kinds of errors. Let them debate before you reveal.

WATCH FOR:
- Students who only check the BODMAS order, not the arithmetic - both matter
- Students who spot the 6/2 error quickly - their fact recall is strong

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Error: 6 / 2 should be 3, not 2
- Correct answer: 20 - 3 + 4 = 21

DO:
- Reveal the correction

[Stage 3: We Do Answer | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge question. Spot the error or write CORRECT
- 4 squared - (8 - 5) x 2 = 10
- Walk through it. Is it right or wrong? If wrong, what's the correct answer?

DO:
- Allow 60 seconds
- Whiteboards: CORRECT or actual answer

TEACHER NOTES:
Hinge tests full BODMAS application AND error detection. Correct working: 4 squared = 16, (8-5) = 3, 3x2 = 6, 16-6 = 10. The given answer 10 IS correct. This tests whether students can identify correct work.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Write CORRECT or the actual answer
- Show me
- Scan for: CORRECT
PROCEED:
- 80% or more show CORRECT - move on
PIVOT:
- Most likely misconception: students assume there must be an error
- Reteach: "Walk through every step. If every step is right, the answer is right"
- Re-check: 3 squared + (10 - 7) - 2 = 10 [correct]

WATCH FOR:
- Students who write any other number - they made a working error
- Students who write CORRECT confidently - they have BODMAS

[Stage 3: Hinge CFU | VTLM 2.0: Active Checking]`;

const NOTES_CFU2_A = `SAY:
- The answer IS correct
- 4 squared = 16
- (8 - 5) = 3
- 3 x 2 = 6
- 16 - 6 = 10 - YES, correct

DO:
- Reveal full working

[Stage 3: Hinge Answer | VTLM 2.0: Active Checking]`;

const NOTES_YOUDO = `SAY:
- Independent practice
- Mixed full-BODMAS equations and error analysis
- Show every step
- For error analysis: write CORRECT or write the actual answer

DO:
- Hand out practice sheet
- Allow 12 minutes
- Pull small group for any students still struggling with the rule

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Pull small group. For each equation, write the four BODMAS levels down the side and TICK which appear. Then solve in order
- Extra Notes: This breaks the long equations into manageable level-by-level chunks
EXTENDING PROMPT:
- Task: Challenge: 2 cubed + (12 - 4) x 3 - 6 / 2
- Extra Notes: Two brackets, an order, and all four levels. Answer: 29

WATCH FOR:
- Students who get error analysis wrong but solve the regular ones right - they know the rule but don't apply it diagnostically
- Students who confidently move through both - ready for the calendar project

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Three exit ticket questions
- Show your steps for the first two
- Answer the third in a sentence

DO:
- Hand out exit ticket
- Collect

TEACHER NOTES:
Q1 full BODMAS calculation. Q2 error analysis. Q3 explanation. Tests SC2 and SC3.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- End of week one of BODMAS
- Read SC together
- Self-check
- Tell your partner: how confident do you feel about BODMAS now?

DO:
- Closing slide
- Self-check
- Turn and talk

TEACHER NOTES:
End of week 1. Use the self-check to plan for the calendar project starting next session. If many students self-assess as "Need more practice", a quick 10-minute reteach at the start of session 6 may be needed.

[General: Closing | VTLM 2.0: Reflection]`;

const NOTES_RESOURCES = `SAY:
- Practice sheet for any extra at home or at school

DO:
- Hand out practice sheets

[General: Resources | VTLM 2.0: Resource Awareness]`;

// --- Build -------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres,
    "Full BODMAS",
    "Putting it all together",
    "Session 5 of 10  |  Year 5/6 Numeracy",
    NOTES_TITLE);

  // Slide 2-3: Daily Review
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "BODMAS Recap - This Week", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const dr = [
        "1.  20 - 6 + 4",
        "2.  36 / 6 x 2",
        "3.  3 squared + (5 - 1)",
        "4.  18 - 4 x 3 + 2",
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
        { text: "Use BODMAS - all four levels", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Show every step", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "Same level = left to right", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "5 minutes", options: { bullet: true, fontSize: 14, color: C.ALERT, bold: true } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: 1.7,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) 18    2) 12    3) 13    4) 8", {
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
      addTitle(s, "Mixed Operations Speed Round", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 9, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const fluency = [
        "8 x 9", "144 / 12", "67 + 38",
        "100 - 47", "9 x 12", "96 / 8",
        "245 + 156", "500 - 268",
      ];
      fluency.forEach((q, i) => {
        const col = i % 4;
        const row = Math.floor(i / 4);
        const x = 0.7 + col * 2.16;
        const y = CONTENT_TOP + 0.4 + row * 1.55;
        addTextOnShape(s, q, {
          x, y, w: 1.95, h: 1.15, rectRadius: 0.1,
          fill: { color: STAGE_COLORS["1"] },
        }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_FLUENCY);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "72  |  12  |  105  |  53  |  108  |  12  |  401  |  232", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 6: LI/SC
  liSlide(pres,
    ["I am consolidating BODMAS across all four levels"],
    [
      "I can apply BODMAS in the right order",
      "I can solve a full BODMAS equation correctly",
      "I can find and explain the mistake in a wrong answer",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 7: Viral problem launch
  contentSlide(pres, "Launch", C.PRIMARY, "Another Famous Equation",
    [
      "This one starts arguments online",
      "Use BODMAS to solve it confidently",
      "Show every step",
    ],
    NOTES_VIRAL, FOOTER,
    (slide, lg) => {
      addTextOnShape(slide, "8 / 2 x (2 + 2)", {
        x: lg.rightX, y: lg.panelTopPadded + 0.4, w: lg.rightW, h: 1.6, rectRadius: 0.12,
        fill: { color: C.PRIMARY },
      }, { fontSize: 36, fontFace: FONT_H, color: C.WHITE, bold: true });

      slide.addText("Use BODMAS - what comes first?", {
        x: lg.rightX, y: lg.panelTopPadded + 2.2, w: lg.rightW, h: 0.4,
        fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });

      addTextOnShape(slide, "Show your steps!", {
        x: lg.rightX + 0.4, y: lg.panelTopPadded + 2.7, w: lg.rightW - 0.8, h: 0.45, rectRadius: 0.08,
        fill: { color: C.SECONDARY },
      }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 8: Full worked example
  workedExSlide(pres, 2, "I Do", "All Four Levels",
    [
      "Equation: 2 squared + (15 - 5) / 2 - 3",
      "",
      "B (brackets):",
      "  (15 - 5) = 10  -->  2 squared + 10 / 2 - 3",
      "",
      "O (orders):",
      "  2 squared = 4  -->  4 + 10 / 2 - 3",
      "",
      "D/M (left to right):",
      "  10 / 2 = 5  -->  4 + 5 - 3",
      "",
      "A/S (left to right):",
      "  4 + 5 = 9, 9 - 3 = 6",
      "",
      "Final: 6",
    ],
    NOTES_FULL_EXAMPLE, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.85, { strip: C.PRIMARY });
      slide.addText("BODMAS Sweep", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      const sweeps = [
        { L: "B", text: "(15-5) = 10" },
        { L: "O", text: "2 sq = 4" },
        { L: "D/M", text: "10 / 2 = 5" },
        { L: "A/S", text: "4 + 5 - 3 = 6" },
      ];
      sweeps.forEach((sw, i) => {
        const y = lg.panelTopPadded + 0.45 + i * 0.55;
        addTextOnShape(slide, sw.L, {
          x: lg.rightX + 0.25, y, w: 0.7, h: 0.42, rectRadius: 0.06,
          fill: { color: C.SUCCESS },
        }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
        slide.addText(sw.text, {
          x: lg.rightX + 1.05, y, w: lg.rightW - 1.25, h: 0.42,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });

      addTextOnShape(slide, "ANSWER: 6", {
        x: lg.rightX, y: lg.panelTopPadded + 2.95, w: lg.rightW, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 9: Error analysis intro
  contentSlide(pres, "I Do", C.PRIMARY, "Find the Mistake",
    [
      "Mathematicians check each other's work",
      "Today you become the checker",
      "Walk through each step carefully",
      "Look for: BODMAS errors AND arithmetic errors",
    ],
    NOTES_ERROR_INTRO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.ALERT });
      slide.addText("Two kinds of error", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.36,
        fontSize: 16, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0, align: "center",
      });

      // Error types
      addTextOnShape(slide, "BODMAS error", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.6, w: lg.rightW - 0.4, h: 0.45, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
      slide.addText("Wrong order: did + before x", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.1, w: lg.rightW - 0.4, h: 0.35,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });

      addTextOnShape(slide, "Arithmetic error", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.7, w: lg.rightW - 0.4, h: 0.45, rectRadius: 0.06,
        fill: { color: C.SECONDARY },
      }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
      slide.addText("Wrong calculation: 6/2 written as 2", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.2, w: lg.rightW - 0.4, h: 0.35,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });

      addTextOnShape(slide, "Check both!", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.85, w: lg.rightW - 0.4, h: 0.45, rectRadius: 0.06,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 10-11: CFU 1 - is the student right?
  withReveal(
    () => cfuSlide(pres, "CFU", "Right or Wrong?", "Show Me Boards",
      "A student writes:\n12 - 8 / 4 + 3 = 13\n\nWrite R (right) or W (wrong)",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "RIGHT (8/4=2 first, then 12-2+3 = 13 left to right)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU1_A);
    }
  );

  // Slide 12-13: We Do - find the error
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Find the Error Together",
      [
        "A student claims:",
        "20 - 6 / 2 + 4 = 11",
        "",
        "Walk through with me:",
        "B? No. O? No.",
        "D/M: 6 / 2 = ? ... is it 2?",
        "",
        "What is 6 / 2?",
        "",
        "If the student wrote 2, that's an error",
        "Correct: 6 / 2 = 3",
        "20 - 3 + 4 = 21",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.SECONDARY });
        slide.addText("Detective work", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.36,
          fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
        });
        slide.addText([
          { text: "Check the BODMAS order", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Check each calculation", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Find the FIRST mistake", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
          { text: "Solve correctly", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
        ], {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.6,
          fontFace: FONT_B, margin: 0, valign: "top",
        });

        addTextOnShape(slide, "Two error types: BODMAS or arithmetic", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.6, w: lg.rightW - 0.4, h: 0.5, rectRadius: 0.08,
          fill: { color: C.PRIMARY },
        }, { fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "Error: 6/2 should be 3, not 2  |  Correct answer: 21", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 14-15: Hinge CFU
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Question", "Show Me Boards",
      "4 squared - (8 - 5) x 2 = 10\n\nWrite CORRECT, or write the actual answer.",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "CORRECT  (4 sq=16, (8-5)=3, 3x2=6, 16-6 = 10)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU2_A);
    }
  );

  // Slide 16: You Do
  workedExSlide(pres, 4, "You Do", "Mixed Practice",
    [
      "Section A - Solve full BODMAS",
      "1.  3 squared + (10 - 4) - 2",
      "2.  20 - 12 / 4 + 3 x 2",
      "3.  (5 + 3) x 2 - 2 squared",
      "",
      "Section B - Find the error or write CORRECT",
      "4.  6 + 4 x 2 = 20",
      "5.  18 / 6 + 7 = 10",
      "",
      "Challenge: 2 cubed + (12 - 4) x 3 - 6 / 2",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.ALERT });
      slide.addText("Mixed work", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.36,
        fontSize: 15, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "A: solve confidently", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "B: spot the error if any", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Show every step", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Use BODMAS reference", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.8,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addTextOnShape(slide, "12 minutes", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.45, w: lg.rightW - 0.4, h: 0.4, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 17: Exit Ticket
  exitTicketSlide(pres,
    [
      "Solve:  4 + 3 x (8 - 5) - 2 squared   (show steps)",
      "Find error or write CORRECT:  10 - 6 / 2 = 2",
      "Tell your teacher: which step of BODMAS gives you the most trouble?",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 18: Closing
  closingSlide(pres, {
    reflectionPrompt: "End of week 1 - tell your partner: how confident do you feel about BODMAS now?",
    scItems: [
      "I can apply BODMAS in the right order",
      "I can solve a full BODMAS equation correctly",
      "I can find and explain the mistake in a wrong answer",
    ],
    selfAssessment: { prompt: "Self-check", options: ["Got it", "Getting there", "Need more practice"] },
  }, NOTES_CLOSING);

  // Slide 19: Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "BOD_Session5_Full_BODMAS.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: PRACTICE_RES.name });
    let y = addPdfHeader(doc, PRACTICE_RES.name, {
      subtitle: "Full BODMAS Mixed Practice",
      color: C.NAVY,
      lessonInfo: "Session 5 of 10 | Year 5/6 Numeracy",
    });
    y = addTipBox(doc, "Mixed practice across all four BODMAS levels. Section B asks you to find the error in someone else's work.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A - Solve full BODMAS", y, { color: C.NAVY });
    y = addProblem(doc, 1, "3 squared + (10 - 4) - 2 =                    Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 2, "20 - 12 / 4 + 3 x 2 =                         Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 3, "(5 + 3) x 2 - 2 squared =                     Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 4, "4 squared / 2 + 5 - 3 =                       Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 5, "30 - (8 + 4) / 3 + 2 =                        Show every step", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Section B - Find error or write CORRECT", y, { color: C.NAVY });
    y = addProblem(doc, 6, "6 + 4 x 2 = 20                                R or W? If wrong, write the correct answer", y, { color: C.NAVY });
    y = addProblem(doc, 7, "18 / 6 + 7 = 10                               R or W? If wrong, write the correct answer", y, { color: C.NAVY });
    y = addProblem(doc, 8, "(8 - 3) x 4 = 20                              R or W? If wrong, write the correct answer", y, { color: C.NAVY });
    y = addProblem(doc, 9, "5 squared - 10 / 2 = 7                        R or W? If wrong, write the correct answer", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addProblem(doc, 10, "2 cubed + (12 - 4) x 3 - 6 / 2 =             Show every step", y, { color: C.NAVY });

    addPdfFooter(doc, "Session 5 | Mixed Practice | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, PRACTICE_RES.fileName));
    console.log("PDF written: " + PRACTICE_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Mixed Practice Answer Key",
      color: C.NAVY,
      lessonInfo: "Session 5 of 10 | Year 5/6 Numeracy",
    });

    y = addSectionHeading(doc, "Section A", y, { color: C.NAVY });
    y = addBodyText(doc, "1.  3 squared + (10 - 4) - 2  -->  9 + 6 - 2  -->  15 - 2 = 13", y);
    y = addBodyText(doc, "2.  20 - 12 / 4 + 3 x 2  -->  20 - 3 + 6  -->  17 + 6 = 23", y);
    y = addBodyText(doc, "3.  (5 + 3) x 2 - 2 squared  -->  8 x 2 - 4  -->  16 - 4 = 12", y);
    y = addBodyText(doc, "4.  4 squared / 2 + 5 - 3  -->  16 / 2 + 5 - 3  -->  8 + 5 - 3 = 10", y);
    y = addBodyText(doc, "5.  30 - (8 + 4) / 3 + 2  -->  30 - 12 / 3 + 2  -->  30 - 4 + 2 = 28", y);

    y = addSectionHeading(doc, "Section B - Errors", y, { color: C.NAVY });
    y = addBodyText(doc, "6.  6 + 4 x 2 = 20   WRONG (should be 4 x 2 = 8 first, then 6 + 8 = 14)", y);
    y = addBodyText(doc, "7.  18 / 6 + 7 = 10  CORRECT (3 + 7 = 10)", y);
    y = addBodyText(doc, "8.  (8 - 3) x 4 = 20 CORRECT (5 x 4 = 20)", y);
    y = addBodyText(doc, "9.  5 squared - 10 / 2 = 7  WRONG (25 - 5 = 20, not 7)", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "10.  2 cubed + (12 - 4) x 3 - 6 / 2", y);
    y = addBodyText(doc, "     B: (12 - 4) = 8", y);
    y = addBodyText(doc, "     O: 2 cubed = 8", y);
    y = addBodyText(doc, "     M and D: 8 x 3 = 24, 6 / 2 = 3", y);
    y = addBodyText(doc, "     A/S: 8 + 24 - 3  -->  32 - 3 = 29", y);

    addPdfFooter(doc, "Session 5 | Answer Key | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Session 5 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
