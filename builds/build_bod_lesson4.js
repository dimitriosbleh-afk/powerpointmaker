"use strict";

// BODMAS Unit - Session 4: Addition and Subtraction (left to right)
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

const SESSION = 4;
const FOOTER = "BODMAS | Session 4 of 10 | Year 5/6 Numeracy";
const OUT_DIR = "output/BOD_Session4_Addition_Subtraction";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const PRACTICE_RES = makeSessionResource(SESSION, "Practice Sheet", "Addition and subtraction left-to-right plus all-BODMAS practice.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference.");
const RESOURCE_ITEMS = [PRACTICE_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// --- Teacher Notes -----------------------------------------------------------

const NOTES_TITLE = `SAY:
- Day 4 of BODMAS
- Today is the same idea as yesterday but with the LAST two letters
- A and S are at the same level - work left to right

DO:
- Display title slide
- BODMAS reference visible

TEACHER NOTES:
Session 4 of 10. Today mirrors yesterday's lesson but for A and S. Same rule: same-level operations work left to right. By the end, students should fluently apply BODMAS across all four levels.

WATCH FOR:
- Students who think yesterday's lesson was "just for division" - we are generalising the rule today

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Daily Review. Yesterday's M and D rule
- Solve each one - work left to right for x and /

DO:
- Display 4 review equations
- Allow 4 minutes

TEACHER NOTES:
Daily Review keeps yesterday's M and D rule fresh. The rule must transfer cleanly to today's A and S.

WATCH FOR:
- Students reverting to "M before D" - the persistent error

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Tick or fix
- Hands up if you got all four

DO:
- Reveal answers

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency. Mixed addition and subtraction facts
- Quick whiteboard responses
- We're warming up the A and S brain muscles for today

DO:
- Display 8 facts
- Brisk - 6 seconds per fact
- Whiteboard show-me

TEACHER NOTES:
Mixed A and S fact recall. Builds fluency and primes the A-and-S thinking for today.

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_FLUENCY_A = `SAY:
- Tick or fix
- Note any you missed

DO:
- Reveal answers

[Stage 1: Fluency Answers | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Today's intention: A and S are also at the same level
- Read SC together
- Ask: True or false - we always do A before S? [FALSE - same as yesterday]

DO:
- Choral read
- Quick true/false

TEACHER NOTES:
Today's threshold concept extends yesterday's rule to A and S. Students should now hold a unified rule: D and M are equal, A and S are equal, both work left to right.

WATCH FOR:
- Students who say "TRUE - A before S because BODMAS" - we are teaching against this

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_HOOK = `SAY:
- Look at this equation: 20 - 5 + 3
- On your whiteboards. Don't talk yet
- Show me

DO:
- Display 20 - 5 + 3
- Wait for whiteboards
- Tally: how many wrote 18? How many wrote 12?

TEACHER NOTES:
The hook surfaces the misconception in A and S. Students who do A first get 5+3=8, then 20-8=12. Correct (left to right) gets 20-5=15, then 15+3=18. Most will get this wrong if they applied yesterday's lesson incorrectly or remembered "A before S".

MISCONCEPTIONS:
- Misconception: A always comes before S
  Why: BODMAS lists A before S
  Impact: Wrong answers on equations like 20 - 5 + 3 (correct: 18, wrong: 12)
  Quick correction: "A and S - SAME level. Whichever one comes FIRST as you read left to right is the one you do first"

WATCH FOR:
- Students who got 12 - they applied "A before S"
- Students who got 18 - they have transferred yesterday's rule correctly

[Stage 2: Launch | VTLM 2.0: Engagement]`;

const NOTES_RULE = `SAY:
- Same rule as yesterday, different letters
- A and S are at the SAME level
- Work LEFT to RIGHT
- Whichever comes first in the equation gets done first
- 20 - 5 + 3 - subtraction comes first reading L to R
- 20 - 5 = 15, then 15 + 3 = 18

DO:
- Write the rule on the board
- Use a colour to show the order

TEACHER NOTES:
The explicit rule. Same as yesterday's M and D rule. The two rules combine: M/D = same level, A/S = same level. The full BODMAS now has just FOUR levels, not six.

WATCH FOR:
- Students who notice "this is the same rule as yesterday" - excellent transfer

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_BIG_PICTURE = `SAY:
- Now you know the FULL BODMAS picture
- Six letters but only FOUR levels
- B by itself
- O by itself
- D and M together (left to right)
- A and S together (left to right)
- Same level always means: work left to right

DO:
- Show the BODMAS pyramid visual
- Trace each level with a finger
- Emphasise the four levels

TEACHER NOTES:
This big-picture moment consolidates the entire unit so far. Students should leave today seeing BODMAS as four levels, not six. This becomes the working mental model for the rest of the unit.

WATCH FOR:
- Students who say "ohhh!" or look surprised - the picture is forming
- Students who still see six steps - they may need the visual longer

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. Whiteboards
- 15 - 4 + 6
- Final answer

DO:
- Students compute and show
- Scan for "17"
- 17 means correct (15-4=11, 11+6=17)
- 5 means they did A before S (15 - 10 = 5) - error

TEACHER NOTES:
Direct CFU on today's rule. Wrong answer = "A before S" misconception.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Final answer
- Show me
- Scan for: 17
PROCEED:
- 80% or more show 17 - move on
PIVOT:
- Most likely misconception: A before S
- Reteach: "A and S - SAME level. Subtraction is FIRST reading left to right"
- Re-check: 20 - 7 + 3 [answer: 16]

WATCH FOR:
- Students showing 5 - persistent BODMAS-literal error
- Students showing 17 - rule has landed

[Stage 2: CFU | VTLM 2.0: Active Checking]`;

const NOTES_CFU1_A = `SAY:
- Answer: 17
- 15 - 4 = 11, then 11 + 6 = 17
- Subtraction first because it came first reading left to right

DO:
- Reveal full working

[Stage 2: CFU Answer | VTLM 2.0: Active Checking]`;

const NOTES_WEDO_Q = `SAY:
- Together. All four BODMAS levels
- 20 - 4 x 3 + 6 / 2
- What's first? B - no brackets. O - no orders
- D/M: x and / - both at same level. 4 x 3 = 12 (came first), then 6 / 2 = 3
- Equation now: 20 - 12 + 3
- A/S: subtraction first (came first L to R). 20 - 12 = 8, then 8 + 3 = 11

DO:
- Walk through with class
- Pause for choral response at each step
- Show the sub-steps clearly

TEACHER NOTES:
This is the key consolidation - all four levels in one equation. Students should see the structure: do the higher-level steps first, then chunk back through the lower-level steps left to right.

WATCH FOR:
- Students who try to do A before S in the final step - reinforce
- Students who jump levels - reinforce: do all D/M first, then all A/S

[Stage 3: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_A = `SAY:
- Final: 11
- Tick or fix

DO:
- Reveal answer

[Stage 3: We Do Answer | VTLM 2.0: Guided Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge question. Whiteboards
- 30 - 12 / 4 + 5
- Final answer
- Show your steps in your book too

DO:
- Students compute and show
- Scan for "32"
- 32 means correct (12/4=3, 30-3=27, 27+5=32)

TEACHER NOTES:
Hinge tests both the M/D rule (do division first) AND the A/S rule (subtract first reading L to R). The threshold question for today.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Final answer
- Show me
- Scan for: 32
PROCEED:
- 80% or more show 32 - go to You Do
PIVOT:
- Most likely misconception: did A before S, or jumped to A/S before D/M
- Reteach: "Four levels - which apply here? D/M first (the division), then A/S (subtraction first L to R)"
- Re-check: 40 - 8 / 2 + 3 [answer: 39]

WATCH FOR:
- Students showing 32 confidently - solid all-BODMAS
- Students showing wrong answers - flag for small group

[Stage 3: Hinge CFU | VTLM 2.0: Active Checking]`;

const NOTES_CFU2_A = `SAY:
- Answer: 32
- D/M first: 12 / 4 = 3
- A/S left to right: 30 - 3 = 27, 27 + 5 = 32

DO:
- Reveal full working

[Stage 3: Hinge Answer | VTLM 2.0: Active Checking]`;

const NOTES_YOUDO = `SAY:
- Independent practice
- Show every step
- Check yourself: did you do the right level first?

DO:
- Hand out practice sheet
- Allow 10 minutes
- Pull small group as needed

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Pull small group. Use BODMAS reference. For each equation, write B-O-D/M-A/S down the side and tick which appear
- Extra Notes: Solve in BODMAS order. Build the habit
EXTENDING PROMPT:
- Task: Challenge: 50 - 6 x 3 + 4 squared - (2 + 1)
- Extra Notes: All four levels. Answer: 39

WATCH FOR:
- Students who skip levels - reinforce the order
- Students who lose track in long equations - prompt them to rewrite the equation after each step

[Stage 4: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_EXIT = `SAY:
- Three exit ticket questions
- Show your steps

DO:
- Hand out exit ticket
- Collect

TEACHER NOTES:
Q1 simple A/S, Q2 mixed all levels, Q3 explanation. Tests SC2.

[Stage 5: Exit Ticket | VTLM 2.0: Evidence of Learning]`;

const NOTES_CLOSING = `SAY:
- Read SC together
- Self-check
- Tell your partner: how many LEVELS in BODMAS now?
- Answer: four. B, O, D/M, A/S

DO:
- Closing slide
- Self-check
- Turn and talk

TEACHER NOTES:
Closing reinforces the four-level mental model. Tomorrow we put it all together.

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
    "A and S - Same Level",
    "Addition and Subtraction - work left to right",
    "Session 4 of 10  |  Year 5/6 Numeracy",
    NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2-3: Daily Review
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "M and D Recap from Yesterday", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const dr = [
        "1.  24 / 6 x 5",
        "2.  4 x 18 / 6",
        "3.  60 / 5 x 2",
        "4.  9 x 8 / 12",
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
        { text: "M and D - same level", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Strict left to right", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
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
      addTextOnShape(slide, "1) 20    2) 12    3) 24    4) 6", {
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
      addTitle(s, "Mixed Addition and Subtraction", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 9, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const fluency = [
        "47 + 29", "82 - 35", "156 + 78",
        "234 - 89", "65 + 47", "120 - 38",
        "99 + 56", "303 - 147",
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
      addTextOnShape(slide, "76  |  47  |  234  |  145  |  112  |  82  |  155  |  156", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FLUENCY_A);
    }
  );

  // Slide 6: LI/SC
  liSlide(pres,
    ["I am learning that Addition and Subtraction are at the SAME LEVEL"],
    [
      "I can identify which of A or S comes first reading left to right",
      "I can solve a full BODMAS equation with all four levels",
      "I can explain that BODMAS has FOUR levels, not six",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 7: Hook
  contentSlide(pres, "Launch", C.PRIMARY, "What does this equal?",
    [
      "Work it out on your whiteboard",
      "Don't talk yet",
      "Show me",
    ],
    NOTES_HOOK, FOOTER,
    (slide, lg) => {
      addTextOnShape(slide, "20 - 5 + 3 = ?", {
        x: lg.rightX, y: lg.panelTopPadded, w: lg.rightW, h: 1.6, rectRadius: 0.12,
        fill: { color: C.PRIMARY },
      }, { fontSize: 40, fontFace: FONT_H, color: C.WHITE, bold: true });

      addTextOnShape(slide, "Some say:\n18", {
        x: lg.rightX, y: lg.panelTopPadded + 1.85, w: 2.0, h: 1.3, rectRadius: 0.1,
        fill: { color: C.SECONDARY },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });

      addTextOnShape(slide, "Some say:\n12", {
        x: lg.rightX + 2.2, y: lg.panelTopPadded + 1.85, w: 2.0, h: 1.3, rectRadius: 0.1,
        fill: { color: C.ACCENT },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });

      slide.addText("Sound familiar?", {
        x: lg.rightX, y: lg.panelTopPadded + 3.25, w: lg.rightW, h: 0.4,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });
    }
  );

  // Slide 8: The rule
  contentSlide(pres, "I Do", C.PRIMARY, "A and S are at the SAME LEVEL",
    [
      "Same idea as yesterday with M and D",
      "BODMAS lists A before S - but they are EQUAL",
      "Work LEFT TO RIGHT when you reach this step",
      "",
      "20 - 5 + 3:",
      "  Subtraction comes first (reading L to R)",
      "  20 - 5 = 15, then 15 + 3 = 18",
    ],
    NOTES_RULE, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.ALERT });
      slide.addText("THE RULE", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.4,
        fontSize: 16, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0, align: "center",
      });

      slide.addText([
        { text: "A and S = ", options: { fontSize: 22, color: C.CHARCOAL, bold: true, breakLine: false } },
        { text: "SAME LEVEL", options: { fontSize: 22, color: C.ALERT, bold: true, breakLine: true } },
      ], {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.65, w: lg.rightW - 0.4, h: 0.6,
        fontFace: FONT_H, margin: 0, align: "center",
      });

      slide.addText("--> Read LEFT to RIGHT -->", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 1.4, w: lg.rightW - 0.4, h: 0.5,
        fontSize: 18, fontFace: FONT_B, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });

      slide.addText("Whichever you see first --> do first", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.0, w: lg.rightW - 0.4, h: 0.4,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, italic: true, align: "center", margin: 0,
      });

      addTextOnShape(slide, "20 - 5 + 3 = 18", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.6, w: lg.rightW - 0.4, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 9: Big picture - 4 levels
  contentSlide(pres, "I Do", C.PRIMARY, "BODMAS - The Big Picture (Four Levels)",
    [
      "Six letters - but only FOUR levels",
      "Same level always means: work left to right",
      "",
      "Level 1: B (Brackets)",
      "Level 2: O (Orders)",
      "Level 3: D and M (left to right)",
      "Level 4: A and S (left to right)",
    ],
    NOTES_BIG_PICTURE, FOOTER,
    (slide, lg) => {
      // 4-level pyramid visual
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.PRIMARY });
      slide.addText("BODMAS Pyramid", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      const levels = [
        { L: "1", letters: "B", word: "Brackets",       color: C.PRIMARY },
        { L: "2", letters: "O", word: "Orders",         color: C.SECONDARY },
        { L: "3", letters: "D / M", word: "Left to right", color: C.ACCENT },
        { L: "4", letters: "A / S", word: "Left to right", color: C.ALERT },
      ];
      levels.forEach((lev, i) => {
        const y = lg.panelTopPadded + 0.50 + i * 0.65;
        // Level number
        addTextOnShape(slide, lev.L, {
          x: lg.rightX + 0.2, y, w: 0.5, h: 0.55, rectRadius: 0.06,
          fill: { color: lev.color },
        }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
        // Letters
        addTextOnShape(slide, lev.letters, {
          x: lg.rightX + 0.8, y, w: 1.3, h: 0.55, rectRadius: 0.06,
          fill: { color: lev.color },
        }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
        // Word
        slide.addText(lev.word, {
          x: lg.rightX + 2.2, y, w: lg.rightW - 2.4, h: 0.55,
          fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });
    }
  );

  // Slide 10-11: CFU 1
  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check", "Show Me Boards",
      "15 - 4 + 6\n\nFinal answer.",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Answer: 17   (15 - 4 = 11, then 11 + 6 = 17)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU1_A);
    }
  );

  // Slide 12-13: We Do
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Together - All Four Levels",
      [
        "Equation: 20 - 4 x 3 + 6 / 2",
        "",
        "Brackets? No. Orders? No.",
        "D/M (left to right):",
        "  4 x 3 = 12 (came first)",
        "  6 / 2 = 3",
        "  Now: 20 - 12 + 3",
        "",
        "A/S (left to right):",
        "  20 - 12 = 8",
        "  8 + 3 = 11",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.4, { strip: C.SECONDARY });
        slide.addText("Levels in this equation", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.36,
          fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
        });

        const used = [
          { L: "B", on: false, text: "No brackets" },
          { L: "O", on: false, text: "No orders" },
          { L: "D/M", on: true, text: "Used: x and /" },
          { L: "A/S", on: true, text: "Used: - and +" },
        ];
        used.forEach((u, i) => {
          const y = lg.panelTopPadded + 0.55 + i * 0.55;
          addTextOnShape(slide, u.L, {
            x: lg.rightX + 0.25, y, w: 0.7, h: 0.42, rectRadius: 0.06,
            fill: { color: u.on ? C.SUCCESS : C.MUTED },
          }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
          slide.addText(u.text, {
            x: lg.rightX + 1.05, y, w: lg.rightW - 1.25, h: 0.42,
            fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
          });
        });

        addTextOnShape(slide, "Same-level = L to R", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 2.95, w: lg.rightW - 0.4, h: 0.4, rectRadius: 0.06,
          fill: { color: C.PRIMARY },
        }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "Final: 11   |   D/M: 12 and 3   |   A/S: 20-12=8, 8+3=11", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 14-15: Hinge CFU
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Question", "Show Me Boards",
      "30 - 12 / 4 + 5\n\nFinal answer.",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Answer: 32   (12/4=3, 30-3=27, 27+5=32)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU2_A);
    }
  );

  // Slide 16: You Do
  workedExSlide(pres, 4, "You Do", "Independent Practice",
    [
      "Show every step",
      "Tick off the BODMAS levels you used",
      "",
      "1.  18 - 7 + 4",
      "2.  25 - 6 + 9 - 3",
      "3.  20 - 4 x 3 + 5",
      "4.  6 + 18 / 3 - 2",
      "5.  3 squared - 4 + 6",
      "",
      "Challenge: 50 - 6 x 3 + 4 squared - (2 + 1)",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.ALERT });
      slide.addText("Four levels", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.36,
        fontSize: 15, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "1. Brackets", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "2. Orders (powers)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "3. D and M (L to R)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "4. A and S (L to R)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
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
      "Solve:  25 - 8 + 3   (show steps)",
      "Solve:  20 - 6 / 2 + 4   (show steps)",
      "Tell your teacher: How many LEVELS in BODMAS?",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 18: Closing
  closingSlide(pres, {
    reflectionPrompt: "Tell your partner: BODMAS has six letters but only how many levels?",
    scItems: [
      "I can identify which of A or S comes first reading left to right",
      "I can solve a full BODMAS equation with all four levels",
      "I can explain that BODMAS has FOUR levels, not six",
    ],
    selfAssessment: { prompt: "Self-check", options: ["Got it", "Getting there", "Need more practice"] },
  }, NOTES_CLOSING);

  // Slide 19: Resources


  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "BOD_Session4_Addition_Subtraction.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // --- PDFs ------------------------------------------------------------------

  await (async () => {
    const doc = createPdf({ title: PRACTICE_RES.name });
    let y = addPdfHeader(doc, PRACTICE_RES.name, {
      subtitle: "A and S - Same Level",
      color: C.NAVY,
      lessonInfo: "Session 4 of 10 | Year 5/6 Numeracy",
    });
    y = addTipBox(doc, "A and S are same-level. Work LEFT TO RIGHT. Same rule as yesterday's M and D.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A - A and S only", y, { color: C.NAVY });
    y = addProblem(doc, 1, "18 - 7 + 4 =                          Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 2, "25 - 6 + 9 - 3 =                      Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 3, "100 - 25 + 10 - 5 =                   Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 4, "40 + 15 - 8 + 12 =                    Show every step", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Section B - Mixed (D/M and A/S)", y, { color: C.NAVY });
    y = addProblem(doc, 5, "20 - 4 x 3 + 5 =                      Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 6, "6 + 18 / 3 - 2 =                      Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 7, "30 - 12 / 4 + 5 =                     Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 8, "8 x 4 - 20 + 6 =                      Show every step", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Section C - All four levels", y, { color: C.NAVY });
    y = addProblem(doc, 9, "3 squared - 4 + 6 =                   Show every step", y, { color: C.NAVY });
    y = addProblem(doc, 10, "(8 + 4) - 2 squared =                Show every step", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addProblem(doc, 11, "50 - 6 x 3 + 4 squared - (2 + 1) =   Show every step", y, { color: C.NAVY });

    addPdfFooter(doc, "Session 4 | Practice Sheet | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, PRACTICE_RES.fileName));
    console.log("PDF written: " + PRACTICE_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "A and S Practice Answer Key",
      color: C.NAVY,
      lessonInfo: "Session 4 of 10 | Year 5/6 Numeracy",
    });

    y = addSectionHeading(doc, "Section A - A and S only", y, { color: C.NAVY });
    y = addBodyText(doc, "1.  18 - 7 + 4    -->  11 + 4 = 15", y);
    y = addBodyText(doc, "2.  25 - 6 + 9 - 3  -->  19 + 9 - 3  -->  28 - 3 = 25", y);
    y = addBodyText(doc, "3.  100 - 25 + 10 - 5  -->  75 + 10 - 5  -->  85 - 5 = 80", y);
    y = addBodyText(doc, "4.  40 + 15 - 8 + 12  -->  55 - 8 + 12  -->  47 + 12 = 59", y);

    y = addSectionHeading(doc, "Section B - Mixed", y, { color: C.NAVY });
    y = addBodyText(doc, "5.  20 - 4 x 3 + 5  -->  20 - 12 + 5  -->  8 + 5 = 13", y);
    y = addBodyText(doc, "6.  6 + 18 / 3 - 2  -->  6 + 6 - 2  -->  12 - 2 = 10", y);
    y = addBodyText(doc, "7.  30 - 12 / 4 + 5  -->  30 - 3 + 5  -->  27 + 5 = 32", y);
    y = addBodyText(doc, "8.  8 x 4 - 20 + 6  -->  32 - 20 + 6  -->  12 + 6 = 18", y);

    y = addSectionHeading(doc, "Section C - All four levels", y, { color: C.NAVY });
    y = addBodyText(doc, "9.  3 squared - 4 + 6  -->  9 - 4 + 6  -->  5 + 6 = 11", y);
    y = addBodyText(doc, "10.  (8 + 4) - 2 squared  -->  12 - 4 = 8", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "11.  50 - 6 x 3 + 4 squared - (2 + 1)", y);
    y = addBodyText(doc, "     B: (2 + 1) = 3", y);
    y = addBodyText(doc, "     O: 4 squared = 16", y);
    y = addBodyText(doc, "     M: 6 x 3 = 18", y);
    y = addBodyText(doc, "     A/S left to right: 50 - 18 + 16 - 3  -->  32 + 16 - 3  -->  48 - 3 = 45", y);

    addPdfFooter(doc, "Session 4 | Answer Key | Year 5/6");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Session 4 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
