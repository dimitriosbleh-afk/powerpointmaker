"use strict";

// Teen Numbers Unit — Session 3: Ordering Numbers 0-20
// Term 2 Week 1, Foundation Numeracy, Variant 0
// DR: Patterns — I can recognise AB, AAB patterns
// Fluency: Ordering numbers up to at least 20
// VC2MFN01 — name, represent and order numbers including zero to at least 20

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "foundation", 0);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle,
  withReveal, runSlideDiagnostics, addNumberLine,
  STAGE_COLORS,
} = T;

const SESSION = 3;
const UNIT_TITLE = "Teen Numbers";
const FOOTER = "Teen Numbers | Session 3 of 5 | Foundation Numeracy";
const OUT_DIR = "output/Teen_Session3_Ordering_0_to_20";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Ordering and missing number practice 0-20.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with correct ordering.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Number line with visual guides for ordering.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Ordering beyond 20 and backwards ordering.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- We have been learning about teen numbers all week
- Today we are going to put numbers in ORDER from 0 to 20
- This means knowing which numbers come first and which come after

DO:
- Display title slide
- Have number cards 0-20 and a number line ready

TEACHER NOTES:
Session 3 of 5. Now that students can name 0-20, they need to sequence them. Ordering is the bridge between naming and understanding magnitude -- knowing that 15 is bigger than 12 because it comes after.

WATCH FOR:
- Students who are confident with naming but unsure about ordering -- naming and sequencing are different skills

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Let's warm up with some pattern work
- Look at each pattern. What comes next?
- Think about the rule -- what is repeating?

DO:
- Display pattern problems
- Students write answers on whiteboards
- Allow 15 seconds per pattern

TEACHER NOTES:
Daily Review retrieves pattern recognition from Term 1. AB and AAB patterns build the sequential thinking needed for number ordering later in this lesson.

WATCH FOR:
- Students who identify the repeating unit quickly -- pattern recognition is strong
- Students who struggle to see the repeat -- they may need concrete materials (physical pattern blocks)

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers
- Fix any that are different
- Ask: What was repeating in pattern 2? [circle, circle, square -- AAB]

DO:
- Click to reveal answers
- Students tick and fix
- Discuss the repeating unit for each pattern

TEACHER NOTES:
Naming the pattern type (AB, AAB) builds mathematical language. Students do not need to memorise the labels, but hearing them normalises the vocabulary.

WATCH FOR:
- Students who got all patterns correct -- pattern thinking is secure
- Students who struggled with AAB (three-part) but got AB -- they need the simpler pattern type first

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time. We are going to put numbers in order
- I will hold up number cards. You tell me which is bigger.
- Then we will put them in order from smallest to biggest
- Ready? Which is bigger: 5 or 12?

DO:
- Hold up pairs of number cards
- Students point to or call out the bigger number
- Then arrange 5 number cards in order (whole class activity)
- Use cards: 3, 11, 7, 15, 1

TEACHER NOTES:
Fluency builds ordering automaticity using a game-like format. Comparing pairs first, then ordering a set, scaffolds the complexity. Keep the pace brisk.

WATCH FOR:
- Students who compare correctly but struggle to order a set of 5 -- ordering multiple numbers is harder than comparing two
- Students who order correctly and quickly -- ordering is becoming automatic

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me
- Read from slide: "We are learning to put the numbers from 0 to 20 in order"
- Now our success criteria.
- Ask: If I count 1, 2, 3, 4, 5 -- am I putting numbers in order? [Yes!]

DO:
- Choral read LI and SC
- Quick check that students understand "in order"

TEACHER NOTES:
Ordering is the application of counting knowledge. SC1 is the floor (count 0-20). SC2 is the core (order a set). SC3 extends to finding missing numbers in a sequence.

WATCH FOR:
- Students who connect counting to ordering immediately -- strong foundation
- Students who are unsure what "in order" means -- clarify with a concrete example before I Do

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me put numbers in order
- I am going to use a number line. A number line shows numbers in order from smallest to biggest.
- Here is 0 on the left. Numbers get bigger as we move to the right.
- I put 1 next, then 2, then 3... all the way to 20.
- If I want to know which number comes after 14, I look at my number line: 15!

DO:
- Point to the number line on screen
- Trace along it left to right
- Demonstrate finding "what comes after" by pointing
- Show that numbers get bigger going right

TEACHER NOTES:
The number line is the key visual tool for ordering. Foundation students need to see that the number line is a physical path where bigger numbers are further right. This builds spatial understanding of magnitude.

WATCH FOR:
- Students who can follow the number line with their eyes -- spatial-numerical association is forming
- Students who point in the wrong direction (right to left) -- they need explicit direction modelling

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now watch me find a missing number
- Look: 8, 9, ___, 11. What number is missing?
- I use my number line. After 9 comes... 10! The missing number is 10.
- Another one: 15, 16, ___, 18. After 16 comes... 17!
- My strategy: say the numbers in your head and listen for the gap.

DO:
- Point to the number line for each example
- Trace the sequence to find the gap
- Emphasise the "say it in your head" strategy

TEACHER NOTES:
Finding missing numbers is harder than reciting the sequence because students must hold the surrounding numbers in working memory while identifying the gap. The number line reduces working memory load by making the sequence visible.

WATCH FOR:
- Students who find the missing number by counting on their fingers -- they are using the sequence even without the visual
- Students who shout the answer before you point -- they can hear the gap in the sequence

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. What number comes after 13?
- Use the number line if you need to.
- Show me on your whiteboard.

DO:
- Students write on whiteboards
- Scan for 14

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "What number comes after 13? You can look at the number line. Write it on your whiteboard. Show me."
- Scan for: 14. Students who write 14 quickly understand the sequence.
PROCEED: If 80%+ show 14, move to the We Do.
PIVOT: Most likely misconception -- students write 12 (going backwards instead of forwards). Reteach: "After means the NEXT number. When we count, 13 comes then... 14. After 13 is 14. After means bigger, not smaller."

TEACHER NOTES:
This CFU checks whether students can use the sequence to find the next number. The number line is available as a scaffold.

WATCH FOR:
- Students who answer without looking at the number line -- they have internalised the sequence
- Readiness signal: instant correct answers

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Partner time. Find the missing numbers
- Use the number line to help you
- Write the missing number on your whiteboard for each one

DO:
- Display four missing number problems
- Partners discuss and write on whiteboards
- Allow 20 seconds per problem

TEACHER NOTES:
We Do uses missing numbers in the 0-20 range. Problems progress: easy gap (single missing in familiar range), then harder (teen number gaps, gap near 20).

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 3 Enabling Scaffold which has a complete number line 0-20 with some numbers lightly printed. Students trace the printed numbers and fill in the blanks.
- Extra Notes: Distribute the Session 3 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Use the Session 3 Extension to order numbers beyond 20 and try ordering numbers from biggest to smallest (backwards ordering).
- Extra Notes: Self-contained investigation.

WATCH FOR:
- Students who find all missing numbers quickly -- sequencing is strong
- Students who struggle with teen number gaps (e.g., 14, ___, 16) -- they may still be shaky on 11-19 names

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check. 5, 6, ___, 8. What is missing? [7]
- 12, ___, 14, 15. What is missing? [13]
- 17, 18, ___, 20. What is missing? [19]
- 0, 1, 2, ___, 4. What is missing? [3]

DO:
- Reveal one at a time
- Cold call for each
- Point to the number line to verify

TEACHER NOTES:
The variety (low numbers, teen numbers, near 20) checks whether ordering works across the whole 0-20 range, not just in familiar territory.

WATCH FOR:
- Students who got all four -- ordering is secure across the range
- Students who struggled with 13 or 19 -- teen number gaps are the weak spot

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. I am going to show you numbers that are NOT in order
- You need to tell me: which number is in the WRONG spot?
- Look: 1, 2, 3, 5, 4. Which number is wrong?
- Think... whisper to your partner... now tell me!

DO:
- Students identify that 5 and 4 are swapped
- Cold call for explanation

CFU CHECKPOINT:
Technique: Turn and Talk
Script:
- Say: "Look at these numbers: 1, 2, 3, 5, 4. Something is wrong. Whisper to your partner -- which number is in the wrong spot?"
- Listen for: students identifying that 5 and 4 are swapped. "4 should come before 5."
PROCEED: If 80%+ identify the swap, move to You Do.
PIVOT: Most likely misconception -- students say the sequence "looks right" because all the numbers are present. Reteach: "Count with me: 1, 2, 3... what comes next? 4! But here it says 5. That is the wrong order. 4 comes before 5."

TEACHER NOTES:
Error detection is harder than sequencing because students must hold the correct sequence in mind and compare. This is a strong check of internalised ordering.

WATCH FOR:
- Students who spot the swap instantly -- strong sequence knowledge
- Readiness signal: students explaining WHY it is wrong, not just pointing

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice time
- On your worksheet, you will put numbers in order and find missing numbers
- First: Read the numbers. Next: Find the gap or put them in order. Then: Write the answer.
- You have 6 minutes.

DO:
- Distribute Session 3 Worksheet
- Students work independently
- Circulate: start with students who struggled
- Distribute scaffolds/extensions as needed

TEACHER NOTES:
You Do uses different sequences from the We Do. Section A is missing numbers; Section B is ordering a set from smallest to biggest. Different content ensures genuine transfer.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Complete the enabling scaffold with a printed number line. Students fill in blanks on the number line itself, with every second number already given.
- Extra Notes: Distribute the Session 3 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Use the Session 3 Extension to order numbers beyond 20 and practise backwards ordering (biggest to smallest).
- Extra Notes: Self-contained investigation.

WATCH FOR:
- Students who order sets correctly -- they can sequence without a number line
- Students who still need the number line -- that is okay at this stage
- Readiness signal: completing 6+ problems in 4 minutes

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Three quick questions. 2 minutes.

DO:
- Display exit ticket
- Students respond on whiteboards
- Sort responses

TEACHER NOTES:
Q1 assesses SC1 (counting sequence). Q2 assesses SC2 (ordering). Q3 assesses SC3 (missing number). Sort for Session 4 planning.

WATCH FOR:
- Students who get all three -- ready for writing teen numbers in Session 4
- Students who struggle with Q3 (missing number) -- they may need more number line support

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Let's check our success criteria.
- SC1: I can count from 0 to 20 in the right order. Thumbs?
- SC2: I can put numbers in order from smallest to biggest. Thumbs?
- SC3: I can find a missing number in a counting sequence. Thumbs?
- Turn and talk: What number comes between 16 and 18?

DO:
- Display SC
- Thumbs check
- Turn and Talk (answer: 17)
- Cold call 2-3 students

TEACHER NOTES:
The Turn and Talk applies today's skill in one focused question. Students who answer "17" instantly have internalised teen number ordering.

WATCH FOR:
- Students who answer 17 without hesitation -- ordering is secure
- Students who need to count from a lower number to get there -- they are getting there but not yet automatic

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Today's resources are linked here

DO:
- Point out resources

TEACHER NOTES:
Resource slide for Session 3.

WATCH FOR:
- N/A

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Session 3: Ordering Numbers 0 to 20",
    "Foundation Numeracy | Session 3 of 5 | Term 2 Week 1", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2-3: Daily Review — Patterns AB, AAB (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Pattern Power", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      s.addText([
        { text: "What comes next?", options: { fontSize: 16, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "1.  red, blue, red, blue, red, ___", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true, paraSpaceAfter: 10 } },
        { text: "2.  clap, clap, stomp, clap, clap, ___", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true, paraSpaceAfter: 10 } },
        { text: "3.  star, circle, star, circle, star, ___", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true, paraSpaceAfter: 10 } },
        { text: "4.  big, big, small, big, big, ___", options: { fontSize: 14, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboards:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "Write or draw what comes next.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Look for the repeat!", options: { fontSize: 14, bold: true, color: C.ALERT, breakLine: true } },
        { text: "AB pattern = 2 parts repeating", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "AAB pattern = 3 parts repeating", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) blue (AB)    2) stomp (AAB)    3) circle (AB)    4) small (AAB)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 12, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Ordering numbers
  contentSlide(pres, "Fluency", STAGE_COLORS["1"], "Order These Numbers",
    [
      "Which is bigger?",
      "  5 or 12?          8 or 3?",
      "  14 or 11?         19 or 20?",
      "",
      "Now put these in order (smallest to biggest):",
      "  3, 11, 7, 15, 1",
    ],
    NOTES_FLUENCY, FOOTER,
    (slide, lg) => {
      // Number line 0-20 on the right
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.8, { strip: STAGE_COLORS["1"] });
      slide.addText("Number Line", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.25,
        fontSize: 13, fontFace: FONT_H, color: STAGE_COLORS["1"], bold: true, margin: 0, align: "center",
      });

      // Vertical number sequence
      const nums = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20];
      const gridCols = 7;
      const gridRows = 3;
      const gCellW = (lg.rightW - 0.4) / gridCols;
      const gCellH = 0.5;
      nums.forEach((n, i) => {
        const col = i % gridCols;
        const row = Math.floor(i / gridCols);
        const cx = lg.rightX + 0.2 + col * gCellW;
        const cy = lg.panelTopPadded + 0.4 + row * gCellH;
        slide.addShape("roundRect", {
          x: cx, y: cy, w: gCellW - 0.04, h: gCellH - 0.06, rectRadius: 0.04,
          fill: { color: n >= 11 && n <= 20 ? C.PRIMARY : C.BG_LIGHT },
          line: { color: C.PRIMARY, width: 0.8 },
        });
        slide.addText(String(n), {
          x: cx, y: cy, w: gCellW - 0.04, h: gCellH - 0.06,
          fontSize: 13, fontFace: FONT_H, color: n >= 11 && n <= 20 ? C.WHITE : C.CHARCOAL,
          align: "center", valign: "middle", bold: true, margin: 0,
        });
      });

      addTextOnShape(slide, "Answer: 1, 3, 7, 11, 15", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 2.2, w: lg.rightW - 0.6, h: 0.35, rectRadius: 0.06,
        fill: { color: C.SUCCESS },
      }, { fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 5: LI/SC
  liSlide(pres,
    ["We are learning to put the numbers from 0 to 20 in order"],
    [
      "I can count from 0 to 20 in the right order",
      "I can put numbers in order from smallest to biggest",
      "I can find a missing number in a counting sequence",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — Number line introduction
  workedExSlide(pres, 2, "I Do", "Numbers in Order: The Number Line",
    [
      "A number line shows numbers in ORDER",
      "Small numbers are on the LEFT",
      "Big numbers are on the RIGHT",
      "",
      "0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10,",
      "11, 12, 13, 14, 15, 16, 17, 18, 19, 20",
      "",
      "What comes after 14? Look at the line: 15!",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.3, { strip: C.PRIMARY });
      slide.addText("Number Line 0-20", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.2, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      // Draw a vertical number path (more readable for Foundation)
      for (let i = 0; i <= 20; i++) {
        const col = i <= 10 ? 0 : 1;
        const row = i <= 10 ? i : i - 11;
        const cx = lg.rightX + 0.2 + col * 2.0;
        const cy = lg.panelTopPadded + 0.45 + row * 0.25;
        const isHighlight = i === 14 || i === 15;
        slide.addShape("roundRect", {
          x: cx, y: cy, w: 1.6, h: 0.22, rectRadius: 0.04,
          fill: { color: isHighlight ? C.ALERT : (i >= 11 ? C.PRIMARY : C.BG_LIGHT) },
          line: { color: C.PRIMARY, width: 0.6 },
        });
        slide.addText(String(i), {
          x: cx, y: cy, w: 1.6, h: 0.22,
          fontSize: 11, fontFace: FONT_H, color: isHighlight ? C.WHITE : (i >= 11 ? C.WHITE : C.CHARCOAL),
          align: "center", valign: "middle", bold: true, margin: 0,
        });
      }
    }
  );

  // Slide 7: I Do — Finding missing numbers
  workedExSlide(pres, 2, "I Do", "Finding Missing Numbers",
    [
      "Strategy: Say the numbers in your head.",
      "Listen for the gap!",
      "",
      "Example 1:  8, 9, ___, 11",
      "  Say: 8, 9, TEN, 11. The gap is 10!",
      "",
      "Example 2:  15, 16, ___, 18",
      "  Say: 15, 16, SEVENTEEN, 18.",
      "  The gap is 17!",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.2, { strip: C.SECONDARY });
      slide.addText("My Strategy:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
      });
      slide.addText([
        { text: "1. Read the numbers around the gap", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "2. Say the counting sequence in your head", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "3. Listen for the missing number", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "4. Write it in the gap!", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.5, h: 1.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Slide 8-9: CFU (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "What Comes After 13?", "Show Me Boards",
      "What number comes AFTER 13?\n\nUse the number line if you need to.\n\nWrite it on your whiteboard!",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "14!  After 13 comes 14.", {
        x: 1.0, y: 4.0, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10-11: We Do (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Find the Missing Number", { y: 0.65, fontSize: 22, color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });
      s.addText([
        { text: "With your partner:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "Find the missing number in each sequence.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "A.  5, 6, ___, 8", options: { fontSize: 16, bold: true, color: C.PRIMARY, breakLine: true, paraSpaceAfter: 12 } },
        { text: "B.  12, ___, 14, 15", options: { fontSize: 16, bold: true, color: C.PRIMARY, breakLine: true, paraSpaceAfter: 12 } },
        { text: "C.  17, 18, ___, 20", options: { fontSize: 16, bold: true, color: C.PRIMARY, breakLine: true, paraSpaceAfter: 12 } },
        { text: "D.  0, 1, 2, ___, 4", options: { fontSize: 16, bold: true, color: C.PRIMARY } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      s.addText([
        { text: "Use the number line!", options: { fontSize: 15, bold: true, color: C.ACCENT, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Say the numbers in your head.", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Listen for the gap.", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Write the missing number.", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "A) 7    B) 13    C) 19    D) 3", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 12-13: CFU 2 Hinge (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Spot the Mistake!", "Turn and Talk",
      "These numbers are supposed to be in order:\n\n1,  2,  3,  5,  4\n\nWhich number is in the WRONG spot?\nWhisper to your partner!",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "5 and 4 are swapped! Correct order: 1, 2, 3, 4, 5", {
        x: 0.5, y: 4.0, w: 9, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do
  workedExSlide(pres, 4, "You Do", "Your Turn: Ordering Practice",
    [
      "First: Read the numbers on your worksheet.",
      "Next: Find the missing number or put them in order.",
      "Then: Write your answer.",
      "",
      "Use the number line at the top of your page",
      "if you need help!",
      "",
      "You have 6 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.6, { strip: C.ALERT });
      slide.addText("Strategy:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Say it in your head", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Listen for the gap", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Use the number line to check", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.5, h: 0.9,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "What number comes after 17?",
      "Put these in order (smallest to biggest): 9, 3, 15, 7",
      "Find the missing number: 11, 12, ___, 14",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 16: Closing
  closingSlide(pres,
    "What number comes between 16 and 18? Tell your partner!",
    [
      "I can count from 0 to 20 in the right order",
      "I can put numbers in order from smallest to biggest",
      "I can find a missing number in a counting sequence",
    ],
    NOTES_CLOSING);

  // Slide 17: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Teen_Session3_Ordering_0_to_20.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ──────────────────────────────────────────────────────────────────

  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, { subtitle: "Ordering Numbers 0-20", color: C.PRIMARY, lessonInfo: "Session 3 of 5 | Foundation Numeracy" });
    y = addTipBox(doc, "Say the numbers in your head. Listen for the gap. Write the missing number!", y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section A: Find the missing number", y, { color: C.PRIMARY });
    y = addProblem(doc, 1, "3, 4, ___, 6, 7", y, { color: C.PRIMARY });
    y = addProblem(doc, 2, "10, ___, 12, 13", y, { color: C.PRIMARY });
    y = addProblem(doc, 3, "14, 15, ___, 17", y, { color: C.PRIMARY });
    y = addProblem(doc, 4, "18, ___, 20", y, { color: C.PRIMARY });
    y = addProblem(doc, 5, "6, 7, 8, ___, 10", y, { color: C.PRIMARY });
    y = addProblem(doc, 6, "0, 1, ___, 3, 4", y, { color: C.PRIMARY });

    y = addSectionHeading(doc, "Section B: Put in order (smallest to biggest)", y, { color: C.PRIMARY });
    y = addProblem(doc, 7, "5, 2, 8, 1, 4   ->   ___, ___, ___, ___, ___", y, { color: C.PRIMARY });
    y = addProblem(doc, 8, "13, 10, 16, 11, 14   ->   ___, ___, ___, ___, ___", y, { color: C.PRIMARY });

    addPdfFooter(doc, "Session 3 | Teen Numbers | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF: " + WORKSHEET_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, { subtitle: "Ordering Numbers 0-20", color: C.PRIMARY, lessonInfo: "Session 3 of 5 | Foundation Numeracy" });
    y = addSectionHeading(doc, "Section A", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1. 5    2. 11    3. 16    4. 19    5. 9    6. 2", y);
    y = addSectionHeading(doc, "Section B", y, { color: C.PRIMARY });
    y = addBodyText(doc, "7. 1, 2, 4, 5, 8", y);
    y = addBodyText(doc, "8. 10, 11, 13, 14, 16", y);
    addPdfFooter(doc, "Session 3 | Answer Key | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF: " + ANSWER_KEY_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, { subtitle: "Number Line Helper", color: C.ACCENT, lessonInfo: "Session 3 of 5 | Foundation Numeracy" });
    y = addTipBox(doc, "Use the number line below. Every second number is given. Fill in the blanks!", y, { color: C.ACCENT });
    y = addSectionHeading(doc, "Fill in the missing numbers on the number line", y, { color: C.PRIMARY });
    y = addBodyText(doc, "0, ___, 2, ___, 4, ___, 6, ___, 8, ___, 10", y);
    y = addBodyText(doc, "10, ___, 12, ___, 14, ___, 16, ___, 18, ___, 20", y);
    y += 10;
    y = addSectionHeading(doc, "Now find these missing numbers", y, { color: C.PRIMARY });
    y = addProblem(doc, 1, "2, 3, ___, 5", y, { color: C.PRIMARY });
    y = addProblem(doc, 2, "11, ___, 13", y, { color: C.PRIMARY });
    y = addProblem(doc, 3, "16, 17, ___, 19", y, { color: C.PRIMARY });
    addPdfFooter(doc, "Session 3 | Enabling Scaffold | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF: " + ENABLING_RES.fileName);
  })();

  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, { subtitle: "Ordering Beyond 20", color: C.PRIMARY, lessonInfo: "Session 3 of 5 | Foundation Numeracy" });
    y = addSectionHeading(doc, "Numbers Keep Going!", y, { color: C.PRIMARY });
    y = addBodyText(doc, "After 20 comes 21, then 22, then 23...", y);
    y = addBodyText(doc, "The pattern is the same: just keep counting!", y);
    y = addSectionHeading(doc, "Fill in the missing numbers", y, { color: C.PRIMARY });
    y = addProblem(doc, 1, "18, 19, 20, ___, ___", y, { color: C.PRIMARY });
    y = addProblem(doc, 2, "20, ___, 22, ___, 24", y, { color: C.PRIMARY });
    y = addProblem(doc, 3, "25, ___, ___, 28, 29", y, { color: C.PRIMARY });
    y = addSectionHeading(doc, "Backwards Ordering (biggest to smallest)", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Can you put these in order from BIGGEST to smallest?", y);
    y = addProblem(doc, 4, "5, 12, 3, 18, 9   ->   ___, ___, ___, ___, ___", y, { color: C.PRIMARY });
    y = addProblem(doc, 5, "20, 15, 8, 17, 2   ->   ___, ___, ___, ___, ___", y, { color: C.PRIMARY });
    y = addTipBox(doc, "Challenge: Can you count backwards from 20 to 0? Try it!", y, { color: C.ACCENT });
    addPdfFooter(doc, "Session 3 | Extension | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 3 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
