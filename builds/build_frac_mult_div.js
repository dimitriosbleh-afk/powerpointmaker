"use strict";

// Fractions — Session 1: Multiplying and Dividing Fractions
// Year 6 Numeracy, Variant 2
// DR: Coordinates and area, perimeter, volume
// Fluency: Division
// VC2M6N07 — solve problems involving multiplication/division of fractions

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", 2);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle,
  withReveal,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const FOOTER = "Multiplying and Dividing Fractions | Year 6 Numeracy";
const OUT_DIR = "output/Fractions_Multiplying_Dividing";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES  = makeSessionResource(SESSION, "Worksheet",         "Independent practice - multiplying and dividing fractions.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key",        "Teacher reference with worked answers.");
const ENABLING_RES   = makeSessionResource(SESSION, "Enabling Scaffold", "Pre-drawn area models and bar models for multiplying and dividing fractions.");
const EXTENDING_RES  = makeSessionResource(SESSION, "Extension",         "Self-contained investigation - multiplying and dividing mixed numbers.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ───────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Today we are working with fractions - multiplying them and dividing them
- Some of you may have seen these before. If this feels new, that is okay. We will build it together
- By the end of today you will have two strategies you can reach for every time

DO:
- Display title slide as students settle
- Have whiteboards, markers and fraction strips ready
- Distribute the Session 1 Worksheet face-down for now

TEACHER NOTES:
Year 6 lesson. Pitches multiplication and division of simple proper fractions side-by-side so students see how dividing by a fraction connects to multiplying by its reciprocal.

WATCH FOR:
- Students who assume multiplying always makes numbers bigger - flag early, we will address this

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up. Four quick measurement and geometry problems on your whiteboards
- Work through them one at a time. Show your working
- Ask: Which problem is about a 3D shape? [Volume - question 4]

DO:
- Display the four problems
- Allow about 4 minutes total
- Circulate - check that students label units cm, cm squared, and cm cubed correctly

TEACHER NOTES:
Daily Review retrieves prior Term 1 learning on coordinates and measurement. Not today's content. Watch for unit errors and formula confusion.

WATCH FOR:
- Students who mix up area and perimeter - prompt: "Area is the space inside. Perimeter is the distance around."
- Students who forget volume multiplies three dimensions

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Tick and fix. Pen down, highlighter up
- If you got it right, tick. If not, write the correct answer in a different colour
- Fixing it is the learning
- Ask: Which unit did you use for volume? [cubic centimetres, cm cubed]

DO:
- Click to reveal answers
- Students tick-and-fix on their whiteboards
- Scan for students who missed unit labels

TEACHER NOTES:
Tick-and-fix gives immediate feedback. Any student with a persistent unit error needs a quick side conference.

WATCH FOR:
- Students who wrote the right number but no unit - mark it half right, insist on units next time

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Division fluency. 90 seconds
- Work across the row. Short division or mental strategy, your choice
- Ask: What does 'divided by 4' mean? [Split into 4 equal groups, or how many 4s fit]
- Go.

DO:
- Display the six division problems
- Time 90 seconds. Students work on whiteboards
- After time, cold call for answers and one strategy

TEACHER NOTES:
Fluency builds automaticity with short division. Stays close to basic facts and times-tables recall. Separate from DR.

WATCH FOR:
- Students who cannot recall times tables quickly - this will bite when we simplify fractions later
- Students writing the remainder without checking if the division is exact

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me
- Read from slide: We are learning to multiply and divide fractions
- Now the success criteria. Read each one together
- Ask: Which criterion sounds most familiar? [Most will flag SC1]

DO:
- Choral read the LI then each SC
- Brief pause for students to self-assess

TEACHER NOTES:
SC1 is the floor - multiply two simple fractions with support. SC2 is the target - dividing a fraction by a fraction using the reciprocal. SC3 extends to justifying why the reciprocal rule works.

WATCH FOR:
- Students who look anxious at 'divide by a fraction' - reassure: "We will draw it first, then find the shortcut"

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO_MULT = `SAY:
- Watch me multiply two fractions. Two thirds times three quarters
- First I draw a rectangle. I split it vertically into 4 columns - that is my quarters
- Now I split it horizontally into 3 rows - that is my thirds
- I shade 3 of the 4 columns blue - that is three quarters
- I shade 2 of the 3 rows orange - that is two thirds
- Where the two colours overlap is the answer - 6 little squares out of 12 total. Six twelfths. Simplified, one half
- Shortcut: multiply the numerators, multiply the denominators. 2 x 3 = 6. 3 x 4 = 12

DO:
- Point to each stage of the area model on the slide
- Write 2/3 x 3/4 = 6/12 = 1/2 on the board
- Circle the overlap zone

TEACHER NOTES:
The area model makes 'fraction of a fraction' visible. Connecting the picture to the algorithm prevents students from blindly multiplying without understanding.

MISCONCEPTIONS:
- Misconception: When I multiply fractions the answer gets bigger
  Why: Experience with whole-number multiplication
  Impact: Students reject correct answers that are smaller than the factors
  Quick correction: "Two thirds OF three quarters is a part OF a part. A part of something is smaller than the whole thing."

WATCH FOR:
- Students trying to find a common denominator - prompt: "That is for adding. Multiplying is simpler - just multiply across"

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_MULT = `SAY:
- Quick check. On your whiteboards, multiply these two fractions
- One half times three fifths
- 10 seconds. Show me.

DO:
- Students calculate on whiteboards
- Scan for 3/10

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "One half times three fifths. Multiply across. Show me in 10 seconds."
- Scan for: 3/10. Accept 3/10 with working shown.
PROCEED: If 80%+ show 3/10, move to We Do.
PIVOT: Most likely misconception - students write 4/7 (added numerators and denominators). Reteach: "Adding is for same denominators. Multiplying is across. Top times top, bottom times bottom. 1 x 3 = 3, 2 x 5 = 10."

TEACHER NOTES:
First CFU checks whether students execute the multiply-across rule correctly. If most get it, move on.

WATCH FOR:
- Students writing 4/7 (added not multiplied) - this is the classic error
- Students who tried to simplify before multiplying - that is fine, praise it

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with a partner
- Solve three quarters times two fifths on your whiteboards
- Multiply across first, then simplify if you can
- Partner A draws the area model. Partner B writes the calculation. Swap at the end

DO:
- Display the problem
- Partners work together on whiteboards
- Circulate - prompt students who stall
- Allow 90 seconds, then cold call

TEACHER NOTES:
We Do uses a new problem set. Requiring both a drawing and a calculation reinforces the link between the model and the algorithm.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 1 Enabling Scaffold. The area model is pre-drawn as a 4 x 5 grid. Shade 3 columns one colour and 2 rows another colour. Count the overlap.
- Extra Notes: Distribute the Session 1 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Find a second pair of fractions that also multiply to give 6/20. Explain how you know.
- Extra Notes: Accept any equivalent product such as 1/2 x 3/5 x 1/... Students should justify using equivalent fractions.

WATCH FOR:
- Partners who finish fast without simplifying - prompt: "Can you simplify 6/20? What is the highest common factor?"
- Readiness signal: pairs showing 3/10 with a drawn model within 60 seconds

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Check your work. Three quarters times two fifths is 6/20, which simplifies to 3/10
- Ask: How did you simplify? [Divide top and bottom by 2]
- Ask: Could you have simplified before multiplying? [Yes - 2/5 and 3/4 share nothing, but in other problems you can cross-simplify]

DO:
- Reveal the answer
- Cold call two students for the simplification step
- Celebrate thinking, correct errors kindly

TEACHER NOTES:
Revealing confirms the answer. The simplification discussion sets up efficient methods for future lessons.

WATCH FOR:
- Students who wrote 6/20 without simplifying - redirect: "What is the highest common factor of 6 and 20?"

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_IDO_DIV = `SAY:
- Now dividing. Watch me
- Three quarters divided by one half. I ask myself: how many halves fit into three quarters?
- I draw a bar showing three quarters. I draw another bar showing one half underneath
- One half fits ONCE into three quarters, with one quarter left over
- That leftover one quarter is HALF of a half. So the answer is one-and-a-half, which is 3/2
- The shortcut is called keep, change, flip. Keep the first fraction. Change the sign to multiply. Flip the second fraction upside down
- 3/4 divided by 1/2 becomes 3/4 times 2/1. Multiply across: 6/4. Simplify: 3/2 or 1 and a half
- Both the drawing and the shortcut agree

DO:
- Walk through the bar model visually on the slide
- Write the keep-change-flip transformation on the board
- Circle the word 'reciprocal' if it comes up

TEACHER NOTES:
The bar model makes division visible. Students often meet keep-change-flip as a rule with no meaning. Grounding it in 'how many of these fit into that' is the conceptual bridge.

MISCONCEPTIONS:
- Misconception: Flip the FIRST fraction, not the second
  Why: Students apply the rule carelessly
  Impact: Wrong answer, and they cannot check it
  Quick correction: "Keep the first as is. Change the sign. Flip only the second. Say it three times."

WATCH FOR:
- Students who ask "but why does flipping work?" - honour it: "Because dividing by a half is the same as doubling. Dividing by a third is the same as tripling. Flipping makes this work for any fraction."

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_DIV = `SAY:
- Hinge check. Two thirds divided by one quarter
- Use keep, change, flip
- 20 seconds. Show me on my signal.

DO:
- Students calculate on whiteboards
- Scan for 8/3 or 2 and two thirds

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Two thirds divided by one quarter. Keep, change, flip. Show me in 20 seconds."
- Scan for: 8/3 or 2 2/3. Accept either form.
PROCEED: If 80%+ show 8/3, move to You Do.
PIVOT: Most likely misconception - students flip the first fraction instead of the second. Reteach: "Keep the first. Change the sign. Flip only the second. 2/3 stays as 2/3. Sign becomes times. 1/4 flips to 4/1. So 2/3 times 4/1 = 8/3."

TEACHER NOTES:
Hinge question tests whether students apply keep-change-flip correctly. A clean 80%+ means the rule has landed with meaning.

WATCH FOR:
- Students who get 3/8 (flipped the wrong fraction) - this is the most common error
- Students who write 8/3 and stop - that is fine; converting to a mixed number is optional here

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice. Worksheet in front of you
- First: read each problem. Next: choose multiply or divide. Then: solve and simplify
- Circle your answer. Show your working
- If you finish early, try the challenge problems
- 10 minutes. Go.

DO:
- Distribute Session 1 Worksheet
- Students work independently
- Circulate: prioritise students who struggled in the CFUs
- For enabling students, distribute the Session 1 Enabling Scaffold

TEACHER NOTES:
You Do uses different numbers from the We Do. Problems alternate multiplication and division to force students to read before calculating. Challenge problems extend to mixed numbers.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 1 Enabling Scaffold - area models are pre-drawn for multiplication and bar models for division. Shade and count.
- Extra Notes: Distribute the Session 1 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Investigate multiplying and dividing mixed numbers. Convert to improper fractions first, then apply the rule.
- Extra Notes: Distribute the Session 1 Extension PDF. Self-contained.

WATCH FOR:
- Students who skip the simplification step - redirect: "Is your answer in its simplest form?"
- Students who confuse the operation - prompt: "Read the sign again. What are you being asked to do?"
- Readiness signal: 80%+ of students completing 6+ problems correctly within 10 minutes

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Four quick questions in your workbook
- Work on your own, no help
- Three minutes. Go.

DO:
- Display the exit ticket questions
- Students work silently
- Collect or scan responses - sort into three piles: secure, developing, beginning

TEACHER NOTES:
Exit ticket checks SC1 (multiply) and SC2 (divide) directly. Q4 asks for reasoning, which checks SC3.

WATCH FOR:
- Students who finish in under a minute with all correct - they have automated both operations
- Students who confuse multiplication and division - they need targeted reteach next lesson

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Check the success criteria. Thumbs up, sideways, or down
- SC1: I can multiply two simple fractions. Thumbs?
- SC2: I can divide a fraction by a fraction using the reciprocal. Thumbs?
- SC3: I can explain why 'keep, change, flip' works. Thumbs?
- Turn and tell your partner: which was easier for you today - multiplying or dividing?

DO:
- Display success criteria
- Run thumbs check for each SC
- 30 seconds for Turn and Tell
- Cold call 2-3 students

TEACHER NOTES:
Closing brings the lesson full circle. Self-assessment informs planning for the next fractions lesson.

WATCH FOR:
- Students showing thumbs down on SC2 - schedule small-group reteach on keep-change-flip

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- These are today's resources
- Worksheet and answer key are linked

DO:
- Point out each resource and purpose
- Remind students: enabling scaffold is for students who need the visual model
- Extension is for students who finished early

TEACHER NOTES:
Resource slide with clickable links to all companion PDFs.

WATCH FOR:
- N/A - teacher reference only

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ───────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres,
    "Multiplying and Dividing Fractions",
    "Two operations, one strong strategy",
    "Year 6 Numeracy | Session 1",
    NOTES_TITLE);

  // Slides 2-3: Daily Review (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Coordinates, Area, Perimeter and Volume", { y: 0.65, fontSize: 20, color: STAGE_COLORS["1"] });

      // Left card - problems
      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const problems = [
        "1.  Plot (4, 3) on a coordinate grid.",
        "     Which quadrant is it in?",
        "",
        "2.  Area of a rectangle",
        "     8 cm x 5 cm = ?",
        "",
        "3.  Perimeter of a square",
        "     side length 6 cm = ?",
        "",
        "4.  Volume of a cuboid",
        "     3 cm x 4 cm x 5 cm = ?",
      ];
      s.addText(problems.map((p, i) => ({
        text: p,
        options: { fontSize: 13, color: C.CHARCOAL, breakLine: i < problems.length - 1 },
      })), {
        x: 0.75, y: CONTENT_TOP + 0.12, w: 4.1, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right card - instructions / reminders
      addCard(s, 5.3, CONTENT_TOP, 4.2, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboards:", options: { fontSize: 15, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "Show your working clearly.", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Remember your units:", options: { fontSize: 14, bold: true, color: C.ALERT, breakLine: true } },
        { text: "Perimeter: cm (length)", options: { bullet: true, fontSize: 12, color: C.CHARCOAL, breakLine: true } },
        { text: "Area: square cm (cm squared)", options: { bullet: true, fontSize: 12, color: C.CHARCOAL, breakLine: true } },
        { text: "Volume: cubic cm (cm cubed)", options: { bullet: true, fontSize: 12, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Check: does the unit match what you measured?", options: { fontSize: 12, italic: true, color: C.MUTED } },
      ], {
        x: 5.55, y: CONTENT_TOP + 0.12, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) first quadrant   2) 40 cm²   3) 24 cm   4) 60 cm³", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slides 4-5: Fluency - Division (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Fluency");
      addTitle(s, "Division Fluency", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      // Left: problems
      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const divProblems = [
        "1.   72 ÷ 8 =",
        "2.   56 ÷ 7 =",
        "3.   132 ÷ 4 =",
        "4.   156 ÷ 6 =",
        "5.   204 ÷ 3 =",
        "6.   625 ÷ 5 =",
      ];
      s.addText(divProblems.map((p, i) => ({
        text: p,
        options: { fontSize: 16, color: C.CHARCOAL, breakLine: i < divProblems.length - 1, paraSpaceAfter: 10 },
      })), {
        x: 0.9, y: CONTENT_TOP + 0.15, w: 3.9, h: SAFE_BOTTOM - CONTENT_TOP - 0.25,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right: instructions
      addCard(s, 5.3, CONTENT_TOP, 4.2, 2.4, { strip: C.SECONDARY });
      s.addText([
        { text: "Whiteboards - 90 seconds", options: { fontSize: 17, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Choose your strategy:", options: { fontSize: 14, bold: true, color: C.CHARCOAL, breakLine: true } },
        { text: "Times tables recall", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Short division", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Partitioning", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Check: is there a remainder or an exact answer?", options: { fontSize: 12, italic: true, color: C.MUTED } },
      ], {
        x: 5.55, y: CONTENT_TOP + 0.15, w: 3.8, h: 2.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_FLUENCY);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) 9   2) 8   3) 33   4) 26   5) 68   6) 125", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 6: LI/SC
  liSlide(pres,
    ["We are learning to multiply and divide fractions"],
    [
      "I can multiply two simple fractions and simplify the answer",
      "I can divide a fraction by a fraction using the reciprocal",
      "I can explain why 'keep, change, flip' works when dividing",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 7: I Do - Multiplying fractions (area model)
  workedExSlide(pres, 2, "I Do", "Multiplying Fractions: 2/3 x 3/4",
    [
      "Think: two thirds OF three quarters",
      "Draw a rectangle. Split vertically into 4 (quarters).",
      "Split horizontally into 3 (thirds).",
      "Shade 3 columns one colour (three quarters).",
      "Shade 2 rows another colour (two thirds).",
      "The OVERLAP is the answer.",
      "6 little squares out of 12 = 6/12 = 1/2",
      "",
      "Shortcut: multiply across.",
      "    2/3 x 3/4 = (2 x 3) / (3 x 4) = 6/12 = 1/2",
    ],
    NOTES_IDO_MULT, FOOTER,
    (slide, lg) => {
      // Title above grid
      slide.addText("2/3 x 3/4 = ?", {
        x: lg.rightX, y: lg.panelTopPadded, w: lg.rightW, h: 0.34,
        fontSize: 17, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });

      // Area model: 4 cols x 3 rows grid (positioned below the title with clearance)
      const gridX = lg.rightX + 0.3;
      const gridY = lg.panelTopPadded + 0.52;
      const gridW = 3.2;
      const gridH = 2.25;
      const cols = 4, rows = 3;
      const cellW = gridW / cols;
      const cellH = gridH / rows;

      // Draw cells - shade overlap (first 3 cols, first 2 rows) in ACCENT, rest show lighter shading
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const isCol = c < 3;       // three quarters (blue)
          const isRow = r < 2;       // two thirds (orange)
          const isOverlap = isCol && isRow;

          let fill;
          if (isOverlap) fill = C.ACCENT;
          else if (isCol) fill = "BFD4F5";   // light blue for quarters only
          else if (isRow) fill = "FCE3C4";   // light orange for thirds only
          else fill = C.WHITE;

          slide.addShape("rect", {
            x: gridX + c * cellW, y: gridY + r * cellH, w: cellW, h: cellH,
            fill: { color: fill },
            line: { color: C.CHARCOAL, width: 1 },
          });
        }
      }

      // Legend below grid
      const legendY = gridY + gridH + 0.12;
      slide.addShape("rect", { x: gridX, y: legendY, w: 0.2, h: 0.2, fill: { color: C.ACCENT } });
      slide.addText("overlap = 6/12 = 1/2", {
        x: gridX + 0.25, y: legendY - 0.02, w: 2.2, h: 0.25,
        fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0, valign: "middle",
      });
    }
  );

  // Slides 8-9: CFU - Multiplying (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Quick Check: Multiplying", "Show Me Boards",
      "On your whiteboard:\n\n1/2 x 3/5 = ?\n\nMultiply across. Simplify if you can.",
      NOTES_CFU_MULT, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Answer: 1/2 x 3/5 = 3/10", {
        x: 1.5, y: 4.3, w: 7, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 10-11: We Do - Multiplying (withReveal)
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "We Do: 3/4 x 2/5",
      [
        "With a partner, solve 3/4 x 2/5",
        "",
        "Partner A: draw the area model",
        "Partner B: do the calculation",
        "Swap and check each other.",
        "",
        "Remember: multiply across, then simplify.",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.7, { strip: C.SECONDARY });
        slide.addText("Steps", {
          x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.3,
          fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        const steps = [
          { label: "1", desc: "Multiply numerators", color: C.PRIMARY },
          { label: "2", desc: "Multiply denominators", color: C.PRIMARY },
          { label: "3", desc: "Simplify the fraction", color: C.ACCENT },
        ];
        steps.forEach((step, i) => {
          const sy = lg.panelTopPadded + 0.5 + i * 0.55;
          addTextOnShape(slide, step.label, {
            x: lg.rightX + 0.25, y: sy, w: 0.42, h: 0.42, rectRadius: 0.21,
            fill: { color: step.color },
          }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
          slide.addText(step.desc, {
            x: lg.rightX + 0.85, y: sy, w: lg.rightW - 1.0, h: 0.42,
            fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0, valign: "middle",
          });
        });
        // Example hint at bottom
        slide.addText("Hint: 3 x 2 over 4 x 5", {
          x: lg.rightX + 0.25, y: lg.panelTopPadded + 2.25, w: lg.rightW - 0.4, h: 0.3,
          fontSize: 13, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
        });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "3/4 x 2/5 = 6/20 = 3/10  (divide top and bottom by 2)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 12: I Do - Dividing fractions (bar model + keep/change/flip)
  workedExSlide(pres, 2, "I Do", "Dividing Fractions: 3/4 ÷ 1/2",
    [
      "Ask: how many halves fit into three quarters?",
      "Draw the 3/4 bar.",
      "Mark where 1/2 sits.",
      "One half fits ONCE, with 1/4 left over.",
      "1/4 is HALF of a half - so another 0.5 halves fit.",
      "Answer: 1 and a half, or 3/2.",
      "",
      "Shortcut: KEEP, CHANGE, FLIP",
      "    3/4 ÷ 1/2  ->  3/4 x 2/1 = 6/4 = 3/2",
      "",
      "Same answer. The shortcut works because dividing",
      "by a fraction = multiplying by its reciprocal.",
    ],
    NOTES_IDO_DIV, FOOTER,
    (slide, lg) => {
      // Bar model visual: labels placed to the LEFT of each bar to avoid title overlap.
      const labelW = 0.5;
      const barX = lg.rightX + labelW + 0.1;
      const barW = lg.rightW - labelW - 0.25;
      const barH = 0.5;

      // Title
      slide.addText("How many 1/2s fit into 3/4?", {
        x: lg.rightX, y: lg.panelTopPadded - 0.05, w: lg.rightW, h: 0.32,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });

      // Three-quarters bar (4 segments, first 3 shaded). Label to the left of the bar.
      const bar1Y = lg.panelTopPadded + 0.6;
      slide.addText("3/4", {
        x: lg.rightX + 0.05, y: bar1Y + (barH - 0.28) / 2, w: labelW, h: 0.28,
        fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0, valign: "middle",
      });
      for (let i = 0; i < 4; i++) {
        slide.addShape("rect", {
          x: barX + (i * barW / 4), y: bar1Y, w: barW / 4, h: barH,
          fill: { color: i < 3 ? C.PRIMARY : C.WHITE },
          line: { color: C.CHARCOAL, width: 1 },
        });
      }

      // Half bar (2 segments, first shaded) - same total length.
      const bar2Y = bar1Y + barH + 0.45;
      slide.addText("1/2", {
        x: lg.rightX + 0.05, y: bar2Y + (barH - 0.28) / 2, w: labelW, h: 0.28,
        fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0, valign: "middle",
      });
      for (let i = 0; i < 2; i++) {
        slide.addShape("rect", {
          x: barX + (i * barW / 2), y: bar2Y, w: barW / 2, h: barH,
          fill: { color: i < 1 ? C.SECONDARY : C.WHITE },
          line: { color: C.CHARCOAL, width: 1 },
        });
      }

      // Answer callout
      addTextOnShape(slide, "1 and a half halves fit = 3/2", {
        x: lg.rightX + 0.1, y: bar2Y + barH + 0.3, w: lg.rightW - 0.2, h: 0.42, rectRadius: 0.08,
        fill: { color: C.ACCENT },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slides 13-14: CFU - Dividing (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Check: Dividing", "Show Me Boards",
      "On your whiteboard:\n\n2/3 ÷ 1/4 = ?\n\nUse KEEP, CHANGE, FLIP.",
      NOTES_CFU_DIV, FOOTER),
    (slide) => {
      addTextOnShape(slide, "2/3 ÷ 1/4 = 2/3 x 4/1 = 8/3 (or 2 and 2/3)", {
        x: 0.8, y: 4.3, w: 8.4, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 15: You Do
  workedExSlide(pres, 4, "You Do", "You Do: Mixed Practice",
    [
      "First:  read each problem. Multiply or divide?",
      "Next:   apply the correct rule.",
      "            x  -> multiply across",
      "            ÷  -> keep, change, flip",
      "Then:   simplify if you can.",
      "",
      "Work on your own. Show your working.",
      "If you finish early, try the challenge section.",
      "",
      "You have 10 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.3, { strip: C.ALERT });
      slide.addText("Remember", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 16, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Multiplying: multiply across", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Dividing: keep, change, flip", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Simplify at the end", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Check: does your answer make sense?", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.5, w: lg.rightW - 0.5, h: 1.5,
        fontFace: FONT_B, margin: 0, valign: "top", paraSpaceAfter: 4,
      });
      // Small worked example box
      addCard(slide, lg.rightX + 0.2, lg.panelTopPadded + 2.1, lg.rightW - 0.4, 1.0, { strip: C.PRIMARY, fill: C.BG_LIGHT });
      slide.addText("Example:", {
        x: lg.rightX + 0.35, y: lg.panelTopPadded + 2.18, w: lg.rightW - 0.7, h: 0.25,
        fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
      });
      slide.addText("1/3 ÷ 2/5 = 1/3 x 5/2 = 5/6", {
        x: lg.rightX + 0.35, y: lg.panelTopPadded + 2.45, w: lg.rightW - 0.7, h: 0.45,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });
    }
  );

  // Slide 16: Exit Ticket
  exitTicketSlide(pres,
    [
      "2/5 x 1/2 = ?",
      "1/2 ÷ 1/3 = ?",
      "Simplify 8/12 to its lowest terms.",
      "Explain in one sentence why dividing by 1/4 gives the same answer as multiplying by 4.",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 17: Closing
  closingSlide(pres,
    "Which was easier for you today - multiplying fractions or dividing fractions? Tell your partner why.",
    [
      "I can multiply two simple fractions and simplify",
      "I can divide a fraction by a fraction using the reciprocal",
      "I can explain why 'keep, change, flip' works",
    ],
    NOTES_CLOSING);

  // Slide 18: Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Fractions_Multiplying_Dividing.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Generate PDFs ─────────────────────────────────────────────────────────

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Multiplying and Dividing Fractions",
      color: C.NAVY,
      lessonInfo: "Session 1 | Year 6 Numeracy",
    });
    y = addTipBox(doc,
      "For each problem: work out the answer and simplify. Show your working next to each problem.",
      y, { color: C.TEAL });

    y = addSectionHeading(doc, "Multiplying Fractions", y, { color: C.NAVY });
    y = addProblem(doc, 1, "1/3 x 1/2 = ____", y, { color: C.NAVY });
    y = addProblem(doc, 2, "2/5 x 3/4 = ____", y, { color: C.NAVY });
    y = addProblem(doc, 3, "3/7 x 2/3 = ____", y, { color: C.NAVY });
    y = addProblem(doc, 4, "4/5 x 1/2 = ____  (simplify)", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Dividing Fractions", y, { color: C.NAVY });
    y = addProblem(doc, 5, "1/2 ÷ 1/4 = ____", y, { color: C.NAVY });
    y = addProblem(doc, 6, "2/3 ÷ 1/6 = ____", y, { color: C.NAVY });
    y = addProblem(doc, 7, "3/5 ÷ 2/5 = ____", y, { color: C.NAVY });
    y = addProblem(doc, 8, "5/6 ÷ 1/3 = ____", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Mixed Practice", y, { color: C.NAVY });
    y = addProblem(doc, 9,  "1/4 x 3/8 = ____", y, { color: C.NAVY });
    y = addProblem(doc, 10, "7/8 ÷ 1/2 = ____", y, { color: C.NAVY });
    y = addProblem(doc, 11, "3/10 x 2/3 = ____", y, { color: C.NAVY });
    y = addProblem(doc, 12, "4/9 ÷ 2/3 = ____", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addProblem(doc, 13,
      "A recipe uses 3/4 cup of flour. You want to make half the recipe. How much flour do you need? Show your working.",
      y, { color: C.NAVY, writeLines: [{ label: "" }] });

    addPdfFooter(doc, "Session 1 | Multiplying and Dividing Fractions | Year 6");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Multiplying and Dividing Fractions",
      color: C.NAVY,
      lessonInfo: "Session 1 | Year 6 Numeracy",
    });

    y = addSectionHeading(doc, "Multiplying Fractions", y, { color: C.NAVY });
    y = addBodyText(doc, "1.  1/3 x 1/2 = 1/6", y);
    y = addBodyText(doc, "2.  2/5 x 3/4 = 6/20 = 3/10", y);
    y = addBodyText(doc, "3.  3/7 x 2/3 = 6/21 = 2/7", y);
    y = addBodyText(doc, "4.  4/5 x 1/2 = 4/10 = 2/5", y);

    y = addSectionHeading(doc, "Dividing Fractions", y, { color: C.NAVY });
    y = addBodyText(doc, "5.  1/2 ÷ 1/4 = 1/2 x 4/1 = 4/2 = 2", y);
    y = addBodyText(doc, "6.  2/3 ÷ 1/6 = 2/3 x 6/1 = 12/3 = 4", y);
    y = addBodyText(doc, "7.  3/5 ÷ 2/5 = 3/5 x 5/2 = 15/10 = 3/2 or 1 1/2", y);
    y = addBodyText(doc, "8.  5/6 ÷ 1/3 = 5/6 x 3/1 = 15/6 = 5/2 or 2 1/2", y);

    y = addSectionHeading(doc, "Mixed Practice", y, { color: C.NAVY });
    y = addBodyText(doc, "9.   1/4 x 3/8 = 3/32", y);
    y = addBodyText(doc, "10.  7/8 ÷ 1/2 = 7/8 x 2/1 = 14/8 = 7/4 or 1 3/4", y);
    y = addBodyText(doc, "11.  3/10 x 2/3 = 6/30 = 1/5", y);
    y = addBodyText(doc, "12.  4/9 ÷ 2/3 = 4/9 x 3/2 = 12/18 = 2/3", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "13.  Half of 3/4 = 1/2 x 3/4 = 3/8 cup of flour.", y);

    addPdfFooter(doc, "Session 1 | Answer Key | Year 6");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold - pre-drawn area models and bar models
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Visual Models for Multiplying and Dividing Fractions",
      color: C.TEAL,
      lessonInfo: "Session 1 | Year 6 Numeracy",
    });
    y = addTipBox(doc,
      "The models are drawn for you. Shade and count. Then write your answer.",
      y, { color: C.TEAL });

    // ─── Multiplying section ───
    y = addSectionHeading(doc, "Multiplying: Area Models", y, { color: C.NAVY });
    y = addBodyText(doc, "Problem 1:  1/2 x 1/3", y);
    y = addBodyText(doc, "Shade 1 of 2 columns blue. Shade 1 of 3 rows orange. Count the overlap.", y);

    // Draw 2x3 area model (2 cols, 3 rows)
    y += 4;
    {
      const gx = 70, gy = y, cw = 55, ch = 35;
      for (let r = 0; r < 3; r++) for (let c = 0; c < 2; c++) {
        doc.rect(gx + c * cw, gy + r * ch, cw, ch).lineWidth(0.6).stroke("#666666");
      }
      doc.fontSize(9).fillColor("#666666")
        .text("2 columns x 3 rows = 6 total parts", gx + 2 * cw + 15, gy + 10);
      y = gy + 3 * ch + 8;
    }
    y = addBodyText(doc, "Answer:  1/2 x 1/3 = ____/____ = ____", y);

    y += 8;
    y = addBodyText(doc, "Problem 2:  2/3 x 1/4", y);
    y = addBodyText(doc, "Shade 1 of 4 columns blue. Shade 2 of 3 rows orange. Count the overlap.", y);

    // Draw 4x3 area model
    y += 4;
    {
      const gx = 70, gy = y, cw = 35, ch = 35;
      for (let r = 0; r < 3; r++) for (let c = 0; c < 4; c++) {
        doc.rect(gx + c * cw, gy + r * ch, cw, ch).lineWidth(0.6).stroke("#666666");
      }
      doc.fontSize(9).fillColor("#666666")
        .text("4 columns x 3 rows = 12 total parts", gx + 4 * cw + 15, gy + 10);
      y = gy + 3 * ch + 8;
    }
    y = addBodyText(doc, "Answer:  2/3 x 1/4 = ____/____ = ____ (simplified)", y);

    // ─── Dividing section ───
    y += 10;
    if (y > 640) { doc.addPage(); y = 60; }
    y = addSectionHeading(doc, "Dividing: Bar Models", y, { color: C.NAVY });
    y = addBodyText(doc, "Problem 3:  3/4 ÷ 1/4  - How many quarters fit into 3/4?", y);

    y += 4;
    {
      const gx = 70, gy = y, w = 240, h = 26;
      doc.fontSize(9).fillColor("#333333").text("3/4", gx - 20, gy + 8);
      for (let i = 0; i < 4; i++) {
        doc.rect(gx + i * (w / 4), gy, w / 4, h).lineWidth(0.6)
          .fillAndStroke(i < 3 ? "#C8DDF8" : "#FFFFFF", "#666666");
      }
      doc.fillColor("#333333").fontSize(9)
        .text("Count the 1/4 blocks shaded.", gx + w + 12, gy + 8);
      y = gy + h + 12;
    }
    y = addBodyText(doc, "Answer:  3/4 ÷ 1/4 = ____", y);

    y += 8;
    y = addBodyText(doc, "Problem 4:  1/2 ÷ 1/6  - How many sixths fit into 1/2?", y);

    y += 4;
    {
      const gx = 70, gy = y, w = 240, h = 26;
      doc.fontSize(9).fillColor("#333333").text("1/2", gx - 20, gy + 8);
      for (let i = 0; i < 6; i++) {
        doc.rect(gx + i * (w / 6), gy, w / 6, h).lineWidth(0.6)
          .fillAndStroke(i < 3 ? "#C8DDF8" : "#FFFFFF", "#666666");
      }
      doc.fillColor("#333333").fontSize(9)
        .text("Count the 1/6 blocks shaded.", gx + w + 12, gy + 8);
      y = gy + h + 12;
    }
    y = addBodyText(doc, "Answer:  1/2 ÷ 1/6 = ____", y);

    y += 8;
    y = addBodyText(doc, "Check with keep, change, flip:  1/2 ÷ 1/6 = 1/2 x 6/1 = 6/2 = ____", y);

    addPdfFooter(doc, "Session 1 | Enabling Scaffold | Year 6");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Extension - Mixed Numbers
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Mixed Numbers Investigation",
      color: C.NAVY,
      lessonInfo: "Session 1 | Year 6 Numeracy",
    });

    y = addSectionHeading(doc, "Mixed Numbers Recap", y, { color: C.NAVY });
    y = addBodyText(doc, "A mixed number has a whole part and a fraction part. Example: 2 1/3.", y);
    y = addBodyText(doc, "To multiply or divide with mixed numbers, first turn them into improper fractions.", y);
    y = addBodyText(doc, "Rule: whole number x denominator + numerator, over the denominator.", y);
    y = addBodyText(doc, "Example: 2 1/3 = (2 x 3 + 1)/3 = 7/3", y);

    y = addSectionHeading(doc, "Worked Example", y, { color: C.NAVY });
    y = addBodyText(doc, "Calculate 1 1/2 x 2 1/3", y);
    y = addBodyText(doc, "Step 1: convert. 1 1/2 = 3/2. 2 1/3 = 7/3.", y);
    y = addBodyText(doc, "Step 2: multiply. 3/2 x 7/3 = 21/6.", y);
    y = addBodyText(doc, "Step 3: simplify. 21/6 = 7/2 = 3 1/2.", y);

    y = addSectionHeading(doc, "Your Investigation", y, { color: C.NAVY });
    y = addBodyText(doc, "Convert each mixed number, then solve. Show all steps.", y);

    y = addProblem(doc, 1, "1 1/4 x 2/3 = ____", y, { color: C.NAVY });
    y = addProblem(doc, 2, "2 1/2 x 1 1/5 = ____", y, { color: C.NAVY });
    y = addProblem(doc, 3, "3 1/3 ÷ 5/6 = ____", y, { color: C.NAVY });
    y = addProblem(doc, 4, "4 1/2 ÷ 1 1/2 = ____", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Think About It", y, { color: C.NAVY });
    y = addBodyText(doc, "When you multiply two numbers both bigger than 1, what happens to the answer? Is it always bigger than both factors?", y);
    y = addWriteLine(doc, "I noticed:", y);
    y = addWriteLine(doc, "", y);

    addPdfFooter(doc, "Session 1 | Extension | Year 6");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 1 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
