"use strict";

// Fractions - Session 1: Adding and Subtracting Fractions
// Year 6 Numeracy, Variant 2
// DR: Coordinates and area, perimeter, volume
// Fluency: Division

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");

const T = createTheme("numeracy", "grade56", 2);
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const FOOTER = "Adding and Subtracting Fractions | Year 6 Numeracy";
const OUT_DIR = "output/Fractions_Adding_Subtracting";

function addFractionStrip(slide, x, y, w, h, numerator, denominator, opts) {
  const o = opts || {};
  const fillColor = o.fillColor || C.PRIMARY;
  const label = String(o.label || (String(numerator) + "/" + String(denominator)));
  const labelW = o.labelW || 0.7;
  const segW = w / denominator;

  slide.addText(label, {
    x,
    y: y + (h - 0.26) / 2,
    w: labelW,
    h: 0.26,
    fontSize: o.labelFontSize || 11,
    fontFace: FONT_B,
    color: o.labelColor || C.CHARCOAL,
    bold: Boolean(o.labelBold),
    align: "left",
    valign: "middle",
    margin: 0,
  });

  for (let i = 0; i < denominator; i++) {
    slide.addShape("rect", {
      x: x + labelW + i * segW,
      y,
      w: segW,
      h,
      fill: { color: i < numerator ? fillColor : C.WHITE },
      line: { color: o.lineColor || C.CHARCOAL, width: 1 },
    });
  }
}

function addCrossOutSegments(slide, x, y, w, h, denominator, startIndex, count, opts) {
  const o = opts || {};
  const labelW = o.labelW || 0.7;
  const segW = w / denominator;
  const color = o.color || C.ALERT;

  for (let i = startIndex; i < startIndex + count; i++) {
    const segX = x + labelW + i * segW;
    slide.addShape("line", {
      x: segX + 0.02,
      y: y + 0.02,
      w: segW - 0.04,
      h: h - 0.04,
      line: { color, width: 2.2 },
    });
    slide.addShape("line", {
      x: segX + 0.02,
      y: y + h - 0.02,
      w: segW - 0.04,
      h: -(h - 0.04),
      line: { color, width: 2.2 },
    });
  }
}

function addPracticeCard(slide, x, y, w, h, title, problem, color) {
  addCard(slide, x, y, w, h, { strip: color, fill: C.WHITE });
  slide.addText(String(title), {
    x: x + 0.18,
    y: y + 0.1,
    w: w - 0.36,
    h: 0.24,
    fontSize: 11,
    fontFace: FONT_B,
    color,
    bold: true,
    margin: 0,
  });
  slide.addText(String(problem), {
    x: x + 0.18,
    y: y + 0.38,
    w: w - 0.36,
    h: h - 0.48,
    fontSize: 18,
    fontFace: FONT_H,
    color: C.CHARCOAL,
    bold: true,
    align: "center",
    valign: "middle",
    margin: 0,
  });
}

const NOTES_TITLE = `SAY:
- Today we are adding and subtracting fractions
- The key move is to slow down and make the parts the same size first
- If this feels new, that is okay. We will build it together

DO:
- Display title slide as students settle
- Have mini-whiteboards, markers and fraction strips ready
- Draw one blank whole bar on the board ready for the first model

TEACHER NOTES:
Year 6 lesson. The strategy focus is common denominators with related fractions. Visual fraction strips stay visible before students move to symbols alone.

WATCH FOR:
- Students who want to add the denominators straight away
- Students who say thirds and sixths are already the same-sized parts

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up. Four quick review problems on your whiteboards
- Work through them one at a time and show your thinking
- Ask: Which one is the 3D measurement question? [Volume - question 4]

DO:
- Display the four problems
- Allow about 4 minutes total
- Circulate and check that students include units

TEACHER NOTES:
Daily Review retrieves prior learning on coordinates and measurement. It is not today's new content.

WATCH FOR:
- Students who mix up area and perimeter
- Students who forget that volume uses cubic units

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Tick and fix. Pen down, marker up
- If you got it right, tick it. If not, fix it in a new colour
- Ask: Which unit belongs with volume? [cubic centimetres, cm cubed]

DO:
- Click to reveal answers
- Students tick and fix on their whiteboards
- Pause briefly on any unit errors

TEACHER NOTES:
Immediate feedback keeps the review brisk. Any student with repeated unit errors needs a quick conference before the lesson moves on.

WATCH FOR:
- Students who wrote the right number but no unit

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Division fluency. Ninety seconds
- Work across the row. Mental strategy or short division, your choice
- Ask: Why does division matter for fractions? [It helps us simplify]
- Go

DO:
- Display the six division facts
- Time 90 seconds on whiteboards
- Cold call for both answers and strategies

TEACHER NOTES:
Division fluency supports simplification later in the lesson. This routine is separate from the Daily Review and stays in Number.

WATCH FOR:
- Students who hesitate on times-table facts
- Students who do not notice when an answer should be exact

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me
- Read from slide: We are learning to add and subtract fractions by making the parts the same size
- Now read each success criterion together
- Ask: Which criterion sounds like the first step? [SC1]

DO:
- Choral read the LI and each SC
- Brief pause for students to self-assess with fingers 1 to 3

TEACHER NOTES:
SC1 is the access point. SC2 is the core target for the lesson. SC3 pushes explanation of why the denominator stays the same once the parts match.

WATCH FOR:
- Students who think SC2 means only addition or only subtraction

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch the strategy, not just the answer
- Example one: one third plus one sixth. The parts are not the same size yet
- I rename one third as two sixths. Now I can add two sixths and one sixth to make three sixths, or one half
- Example two: five sixths minus one third. I rename one third as two sixths. Now I subtract five sixths minus two sixths to get three sixths, or one half
- The denominator tells me the size of the parts. Once the parts match, I keep that denominator

DO:
- Point to each strip as you rename the fraction
- Write the rename step on the board beside the slide
- Circle the denominator after the rename to emphasise same-sized parts

TEACHER NOTES:
This I Do models one addition and one subtraction example using the same common-denominator routine. The visual keeps the meaning of the denominator clear.

MISCONCEPTIONS:
- Misconception: Add or subtract the denominator as well
  Why: Students treat the fraction like two separate whole numbers
  Impact: They lose the size of the part and produce impossible answers
  Quick correction: "The denominator names the size of the part. If the parts stay sixths, the answer is still in sixths."

WATCH FOR:
- Students who rename one third as one sixth
- Students who combine numerators before the denominators match

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU_ADD = `SAY:
- Quick check. Rename first, then add
- One quarter plus two eighths
- Show the rename and the answer on your whiteboard

DO:
- Students solve on whiteboards
- Scan for one quarter renamed as two eighths, then four eighths, then one half

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "One quarter plus two eighths. Rename first. Show me the renamed fraction and the answer."
- Scan for: 1/4 = 2/8, then 2/8 + 2/8 = 4/8 = 1/2.
PROCEED: If 80% or more show one half with a correct rename, move to the subtraction board build.
PIVOT: Most likely misconception - students write 3/12 by adding both numerators and denominators. Reteach with strips: "Quarter means bigger parts than eighths. We must rename to the same-sized parts before we combine."

TEACHER NOTES:
This CFU checks whether students can execute the rename step independently before adding.

WATCH FOR:
- Students who stop at four eighths and do not simplify
- Students who rename two eighths instead of one quarter

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Build this one with me on the board
- We have five sixths and we need to take away one third
- Ask: Can I subtract sixths and thirds yet? [No - the parts are different sizes]
- Rename one third as two sixths, then cross out two sixths from the strip

DO:
- Students copy the strip into their workbook or onto whiteboards
- Teacher crosses out two sixths on the board
- Cold call after think time before revealing the answer

TEACHER NOTES:
This is the board-build moment. Students see subtraction as removing same-sized parts from the strip, not as a memorised trick.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use physical fraction strips. Match one third to sixth pieces first, then remove two sixth pieces from five sixths.
- Extra Notes: Keep students on the concrete materials for the full problem.
EXTENDING PROMPT:
- Task: Create a different subtraction problem that also equals one half.
- Extra Notes: Students should prove it with either strips or an equation.

WATCH FOR:
- Students who try to subtract five minus one before renaming
- Readiness signal: students can say "one third is two sixths" without prompting

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Check your work. One third is two sixths
- So five sixths minus two sixths equals three sixths, or one half
- Ask: Why did the denominator stay as sixths? [Because the parts were sixths the whole time]

DO:
- Reveal the crossed-out sixths and final answer
- Ask one student to explain the rename step and another to explain the subtraction

TEACHER NOTES:
The reveal makes the subtraction visible. The explanation question checks whether students are attending to the part size, not only the final number.

WATCH FOR:
- Students who say the answer is three because they are ignoring the denominator

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU_SUB = `SAY:
- Hinge check. Three quarters minus one eighth
- Rename first, then subtract
- Twenty seconds. Show me

DO:
- Students solve on whiteboards
- Scan for six eighths minus one eighth equals five eighths

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Three quarters minus one eighth. Rename first. Show me in 20 seconds."
- Scan for: 3/4 = 6/8, then 6/8 - 1/8 = 5/8.
PROCEED: If 80% or more show five eighths, move to independent practice.
PIVOT: Most likely misconception - students write 2/12 or 2/8 because they subtract before renaming. Reteach with the strip: "Quarters are bigger pieces than eighths. Rename three quarters as six eighths first, then subtract one eighth."

TEACHER NOTES:
This hinge question checks whether students can transfer the same strategy to subtraction without teacher support.

WATCH FOR:
- Students who rename one eighth instead of three quarters
- Students who get five eighths but cannot show the rename

[Stage 3: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice in your workbook
- First: choose a common denominator
- Next: rename the fractions
- Then: add or subtract and simplify
- If you get stuck, sketch a strip

DO:
- Students complete the four problems in their workbook
- Circulate and prioritise students who missed a CFU
- Prompt students to show the rename step, not only the final answer

TEACHER NOTES:
You Do changes the numbers and mixes operations so students must choose the strategy independently. No printed sheet is needed here.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Solve only the first two problems using physical fraction strips before writing the equations.
- Extra Notes: Keep the model beside the student while they record the symbols.
EXTENDING PROMPT:
- Task: Write one addition and one subtraction problem that both equal one whole.
- Extra Notes: Students should solve and justify each one.

WATCH FOR:
- Students who pick a denominator that is not a common multiple
- Students who simplify incorrectly at the end
- Readiness signal: students record the rename step automatically

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Four quick checks in your workbook
- Work on your own and show the rename step when you need it
- Three minutes. Go

DO:
- Display the exit ticket questions
- Students work silently
- Collect or scan responses and sort into secure, developing and beginning

TEACHER NOTES:
The exit ticket checks SC1, SC2 and SC3. The two calculation items assess the target directly.

WATCH FOR:
- Students who can calculate but cannot explain why the denominator stays the same
- Students who still avoid renaming before they start

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Check the success criteria. Thumbs up, sideways or down
- SC1: I can rename simple fractions to a common denominator. Thumbs?
- SC2: I can add or subtract related fractions and simplify. Thumbs?
- SC3: I can explain why the denominator stays the same once the parts match. Thumbs?
- Turn and tell your partner: which step helps you most - rename, combine or simplify?

DO:
- Display the success criteria
- Run the thumbs check for each one
- Cold call two or three students after turn and talk

TEACHER NOTES:
Closing returns to the lesson language and gives the teacher a quick read on which part of the process still needs support.

WATCH FOR:
- Students showing thumbs down on SC1 or SC3 - they need more concrete modelling next lesson

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- No printed student resources are needed for this lesson
- We are using whiteboards, workbooks and fraction strips instead

DO:
- Point out the materials list
- Check that fraction strips or teacher-drawn bars are ready before the lesson starts

TEACHER NOTES:
Teacher reference slide. The lesson is intentionally workbook and manipulative based, so there are no companion handouts.

WATCH FOR:
- N/A - teacher reference only

[General: Resources | VTLM 2.0: Planning]`;

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(
    pres,
    "Adding and Subtracting Fractions",
    "Make the parts the same size first",
    "Year 6 Numeracy | Session 1",
    NOTES_TITLE
  );

  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Coordinates, Area, Perimeter and Volume", {
        y: 0.65,
        fontSize: 20,
        color: STAGE_COLORS["1"],
      });

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
        text: String(p),
        options: { fontSize: 13, color: C.CHARCOAL, breakLine: i < problems.length - 1 },
      })), {
        x: 0.75,
        y: CONTENT_TOP + 0.12,
        w: 4.1,
        h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
        fontFace: FONT_B,
        margin: 0,
        valign: "top",
      });

      addCard(s, 5.3, CONTENT_TOP, 4.2, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboards:", options: { fontSize: 15, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "Show your working clearly.", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Remember your units:", options: { fontSize: 14, bold: true, color: C.ALERT, breakLine: true } },
        { text: "Perimeter: cm", options: { bullet: true, fontSize: 12, color: C.CHARCOAL, breakLine: true } },
        { text: "Area: square cm", options: { bullet: true, fontSize: 12, color: C.CHARCOAL, breakLine: true } },
        { text: "Volume: cubic cm", options: { bullet: true, fontSize: 12, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Check: does the unit match the measurement?", options: { fontSize: 12, italic: true, color: C.MUTED } },
      ], {
        x: 5.55,
        y: CONTENT_TOP + 0.12,
        w: 3.8,
        h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
        fontFace: FONT_B,
        margin: 0,
        valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      runSlideDiagnostics(s, pres, { respectSafeBottom: false });
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) first quadrant   2) 40 cm squared   3) 24 cm   4) 60 cm cubed", {
        x: 0.5,
        y: 4.55,
        w: 9,
        h: 0.5,
        rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, {
        fontSize: 15,
        fontFace: FONT_H,
        color: C.WHITE,
        bold: true,
      });
      slide.addNotes(NOTES_DR_A);
    }
  );

  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Fluency");
      addTitle(s, "Division Fluency", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.5, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const divProblems = [
        "1.   84 / 7 =",
        "2.   96 / 8 =",
        "3.   144 / 12 =",
        "4.   156 / 4 =",
        "5.   216 / 6 =",
        "6.   315 / 9 =",
      ];
      s.addText(divProblems.map((p, i) => ({
        text: String(p),
        options: { fontSize: 16, color: C.CHARCOAL, breakLine: i < divProblems.length - 1, paraSpaceAfter: 10 },
      })), {
        x: 0.9,
        y: CONTENT_TOP + 0.15,
        w: 3.9,
        h: SAFE_BOTTOM - CONTENT_TOP - 0.25,
        fontFace: FONT_B,
        margin: 0,
        valign: "top",
      });

      addCard(s, 5.3, CONTENT_TOP, 4.2, 2.45, { strip: C.SECONDARY });
      s.addText([
        { text: "Whiteboards - 90 seconds", options: { fontSize: 17, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Choose your strategy:", options: { fontSize: 14, bold: true, color: C.CHARCOAL, breakLine: true } },
        { text: "Times tables recall", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Short division", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Partitioning", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Check: can you use this fact to simplify a fraction later?", options: { fontSize: 12, italic: true, color: C.MUTED } },
      ], {
        x: 5.55,
        y: CONTENT_TOP + 0.15,
        w: 3.8,
        h: 2.25,
        fontFace: FONT_B,
        margin: 0,
        valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_FLUENCY);
      runSlideDiagnostics(s, pres, { respectSafeBottom: false });
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "1) 12   2) 12   3) 12   4) 39   5) 36   6) 35", {
        x: 0.5,
        y: 4.55,
        w: 9,
        h: 0.5,
        rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, {
        fontSize: 16,
        fontFace: FONT_H,
        color: C.WHITE,
        bold: true,
      });
    }
  );

  liSlide(
    pres,
    ["We are learning to add and subtract fractions by making the parts the same size"],
    [
      "I can rename simple fractions to a common denominator",
      "I can add or subtract related fractions and simplify the answer",
      "I can explain why the denominator stays the same once the parts match",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  workedExSlide(
    pres,
    2,
    "I Do",
    "Same-Size Parts First",
    [
      "Look at the denominators. Are the parts the same size?",
      "If not, rename to a common denominator.",
      "Add or subtract the numerators only.",
      "Keep the denominator because the parts match.",
      "Simplify if you can.",
      "Reminder: 1/3 = 2/6.",
      "Reminder: 1/4 = 2/8.",
    ],
    NOTES_IDO,
    FOOTER,
    (slide, lg) => {
      const cardX = lg.rightX;
      const cardW = lg.rightW;
      const topY = lg.panelTopPadded;
      const miniH = 1.6;
      const gap = 0.18;

      addCard(slide, cardX, topY, cardW, miniH, { strip: C.PRIMARY, fill: C.WHITE });
      slide.addText("Add: 1/3 + 1/6", {
        x: cardX + 0.18,
        y: topY + 0.08,
        w: cardW - 0.36,
        h: 0.24,
        fontSize: 13,
        fontFace: FONT_H,
        color: C.PRIMARY,
        bold: true,
        margin: 0,
      });
      slide.addText("Rename 1/3 as 2/6 first", {
        x: cardX + 0.18,
        y: topY + 0.34,
        w: cardW - 0.36,
        h: 0.2,
        fontSize: 10.5,
        fontFace: FONT_B,
        color: C.MUTED,
        italic: true,
        margin: 0,
      });
      addFractionStrip(slide, cardX + 0.22, topY + 0.58, 2.55, 0.16, 2, 6, {
        label: "2/6",
        fillColor: C.PRIMARY,
      });
      addFractionStrip(slide, cardX + 0.22, topY + 0.84, 2.55, 0.16, 1, 6, {
        label: "1/6",
        fillColor: C.SECONDARY,
      });
      addFractionStrip(slide, cardX + 0.22, topY + 1.1, 2.55, 0.16, 3, 6, {
        label: "3/6",
        fillColor: C.ACCENT,
        labelBold: true,
      });
      addTextOnShape(slide, "2/6 + 1/6 = 3/6 = 1/2", {
        x: cardX + 2.95,
        y: topY + 0.66,
        w: 1.0,
        h: 0.56,
        rectRadius: 0.08,
        fill: { color: C.ACCENT },
      }, {
        fontSize: 11,
        fontFace: FONT_B,
        color: C.WHITE,
        bold: true,
        align: "center",
        valign: "middle",
      });

      const bottomY = topY + miniH + gap;
      addCard(slide, cardX, bottomY, cardW, miniH, { strip: C.SECONDARY, fill: C.WHITE });
      slide.addText("Subtract: 5/6 - 1/3", {
        x: cardX + 0.18,
        y: bottomY + 0.08,
        w: cardW - 0.36,
        h: 0.24,
        fontSize: 13,
        fontFace: FONT_H,
        color: C.SECONDARY,
        bold: true,
        margin: 0,
      });
      slide.addText("Rename 1/3 as 2/6 first", {
        x: cardX + 0.18,
        y: bottomY + 0.34,
        w: cardW - 0.36,
        h: 0.2,
        fontSize: 10.5,
        fontFace: FONT_B,
        color: C.MUTED,
        italic: true,
        margin: 0,
      });
      addFractionStrip(slide, cardX + 0.22, bottomY + 0.58, 2.55, 0.16, 5, 6, {
        label: "5/6",
        fillColor: C.SECONDARY,
      });
      addFractionStrip(slide, cardX + 0.22, bottomY + 0.84, 2.55, 0.16, 2, 6, {
        label: "2/6",
        fillColor: C.ALERT,
      });
      addFractionStrip(slide, cardX + 0.22, bottomY + 1.1, 2.55, 0.16, 3, 6, {
        label: "3/6",
        fillColor: C.SUCCESS,
        labelBold: true,
      });
      addTextOnShape(slide, "5/6 - 2/6 = 3/6 = 1/2", {
        x: cardX + 2.95,
        y: bottomY + 0.66,
        w: 1.0,
        h: 0.56,
        rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, {
        fontSize: 11,
        fontFace: FONT_B,
        color: C.WHITE,
        bold: true,
        align: "center",
        valign: "middle",
      });
    }
  );

  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Quick Check: Addition",
      "Show Me Boards",
      "Rename first.\n\n1/4 + 2/8 = ?\n\nShow the rename step before you add.",
      NOTES_CFU_ADD,
      FOOTER
    ),
    (slide) => {
      addTextOnShape(slide, "1/4 = 2/8, so 2/8 + 2/8 = 4/8 = 1/2", {
        x: 1.15,
        y: 4.3,
        w: 7.7,
        h: 0.55,
        rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, {
        fontSize: 17,
        fontFace: FONT_H,
        color: C.WHITE,
        bold: true,
      });
    }
  );

  withReveal(
    () => workedExSlide(
      pres,
      3,
      "We Do",
      "Board Build: 5/6 - 1/3",
      [
        "Build this with me on the board.",
        "First: rename 1/3 as 2/6.",
        "Next: cross out 2 sixths from 5 sixths.",
        "Then: read what is left.",
        "Say the sentence: same-size parts first.",
      ],
      NOTES_WEDO_Q,
      FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY, fill: C.WHITE });
        slide.addText("Build it together", {
          x: lg.rightX + 0.2,
          y: lg.panelTopPadded + 0.1,
          w: lg.rightW - 0.4,
          h: 0.25,
          fontSize: 14,
          fontFace: FONT_H,
          color: C.SECONDARY,
          bold: true,
          margin: 0,
        });
        slide.addText("5/6 - 1/3", {
          x: lg.rightX + 0.2,
          y: lg.panelTopPadded + 0.36,
          w: lg.rightW - 0.4,
          h: 0.28,
          fontSize: 18,
          fontFace: FONT_H,
          color: C.CHARCOAL,
          bold: true,
          margin: 0,
          align: "center",
        });

        addFractionStrip(slide, lg.rightX + 0.22, lg.panelTopPadded + 0.78, 2.7, 0.22, 5, 6, {
          label: "5/6",
          fillColor: C.SECONDARY,
          labelBold: true,
        });
        addFractionStrip(slide, lg.rightX + 0.22, lg.panelTopPadded + 1.18, 2.7, 0.22, 2, 6, {
          label: "2/6",
          fillColor: C.ALERT,
          labelBold: true,
        });

        slide.addText("1/3 becomes 2/6", {
          x: lg.rightX + 3.08,
          y: lg.panelTopPadded + 1.14,
          w: 0.92,
          h: 0.28,
          fontSize: 10.5,
          fontFace: FONT_B,
          color: C.MUTED,
          italic: true,
          margin: 0,
        });

        addTextOnShape(slide, "What is left?", {
          x: lg.rightX + 0.32,
          y: lg.panelTopPadded + 2.0,
          w: lg.rightW - 0.64,
          h: 0.5,
          rectRadius: 0.08,
          fill: { color: C.BG_LIGHT },
          line: { color: C.SECONDARY, width: 1.2 },
        }, {
          fontSize: 15,
          fontFace: FONT_H,
          color: C.SECONDARY,
          bold: true,
        });
      }
    ),
    (slide) => {
      addCrossOutSegments(slide, 5.52, 2.16, 2.7, 0.22, 6, 3, 2, {
        labelW: 0.7,
        color: C.ALERT,
      });
      addTextOnShape(slide, "5/6 - 2/6 = 3/6 = 1/2", {
        x: 5.6,
        y: 3.6,
        w: 3.5,
        h: 0.52,
        rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, {
        fontSize: 16,
        fontFace: FONT_H,
        color: C.WHITE,
        bold: true,
      });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Hinge Check: Subtraction",
      "Show Me Boards",
      "Rename first.\n\n3/4 - 1/8 = ?\n\nShow the rename step before you subtract.",
      NOTES_CFU_SUB,
      FOOTER
    ),
    (slide) => {
      addTextOnShape(slide, "3/4 = 6/8, so 6/8 - 1/8 = 5/8", {
        x: 1.35,
        y: 4.3,
        w: 7.3,
        h: 0.55,
        rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, {
        fontSize: 17,
        fontFace: FONT_H,
        color: C.WHITE,
        bold: true,
      });
    }
  );

  workedExSlide(
    pres,
    4,
    "You Do",
    "You Do: Mixed Practice",
    [
      "First: choose a common denominator.",
      "Next: rename the fractions.",
      "Then: add or subtract and simplify.",
      "Record each problem in your workbook.",
      "Show the rename step every time.",
      "If you get stuck, sketch a fraction strip.",
    ],
    NOTES_YOUDO,
    FOOTER,
    (slide, lg) => {
      const cardW = 1.95;
      const cardH = 1.0;
      const leftX = lg.rightX;
      const rightX = lg.rightX + 2.15;
      const topY = lg.panelTopPadded;
      const bottomY = topY + 1.2;

      addPracticeCard(slide, leftX, topY, cardW, cardH, "Problem A", "1/2 + 1/3", C.PRIMARY);
      addPracticeCard(slide, rightX, topY, cardW, cardH, "Problem B", "7/8 - 1/4", C.SECONDARY);
      addPracticeCard(slide, leftX, bottomY, cardW, cardH, "Problem C", "3/5 + 1/10", C.ACCENT);
      addPracticeCard(slide, rightX, bottomY, cardW, cardH, "Problem D", "5/6 - 1/2", C.ALERT);

      addTextOnShape(slide, "Challenge: write one fraction problem that equals 1 whole", {
        x: lg.rightX,
        y: bottomY + 1.18,
        w: lg.rightW,
        h: 0.45,
        rectRadius: 0.08,
        fill: { color: C.BG_LIGHT },
        line: { color: C.ALERT, width: 1.1 },
      }, {
        fontSize: 12.5,
        fontFace: FONT_B,
        color: C.ALERT,
        bold: true,
      });
    }
  );

  exitTicketSlide(
    pres,
    [
      "2/3 + 1/6 = ?",
      "5/8 - 1/4 = ?",
      "Which denominator helps with 1/3 and 1/9?",
      "Why does the denominator stay the same once the parts are the same size?",
    ],
    NOTES_EXIT,
    FOOTER
  );

  closingSlide(
    pres,
    "Which step helps you most when fractions look different: rename, combine or simplify? Tell your partner why.",
    [
      "Rename first when the denominators do not match",
      "Add or subtract the numerators once the parts match",
      "Simplify if you can",
    ],
    NOTES_CLOSING
  );

  contentSlide(
    pres,
    "Resources",
    C.SECONDARY,
    "Teacher Resources",
    [
      "No printed student resources required.",
      "Materials: mini-whiteboards, markers, fraction strips or teacher-drawn bars, workbooks.",
      "Students complete the independent practice and exit ticket in their workbook.",
      "No websites or external links required for this lesson.",
    ],
    NOTES_RESOURCES,
    FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.SECONDARY, fill: C.WHITE });
      slide.addText("Have ready", {
        x: lg.rightX + 0.2,
        y: lg.panelTopPadded + 0.12,
        w: lg.rightW - 0.4,
        h: 0.26,
        fontSize: 16,
        fontFace: FONT_H,
        color: C.SECONDARY,
        bold: true,
        margin: 0,
      });

      const tileY = lg.panelTopPadded + 0.55;
      const tileW = 1.15;
      ["Whiteboards", "Markers", "Fraction strips"].forEach((label, index) => {
        const tx = lg.rightX + 0.22 + index * 1.28;
        slide.addShape("roundRect", {
          x: tx,
          y: tileY,
          w: tileW,
          h: 1.0,
          rectRadius: 0.08,
          fill: { color: index === 0 ? C.PRIMARY : (index === 1 ? C.ACCENT : C.SUCCESS) },
        });
        slide.addText(label, {
          x: tx + 0.08,
          y: tileY + 0.24,
          w: tileW - 0.16,
          h: 0.52,
          fontSize: 12,
          fontFace: FONT_B,
          color: C.WHITE,
          bold: true,
          align: "center",
          valign: "middle",
          margin: 0,
        });
      });

      addTextOnShape(slide, "Board prep: draw one whole bar and one sixths strip before class", {
        x: lg.rightX + 0.2,
        y: tileY + 1.28,
        w: lg.rightW - 0.4,
        h: 0.52,
        rectRadius: 0.08,
        fill: { color: C.BG_LIGHT },
        line: { color: C.SECONDARY, width: 1.1 },
      }, {
        fontSize: 12.5,
        fontFace: FONT_B,
        color: C.SECONDARY,
        bold: true,
      });
    }
  );

  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Fractions_Adding_Subtracting.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
  console.log("Session " + SESSION + " build complete.");
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
