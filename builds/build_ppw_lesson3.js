"use strict";

// Part Part Whole Unit — Session 3: Tens Frame PPW
// Foundation Numeracy | Term 2 Week 4 | Variant 3
// DR: Subitising numbers 11-20
// Fluency: Ordering numbers to 20
// VC2MFN04 — partition and combine collections up to 10 using part-part-whole relationships and subitising

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "foundation", weekToVariant(4));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 3;
const UNIT_TITLE = "Part, Part, Whole";
const FOOTER = "Part, Part, Whole | Session 3 of 4 | Foundation Numeracy";
const OUT_DIR = "output/PPW_Session3_Tens_Frame_PPW";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Tens Frame Recording Sheet", "Build numbers on a tens frame with 2 colours to show parts and whole.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Sample tens frame builds for whole numbers 5-10.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Pre-filled tens frames for numbers 4-6.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Double tens frames for teen numbers up to 20.");

const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

// Draw a tens frame with two-colour counters: first `partA` cells red, next `partB` yellow, rest empty
function drawTenFrameTwoColour(slide, x, y, w, h, partA, partB, opts) {
  const o = opts || {};
  const cols = 5;
  const rows = 2;
  const cellW = w / cols;
  const cellH = h / rows;
  const colorA = o.colorA || "D64545"; // red
  const colorB = o.colorB || "F4C430"; // yellow
  const borderColor = o.borderColor || C.PRIMARY;

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const idx = r * cols + c;
      const cx = x + c * cellW;
      const cy = y + r * cellH;
      slide.addShape("rect", {
        x: cx, y: cy, w: cellW, h: cellH,
        fill: { color: C.WHITE },
        line: { color: borderColor, width: 1.5 },
      });
      let dotColor = null;
      if (idx < partA) dotColor = colorA;
      else if (idx < partA + partB) dotColor = colorB;
      if (dotColor) {
        const dotSize = Math.min(cellW, cellH) * 0.62;
        slide.addShape("roundRect", {
          x: cx + (cellW - dotSize) / 2,
          y: cy + (cellH - dotSize) / 2,
          w: dotSize, h: dotSize, rectRadius: dotSize / 2,
          fill: { color: dotColor },
          line: { color: C.CHARCOAL, width: 0.7 },
        });
      }
    }
  }
}

// Draw dots in a loose scattered pattern for subitising (no frame around)
function drawDotCluster(slide, x, y, w, h, count, opts) {
  const o = opts || {};
  const dotSize = o.dotSize || 0.28;
  const color = o.color || C.PRIMARY;
  // Predefined positions for 11-20 arrangements; keeps them grouped
  const layouts = {
    11: [[0,0],[1,0],[2,0],[3,0],[4,0],[0,1],[1,1],[2,1],[3,1],[4,1], [1.5,2]],
    12: [[0,0],[1,0],[2,0],[3,0],[4,0],[0,1],[1,1],[2,1],[3,1],[4,1], [1,2],[2.5,2]],
    13: [[0,0],[1,0],[2,0],[3,0],[4,0],[0,1],[1,1],[2,1],[3,1],[4,1], [0.5,2],[2,2],[3.5,2]],
    14: [[0,0],[1,0],[2,0],[3,0],[4,0],[0,1],[1,1],[2,1],[3,1],[4,1], [0,2],[1.5,2],[3,2],[4.5,2]],
    15: [[0,0],[1,0],[2,0],[3,0],[4,0],[0,1],[1,1],[2,1],[3,1],[4,1], [0,2],[1,2],[2,2],[3,2],[4,2]],
    16: [[0,0],[1,0],[2,0],[3,0],[4,0],[0,1],[1,1],[2,1],[3,1],[4,1], [0,2],[1,2],[2,2],[3,2],[4,2],[2,3]],
    17: [[0,0],[1,0],[2,0],[3,0],[4,0],[0,1],[1,1],[2,1],[3,1],[4,1], [0,2],[1,2],[2,2],[3,2],[4,2],[1.5,3],[3,3]],
    18: [[0,0],[1,0],[2,0],[3,0],[4,0],[0,1],[1,1],[2,1],[3,1],[4,1], [0,2],[1,2],[2,2],[3,2],[4,2],[0.5,3],[2,3],[3.5,3]],
    19: [[0,0],[1,0],[2,0],[3,0],[4,0],[0,1],[1,1],[2,1],[3,1],[4,1], [0,2],[1,2],[2,2],[3,2],[4,2],[0,3],[1.5,3],[3,3],[4.5,3]],
    20: [[0,0],[1,0],[2,0],[3,0],[4,0],[0,1],[1,1],[2,1],[3,1],[4,1], [0,2],[1,2],[2,2],[3,2],[4,2],[0,3],[1,3],[2,3],[3,3],[4,3]],
  };
  const pts = layouts[count] || layouts[15];
  const cellW = w / 5;
  const cellH = h / 4;
  pts.forEach((p) => {
    const cx = x + p[0] * cellW + (cellW - dotSize) / 2;
    const cy = y + p[1] * cellH + (cellH - dotSize) / 2;
    slide.addShape("roundRect", {
      x: cx, y: cy, w: dotSize, h: dotSize, rectRadius: dotSize / 2,
      fill: { color },
      line: { color: C.CHARCOAL, width: 0.5 },
    });
  });
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Good morning! So far we have used counters and unifix cubes
- Today we use a very special tool: the TENS FRAME
- A tens frame helps us see numbers and their parts

DO:
- Have tens frames and two-colour counters ready on each table
- Display title slide as students settle

TEACHER NOTES:
Session 3 of 4. This session introduces the tens frame as a PPW tool. The tens frame gives a fixed visual structure: 10 cells in 2 rows of 5. This makes the whole visible and the parts easy to count.

WATCH FOR:
- Students who have used tens frames before - they can lead their table
- Students who look unsure - model filling it cell by cell

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Quick eyes! Look at the dots. Do not count one by one.
- How many dots do you see?
- Turn and tell your partner the number

DO:
- Flash each dot picture for 3 seconds
- Partners whisper the number
- Cold call 2 students per picture

TEACHER NOTES:
Daily Review on subitising 11-20. Subitising is quickly seeing a quantity without counting. Showing the dots briefly forces the eye to see groups rather than count individually.

WATCH FOR:
- Students who call the number instantly - subitising is secure
- Students who count in air - that is okay, counting is a fallback

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Let's check. First picture - how many? [13]
- Second - [17]
- Third - [20]
- Tick or fix your answer.

DO:
- Reveal each answer
- Celebrate students who "just knew" the number
- Remind: subitising gets easier with practice

TEACHER NOTES:
Tick-and-fix. Note which students can subitise reliably in the teen range - they will find today's tens frame work easy. Others will rely more on the visual structure.

WATCH FOR:
- Students who got all 3 - strong subitising
- Students who missed by 1 or 2 - they are close; more practice needed

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Number fluency time
- Look at the numbers on the slide. Some are in the right order. Some are not.
- Read them with me, then say which ones are out of order

DO:
- Read the number sequence aloud with the class
- Students point to or call out the out-of-order numbers
- Correct together as a class

TEACHER NOTES:
Ordering numbers to 20 is the fluency focus. Keeping the number line in view builds familiarity with the sequence that anchors counting and PPW.

WATCH FOR:
- Students who spot out-of-order numbers quickly - sequence is secure
- Students who read them all in order without noticing - prompt: "Wait - what comes after 11?"

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Our learning intention is the same as yesterday
- We are learning to partition numbers to 10 into two parts using materials
- Today the material is the TENS FRAME with counters

DO:
- Choral read LI and SC
- Hold up a tens frame so students can see the tool

TEACHER NOTES:
Same LI and SC as Sessions 1 and 2. The repetition helps Foundation students know where they are going. By the end of this session, most students should be able to place counters on a tens frame to show two parts of a number to 10.

WATCH FOR:
- Students confident in the language - they are ready for today's tool
- Students still new to the language - the tens frame will anchor it visually

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. I have a tens frame and coloured counters.
- I am going to show the number 8.
- I place 5 red counters in the top row - count them with me: 1, 2, 3, 4, 5
- Now I add 3 yellow counters in the bottom row: 1, 2, 3
- The whole is 8. One part is 5 red. The other part is 3 yellow.

DO:
- Use a real tens frame and counters under the document camera
- Place counters one at a time, saying each number
- Point to the slide visual as you say the parts

TEACHER NOTES:
The tens frame gives a fixed "out of 10" visual. When students see the filled cells, they can also see the empty cells - which prepares them to see 8 + 2 = 10. Keep that in mind for Session 4.

MISCONCEPTIONS:
- Misconception: Students think the whole is 10 because the tens frame has 10 cells
  Why: The frame itself looks like "10 something" so they conflate the frame with the quantity
  Impact: They say "10" when the counters only show 8
  Quick correction: "The frame is our house. But we only count the counters we put in. 5 red and 3 yellow. That is 8 counters, even though the house has room for 10."

WATCH FOR:
- Students who count the counters only - they separate frame from quantity
- Students who count the empty cells too - reteach with physical demo

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Let's show 8 a different way
- Same tens frame. Same 8 counters.
- This time: 2 red and 6 yellow.
- Count red: 1, 2. Count yellow: 1, 2, 3, 4, 5, 6.
- The whole is still 8. Now the parts are 2 and 6.

DO:
- Place counters: 2 red, then 6 yellow
- Point to both tens frames on the slide side by side
- Repeat: "Same whole, different parts"

TEACHER NOTES:
Same whole, different parts - reinforced with a fixed frame so students see that the TOTAL counters stays the same even when one part grows and the other shrinks.

WATCH FOR:
- Students who say "8 = 2 + 6" - applying the pattern
- Students pointing at both frames showing the tens - they are comparing

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Your eyes on the tens frame
- How many YELLOW counters do you see?
- Show me on your fingers
- Then say: The whole is ___. The parts are ___ and ___.

DO:
- Display tens frame: 4 red + 3 yellow (whole of 7)
- Scan fingers for the yellow count
- Cold call 2 students for the full sentence

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "How many yellow? Show me on your fingers."
- Scan for: 3 fingers
PROCEED: If 80%+ show 3, move to We Do
PIVOT: Most likely misconception - students count all dots including red. Reteach: "Point at only yellow dots. Count only those. How many yellow?"

TEACHER NOTES:
This CFU has students isolate one colour group on the tens frame. If most of them struggle, you need more I Do before We Do.

WATCH FOR:
- Students showing 3 quickly - they can isolate colours
- Students showing 7 - they are counting all dots; reteach one part at a time

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me
- The number card says 9
- Show 9 on your tens frame using 2 colours
- Find ONE way. Then find a DIFFERENT way.
- Show me when you are done.

DO:
- Display number card: 9
- Each partner pair has a tens frame and 10 counters (or use whiteboard)
- Allow 2 minutes
- Cold call 3 pairs to share their ways

TEACHER NOTES:
We Do takes the I Do routine into partner practice. Having students find two ways stretches the "many ways" idea.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Session 3 Enabling Scaffold has tens frames pre-filled for 4, 5, and 6. Students circle one red group and one yellow group.
- Extra Notes: Lower numbers, one tens frame per page.
EXTENDING PROMPT:
- Task: Use the Session 3 Extension - double tens frames for teen numbers (11-20). Show one way to make each teen number.
- Extra Notes: Extends the concept to teen numbers using two tens frames.

WATCH FOR:
- Students placing counters fluently - tens frame routine is secure
- Students who place all red then all yellow - that is correct; just check the count

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's see different ways to make 9 on a tens frame
- 9 = 5 and 4    or    9 = 6 and 3    or    9 = 8 and 1...
- All of these show 9 counters, just in different colours.
- Notice: the frame has one empty cell each time. Why? Because 9 is 1 less than 10.

DO:
- Reveal the answers
- Celebrate the variety of responses
- Point to the empty cell - this will be useful in Session 4

TEACHER NOTES:
The empty cell on a tens frame for whole of 9 is a teaching moment. It shows 9 is "almost 10" or "10 minus 1". Leave this observation for students to notice; do not hammer it.

WATCH FOR:
- Students who notice the empty cell - they are ready for bonds to 10
- Students who just record their own way - that is fine for today

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. Look at the tens frame on the slide.
- Whole is 7. Parts are 4 and 3. Is this right?
- Thumbs up or thumbs down.

DO:
- Display tens frame: 4 red + 3 yellow
- Scan thumbs
- Follow-up: "How do you know?" Cold call 2 students

CFU CHECKPOINT:
Technique: Thumbs Up/Down
Script:
- Say: "Tens frame shows 4 red and 3 yellow. Whole is 7. Right or wrong?"
- Scan for: thumbs up
PROCEED: If 80%+ get it, move to You Do
PIVOT: Most likely misconception - students count empty cells. Reteach: "Only count dots. Red dots: 1, 2, 3, 4. Yellow dots: 1, 2, 3. Total dots: 7. So yes, 7 is right."

TEACHER NOTES:
This hinge check is positive. For Foundation, one positive example is often enough - the negative example can be reserved for when students are very confident.

WATCH FOR:
- Thumbs up with confidence - students can match tens frame to whole-and-parts
- Hesitation - more tens frame practice needed in You Do

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Your turn on your own or with a partner
- First: Take a number card. Read the number.
- Next: Place counters on your tens frame using 2 colours.
- Then: Draw it on your sheet. Write the parts.
- You have 10 minutes.

DO:
- Hand out the Session 3 Tens Frame Recording Sheet
- Each student has a tens frame + 10 two-colour counters + number cards 1-10
- Circulate - start with strugglers
- Enabler students get Session 3 Enabling Scaffold
- Extender students get Session 3 Extension (double tens frames for teens)

TEACHER NOTES:
Students who can flexibly show a number on a tens frame in different ways are ready for Session 4. Students who can show one way reliably have met the core target.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Session 3 Enabling Scaffold - tens frames for 4, 5, 6. Counters are pre-drawn. Students circle the red group and yellow group.
- Extra Notes: One tens frame per page, larger circles.
EXTENDING PROMPT:
- Task: Session 3 Extension - use double tens frames to show teen numbers (11, 13, 15). One way for each.
- Extra Notes: Self-contained with blank double tens frames.

WATCH FOR:
- Students filling tens frames cell-by-cell in order - routine is secure
- Students placing counters randomly - redirect to left-to-right, top-to-bottom filling

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket. Work on your own.
- Look at the tens frame. Answer each question.
- You have 3 minutes.

DO:
- Display exit ticket slide
- Students respond on whiteboards
- Collect or scan to sort

TEACHER NOTES:
Q1 checks SC1 (identify parts from a tens frame). Q2 checks SC2 (show parts differently). Q3 checks SC3 (explain with the language).

WATCH FOR:
- Students completing all three accurately - ready for Session 4
- Students who only read the tens frame (not make their own) - they need more Session 4 practice

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Let's review our learning
- Thumbs up, sideways, or down for each success criterion
- Turn and tell: What is special about a tens frame?

DO:
- Show SC on screen
- Thumbs check each
- 30 seconds Turn and Talk
- Cold call 2 students

TEACHER NOTES:
Closing prompts students to articulate the tool's benefit. "It has 10 cells" or "it shows the parts clearly" are great student responses.

WATCH FOR:
- Students who say "it helps me see parts" - concept is clear
- Students who say "it is blue and white" - focus on the tool, not the concept; redirect

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Today's resources
- The Tens Frame Recording Sheet is the main one
- Enabler and extension are on the differentiation shelf

DO:
- Point out each resource

TEACHER NOTES:
Materials needed: tens frames (1 per student), 10 two-colour counters per student, number cards 1-10, whiteboards.

WATCH FOR:
- N/A - teacher reference only

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Session 3: Tens Frame Part Part Whole",
    "Foundation Numeracy | Session 3 of 4 | Term 2 Week 4", NOTES_TITLE);

  // Slide 2-3: Daily Review — Subitising 11-20 (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Quick Eyes! How Many Dots?", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      // Left: instruction
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      s.addText([
        { text: "Look quickly.", options: { fontSize: 18, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Do NOT count one by one.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "See the groups.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Whisper the number.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right: three dot clusters
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      const clusters = [
        { label: "A.", count: 13 },
        { label: "B.", count: 17 },
        { label: "C.", count: 20 },
      ];
      clusters.forEach((cl, i) => {
        const rowY = CONTENT_TOP + 0.15 + i * 1.10;
        s.addText(cl.label, {
          x: 5.3, y: rowY + 0.1, w: 0.35, h: 0.5,
          fontSize: 20, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
          align: "center", valign: "middle",
        });
        drawDotCluster(s, 5.75, rowY, 3.5, 1.05, cl.count, { dotSize: 0.18, color: C.PRIMARY });
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "A) 13     B) 17     C) 20", {
        x: 0.8, y: 4.55, w: 8.4, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Ordering numbers to 20
  const sFluency = pres.addSlide();
  addTopBar(sFluency, STAGE_COLORS["1"]);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Order the Numbers", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

  addCard(sFluency, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  sFluency.addText([
    { text: "Read them with me", options: { fontSize: 18, bold: true, color: C.PRIMARY, breakLine: true } },
    { text: "", options: { fontSize: 8, breakLine: true } },
    { text: "Point to: 11, 12, 13...", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
    { text: "What comes after 15?", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
    { text: "What comes before 20?", options: { bullet: true, fontSize: 15, color: C.CHARCOAL } },
  ], {
    x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  // Right: number line 1-20 in two rows
  addCard(sFluency, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
  const flGridStartX = 5.35;
  const flGridStartY = CONTENT_TOP + 0.25;
  const flCellW = 0.40;
  const flCellH = 0.55;
  for (let i = 0; i < 20; i++) {
    const row = Math.floor(i / 10);
    const col = i % 10;
    const num = i + 1;
    const isTeen = num >= 11;
    sFluency.addShape("roundRect", {
      x: flGridStartX + col * flCellW,
      y: flGridStartY + row * (flCellH + 0.10),
      w: flCellW - 0.04,
      h: flCellH,
      rectRadius: 0.06,
      fill: { color: isTeen ? C.ACCENT : C.PRIMARY },
      line: { color: isTeen ? C.ACCENT : C.PRIMARY, width: 1 },
    });
    sFluency.addText(String(num), {
      x: flGridStartX + col * flCellW,
      y: flGridStartY + row * (flCellH + 0.10),
      w: flCellW - 0.04,
      h: flCellH,
      fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
  }
  sFluency.addText("Teen numbers!", {
    x: flGridStartX, y: flGridStartY + 2 * (flCellH + 0.10) + 0.05, w: 4, h: 0.35,
    fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, italic: true, margin: 0,
    align: "center",
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    ["We are learning to partition numbers to 10 into two parts using materials"],
    [
      "I can show a number as two parts",
      "I can find different ways to make the same number",
      "I can use pictures, objects, or numbers to explain my thinking",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — Show 8 on a tens frame
  workedExSlide(pres, 2, "I Do", "Show 8 on a Tens Frame",
    [
      "I need 8 counters",
      "",
      "5 red in the top row",
      "3 yellow in the bottom row",
      "",
      "Count red: 5",
      "Count yellow: 3",
      "",
      "Whole = 8",
      "Parts = 5 and 3",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.PRIMARY });

      slide.addText("Tens Frame", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.2, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      // Tens frame: 5 red + 3 yellow
      const tfW = 3.2;
      const tfH = 1.28;
      const tfX = lg.rightX + (lg.rightW - tfW) / 2;
      const tfY = lg.panelTopPadded + 0.5;
      drawTenFrameTwoColour(slide, tfX, tfY, tfW, tfH, 5, 3);

      slide.addText("8 = 5 + 3", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 2.0, w: lg.rightW - 0.2, h: 0.4,
        fontSize: 22, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0, align: "center",
      });
      slide.addText("2 cells empty", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 2.5, w: lg.rightW - 0.2, h: 0.3,
        fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0, align: "center",
      });
    }
  );

  // Slide 7: I Do #2 — 8 a different way
  workedExSlide(pres, 2, "I Do", "Same 8, Different Parts",
    [
      "Same 8 counters",
      "",
      "Way 1: 5 red and 3 yellow",
      "",
      "Way 2: 2 red and 6 yellow",
      "",
      "Both show 8",
      "",
      "Same whole!",
      "Different parts!",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.5, { strip: C.SECONDARY });

      // Way 1: 5 red + 3 yellow
      slide.addText("Way 1:  8 = 5 + 3", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.2, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      drawTenFrameTwoColour(slide, lg.rightX + 0.35, lg.panelTopPadded + 0.45, 3.5, 1.1, 5, 3);

      // Way 2: 2 red + 6 yellow
      slide.addText("Way 2:  8 = 2 + 6", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 1.75, w: lg.rightW - 0.2, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0, align: "center",
      });
      drawTenFrameTwoColour(slide, lg.rightX + 0.35, lg.panelTopPadded + 2.10, 3.5, 1.1, 2, 6);
    }
  );

  // Slide 8-9: CFU — tens frame 4 red + 3 yellow (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["2"]);
      addStageBadge(s, 2, "CFU");
      addTitle(s, "Count the Yellow Dots", { y: 0.65, fontSize: 24, color: STAGE_COLORS["2"] });

      // Left: prompt
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["2"] });
      s.addText([
        { text: "Show me on your fingers:", options: { fontSize: 18, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "How many YELLOW?", options: { fontSize: 22, bold: true, color: C.ALERT, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Then say:", options: { fontSize: 14, color: C.MUTED, breakLine: true } },
        { text: "The whole is ___.", options: { fontSize: 14, italic: true, color: C.CHARCOAL, breakLine: true } },
        { text: "The parts are ___ and ___.", options: { fontSize: 14, italic: true, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.2, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right: tens frame 4 red + 3 yellow
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      s.addText("The Tens Frame", {
        x: 5.35, y: CONTENT_TOP + 0.1, w: 4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0, align: "center",
      });
      drawTenFrameTwoColour(s, 5.5, CONTENT_TOP + 0.55, 3.7, 1.5, 4, 3);

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "3 yellow!   Whole is 7.  Parts are 4 and 3.", {
        x: 1.0, y: 4.1, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10-11: We Do — number card 9 (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Make 9 on a Tens Frame", { y: 0.65, fontSize: 24, color: STAGE_COLORS["3"] });

      // Left: prompt
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });
      s.addText([
        { text: "With your partner:", options: { fontSize: 18, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Fill your tens frame", options: { bullet: true, fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "Use 2 colours", options: { bullet: true, fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "Find 2 different ways", options: { bullet: true, fontSize: 16, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "9 = ___ and ___", options: { fontSize: 18, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "9 = ___ and ___", options: { fontSize: 18, bold: true, color: C.PRIMARY } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right: big number card 9
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      s.addText("Number Card:", {
        x: 5.35, y: CONTENT_TOP + 0.1, w: 4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0, align: "center",
      });
      s.addShape("roundRect", {
        x: 6.3, y: CONTENT_TOP + 0.5, w: 2.1, h: 2.0, rectRadius: 0.15,
        fill: { color: C.WHITE },
        line: { color: C.ACCENT, width: 4 },
      });
      s.addText("9", {
        x: 6.3, y: CONTENT_TOP + 0.5, w: 2.1, h: 2.0,
        fontSize: 130, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText("Fill your tens frame with 9", {
        x: 5.35, y: CONTENT_TOP + 2.7, w: 4, h: 0.35,
        fontSize: 15, fontFace: FONT_B, color: C.CHARCOAL, italic: true, margin: 0, align: "center",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      runSlideDiagnostics(s, pres, { respectSafeBottom: false });
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "9 = 5+4, 6+3, 8+1, 4+5, 7+2, 9+0 (or any pair that makes 9)", {
        x: 0.4, y: 4.55, w: 9.2, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 12-13: CFU Hinge — tens frame check (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["2"]);
      addStageBadge(s, 2, "CFU");
      addTitle(s, "Is This Right?", { y: 0.65, fontSize: 24, color: STAGE_COLORS["2"] });

      // Left: prompt
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["2"] });
      s.addText([
        { text: "Whole = 7", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "Parts = 4 and 3", options: { fontSize: 22, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Is this right?", options: { fontSize: 18, bold: true, color: C.ALERT, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "Thumbs UP = yes", options: { fontSize: 15, color: C.SUCCESS, breakLine: true } },
        { text: "Thumbs DOWN = no", options: { fontSize: 15, color: C.ALERT } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.2, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Right: tens frame showing 4 red + 3 yellow
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      s.addText("Check the tens frame:", {
        x: 5.35, y: CONTENT_TOP + 0.15, w: 4, h: 0.3,
        fontSize: 15, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0, align: "center",
      });
      drawTenFrameTwoColour(s, 5.5, CONTENT_TOP + 0.65, 3.7, 1.5, 4, 3);

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU2);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Thumbs UP!  4 + 3 = 7", {
        x: 1.0, y: 4.1, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do
  workedExSlide(pres, 4, "You Do", "Your Turn: Tens Frame Cards",
    [
      "First: Pick a number card (1-10)",
      "",
      "Next: Fill your tens frame with counters in 2 colours",
      "",
      "Then: Draw it on your sheet. Write the parts.",
      "",
      "Try to do 3 different cards.",
      "",
      "You have 10 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.0, { strip: C.ALERT });
      slide.addText("You need:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.35,
        fontSize: 16, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "a tens frame", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "10 two-colour counters", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "number cards 1-10", options: { bullet: true, fontSize: 15, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Tip card
      addCard(slide, lg.rightX, lg.panelTopPadded + 2.2, lg.rightW, 1.1, { strip: C.PRIMARY });
      slide.addText("Fill left to right, top to bottom", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.35, w: lg.rightW - 0.3, h: 0.7,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, italic: true, margin: 0,
        align: "center", valign: "middle",
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "Tens frame with 6 red + 2 yellow. What is the whole? What are the parts?",
      "Draw a tens frame showing the number 7 in TWO different ways.",
      "Tell your teacher: why does a tens frame help us see parts?",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 16: Closing
  closingSlide(pres,
    "Turn and tell: What is one way to make 10 on a tens frame?",
    [
      "I can show a number as two parts",
      "I can find different ways to make the same number",
      "I can use pictures, objects, or numbers to explain my thinking",
    ],
    NOTES_CLOSING);

  // Slide 17: Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "PPW_Session3_Tens_Frame_PPW.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ─────────────────────────────────────────────────────────────────

  // Helper: draw a tens frame in PDF
  function drawPdfTensFrame(doc, x, y, cellSize, preFilled) {
    // preFilled: array of colours (null for empty) length 10
    for (let i = 0; i < 10; i++) {
      const row = Math.floor(i / 5);
      const col = i % 5;
      const cx = x + col * cellSize;
      const cy = y + row * cellSize;
      doc.rect(cx, cy, cellSize, cellSize).lineWidth(1.2).stroke("#1F3A63");
      if (preFilled && preFilled[i]) {
        const dotR = cellSize * 0.32;
        doc.circle(cx + cellSize / 2, cy + cellSize / 2, dotR)
          .lineWidth(0.8).fillAndStroke(preFilled[i], "#333333");
        doc.fillColor("#333333");
      }
    }
  }

  // Tens Frame Recording Sheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Fill the tens frames with 2 colours. Write the parts.",
      color: C.PRIMARY,
      lessonInfo: "Session 3 of 4 | Foundation Numeracy",
    });
    y = addTipBox(doc, "Colour some dots RED and some dots YELLOW. Write how many of each colour and the whole.", y, { color: C.ACCENT });

    // Draw 4 empty tens frames with spaces to write
    const drawBlock = (label, ty, cellSize) => {
      doc.fontSize(13).fillColor("#333333").text(label, 60, ty);
      drawPdfTensFrame(doc, 60, ty + 22, cellSize, null);
      doc.fontSize(12).fillColor("#333333")
        .text("Red: _____", 60 + 5 * cellSize + 20, ty + 30)
        .text("Yellow: _____", 60 + 5 * cellSize + 20, ty + 50)
        .text("Whole: _____", 60 + 5 * cellSize + 20, ty + 70);
      return ty + 22 + 2 * cellSize + 20;
    };

    y = addSectionHeading(doc, "Card 1:  I picked ______", y, { color: C.PRIMARY });
    y = drawBlock("Tens frame:", y, 40);

    y = addSectionHeading(doc, "Card 2:  I picked ______", y, { color: C.PRIMARY });
    y = drawBlock("Tens frame:", y, 40);

    y = addSectionHeading(doc, "Card 3:  I picked ______", y, { color: C.PRIMARY });
    y = drawBlock("Tens frame:", y, 40);

    addPdfFooter(doc, "Session 3 | Tens Frame Recording Sheet | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Sample tens frame splits for each whole",
      color: C.PRIMARY,
      lessonInfo: "Session 3 of 4 | Foundation Numeracy",
    });
    y = addBodyText(doc, "Any correct PPW pair is right. These are common examples students record.", y);

    const samples = [
      { whole: 5, ways: "5+0, 4+1, 3+2, 2+3, 1+4, 0+5" },
      { whole: 6, ways: "6+0, 5+1, 4+2, 3+3, 2+4, 1+5, 0+6" },
      { whole: 7, ways: "7+0, 6+1, 5+2, 4+3, 3+4, 2+5, 1+6, 0+7" },
      { whole: 8, ways: "8+0, 7+1, 6+2, 5+3, 4+4, 3+5, 2+6, 1+7, 0+8" },
      { whole: 9, ways: "9+0, 8+1, 7+2, 6+3, 5+4, 4+5, 3+6, 2+7, 1+8, 0+9" },
      { whole: 10, ways: "10+0, 9+1, 8+2, 7+3, 6+4, 5+5, 4+6, 3+7, 2+8, 1+9, 0+10" },
    ];
    samples.forEach((s) => {
      y = addSectionHeading(doc, "Whole = " + s.whole, y, { color: C.PRIMARY });
      y = addBodyText(doc, s.ways, y);
    });

    y = addTipBox(doc, "A whole of N has N + 1 different PPW pairs. Students who notice this pattern are ready for Session 4.", y, { color: C.ACCENT });

    addPdfFooter(doc, "Session 3 | Answer Key | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold — pre-filled tens frames for 4, 5, 6
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Count the RED dots. Count the YELLOW dots.",
      color: C.ACCENT,
      lessonInfo: "Session 3 of 4 | Foundation Numeracy",
    });
    y = addTipBox(doc, "The tens frames are filled for you. Count each colour. Write the parts and the whole.", y, { color: C.ACCENT });

    const buildFrame = (redCount, yellowCount) => {
      const arr = [];
      for (let i = 0; i < 10; i++) {
        if (i < redCount) arr.push("#D64545");
        else if (i < redCount + yellowCount) arr.push("#F4C430");
        else arr.push(null);
      }
      return arr;
    };

    // Whole of 4: 3 red + 1 yellow
    y = addSectionHeading(doc, "Whole = 4", y, { color: C.PRIMARY });
    drawPdfTensFrame(doc, 60, y, 48, buildFrame(3, 1));
    doc.fontSize(13).fillColor("#333333")
      .text("Red: _____", 60 + 5 * 48 + 20, y + 30)
      .text("Yellow: _____", 60 + 5 * 48 + 20, y + 52)
      .text("Whole: 4", 60 + 5 * 48 + 20, y + 74);
    y = y + 2 * 48 + 25;

    // Whole of 5: 2 red + 3 yellow
    y = addSectionHeading(doc, "Whole = 5", y, { color: C.PRIMARY });
    drawPdfTensFrame(doc, 60, y, 48, buildFrame(2, 3));
    doc.fontSize(13).fillColor("#333333")
      .text("Red: _____", 60 + 5 * 48 + 20, y + 30)
      .text("Yellow: _____", 60 + 5 * 48 + 20, y + 52)
      .text("Whole: 5", 60 + 5 * 48 + 20, y + 74);
    y = y + 2 * 48 + 25;

    // Whole of 6: 4 red + 2 yellow
    y = addSectionHeading(doc, "Whole = 6", y, { color: C.PRIMARY });
    drawPdfTensFrame(doc, 60, y, 48, buildFrame(4, 2));
    doc.fontSize(13).fillColor("#333333")
      .text("Red: _____", 60 + 5 * 48 + 20, y + 30)
      .text("Yellow: _____", 60 + 5 * 48 + 20, y + 52)
      .text("Whole: 6", 60 + 5 * 48 + 20, y + 74);
    y = y + 2 * 48 + 25;

    addPdfFooter(doc, "Session 3 | Enabling Scaffold | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Extension — double tens frames for teen numbers
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Show a teen number on two tens frames",
      color: C.PRIMARY,
      lessonInfo: "Session 3 of 4 | Foundation Numeracy",
    });
    y = addBodyText(doc, "Challenge: teen numbers need TWO tens frames. Fill the first frame with 10. Put the extras in the second frame.", y);

    const drawDoubleFrame = (ty, cellSize) => {
      drawPdfTensFrame(doc, 60, ty, cellSize, null);
      drawPdfTensFrame(doc, 60 + 5 * cellSize + 30, ty, cellSize, null);
    };

    y = addSectionHeading(doc, "Show 11", y, { color: C.PRIMARY });
    drawDoubleFrame(y, 32);
    doc.fontSize(12).fillColor("#333333")
      .text("Tens frame 1: _____ dots", 60, y + 2 * 32 + 10)
      .text("Tens frame 2: _____ dots", 60, y + 2 * 32 + 28)
      .text("Total: 11", 60, y + 2 * 32 + 46);
    y = y + 2 * 32 + 65;

    y = addSectionHeading(doc, "Show 13", y, { color: C.PRIMARY });
    drawDoubleFrame(y, 32);
    doc.fontSize(12).fillColor("#333333")
      .text("Tens frame 1: _____ dots", 60, y + 2 * 32 + 10)
      .text("Tens frame 2: _____ dots", 60, y + 2 * 32 + 28)
      .text("Total: 13", 60, y + 2 * 32 + 46);
    y = y + 2 * 32 + 65;

    y = addSectionHeading(doc, "Show 15", y, { color: C.PRIMARY });
    drawDoubleFrame(y, 32);
    doc.fontSize(12).fillColor("#333333")
      .text("Tens frame 1: _____ dots", 60, y + 2 * 32 + 10)
      .text("Tens frame 2: _____ dots", 60, y + 2 * 32 + 28)
      .text("Total: 15", 60, y + 2 * 32 + 46);
    y = y + 2 * 32 + 65;

    y = addTipBox(doc, "Notice: the first tens frame is always full (10). The second frame holds the extras. That is how teen numbers work.", y, { color: C.ACCENT });

    addPdfFooter(doc, "Session 3 | Extension | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 3 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
