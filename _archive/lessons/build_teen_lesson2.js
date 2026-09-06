"use strict";

// Teen Numbers Unit — Session 2: Teen Numbers 16-20
// Term 2 Week 1, Foundation Numeracy, Variant 0
// DR: Digit Formations — I can write the teen numbers
// Fluency: Backwards counting from 10
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
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 2;
const UNIT_TITLE = "Teen Numbers";
const FOOTER = "Teen Numbers | Session 2 of 5 | Foundation Numeracy";
const OUT_DIR = "output/Teen_Session2_Numbers_16_to_20";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Ten frame practice for teen numbers 16-20.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with completed ten frames.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Pre-drawn ten frames for 16-18 with visual guides.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Self-contained investigation: all teen numbers 11-20 review and beyond.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Ten frame helpers ──────────────────────────────────────────────────────

function drawTenFrame(slide, x, y, w, h, filled, opts) {
  const o = opts || {};
  const cols = 5;
  const rows = 2;
  const cellW = w / cols;
  const cellH = h / rows;
  const borderColor = o.borderColor || C.PRIMARY;

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const idx = r * cols + c;
      const cx = x + c * cellW;
      const cy = y + r * cellH;
      slide.addShape("rect", {
        x: cx, y: cy, w: cellW, h: cellH,
        fill: { color: idx < filled ? (o.fillColor || C.PRIMARY) : (o.emptyColor || C.WHITE) },
        line: { color: borderColor, width: 1.5 },
      });
      if (idx < filled) {
        const dotSize = Math.min(cellW, cellH) * 0.55;
        slide.addShape("roundRect", {
          x: cx + (cellW - dotSize) / 2,
          y: cy + (cellH - dotSize) / 2,
          w: dotSize, h: dotSize, rectRadius: dotSize / 2,
          fill: { color: o.dotColor || C.ACCENT },
        });
      }
    }
  }
}

function drawTenFrameWithExtras(slide, x, y, w, h, total, opts) {
  const o = opts || {};
  drawTenFrame(slide, x, y, w, h, Math.min(total, 10), o);
  const extras = total - 10;
  if (extras > 0) {
    const cellW = w / 5;
    const cellH = h / 2;
    const dotSize = Math.min(cellW, cellH) * 0.55;
    const extraY = y + h + 0.08;
    for (let i = 0; i < extras; i++) {
      const cx = x + i * cellW;
      slide.addShape("rect", {
        x: cx, y: extraY, w: cellW, h: cellH,
        fill: { color: C.WHITE },
        line: { color: C.SECONDARY, width: 1.5 },
      });
      slide.addShape("roundRect", {
        x: cx + (cellW - dotSize) / 2,
        y: extraY + (cellH - dotSize) / 2,
        w: dotSize, h: dotSize, rectRadius: dotSize / 2,
        fill: { color: o.extraDotColor || C.ALERT },
      });
    }
  }
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Yesterday we met the teen numbers 11 to 15
- Today we are going to learn the rest -- 16, 17, 18, 19, and 20
- By the end of today, you will know ALL the teen numbers

DO:
- Display title slide as students settle
- Have counters and ten frames ready

TEACHER NOTES:
Session 2 of 5. Extends the teen number concept to 16-20. Students who grasped the pattern in Session 1 should transfer it easily. The key challenge is 20, where the extras row is completely full.

WATCH FOR:
- Students who remember "ten and some more" from yesterday -- prior learning is intact

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Let's practise writing our teen numbers
- I will show you some numbers. Write each one on your whiteboard
- Make your numbers nice and clear

DO:
- Display the numbers to write: 11, 14, 13, 15, 12
- Allow 10 seconds per number
- Students write on whiteboards and hold up

TEACHER NOTES:
Daily Review practises digit formation for teen numbers from Session 1. Writing reinforces recognition. After the break, fine motor skills may need warming up.

WATCH FOR:
- Students who reverse digits (writing 41 instead of 14) -- this is common at Foundation
- Students who form digits correctly and quickly -- fine motor and number recognition are strong

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your numbers. Do they match?
- Fix any that are different
- Remember: the 1 comes first in all teen numbers (11, 12, 13, 14, 15)

DO:
- Click to reveal correct formations
- Students compare and fix
- Point out the 1 at the start of every teen number

TEACHER NOTES:
Tick-and-fix for digit formation. The "1 comes first" observation reinforces the tens-and-ones structure from a writing perspective.

WATCH FOR:
- Students who reversed digits -- address privately, not publicly
- Students who self-corrected -- metacognition is developing

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time -- today we are counting BACKWARDS
- Start at 10 and count back to 1 with me
- 10, 9, 8, 7, 6, 5, 4, 3, 2, 1... blast off!
- Now let's try it faster. Ready? Go!

DO:
- Lead backward count 10 to 1 (whole class, 3 rounds)
- Round 1: normal pace
- Round 2: faster
- Round 3: whisper count
- Use hand gestures (lowering hand as numbers decrease)

TEACHER NOTES:
Backward counting builds number sequence flexibility. Students who can only count forward have a rigid number sense. Backward counting from 10 is achievable for most Foundation students and prepares for subtraction concepts later.

WATCH FOR:
- Students who hesitate at 7 or 6 -- the middle of the sequence is often less automatic
- Students who count back fluently -- backward counting is automated

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me
- Read from slide: "We are learning to name and recognise the numbers 16 to 20"
- Now our success criteria. Let's read them together.
- Ask: What did we learn about teen numbers yesterday? [They have a ten and some more]

DO:
- Choral read LI and SC
- Quick recall of yesterday's key idea
- Accept and affirm student responses

TEACHER NOTES:
Connecting to Session 1 activates prior knowledge. SC1 is the floor (saying 16-20). SC2 is the core target (representing with ten frames). SC3 extends to explaining the structure.

WATCH FOR:
- Students who remember "ten and some more" -- Session 1 learning has stuck
- Students who seem unsure -- they may need a quick recap before I Do

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Yesterday we learned 11 to 15. Today we keep going
- Watch me. 16 = ten and 6 more. Six-teen. Hear "six" in "sixteen"?
- 17 = ten and 7 more. Seven-teen. Hear "seven" in "seventeen"?
- 18 = ten and 8 more. Eight-teen. Hear "eight" in "eighteen"?
- The pattern is the same. Full ten frame plus extras.

DO:
- Build each number on the ten frame visual
- Point to the extras for each
- Have students repeat each number name
- Emphasise the consistent pattern

TEACHER NOTES:
This I Do follows the same structure as Session 1 to build consistency. The naming pattern is very transparent for 16-18 (six-teen, seven-teen, eight-teen), which helps students see the structure in the name itself.

WATCH FOR:
- Students who predict the next number before you build it -- pattern recognition is strong
- Students who can hear the ones digit in the name -- phonological-numerical connection is forming

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now the tricky ones. 19 = ten and 9 more. Nine-teen.
- And 20. Twenty is special. Watch.
- 20 = ten and 10 more. That means BOTH rows of the ten frame are full!
- 20 is actually TWO tens. Two full ten frames.
- That is why we call it twenty -- two-ty means two tens.

DO:
- Build 19 on the ten frame (full ten + 9 extras)
- Build 20: show both the ten frame completely full AND two separate full ten frames
- Emphasise that 20 = 2 tens

TEACHER NOTES:
20 is the conceptual leap. It is the first time students see a number made of two complete tens with no extras. This is the bridge to understanding place value. Do not rush this -- let the visual speak.

MISCONCEPTIONS:
- Misconception: 20 follows the same "teen" pattern as 11-19
  Why: Students generalise the "ten and ones" pattern without realising 20 is qualitatively different -- it has no leftover ones
  Impact: If students do not see 20 as "two tens," they will struggle with the decade numbers (30, 40, 50) later
  Quick correction: "20 is special. There are no extras left over. Both rows are completely full. That is two whole tens."

WATCH FOR:
- Students who look surprised when both rows fill up -- they are noticing something new
- Students who say "two tens" or "both full" -- conceptual understanding is emerging

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. I am going to show you a ten frame
- Count the full ten and the extras
- Write the number on your whiteboard. Ready... show me.

DO:
- Display a ten frame showing 17 (full ten + 7 extras)
- Students write on whiteboards
- Scan for 17

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Look at this ten frame. Full ten plus extras. Write the number on your whiteboard. You have 5 seconds. Show me."
- Scan for: 17. Students who write 17 quickly can read the ten frame for numbers beyond 15.
PROCEED: If 80%+ show 17, move to the We Do.
PIVOT: Most likely misconception -- students write 7 (counting only the extras, forgetting the ten). Reteach: "Remember, the full ten frame means we start at 10. Then we count the extras: 1, 2, 3, 4, 5, 6, 7. So it is 10 and 7 more -- 17. Always start with 10."

TEACHER NOTES:
This CFU checks transfer from Session 1 to the new numbers 16-20. If students can read 17 from a ten frame, the concept has generalised.

WATCH FOR:
- Students who write 7 instead of 17 -- they are counting extras without the ten
- Readiness signal: students writing 17 without hesitation

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Partner time. I will show you ten frames for bigger teen numbers
- Work together. Count the ten, count the extras, write the number
- Hold up your whiteboards when ready

DO:
- Display ten frames for 16, 19, 18, 20 one at a time
- Partners discuss and write on whiteboards
- Allow 15-20 seconds per number

TEACHER NOTES:
We Do uses the new numbers 16-20. Including 20 checks whether students recognise the two-tens structure. Different content from Session 1 We Do ensures genuine practice, not recall.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 2 Enabling Scaffold with pre-shaded ten frames for 16, 17, 18. Students count extras and write the number with visual cues.
- Extra Notes: Distribute the Session 2 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: After identifying each number, use the Session 2 Extension to match ALL teen numbers 11-20 to their ten frame representations and explore what comes after 20.
- Extra Notes: Self-contained investigation with visual examples.

WATCH FOR:
- Students who correctly identify 20 as "two full tens" -- they understand the leap beyond teen
- Students who write 10 for the 20 frame -- they may be reading only the extras row

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check. First: full ten and 6 extras. [16] Sixteen!
- Second: full ten and 9 extras. [19] Nineteen! That is a lot of extras.
- Third: full ten and 8 extras. [18] Eighteen!
- Last one: BOTH rows full. No extras left. [20] Twenty! Two whole tens.

DO:
- Reveal answers one at a time
- Cold call for each
- For 20, emphasise "two tens, no extras"

TEACHER NOTES:
The reveal reinforces the naming pattern. Highlighting 20 as "two tens" is deliberate -- this is the most important conceptual moment in the session.

WATCH FOR:
- Students who self-correct -- learning is active
- Students who answered 20 correctly -- they have grasped the two-tens concept

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. I will say a number. You show me on your fingers how many EXTRAS it has after the ten.
- Ready? The number is 18. How many extras? Show me.
- Now: The number is 20. How many extras? Show me.

DO:
- Students show 8 fingers for 18
- Then 10 fingers (or "none -- it is two full tens") for 20
- Cold call to explain

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "Show me on your fingers how many extras. 18 has how many extras? Show me." Scan for: 8 fingers.
- Then: "20 has how many extras? Show me." Scan for: students showing 10 fingers or saying "none" or "zero extras, it is two tens."
PROCEED: If 80%+ show 8 for 18, and most students recognise that 20 is different, move to You Do.
PIVOT: Most likely misconception -- students show 0 fingers for 20 because they think "full = no extras." Clarify: "Good thinking! 20 has no LEFTOVER extras because the extras row is also full. 20 = 10 + 10. Two full tens."

TEACHER NOTES:
Two-part hinge checks both a standard teen number (18) and the conceptual outlier (20). The 20 question reveals depth of understanding.

WATCH FOR:
- Students who pause on 20 -- they are thinking, which is good
- Readiness signal: students explaining that 20 is "two tens" without prompting

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice time
- On your worksheet you will see ten frames for numbers 16 to 20
- First: Count the ten and the extras. Next: Write the number. Then: Draw your own ten frames.
- You have 6 minutes.

DO:
- Distribute Session 2 Worksheet
- Students work independently
- Circulate: start with students who struggled during We Do
- Distribute enabling scaffold or extension as needed

TEACHER NOTES:
You Do covers 16-20 with different arrangements from the We Do. Section B asks students to draw counters for given numbers, checking they can represent as well as identify.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Complete the enabling scaffold with pre-shaded ten frames for 16, 17, 18. The tens are already drawn; students count extras and write the number.
- Extra Notes: Distribute the Session 2 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Use the Session 2 Extension to review all teen numbers 11-20, then investigate what comes after 20 using two ten frames.
- Extra Notes: Self-contained investigation.

WATCH FOR:
- Students who represent 20 correctly with two full ten frames -- deep understanding
- Students who complete Section A quickly -- direct to Section B (drawing)
- Readiness signal: completing 5+ problems correctly within 4 minutes

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket time. Three quick questions.
- Work on your own. 2 minutes.

DO:
- Display exit ticket
- Students respond on whiteboards
- Sort responses: secure, developing, beginning

TEACHER NOTES:
Exit ticket assesses SC1 (Q1 -- say/write 16-20) and SC2 (Q2 -- ten frame identification). Q3 targets SC3 (explain 20 as two tens).

WATCH FOR:
- Students who get Q1-Q2 correct but struggle with Q3 -- naming is secure but not structure
- Students who explain 20 clearly -- ready for ordering in Session 3

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Let's check our success criteria.
- SC1: I can say the numbers 16, 17, 18, 19, and 20. Thumbs?
- SC2: I can show a teen number using a ten frame and extra ones. Thumbs?
- SC3: I can explain what each teen number means using tens and ones. Thumbs?
- Turn and talk: What is special about the number 20?

DO:
- Display SC on screen
- Run thumbs check
- Allow 30 seconds for Turn and Talk
- Cold call 2-3 students

TEACHER NOTES:
The Turn and Talk targets the key conceptual move: 20 = two tens. Students who can articulate this are ready for ordering the full 0-20 range in Session 3.

WATCH FOR:
- Students who say "two tens" or "both rows full" -- the concept has landed
- Students who are still unsure about 20 -- they need more support in future sessions

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- These are today's resources

DO:
- Point out each resource

TEACHER NOTES:
Resource slide provides clickable links to all companion PDFs.

WATCH FOR:
- N/A

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Session 2: Teen Numbers 16-20",
    "Foundation Numeracy | Session 2 of 5 | Term 2 Week 1", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2-3: Daily Review — Digit Formations (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Writing Teen Numbers", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      s.addText([
        { text: "Write these numbers on your whiteboard:", options: { fontSize: 15, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "1.  11", options: { fontSize: 20, color: C.CHARCOAL, breakLine: true, paraSpaceAfter: 6 } },
        { text: "2.  14", options: { fontSize: 20, color: C.CHARCOAL, breakLine: true, paraSpaceAfter: 6 } },
        { text: "3.  13", options: { fontSize: 20, color: C.CHARCOAL, breakLine: true, paraSpaceAfter: 6 } },
        { text: "4.  15", options: { fontSize: 20, color: C.CHARCOAL, breakLine: true, paraSpaceAfter: 6 } },
        { text: "5.  12", options: { fontSize: 20, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboards:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "Write each number carefully.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Tips:", options: { fontSize: 14, bold: true, color: C.ALERT, breakLine: true } },
        { text: "The 1 always comes first", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Make your numbers big and clear", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Take your time!", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Check: 1) 11   2) 14   3) 13   4) 15   5) 12   --  the 1 comes first!", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Backwards counting from 10
  const sFluency = pres.addSlide();
  addTopBar(sFluency, STAGE_COLORS["1"]);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Countdown: 10 to 1", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

  addCard(sFluency, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  sFluency.addText([
    { text: "Round 1: Count Back Together", options: { fontSize: 15, bold: true, color: C.PRIMARY, breakLine: true } },
    { text: "10... 9... 8... 7... 6...", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
    { text: "5... 4... 3... 2... 1...", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
    { text: "BLAST OFF!", options: { fontSize: 16, bold: true, color: C.ALERT, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "Round 2: Faster!", options: { fontSize: 15, bold: true, color: C.PRIMARY, breakLine: true } },
    { text: "Same countdown, double speed.", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "Round 3: Whisper Countdown", options: { fontSize: 15, bold: true, color: C.PRIMARY, breakLine: true } },
    { text: "Whisper the whole way down.", options: { fontSize: 13, color: C.CHARCOAL } },
  ], {
    x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  // Rocket countdown visual
  addCard(sFluency, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
  const countNums = [10, 9, 8, 7, 6, 5, 4, 3, 2, 1];
  const colWidth = 0.75;
  const rowHeight = 0.32;
  countNums.forEach((num, i) => {
    const col = i % 5;
    const row = Math.floor(i / 5);
    sFluency.addShape("roundRect", {
      x: 5.4 + col * colWidth,
      y: CONTENT_TOP + 0.15 + row * rowHeight * 2.2,
      w: colWidth - 0.06, h: rowHeight,
      rectRadius: 0.06,
      fill: { color: i < 5 ? C.PRIMARY : C.ACCENT },
    });
    sFluency.addText(String(num), {
      x: 5.4 + col * colWidth,
      y: CONTENT_TOP + 0.15 + row * rowHeight * 2.2,
      w: colWidth - 0.06, h: rowHeight,
      fontSize: 16, fontFace: FONT_H, color: C.WHITE,
      align: "center", valign: "middle", bold: true, margin: 0,
    });
  });
  // Blast off
  addTextOnShape(sFluency, "BLAST OFF!", {
    x: 5.6, y: CONTENT_TOP + 1.8, w: 3.2, h: 0.5, rectRadius: 0.1,
    fill: { color: C.ALERT },
  }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    ["We are learning to name and recognise the numbers 16 to 20"],
    [
      "I can say the numbers 16, 17, 18, 19, and 20",
      "I can show a teen number using a ten frame and extra ones",
      "I can explain what each teen number means using tens and ones",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — 16, 17, 18
  workedExSlide(pres, 2, "I Do", "Teen Numbers: 16, 17, 18",
    [
      "The pattern continues!",
      "",
      "16 = ten and 6 more (sixteen)",
      "   Can you hear SIX in sixteen?",
      "",
      "17 = ten and 7 more (seventeen)",
      "   Can you hear SEVEN in seventeen?",
      "",
      "18 = ten and 8 more (eighteen)",
      "   Can you hear EIGHT in eighteen?",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.3, { strip: C.PRIMARY });
      // 16
      slide.addText("16", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 0.08, w: 0.5, h: 0.35,
        fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      drawTenFrameWithExtras(slide, lg.rightX + 0.65, lg.panelTopPadded + 0.08, 1.3, 0.38, 16);

      // 17
      slide.addText("17", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 0.95, w: 0.5, h: 0.35,
        fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      drawTenFrameWithExtras(slide, lg.rightX + 0.65, lg.panelTopPadded + 0.95, 1.3, 0.38, 17);

      // 18
      slide.addText("18", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 1.82, w: 0.5, h: 0.35,
        fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      drawTenFrameWithExtras(slide, lg.rightX + 0.65, lg.panelTopPadded + 1.82, 1.3, 0.38, 18);
    }
  );

  // Slide 7: I Do — 19 and 20
  workedExSlide(pres, 2, "I Do", "The Big Ones: 19 and 20",
    [
      "19 = ten and 9 more (nineteen)",
      "   Almost a full second row!",
      "",
      "20 = ten and 10 more",
      "   BOTH rows are completely full!",
      "   20 = TWO tens",
      "   That is why it is called twenty",
      "   (twenty sounds like two-ty)",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.3, { strip: C.SECONDARY });
      // 19
      slide.addText("19", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 0.08, w: 0.5, h: 0.35,
        fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      drawTenFrameWithExtras(slide, lg.rightX + 0.65, lg.panelTopPadded + 0.08, 1.3, 0.38, 19);

      // 20 — two full ten frames side by side
      slide.addText("20 = TWO full tens!", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 1.2, w: lg.rightW - 0.2, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0, align: "center",
      });
      // First ten frame
      drawTenFrame(slide, lg.rightX + 0.2, lg.panelTopPadded + 1.55, 1.5, 0.44, 10, { dotColor: C.ACCENT });
      // Second ten frame
      drawTenFrame(slide, lg.rightX + 1.9, lg.panelTopPadded + 1.55, 1.5, 0.44, 10, { dotColor: C.ALERT });

      addTextOnShape(slide, "10  +  10  =  20", {
        x: lg.rightX + 0.4, y: lg.panelTopPadded + 2.2, w: 2.8, h: 0.35, rectRadius: 0.06,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 8-9: CFU (withReveal) — Show Me Boards
  withReveal(
    () => cfuSlide(pres, "CFU", "What Number Is This?", "Show Me Boards",
      "Look at the ten frame.\n\nFull ten + 7 extras.\n\nWrite the number on your whiteboard!",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "17!  Ten and 7 more = seventeen", {
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
      addTitle(s, "What Teen Number Is This?", { y: 0.65, fontSize: 22, color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });
      s.addText([
        { text: "With your partner:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Look at each ten frame.", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "Count the ten, count the extras.", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "Write the teen number.", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 10, breakLine: true } },
        { text: "A:  ten + ? = ?", options: { fontSize: 15, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "B:  ten + ? = ?", options: { fontSize: 15, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "C:  ten + ? = ?", options: { fontSize: 15, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "D:  ten + ? = ?", options: { fontSize: 15, bold: true, color: C.PRIMARY } },
      ], {
        x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });
      const frames = [
        { label: "A", count: 16 },
        { label: "B", count: 19 },
        { label: "C", count: 18 },
        { label: "D", count: 20 },
      ];
      frames.forEach((f, i) => {
        const frameY = CONTENT_TOP + 0.1 + i * 0.9;
        s.addText(f.label + ".", {
          x: 5.35, y: frameY, w: 0.35, h: 0.3,
          fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
        });
        if (f.count === 20) {
          // Two ten frames for 20
          drawTenFrame(s, 5.75, frameY, 0.9, 0.28, 10, { dotColor: C.ACCENT });
          drawTenFrame(s, 6.75, frameY, 0.9, 0.28, 10, { dotColor: C.ALERT });
        } else {
          drawTenFrameWithExtras(s, 5.75, frameY, 1.1, 0.32, f.count, { dotColor: C.ACCENT });
        }
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      runSlideDiagnostics(s, pres, { respectSafeBottom: false });
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "A) 16    B) 19    C) 18    D) 20 (two tens!)", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 12-13: CFU 2 Hinge (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Check", "Finger Voting",
      "18 has how many extras after the ten?\nShow me on your fingers!\n\n20 has how many extras?\nShow me!",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "18 = 8 extras    20 = special! Two full tens, no leftover extras", {
        x: 0.5, y: 4.0, w: 9, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do
  workedExSlide(pres, 4, "You Do", "Your Turn: Numbers 16-20",
    [
      "First: Look at each ten frame on your worksheet.",
      "Next: Count the ten, then count the extras.",
      "Then: Write the teen number.",
      "",
      "Section B: Draw counters to show the number.",
      "",
      "You have 6 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.6, { strip: C.ALERT });
      slide.addText("Remember:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Full ten frame = 10", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Count the extras", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "20 = two full tens!", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.5, h: 0.9,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "Say these numbers to your teacher: 16, 18, 20",
      "A ten frame shows a full ten and 9 extras. What number is it?",
      "What is special about the number 20? (Hint: how many tens?)",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 16: Closing
  closingSlide(pres,
    "Tell your partner: What is special about the number 20?",
    [
      "I can say the numbers 16, 17, 18, 19, and 20",
      "I can show a teen number using a ten frame and extra ones",
      "I can explain what each teen number means using tens and ones",
    ],
    NOTES_CLOSING);

  // Slide 17: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Teen_Session2_Numbers_16_to_20.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Generate PDFs ──────────────────────────────────────────────────────────

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, { subtitle: "Teen Numbers 16-20", color: C.PRIMARY, lessonInfo: "Session 2 of 5 | Foundation Numeracy" });
    y = addTipBox(doc, "Look at each ten frame. Count the full ten, then count the extras. Write the teen number.", y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section A: What number is this?", y, { color: C.PRIMARY });
    y = addProblem(doc, 1, "A full ten frame and 6 extra counters. What number? ___", y, { color: C.PRIMARY });
    y = addProblem(doc, 2, "A full ten frame and 8 extra counters. What number? ___", y, { color: C.PRIMARY });
    y = addProblem(doc, 3, "A full ten frame and 9 extra counters. What number? ___", y, { color: C.PRIMARY });
    y = addProblem(doc, 4, "TWO full ten frames (both rows full). What number? ___", y, { color: C.PRIMARY });
    y = addProblem(doc, 5, "A full ten frame and 7 extra counters. What number? ___", y, { color: C.PRIMARY });

    y = addSectionHeading(doc, "Section B: Draw the counters", y, { color: C.PRIMARY });
    y = addProblem(doc, 6, "Draw 18 counters using a ten frame and extras.", y, { color: C.PRIMARY, writeLines: [{ label: "How many in the ten?" }, { label: "How many extras?" }] });
    y = addProblem(doc, 7, "Draw 16 counters using a ten frame and extras.", y, { color: C.PRIMARY, writeLines: [{ label: "How many in the ten?" }, { label: "How many extras?" }] });
    y = addProblem(doc, 8, "Draw 20 counters. Hint: you need TWO ten frames!", y, { color: C.PRIMARY, writeLines: [{ label: "How many tens?" }] });

    addPdfFooter(doc, "Session 2 | Teen Numbers | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, { subtitle: "Teen Numbers 16-20", color: C.PRIMARY, lessonInfo: "Session 2 of 5 | Foundation Numeracy" });
    y = addSectionHeading(doc, "Section A", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1. 10 + 6 = 16   2. 10 + 8 = 18   3. 10 + 9 = 19   4. 10 + 10 = 20   5. 10 + 7 = 17", y);
    y = addSectionHeading(doc, "Section B", y, { color: C.PRIMARY });
    y = addBodyText(doc, "6. 18: Ten frame full (10), 8 extras", y);
    y = addBodyText(doc, "7. 16: Ten frame full (10), 6 extras", y);
    y = addBodyText(doc, "8. 20: Two full ten frames (10 + 10 = 20)", y);
    addPdfFooter(doc, "Session 2 | Answer Key | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, { subtitle: "Ten Frames for 16-18", color: C.ACCENT, lessonInfo: "Session 2 of 5 | Foundation Numeracy" });
    y = addTipBox(doc, "The ten frame is already shaded. Count the extra dots and write the number!", y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Count the extras and write the number", y, { color: C.PRIMARY });
    [16, 17, 18].forEach((num, i) => {
      y = addProblem(doc, i + 1, "Ten frame: 10 shaded.  Extra dots: " + (num - 10) + ".  This number is: ___", y, { color: C.PRIMARY });
    });

    addPdfFooter(doc, "Session 2 | Enabling Scaffold | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF: " + ENABLING_RES.fileName);
  })();

  // Extension
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, { subtitle: "All Teen Numbers and Beyond!", color: C.PRIMARY, lessonInfo: "Session 2 of 5 | Foundation Numeracy" });

    y = addSectionHeading(doc, "You Know All the Teen Numbers!", y, { color: C.PRIMARY });
    y = addBodyText(doc, "11 = ten and 1    12 = ten and 2    13 = ten and 3    14 = ten and 4    15 = ten and 5", y);
    y = addBodyText(doc, "16 = ten and 6    17 = ten and 7    18 = ten and 8    19 = ten and 9    20 = two tens", y);

    y = addSectionHeading(doc, "Match the Number to Its Ten Frame", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Draw a line from each number to the correct description:", y);
    y = addProblem(doc, 1, "14    ->    ten and ___ extras", y, { color: C.PRIMARY });
    y = addProblem(doc, 2, "17    ->    ten and ___ extras", y, { color: C.PRIMARY });
    y = addProblem(doc, 3, "20    ->    ___ full tens", y, { color: C.PRIMARY });
    y = addProblem(doc, 4, "11    ->    ten and ___ extras", y, { color: C.PRIMARY });

    y = addSectionHeading(doc, "What Comes After 20?", y, { color: C.PRIMARY });
    y = addBodyText(doc, "20 is two tens. What if we add one more counter?", y);
    y = addBodyText(doc, "Two tens and 1 more = 21 (twenty-one)", y);
    y = addBodyText(doc, "Two tens and 2 more = 22 (twenty-two)", y);
    y = addProblem(doc, 5, "Two tens and 3 more = ___ (twenty-___)", y, { color: C.PRIMARY });
    y = addProblem(doc, 6, "Two tens and 5 more = ___ (twenty-___)", y, { color: C.PRIMARY });

    y = addTipBox(doc, "Challenge: If 20 = two tens, what do you think 30 means? How many tens?", y, { color: C.ACCENT });

    addPdfFooter(doc, "Session 2 | Extension | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 2 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
