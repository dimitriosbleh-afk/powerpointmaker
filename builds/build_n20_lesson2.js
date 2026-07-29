"use strict";

// Numbers to 20 Unit — Session 2: Teen Numbers = Ten and More
// Term 2 Week 1, Foundation Numeracy, Variant 0
// DR: Digit Formations — I can write the teen numbers
// Fluency: Writing teen numbers (air-write, whiteboard practice)
// VC2MFN01 — name, represent and order numbers including zero to at least 20

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  formatSessionResourceFileName, addLinedArea,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "foundation", 0);
const {
  C, FONT_H, FONT_B,
  SLIDE_W, SLIDE_H, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addBadge, addTitle,
  addInstructionCard,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 2;
const UNIT_TITLE = "Numbers to 20";
const FOOTER = "Numbers to 20 | Session 2 of 5 | Foundation Numeracy";
const OUT_DIR = "output/N20_Session2_Teen_Numbers";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Independent practice - building and writing teen numbers.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with answers.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Pre-filled ten frames with numeral tracing guides for teen numbers.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Self-contained investigation into representing teen numbers multiple ways.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helper: draw a ten frame on a slide ─────────────────────────────────────
function drawTenFrame(slide, x, y, filled, opts) {
  const o = opts || {};
  const cellW = o.cellW || 0.42;
  const cellH = o.cellH || 0.42;
  const gap = 0.02;
  for (let row = 0; row < 2; row++) {
    for (let col = 0; col < 5; col++) {
      const idx = row * 5 + col;
      const cx = x + col * (cellW + gap);
      const cy = y + row * (cellH + gap);
      slide.addShape("roundRect", {
        x: cx, y: cy, w: cellW, h: cellH, rectRadius: 0.04,
        fill: { color: idx < filled ? (o.fillColor || C.SECONDARY) : C.WHITE },
        line: { color: o.borderColor || C.PRIMARY, width: 1.2 },
      });
    }
  }
  return { w: 5 * (cellW + gap) - gap, h: 2 * (cellH + gap) - gap };
}

// ─── Teacher Notes ───────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Today we are going to learn something exciting about teen numbers
- Teen numbers are special because they are all made of TEN and some more
- By the end of this lesson you will be able to build teen numbers

DO:
- Display title slide as students settle on the mat
- Have ten frames, counters, and whiteboards ready
- Pre-count bags of counters (at least 20 per pair)

TEACHER NOTES:
Session 2 of 5. This lesson introduces the place value structure of teen numbers (10 + ones) using ten frames as the key visual. Students worked on teen number recognition last week. This session builds the structural understanding that underpins place value.

WATCH FOR:
- Students who seem nervous about teen numbers -- remind them: "You already know how to count these. Today we are going to see how they are BUILT"

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up time! We are practising writing our teen numbers
- I will show you a number. Write it on your whiteboard
- Ready? Write each number as I call it out

DO:
- Call out numbers one at a time: 13, 17, 11, 15, 19
- Allow 10 seconds per number
- Students write on whiteboards, hold face down
- After all 5 numbers, say "Show me!"

TEACHER NOTES:
Daily Review retrieves digit formation for teen numbers -- the focus from last week. Common Foundation errors include reversing digits (31 for 13) and confusing similar-looking teens (16 vs 19). The dictation format tests recall and formation together.

WATCH FOR:
- Students who reverse digits (writing 31 for 13) -- this is very common. Note them for targeted digit formation support
- Students who write clearly and correctly -- digit formation is secure

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers -- tick if you got it right, fix it if you didn't
- Fixing it is the learning
- Ask: Which number was the trickiest to write? Why?

DO:
- Click to reveal the correct numerals
- Students tick correct and fix errors
- Discuss common tricky ones (11, 17, 19 are often confused)

TEACHER NOTES:
Tick-and-fix gives immediate feedback on digit formation. Note which students reversed digits -- they will benefit from the ten frame work in this lesson because seeing the structure (10 + ones) reinforces which digit comes first.

WATCH FOR:
- Students who reversed digits but self-corrected after seeing the answers -- metacognition is developing
- Students who made the same reversal on multiple numbers -- they need a consistent intervention strategy

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time! We are practising writing teen numbers
- I will say a number. You air-write it with your finger, then write it on your whiteboard
- Watch me first: 14. Air-write: 1... 4. Now I write it: 14
- Your turn. Ready?

DO:
- Call out: 12, 16, 18, 11, 14, 20, 13, 15, 17, 19
- For each: students air-write first, then write on whiteboard
- Keep the pace brisk -- about 8 seconds per number
- After all 10, students self-check against the number strip on the wall

TEACHER NOTES:
Air-writing engages motor memory alongside visual recall. The sequence includes all teen numbers to build speed and confidence. Having students air-write before board-writing adds a kinaesthetic rehearsal step.

WATCH FOR:
- Students who air-write the digits in the wrong order -- they are thinking about the digits but not the sequence
- Students who are fast and accurate -- automaticity with teen number formation is building

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read from slide: "We are learning that teen numbers are made of ten and some more"
- Let's read our success criteria together
- Read from slide: "I can say what a teen number is"
- Read from slide: "I can show a teen number using a ten frame and extra counters"
- Read from slide: "I can write the numeral for a teen number I have built"

DO:
- Choral read the LI then each SC, pointing to each
- Hold up a ten frame and counters: "These are our tools for today"

TEACHER NOTES:
This lesson targets the composition of teen numbers: 10 + ones. SC1 is the definition (ultra-achievable -- "a number bigger than 10 and less than 20"). SC2 is the concrete representation. SC3 bridges concrete to symbolic. The ten frame is the key manipulative because it makes the "10" visible and countable.

WATCH FOR:
- Students who seem unsure about what a ten frame is -- show one and count the spaces together

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me build the number 13
- First, I fill my ten frame. Count with me: 1, 2, 3, 4, 5, 6, 7, 8, 9, 10
- My ten frame is FULL. That means I have 10
- But I need 13. 13 is MORE than 10. So I need some extras
- How many more? 13 is 10 and... 3 more
- I put 3 counters outside the frame: 1, 2, 3
- So 13 is a full ten frame PLUS 3 more. 10 and 3 makes 13
- Now I write the number: 1 for the ten, 3 for the extras. 13!

DO:
- Use a physical ten frame and counters under the document camera
- Fill the frame slowly, counting aloud
- Place 3 extras beside the frame
- Write 13 on the whiteboard next to it
- Point: "The 1 tells us there is one ten. The 3 tells us there are 3 extras."

MISCONCEPTIONS:
- Misconception: Each counter in the ten frame represents the teen number (e.g. students think the 3 extra counters ARE 13)
  Why: Students may not yet see the ten frame as representing 10 -- they see it as a container, not a quantity
  Impact: Without understanding that a full ten frame = 10, students cannot compose or decompose teen numbers
  Quick correction: "Count the ten frame: 1, 2, 3... all the way to 10. The frame holds 10. Then we add the extras on top."

TEACHER NOTES:
This is the key conceptual move of the lesson. The ten frame makes 10 concrete and visible -- students can SEE the ten inside every teen number. The verbal pattern "10 and ___ more" is the sentence stem that bridges concrete to language.

WATCH FOR:
- Students who count along enthusiastically -- engagement is high
- Students who look confused when you place extras outside the frame -- they may not yet see the frame as "full = 10"
- Readiness signal: students saying "10 and 3 more" before you do

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Let me build another one. This time: 17
- First I fill my ten frame: 1, 2, 3... 10. Full!
- 17 is 10 and... how many more? 7 more!
- I put 7 extras: 1, 2, 3, 4, 5, 6, 7
- 10 and 7 makes 17. I write: 1, 7. 17!
- See the pattern? Every teen number starts with a full ten frame

DO:
- Build 17 with counters
- Emphasise the "10 and ___ more" pattern
- Write 17 next to the ten frame

TEACHER NOTES:
Second worked example reinforces the same structure with a different number. The repeated "10 and ___ more" pattern is building the schema. Two examples before CFU is appropriate for Foundation because the concept is new.

WATCH FOR:
- Students who call out "7 more!" before you ask -- the pattern is clicking
- Students who still look unsure -- they will benefit from hands-on practice in We Do

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check! I have built a number with my ten frame
- The ten frame is full and there are 5 extra counters
- What number did I build? Show me with your fingers... now!

DO:
- Display a full ten frame + 5 extras on screen
- Students show with fingers (should show 15 -- 10 + 5 fingers)
- Many students may show 1 hand (5) and the other hand (10) = 15

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "The ten frame is full -- that is 10. And 5 extra counters. What number is that? Show me with ALL your fingers. Ready... show me!"
- Scan for: students showing 15 (both hands up, one fully open = 10, other showing 5).
PROCEED: If 80%+ show 15 (some may hold up all 10 fingers then add 5 -- accept any clear 15), move to We Do.
PIVOT: Most likely misconception -- students show 5 because they only counted the extras, ignoring the ten frame. Reteach: "How many in the ten frame? Count with me: 1, 2, 3... 10. Plus 5 extras. 10 and 5 makes... 15."

TEACHER NOTES:
Finger voting with 15 is tricky for Foundation -- some students may not be able to show 15 on their fingers. Accept any clear attempt that shows understanding of 10 + 5. Alternatively, students can write 15 on their whiteboard.

WATCH FOR:
- Students who show 5 (only counting extras) -- they need the full ten frame count
- Students who show 15 confidently -- the composition concept is understood

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn! Each pair gets a ten frame and counters
- I will say a teen number. You build it!
- First: Fill the ten frame. Next: Add the extras. Then: Write the number on your whiteboard
- Ready? Build 14!

DO:
- Distribute ten frames and counters to pairs
- Call out: 14. Allow 60 seconds
- Circulate: check students are filling the frame first
- Cold call: "How many in the frame? [10] How many extras? [4] What number? [14]"
- Repeat with 16 and 11

TEACHER NOTES:
We Do transitions from watching to doing with the same manipulatives. Three numbers (14, 16, 11) provide practice. 11 is deliberately included because "10 and 1 more" is the simplest teen composition.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 2 Enabling Scaffold with pre-filled ten frames. Students only need to add the extras and write the number. The frame is already shown as full (10) with dots pre-printed.
- Extra Notes: Distribute the Session 2 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Build each teen number AND write a number sentence: 10 + ___ = ___. For example, 10 + 4 = 14. Try writing the sentence for every teen number from 11 to 19.
- Extra Notes: Distribute the Session 2 Extension PDF.

WATCH FOR:
- Students who forget to fill the ten frame first and start counting all counters from 1 -- redirect: "Fill the frame first. That gives you 10."
- Students who build correctly and write the numeral with confidence -- they are ready for the extension
- Readiness signal: pairs completing all three numbers with correct numerals

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check. 14: ten frame full plus how many extras? [4] Number? [14]
- 16: ten frame full plus? [6] Number? [16]
- 11: ten frame full plus? [1] Number? [11]
- What is the pattern? [Every teen number is 10 and some more]

DO:
- Reveal answers on screen
- Have students hold up their whiteboards to show their numerals
- Celebrate the "10 and ___ more" pattern

TEACHER NOTES:
The key learning to draw out is the generalisation: ALL teen numbers = 10 + ones. This pattern is the foundation of place value understanding.

WATCH FOR:
- Students who can articulate "10 and ___ more" -- the composition concept has landed
- Students who got the building right but wrote the numeral incorrectly -- building is secure, formation needs more practice

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. I am going to say a number. You tell me how many extras go OUTSIDE the ten frame
- Ready? The number is 18. How many extras? Show me with your fingers

DO:
- Students show 8 fingers
- Follow up: "How do you know?" Cold call 2 students

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "If the number is 18 and the ten frame is full, how many extras? Show me with your fingers. Ready... show me!"
- Scan for: 8 fingers. Students who show 8 quickly can decompose 18 into 10 + 8.
PROCEED: If 80%+ show 8 fingers, move to You Do.
PIVOT: Most likely misconception -- students show a random number because they do not see the relationship between the teen number and the extras. Reteach: "18 is 10 and... let's count from 10. 11, 12, 13, 14, 15, 16, 17, 18. That is 8 more after 10."

TEACHER NOTES:
This hinge reverses the direction: instead of building from 10 + extras, students decompose the teen number to find the ones. This is a harder cognitive move and confirms genuine understanding.

WATCH FOR:
- Students who show 8 instantly -- they understand composition/decomposition
- Students who need to count up from 10 on their fingers -- this is valid and shows developing understanding

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice time. Open your worksheet
- First: Look at each ten frame picture
- Next: Count the full frame (10) and the extras
- Then: Write the teen number
- Section B: Build a number I give you -- draw dots in the frame and extras
- You have 6 minutes. Go!

DO:
- Distribute Session 2 Worksheet
- Students work independently
- Circulate: prioritise students who struggled during We Do
- For enabling students, distribute the Session 2 Enabling Scaffold
- For extending students, distribute the Session 2 Extension

TEACHER NOTES:
You Do uses pictures of ten frames (not physical ones) to check transfer from concrete to pictorial. Section A is read-and-write (given the representation, write the numeral). Section B is draw-and-write (given the numeral, create the representation). This is a harder direction.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 2 Enabling Scaffold with pre-drawn ten frames (full) and dotted numeral guides. Students count the pre-drawn extras and trace the numeral.
- Extra Notes: Distribute the Session 2 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Investigate representing teen numbers multiple ways. The extension PDF shows each teen number as 10 + ones and asks students to write the number sentence, draw a ten frame representation, and show it on a number line.
- Extra Notes: Distribute the Session 2 Extension PDF. Self-contained with worked examples.

WATCH FOR:
- Students who count all dots from 1 instead of recognising the full frame as 10 -- redirect: "Is the frame full? Then that is 10. Just count the extras."
- Students who draw the correct number of extras in Section B -- strong understanding
- Readiness signal: 80%+ completing Section A correctly within 3 minutes

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket time. Three quick questions
- Work on your own. You have 2 minutes

DO:
- Display the exit ticket on screen
- Students write answers in workbooks or on whiteboards
- Collect responses to sort

TEACHER NOTES:
Exit ticket assesses SC2 (Q1 -- read a ten frame) and SC3 (Q2-Q3 -- write the numeral and decompose). Q3 asks "15 is 10 and how many more?" which tests the composition concept directly.

WATCH FOR:
- Students who get Q1-Q2 correct but struggle with Q3 -- they can read ten frames but have not yet internalised the composition pattern
- Students who complete all three instantly and correctly -- ready for representing numbers multiple ways in Session 3

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Let's check our success criteria. Look at the slide
- SC1: I can say what a teen number is. Thumbs up, sideways, or down
- SC2: I can show a teen number using a ten frame and extra counters. Thumbs?
- SC3: I can write the numeral for a teen number I have built. Thumbs?
- Turn and talk: Tell your partner -- what are teen numbers made of? [10 and some more]

DO:
- Display success criteria on screen
- Run thumbs check for each SC
- Allow 30 seconds for Turn and Talk
- Cold call 2 students: "What are teen numbers made of?"

TEACHER NOTES:
The key response is "10 and some more" (or equivalent). Students who can articulate this have grasped the core concept. This understanding is the foundation for all place value work going forward.

WATCH FOR:
- Students who say "10 and some more" confidently -- the composition concept has landed
- Students who show thumbs down on SC2 -- they need more physical ten frame practice in Session 3

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- These are the resources for today's lesson

DO:
- Point out each resource and its purpose

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
  titleSlide(pres, UNIT_TITLE, "Session 2: Teen Numbers = Ten and More", "Foundation Numeracy | Session 2 of 5 | Term 2 Week 1", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2-3: Daily Review (withReveal) — Writing Teen Numbers
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Write the Teen Numbers", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 9, 2.4, { strip: STAGE_COLORS["1"] });
      // Show 5 numbers to write
      const drNums = ["13", "17", "11", "15", "19"];
      drNums.forEach((num, i) => {
        const col = i;
        const cx = 1.0 + col * 1.7;
        s.addShape("roundRect", {
          x: cx, y: CONTENT_TOP + 0.2, w: 1.2, h: 0.8, rectRadius: 0.08,
          fill: { color: C.PRIMARY },
        });
        s.addText(num, {
          x: cx, y: CONTENT_TOP + 0.2, w: 1.2, h: 0.8,
          fontSize: 30, fontFace: FONT_H, color: C.WHITE,
          align: "center", valign: "middle", bold: true, margin: 0,
        });
        // Write line below
        s.addShape("line", {
          x: cx + 0.1, y: CONTENT_TOP + 1.3, w: 1.0, h: 0,
          line: { color: C.MUTED, width: 1.5 },
        });
        s.addText("Write it!", {
          x: cx, y: CONTENT_TOP + 1.45, w: 1.2, h: 0.3,
          fontSize: 10, fontFace: FONT_B, color: C.MUTED, align: "center", margin: 0,
        });
      });

      addCard(s, 0.5, CONTENT_TOP + 2.6, 9, 0.9, { strip: C.ACCENT });
      s.addText([
        { text: "On your whiteboards:", options: { fontSize: 16, bold: true, color: C.ACCENT, breakLine: true } },
        { text: "Write each teen number as I call it out. Hold your board face down until I say show me!", options: { fontSize: 14, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 2.72, w: 8.5, h: 0.7,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Answers: 13   17   11   15   19", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Writing teen numbers (air-write + whiteboard)
  const sFluency = pres.addSlide();
  addTopBar(sFluency, STAGE_COLORS["1"]);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Air-Write and Board-Write", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

  addCard(sFluency, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  const fluencyNums = ["12", "16", "18", "11", "14", "20", "13", "15", "17", "19"];
  sFluency.addText(fluencyNums.map((n, i) => ({
    text: (i + 1) + ".  " + n,
    options: { fontSize: 18, color: C.CHARCOAL, breakLine: i < fluencyNums.length - 1, paraSpaceAfter: 2 },
  })), {
    x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
    fontFace: FONT_H, margin: 0, valign: "top",
  });

  addCard(sFluency, 5.2, CONTENT_TOP, 4.3, 2.0, { strip: C.SECONDARY });
  sFluency.addText([
    { text: "Steps:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
    { text: "1. I say the number", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
    { text: "2. You AIR-WRITE with your finger", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
    { text: "3. You WRITE it on your whiteboard", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
    { text: "4. Self-check against the number strip", options: { fontSize: 14, color: C.CHARCOAL } },
  ], {
    x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: 1.8,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    ["We are learning that teen numbers are made of ten and some more"],
    [
      "I can say what a teen number is",
      "I can show a teen number using a ten frame and extra counters",
      "I can write the numeral for a teen number I have built",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — Building 13
  workedExSlide(pres, 2, "I Do", "Building a Teen Number: 13",
    [
      "Step 1: Fill the ten frame",
      "  Count: 1, 2, 3... 10. FULL!",
      "",
      "Step 2: Add the extras",
      "  13 is 10 and 3 more",
      "  Put 3 counters outside",
      "",
      "Step 3: Write the number",
      "  1 for the ten, 3 for the extras = 13",
    ],
    NOTES_IDO, FOOTER,
    (slide, lg) => {
      // Ten frame visual showing 13
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.8, { strip: C.PRIMARY });
      slide.addText("Building 13", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      // Ten frame (full = 10)
      const tfX = lg.rightX + 0.5;
      const tfY = lg.panelTopPadded + 0.5;
      drawTenFrame(slide, tfX, tfY, 10);

      // Label "10"
      addTextOnShape(slide, "= 10", {
        x: tfX + 2.3, y: tfY + 0.15, w: 0.8, h: 0.4, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });

      // 3 extra counters
      const extraY = tfY + 1.1;
      slide.addText("+ extras:", {
        x: tfX, y: extraY, w: 1.0, h: 0.3,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });
      for (let i = 0; i < 3; i++) {
        slide.addShape("roundRect", {
          x: tfX + 1.0 + i * 0.5, y: extraY, w: 0.4, h: 0.4, rectRadius: 0.2,
          fill: { color: C.ALERT },
        });
      }
      addTextOnShape(slide, "= 3", {
        x: tfX + 2.6, y: extraY, w: 0.6, h: 0.4, rectRadius: 0.06,
        fill: { color: C.ALERT },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });

      // Answer
      addTextOnShape(slide, "10 + 3 = 13", {
        x: lg.rightX + 0.5, y: lg.panelTopPadded + 2.2, w: lg.rightW - 1.0, h: 0.45, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 7: I Do — Building 17
  workedExSlide(pres, 2, "I Do", "Building a Teen Number: 17",
    [
      "Step 1: Fill the ten frame",
      "  Count to 10. FULL!",
      "",
      "Step 2: Add the extras",
      "  17 is 10 and 7 more",
      "  Put 7 counters outside",
      "",
      "Step 3: Write the number",
      "  1 for the ten, 7 for the extras = 17",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.8, { strip: C.PRIMARY });
      slide.addText("Building 17", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      const tfX = lg.rightX + 0.5;
      const tfY = lg.panelTopPadded + 0.5;
      drawTenFrame(slide, tfX, tfY, 10);

      addTextOnShape(slide, "= 10", {
        x: tfX + 2.3, y: tfY + 0.15, w: 0.8, h: 0.4, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });

      // 7 extra counters (row of 5 + row of 2)
      const extraY = tfY + 1.1;
      slide.addText("+ extras:", {
        x: tfX, y: extraY, w: 1.0, h: 0.3,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      });
      for (let i = 0; i < 5; i++) {
        slide.addShape("roundRect", {
          x: tfX + 1.0 + i * 0.4, y: extraY, w: 0.35, h: 0.35, rectRadius: 0.17,
          fill: { color: C.ALERT },
        });
      }
      for (let i = 0; i < 2; i++) {
        slide.addShape("roundRect", {
          x: tfX + 1.0 + i * 0.4, y: extraY + 0.4, w: 0.35, h: 0.35, rectRadius: 0.17,
          fill: { color: C.ALERT },
        });
      }
      addTextOnShape(slide, "= 7", {
        x: tfX + 2.6, y: extraY + 0.1, w: 0.6, h: 0.4, rectRadius: 0.06,
        fill: { color: C.ALERT },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });

      addTextOnShape(slide, "10 + 7 = 17", {
        x: lg.rightX + 0.5, y: lg.panelTopPadded + 2.2, w: lg.rightW - 1.0, h: 0.45, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 8-9: CFU 1 (withReveal) — Full ten frame + 5 extras
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["2"]);
      addStageBadge(s, 2, "I Do");
      addTitle(s, "What Number Did I Build?", { y: 0.65, fontSize: 22, color: STAGE_COLORS["2"] });

      addCard(s, 0.5, CONTENT_TOP, 9, 2.2, { strip: C.PRIMARY });
      // Full ten frame
      drawTenFrame(s, 1.0, CONTENT_TOP + 0.3, 10, { cellW: 0.5, cellH: 0.5 });
      // 5 extras
      s.addText("+", {
        x: 3.8, y: CONTENT_TOP + 0.5, w: 0.5, h: 0.5,
        fontSize: 30, fontFace: FONT_H, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
      });
      for (let i = 0; i < 5; i++) {
        s.addShape("roundRect", {
          x: 4.5 + i * 0.55, y: CONTENT_TOP + 0.55, w: 0.45, h: 0.45, rectRadius: 0.22,
          fill: { color: C.ALERT },
        });
      }
      s.addText("?", {
        x: 7.5, y: CONTENT_TOP + 0.4, w: 1.0, h: 0.7,
        fontSize: 48, fontFace: FONT_H, color: C.ACCENT, align: "center", valign: "middle", bold: true, margin: 0,
      });

      addCard(s, 0.5, CONTENT_TOP + 2.5, 9, 0.9, { strip: C.ACCENT });
      s.addText("Show me with your FINGERS! How many altogether?", {
        x: 0.75, y: CONTENT_TOP + 2.62, w: 8.5, h: 0.65,
        fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, bold: true, margin: 0, valign: "middle",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "15!  (10 + 5 = 15)", {
        x: 2.0, y: 4.2, w: 6.0, h: 0.6, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 10-11: We Do (withReveal) — Build teen numbers with counters
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Build These Teen Numbers!",
      [
        "With your partner:",
        "",
        "1. Fill the ten frame (10)",
        "2. Add the extras outside",
        "3. Write the number on your whiteboard",
        "",
        "Build: 14, then 16, then 11",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.0, { strip: C.SECONDARY });
        slide.addText("10 and ___ more?", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.35,
          fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
        });

        // Three target numbers
        const targets = [
          { num: "14", extras: "4" },
          { num: "16", extras: "6" },
          { num: "11", extras: "1" },
        ];
        targets.forEach((t, i) => {
          addTextOnShape(slide, t.num + "  =  10 + ?", {
            x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55 + i * 0.5, w: lg.rightW - 0.6, h: 0.4, rectRadius: 0.06,
            fill: { color: C.PRIMARY },
          }, { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true });
        });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "14 = 10 + 4     16 = 10 + 6     11 = 10 + 1", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 12-13: CFU 2 Hinge (withReveal) — How many extras for 18?
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Check", "Finger Voting",
      "The number is 18.\nThe ten frame is full.\n\nHow many EXTRAS go outside?\nShow me with your fingers!",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "8 extras!  18 = 10 + 8", {
        x: 1.5, y: 4.0, w: 7.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do
  workedExSlide(pres, 4, "You Do", "Teen Number Practice",
    [
      "First: Look at each ten frame picture.",
      "Next: Count the full frame (10) and extras.",
      "Then: Write the teen number.",
      "",
      "Section B: I say a number, you draw",
      "the ten frame and extras.",
      "",
      "You have 6 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.6, { strip: C.ALERT });
      slide.addText("Remember:", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });

      // Mini ten frame visual
      drawTenFrame(slide, lg.rightX + 0.3, lg.panelTopPadded + 0.5, 10, { cellW: 0.3, cellH: 0.3 });
      slide.addText("= 10", {
        x: lg.rightX + 2.0, y: lg.panelTopPadded + 0.6, w: 0.6, h: 0.3,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
      });
      slide.addText("+ extras = teen number!", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.15, w: lg.rightW - 0.6, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0,
      });
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "This ten frame is full with 2 extras. What number?",
      "Write the number for: ten frame full + 6 extras.",
      "15 is 10 and how many more?  ___",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 16: Closing
  closingSlide(pres,
    "Tell your partner: What are teen numbers made of?",
    [
      "I can say what a teen number is",
      "I can show a teen number using a ten frame and extra counters",
      "I can write the numeral for a teen number I have built",
    ],
    NOTES_CLOSING);

  // Slide 17: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "N20_Session2_Teen_Numbers.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Generate PDFs ─────────────────────────────────────────────────────────

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Teen Numbers = Ten and More",
      color: C.NAVY,
      lessonInfo: "Session 2 of 5 | Foundation Numeracy",
    });
    y = addTipBox(doc, "A full ten frame = 10. Count the extras. Write the teen number!", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A: Read the Ten Frame", y, { color: C.NAVY });

    // Helper to draw a PDF ten frame with extras
    const drawPdfTenFrame = (doc, filled, extras, taskNum, y) => {
      doc.fontSize(12).font("Sans-Bold").fillColor("#333333").text(taskNum + ".", 55, y + 8);
      const cellSize = 22;
      const gap = 2;
      const startX = 80;
      // 2 rows x 5 cols
      for (let row = 0; row < 2; row++) {
        for (let col = 0; col < 5; col++) {
          const cx = startX + col * (cellSize + gap);
          const cy = y + row * (cellSize + gap);
          doc.rect(cx, cy, cellSize, cellSize).lineWidth(1).strokeColor("#" + C.NAVY).stroke();
          const idx = row * 5 + col;
          if (idx < filled) {
            doc.circle(cx + cellSize / 2, cy + cellSize / 2, 8).fill("#" + C.SECONDARY);
          }
        }
      }
      // Extra counters
      if (extras > 0) {
        const extraX = startX + 5 * (cellSize + gap) + 15;
        doc.fontSize(16).font("Sans-Bold").fillColor("#333333").text("+", extraX - 15, y + 10);
        for (let i = 0; i < extras; i++) {
          doc.circle(extraX + i * 22, y + cellSize / 2, 8).fill("#" + C.ALERT);
        }
      }
      const ansX = startX + 5 * (cellSize + gap) + 15 + Math.max(extras, 1) * 22 + 20;
      doc.fontSize(14).font("Sans").fillColor("#333333").text("= _____", ansX, y + 10);
      return y + cellSize * 2 + gap + 18;
    };

    y = drawPdfTenFrame(doc, 10, 3, "1", y);   // 13
    y = drawPdfTenFrame(doc, 10, 6, "2", y);   // 16
    y = drawPdfTenFrame(doc, 10, 1, "3", y);   // 11
    y = drawPdfTenFrame(doc, 10, 8, "4", y);   // 18
    y = drawPdfTenFrame(doc, 10, 4, "5", y);   // 14

    y = addSectionHeading(doc, "Section B: Build the Number", y, { color: C.NAVY });
    y = addBodyText(doc, "Draw dots in the ten frame and add extras to make each number.", y);

    const drawBlankTenFrame = (doc, targetNum, taskNum, y) => {
      doc.fontSize(12).font("Sans-Bold").fillColor("#333333").text(taskNum + ".  Make " + targetNum, 55, y + 8);
      const cellSize = 22;
      const gap = 2;
      const startX = 180;
      for (let row = 0; row < 2; row++) {
        for (let col = 0; col < 5; col++) {
          doc.rect(startX + col * (cellSize + gap), y + row * (cellSize + gap), cellSize, cellSize)
            .lineWidth(1).strokeColor("#" + C.NAVY).stroke();
        }
      }
      doc.fontSize(14).font("Sans").fillColor("#333333").text("+ extras: _____     = _____", startX + 5 * (cellSize + gap) + 10, y + 10);
      return y + cellSize * 2 + gap + 18;
    };

    y = drawBlankTenFrame(doc, 15, "6", y);
    y = drawBlankTenFrame(doc, 12, "7", y);
    y = drawBlankTenFrame(doc, 19, "8", y);
    y = drawBlankTenFrame(doc, 17, "9", y);

    addPdfFooter(doc, "Session 2 | Numbers to 20 | Foundation");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Teen Numbers = Ten and More",
      color: C.NAVY,
      lessonInfo: "Session 2 of 5 | Foundation Numeracy",
    });
    y = addSectionHeading(doc, "Section A", y, { color: C.NAVY });
    y = addBodyText(doc, "1. Full ten frame + 3 extras = 13", y);
    y = addBodyText(doc, "2. Full ten frame + 6 extras = 16", y);
    y = addBodyText(doc, "3. Full ten frame + 1 extra = 11", y);
    y = addBodyText(doc, "4. Full ten frame + 8 extras = 18", y);
    y = addBodyText(doc, "5. Full ten frame + 4 extras = 14", y);
    y = addSectionHeading(doc, "Section B", y, { color: C.NAVY });
    y = addBodyText(doc, "6. Make 15: full frame + 5 extras", y);
    y = addBodyText(doc, "7. Make 12: full frame + 2 extras", y);
    y = addBodyText(doc, "8. Make 19: full frame + 9 extras", y);
    y = addBodyText(doc, "9. Make 17: full frame + 7 extras", y);
    addPdfFooter(doc, "Session 2 | Answer Key | Foundation");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold — Pre-filled ten frames with numeral tracing
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Teen Numbers - Ten Frame Support",
      color: C.TEAL,
      lessonInfo: "Session 2 of 5 | Foundation Numeracy",
    });
    y = addTipBox(doc, "The ten frame is already full (10). Count the extras. Write the number by tracing the dots!", y, { color: C.TEAL });

    // Pre-filled ten frames with extras already drawn, dotted numeral to trace
    const drawPrefilledTask = (doc, extras, numeral, taskNum, y) => {
      doc.fontSize(12).font("Sans-Bold").fillColor("#333333").text(taskNum + ".", 55, y + 8);
      const cellSize = 22;
      const gap = 2;
      const startX = 80;
      // Pre-filled ten frame
      for (let row = 0; row < 2; row++) {
        for (let col = 0; col < 5; col++) {
          const cx = startX + col * (cellSize + gap);
          const cy = y + row * (cellSize + gap);
          doc.rect(cx, cy, cellSize, cellSize).lineWidth(1).strokeColor("#" + C.NAVY).stroke();
          doc.circle(cx + cellSize / 2, cy + cellSize / 2, 8).fill("#" + C.SECONDARY);
        }
      }
      doc.fontSize(10).font("Sans").fillColor("#" + C.NAVY).text("10", startX + 2.5 * (cellSize + gap) - 6, y + 2 * (cellSize + gap) + 4);
      // Extras
      const extraX = startX + 5 * (cellSize + gap) + 20;
      doc.fontSize(14).font("Sans-Bold").fillColor("#333333").text("+", extraX - 15, y + 10);
      for (let i = 0; i < extras; i++) {
        doc.circle(extraX + i * 22, y + cellSize / 2, 8).fill("#" + C.ALERT);
      }
      doc.fontSize(10).font("Sans").fillColor("#" + C.ALERT).text(String(extras), extraX + extras * 11 - 3, y + 2 * (cellSize + gap) + 4);
      // Dotted numeral to trace
      const traceX = extraX + extras * 22 + 30;
      doc.fontSize(36).font("Sans").fillColor("#CCCCCC").text(numeral, traceX, y - 2);
      doc.fontSize(12).font("Sans").fillColor("#999999").text("trace it!", traceX, y + 38);
      return y + cellSize * 2 + gap + 25;
    };

    y = addSectionHeading(doc, "Count the extras and trace the number", y, { color: C.NAVY });
    y = drawPrefilledTask(doc, 2, "12", "1", y);
    y = drawPrefilledTask(doc, 5, "15", "2", y);
    y = drawPrefilledTask(doc, 3, "13", "3", y);
    y = drawPrefilledTask(doc, 7, "17", "4", y);
    y = drawPrefilledTask(doc, 1, "11", "5", y);
    y = drawPrefilledTask(doc, 9, "19", "6", y);

    addPdfFooter(doc, "Session 2 | Enabling Scaffold | Foundation");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Extension — Representing teen numbers multiple ways
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Teen Numbers Many Ways",
      color: C.TEAL,
      lessonInfo: "Session 2 of 5 | Foundation Numeracy",
    });

    y = addSectionHeading(doc, "How many ways can you show a teen number?", y, { color: C.NAVY });
    y = addBodyText(doc, "Every teen number is 10 and some more. You can show this in different ways!", y);

    y = addSectionHeading(doc, "Example: The number 14", y, { color: C.NAVY });
    y = addBodyText(doc, "Way 1 - Ten frame: Full frame + 4 extras", y);
    y = addBodyText(doc, "Way 2 - Number sentence: 10 + 4 = 14", y);
    y = addBodyText(doc, "Way 3 - Words: fourteen = ten and four more", y);
    y = addBodyText(doc, "Way 4 - Number line: mark 14 on the line between 0 and 20", y);

    y = addSectionHeading(doc, "Your Turn", y, { color: C.NAVY });
    y = addBodyText(doc, "Show each number THREE different ways.", y);

    const drawExtTask = (doc, num, taskNum, y) => {
      doc.fontSize(14).font("Sans-Bold").fillColor("#" + C.NAVY).text(taskNum + ". Show the number " + num, 55, y);
      y += 20;
      // Ten frame space
      const cellSize = 18;
      const gap = 2;
      doc.fontSize(10).font("Sans").fillColor("#333333").text("Ten frame:", 70, y + 4);
      for (let row = 0; row < 2; row++) {
        for (let col = 0; col < 5; col++) {
          doc.rect(145 + col * (cellSize + gap), y + row * (cellSize + gap), cellSize, cellSize)
            .lineWidth(0.8).strokeColor("#" + C.NAVY).stroke();
        }
      }
      doc.text("+ extras", 145 + 5 * (cellSize + gap) + 5, y + 10);
      y += cellSize * 2 + gap + 10;
      // Number sentence
      doc.fontSize(10).font("Sans").fillColor("#333333").text("Number sentence:  10  +  ___  =  ___", 70, y);
      y += 20;
      // Words
      doc.text("In words: ___________________________________", 70, y);
      y += 25;
      return y;
    };

    y = drawExtTask(doc, "16", "1", y);
    y = drawExtTask(doc, "13", "2", y);
    y = drawExtTask(doc, "18", "3", y);
    y = drawExtTask(doc, "11", "4", y);

    y = addSectionHeading(doc, "Challenge", y, { color: C.NAVY });
    y = addBodyText(doc, "Can you put these teen numbers in order from smallest to biggest?", y);
    y = addBodyText(doc, "15,  11,  19,  13,  17", y);
    y = addWriteLine(doc, "Order:", y);

    addPdfFooter(doc, "Session 2 | Extension | Foundation");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 2 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
