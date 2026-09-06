"use strict";

// Teen Numbers Unit — Session 1: Introducing Teen Numbers (11-15)
// Term 2 Week 1, Foundation Numeracy, Variant 0
// DR: Counting — I can count up to at least 20 fluently by 1s
// Fluency: Counting from 1-20
// VC2MFN01 — name, represent and order numbers including zero to at least 20

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "foundation", 0);
const {
  C, FONT_H, FONT_B,
  SLIDE_W, SLIDE_H, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addBadge, addTitle,
  withReveal, runSlideDiagnostics, addNumberLine,
  STAGE_COLORS, STAGE_LABELS, addInstructionCard,
} = T;

const SESSION = 1;
const UNIT_TITLE = "Teen Numbers";
const FOOTER = "Teen Numbers | Session 1 of 5 | Foundation Numeracy";
const OUT_DIR = "output/Teen_Session1_Introducing_11_to_15";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Ten frame practice for teen numbers 11-15.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with completed ten frames.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Pre-drawn ten frames with visual support.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Self-contained investigation: teen numbers beyond 15.");

const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helper: draw a ten frame on a slide ───────────────────────────────────

function drawTenFrame(slide, x, y, w, h, filled, opts) {
  const o = opts || {};
  const cols = 5;
  const rows = 2;
  const cellW = w / cols;
  const cellH = h / rows;
  const fillColor = o.fillColor || C.PRIMARY;
  const emptyColor = o.emptyColor || C.WHITE;
  const borderColor = o.borderColor || C.PRIMARY;

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const idx = r * cols + c;
      const cx = x + c * cellW;
      const cy = y + r * cellH;
      slide.addShape("rect", {
        x: cx, y: cy, w: cellW, h: cellH,
        fill: { color: idx < filled ? fillColor : emptyColor },
        line: { color: borderColor, width: 1.5 },
      });
      if (idx < filled) {
        // Draw a dot/counter in the cell
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
  // Draw a full ten frame + extra counters below for numbers > 10
  const o = opts || {};
  const tenFrameFilled = Math.min(total, 10);
  drawTenFrame(slide, x, y, w, h, tenFrameFilled, o);

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
- Welcome back! We have been on holidays and now we are ready to learn
- This week we are going to learn about teen numbers
- Teen numbers are the numbers from 11 to 20

DO:
- Display title slide as students settle on the mat
- Have counters and ten frames ready

TEACHER NOTES:
Session 1 of 5. First lesson back from holidays -- keep the pace gentle and achievable. This session introduces teen numbers 11-15 only, building the concept that teen means "ten and some more."

WATCH FOR:
- Students who seem unsettled after the break -- use a calm, welcoming tone to re-establish routines

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Let's warm up our counting brains
- I want you to count from 1 to 20 with me
- Then try the counting challenges on your own

DO:
- Lead a choral count 1-20 first (whole class together)
- Then display the challenges
- Students show answers on whiteboards or fingers

TEACHER NOTES:
Daily Review retrieves counting fluency from Term 1. After a 16-day break, some students may be rusty. The choral count rebuilds confidence before individual challenges.

WATCH FOR:
- Students who hesitate or lose track after 12 or 13 -- teen numbers are the common stumbling point
- Students who count fluently to 20 -- counting automaticity is intact

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers -- tick if you got it right, fix it if you did not
- Fixing it is the learning
- Ask: What number comes after 13? [14]

DO:
- Click to reveal answers
- Students tick correct responses and fix errors
- Scan for students who struggled with the teen numbers specifically

TEACHER NOTES:
Tick-and-fix gives immediate feedback on counting recall. Note which students hesitate in the 11-19 range -- this unit targets exactly that gap.

WATCH FOR:
- Students who got all answers correct quickly -- counting to 20 is secure
- Students who made errors in the teen range (11-19) -- they are the target learners for this unit

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time -- let's get our counting muscles warmed up
- We are going to count together from 1 to 20
- Clap on every number. Ready? 1, 2, 3...
- Now let's try counting from 1 to 20 but WHISPER the numbers 1 to 10 and SHOUT the numbers 11 to 20

DO:
- Lead clap-and-count 1-20 (whole class, together)
- Then whisper/shout activity to highlight the teen numbers
- Repeat 2-3 times, getting faster each time

TEACHER NOTES:
Fluency builds counting automaticity. The whisper/shout activity draws attention to the shift at 11 -- priming students for the lesson focus on teen numbers. This is about rhythm and confidence, not new teaching.

WATCH FOR:
- Students who cannot keep the clapping rhythm -- they may need counting support
- Students who shout with confidence from 11-20 -- counting is automated

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention with me
- Read from slide: "We are learning to name and recognise the teen numbers 11 to 15"
- Now our success criteria. Read each one.
- Ask: Has anyone heard the word "teen" before? What do you think it means?

DO:
- Choral read the LI then each SC
- Accept student ideas about what "teen" means
- Do not correct yet -- the I Do will build the understanding

TEACHER NOTES:
This is the gateway to the teen numbers unit. The word "teen" signals "ten and more." SC1 is the floor -- every student should be able to say the numbers. SC2 is the core target. SC3 extends to conceptual understanding of place value at Foundation level.

WATCH FOR:
- Students who already know what teen means -- prior knowledge is present
- Students who look blank -- they are meeting this concept for the first time, and that is okay

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO1 = `SAY:
- Watch me. I am going to show you what teen numbers are
- Teen numbers are numbers that have a TEN and some more
- Look at this ten frame. It is full -- that is 10
- Now I add one more counter. 10 and 1 more makes... 11. Eleven.
- 11 is a teen number. It means ten and one more
- Now watch. I add another counter. 10 and 2 more makes... 12. Twelve.

DO:
- Point to the ten frame visual on screen
- Physically demonstrate with real counters if possible
- Say each number name slowly and clearly
- Emphasise "ten and ___ more" for each number

TEACHER NOTES:
This I Do establishes the core concept: teen = ten + extra ones. The ten frame makes this visible and concrete. Foundation students need the physical/visual anchor before the abstract number name. Keep the pace slow and repeat the pattern "ten and ___ more" for each number.

MISCONCEPTIONS:
- Misconception: Students think 11 and 12 are completely separate from 10
  Why: The words "eleven" and "twelve" do not sound like "ten-one" or "ten-two" the way thirteen sounds like "three-teen"
  Impact: Students cannot see the ten-and-ones structure in these numbers, which blocks place value understanding
  Quick correction: "Eleven means ten and one more. I can show you -- look, a full ten frame and one extra."

WATCH FOR:
- Students who look at the ten frame and can see it is "full" plus extras -- the visual anchor is working
- Students who seem confused by the connection between the full frame and the number name

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Let's keep going. Watch me build 13, 14, and 15
- 10 and 3 more makes... 13. Thirteen. Can you hear "three" in "thirteen"?
- 10 and 4 more makes... 14. Fourteen. Can you hear "four" in "fourteen"?
- 10 and 5 more makes... 15. Fifteen. Can you hear "five" in "fifteen"?
- Every teen number starts with a full ten and then has extra ones

DO:
- Build each number on the ten frame visual
- Point to the "extras" for each number
- Have students repeat each number name after you
- Emphasise the pattern: teen numbers have "ten and some more"

TEACHER NOTES:
This I Do extends to 13-15, where the naming pattern becomes more transparent (thir-teen, four-teen, fif-teen). Drawing attention to the sound pattern helps students connect the name to the structure.

WATCH FOR:
- Students who start saying the numbers before you -- they are picking up the pattern
- Students who can hear "three" in "thirteen" -- phonological awareness is supporting number sense

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. I am going to show you a ten frame
- Count the counters and tell me the number
- Show me on your fingers how many extras there are after the ten

DO:
- Display a ten frame showing 13 (10 + 3)
- Students show 3 fingers for the extras
- Then cold call: "What number is this?" [13]

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "Look at the ten frame. The ten frame is full -- that is 10. Now count the extras. Show me on your fingers how many extras. Ready... show me."
- Scan for: 3 fingers. Students who show 3 quickly have grasped the ten-and-ones structure.
PROCEED: If 80%+ show 3 fingers, move to the We Do.
PIVOT: Most likely misconception -- students count ALL the counters from 1 instead of seeing the full ten plus extras. Reteach: "The ten frame is full. That means we already have 10. We do not need to count those again. Just count the extra ones: 1, 2, 3. So it is 10 and 3 more -- 13."

TEACHER NOTES:
This CFU checks whether students can decompose a teen number into ten + ones using the visual. Finger voting is quick and universal for Foundation.

WATCH FOR:
- Students who count from 1 instead of recognising the full ten frame -- they need more support with subitising 10
- Readiness signal: students showing 3 fingers instantly without counting the ten frame cells individually

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn! I will show you a ten frame picture
- With your partner, work out which teen number it shows
- Hold up your whiteboard with the number when you are ready

DO:
- Display ten frames for 11, 14, 12, 15 one at a time
- Partners discuss and write the number on whiteboards
- Allow 20 seconds per number
- Cold call for each answer

TEACHER NOTES:
We Do uses the same ten frame format as I Do but asks students to identify the number rather than watch the teacher build it. Working with a partner supports students who are still building confidence after the break.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 1 Enabling Scaffold which has pre-drawn ten frames with the ten already shaded. Students only need to count the extra ones and write the number.
- Extra Notes: Distribute the Session 1 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: After identifying each number, draw your own ten frame to show a teen number your partner calls out. Then try numbers beyond 15 (16, 17).
- Extra Notes: Extending students work with the Session 1 Extension PDF.

WATCH FOR:
- Students who write the correct number quickly -- they can read the ten frame structure
- Students who write single-digit numbers (e.g., writing 4 instead of 14) -- they are counting extras but not adding the ten

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check. First ten frame: a full ten and 1 more. What number? [11] Eleven!
- Second: a full ten and 4 more. [14] Fourteen!
- Third: a full ten and 2 more. [12] Twelve!
- Fourth: a full ten and 5 more. [15] Fifteen!

DO:
- Reveal answers one at a time
- Cold call different students for each
- For each, repeat "ten and ___ more makes ___"

TEACHER NOTES:
Reinforcing the "ten and ___ more" language pattern with each reveal builds the connection between the visual and the number name.

WATCH FOR:
- Students who self-correct when they hear the answer -- learning is happening in the moment
- Students who got all four correct -- they are ready for independent practice

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. Thumbs up if I say the right number, thumbs down if I say the wrong number
- This ten frame shows a full ten and 3 more. Is this 13?
- Show me... now.

DO:
- Students show thumbs up (correct -- it is 13)
- Follow up: "How do you know?" Cold call 2 students
- Then try a wrong one: "This shows a full ten and 2 more. Is this 15?" [Thumbs down -- it is 12]

CFU CHECKPOINT:
Technique: Thumbs Up/Down
Script:
- Say: "Thumbs up if I am right, thumbs down if I am wrong. This ten frame has a full ten and 3 extras. Is this 13? Think... show me."
- Scan for: thumbs up. Then: "This one has a full ten and 2 extras. Is this 15?" Scan for: thumbs down.
PROCEED: If 80%+ get both correct, move to You Do.
PIVOT: Most likely misconception -- students confuse the extras count with the total. Reteach using physical counters: "Let's count together. Here is my full ten. Now my extras: 1, 2, 3. Ten and 3 more is 13, not 3."

TEACHER NOTES:
Two-part hinge checks both positive and negative identification. The wrong example is critical -- it confirms students are not just agreeing with everything.

WATCH FOR:
- Students who show thumbs up for both -- they may be agreeing rather than checking
- Readiness signal: confident thumbs down on the wrong example with quick correction to "12"

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice time
- On your worksheet you will see ten frames
- First: Count the full ten. Next: Count the extras. Then: Write the teen number.
- You have 6 minutes. Go.

DO:
- Distribute Session 1 Worksheet
- Students work independently
- Circulate: start with students who struggled during We Do
- For enabling students, distribute the Session 1 Enabling Scaffold
- For extending students, distribute the Session 1 Extension

TEACHER NOTES:
You Do uses different ten frame arrangements from the We Do (different numbers, some with visual distractors). The worksheet progresses: identify numbers, then draw counters to match a given number.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Complete the enabling scaffold where the ten is pre-shaded. Students count only the extras and trace the number. Simpler numbers (11, 12, 13) only.
- Extra Notes: Distribute the Session 1 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Use the Session 1 Extension to explore teen numbers beyond 15. Draw ten frames for 16, 17, 18 and discover the pattern continues.
- Extra Notes: Self-contained investigation with visual examples.

WATCH FOR:
- Students who write the correct numbers quickly -- ten frame reading is secured
- Students who draw the correct number of counters when asked to "show 14" -- they can represent as well as identify
- Readiness signal: completing 6+ problems correctly within 4 minutes

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket time. Three quick questions.
- Work on your own. You have 2 minutes.

DO:
- Display the exit ticket
- Students respond on whiteboards or in workbooks
- Collect or scan responses to sort into secure, developing, beginning

TEACHER NOTES:
Exit ticket assesses SC1 (Q1 -- say the number) and SC2 (Q2 -- identify from ten frame). Q3 touches SC3 (explain the structure). Sort results to plan Session 2 support.

WATCH FOR:
- Students who get Q1 and Q2 correct but struggle with Q3 -- they can identify but not yet explain the structure
- Students who complete all three -- they are ready for numbers 16-20 in Session 2

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Let's check our success criteria. Look at the slide.
- SC1: I can say the numbers 11, 12, 13, 14, and 15. Thumbs up, sideways, or down.
- SC2: I can match a teen number to its ten frame picture. Thumbs?
- SC3: I can explain that teen numbers have a ten and extra ones. Thumbs?
- Turn and talk: Tell your partner what "teen" means in teen numbers.

DO:
- Display success criteria on screen
- Run thumbs check for each SC
- Allow 30 seconds for Turn and Talk
- Cold call 2-3 students to share

TEACHER NOTES:
Closing brings the lesson full circle. The Turn and Talk targets the core concept: teen = ten + more. Students who can say "it means ten and some more" have grasped the key idea.

WATCH FOR:
- Students who say "teen means ten and more" -- the concept is forming
- Students who still seem unsure -- they will get more practice in Sessions 2-5

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- These are the resources for today's lesson
- The worksheet and answer key are linked here

DO:
- Point out each resource and its purpose

TEACHER NOTES:
Resource slide provides clickable links to all companion PDFs for this session.

WATCH FOR:
- N/A -- resource slide is teacher reference only

[General: Resources | VTLM 2.0: Planning]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, UNIT_TITLE, "Session 1: Introducing Teen Numbers 11-15",
    "Foundation Numeracy | Session 1 of 5 | Term 2 Week 1", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2-3: Daily Review (withReveal) — Counting to 20 by 1s
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Counting to 20", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
      const challenges = [
        "1.  What number comes after 7?",
        "2.  What number comes after 11?",
        "3.  What number comes after 15?",
        "4.  What number comes after 19?",
        "5.  Count from 8 to 14. What numbers did you say?",
      ];
      s.addText(challenges.map((p, i) => ({
        text: p,
        options: { fontSize: 14, color: C.CHARCOAL, breakLine: i < challenges.length - 1, paraSpaceAfter: 10 },
      })), {
        x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
      s.addText([
        { text: "On your whiteboards:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "Write the answer to each question.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "", options: { fontSize: 8, breakLine: true } },
        { text: "Remember:", options: { fontSize: 14, bold: true, color: C.ALERT, breakLine: true } },
        { text: "Count carefully!", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Say the numbers in your head first", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: 5.45, y: CONTENT_TOP + 0.15, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.3,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Answers:  1) 8    2) 12    3) 16    4) 20    5) 8, 9, 10, 11, 12, 13, 14", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 12, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Counting from 1-20
  const sFluency = pres.addSlide();
  addTopBar(sFluency, STAGE_COLORS["1"]);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Count from 1 to 20", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

  addCard(sFluency, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  sFluency.addText([
    { text: "Round 1: Clap and Count", options: { fontSize: 15, bold: true, color: C.PRIMARY, breakLine: true } },
    { text: "Count 1 to 20 together.", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
    { text: "Clap on every number!", options: { fontSize: 13, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "Round 2: Whisper and Shout", options: { fontSize: 15, bold: true, color: C.PRIMARY, breakLine: true } },
    { text: "WHISPER 1 to 10", options: { fontSize: 13, color: C.MUTED, breakLine: true } },
    { text: "SHOUT 11 to 20!", options: { fontSize: 14, bold: true, color: C.ALERT, breakLine: true } },
    { text: "", options: { fontSize: 10, breakLine: true } },
    { text: "Round 3: Speed Count", options: { fontSize: 15, bold: true, color: C.PRIMARY, breakLine: true } },
    { text: "How fast can we count 1-20?", options: { fontSize: 13, color: C.CHARCOAL } },
  ], {
    x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addCard(sFluency, 5.2, CONTENT_TOP, 4.3, 2.8, { strip: C.SECONDARY });
  // Number grid 1-20 in the right card
  const gridStartX = 5.45;
  const gridStartY = CONTENT_TOP + 0.15;
  const gridCellW = 0.72;
  const gridCellH = 0.5;
  for (let i = 0; i < 20; i++) {
    const row = Math.floor(i / 5);
    const col = i % 5;
    const num = i + 1;
    const isTeenRange = num >= 11;
    sFluency.addShape("roundRect", {
      x: gridStartX + col * gridCellW,
      y: gridStartY + row * gridCellH,
      w: gridCellW - 0.04,
      h: gridCellH - 0.04,
      rectRadius: 0.06,
      fill: { color: isTeenRange ? C.PRIMARY : C.BG_CARD },
      line: { color: C.PRIMARY, width: 1 },
    });
    sFluency.addText(String(num), {
      x: gridStartX + col * gridCellW,
      y: gridStartY + row * gridCellH,
      w: gridCellW - 0.04,
      h: gridCellH - 0.04,
      fontSize: 18, fontFace: FONT_H, color: isTeenRange ? C.WHITE : C.CHARCOAL,
      align: "center", valign: "middle", bold: true, margin: 0,
    });
  }

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    ["We are learning to name and recognise the teen numbers 11 to 15"],
    [
      "I can say the numbers 11, 12, 13, 14, and 15",
      "I can match a teen number to its ten frame picture",
      "I can explain that teen numbers have a ten and extra ones",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — What are teen numbers? (11, 12)
  workedExSlide(pres, 2, "I Do", "What Are Teen Numbers?",
    [
      "Teen numbers have a TEN and some more",
      "",
      "11 = ten and 1 more (eleven)",
      "12 = ten and 2 more (twelve)",
      "",
      "Look at the ten frames:",
      "A full ten frame = 10",
      "The extra counters = the ones",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      // Ten frame for 11
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.PRIMARY });
      slide.addText("11 = ten and 1 more", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      drawTenFrameWithExtras(slide, lg.rightX + 0.3, lg.panelTopPadded + 0.45, 1.6, 0.56, 11);

      // Ten frame for 12
      slide.addText("12 = ten and 2 more", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 1.55, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      drawTenFrameWithExtras(slide, lg.rightX + 0.3, lg.panelTopPadded + 1.92, 1.6, 0.56, 12);
    }
  );

  // Slide 7: I Do — Teen numbers 13, 14, 15
  workedExSlide(pres, 2, "I Do", "Teen Numbers: 13, 14, 15",
    [
      "13 = ten and 3 more (thirteen)",
      "   Can you hear THREE in thirteen?",
      "",
      "14 = ten and 4 more (fourteen)",
      "   Can you hear FOUR in fourteen?",
      "",
      "15 = ten and 5 more (fifteen)",
      "   Can you hear FIVE in fifteen?",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.5, { strip: C.SECONDARY });
      // 13
      slide.addText("13", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 0.08, w: 0.5, h: 0.35,
        fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      drawTenFrameWithExtras(slide, lg.rightX + 0.65, lg.panelTopPadded + 0.08, 1.3, 0.44, 13, { dotColor: C.ACCENT });

      // 14
      slide.addText("14", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 1.05, w: 0.5, h: 0.35,
        fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      drawTenFrameWithExtras(slide, lg.rightX + 0.65, lg.panelTopPadded + 1.05, 1.3, 0.44, 14, { dotColor: C.ACCENT });

      // 15
      slide.addText("15", {
        x: lg.rightX + 0.1, y: lg.panelTopPadded + 2.02, w: 0.5, h: 0.35,
        fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      drawTenFrameWithExtras(slide, lg.rightX + 0.65, lg.panelTopPadded + 2.02, 1.3, 0.44, 15, { dotColor: C.ACCENT });
    }
  );

  // Slide 8-9: CFU 1 (withReveal) — Finger Voting
  withReveal(
    () => cfuSlide(pres, "CFU", "How Many Extras?", "Finger Voting",
      "Look at this ten frame.\n\nThe ten frame is FULL (that is 10).\nHow many EXTRA counters are there?\n\nShow me on your fingers!",
      NOTES_CFU1, FOOTER),
    (slide) => {
      addTextOnShape(slide, "3 extras!  10 + 3 = 13  (thirteen)", {
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

      // Left: instruction card
      addCard(s, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["3"] });
      s.addText([
        { text: "With your partner:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
        { text: "", options: { fontSize: 6, breakLine: true } },
        { text: "Look at each ten frame.", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "Count the full ten.", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "Count the extras.", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
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

      // Right: 4 ten frames
      addCard(s, 5.2, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: C.ACCENT });

      const frames = [
        { label: "A", count: 11 },
        { label: "B", count: 14 },
        { label: "C", count: 12 },
        { label: "D", count: 15 },
      ];
      frames.forEach((f, i) => {
        const frameY = CONTENT_TOP + 0.1 + i * 0.9;
        s.addText(f.label + ".", {
          x: 5.35, y: frameY, w: 0.35, h: 0.3,
          fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
        });
        drawTenFrameWithExtras(s, 5.75, frameY, 1.25, 0.38, f.count, { dotColor: C.ACCENT });
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      runSlideDiagnostics(s, pres, { respectSafeBottom: false });
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "A) 11    B) 14    C) 12    D) 15", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 12-13: CFU 2 Hinge (withReveal)
  withReveal(
    () => cfuSlide(pres, "CFU", "Hinge Check", "Thumbs Up/Down",
      "This ten frame shows a full ten and 2 extras.\n\nIs this number 15?\n\nThumbs up = YES    Thumbs down = NO",
      NOTES_CFU2, FOOTER),
    (slide) => {
      addTextOnShape(slide, "Thumbs DOWN!  Ten and 2 more = 12, not 15", {
        x: 1.0, y: 4.0, w: 8, h: 0.55, rectRadius: 0.08,
        fill: { color: C.ALERT },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 14: You Do
  workedExSlide(pres, 4, "You Do", "Your Turn: Teen Numbers",
    [
      "First: Look at each ten frame on your worksheet.",
      "Next: Count the full ten, then count the extras.",
      "Then: Write the teen number next to each picture.",
      "",
      "If you finish early, try drawing your own ten",
      "frames for the numbers your teacher calls out.",
      "",
      "You have 6 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.8, { strip: C.ALERT });
      slide.addText("Remember:", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Full ten frame = 10", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Count the extras", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "10 + extras = teen number", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.5, h: 1.0,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      // Example ten frame
      slide.addText("Example: 14", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 2.1, w: lg.rightW - 0.3, h: 0.25,
        fontSize: 12, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      drawTenFrameWithExtras(slide, lg.rightX + 0.6, lg.panelTopPadded + 2.4, 1.2, 0.4, 14);
    }
  );

  // Slide 15: Exit Ticket
  exitTicketSlide(pres,
    [
      "Say these numbers to your teacher: 11, 13, 15",
      "Look at this ten frame (full ten + 4 extras). What number is it?",
      "Why is 13 called a teen number? (Hint: what does it have?)",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 16: Closing
  closingSlide(pres,
    "Tell your partner: What does TEEN mean in teen numbers?",
    [
      "I can say the numbers 11, 12, 13, 14, and 15",
      "I can match a teen number to its ten frame picture",
      "I can explain that teen numbers have a ten and extra ones",
    ],
    NOTES_CLOSING);

  // Slide 17: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "Teen_Session1_Introducing_11_to_15.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Generate PDFs ──────────────────────────────────────────────────────────

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Teen Numbers 11-15: Ten Frame Practice",
      color: C.PRIMARY,
      lessonInfo: "Session 1 of 5 | Foundation Numeracy",
    });
    y = addTipBox(doc, "Look at each ten frame. Count the full ten, then count the extras. Write the teen number.", y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Section A: What number is this?", y, { color: C.PRIMARY });
    y = addProblem(doc, 1, "A full ten frame and 1 extra counter. What number? ___", y, { color: C.PRIMARY });
    y = addProblem(doc, 2, "A full ten frame and 3 extra counters. What number? ___", y, { color: C.PRIMARY });
    y = addProblem(doc, 3, "A full ten frame and 5 extra counters. What number? ___", y, { color: C.PRIMARY });
    y = addProblem(doc, 4, "A full ten frame and 2 extra counters. What number? ___", y, { color: C.PRIMARY });
    y = addProblem(doc, 5, "A full ten frame and 4 extra counters. What number? ___", y, { color: C.PRIMARY });

    y = addSectionHeading(doc, "Section B: Draw the counters", y, { color: C.PRIMARY });
    y = addProblem(doc, 6, "Draw 13 counters using a ten frame and extras.", y, {
      color: C.PRIMARY,
      writeLines: [{ label: "How many in the ten frame?" }, { label: "How many extras?" }],
    });
    y = addProblem(doc, 7, "Draw 11 counters using a ten frame and extras.", y, {
      color: C.PRIMARY,
      writeLines: [{ label: "How many in the ten frame?" }, { label: "How many extras?" }],
    });
    y = addProblem(doc, 8, "Draw 15 counters using a ten frame and extras.", y, {
      color: C.PRIMARY,
      writeLines: [{ label: "How many in the ten frame?" }, { label: "How many extras?" }],
    });

    addPdfFooter(doc, "Session 1 | Teen Numbers | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Teen Numbers 11-15",
      color: C.PRIMARY,
      lessonInfo: "Session 1 of 5 | Foundation Numeracy",
    });
    y = addSectionHeading(doc, "Section A", y, { color: C.PRIMARY });
    y = addBodyText(doc, "1. 10 + 1 = 11 (eleven)", y);
    y = addBodyText(doc, "2. 10 + 3 = 13 (thirteen)", y);
    y = addBodyText(doc, "3. 10 + 5 = 15 (fifteen)", y);
    y = addBodyText(doc, "4. 10 + 2 = 12 (twelve)", y);
    y = addBodyText(doc, "5. 10 + 4 = 14 (fourteen)", y);
    y = addSectionHeading(doc, "Section B", y, { color: C.PRIMARY });
    y = addBodyText(doc, "6. 13: Ten frame full (10), 3 extras", y);
    y = addBodyText(doc, "7. 11: Ten frame full (10), 1 extra", y);
    y = addBodyText(doc, "8. 15: Ten frame full (10), 5 extras", y);
    addPdfFooter(doc, "Session 1 | Answer Key | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Ten Frames With Visual Support",
      color: C.ACCENT,
      lessonInfo: "Session 1 of 5 | Foundation Numeracy",
    });
    y = addTipBox(doc, "The ten frame is already shaded for you. Just count the extra dots and write the number!", y, { color: C.ACCENT });

    y = addSectionHeading(doc, "Count the extras and write the number", y, { color: C.PRIMARY });
    const scaffoldNums = [11, 12, 13];
    scaffoldNums.forEach((num, i) => {
      const extras = num - 10;
      y = addProblem(doc, i + 1,
        "Ten frame: 10 shaded.  Extra dots: " + extras + ".  This number is: ___",
        y, { color: C.PRIMARY });
    });

    y = addSectionHeading(doc, "Now try these (count carefully)", y, { color: C.PRIMARY });
    [14, 15].forEach((num, i) => {
      const extras = num - 10;
      y = addProblem(doc, i + 4,
        "Ten frame: 10 shaded.  Extra dots: " + extras + ".  This number is: ___",
        y, { color: C.PRIMARY });
    });

    addPdfFooter(doc, "Session 1 | Enabling Scaffold | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Extension — Teen Numbers Beyond 15
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "What Comes After 15?",
      color: C.PRIMARY,
      lessonInfo: "Session 1 of 5 | Foundation Numeracy",
    });

    y = addSectionHeading(doc, "The Pattern Continues!", y, { color: C.PRIMARY });
    y = addBodyText(doc, "You know that teen numbers have a TEN and some more.", y);
    y = addBodyText(doc, "11 = ten and 1    12 = ten and 2    13 = ten and 3", y);
    y = addBodyText(doc, "14 = ten and 4    15 = ten and 5", y);
    y = addBodyText(doc, "What comes next? The pattern keeps going!", y);

    y = addSectionHeading(doc, "Meet the Next Teen Numbers", y, { color: C.PRIMARY });
    y = addBodyText(doc, "16 = ten and 6 more (sixteen) -- can you hear SIX in sixteen?", y);
    y = addBodyText(doc, "17 = ten and 7 more (seventeen) -- can you hear SEVEN in seventeen?", y);
    y = addBodyText(doc, "18 = ten and 8 more (eighteen) -- can you hear EIGHT in eighteen?", y);

    y = addSectionHeading(doc, "Your Investigation", y, { color: C.PRIMARY });
    y = addBodyText(doc, "Draw ten frames with extras for these numbers:", y);
    y = addProblem(doc, 1, "Draw 16 using a ten frame and extras.", y, {
      color: C.PRIMARY, writeLines: [{ label: "Extras:" }, { label: "Number:" }],
    });
    y = addProblem(doc, 2, "Draw 17 using a ten frame and extras.", y, {
      color: C.PRIMARY, writeLines: [{ label: "Extras:" }, { label: "Number:" }],
    });
    y = addProblem(doc, 3, "Draw 18 using a ten frame and extras.", y, {
      color: C.PRIMARY, writeLines: [{ label: "Extras:" }, { label: "Number:" }],
    });

    y = addTipBox(doc, "Challenge: Can you figure out 19 and 20? What happens when the extras row is full too?", y, { color: C.ACCENT });

    addPdfFooter(doc, "Session 1 | Extension Investigation | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 1 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
