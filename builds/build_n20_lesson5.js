"use strict";

// Numbers to 20 Unit — Session 5: Putting It All Together
// Term 2 Week 1, Foundation Numeracy, Variant 0
// DR: Teen Numbers — I can order the numbers from 0 to at least 20
// Fluency: Sequencing numbers 1-10
// VC2MFN01 — name, represent and order numbers including zero to at least 20

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  formatSessionResourceFileName,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "foundation", 0);
const {
  C, FONT_H, FONT_B,
  SLIDE_W, SLIDE_H, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide, addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addBadge, addTitle,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 5;
const UNIT_TITLE = "Numbers to 20";
const FOOTER = "Numbers to 20 | Session 5 of 5 | Foundation Numeracy";
const OUT_DIR = "output/N20_Session5_Putting_It_Together";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Independent practice - count, represent, and order numbers to 20.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with answers.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Structured support mat with ten frames and number line for numbers to 10.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Self-contained investigation into numbers beyond 20.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helper: draw ten frame on slide ────────────────────────────────────────
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
}

// ─── Teacher Notes ───────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Last session for this week! You have done an amazing job with numbers to 20
- Today we are putting everything together -- counting, showing, and ordering
- Think of this as your chance to show off what you have learned

DO:
- Display title slide
- Have all materials ready: ten frames, counters, number cards, whiteboards

TEACHER NOTES:
Session 5 of 5. Consolidation lesson that draws together counting (Session 1), teen number composition (Session 2), multiple representations (Session 3), and ordering (Session 4). The exit ticket doubles as an end-of-week assessment for the unit.

WATCH FOR:
- Students who seem confident and ready -- challenge them with the extension
- Students who are still building skills -- the enabling scaffold provides structured support

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Warm up! Put these numbers in order from 0 to 20
- Some numbers are already in place. Fill in the missing ones
- Write the missing numbers on your whiteboard

DO:
- Display a number strip 0-20 with 6 numbers missing
- Allow 90 seconds for students to identify and write missing numbers
- Hold whiteboards face down until signal

TEACHER NOTES:
Daily Review targets teen number ordering (the focus from last week and reinforced in Session 4). The missing number format requires students to hold the sequence in memory and identify gaps -- a consolidation of counting sequence knowledge.

WATCH FOR:
- Students who fill in all gaps quickly -- counting sequence to 20 is automated
- Students who struggle at specific points (e.g., the 12-13 or 17-18 transitions) -- note for future fluency work

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers -- tick if you got it right, fix it if you didn't
- Let's count from 0 to 20 together to check. Ready? 0, 1, 2...
- Ask: Which missing number was the trickiest? Why?

DO:
- Click to reveal answers
- Students tick correct and fix errors
- Lead a choral count 0-20

TEACHER NOTES:
Tick-and-fix plus a choral count consolidates the counting sequence. Students who can count 0-20 fluently have the prerequisite for all number work in the coming weeks.

WATCH FOR:
- Students who self-corrected and can explain the pattern -- metacognition is developing
- Students who still have gaps in the teen sequence -- flag for next week's fluency focus

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time! Number card race
- I will show you number cards 1 to 10 in a jumbled order
- Put them in order on your desk as fast as you can
- Ready... go!

DO:
- Each pair gets number cards 1-10 face down
- On "go," pairs flip and arrange in order left to right
- Time the class -- aim for under 30 seconds
- Repeat to beat their time
- After racing, cold call: "What comes after 6? Before 4? Between 7 and 9?"

TEACHER NOTES:
Sequencing 1-10 builds speed with the foundational counting sequence. The physical sorting is engaging and kinaesthetic. Racing adds urgency and energy to the end-of-week lesson.

WATCH FOR:
- Pairs who finish quickly with correct order -- strong sequencing skills
- Students who place cards out of order and need to swap -- they know the sequence but are still building automaticity

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read from slide: "We are learning to count, show, and order numbers to 20"
- Our success criteria:
- Read from slide: "I can count a collection of objects up to 20"
- Read from slide: "I can show a number using a ten frame"
- Read from slide: "I can put numbers in order and find missing numbers"

DO:
- Choral read LI and each SC
- Say: "These are all the things we have learned this week!"

TEACHER NOTES:
This lesson consolidates all three prior sessions. SC1 targets Sessions 1-2 (counting and teen composition). SC2 targets Sessions 2-3 (ten frame representation). SC3 targets Session 4 (ordering). The exit ticket assesses all three SC.

WATCH FOR:
- Students who show confidence in reading the SC -- they recognise these skills from earlier in the week

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me solve a number challenge. The challenge says: "Count the collection, show it on a ten frame, then put these numbers in order."
- First: I count the objects. Touch... 1, 2, 3... 14. There are 14.
- Next: I show 14 on the ten frame. Fill it up -- 10! Plus 4 extras. 10 and 4 makes 14.
- Then: I need to order 14 with these other numbers: 8 and 19. Which is smallest? 8. Then 14. Then 19.
- Order: 8, 14, 19

DO:
- Use physical counters and a ten frame under the document camera
- Count a collection of 14 objects
- Build 14 on the ten frame
- Write 8, 14, 19 on cards and arrange in order
- Narrate each step clearly

TEACHER NOTES:
This I Do models a multi-step task that combines counting, representing, and ordering -- the three skills from this week. The think-aloud makes the decision process visible at each step. It is deliberately a longer worked example because students are applying multiple skills.

WATCH FOR:
- Students who anticipate the next step -- they remember the skills from earlier sessions
- Students who seem overwhelmed by multiple steps -- the "First, Next, Then" structure helps manage cognitive load
- Readiness signal: students able to identify "First count, then show, then order" as the steps

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. Look at this ten frame
- The frame is full and there are 7 extras
- What is the number? Write it. And is it bigger or smaller than 15?
- Write: the number AND bigger or smaller on your whiteboard

DO:
- Display full ten frame + 7 extras
- Students write on whiteboards (17, bigger)
- Scan for correct number AND comparison

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Full ten frame plus 7 extras. What number? Is it bigger or smaller than 15? Write BOTH on your whiteboard. Ready... show me!"
- Scan for: 17, bigger. Students who get both correct can read ten frames AND compare.
PROCEED: If 80%+ show 17 and bigger, move to We Do.
PIVOT: Most likely misconception -- students write 7 (only counting extras). Reteach: "The frame is full. Full frame = 10. Plus 7 extras. 10 and 7 is 17. Is 17 bigger or smaller than 15? Count: 15, 16, 17. 17 comes AFTER 15, so it is bigger."

TEACHER NOTES:
This CFU combines two skills in one check: reading a ten frame (Session 2-3) and comparing (Session 4). It tests whether students can integrate skills from across the week.

WATCH FOR:
- Students who get 17 but struggle with the comparison -- the representation is secure, ordering needs more practice
- Students who get both correct quickly -- ready for the full multi-step task

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn! With your partner, complete the number challenge
- Step 1: Count the objects in your bag (new bags today -- bigger!)
- Step 2: Show the number on your ten frame
- Step 3: I will give you two more numbers. Put all three in order
- Ready? Count your collection first

DO:
- Distribute bags with 11-18 objects (varied across tables)
- Partners count and build on ten frames
- After building, write on board: "Now order your number with 6 and 20"
- Partners write the order on their whiteboards
- Cold call pairs to share

TEACHER NOTES:
We Do combines all three skills in a supported context. The bags contain teen-number quantities (11-18) so students practise teen composition. Ordering with 6 and 20 gives a known small number and the endpoint of the counting range.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 5 Enabling Scaffold with bags containing 5-10 objects only. Students count, build on a pre-printed ten frame (not full), and order with two given single-digit numbers.
- Extra Notes: Distribute the Session 5 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: After completing the main task, investigate collections beyond 20. The extension PDF provides a second ten frame for counting past 20 (two full frames = 20, plus extras). Students count, represent, and order numbers up to 30.
- Extra Notes: Distribute the Session 5 Extension PDF. Self-contained.

WATCH FOR:
- Students who count, build, and order fluently -- all three skills are consolidated
- Students who get stuck on one step -- identify which skill needs more support for next week
- Readiness signal: pairs completing all three steps within 3 minutes with correct results

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check. Table 2, what was in your bag? [e.g. 13] Show your ten frame. [Full + 3]
- Now order: 6, 13, 20. Smallest first? [6, 13, 20]
- Table 4, your bag? [e.g. 16] Order with 6 and 20? [6, 16, 20]
- Ask: What did everyone have in common? [6 is always first, 20 is always last]

DO:
- Cold call 3-4 tables
- Draw out the pattern: 6 is always smallest, 20 is always biggest
- Different tables had different middle numbers

TEACHER NOTES:
The shared structure (6, ___, 20) makes comparison natural -- students can see that only the middle number changes across tables. This reinforces ordering skills.

WATCH FOR:
- Students who notice the pattern independently -- strong reasoning
- Students who got the ordering wrong -- check which specific skill (counting, representing, or ordering) was the weak link

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Last check before independent practice
- I have three numbers: 11, 5, 18
- Choral response: Put them in order, smallest to biggest. Ready? Shout it out!

DO:
- Display the three numbers
- Count down: 3, 2, 1... students call out the order
- Follow up: "How do you know 5 is smallest?" [It comes first when counting]

CFU CHECKPOINT:
Technique: Choral Response
Script:
- Say: "Order these from smallest to biggest: 11, 5, 18. On my signal, call it out. 3, 2, 1... go!"
- Listen for: "5, 11, 18" -- clear, confident choral response.
PROCEED: If the choral response is clear and correct, move to You Do.
PIVOT: Most likely misconception -- students call out numbers in the original order (11, 5, 18) without reordering. Reteach: "Which one comes first when we count? Is it 5, 11, or 18? [5] So 5 goes first. Then which? [11] And last? [18]."

TEACHER NOTES:
Choral response is a quick, whole-class check. The energy of calling out together suits an end-of-week lesson and Foundation attention spans.

WATCH FOR:
- Confident, immediate choral response -- the class is ready
- Hesitation or mixed responses -- some students may need the number line visual

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice time. This is your big chance to show everything you have learned this week!
- Open your worksheet
- First: Count each collection and write the number
- Next: Show each number on a ten frame
- Then: Put sets of numbers in order
- You have 8 minutes. Go!

DO:
- Distribute Session 5 Worksheet
- Students work independently
- Circulate: this is the end-of-week assessment -- note students' strengths and gaps
- For enabling students, distribute the Session 5 Enabling Scaffold
- For extending students, distribute the Session 5 Extension

TEACHER NOTES:
This You Do combines all skills from the week into one assessment-style worksheet. Section A is counting, Section B is ten frame representation, Section C is ordering. Use results alongside the exit ticket to plan next week's focus.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 5 Enabling Scaffold with numbers to 10 only, pre-printed ten frames (no extras needed for numbers to 10), and ordering with number line support.
- Extra Notes: Distribute the Session 5 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Investigate numbers beyond 20 using double ten frames (20 + extras). The extension PDF teaches the concept and provides structured practice.
- Extra Notes: Distribute the Session 5 Extension PDF. Self-contained.

WATCH FOR:
- Students who complete all three sections correctly -- consolidation is strong across all skills
- Students who are strong in one section but weak in another -- targeted support for next week
- Readiness signal: 80%+ completing Sections A and B correctly within 5 minutes

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket time. Four questions to show what you know
- Work on your own. You have 3 minutes

DO:
- Display exit ticket on screen
- Students write answers in workbooks or on whiteboards
- Collect or photograph responses -- this is the end-of-week assessment
- Sort into secure, developing, beginning for each SC

TEACHER NOTES:
End-of-week assessment. Q1 assesses counting (SC1), Q2 assesses ten frame reading (SC2), Q3 assesses ordering (SC3), Q4 assesses missing numbers (SC3 depth). Use results to plan next week -- students who are secure can extend, students who are developing get reinforcement, students who are beginning need targeted intervention.

WATCH FOR:
- Students who complete all four correctly -- ready to extend into numbers beyond 20 next week
- Students who get Q1-Q2 but struggle with Q3-Q4 -- counting and representation are secure, ordering needs more work
- Students who struggle across all questions -- they need continued foundational number work

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- What an amazing week! Let's check our success criteria one last time
- SC1: I can count a collection of objects up to 20. Thumbs?
- SC2: I can show a number using a ten frame. Thumbs?
- SC3: I can put numbers in order and find missing numbers. Thumbs?
- Turn and talk: What was the MOST important thing you learned this week about numbers? Tell your partner

DO:
- Display success criteria
- Run thumbs check for each SC
- Allow 30 seconds for Turn and Talk
- Cold call 3-4 students to share
- Celebrate the week's learning: "You went from counting to ordering in five sessions!"

TEACHER NOTES:
End-of-week closing should be celebratory. Students have built counting, composition, representation, and ordering skills over five sessions. The Turn and Talk gives students ownership of their learning -- they identify what mattered most to them.

WATCH FOR:
- Students who name specific strategies ("touch and count," "fill the ten frame," "think about counting order") -- they have internalised the methods
- Students who name a specific moment of learning -- metacognition about their own growth

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- These are the resources for today's lesson

DO:
- Point out each resource

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
  titleSlide(pres, UNIT_TITLE, "Session 5: Putting It All Together", "Foundation Numeracy | Session 5 of 5 | Term 2 Week 1", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2-3: Daily Review (withReveal) — Order 0-20, fill gaps
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Fill the Number Strip: 0 to 20", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      // Number strip with missing numbers
      const nums = [];
      for (let i = 0; i <= 20; i++) nums.push(i);
      const missing = new Set([3, 8, 12, 15, 17, 20]);

      // Row 1: 0-10
      const cellW = 0.75;
      const cellH = 0.48;
      const startX = 0.6;
      for (let i = 0; i <= 10; i++) {
        const cx = startX + i * (cellW + 0.06);
        s.addShape("roundRect", {
          x: cx, y: CONTENT_TOP + 0.15, w: cellW, h: cellH, rectRadius: 0.06,
          fill: { color: missing.has(i) ? C.WHITE : C.PRIMARY },
          line: { color: C.PRIMARY, width: 1.2 },
        });
        s.addText(missing.has(i) ? "?" : String(i), {
          x: cx, y: CONTENT_TOP + 0.15, w: cellW, h: cellH,
          fontSize: 18, fontFace: FONT_H, color: missing.has(i) ? C.ALERT : C.WHITE,
          align: "center", valign: "middle", bold: true, margin: 0,
        });
      }
      // Row 2: 11-20
      for (let i = 11; i <= 20; i++) {
        const cx = startX + (i - 11) * (cellW + 0.06);
        s.addShape("roundRect", {
          x: cx, y: CONTENT_TOP + 0.15 + cellH + 0.1, w: cellW, h: cellH, rectRadius: 0.06,
          fill: { color: missing.has(i) ? C.WHITE : C.SECONDARY },
          line: { color: C.SECONDARY, width: 1.2 },
        });
        s.addText(missing.has(i) ? "?" : String(i), {
          x: cx, y: CONTENT_TOP + 0.15 + cellH + 0.1, w: cellW, h: cellH,
          fontSize: 18, fontFace: FONT_H, color: missing.has(i) ? C.ALERT : C.WHITE,
          align: "center", valign: "middle", bold: true, margin: 0,
        });
      }

      addCard(s, 0.5, CONTENT_TOP + 1.55, 9, 0.9, { strip: C.ACCENT });
      s.addText([
        { text: "On your whiteboards:", options: { fontSize: 16, bold: true, color: C.ACCENT, breakLine: true } },
        { text: "Write the 6 missing numbers.", options: { fontSize: 14, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 1.67, w: 8.5, h: 0.65,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Missing: 3, 8, 12, 15, 17, 20", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Sequencing 1-10 (card race)
  const sFluency = pres.addSlide();
  addTopBar(sFluency, STAGE_COLORS["1"]);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Number Card Race: 1 to 10", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

  addCard(sFluency, 0.5, CONTENT_TOP, 9, 2.0, { strip: STAGE_COLORS["1"] });
  // Jumbled number cards visual
  const jumbled = [7, 2, 9, 4, 1, 6, 10, 3, 8, 5];
  jumbled.forEach((n, i) => {
    const cx = 0.7 + (i % 5) * 1.75;
    const cy = CONTENT_TOP + 0.15 + Math.floor(i / 5) * 0.85;
    sFluency.addShape("roundRect", {
      x: cx, y: cy, w: 1.3, h: 0.65, rectRadius: 0.08,
      fill: { color: i % 2 === 0 ? C.PRIMARY : C.SECONDARY },
    });
    sFluency.addText(String(n), {
      x: cx, y: cy, w: 1.3, h: 0.65,
      fontSize: 26, fontFace: FONT_H, color: C.WHITE,
      align: "center", valign: "middle", bold: true, margin: 0,
    });
  });

  addCard(sFluency, 0.5, CONTENT_TOP + 2.2, 9, 1.2, { strip: C.ACCENT });
  sFluency.addText([
    { text: "Put them in order as FAST as you can!", options: { fontSize: 17, bold: true, color: C.ACCENT, breakLine: true } },
    { text: "", options: { fontSize: 4, breakLine: true } },
    { text: "1. Flip your cards over", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
    { text: "2. Arrange 1 to 10, left to right", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
    { text: "3. Hands up when you are done!", options: { bullet: true, fontSize: 14, color: C.CHARCOAL } },
  ], {
    x: 0.75, y: CONTENT_TOP + 2.32, w: 8.5, h: 1.0,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    ["We are learning to count, show, and order numbers to 20"],
    [
      "I can count a collection of objects up to 20",
      "I can show a number using a ten frame",
      "I can put numbers in order and find missing numbers",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — Multi-step number challenge
  workedExSlide(pres, 2, "I Do", "The Number Challenge",
    [
      "Step 1: COUNT the collection",
      "  Touch... 1, 2, 3... 14 objects!",
      "",
      "Step 2: SHOW on a ten frame",
      "  Fill the frame (10) + 4 extras = 14",
      "",
      "Step 3: ORDER with other numbers",
      "  Put 14, 8, and 19 in order:",
      "  8, 14, 19 (smallest to biggest)",
    ],
    NOTES_IDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.PRIMARY });
      slide.addText("Showing 14", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.05, w: lg.rightW - 0.3, h: 0.28,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      // Ten frame
      drawTenFrame(slide, lg.rightX + 0.3, lg.panelTopPadded + 0.38, 10, { cellW: 0.35, cellH: 0.35 });
      // 4 extras
      for (let i = 0; i < 4; i++) {
        slide.addShape("roundRect", {
          x: lg.rightX + 0.3 + i * 0.38, y: lg.panelTopPadded + 1.18, w: 0.3, h: 0.3, rectRadius: 0.15,
          fill: { color: C.ALERT },
        });
      }
      addTextOnShape(slide, "10 + 4 = 14", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 1.6, w: lg.rightW - 0.6, h: 0.35, rectRadius: 0.06,
        fill: { color: C.SECONDARY },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });

      // Order visual
      slide.addText("In order:", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 2.05, w: 1.0, h: 0.25,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0,
      });
      const orderNums = ["8", "14", "19"];
      orderNums.forEach((n, i) => {
        addTextOnShape(slide, n, {
          x: lg.rightX + 0.25 + i * 1.3, y: lg.panelTopPadded + 2.35, w: 0.9, h: 0.5, rectRadius: 0.08,
          fill: { color: C.SUCCESS },
        }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      });
    }
  );

  // Slide 7-8: CFU 1 (withReveal) — Ten frame + comparison
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["2"]);
      addStageBadge(s, 2, "I Do");
      addTitle(s, "Read and Compare", { y: 0.65, fontSize: 22, color: STAGE_COLORS["2"] });

      addCard(s, 0.5, CONTENT_TOP, 9, 2.0, { strip: C.PRIMARY });
      // Full ten frame + 7 extras
      drawTenFrame(s, 1.0, CONTENT_TOP + 0.2, 10, { cellW: 0.5, cellH: 0.5 });
      s.addText("+", {
        x: 3.7, y: CONTENT_TOP + 0.4, w: 0.4, h: 0.4,
        fontSize: 26, fontFace: FONT_H, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
      });
      for (let i = 0; i < 7; i++) {
        const row = Math.floor(i / 4);
        const col = i % 4;
        s.addShape("roundRect", {
          x: 4.3 + col * 0.55, y: CONTENT_TOP + 0.25 + row * 0.55, w: 0.45, h: 0.45, rectRadius: 0.22,
          fill: { color: C.ALERT },
        });
      }

      addCard(s, 0.5, CONTENT_TOP + 2.2, 9, 1.1, { strip: C.ACCENT });
      s.addText([
        { text: "What number is this?  Is it BIGGER or SMALLER than 15?", options: { fontSize: 17, bold: true, color: C.CHARCOAL, breakLine: true } },
        { text: "Write BOTH answers on your whiteboard!", options: { fontSize: 15, color: C.ACCENT } },
      ], {
        x: 0.75, y: CONTENT_TOP + 2.35, w: 8.5, h: 0.8,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "17 — BIGGER than 15!", {
        x: 2.0, y: 4.2, w: 6.0, h: 0.6, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 9-10: We Do (withReveal) — Count, build, order
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Your Number Challenge!",
      [
        "With your partner:",
        "",
        "Step 1: Count your bag of objects",
        "Step 2: Show the number on your",
        "             ten frame",
        "Step 3: Order YOUR number with",
        "             6 and 20",
        "",
        "Write the order on your whiteboard.",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.0, { strip: C.SECONDARY });
        slide.addText("The Challenge:", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
          fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        const steps = [
          { label: "1. COUNT", color: C.PRIMARY },
          { label: "2. SHOW (ten frame)", color: C.SECONDARY },
          { label: "3. ORDER with 6 and 20", color: C.ACCENT },
        ];
        steps.forEach((step, i) => {
          addTextOnShape(slide, step.label, {
            x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.48 + i * 0.5, w: lg.rightW - 0.4, h: 0.4, rectRadius: 0.06,
            fill: { color: step.color },
          }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });
        });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "Order: 6, (your number), 20", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 11-12: CFU 2 Hinge (withReveal) — Choral ordering
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Order These — Call It Out!", { y: 0.65, fontSize: 22, color: STAGE_COLORS["3"] });

      // Three big number cards
      const nums = ["11", "5", "18"];
      nums.forEach((n, i) => {
        const colors = [C.SECONDARY, C.PRIMARY, C.ACCENT];
        addCard(s, 0.8 + i * 3.0, CONTENT_TOP + 0.2, 2.5, 1.8, { strip: colors[i] });
        s.addText(n, {
          x: 0.8 + i * 3.0, y: CONTENT_TOP + 0.2, w: 2.5, h: 1.8,
          fontSize: 60, fontFace: FONT_H, color: colors[i],
          align: "center", valign: "middle", bold: true, margin: 0,
        });
      });

      addCard(s, 0.5, CONTENT_TOP + 2.3, 9, 0.9, { strip: C.ALERT });
      s.addText([
        { text: "Smallest to biggest! Shout it out on my signal!", options: { fontSize: 18, bold: true, color: C.ALERT, breakLine: true } },
        { text: "3... 2... 1... GO!", options: { fontSize: 16, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 2.42, w: 8.5, h: 0.65,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU2);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "5,  11,  18", {
        x: 2.0, y: 4.2, w: 6.0, h: 0.6, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 13: You Do
  workedExSlide(pres, 4, "You Do", "Show What You Know!",
    [
      "First: Count each collection and write",
      "         the number.",
      "Next: Show each number on a ten frame.",
      "Then: Put sets of numbers in order.",
      "",
      "This is your big chance to show",
      "everything you learned this week!",
      "You have 8 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.6, { strip: C.ALERT });
      slide.addText("This Week's Skills:", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Count (touch, say, move)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Show (ten frame + extras)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, breakLine: true } },
        { text: "Order (smallest to biggest)", options: { bullet: true, fontSize: 13, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.45, w: lg.rightW - 0.5, h: 0.9,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Slide 14: Exit Ticket
  exitTicketSlide(pres,
    [
      "Count the objects: How many? (15 objects shown)",
      "Ten frame full + 3 extras = what number?",
      "Order these: 16, 4, 11   smallest to biggest",
      "Missing number: 13, ___, 15, 16",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 15: Closing
  closingSlide(pres,
    "What was the MOST important thing you learned about numbers this week? Tell your partner!",
    [
      "I can count a collection of objects up to 20",
      "I can show a number using a ten frame",
      "I can put numbers in order and find missing numbers",
    ],
    NOTES_CLOSING);

  // Slide 16: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "N20_Session5_Putting_It_Together.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Generate PDFs ─────────────────────────────────────────────────────────

  // Worksheet — Combined skills assessment
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Numbers to 20 - Show What You Know!",
      color: C.NAVY,
      lessonInfo: "Session 5 of 5 | Foundation Numeracy",
    });
    y = addTipBox(doc, "Use everything you learned this week: touch and count, use ten frames, and put numbers in order!", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A: Count and Write", y, { color: C.NAVY });
    const drawDots = (doc, count, num, y) => {
      doc.fontSize(12).font("Sans-Bold").fillColor("#333333").text(num + ".", 55, y + 4);
      const dotR = 7;
      const gap = 20;
      const startX = 90;
      for (let i = 0; i < count; i++) {
        const row = Math.floor(i / 8);
        const col = i % 8;
        doc.circle(startX + col * gap, y + 10 + row * gap, dotR).fill("#" + C.SECONDARY);
      }
      const ansX = startX + Math.min(count, 8) * gap + 30;
      doc.fontSize(14).font("Sans").fillColor("#333333").text("How many? _____", ansX, y + 4);
      const rows = Math.ceil(count / 8);
      return y + rows * gap + 20;
    };
    y = drawDots(doc, 8, "1", y);
    y = drawDots(doc, 13, "2", y);
    y = drawDots(doc, 16, "3", y);

    y = addSectionHeading(doc, "Section B: Ten Frame", y, { color: C.NAVY });
    const drawPdfTenFrame = (doc, extras, taskNum, y) => {
      doc.fontSize(12).font("Sans-Bold").fillColor("#333333").text(taskNum + ".", 55, y + 8);
      const cellSize = 20;
      const gap = 2;
      const startX = 80;
      for (let row = 0; row < 2; row++) {
        for (let col = 0; col < 5; col++) {
          doc.rect(startX + col * (cellSize + gap), y + row * (cellSize + gap), cellSize, cellSize)
            .lineWidth(0.8).strokeColor("#" + C.NAVY).stroke();
          doc.circle(startX + col * (cellSize + gap) + cellSize / 2, y + row * (cellSize + gap) + cellSize / 2, 7)
            .fill("#" + C.SECONDARY);
        }
      }
      if (extras > 0) {
        const extraX = startX + 5 * (cellSize + gap) + 15;
        doc.fontSize(14).font("Sans-Bold").fillColor("#333333").text("+", extraX - 12, y + 10);
        for (let i = 0; i < extras; i++) {
          doc.circle(extraX + i * 20, y + cellSize / 2 + gap, 7).fill("#" + C.ALERT);
        }
      }
      const ansX = startX + 5 * (cellSize + gap) + 15 + Math.max(extras, 1) * 20 + 20;
      doc.fontSize(14).font("Sans").fillColor("#333333").text("= _____", ansX, y + 10);
      return y + cellSize * 2 + gap + 16;
    };
    y = drawPdfTenFrame(doc, 5, "4", y);
    y = drawPdfTenFrame(doc, 8, "5", y);
    y = drawPdfTenFrame(doc, 2, "6", y);

    y = addSectionHeading(doc, "Section C: Order (Smallest to Biggest)", y, { color: C.NAVY });
    y = addProblem(doc, 7, "Order:  9,  3,  15       ___,  ___,  ___", y, { color: C.NAVY });
    y = addProblem(doc, 8, "Order:  18,  7,  12      ___,  ___,  ___", y, { color: C.NAVY });
    y = addProblem(doc, 9, "Missing:  11,  ___,  13,  14", y, { color: C.NAVY });
    y = addProblem(doc, 10, "Missing:  ___,  17,  18,  19,  20", y, { color: C.NAVY });

    addPdfFooter(doc, "Session 5 | Numbers to 20 | Foundation");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Numbers to 20 - Show What You Know!",
      color: C.NAVY,
      lessonInfo: "Session 5 of 5 | Foundation Numeracy",
    });
    y = addSectionHeading(doc, "Section A", y, { color: C.NAVY });
    y = addBodyText(doc, "1. 8 dots -> 8", y);
    y = addBodyText(doc, "2. 13 dots -> 13", y);
    y = addBodyText(doc, "3. 16 dots -> 16", y);
    y = addSectionHeading(doc, "Section B", y, { color: C.NAVY });
    y = addBodyText(doc, "4. Full frame + 5 extras = 15", y);
    y = addBodyText(doc, "5. Full frame + 8 extras = 18", y);
    y = addBodyText(doc, "6. Full frame + 2 extras = 12", y);
    y = addSectionHeading(doc, "Section C", y, { color: C.NAVY });
    y = addBodyText(doc, "7. 3, 9, 15", y);
    y = addBodyText(doc, "8. 7, 12, 18", y);
    y = addBodyText(doc, "9. Missing: 12", y);
    y = addBodyText(doc, "10. Missing: 16", y);
    addPdfFooter(doc, "Session 5 | Answer Key | Foundation");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold — Numbers to 10, structured support
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Numbers to 10 - Practice Mat",
      color: C.TEAL,
      lessonInfo: "Session 5 of 5 | Foundation Numeracy",
    });
    y = addTipBox(doc, "Count carefully. Use the ten frame. Use the number line to order!", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Count and Write (to 10)", y, { color: C.NAVY });
    const drawSmallDots = (doc, count, num, y) => {
      doc.fontSize(12).font("Sans-Bold").fillColor("#333333").text(num + ".", 55, y + 4);
      const dotR = 8;
      const gap = 22;
      const startX = 90;
      for (let i = 0; i < count; i++) {
        doc.circle(startX + i * gap, y + 10, dotR).fill("#" + C.SECONDARY);
      }
      doc.fontSize(14).font("Sans").fillColor("#333333").text("How many? _____", startX + count * gap + 15, y + 2);
      return y + 34;
    };
    y = drawSmallDots(doc, 4, "1", y);
    y = drawSmallDots(doc, 7, "2", y);
    y = drawSmallDots(doc, 3, "3", y);
    y = drawSmallDots(doc, 9, "4", y);

    y = addSectionHeading(doc, "Show on a Ten Frame", y, { color: C.NAVY });
    const drawBlankFrame = (doc, num, taskNum, y) => {
      doc.fontSize(12).font("Sans-Bold").fillColor("#333333").text(taskNum + ".  Show " + num, 55, y + 8);
      const cellSize = 20;
      const gap = 2;
      const startX = 160;
      for (let row = 0; row < 2; row++) {
        for (let col = 0; col < 5; col++) {
          doc.rect(startX + col * (cellSize + gap), y + row * (cellSize + gap), cellSize, cellSize)
            .lineWidth(0.8).strokeColor("#" + C.NAVY).stroke();
        }
      }
      return y + cellSize * 2 + gap + 14;
    };
    y = drawBlankFrame(doc, 6, "5", y);
    y = drawBlankFrame(doc, 8, "6", y);
    y = drawBlankFrame(doc, 10, "7", y);

    y = addSectionHeading(doc, "Order with Number Line", y, { color: C.NAVY });
    // Number line 0-10
    const lineX = 60;
    const lineW = 450;
    doc.moveTo(lineX, y + 10).lineTo(lineX + lineW, y + 10).lineWidth(2).strokeColor("#" + C.NAVY).stroke();
    for (let i = 0; i <= 10; i++) {
      const tx = lineX + (i / 10) * lineW;
      doc.moveTo(tx, y + 4).lineTo(tx, y + 16).lineWidth(1.5).strokeColor("#" + C.NAVY).stroke();
      doc.fontSize(11).font("Sans-Bold").fillColor("#333333").text(String(i), tx - 5, y + 20);
    }
    y += 40;
    y = addProblem(doc, 8, "Order:  5,  2,  8       ___,  ___,  ___", y, { color: C.NAVY });
    y = addProblem(doc, 9, "Order:  10,  3,  7      ___,  ___,  ___", y, { color: C.NAVY });

    addPdfFooter(doc, "Session 5 | Enabling Scaffold | Foundation");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Extension — Numbers beyond 20
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Numbers Beyond 20",
      color: C.TEAL,
      lessonInfo: "Session 5 of 5 | Foundation Numeracy",
    });

    y = addSectionHeading(doc, "What happens after 20?", y, { color: C.NAVY });
    y = addBodyText(doc, "After 20, we start the twenties: 21, 22, 23, 24, 25, 26, 27, 28, 29, 30.", y);
    y = addBodyText(doc, "To show numbers bigger than 20, we use TWO ten frames!", y);

    y = addSectionHeading(doc, "Example: Showing 23", y, { color: C.NAVY });
    y = addBodyText(doc, "Ten frame 1: FULL (10)", y);
    y = addBodyText(doc, "Ten frame 2: FULL (10)", y);
    y = addBodyText(doc, "Extras: 3", y);
    y = addBodyText(doc, "10 + 10 + 3 = 23!", y);

    // Draw double ten frame example
    const cellSize = 18;
    const gap = 2;
    const startX = 80;
    for (let frame = 0; frame < 2; frame++) {
      for (let row = 0; row < 2; row++) {
        for (let col = 0; col < 5; col++) {
          doc.rect(startX + frame * 130 + col * (cellSize + gap), y + row * (cellSize + gap), cellSize, cellSize)
            .lineWidth(0.8).strokeColor("#" + C.NAVY).stroke();
          doc.circle(startX + frame * 130 + col * (cellSize + gap) + cellSize / 2, y + row * (cellSize + gap) + cellSize / 2, 6)
            .fill("#" + C.SECONDARY);
        }
      }
    }
    // 3 extras
    for (let i = 0; i < 3; i++) {
      doc.circle(startX + 2 * 130 + 10 + i * 20, y + cellSize / 2, 6).fill("#" + C.ALERT);
    }
    doc.fontSize(12).font("Sans-Bold").fillColor("#" + C.NAVY).text("= 23", startX + 2 * 130 + 10 + 3 * 20 + 10, y + 4);
    y += cellSize * 2 + gap + 20;

    y = addSectionHeading(doc, "Your Investigation", y, { color: C.NAVY });
    y = addBodyText(doc, "Use two ten frames to show each number. Draw dots and write the number.", y);

    const drawDoubleFrame = (doc, num, taskNum, y) => {
      doc.fontSize(12).font("Sans-Bold").fillColor("#333333").text(taskNum + ".  Show " + num, 55, y + 8);
      const tfX = 160;
      for (let frame = 0; frame < 2; frame++) {
        for (let row = 0; row < 2; row++) {
          for (let col = 0; col < 5; col++) {
            doc.rect(tfX + frame * 120 + col * (cellSize + gap), y + row * (cellSize + gap), cellSize, cellSize)
              .lineWidth(0.6).strokeColor("#" + C.NAVY).stroke();
          }
        }
      }
      doc.fontSize(12).font("Sans").fillColor("#333333").text("+ extras: _____  = _____", tfX + 2 * 120 + 5, y + 8);
      return y + cellSize * 2 + gap + 16;
    };

    y = drawDoubleFrame(doc, 22, "1", y);
    y = drawDoubleFrame(doc, 25, "2", y);
    y = drawDoubleFrame(doc, 28, "3", y);

    y = addSectionHeading(doc, "Order These Big Numbers!", y, { color: C.NAVY });
    y = addProblem(doc, 4, "Order:  25,  13,  30     ___,  ___,  ___", y, { color: C.NAVY });
    y = addProblem(doc, 5, "Order:  22,  18,  27,  9  ___,  ___,  ___,  ___", y, { color: C.NAVY });

    y = addSectionHeading(doc, "Did You Know?", y, { color: C.NAVY });
    y = addBodyText(doc, "Numbers keep going forever! After 30 comes 31, 32, 33... all the way to 100 and beyond. You will learn about bigger numbers as you grow!", y);

    addPdfFooter(doc, "Session 5 | Extension | Foundation");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 5 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
