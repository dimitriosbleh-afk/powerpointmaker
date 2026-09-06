"use strict";

// Numbers to 20 Unit — Session 3: Representing Numbers to 20
// Term 2 Week 1, Foundation Numeracy, Variant 0
// DR: Subitising — I can recognise how many items are in a collection, up to 5, without counting
// Fluency: Numbers after from 1-20
// VC2MFN01 — name, represent and order numbers including zero to at least 20

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
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
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 3;
const UNIT_TITLE = "Numbers to 20";
const FOOTER = "Numbers to 20 | Session 3 of 5 | Foundation Numeracy";
const OUT_DIR = "output/N20_Session3_Representing_Numbers";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Independent practice - showing numbers to 20 in different ways.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with answers.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Structured representation mat with ten frame templates and numeral guides.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Self-contained investigation into number bonds to 20.");
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
- Welcome to Session 3! You have been doing great counting and building numbers
- Today we are going to show numbers in different ways -- counters, ten frames, drawings, and fingers
- Being able to show a number lots of ways helps us really understand it

DO:
- Display title slide
- Have ten frames, counters, whiteboards, and crayons ready

TEACHER NOTES:
Session 3 of 5. This lesson extends from Session 2 (teen number composition) to multiple representations. Students move between concrete (counters), pictorial (ten frames, drawings), and symbolic (numerals) representations. This is the CRA progression in action.

WATCH FOR:
- Students who are still shaky on teen number composition from Session 2 -- they may need the enabling scaffold with pre-filled frames

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Quick warm up! I am going to flash some dots on the screen
- Do NOT count them! Just look and tell me how many
- This is called subitising -- seeing how many without counting
- Ready? Look... how many?

DO:
- Flash dot patterns for 2-3 seconds each: 3 (triangle), 5 (dice), 4 (square), 2 (pair), 5 (different pattern)
- After each flash, students show answer with fingers
- Keep the pace quick -- the point is instant recognition, not counting

TEACHER NOTES:
Daily Review targets subitising to 5. Subitising is the ability to instantly recognise small quantities without counting. Standard arrangements (dice, domino patterns) are easiest. This is a separate skill from counting but supports it.

WATCH FOR:
- Students who count on their fingers during the flash -- they are counting, not subitising. That is okay at Foundation but note them
- Students who respond instantly with correct answers -- subitising is developing well

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers
- 3, 5, 4, 2, 5
- Ask: Which one was easiest to see straight away? [Usually 2 or 3]
- The patterns help us see how many without counting

DO:
- Reveal each pattern with its number
- Students self-check

TEACHER NOTES:
Discussing why patterns help (structured arrangements are easier to subitise) builds metacognition about number recognition.

WATCH FOR:
- Students who got all 5 correct -- strong subitising
- Students who missed the 5s -- they may need more structured pattern practice

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time! Numbers after
- I say a number, you tell me what comes AFTER
- Ready? What comes after 7? [8] After 12? [13] After 19? [20]
- Now try on your whiteboards. I will say the number, you write the one that comes after

DO:
- Verbal round first (8 questions, rapid fire)
- Then whiteboard round with: 5, 11, 14, 9, 16, 18, 13, 7
- Students write the number AFTER on their whiteboard
- Self-check against class answers

TEACHER NOTES:
"Numbers after" builds the counting sequence by testing whether students know what comes next without counting from 1. This is an important step toward number sense -- knowing the position of a number in the sequence. The transition from single digits to teens (9 -> 10, 19 -> 20) is the key challenge.

WATCH FOR:
- Students who struggle at the 9 -> 10 and 19 -> 20 transitions -- decade crossing is cognitively harder
- Students who are fast and accurate -- sequence knowledge is automated

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read from slide: "We are learning to show numbers to 20 in different ways"
- Our success criteria:
- Read from slide: "I can show a number using counters"
- Read from slide: "I can show a number using a ten frame"
- Read from slide: "I can draw a picture to show a number"

DO:
- Choral read LI and each SC
- Hold up counters, a ten frame, and a whiteboard: "These are our three tools today"

TEACHER NOTES:
Three representations are the focus: concrete (counters), structured pictorial (ten frame), and free pictorial (drawing). SC1 is ultra-achievable. SC2 builds on Session 2. SC3 extends to unstructured drawing which requires students to create their own representation.

WATCH FOR:
- Students who seem confident about ten frames after Session 2 -- they are ready to move quickly through SC2

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me show the number 16 in three different ways
- Way 1: Counters. I count out 16 counters. 1, 2, 3... 16.
- Way 2: Ten frame. I fill the frame -- that is 10. Then 6 extras. 10 and 6 makes 16.
- Way 3: Drawing. I draw 16 circles. I can group them to make it easier to see -- I draw 10 in a row and 6 more.
- All three ways show the same number: 16
- Ask: Which way is easiest for you to see that it is 16?

DO:
- Under the document camera, show 16 counters in a pile
- Then arrange them on a ten frame (10 + 6)
- Then draw 16 circles on the whiteboard (grouped: 10 + 6)
- Point to each: "Same number, different way of showing it"

TEACHER NOTES:
The key insight is that the SAME number can be represented multiple ways. The ten frame representation is the most structured and most useful for building place value understanding. Grouping the drawing into 10 + extras mirrors the ten frame structure.

MISCONCEPTIONS:
- Misconception: Each representation is a different number
  Why: Young students may think that 16 counters is a different quantity from 16 drawn circles because they look different
  Impact: Without understanding conservation of number across representations, students cannot flexibly move between models
  Quick correction: "Let's count each one. Counters: 16. Ten frame: 10 + 6 = 16. Drawing: 16 circles. Same number every time!"

WATCH FOR:
- Students who see the ten frame as the easiest to read -- the structure helps!
- Students who prefer the pile of counters -- they may not yet see the benefit of organised representations

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. I have shown a number on the screen using a ten frame
- Full frame plus 4 extras
- Write the number on your whiteboard. Ready... show me!

DO:
- Display a full ten frame + 4 extras
- Students write the number (14) on whiteboards
- Scan for 14

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Say: "Look at the ten frame. Full frame plus 4 extras. What number is this? Write it. Ready... show me!"
- Scan for: 14. Students who write 14 quickly can read ten frame representations.
PROCEED: If 80%+ show 14, move to We Do.
PIVOT: Most likely misconception -- students write 4 because they only counted the extras. Reteach: "The frame is full. How many in a full frame? [10] Plus 4 extras. 10 and 4 makes...? [14]"

TEACHER NOTES:
This CFU checks transfer from Session 2 -- can students read a ten frame representation and write the numeral?

WATCH FOR:
- Students who write 14 instantly -- Session 2 learning has stuck
- Students who write 4 -- they need the "full frame = 10" reminder

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn! With your partner, show the number 12 in three ways
- Way 1: Count out 12 counters
- Way 2: Show 12 on your ten frame (fill it, add extras)
- Way 3: Draw 12 circles on your whiteboard -- try grouping them as 10 and 2

DO:
- Partners work together
- Circulate: check all three representations
- After 2 minutes, cold call 3 pairs to show each representation
- Repeat with number 18

TEACHER NOTES:
We Do requires students to create all three representations -- this is harder than reading a given representation. Working in pairs provides support. Two numbers (12 and 18) give practice with small and large extras.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 3 Enabling Scaffold with pre-drawn ten frames and numeral guides. Students only need to add counters to the frame and draw dots in the provided spaces. The number is written in dotted lines to trace.
- Extra Notes: Distribute the Session 3 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Investigate how many different ways you can break 12 into two groups using counters or drawings (e.g. 10+2, 9+3, 8+4...). The extension PDF teaches partitioning and provides a recording table.
- Extra Notes: Distribute the Session 3 Extension PDF. Self-contained with worked examples.

WATCH FOR:
- Students who can do the ten frame but struggle with the free drawing -- suggest grouping: "Try drawing 10 in a line, then the extras underneath"
- Students who create all three representations quickly and accurately -- they understand multiple representations
- Readiness signal: both partners showing three correct representations within 2 minutes

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check. Table 2, show us your 12 with counters. How many? [12]
- Table 5, show your ten frame for 12. Full frame plus? [2 extras]
- Table 1, show your drawing of 12. How did you group them? [10 and 2]
- Now 18 -- who can show all three ways?

DO:
- Cold call different tables for each representation
- Celebrate grouping strategies in drawings
- Emphasise: "Same number, three different ways!"

TEACHER NOTES:
Drawing out the grouping strategy in free drawings is important -- it bridges to the structured ten frame representation and builds the "10 and ones" schema.

WATCH FOR:
- Students who drew 12 circles in a random arrangement -- they made the number but may not see the 10+2 structure
- Students who grouped their drawing as 10+2 without being told -- strong understanding

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check! I am going to show you a drawing of some objects
- I drew them in a special way. Tell me the number. Show me with your fingers

DO:
- Display a grouped drawing: 10 circles in a line + 5 below
- Students show 15 with fingers
- Follow up: "How did the grouping help you?" [I could see the 10 quickly]

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "Look at my drawing. I grouped them. How many altogether? Show me with fingers. Ready... show me!"
- Scan for: 15 (10 + 5 grouped drawing).
PROCEED: If 80%+ show 15, move to You Do.
PIVOT: Most likely misconception -- students try to count all dots one by one and lose track. Reteach: "See the top row? That is 10 -- I can see a full row. Now count the extras below: 1, 2, 3, 4, 5. 10 and 5 makes 15."

TEACHER NOTES:
This hinge checks whether students can read a grouped pictorial representation and connect it to the ten-and-ones structure.

WATCH FOR:
- Students who respond instantly -- they are subitising the group structure
- Students who count from 1 -- they are accurate but not yet using the grouping

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice time. Open your worksheet
- First: Look at each number
- Next: Show it on the ten frame (draw dots)
- Then: Draw the number as a picture (group them!)
- You have 6 minutes. Go!

DO:
- Distribute Session 3 Worksheet
- Students work independently
- Circulate: prioritise students who struggled with drawings
- For enabling students, distribute the Session 3 Enabling Scaffold
- For extending students, distribute the Session 3 Extension

TEACHER NOTES:
You Do requires students to create representations (not just read them). This is the harder direction of the CRA model. The worksheet provides blank ten frames and drawing spaces for each number.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 3 Enabling Scaffold with numbers to 12 only and pre-drawn ten frames (full frame provided). Students add extras and trace the numeral.
- Extra Notes: Distribute the Session 3 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Investigate number bonds to 20 -- how many ways can you split 12 into two groups? The extension PDF teaches partitioning with worked examples and a recording table.
- Extra Notes: Distribute the Session 3 Extension PDF. Self-contained.

WATCH FOR:
- Students who draw dots without grouping -- suggest: "Can you put them in rows of 10?"
- Students who complete the ten frame correctly but struggle with free drawing -- the structured representation is secure
- Readiness signal: 80%+ completing 4+ numbers with correct representations within 4 minutes

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket time. Three quick questions
- Work on your own. You have 2 minutes

DO:
- Display exit ticket questions on screen
- Students write answers in workbooks or on whiteboards

TEACHER NOTES:
Q1 assesses SC2 (show on ten frame), Q2 assesses SC3 (draw a picture), Q3 checks cross-representation understanding (match a drawing to its numeral).

WATCH FOR:
- Students who get Q1 correct but struggle with Q2 -- structured representations are secure, free drawing needs more practice
- Students who complete all three quickly -- ready for ordering in Session 4

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Let's check our success criteria
- SC1: I can show a number using counters. Thumbs?
- SC2: I can show a number using a ten frame. Thumbs?
- SC3: I can draw a picture to show a number. Thumbs?
- Turn and talk: What is your FAVOURITE way to show a number? Tell your partner why

DO:
- Display success criteria
- Run thumbs check for each SC
- Allow 30 seconds for Turn and Talk
- Cold call 2-3 students to share their favourite representation

TEACHER NOTES:
The discussion about favourite representations builds metacognition -- students are becoming aware of which tools help them think about numbers. Most Foundation students prefer the ten frame because it is structured.

WATCH FOR:
- Students who can articulate why they prefer a particular representation -- metacognition is developing
- Students who show thumbs down on SC3 -- they need more practice with free drawing

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
  titleSlide(pres, UNIT_TITLE, "Session 3: Showing Numbers Many Ways", "Foundation Numeracy | Session 3 of 5 | Term 2 Week 1", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2-3: Daily Review (withReveal) — Subitising to 5
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Subitising — How Many Can You See?", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      // 5 subitising patterns
      const patterns = [
        { n: 3, dots: [[0.5,0],[0,0.8],[1,0.8]] },
        { n: 5, dots: [[0,0],[1,0],[0.5,0.5],[0,1],[1,1]] },
        { n: 4, dots: [[0,0],[1,0],[0,1],[1,1]] },
        { n: 2, dots: [[0,0.5],[1,0.5]] },
        { n: 5, dots: [[0.5,0],[0,0.4],[1,0.4],[0,0.8],[1,0.8]] },
      ];

      patterns.forEach((pat, i) => {
        const baseX = 0.7 + i * 1.8;
        const baseY = CONTENT_TOP + 0.2;
        // Card for each pattern
        addCard(s, baseX, baseY, 1.5, 1.6, { strip: STAGE_COLORS["1"] });
        s.addText(String(i + 1), {
          x: baseX + 0.1, y: baseY + 0.02, w: 0.3, h: 0.25,
          fontSize: 10, fontFace: FONT_B, color: C.MUTED, margin: 0,
        });
        // Draw dots
        pat.dots.forEach(([dx, dy]) => {
          s.addShape("roundRect", {
            x: baseX + 0.25 + dx * 0.7, y: baseY + 0.35 + dy * 0.7, w: 0.3, h: 0.3, rectRadius: 0.15,
            fill: { color: C.PRIMARY },
          });
        });
        // "?" below
        s.addText("?", {
          x: baseX + 0.4, y: baseY + 1.2, w: 0.7, h: 0.3,
          fontSize: 20, fontFace: FONT_H, color: C.ALERT, align: "center", valign: "middle", bold: true, margin: 0,
        });
      });

      addCard(s, 0.5, CONTENT_TOP + 2.1, 9, 1.0, { strip: C.ACCENT });
      s.addText([
        { text: "Do NOT count! Just LOOK.", options: { fontSize: 18, bold: true, color: C.ACCENT, breakLine: true } },
        { text: "Show me with your fingers how many you see in each group.", options: { fontSize: 14, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 2.22, w: 8.5, h: 0.8,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Answers: 3, 5, 4, 2, 5", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Numbers after 1-20
  const sFluency = pres.addSlide();
  addTopBar(sFluency, STAGE_COLORS["1"]);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "What Comes After?", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

  addCard(sFluency, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  const afterNums = [
    { q: "After 5?", a: "6" }, { q: "After 11?", a: "12" },
    { q: "After 14?", a: "15" }, { q: "After 9?", a: "10" },
    { q: "After 16?", a: "17" }, { q: "After 18?", a: "19" },
    { q: "After 13?", a: "14" }, { q: "After 7?", a: "8" },
  ];
  sFluency.addText(afterNums.map((item, i) => ({
    text: (i + 1) + ".  " + item.q + "   ___",
    options: { fontSize: 16, color: C.CHARCOAL, breakLine: i < afterNums.length - 1, paraSpaceAfter: 6 },
  })), {
    x: 0.75, y: CONTENT_TOP + 0.1, w: 3.8, h: SAFE_BOTTOM - CONTENT_TOP - 0.2,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addCard(sFluency, 5.2, CONTENT_TOP, 4.3, 1.8, { strip: C.SECONDARY });
  sFluency.addText([
    { text: "Verbal round first!", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
    { text: "I say a number.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
    { text: "You say what comes AFTER.", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Then write the answers on your whiteboard!", options: { fontSize: 14, color: C.CHARCOAL, bold: true } },
  ], {
    x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: 1.6,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addTextOnShape(sFluency, "Answers: 6, 12, 15, 10, 17, 19, 14, 8", {
    x: 5.2, y: CONTENT_TOP + 2.1, w: 4.3, h: 0.4, rectRadius: 0.06,
    fill: { color: C.MUTED },
  }, { fontSize: 11, fontFace: FONT_B, color: C.WHITE });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    ["We are learning to show numbers to 20 in different ways"],
    [
      "I can show a number using counters",
      "I can show a number using a ten frame",
      "I can draw a picture to show a number",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — Showing 16 three ways
  workedExSlide(pres, 2, "I Do", "Three Ways to Show 16",
    [
      "Way 1: Counters",
      "  Count out 16 counters",
      "",
      "Way 2: Ten frame",
      "  Full frame (10) + 6 extras",
      "",
      "Way 3: Drawing",
      "  Draw 16 circles grouped as 10 + 6",
      "",
      "Same number, different ways!",
    ],
    NOTES_IDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.2, { strip: C.PRIMARY });
      slide.addText("Showing 16", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.05, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      // Ten frame showing 16
      const tfX = lg.rightX + 0.3;
      const tfY = lg.panelTopPadded + 0.42;
      drawTenFrame(slide, tfX, tfY, 10, { cellW: 0.35, cellH: 0.35 });
      // 6 extras
      for (let i = 0; i < 6; i++) {
        slide.addShape("roundRect", {
          x: tfX + i * 0.35, y: tfY + 0.78, w: 0.3, h: 0.3, rectRadius: 0.15,
          fill: { color: C.ALERT },
        });
      }
      addTextOnShape(slide, "10 + 6 = 16", {
        x: lg.rightX + 0.3, y: tfY + 1.2, w: lg.rightW - 0.6, h: 0.35, rectRadius: 0.06,
        fill: { color: C.SECONDARY },
      }, { fontSize: 14, fontFace: FONT_H, color: C.WHITE, bold: true });

      // Drawing: grouped circles
      slide.addText("Drawing (grouped):", {
        x: lg.rightX + 0.3, y: tfY + 1.7, w: lg.rightW - 0.6, h: 0.25,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0,
      });
      // 10 small circles in a row
      for (let i = 0; i < 10; i++) {
        slide.addShape("roundRect", {
          x: lg.rightX + 0.3 + i * 0.35, y: tfY + 2.0, w: 0.25, h: 0.25, rectRadius: 0.12,
          fill: { color: C.PRIMARY },
        });
      }
      // 6 more below
      for (let i = 0; i < 6; i++) {
        slide.addShape("roundRect", {
          x: lg.rightX + 0.3 + i * 0.35, y: tfY + 2.3, w: 0.25, h: 0.25, rectRadius: 0.12,
          fill: { color: C.ALERT },
        });
      }
    }
  );

  // Slide 7-8: CFU 1 (withReveal) — Read a ten frame
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["2"]);
      addStageBadge(s, 2, "I Do");
      addTitle(s, "What Number Is This?", { y: 0.65, fontSize: 22, color: STAGE_COLORS["2"] });

      addCard(s, 0.5, CONTENT_TOP, 9, 2.0, { strip: C.PRIMARY });
      // Full ten frame + 4 extras
      drawTenFrame(s, 1.5, CONTENT_TOP + 0.3, 10, { cellW: 0.5, cellH: 0.5 });
      s.addText("+", {
        x: 4.3, y: CONTENT_TOP + 0.5, w: 0.5, h: 0.5,
        fontSize: 30, fontFace: FONT_H, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
      });
      for (let i = 0; i < 4; i++) {
        s.addShape("roundRect", {
          x: 5.0 + i * 0.6, y: CONTENT_TOP + 0.55, w: 0.45, h: 0.45, rectRadius: 0.22,
          fill: { color: C.ALERT },
        });
      }

      addCard(s, 0.5, CONTENT_TOP + 2.3, 9, 0.9, { strip: C.ACCENT });
      s.addText("Write the number on your whiteboard!", {
        x: 0.75, y: CONTENT_TOP + 2.45, w: 8.5, h: 0.55,
        fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, bold: true, margin: 0, valign: "middle",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "14!  (10 + 4 = 14)", {
        x: 2.5, y: 4.2, w: 5.0, h: 0.6, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 9-10: We Do (withReveal) — Show 12 and 18 three ways
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Show It Three Ways!",
      [
        "With your partner, show each number",
        "in THREE ways:",
        "",
        "1. Counters (count them out)",
        "2. Ten frame (fill it + extras)",
        "3. Drawing (group as 10 + ones)",
        "",
        "Number 1: 12     Number 2: 18",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.2, { strip: C.SECONDARY });
        slide.addText("Three Ways:", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
          fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });

        // Visual icons for three representations
        const items = [
          { label: "Counters", icon: "Count them out" },
          { label: "Ten Frame", icon: "Fill + extras" },
          { label: "Drawing", icon: "Group as 10 + ones" },
        ];
        items.forEach((item, i) => {
          addTextOnShape(slide, item.label + ": " + item.icon, {
            x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.5 + i * 0.55, w: lg.rightW - 0.4, h: 0.42, rectRadius: 0.06,
            fill: { color: i === 1 ? C.PRIMARY : C.ACCENT },
          }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });
        });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "12 = 10 + 2      18 = 10 + 8", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 11-12: CFU 2 Hinge (withReveal) — Read a grouped drawing
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "How Many Altogether?", { y: 0.65, fontSize: 22, color: STAGE_COLORS["3"] });

      addCard(s, 0.5, CONTENT_TOP, 9, 2.2, { strip: C.PRIMARY });
      // Grouped drawing: 10 circles in a row + 5 below
      for (let i = 0; i < 10; i++) {
        s.addShape("roundRect", {
          x: 0.8 + i * 0.8, y: CONTENT_TOP + 0.3, w: 0.6, h: 0.6, rectRadius: 0.3,
          fill: { color: C.PRIMARY },
        });
      }
      for (let i = 0; i < 5; i++) {
        s.addShape("roundRect", {
          x: 0.8 + i * 0.8, y: CONTENT_TOP + 1.1, w: 0.6, h: 0.6, rectRadius: 0.3,
          fill: { color: C.ALERT },
        });
      }

      addCard(s, 0.5, CONTENT_TOP + 2.4, 9, 0.9, { strip: C.ACCENT });
      s.addText("Show me with your FINGERS! How many altogether?", {
        x: 0.75, y: CONTENT_TOP + 2.52, w: 8.5, h: 0.65,
        fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, bold: true, margin: 0, valign: "middle",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU2);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "15!  (10 + 5 = 15)", {
        x: 2.0, y: 4.2, w: 6.0, h: 0.6, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 13: You Do
  workedExSlide(pres, 4, "You Do", "Show Numbers Your Way",
    [
      "First: Look at each number on your worksheet.",
      "Next: Show it on the ten frame (draw dots).",
      "Then: Draw it as a picture (group them!).",
      "",
      "Try to group your drawings as 10 + ones.",
      "You have 6 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.5, { strip: C.ALERT });
      slide.addText("Remember:", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Ten frame = fill 10 + extras", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "Drawing = group as 10 + ones", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "Same number, different ways!", options: { bullet: true, fontSize: 14, color: C.PRIMARY, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.5, h: 0.9,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Slide 14: Exit Ticket
  exitTicketSlide(pres,
    [
      "Show 13 on a ten frame. (Draw it.)",
      "Draw a picture of 17. Group your circles!",
      "Match: ten frame full + 9 extras = what number?",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 15: Closing
  closingSlide(pres,
    "What is your FAVOURITE way to show a number? Tell your partner why.",
    [
      "I can show a number using counters",
      "I can show a number using a ten frame",
      "I can draw a picture to show a number",
    ],
    NOTES_CLOSING);

  // Slide 16: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "N20_Session3_Representing_Numbers.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Generate PDFs ─────────────────────────────────────────────────────────

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Showing Numbers Many Ways",
      color: C.NAVY,
      lessonInfo: "Session 3 of 5 | Foundation Numeracy",
    });
    y = addTipBox(doc, "Show each number using a ten frame AND a drawing. Group your drawings as 10 + ones!", y, { color: C.TEAL });

    const drawRepTask = (doc, num, taskNum, y) => {
      doc.fontSize(14).font("Sans-Bold").fillColor("#" + C.NAVY).text(taskNum + ".  Show the number " + num, 55, y);
      y += 22;
      // Ten frame
      doc.fontSize(10).font("Sans").fillColor("#333333").text("Ten frame:", 70, y);
      const cellSize = 20;
      const gap = 2;
      const tfX = 140;
      for (let row = 0; row < 2; row++) {
        for (let col = 0; col < 5; col++) {
          doc.rect(tfX + col * (cellSize + gap), y + row * (cellSize + gap), cellSize, cellSize)
            .lineWidth(0.8).strokeColor("#" + C.NAVY).stroke();
        }
      }
      doc.text("+ extras:", tfX + 5 * (cellSize + gap) + 5, y + 10);
      y += cellSize * 2 + gap + 10;
      // Drawing space
      doc.fontSize(10).font("Sans").fillColor("#333333").text("Drawing:", 70, y);
      doc.rect(140, y - 2, 350, 40).lineWidth(0.5).strokeColor("#CCCCCC").stroke();
      y += 48;
      return y;
    };

    y = addSectionHeading(doc, "Show each number two ways", y, { color: C.NAVY });
    y = drawRepTask(doc, 11, "1", y);
    y = drawRepTask(doc, 14, "2", y);
    y = drawRepTask(doc, 16, "3", y);
    y = drawRepTask(doc, 19, "4", y);
    y = drawRepTask(doc, 13, "5", y);
    y = drawRepTask(doc, 20, "6", y);

    addPdfFooter(doc, "Session 3 | Numbers to 20 | Foundation");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Showing Numbers Many Ways",
      color: C.NAVY,
      lessonInfo: "Session 3 of 5 | Foundation Numeracy",
    });
    y = addSectionHeading(doc, "Answers", y, { color: C.NAVY });
    y = addBodyText(doc, "1. 11: Full frame + 1 extra. Drawing: 10 circles + 1 circle.", y);
    y = addBodyText(doc, "2. 14: Full frame + 4 extras. Drawing: 10 circles + 4 circles.", y);
    y = addBodyText(doc, "3. 16: Full frame + 6 extras. Drawing: 10 circles + 6 circles.", y);
    y = addBodyText(doc, "4. 19: Full frame + 9 extras. Drawing: 10 circles + 9 circles.", y);
    y = addBodyText(doc, "5. 13: Full frame + 3 extras. Drawing: 10 circles + 3 circles.", y);
    y = addBodyText(doc, "6. 20: Two full ten frames! Drawing: 10 circles + 10 circles.", y);
    addPdfFooter(doc, "Session 3 | Answer Key | Foundation");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Representation Mat",
      color: C.TEAL,
      lessonInfo: "Session 3 of 5 | Foundation Numeracy",
    });
    y = addTipBox(doc, "The ten frame is already full (10). Add the extras. Trace the number!", y, { color: C.TEAL });

    const drawEnablingTask = (doc, extras, numeral, taskNum, y) => {
      doc.fontSize(14).font("Sans-Bold").fillColor("#" + C.NAVY).text(taskNum + ".  Make " + numeral, 55, y);
      y += 20;
      const cellSize = 20;
      const gap = 2;
      const tfX = 80;
      // Pre-filled full ten frame
      for (let row = 0; row < 2; row++) {
        for (let col = 0; col < 5; col++) {
          doc.rect(tfX + col * (cellSize + gap), y + row * (cellSize + gap), cellSize, cellSize)
            .lineWidth(0.8).strokeColor("#" + C.NAVY).stroke();
          doc.circle(tfX + col * (cellSize + gap) + cellSize / 2, y + row * (cellSize + gap) + cellSize / 2, 7)
            .fill("#" + C.SECONDARY);
        }
      }
      // Extras space with circles to fill
      const extraX = tfX + 5 * (cellSize + gap) + 15;
      doc.fontSize(14).font("Sans-Bold").fillColor("#333333").text("+", extraX - 12, y + 8);
      for (let i = 0; i < extras; i++) {
        doc.circle(extraX + i * 22, y + cellSize / 2 + gap, 8)
          .lineWidth(1).strokeColor("#" + C.ALERT).stroke();
      }
      // Dotted numeral
      const traceX = extraX + Math.max(extras, 1) * 22 + 30;
      doc.fontSize(32).font("Sans").fillColor("#CCCCCC").text(numeral, traceX, y - 4);
      return y + cellSize * 2 + gap + 18;
    };

    y = addSectionHeading(doc, "Add extras and trace the number", y, { color: C.NAVY });
    y = drawEnablingTask(doc, 1, "11", "1", y);
    y = drawEnablingTask(doc, 2, "12", "2", y);
    y = drawEnablingTask(doc, 4, "14", "3", y);
    y = drawEnablingTask(doc, 5, "15", "4", y);
    y = drawEnablingTask(doc, 3, "13", "5", y);
    y = drawEnablingTask(doc, 6, "16", "6", y);

    addPdfFooter(doc, "Session 3 | Enabling Scaffold | Foundation");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Extension — Number bonds to 20
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Number Bonds - Breaking Numbers Apart",
      color: C.TEAL,
      lessonInfo: "Session 3 of 5 | Foundation Numeracy",
    });

    y = addSectionHeading(doc, "What are number bonds?", y, { color: C.NAVY });
    y = addBodyText(doc, "A number bond shows how a number can be split into two parts.", y);
    y = addBodyText(doc, "For example, 12 can be split into 10 and 2, or 9 and 3, or 8 and 4!", y);

    y = addSectionHeading(doc, "Example: Bonds of 12", y, { color: C.NAVY });
    y = addBodyText(doc, "12 = 10 + 2", y);
    y = addBodyText(doc, "12 = 9 + 3", y);
    y = addBodyText(doc, "12 = 8 + 4", y);
    y = addBodyText(doc, "12 = 7 + 5", y);
    y = addBodyText(doc, "12 = 6 + 6", y);
    y = addBodyText(doc, "The parts always add up to 12!", y);

    y = addSectionHeading(doc, "Your Investigation", y, { color: C.NAVY });
    y = addBodyText(doc, "Find as many ways as you can to split each number into two parts.", y);
    y = addBodyText(doc, "Use your counters to help! Put some in one group and some in another.", y);

    y = addSectionHeading(doc, "Bonds of 10", y, { color: C.NAVY });
    y = addWriteLine(doc, "10 = ___ + ___", y);
    y = addWriteLine(doc, "10 = ___ + ___", y);
    y = addWriteLine(doc, "10 = ___ + ___", y);
    y = addWriteLine(doc, "10 = ___ + ___", y);
    y = addWriteLine(doc, "10 = ___ + ___", y);

    y = addSectionHeading(doc, "Bonds of 15", y, { color: C.NAVY });
    y = addWriteLine(doc, "15 = ___ + ___", y);
    y = addWriteLine(doc, "15 = ___ + ___", y);
    y = addWriteLine(doc, "15 = ___ + ___", y);
    y = addWriteLine(doc, "15 = ___ + ___", y);

    y = addSectionHeading(doc, "Did You Know?", y, { color: C.NAVY });
    y = addBodyText(doc, "The number 10 has the MOST bonds because it is a special number in our number system. Knowing bonds of 10 helps with all sorts of maths!", y);

    addPdfFooter(doc, "Session 3 | Extension | Foundation");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 3 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
