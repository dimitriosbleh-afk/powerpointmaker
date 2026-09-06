"use strict";

// Numbers to 20 Unit — Session 1: Counting Collections to 10
// Term 2 Week 1, Foundation Numeracy, Variant 0
// DR: Counting — I can count up to at least 20 fluently by 1s
// Fluency: Counting from 1-20 (choral count, count-around)
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

const SESSION = 1;
const UNIT_TITLE = "Numbers to 20";
const FOOTER = "Numbers to 20 | Session 1 of 5 | Foundation Numeracy";
const OUT_DIR = "output/N20_Session1_Counting_Collections";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Worksheet", "Independent practice - counting and matching collections to 10.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key", "Teacher reference with answers.");
const ENABLING_RES = makeSessionResource(SESSION, "Enabling Scaffold", "Pre-structured counting mat with touch-point guides for collections to 5.");
const EXTENDING_RES = makeSessionResource(SESSION, "Extension", "Self-contained investigation into counting collections beyond 10.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, ENABLING_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Teacher Notes ───────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Welcome back! We have been on holidays and now we are ready to learn
- This week we are working with numbers to 20
- Today we are going to count collections -- groups of objects

DO:
- Display title slide as students settle on the mat
- Have counting collections (mini erasers, teddy bears, cubes) ready in zip-lock bags
- Distribute whiteboards and markers

TEACHER NOTES:
Session 1 of 5. This is the first maths lesson back from a 16-day break. Content is pitched gently to rebuild routines and confidence. The focus is on 1-1 correspondence and cardinality with collections to 10 before extending to teen numbers in Sessions 2-3.

WATCH FOR:
- Students who seem unsettled after the break -- keep the tone warm and encouraging
- Students who have forgotten routines -- briefly re-establish mat expectations

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_DR_Q = `SAY:
- Let's warm up with some counting
- I am going to show you a number strip. Some numbers are missing
- On your whiteboards, write the missing numbers
- Ready? Have a go

DO:
- Display the counting strip with gaps
- Allow 60 seconds for students to write missing numbers on whiteboards
- Hold whiteboards face down until signal to show

TEACHER NOTES:
Daily Review retrieves counting fluency from last term. The counting strip tests sequential counting to 20, which was the focus last week. Gaps are placed at teen numbers since those are trickier after a break.

WATCH FOR:
- Students who reverse teen number digits (e.g. writing 31 for 13) -- this is common and will be addressed in Session 2
- Students who skip numbers in the teen sequence -- note for targeted support

[Stage 1: Daily Review | VTLM 2.0: Retention and Recall]`;

const NOTES_DR_A = `SAY:
- Check your answers -- tick if you got it right, fix it if you didn't
- Fixing it is the learning
- Let's count together from 1 to 20. Ready? 1, 2, 3...

DO:
- Click to reveal answers
- Students tick correct responses and fix errors
- Lead a choral count 1-20 after checking
- Scan for students who struggled with teen numbers

TEACHER NOTES:
Tick-and-fix gives immediate feedback. The follow-up choral count rebuilds the counting sequence after the holiday break.

WATCH FOR:
- Students who got all numbers correct -- counting fluency retained well
- Students who missed teen numbers -- they will need extra support in Sessions 2-3

[Stage 1: Daily Review Answers | VTLM 2.0: Retention and Recall]`;

const NOTES_FLUENCY = `SAY:
- Fluency time! Stand up, everyone
- We are going to do a count-around the circle
- I will start at 1. The person next to me says 2. Keep going around
- If you get stuck, your neighbour can help

DO:
- Students stand in a circle (or at desks in a wave pattern)
- Count from 1 to 20, one number per student
- If the class is small, go around twice
- Repeat with increasing speed
- Sit down after 2-3 rounds

TEACHER NOTES:
Count-around builds counting fluency in an active, engaging format suited to Foundation. Standing up shifts the energy after mat time. The neighbour-help rule keeps it safe for students who are rebuilding confidence after the break.

WATCH FOR:
- Students who hesitate at the transition from 9 to 10 or 12 to 13 -- these are common stumble points
- Students who are confident and quick -- they can lead the next round

[Stage 1: Fluency | VTLM 2.0: Automaticity]`;

const NOTES_LI_SC = `SAY:
- Read from slide: "We are learning to count groups of things up to 10"
- Let's read our success criteria together
- Read from slide: "I can touch and count each object one at a time"
- Read from slide: "I can say how many are in a group"
- Read from slide: "I can match a number to the right group"

DO:
- Choral read the LI then each SC, pointing to each on screen
- Use fingers to demonstrate: hold up 5 fingers and count them together

TEACHER NOTES:
This lesson targets the foundational counting principles: 1-1 correspondence (SC1), cardinality (SC2), and numeral-quantity matching (SC3). These are deliberately gentle for the first session back. SC1 is ultra-achievable -- every student can touch objects and count.

WATCH FOR:
- Students who seem confused by "success criteria" -- they may have forgotten the routine. Briefly explain: "These are the things you will be able to do by the end of our lesson"

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO = `SAY:
- Watch me. I have a group of teddies. I need to find out how many
- My rule: I touch each one, I say the number, I move it to the other side
- Watch: touch... 1. Touch... 2. Touch... 3. Touch... 4. Touch... 5. Touch... 6.
- How many teddies? 6. The last number I said tells me how many altogether
- That is called "how many" -- the last number is the answer
- Now watch what happens if I count them a different way -- I'll start from this end
- Touch... 1. Touch... 2. Touch... 3. Touch... 4. Touch... 5. Touch... 6.
- Still 6! It doesn't matter which end I start from

DO:
- Use physical teddies or counters under the document camera
- Exaggerate the touch-and-move action so students can see 1-1 correspondence
- Line objects up in a row for the count
- Demonstrate counting from both ends to show conservation

MISCONCEPTIONS:
- Misconception: The last number is just the name for the last object, not the total
  Why: Young children may not yet understand cardinality -- that the last number in a count tells the total quantity
  Impact: Students will be able to recite the counting sequence but not answer "how many" questions
  Quick correction: After counting, sweep your hand across all objects and say "So there are 6 altogether. 6 is how many."

TEACHER NOTES:
This I Do models the three key counting principles: 1-1 correspondence (touch each once), stable order (say numbers in sequence), and cardinality (last number = total). The touch-and-move strategy makes counting physical and traceable. Foundation students need to see the count as a physical, deliberate act.

WATCH FOR:
- Students who mouth the numbers along with you -- good sign of engagement
- Students who seem distracted -- bring them back with "Watch my hands"
- Readiness signal: students able to answer "How many?" after watching the count

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. I have a group of cubes on the screen
- How many cubes? Count in your head. Show me with your fingers

DO:
- Display a group of 4 cubes arranged in a line
- Allow 5 seconds for students to finger-count
- Say: "Show me on my signal... now!"
- Scan for 4 fingers

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- Say: "Count the cubes. When I say show me, hold up the right number of fingers. Ready... show me!"
- Scan for: 4 fingers. Students who show 4 quickly have understood 1-1 correspondence and cardinality.
PROCEED: If 80%+ show 4 fingers, move to We Do.
PIVOT: Most likely misconception -- students show a number other than 4 because they double-counted or skipped an object. Reteach: point to each cube slowly, counting aloud together as a class. "Touch, 1. Touch, 2. Touch, 3. Touch, 4. How many? 4. The last number tells us how many."

TEACHER NOTES:
Finger voting is ideal for Foundation CFU -- fast, visible, and every student responds. The image uses a line arrangement to match the I Do modelling.

WATCH FOR:
- Students who show fingers before counting -- they may be guessing
- Students who count on their fingers rather than counting the objects on screen -- redirect: "Count the cubes, then show me with your fingers"

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn! I am going to give each table a zip-lock bag of objects
- Tip them out carefully. Count how many are in your bag
- Remember our rule: touch each one, say the number, move it
- The last number you say is how many altogether
- Write the number on your whiteboard

DO:
- Distribute pre-counted zip-lock bags (each has 5-8 objects)
- Circulate to watch counting technique -- are students touching each object once?
- Allow 2 minutes
- Cold call 3-4 students: "How many did you count? Show us how you counted."

TEACHER NOTES:
We Do uses physical manipulatives so students practise the touch-and-move strategy modelled in I Do. Bags are pre-counted so the teacher can verify accuracy. Different tables get different quantities (5, 6, 7, 8) to create variety in sharing.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 1 Enabling Scaffold -- a counting mat with numbered touch-point circles (1-5). Students place one object on each circle and say the number. The structure prevents double-counting.
- Extra Notes: Distribute the Session 1 Enabling Scaffold PDF for students who need the physical structure.
EXTENDING PROMPT:
- Task: Count your collection, then use a ten frame to show the number. Draw the ten frame on your whiteboard with the right number of dots.
- Extra Notes: Students who finish quickly get a second bag with 9-12 objects (crossing into teen territory).

WATCH FOR:
- Students who rush and skip objects -- slow them down: "Touch... say the number... move it"
- Students who lose track and recount from 1 each time -- help them keep a clear "counted" and "not yet counted" pile
- Readiness signal: students who count accurately and write the correct numeral confidently

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Let's check. Table 1, how many in your bag? [e.g. 7] Show us how you counted
- Table 3, yours? [e.g. 5]
- Ask: How do you know that is the right number? [Because it was the last number I said]
- Great! The last number tells us how many altogether

DO:
- Reveal answer counts for each table group
- Cold call students to demonstrate their counting technique
- Praise the touch-and-move strategy when you see it

TEACHER NOTES:
The key response to listen for is cardinality language: "because it was the last number I said" or "because I counted 7." This confirms SC2.

WATCH FOR:
- Students who say the right number but cannot explain how they know -- they may be subitising small quantities rather than counting
- Students who demonstrate clear touch-and-move -- affirm this publicly

[Stage 3: We Do Answers | VTLM 2.0: Scaffold Practice]`;

const NOTES_CFU2 = `SAY:
- Hinge check. Look at the screen
- I have two groups of dots. Which group has more? Show me -- point left or right
- Think... show me now

DO:
- Display two groups: left has 3 dots, right has 7 dots
- Students point to the group with more
- Follow up: "How do you know?" Cold call 2 students

CFU CHECKPOINT:
Technique: Point and Show
Script:
- Say: "Look at the two groups. Which group has MORE? Point to the group with more. Ready... point!"
- Scan for: students pointing right (7 dots). Most students should identify the larger group.
PROCEED: If 80%+ point to the right group, move to You Do.
PIVOT: Most likely misconception -- students point to the group that is more spread out even though it has fewer dots (they confuse area with quantity). Reteach: "Let's count each group. Left: 1, 2, 3. Right: 1, 2, 3, 4, 5, 6, 7. Which number is bigger? 7. So the right group has more."

TEACHER NOTES:
This hinge checks whether students can apply counting to make a comparison. It previews the ordering work in Session 4.

WATCH FOR:
- Students who respond instantly without counting -- they may be subitising (fine for 3 vs 7)
- Students who point to the more spread-out group regardless of quantity -- they need counting to compare

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent practice time
- Open your worksheet
- First: Count the objects in each group
- Next: Write the number
- Then: Draw a line to match the group to the right number
- You have 6 minutes. Go!

DO:
- Distribute Session 1 Worksheet
- Students work independently at desks
- Circulate: prioritise students who struggled during We Do
- For enabling students, distribute the Session 1 Enabling Scaffold
- For extending students, distribute the Session 1 Extension

TEACHER NOTES:
You Do uses different collections from the We Do (pictures instead of physical objects). This checks transfer from concrete manipulation to pictorial counting. The worksheet progresses from collections of 3-5 (Section A) to 6-10 (Section B).

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 1 Enabling Scaffold with counting mats for collections to 5 only. Students place a counter on each picture as they count, preventing double-counting. Numbers are pre-printed beside each collection.
- Extra Notes: Distribute the Session 1 Enabling Scaffold PDF.
EXTENDING PROMPT:
- Task: Investigate counting collections beyond 10 using ten frames. The extension shows collections of 11-15 objects. Students fill a ten frame, count the extras, and write the teen number.
- Extra Notes: Distribute the Session 1 Extension PDF. This is self-contained with ten frame templates and worked examples.

WATCH FOR:
- Students who count pictures by touching the page -- good strategy, affirm it
- Students who lose track with larger collections (7-10) -- suggest crossing out each object as they count
- Readiness signal: 80%+ completing Section A correctly within 3 minutes

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket time. Three quick questions
- Count each group and write the number
- Work on your own. You have 2 minutes

DO:
- Display the exit ticket on screen
- Students write answers in workbooks or on whiteboards
- Collect or photograph responses to sort into secure, developing, beginning

TEACHER NOTES:
Exit ticket assesses SC1 (touch-count) and SC2 (cardinality -- writing how many). Q1 targets 4 objects, Q2 targets 7, Q3 asks students to draw a group of 5. Sort results to plan support for Session 2.

WATCH FOR:
- Students who get Q1-Q2 correct but struggle with Q3 (drawing) -- they can count but struggle to create a representation
- Students who complete all three instantly -- ready for teen number work in Session 2

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Let's check our success criteria. Look at the slide
- SC1: I can touch and count each object one at a time. Thumbs up, sideways, or down
- SC2: I can say how many are in a group. Thumbs?
- SC3: I can match a number to the right group. Thumbs?
- Turn and talk: Tell your partner -- what is the rule for counting? [Touch each one, say the number, the last number tells you how many]

DO:
- Display success criteria on screen
- Run thumbs check for each SC
- Allow 30 seconds for Turn and Talk
- Cold call 2 students to share

TEACHER NOTES:
Closing brings the lesson full circle. The Turn and Talk targets the cardinality principle -- the most important understanding from today. Students who can say "the last number tells me how many" have grasped the core concept.

WATCH FOR:
- Students who show thumbs down on SC1 -- they need more physical counting practice in Session 2
- Students who articulate the counting rule clearly -- they are ready for teen numbers

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- These are the resources for today's lesson
- The worksheet and answer key are linked here if you need to reprint

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
  titleSlide(pres, UNIT_TITLE, "Session 1: Counting Collections to 10", "Foundation Numeracy | Session 1 of 5 | Term 2 Week 1", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 2-3: Daily Review (withReveal) — Counting to 20
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Count to 20 — Fill the Gaps!", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

      // Number strip with gaps
      addCard(s, 0.5, CONTENT_TOP, 9, 1.6, { strip: STAGE_COLORS["1"] });
      const numbers = [
        { n: "1", show: true }, { n: "2", show: true }, { n: "3", show: true },
        { n: "4", show: true }, { n: "5", show: true }, { n: "6", show: false },
        { n: "7", show: true }, { n: "8", show: true }, { n: "9", show: true },
        { n: "10", show: true },
      ];
      const numbers2 = [
        { n: "11", show: true }, { n: "12", show: false }, { n: "13", show: true },
        { n: "14", show: true }, { n: "15", show: false }, { n: "16", show: true },
        { n: "17", show: true }, { n: "18", show: false }, { n: "19", show: true },
        { n: "20", show: true },
      ];

      const cellW = 0.82;
      const cellH = 0.5;
      const startX = 0.75;
      // Row 1: 1-10
      numbers.forEach((item, i) => {
        const cx = startX + i * (cellW + 0.08);
        s.addShape("roundRect", {
          x: cx, y: CONTENT_TOP + 0.15, w: cellW, h: cellH, rectRadius: 0.08,
          fill: { color: item.show ? C.PRIMARY : C.WHITE },
          line: { color: C.PRIMARY, width: 1.5 },
        });
        s.addText(item.show ? item.n : "?", {
          x: cx, y: CONTENT_TOP + 0.15, w: cellW, h: cellH,
          fontSize: 20, fontFace: FONT_H, color: item.show ? C.WHITE : C.ALERT,
          align: "center", valign: "middle", bold: true, margin: 0,
        });
      });
      // Row 2: 11-20
      numbers2.forEach((item, i) => {
        const cx = startX + i * (cellW + 0.08);
        s.addShape("roundRect", {
          x: cx, y: CONTENT_TOP + 0.15 + cellH + 0.12, w: cellW, h: cellH, rectRadius: 0.08,
          fill: { color: item.show ? C.SECONDARY : C.WHITE },
          line: { color: C.SECONDARY, width: 1.5 },
        });
        s.addText(item.show ? item.n : "?", {
          x: cx, y: CONTENT_TOP + 0.15 + cellH + 0.12, w: cellW, h: cellH,
          fontSize: 20, fontFace: FONT_H, color: item.show ? C.WHITE : C.ALERT,
          align: "center", valign: "middle", bold: true, margin: 0,
        });
      });

      // Instruction card
      addCard(s, 0.5, CONTENT_TOP + 2.0, 9, 1.3, { strip: C.ACCENT });
      s.addText([
        { text: "On your whiteboards:", options: { fontSize: 17, bold: true, color: C.ACCENT, breakLine: true } },
        { text: "Write the missing numbers in the counting strip.", options: { fontSize: 15, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 2.12, w: 8.5, h: 1.0,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Answers: 6, 12, 15, 18", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR_A);
    }
  );

  // Slide 4: Fluency — Counting from 1-20 (count-around)
  const sFluency = pres.addSlide();
  addTopBar(sFluency, STAGE_COLORS["1"]);
  addStageBadge(sFluency, 1, "Fluency");
  addTitle(sFluency, "Count-Around: 1 to 20", { y: 0.65, fontSize: 22, color: STAGE_COLORS["1"] });

  addCard(sFluency, 0.5, CONTENT_TOP, 4.3, SAFE_BOTTOM - CONTENT_TOP, { strip: STAGE_COLORS["1"] });
  // Visual number path 1-20
  const pathNums = [];
  for (let i = 1; i <= 20; i++) pathNums.push(String(i));
  sFluency.addText(pathNums.join("    "), {
    x: 0.75, y: CONTENT_TOP + 0.15, w: 3.8, h: 1.0,
    fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, margin: 0, valign: "top",
    fit: "shrink",
  });
  sFluency.addText([
    { text: "Count-Around Rules:", options: { fontSize: 15, bold: true, color: C.PRIMARY, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Stand up in a circle", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
    { text: "Each person says the next number", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
    { text: "If you get stuck, your neighbour helps", options: { bullet: true, fontSize: 14, color: C.CHARCOAL, breakLine: true } },
    { text: "Count from 1 to 20!", options: { bullet: true, fontSize: 14, color: C.CHARCOAL } },
  ], {
    x: 0.75, y: CONTENT_TOP + 1.3, w: 3.8, h: 2.2,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addCard(sFluency, 5.2, CONTENT_TOP, 4.3, 2.2, { strip: C.SECONDARY });
  sFluency.addText([
    { text: "Round 1:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
    { text: "Count 1 to 20 slowly", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Round 2:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
    { text: "Count 1 to 20 faster!", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
    { text: "", options: { fontSize: 6, breakLine: true } },
    { text: "Round 3:", options: { fontSize: 16, bold: true, color: C.SECONDARY, breakLine: true } },
    { text: "Count BACKWARDS from 20 to 1", options: { fontSize: 14, color: C.CHARCOAL } },
  ], {
    x: 5.45, y: CONTENT_TOP + 0.1, w: 3.8, h: 2.0,
    fontFace: FONT_B, margin: 0, valign: "top",
  });

  addFooter(sFluency, FOOTER);
  sFluency.addNotes(NOTES_FLUENCY);

  // Slide 5: LI/SC
  liSlide(pres,
    ["We are learning to count groups of things up to 10"],
    [
      "I can touch and count each object one at a time",
      "I can say how many are in a group",
      "I can match a number to the right group",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 6: I Do — How to Count a Collection
  workedExSlide(pres, 2, "I Do", "How to Count a Collection",
    [
      "My counting rule:",
      "",
      "1. Touch each object",
      "2. Say the number",
      "3. Move it to the other side",
      "",
      "The LAST number I say tells me how many!",
    ],
    NOTES_IDO, FOOTER,
    (slide, lg) => {
      // Visual: counting objects
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.0, { strip: C.PRIMARY });
      slide.addText("Counting 6 Teddies", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.35,
        fontSize: 15, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });

      // Draw 6 circles representing teddies in a line
      const dotY = lg.panelTopPadded + 0.7;
      for (let i = 0; i < 6; i++) {
        const dotX = lg.rightX + 0.35 + i * 0.6;
        slide.addShape("roundRect", {
          x: dotX, y: dotY, w: 0.45, h: 0.45, rectRadius: 0.22,
          fill: { color: i < 4 ? C.SUCCESS : C.SECONDARY },
          line: { color: C.CHARCOAL, width: 0.5 },
        });
        slide.addText(String(i + 1), {
          x: dotX, y: dotY, w: 0.45, h: 0.45,
          fontSize: 16, fontFace: FONT_H, color: C.WHITE,
          align: "center", valign: "middle", bold: true, margin: 0,
        });
      }

      // "How many?" answer
      addTextOnShape(slide, "How many? 6", {
        x: lg.rightX + 0.5, y: lg.panelTopPadded + 1.45, w: lg.rightW - 1.0, h: 0.4, rectRadius: 0.08,
        fill: { color: C.ALERT },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });

      // Key reminder
      addTextOnShape(slide, "Last number = how many!", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 2.2, w: lg.rightW - 0.6, h: 0.4, rectRadius: 0.06,
        fill: { color: C.PRIMARY },
      }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // Slide 7-8: CFU 1 (withReveal)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["2"]);
      addStageBadge(s, 2, "I Do");
      addTitle(s, "How Many Cubes?", { y: 0.65, fontSize: 22, color: STAGE_COLORS["2"] });

      // Show 4 cubes in a line
      addCard(s, 0.5, CONTENT_TOP, 9, 2.0, { strip: C.PRIMARY });
      for (let i = 0; i < 4; i++) {
        const cubeX = 1.5 + i * 1.6;
        s.addShape("roundRect", {
          x: cubeX, y: CONTENT_TOP + 0.5, w: 1.0, h: 1.0, rectRadius: 0.1,
          fill: { color: C.SECONDARY },
          line: { color: C.PRIMARY, width: 2 },
        });
      }

      addCard(s, 0.5, CONTENT_TOP + 2.3, 9, 1.2, { strip: C.ACCENT });
      s.addText([
        { text: "Count the cubes.", options: { fontSize: 18, bold: true, color: C.PRIMARY, breakLine: true } },
        { text: "Show me with your FINGERS how many.", options: { fontSize: 16, color: C.CHARCOAL } },
      ], {
        x: 0.75, y: CONTENT_TOP + 2.42, w: 8.5, h: 1.0,
        fontFace: FONT_B, margin: 0, valign: "top",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "4 cubes!", {
        x: 3.0, y: 4.2, w: 4.0, h: 0.6, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 9-10: We Do (withReveal) — Count your collection
  withReveal(
    () => workedExSlide(pres, 3, "We Do", "Count Your Collection",
      [
        "At your table:",
        "",
        "1. Tip out your bag of objects",
        "2. Touch each one and say the number",
        "3. Move each object as you count",
        "4. Write how many on your whiteboard",
      ],
      NOTES_WEDO_Q, FOOTER,
      (slide, lg) => {
        addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.0, { strip: C.SECONDARY });
        slide.addText("Remember:", {
          x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
          fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
        });
        slide.addText([
          { text: "Touch it", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
          { text: "Say the number", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
          { text: "Move it", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
          { text: "The LAST number = how many!", options: { bullet: true, fontSize: 15, color: C.ALERT, bold: true } },
        ], {
          x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.45, w: lg.rightW - 0.5, h: 1.2,
          fontFace: FONT_B, margin: 0, valign: "top",
        });
      }
    ),
    (slide) => {
      addTextOnShape(slide, "Check your count with your partner!", {
        x: 0.5, y: 4.55, w: 9, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 11-12: CFU 2 Hinge (withReveal) — Which has more?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Which Group Has More?", { y: 0.65, fontSize: 22, color: STAGE_COLORS["3"] });

      // Two groups of dots
      addCard(s, 0.5, CONTENT_TOP, 4.3, 2.5, { strip: C.PRIMARY });
      s.addText("Group A", {
        x: 0.75, y: CONTENT_TOP + 0.08, w: 3.8, h: 0.3,
        fontSize: 16, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0, align: "center",
      });
      // 3 dots
      for (let i = 0; i < 3; i++) {
        s.addShape("roundRect", {
          x: 1.5 + i * 0.9, y: CONTENT_TOP + 0.6, w: 0.7, h: 0.7, rectRadius: 0.35,
          fill: { color: C.PRIMARY },
        });
      }

      addCard(s, 5.2, CONTENT_TOP, 4.3, 2.5, { strip: C.SECONDARY });
      s.addText("Group B", {
        x: 5.45, y: CONTENT_TOP + 0.08, w: 3.8, h: 0.3,
        fontSize: 16, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0, align: "center",
      });
      // 7 dots (two rows: 4 + 3)
      for (let i = 0; i < 4; i++) {
        s.addShape("roundRect", {
          x: 5.8 + i * 0.7, y: CONTENT_TOP + 0.5, w: 0.55, h: 0.55, rectRadius: 0.27,
          fill: { color: C.SECONDARY },
        });
      }
      for (let i = 0; i < 3; i++) {
        s.addShape("roundRect", {
          x: 6.15 + i * 0.7, y: CONTENT_TOP + 1.15, w: 0.55, h: 0.55, rectRadius: 0.27,
          fill: { color: C.SECONDARY },
        });
      }

      addCard(s, 0.5, CONTENT_TOP + 2.7, 9, 0.8, { strip: C.ACCENT });
      s.addText("Point to the group that has MORE.", {
        x: 0.75, y: CONTENT_TOP + 2.82, w: 8.5, h: 0.55,
        fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, bold: true, margin: 0, valign: "middle",
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU2);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Group B has more! (3 vs 7)", {
        x: 1.5, y: 4.2, w: 7.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 13: You Do
  workedExSlide(pres, 4, "You Do", "Count and Match",
    [
      "First: Count the objects in each group.",
      "Next: Write the number.",
      "Then: Draw a line to match the group to",
      "        the right number.",
      "",
      "Show your counting!",
      "You have 6 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.5, { strip: C.ALERT });
      slide.addText("Counting Rule:", {
        x: lg.rightX + 0.15, y: lg.panelTopPadded + 0.08, w: lg.rightW - 0.3, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "1. Touch it", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "2. Say the number", options: { fontSize: 14, color: C.CHARCOAL, breakLine: true } },
        { text: "3. Last number = how many", options: { fontSize: 14, color: C.CHARCOAL, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.42, w: lg.rightW - 0.5, h: 0.9,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Slide 14: Exit Ticket
  exitTicketSlide(pres,
    [
      "Count the stars. How many?  (4 stars shown)",
      "Count the hearts. How many?  (7 hearts shown)",
      "Draw a group of 5 circles.",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 15: Closing
  closingSlide(pres,
    "Tell your partner: What is the rule for counting a group of objects?",
    [
      "I can touch and count each object one at a time",
      "I can say how many are in a group",
      "I can match a number to the right group",
    ],
    NOTES_CLOSING);

  // Slide 16: Resources


  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "N20_Session1_Counting_Collections.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Generate PDFs ─────────────────────────────────────────────────────────

  // Worksheet
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Counting Collections to 10",
      color: C.NAVY,
      lessonInfo: "Session 1 of 5 | Foundation Numeracy",
    });
    y = addTipBox(doc, "Count each group carefully. Touch each object as you count. Write the number.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A: Count and Write (to 5)", y, { color: C.NAVY });

    // Draw simple dot collections with write lines
    const drawDots = (doc, count, y, num) => {
      const dotR = 8;
      const gap = 24;
      const startX = 100;
      doc.fontSize(12).fillColor("#333333").text(num + ".", 55, y + 4);
      for (let i = 0; i < count; i++) {
        doc.circle(startX + i * gap, y + 10, dotR).fill("#" + C.SECONDARY);
      }
      doc.fontSize(14).fillColor("#333333").text("How many? _____", 100 + count * gap + 20, y + 2);
      return y + 38;
    };

    y = drawDots(doc, 3, y, "1");
    y = drawDots(doc, 5, y, "2");
    y = drawDots(doc, 2, y, "3");
    y = drawDots(doc, 4, y, "4");

    y = addSectionHeading(doc, "Section B: Count and Write (to 10)", y, { color: C.NAVY });
    y = drawDots(doc, 7, y, "5");
    y = drawDots(doc, 6, y, "6");
    y = drawDots(doc, 9, y, "7");
    y = drawDots(doc, 8, y, "8");
    y = drawDots(doc, 10, y, "9");

    y = addSectionHeading(doc, "Section C: Draw and Count", y, { color: C.NAVY });
    y = addProblem(doc, 10, "Draw 6 circles. Write the number.", y, { color: C.NAVY, writeLines: [{ label: "How many?" }] });
    y = addProblem(doc, 11, "Draw 8 circles. Write the number.", y, { color: C.NAVY, writeLines: [{ label: "How many?" }] });

    addPdfFooter(doc, "Session 1 | Numbers to 20 | Foundation");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Counting Collections to 10",
      color: C.NAVY,
      lessonInfo: "Session 1 of 5 | Foundation Numeracy",
    });
    y = addSectionHeading(doc, "Section A", y, { color: C.NAVY });
    y = addBodyText(doc, "1. 3 dots -> 3", y);
    y = addBodyText(doc, "2. 5 dots -> 5", y);
    y = addBodyText(doc, "3. 2 dots -> 2", y);
    y = addBodyText(doc, "4. 4 dots -> 4", y);
    y = addSectionHeading(doc, "Section B", y, { color: C.NAVY });
    y = addBodyText(doc, "5. 7 dots -> 7", y);
    y = addBodyText(doc, "6. 6 dots -> 6", y);
    y = addBodyText(doc, "7. 9 dots -> 9", y);
    y = addBodyText(doc, "8. 8 dots -> 8", y);
    y = addBodyText(doc, "9. 10 dots -> 10", y);
    y = addSectionHeading(doc, "Section C", y, { color: C.NAVY });
    y = addBodyText(doc, "10. Accept any drawing of 6 circles. Answer: 6", y);
    y = addBodyText(doc, "11. Accept any drawing of 8 circles. Answer: 8", y);
    addPdfFooter(doc, "Session 1 | Answer Key | Foundation");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Enabling Scaffold — Counting mat with touch-point guides
  await (async () => {
    const doc = createPdf({ title: ENABLING_RES.name });
    let y = addPdfHeader(doc, ENABLING_RES.name, {
      subtitle: "Counting Mat - Touch and Count",
      color: C.TEAL,
      lessonInfo: "Session 1 of 5 | Foundation Numeracy",
    });
    y = addTipBox(doc, "Put one counter on each circle as you count. Say the number. The last number tells you how many!", y, { color: C.TEAL });

    // Draw counting mats with numbered circles for collections to 5
    const drawCountingMat = (doc, count, label, y) => {
      doc.fontSize(14).font("Sans-Bold").fillColor("#" + C.NAVY).text(label, 55, y);
      y += 22;
      const circleR = 18;
      const gap = 48;
      const startX = 90;
      for (let i = 0; i < count; i++) {
        doc.circle(startX + i * gap, y + circleR, circleR)
          .lineWidth(1.5).strokeColor("#" + C.NAVY).stroke();
        doc.fontSize(12).fillColor("#999999").text(String(i + 1), startX + i * gap - 5, y + circleR - 6);
      }
      doc.fontSize(14).font("Sans").fillColor("#333333")
        .text("How many? _____", startX + count * gap + 30, y + circleR - 8);
      return y + circleR * 2 + 20;
    };

    y = addSectionHeading(doc, "Count to 5 (with number guides)", y, { color: C.NAVY });
    y = drawCountingMat(doc, 3, "Collection 1:", y);
    y = drawCountingMat(doc, 5, "Collection 2:", y);
    y = drawCountingMat(doc, 2, "Collection 3:", y);
    y = drawCountingMat(doc, 4, "Collection 4:", y);

    y = addSectionHeading(doc, "Now try without number guides", y, { color: C.NAVY });
    const drawBlankMat = (doc, count, label, y) => {
      doc.fontSize(14).font("Sans-Bold").fillColor("#" + C.NAVY).text(label, 55, y);
      y += 22;
      const circleR = 18;
      const gap = 48;
      const startX = 90;
      for (let i = 0; i < count; i++) {
        doc.circle(startX + i * gap, y + circleR, circleR)
          .lineWidth(1.5).strokeColor("#" + C.NAVY).stroke();
      }
      doc.fontSize(14).font("Sans").fillColor("#333333")
        .text("How many? _____", startX + count * gap + 30, y + circleR - 8);
      return y + circleR * 2 + 20;
    };
    y = drawBlankMat(doc, 4, "Collection 5:", y);
    y = drawBlankMat(doc, 5, "Collection 6:", y);

    addPdfFooter(doc, "Session 1 | Enabling Scaffold | Foundation");
    await writePdf(doc, path.join(OUT_DIR, ENABLING_RES.fileName));
    console.log("PDF written: " + ENABLING_RES.fileName);
  })();

  // Extension — Counting Beyond 10 with Ten Frames
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Counting Beyond 10 with Ten Frames",
      color: C.TEAL,
      lessonInfo: "Session 1 of 5 | Foundation Numeracy",
    });

    y = addSectionHeading(doc, "What is a Ten Frame?", y, { color: C.NAVY });
    y = addBodyText(doc, "A ten frame is a special box with 10 spaces. When the frame is FULL, that means 10!", y);
    y = addBodyText(doc, "If there are more objects, they go OUTSIDE the frame. Count the frame (10) plus the extras.", y);

    y = addSectionHeading(doc, "Example: Counting 13", y, { color: C.NAVY });
    y = addBodyText(doc, "Step 1: Fill the ten frame (that is 10).", y);
    y = addBodyText(doc, "Step 2: Count the extras: 1, 2, 3.", y);
    y = addBodyText(doc, "Step 3: 10 and 3 more = 13.", y);

    // Draw a simple ten frame example
    const frameX = 80;
    const cellSize = 30;
    for (let row = 0; row < 2; row++) {
      for (let col = 0; col < 5; col++) {
        doc.rect(frameX + col * cellSize, y + row * cellSize, cellSize, cellSize)
          .lineWidth(1).strokeColor("#" + C.NAVY).stroke();
        // Fill all 10 cells
        doc.circle(frameX + col * cellSize + cellSize / 2, y + row * cellSize + cellSize / 2, 10)
          .fill("#" + C.SECONDARY);
      }
    }
    // Draw 3 extras outside
    for (let i = 0; i < 3; i++) {
      doc.circle(frameX + 5 * cellSize + 30 + i * 30, y + cellSize / 2, 10)
        .fill("#" + C.ALERT);
    }
    doc.fontSize(14).font("Sans-Bold").fillColor("#" + C.NAVY)
      .text("10 + 3 = 13", frameX + 5 * cellSize + 30, y + cellSize + 10);
    y += cellSize * 2 + 40;

    y = addSectionHeading(doc, "Your Investigation", y, { color: C.NAVY });
    y = addBodyText(doc, "Count each collection. Use the ten frame to help you.", y);

    // Draw blank ten frames with dot collections
    const drawTenFrameTask = (doc, total, taskNum, y) => {
      doc.fontSize(12).font("Sans-Bold").fillColor("#333333").text(taskNum + ".", 55, y + 4);
      // Draw dots for the collection
      const dotR = 6;
      const gap = 18;
      const startX = 80;
      for (let i = 0; i < total; i++) {
        const row = Math.floor(i / 6);
        const col = i % 6;
        doc.circle(startX + col * gap, y + 10 + row * gap, dotR).fill("#" + C.SECONDARY);
      }
      // Blank ten frame
      const tfX = 80 + 7 * gap;
      for (let row = 0; row < 2; row++) {
        for (let col = 0; col < 5; col++) {
          doc.rect(tfX + col * 24, y + row * 24, 24, 24)
            .lineWidth(0.8).strokeColor("#" + C.NAVY).stroke();
        }
      }
      doc.fontSize(14).font("Sans").fillColor("#333333")
        .text("How many? _____", tfX + 5 * 24 + 15, y + 14);
      return y + 56;
    };

    y = drawTenFrameTask(doc, 11, "1", y);
    y = drawTenFrameTask(doc, 13, "2", y);
    y = drawTenFrameTask(doc, 15, "3", y);
    y = drawTenFrameTask(doc, 12, "4", y);

    y = addSectionHeading(doc, "Did You Know?", y, { color: C.NAVY });
    y = addBodyText(doc, "When a number is bigger than 10, it is called a TEEN number! Teen numbers are made of 10 plus some more. We will learn more about teen numbers in our next lesson.", y);

    addPdfFooter(doc, "Session 1 | Extension | Foundation");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();

  console.log("Session 1 build complete.");
}

build().catch(err => { console.error(err); process.exit(1); });
