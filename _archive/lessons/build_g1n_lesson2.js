"use strict";

// Grade 1 Numeracy — Lesson 2: Ordering Numbers on a Number Line
// AC9M1N01 — recognise, represent and order numbers to at least 120 using
//            physical and virtual materials, numerals, number lines and charts.
// Daily Review: Counting and place value (reading numerals - retrieve Lesson 1).
// Fluency:      Skip counting, addition and patterns (one more / one less to 30).

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade1", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, cfuSlide, closingSlide,
  workedExSlide, exitTicketSlide,
  dailyReviewSlide, fluencySlide,
  addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  addInstructionCard,
  withReveal,
  addTensFrame, addNumberTrack, addNumberLine,
  STAGE_COLORS,
} = T;

const SESSION = 2;
const TOTAL = 10;
const UNIT_TITLE = "Numbers to 120";
const LESSON_TITLE = "Ordering Numbers on a Number Line";
const FOOTER = `${UNIT_TITLE} | Lesson ${SESSION} of ${TOTAL} | Year 1 Numeracy`;
const OUT_DIR = `output/G1N_Lesson${SESSION}_Ordering_On_Number_Line`;
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 2 Order on the Number Line",
  "Cut, sort and paste numeral cards onto the number line.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 2 Answer Key",
  "Teacher reference with the correct number-line order.");
const EXTENDING_RES = makeSessionResource(SESSION, "Lesson 2 Extension",
  "Order tricky number sets including teens and twenties together.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

function drawNumeralCard(slide, x, y, w, h, numeral, opts) {
  const o = opts || {};
  const fill = o.fill || C.WHITE;
  const stroke = o.stroke || C.PRIMARY;
  const textColor = o.color || C.PRIMARY;
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: 0.14,
    fill: { color: fill },
    line: { color: stroke, width: 2.5 },
  });
  slide.addText(String(numeral), {
    x, y, w, h,
    fontSize: o.fontSize || 60, fontFace: FONT_H, color: textColor,
    bold: true, align: "center", valign: "middle", margin: 0,
  });
}

// ─── Teacher Notes ──────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Yesterday we read and wrote numbers to 30
- Today we put numbers in order on a number line
- Like finding where each number lives

DO:
- Display the title slide as students arrive
- Have numeral cards 0-50 ready
- Have a long string and pegs ready for the bigger number line build

TEACHER NOTES:
Lesson 2 of 10. The big skill - which number comes before, which comes after, and where a missing number sits.

WATCH FOR:
- Students who can read numbers but cannot order them - common
- Students who count by 1s every time - help them see "one more / one less"

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are our materials for today
- We will use real number cards and a long line on the floor or wall

DO:
- Point to each resource
- Show students the long string number line at the front of the room

TEACHER NOTES:
The cut-and-paste worksheet is the core student resource. The Extension stretches into mixed teens/twenties for early finishers.

WATCH FOR:
- Missing scissors or glue - prepare before lesson

[General: Resources | VTLM 2.0: Planning]`;

const NOTES_LAUNCH = `SAY:
- I have three numeral cards - 7, 5 and 9
- Which one is the smallest? Which one is the biggest?
- Tell your partner the order from smallest to biggest

DO:
- Show three cards 7, 5, 9 on the slide
- Allow 30 seconds partner talk
- Cold call one pair to share

TEACHER NOTES:
Activates ordering with small numbers. The "before / between / after" language enters in I Do.

WATCH FOR:
- Students who say "5, 9, 7" - they are reading left to right, not ordering by size; reteach
- Students who use fingers to compare - good strategy

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_DR1_Q = `SAY:
- Daily Review - read this numeral together when I point
- Eyes up

DO:
- Display 19 in big numerals
- Point firmly
- Listen for "nineteen"

TEACHER NOTES:
Daily Review 1 of 3. Retrieves reading skill from Lesson 1.

WATCH FOR:
- Students who say "ninety" - they hear "nine" and add "ty"; reread digit by digit

[Stage 1: Daily Review 1 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR1_A = `SAY:
- The number is nineteen
- 1 ten and 9 ones is nineteen

DO:
- Reveal "nineteen"
- Tap the digits as you read

TEACHER NOTES:
Quick retrieval. Keep pace brisk.

WATCH FOR:
- Confident chorus - move on

[Stage 1: Daily Review 1 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_Q = `SAY:
- Next one. Read this numeral together

DO:
- Display 24 in big numerals
- Point and cue choral read

TEACHER NOTES:
Mid-twenties numeral retrieval.

WATCH FOR:
- Students who say "twenty four" cleanly - they have decade transitions

[Stage 1: Daily Review 2 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_A = `SAY:
- Twenty-four. 2 tens and 4 ones

DO:
- Reveal "twenty-four"

TEACHER NOTES:
Brief reveal. Keep tempo up.

WATCH FOR:
- Whole-class chorus = ready to move on

[Stage 1: Daily Review 2 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_Q = `SAY:
- Last one. Which number is bigger - 12 or 21?
- Show me on your fingers - 1 for the first, 2 for the second

DO:
- Display 12 and 21 side by side
- Wait 8 seconds
- Signal "show me"

TEACHER NOTES:
This previews today's lesson. Position matters - 12 is one ten, 21 is two tens.

WATCH FOR:
- Students who point at 12 - they may be looking at the bigger digit "2" not the position
- Students who confidently choose 21 - they have place-value sense

[Stage 1: Daily Review 3 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_A = `SAY:
- 21 is bigger. 21 has 2 tens. 12 has only 1 ten
- The tens digit tells us how big the number is

DO:
- Reveal "21 is bigger"
- Tap the 2 in 21, then the 1 in 12

TEACHER NOTES:
Naming "tens digit tells us how big" plants language for the rest of the lesson.

WATCH FOR:
- Students who say "they are the same numbers" - they will need explicit support today

[Stage 1: Daily Review 3 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_FL1 = `SAY:
- Fluency. One more, one more, one more
- I say a number. You say the one MORE
- Ready - 7

DO:
- Display "+1 patterns" with examples
- Say 7. Class says 8. Say 12. Class says 13. Say 19. Class says 20. Say 24. Class says 25

TEACHER NOTES:
"One more" fluency anchors the number-line concept. Crossing the decade (19 to 20, 29 to 30) is the tricky bit.

WATCH FOR:
- Students who pause at 19+1 - mark for support in the You Do

[Stage 1: Fluency 1 | VTLM 2.0: Automaticity]`;

const NOTES_FL2 = `SAY:
- Now one LESS
- I say a number. You say one less
- Ready - 10

DO:
- Display "-1 patterns"
- Say 10. Class says 9. Say 15. Class says 14. Say 20. Class says 19. Say 25. Class says 24

TEACHER NOTES:
"One less" is harder than "one more". Stop at 30. Slow down across decades.

WATCH FOR:
- Students who say "one" instead of "nine" for 10-1 - they have not yet automated the count-back; small-group support

[Stage 1: Fluency 2 | VTLM 2.0: Automaticity]`;

const NOTES_VOCAB = `SAY:
- A new word - between
- Between means "in the middle of two things"
- 6 is between 5 and 7

DO:
- Display 5, 6, 7 with arrows pointing to the 6
- Repeat "between" with the hand gesture

TEACHER NOTES:
"Between" is essential for the missing-number tasks today.

WATCH FOR:
- Students who only learn the word - reinforce with body action; sit between two students to model

[General: Vocabulary | VTLM 2.0: Vocabulary Build]`;

const NOTES_LI_SC = `SAY:
- Read with me - We are learning to put numbers in order from smallest to biggest
- I can point to a number on a number line
- I can put 3 numerals in order from smallest to biggest
- I can find a missing number between two others

DO:
- Choral read LI then each SC

TEACHER NOTES:
SC1 is reachable for everyone. SC2 is core. SC3 stretches to "find the missing one".

WATCH FOR:
- Students who shake their head at SC3 - reassure, that is the stretch

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO_LINE = `SAY:
- Watch how I put numbers on a number line
- A number line goes from small to big, left to right
- The smallest number is on the left. The biggest is on the right
- I peg 5 here, 7 here, and 12 here - in order

DO:
- Use the smartboard to drag numeral cards onto a 0-20 line
- Think aloud - "5 first because it is smallest. Then 7. 12 is bigger so it goes near the end"

TEACHER NOTES:
Modelling the "smallest on the left" routine. Mention that numbers DO not jump around - they go in order.

WATCH FOR:
- Students mimicking the routine - good engagement

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_MISSING = `SAY:
- Now watch - this number line has a number missing
- 10, 11, blank, 13
- What is missing? I look at 11 and add one - that gives me 12
- I check - 12 comes before 13. Yes!

DO:
- Display 10, 11, ?, 13 on a number line
- Tap each number as you count up
- Write 12 in the blank

TEACHER NOTES:
Strategy - count up from the number before, or count back from the number after.

WATCH FOR:
- Students who guess instead of counting - prompt them to count

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU = `SAY:
- Quick check. Look at the screen
- The number line shows 14, blank, 16
- What is the missing number? Write it on your whiteboard

DO:
- Display the line
- 8 seconds thinking time
- Signal "show me"

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Write the missing number on your whiteboard. Show me on three - 1, 2, 3"
- Scan for: 15
PROCEED: If 80%+ show 15, move to We Do.
PIVOT: Most likely misconception - students write 14 again, or 16 again. Reteach by pointing to 14, then a finger jump up - 15. Re-check with 19, ?, 21.

TEACHER NOTES:
Whiteboards give visible response from every student.

WATCH FOR:
- Students who write 13 or 17 - they may be looking at the wrong gap

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_CFU_A = `SAY:
- The answer is 15. 14, 15, 16 in order
- Tick or fix

DO:
- Reveal 15

TEACHER NOTES:
Brief reveal.

WATCH FOR:
- Visible thumbs up = ready for We Do

[Stage 2: CFU Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me
- I have three cards - 23, 17, 9
- On your whiteboard, write them in order - smallest first

DO:
- Display the three numerals on the slide
- Allow 60 seconds
- Cold call one student to read their order

TEACHER NOTES:
This We Do uses ordering rather than missing-number. Students need both strategies in their kit.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Order numbers within a single decade only - 13, 17, 11. Use the number track so students can see position.
- Extra Notes: Sit with small group at front. Show the 0-20 number track.
EXTENDING PROMPT:
- Task: Once ordered, ask what number is one more than the biggest, and one less than the smallest.
- Extra Notes: Whole-class extension; can be done from seats.

WATCH FOR:
- Students writing left to right (23, 17, 9) without sorting - prompt to start from smallest

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- The order from smallest to biggest is 9, 17, 23
- 9 is the smallest because it is a one-digit number
- 17 is in the teens. 23 is in the twenties - the biggest

DO:
- Reveal the order
- Tap each number as you say it

TEACHER NOTES:
Naming "one-digit, teens, twenties" plants the place-value language.

WATCH FOR:
- Tick-and-fix in workbooks

[Stage 3: We Do Answer | VTLM 2.0: Scaffold Practice]`;

const NOTES_HINGE_Q = `SAY:
- Hinge check. The number line is 26, blank, 28
- Write the missing number on your whiteboard
- Show me when I signal

DO:
- Display the line
- 8 seconds
- Signal show me

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Missing number on the line - write it. Show me"
- Scan for: 27
PROCEED: If 80%+ show 27, move to You Do.
PIVOT: Most likely misconception - students get the digit but write 72. Or students struggle in the twenties decade. Reteach with 26 on a tens-frame model + the count-up routine. Re-check with 28, ?, 30.

TEACHER NOTES:
This hinge tests transfer from teens to twenties. The pivot reteaches with manipulatives, not more words.

WATCH FOR:
- Students who write 27 confidently - they have the count-on strategy

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_HINGE_A = `SAY:
- The answer is 27. 26, 27, 28
- Tick or fix

DO:
- Reveal 27
- Quick count 26-27-28 together

TEACHER NOTES:
Final hinge before independent work.

WATCH FOR:
- Confident class = move on

[Stage 3: CFU Hinge Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Independent time
- Take the worksheet and your scissors
- First - cut out the numeral cards
- Next - put them in order on the number line
- Then - glue them down when you and your partner agree

DO:
- Distribute the worksheet
- Circulate to watch for left-to-right ordering errors
- For finishers, hand out the Extension PDF

TEACHER NOTES:
Cut-and-paste is high-engagement and forces ordering before commitment. The worksheet has two number lines - 0-20 and 10-30.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Work with the teacher at the back table. Cut and arrange the smaller 0-15 set first. Lay them out before gluing.
- Extra Notes: Small group with teacher support. No separate enabling PDF.
EXTENDING PROMPT:
- Task: Extension sheet - order 5 numerals from different decades (e.g. 8, 14, 19, 23, 28). Then find the missing number between each pair.
- Extra Notes: Distribute Lesson 2 Extension PDF.

WATCH FOR:
- Students who glue before ordering - prompt them to slide cards into place first
- Students who finish quickly - hand out Extension

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket
- I have three numbers - 18, 11, 25
- Write them in order from smallest to biggest

DO:
- Display the three numerals
- Students write on their slip
- Collect

TEACHER NOTES:
Exit ticket assesses SC2. Sort slips into three piles - correct (11, 18, 25), partial (mostly right), incorrect.

WATCH FOR:
- Slips with the order reversed - students may have written biggest first; reteach tomorrow
- Slips with correct order - they have SC2

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Lets check our success criteria
- I can point to a number on a number line - thumbs
- I can put 3 numerals in order - thumbs
- I can find a missing number - thumbs
- Turn and tell - what is the trick for finding the missing number?

DO:
- Run thumbs check
- Partner talk for 30 seconds
- Cold call one student

TEACHER NOTES:
The reflection prompts students to explain the count-on strategy.

WATCH FOR:
- Students who can articulate the strategy - they own the skill

[General: Closing | VTLM 2.0: Review and Reflect]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, LESSON_TITLE, "Lesson 2 of 10 - Numbers to 120", "Year 1 Numeracy | AC9M1N01", NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 3: Launch - three numeral cards 7, 5, 9
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Launch");
    addTitle(s, "Smallest to biggest?", { color: STAGE_COLORS["1"] });

    drawNumeralCard(s, 1.5, 1.9, 2.0, 1.8, "7", { fontSize: 110 });
    drawNumeralCard(s, 4.0, 1.9, 2.0, 1.8, "5", { fontSize: 110 });
    drawNumeralCard(s, 6.5, 1.9, 2.0, 1.8, "9", { fontSize: 110 });

    addInstructionCard(s, [
      { role: "header", text: "Tell your partner" },
      { role: "body", text: "Which is smallest? Which is biggest?" },
    ], { x: 0.5, y: 4.0, w: 9, h: 1.0, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // Slide 4-5: DR1 — Read 19
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Read this numeral", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 3.5, 1.7, 3.0, 2.5, "19", { fontSize: 160, fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "nineteen", {
        x: 2.5, y: 4.55, w: 5.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR1_A);
    }
  );

  // Slide 6-7: DR2 — Read 24
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Read this numeral", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 3.5, 1.7, 3.0, 2.5, "24", { fontSize: 160, fill: C.SECONDARY, color: C.WHITE, stroke: C.SECONDARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "twenty-four", {
        x: 2.5, y: 4.55, w: 5.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR2_A);
    }
  );

  // Slide 8-9: DR3 — Which is bigger 12 or 21
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Which is bigger?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 2.0, 1.9, 2.5, 2.0, "12", { fontSize: 130 });
      drawNumeralCard(s, 5.5, 1.9, 2.5, 2.0, "21", { fontSize: 130 });

      addInstructionCard(s, [
        { role: "header", text: "Show me 1 or 2" },
        { role: "body", text: "1 for the first card. 2 for the second." },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR3_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "21 is bigger - 2 tens!", {
        x: 2.0, y: 4.55, w: 6.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR3_A);
    }
  );

  // Slide 10: Fluency 1 — one more
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "One MORE", { color: STAGE_COLORS["1"] });

    // 4 example boxes
    const examples = [
      { from: "7", to: "8" },
      { from: "12", to: "13" },
      { from: "19", to: "20" },
      { from: "24", to: "25" },
    ];
    examples.forEach((ex, i) => {
      const x = 0.6 + (i % 2) * 4.5;
      const y = 1.8 + Math.floor(i / 2) * 1.3;
      drawNumeralCard(s, x, y, 1.4, 1.1, ex.from, { fontSize: 48, fill: C.WHITE, color: C.PRIMARY, stroke: C.PRIMARY });
      s.addText("+1", { x: x + 1.5, y: y, w: 0.6, h: 1.1, fontSize: 22, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", valign: "middle", margin: 0 });
      drawNumeralCard(s, x + 2.2, y, 1.4, 1.1, ex.to, { fontSize: 48, fill: C.SUCCESS, color: C.WHITE, stroke: C.SUCCESS });
    });

    addInstructionCard(s, [
      { role: "header", text: "I say a number. You say one MORE." },
    ], { x: 0.5, y: 4.5, w: 9, h: 0.6, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FL1);
  })();

  // Slide 11: Fluency 2 — one less
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "One LESS", { color: STAGE_COLORS["1"] });

    const examples = [
      { from: "10", to: "9" },
      { from: "15", to: "14" },
      { from: "20", to: "19" },
      { from: "25", to: "24" },
    ];
    examples.forEach((ex, i) => {
      const x = 0.6 + (i % 2) * 4.5;
      const y = 1.8 + Math.floor(i / 2) * 1.3;
      drawNumeralCard(s, x, y, 1.4, 1.1, ex.from, { fontSize: 48, fill: C.WHITE, color: C.PRIMARY, stroke: C.PRIMARY });
      s.addText("-1", { x: x + 1.5, y: y, w: 0.6, h: 1.1, fontSize: 22, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", valign: "middle", margin: 0 });
      drawNumeralCard(s, x + 2.2, y, 1.4, 1.1, ex.to, { fontSize: 48, fill: C.SECONDARY, color: C.WHITE, stroke: C.SECONDARY });
    });

    addInstructionCard(s, [
      { role: "header", text: "I say a number. You say one LESS." },
    ], { x: 0.5, y: 4.5, w: 9, h: 0.6, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FL2);
  })();

  // Slide 12: Vocabulary — between
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Word", { color: C.PRIMARY });
    addTitle(s, "Between", { color: C.PRIMARY });

    drawNumeralCard(s, 2.5, 1.9, 1.4, 1.4, "5", { fontSize: 72, fill: C.WHITE, color: C.PRIMARY, stroke: C.PRIMARY });
    drawNumeralCard(s, 4.3, 1.9, 1.4, 1.4, "6", { fontSize: 72, fill: C.SUCCESS, color: C.WHITE, stroke: C.SUCCESS });
    drawNumeralCard(s, 6.1, 1.9, 1.4, 1.4, "7", { fontSize: 72, fill: C.WHITE, color: C.PRIMARY, stroke: C.PRIMARY });
    // Arrows pointing to 6
    s.addText("between", { x: 3.5, y: 3.4, w: 3.0, h: 0.4, fontSize: 18, fontFace: FONT_B, color: C.SUCCESS, bold: true, align: "center", margin: 0 });

    addInstructionCard(s, [
      { role: "header", text: "Between means 'in the middle of two things'." },
      { role: "body", text: "6 is between 5 and 7." },
    ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // Slide 13: LI/SC
  liSlide(pres,
    ["We are learning to put numbers in order from smallest to biggest."],
    [
      "I can point to a number on a number line.",
      "I can put 3 numerals in order from smallest to biggest.",
      "I can find a missing number between two others.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 14: I Do — Put numbers on a number line
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "I put numbers in order", { color: STAGE_COLORS["2"] });

    // Number track 0-20 with 5, 7, 12 highlighted
    addNumberTrack(s, 0.5, 2.0, 9.0, 0, 20, [5, 7, 12], { cellH: 0.7, fontSize: 18 });

    // Labels showing the order
    s.addText("smallest", { x: 0.5, y: 2.85, w: 4, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, align: "left", margin: 0 });
    s.addText("biggest", { x: 5.5, y: 2.85, w: 4, h: 0.3, fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, align: "right", margin: 0 });

    addInstructionCard(s, [
      { role: "header", text: "Smallest on the left, biggest on the right" },
      { role: "body", text: "5, 7, 12 - in order from left to right." },
    ], { x: 0.5, y: 3.5, w: 9, h: 1.55, strip: C.PRIMARY });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_LINE);
  })();

  // Slide 15: I Do — Missing number
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["2"]);
    addStageBadge(s, 2, "I Do");
    addTitle(s, "Find the missing number", { color: STAGE_COLORS["2"] });

    // 4 cells: 10, 11, ?, 13
    drawNumeralCard(s, 1.5, 1.9, 1.6, 1.6, "10", { fontSize: 80 });
    drawNumeralCard(s, 3.3, 1.9, 1.6, 1.6, "11", { fontSize: 80 });
    drawNumeralCard(s, 5.1, 1.9, 1.6, 1.6, "?", { fontSize: 80, color: C.ALERT, stroke: C.ALERT });
    drawNumeralCard(s, 6.9, 1.9, 1.6, 1.6, "13", { fontSize: 80 });

    addInstructionCard(s, [
      { role: "header", text: "Count up from the number before" },
      { role: "body", text: "11 ... then one more is ... 12!" },
    ], { x: 0.5, y: 3.75, w: 9, h: 1.3, strip: C.PRIMARY });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_MISSING);
  })();

  // Slide 16-17: CFU — Missing number 14, ?, 16
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Missing number?", { color: C.ALERT });

      drawNumeralCard(s, 2.0, 1.9, 2.0, 1.8, "14", { fontSize: 90 });
      drawNumeralCard(s, 4.3, 1.9, 2.0, 1.8, "?", { fontSize: 90, color: C.ALERT, stroke: C.ALERT });
      drawNumeralCard(s, 6.6, 1.9, 2.0, 1.8, "16", { fontSize: 90 });

      addInstructionCard(s, [
        { role: "header", text: "On your whiteboard" },
        { role: "body", text: "Write the missing number. Show me!" },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "15", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_CFU_A);
    }
  );

  // Slide 18-19: We Do — Order 23, 17, 9
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Put these in order", { color: STAGE_COLORS["3"] });

      drawNumeralCard(s, 1.5, 1.9, 2.0, 1.8, "23", { fontSize: 90 });
      drawNumeralCard(s, 4.0, 1.9, 2.0, 1.8, "17", { fontSize: 90 });
      drawNumeralCard(s, 6.5, 1.9, 2.0, 1.8, "9", { fontSize: 90 });

      addInstructionCard(s, [
        { role: "header", text: "On your whiteboard" },
        { role: "body", text: "Write them smallest to biggest." },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.SECONDARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "9 - 17 - 23", {
        x: 2.0, y: 4.55, w: 6.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 20-21: Hinge CFU — Missing in twenties (26, ?, 28)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU - Hinge", { color: C.ALERT });
      addTitle(s, "Missing number?", { color: C.ALERT });

      drawNumeralCard(s, 2.0, 1.9, 2.0, 1.8, "26", { fontSize: 90 });
      drawNumeralCard(s, 4.3, 1.9, 2.0, 1.8, "?", { fontSize: 90, color: C.ALERT, stroke: C.ALERT });
      drawNumeralCard(s, 6.6, 1.9, 2.0, 1.8, "28", { fontSize: 90 });

      addInstructionCard(s, [
        { role: "header", text: "Write the missing number" },
        { role: "body", text: "Show me on your whiteboard!" },
      ], { x: 0.5, y: 4.05, w: 9, h: 1.0, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_HINGE_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "27", {
        x: 4.0, y: 4.55, w: 2.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_HINGE_A);
    }
  );

  // Slide 22: You Do
  workedExSlide(pres, 4, "You Do", "Cut, sort and paste",
    [
      "First - cut the numeral cards.",
      "Next - put them in order, smallest first.",
      "Then - glue them onto the number line.",
      "",
      "Show your thinking!",
      "You have 10 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.ALERT });
      slide.addText("Remember", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.35,
        fontSize: 18, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Smallest on the left.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Slide cards before gluing.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Check with your partner!", options: { bullet: true, fontSize: 15, color: C.ALERT, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.7,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Slide 23: Exit Ticket
  exitTicketSlide(pres,
    ["Order from smallest to biggest:  18,  11,  25"],
    NOTES_EXIT, FOOTER);

  // Slide 24: Closing
  closingSlide(pres, {
    reflectionPrompt: "Turn and tell: what is the trick for finding the missing number on a number line?",
    scItems: [
      "I can point to a number on a number line.",
      "I can put 3 numerals in order from smallest to biggest.",
      "I can find a missing number between two others.",
    ],
    selfAssessment: "Thumbs up, sideways or down for each one",
  }, NOTES_CLOSING);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "G1N_Lesson2_Ordering_On_Number_Line.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── PDFs ──────────────────────────────────────────────────────────────────

  function hex(c) { return c.startsWith("#") ? c : "#" + c; }

  // Worksheet — Cut and paste cards onto number lines
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Cut, sort and paste",
      color: C.NAVY,
      lessonInfo: "Lesson 2 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Cut out the numeral cards at the bottom. Put them in order on the number line. Glue them down when you and your partner agree.", y, { color: C.TEAL });

    // Empty number line 1
    y = addSectionHeading(doc, "Number Line 1: 0 to 20", y, { color: C.NAVY });
    function drawEmptyLine(doc, y, start, end, marked) {
      const x = 60;
      const lineW = 475;
      const count = end - start + 1;
      const cellW = lineW / count;
      // Lighter line
      doc.save();
      for (let i = 0; i < count; i += 1) {
        const cx = x + i * cellW;
        const value = start + i;
        const show = marked.indexOf(value) !== -1;
        doc.rect(cx, y, cellW, 38).lineWidth(1).strokeColor(hex(C.NAVY)).stroke();
        if (show) {
          doc.fontSize(14).font("Sans-Bold").fillColor(hex(C.NAVY));
          doc.text(String(value), cx, y + 12, { width: cellW, align: "center" });
        }
      }
      doc.restore();
      return y + 50;
    }
    y = drawEmptyLine(doc, y, 0, 20, [0, 5, 10, 15, 20]);
    y = addBodyText(doc, "Cut cards: 3, 8, 12, 17", y);

    // Empty number line 2
    y = addSectionHeading(doc, "Number Line 2: 10 to 30", y, { color: C.NAVY });
    y = drawEmptyLine(doc, y, 10, 30, [10, 15, 20, 25, 30]);
    y = addBodyText(doc, "Cut cards: 13, 18, 22, 27", y);

    // Cut-out cards section
    y = addSectionHeading(doc, "Cut these out", y, { color: C.NAVY });
    const cuts = ["3", "8", "12", "17", "13", "18", "22", "27"];
    const startX = 60;
    const cardW = 60, cardH = 60, gap = 8;
    const rowsAvailable = (740 - y) / (cardH + gap);
    cuts.forEach((num, i) => {
      const col = i % 8;
      const row = Math.floor(i / 8);
      const cx = startX + col * (cardW + gap);
      const cy = y + row * (cardH + gap);
      doc.rect(cx, cy, cardW, cardH).dash(3, { space: 3 }).lineWidth(1).strokeColor("#777").stroke();
      doc.undash();
      doc.fontSize(28).font("Sans-Bold").fillColor(hex(C.NAVY));
      doc.text(num, cx, cy + 16, { width: cardW, align: "center" });
    });

    addPdfFooter(doc, "Lesson 2 | Numbers to 120 | Year 1");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Order on the number line - ANSWERS",
      color: C.NAVY,
      lessonInfo: "Lesson 2 of 10 | Year 1 Numeracy",
    });
    y = addSectionHeading(doc, "Number Line 1: 0 to 20", y, { color: C.NAVY });
    y = addBodyText(doc, "Cards 3, 8, 12, 17 placed correctly on positions 3, 8, 12, 17.", y);
    y = addSectionHeading(doc, "Number Line 2: 10 to 30", y, { color: C.NAVY });
    y = addBodyText(doc, "Cards 13, 18, 22, 27 placed correctly on positions 13, 18, 22, 27.", y);
    y = addBodyText(doc, "Watch for: students placing cards by guess rather than counting from the labelled marks. Reteach by pointing to each labelled mark and counting up.", y);
    addPdfFooter(doc, "Lesson 2 | Answer Key | Year 1");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Extension - mixed-decade ordering and missing numbers
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Order mixed numbers and find what is missing",
      color: C.TEAL,
      lessonInfo: "Lesson 2 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Order each set from smallest to biggest. Then find the missing number.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Order each set", y, { color: C.NAVY });
    y = addBodyText(doc, "1.  8 , 14 , 19 , 23 , 28   ->   _____ , _____ , _____ , _____ , _____", y);
    y = addBodyText(doc, "2.  21 , 9 , 16 , 30 , 25   ->   _____ , _____ , _____ , _____ , _____", y);
    y = addBodyText(doc, "3.  19 , 11 , 27 , 4 , 22   ->   _____ , _____ , _____ , _____ , _____", y);

    y = addSectionHeading(doc, "Find the missing number", y, { color: C.NAVY });
    y = addBodyText(doc, "4.   28 , _____ , 30", y);
    y = addBodyText(doc, "5.   19 , _____ , 21", y);
    y = addBodyText(doc, "6.   9 , _____ , 11", y);

    addPdfFooter(doc, "Lesson 2 | Extension | Year 1");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
