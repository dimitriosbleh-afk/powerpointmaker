"use strict";

// Grade 1 Numeracy — Lesson 1: Reading and Writing Numerals to 30
// AC9M1N01 — recognise, represent and order numbers to at least 120 using
//            physical and virtual materials, numerals, number lines and charts.
// Daily Review: Counting and place value (tens-frame quantity reads).
// Fluency:      Skip counting, addition and patterns (count 1-30, what comes next).

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addProblem, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

// Variant 0 fixed across all 10 lessons of this unit (CLAUDE.md cohesion rule).
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
  addTensFrame, addFiveFrame, addNumberTrack, addDotCard,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const TOTAL = 10;
const UNIT_TITLE = "Numbers to 120";
const LESSON_TITLE = "Reading & Writing Numerals to 30";
const FOOTER = `${UNIT_TITLE} | Lesson ${SESSION} of ${TOTAL} | Year 1 Numeracy`;
const OUT_DIR = `output/G1N_Lesson${SESSION}_Reading_Writing_Numerals`;
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Lesson 1 Read and Write Numerals",
  "Match numerals to ten-frame pictures, then write the numeral.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Lesson 1 Answer Key",
  "Teacher reference with answers for the read-and-write practice sheet.");
const EXTENDING_RES = makeSessionResource(SESSION, "Lesson 1 Extension",
  "Number-detective task: tricky pairs like 16/60 and 13/30.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Helpers ────────────────────────────────────────────────────────────────

// Draw a numeral card (a coloured roundRect with a big bold numeral inside).
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
- Today we start our unit on Numbers to 120
- We will count, read and write numbers in lots of fun ways
- Today's job - reading and writing numbers all the way to 30

DO:
- Display the title slide as students arrive on the mat
- Have numeral cards 0-30, mini-whiteboards and ten frames ready
- Settle the class with a quick choral count to 10

TEACHER NOTES:
Lesson 1 of 10. This lesson revisits familiar numbers (1-10) and stretches into the teens and twenties. The focus is on numeral recognition and writing, not counting fluency. Confusable pairs like 16/60 sit later in the lesson.

WATCH FOR:
- Students who reverse digits (writing 31 for 13) - normal at this age, plan support
- Students who already read to 30 confidently - they become helpers in We Do

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are the materials we will use today
- Keep the worksheet and answer key handy for after We Do

DO:
- Point to each resource on the slide
- Make sure the materials are stacked at the front of the room before the lesson begins

TEACHER NOTES:
Print the worksheet single-sided. The Extension is for students who finish quickly. Ten frames and numeral cards live on the maths trolley.

WATCH FOR:
- Resource issues - flag before lesson begins, not during

[General: Resources | VTLM 2.0: Planning]`;

const NOTES_LAUNCH = `SAY:
- I am thinking of a number. It has a 2 and a 5
- What number could it be? Tell your partner
- Some of you may remember reading numbers like this last term
- If this feels new, that is okay - we will build it together

DO:
- Allow 30 seconds of partner talk
- Cold call 2 partners to share
- Show "25" written on the board so students can read it aloud

TEACHER NOTES:
A gentle launch that activates digit-reading from Foundation. The "2 and a 5" frame previews place value without naming it. Beginner-safe language honours mixed readiness.

WATCH FOR:
- Students who say "52" or "twenty-five" - both are reasonable; clarify together
- Quiet students - bring them in by pointing to the digits and reading slowly

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_DR1_Q = `SAY:
- Daily Review time - this is to warm up our maths brain
- Look at the ten frame on the screen
- How many counters do you see? Think first, then show me on your fingers

DO:
- Display the ten frame with 7 counters
- Hold up your hand to signal "think"
- Then signal "show me" and scan for 7 fingers

TEACHER NOTES:
Daily Review 1 of 3. The 7-counter ten frame is read as "5 and 2 more" by students who subitise, or one-by-one by counters. Both are fine.

WATCH FOR:
- Students showing 5 or 8 - they are counting only one row or miscounting
- Students who pause and look unsure - prompt with "Start at the top row"

[Stage 1: Daily Review 1 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR1_A = `SAY:
- Show me your answer on the count of 3 - 1, 2, 3
- The answer is 7. The top row is 5 full, the bottom row is 2 more
- Tick if you got it. Fix it if you didn't - fixing is the learning

DO:
- Reveal the answer 7
- Students tick or fix on whiteboards

TEACHER NOTES:
The "5 and 2 more" link previews part-part-whole work in Lessons 5 and 6.

WATCH FOR:
- Students who answer instantly - they are subitising the 5-row and counting the 2
- Students who recount - normal, will speed up with practice

[Stage 1: Daily Review 1 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_Q = `SAY:
- Next one. Look at the ten frame
- How many counters? Show me on your fingers when I signal

DO:
- Display the ten frame fully filled (10 counters)
- Signal "show me" - scan for 10 fingers (both hands up)

TEACHER NOTES:
A full ten frame is an anchor visual for the rest of the unit. Students should recognise "the frame is full" means 10.

WATCH FOR:
- Students who count one at a time - point out the rows: "All 5 here, all 5 here, that's 10"
- Students who flash 10 instantly - good sign

[Stage 1: Daily Review 2 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_A = `SAY:
- The answer is 10. A full ten frame is always 10
- Tick or fix. Say it with me - ten

DO:
- Reveal answer 10
- Lead a quick choral "ten"

TEACHER NOTES:
Naming "a full ten frame is always 10" plants the anchor for Lesson 6 partitioning work.

WATCH FOR:
- Students who can name a full frame without counting - they have one-to-one and cardinality

[Stage 1: Daily Review 2 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_Q = `SAY:
- Last one. Two ten frames this time
- How many counters altogether? Whisper your answer to your partner

DO:
- Display two ten frames - first one full (10), second one with 4
- Allow 20 seconds for partner talk
- Cold call one pair to share

TEACHER NOTES:
Two ten frames is the first taste of teen-number place value. Watch for students who use "10 and 4 more" language.

WATCH FOR:
- Students who count all 14 from 1 - acceptable but slow; nudge to "start at 10"
- Students who say "10 and 4 more is 14" - praise this; it is exactly the language we want

[Stage 1: Daily Review 3 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_A = `SAY:
- The answer is 14. The full ten frame is 10. Four more makes 14
- Read it with me - fourteen

DO:
- Reveal answer 14
- Lead choral "fourteen"
- Note any students who said "forty" - they will need extra support later

TEACHER NOTES:
This previews the tricky teen-twenty pairs that appear later in the lesson (14 vs 40).

WATCH FOR:
- Students who hear "fourteen" and write "40" - common, plan a small group later

[Stage 1: Daily Review 3 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_FL1 = `SAY:
- Fluency time - this is our quick brain warm-up
- Look at the number track
- Count with me from 1 to 30, pointing as we go

DO:
- Display the 1-30 number track
- Point to each number as the class chants
- Repeat twice - the second round faster

TEACHER NOTES:
Group chanting builds the sequence quickly. Pointing forces attention to the numerals themselves, not just the spoken words.

WATCH FOR:
- Students who mouth without sound - tap their shoulder, "louder with us"
- Students who stumble at 20 to 21 transition - flag for later support

[Stage 1: Fluency 1 | VTLM 2.0: Automaticity]`;

const NOTES_FL2 = `SAY:
- Now skip counting by 2s. Watch the highlighted numbers
- Count with me - 2, 4, 6, 8, 10, ready

DO:
- Display the number track with even numbers highlighted
- Lead a choral count of even numbers 2 to 20
- Use the highlighted track as the visual anchor

TEACHER NOTES:
Skip counting by 2s connects to the "skip counting, addition and patterns" fluency focus the user provided. Stop at 20 - keep this lesson focused.

WATCH FOR:
- Students who count 1, 2, 3, 4 quietly - they are not yet skipping; sit beside them next time

[Stage 1: Fluency 2 | VTLM 2.0: Automaticity]`;

const NOTES_FL3_Q = `SAY:
- Pattern check. Look at the numbers - 18, 19, then what
- Write the next number on your whiteboard

DO:
- Display "18, 19, ?" in big numerals
- Allow 8 seconds thinking time
- Signal "show me"

TEACHER NOTES:
A pattern prompt that previews ordering work in Lesson 2. The 19-20 transition trips some students.

WATCH FOR:
- Students who write "10" - they may know the sound "twenty" but cannot write it; show the 20 numeral now
- Students who confidently write 20 - they have the sequence past the decade

[Stage 1: Fluency 3 | VTLM 2.0: Automaticity]`;

const NOTES_FL3_A = `SAY:
- The next number is 20. After 19 comes 20
- Two-zero. The new ten

DO:
- Reveal "20"
- Trace the digits in the air with the class
- Quick choral "after 19 is 20"

TEACHER NOTES:
Naming "two-zero, the new ten" plants language for Lesson 4 (extending to 120) and Lesson 6 (tens and ones).

WATCH FOR:
- Students who chant "twenty" but cannot write "20" - they need writing practice in You Do

[Stage 1: Fluency 3 Answer | VTLM 2.0: Automaticity]`;

const NOTES_VOCAB1 = `SAY:
- A new word for our maths talk - numeral
- A numeral is the way we write a number
- Two is a number. The numeral for two is 2

DO:
- Point to the big "5" on the screen
- Say "this is a numeral. It is the way we write the number five"

TEACHER NOTES:
"Numeral" is the precise word for the written symbol. Use it sparingly today; the rest of the lesson uses "number" for spoken talk.

WATCH FOR:
- Students who mix up "number" and "numeral" - that is fine for now; both work

[General: Vocabulary | VTLM 2.0: Vocabulary Build]`;

const NOTES_LI_SC = `SAY:
- Read with me - We are learning to read and write numerals to 30
- Then the success criteria
- I can read a number out loud
- I can write a number when I hear it
- I can match a numeral to a picture that shows how many

DO:
- Choral read the LI, then each SC
- Point to each line on the slide

TEACHER NOTES:
SC1 is ultra-achievable - reading aloud a numeral the teacher points to. SC2 is the core target and the exit ticket assesses it. SC3 stretches into matching numerals to quantities (a Lesson 3-4 link).

WATCH FOR:
- Students who shake their head at SC2 - reassure them: "We are about to learn how"

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_IDO_READ = `SAY:
- Watch me read a numeral. Here is 23
- The 2 is in the tens place. That tells me 2 tens
- The 3 is in the ones place. That tells me 3 ones
- 2 tens and 3 ones is twenty-three
- I read it out loud - twenty-three

DO:
- Display the numeral 23 large on the slide
- Point to the 2, say "two tens"
- Point to the 3, say "three ones"
- Slide your hand under the whole numeral and say "twenty-three"

TEACHER NOTES:
This I Do builds the place-value anchor that drives the whole unit. Keep it concrete - point and read. Do not over-talk place value names yet; Lesson 6 develops them.

WATCH FOR:
- Students who can read 23 without prompting - call on them in We Do
- Students who say "two three" - praise the effort, then model the full word "twenty-three"

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO_WRITE = `SAY:
- Now I will write a number. Listen first - seventeen
- I hear "teen" - that tells me there is a 1 ten
- I write the 1 first. Then the 7 for seven ones
- One-seven. Seventeen

DO:
- Say "seventeen" clearly
- Write 17 on the board in two steps - the 1 first, then the 7
- Trace each digit slowly so students see the formation

TEACHER NOTES:
Modelling the "1 first, then the ones" routine helps students avoid reversing teens. Foundation students often write "71" for seventeen.

WATCH FOR:
- Students who echo "seventeen" but write 71 in their books - mark for We Do support
- Students who write 17 confidently - move them along quickly

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1 = `SAY:
- Quick check. Look at the numeral on the screen
- Read it together when I point - 1, 2, point

DO:
- Display "26"
- Point firmly on cue
- Listen for "twenty-six"

CFU CHECKPOINT:
Technique: Choral Response
Script:
- "Read this numeral together when I point. Ready - point"
- Scan for: full-class choral "twenty-six"
PROCEED: If 80%+ say twenty-six, move to We Do.
PIVOT: Most likely misconception - students say "two six" or "two and six". Reteach with the 23 example again: "2 tens and 6 ones is twenty-six". Re-check with "29".

TEACHER NOTES:
Choral response gives every student a turn and lets you scan body language for hesitation.

WATCH FOR:
- Quiet students mouthing the wrong sound - sit beside them in We Do

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO_Q = `SAY:
- Your turn with me. I will hold up a numeral card
- You read it together on my signal
- Then your partner will tell you how many counters that would be

DO:
- Hold up cards 12, 18, 25 in turn
- Choral read each card
- Allow 10 seconds of partner talk to picture the quantity

TEACHER NOTES:
This We Do uses physical numeral cards. It bridges I Do (reading) and the upcoming You Do (writing and matching).

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Read numerals to 20 only. Use the matching ten frame as a visual support so students can count to confirm the spoken word.
- Extra Notes: Sit with this small group at the front. Hold up the ten frame alongside the numeral card.
EXTENDING PROMPT:
- Task: After reading the numeral, ask the student to write the numeral that is "one more" on their whiteboard.
- Extra Notes: Use cards 25, 18, 29 - the one-more answer is in the same decade for 25, 18 but crosses the decade for 29.

WATCH FOR:
- Students who pause before reading - they are decoding; give time
- Partners who skip the "how many counters" part - prompt them back

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO_A = `SAY:
- Lets check together - twelve, eighteen, twenty-five
- Eyes back to me - which one had 2 tens and 5 ones?
- Yes - twenty-five

DO:
- Reveal the three numerals on the slide with their spoken names
- Cold call one student to point to the "2 tens and 5 ones" card

TEACHER NOTES:
The reveal also previews tens-and-ones language for Lesson 6.

WATCH FOR:
- Students who point to the right card - they have place value sense, even gently

[Stage 3: We Do Answer | VTLM 2.0: Scaffold Practice]`;

const NOTES_IDO_TRICKY = `SAY:
- Watch this very carefully. Two numbers that sound and look very similar
- This one is 16 - one ten and six ones - sixteen
- This one is 60 - six tens and no ones - sixty
- They sound a bit alike but the digits are in different places

DO:
- Display 16 and 60 side by side, large
- Tap the 1 in 16, say "one ten"
- Tap the 6 in 60, say "six tens"
- Say "sixteen" and "sixty" slowly, exaggerating the ending sounds

MISCONCEPTIONS:
- Misconception: 16 and 60 sound the same, so they are the same
  Why: Young students hear "six" in both words and may not distinguish "-teen" from "-ty"
  Impact: They confuse numbers like 13/30, 14/40, 15/50, 16/60, 17/70, 18/80, 19/90 throughout primary maths
  Quick correction: Use the digits to settle it - "the 1 is in front, so it is sixteen"

TEACHER NOTES:
This is the key tricky-pair moment in Lesson 1. Stay slow. Do not introduce all six pairs - 16 and 60 is enough for today.

WATCH FOR:
- Students who hear no difference between "sixteen" and "sixty" - mark for a small-group phonics support
- Students who use the digits to decide - praise this strategy out loud

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_HINGE_Q = `SAY:
- Hinge check. Three cards are on the screen - 16, 60, 61
- I will say a number out loud
- Hold up the fingers that match - one, two or three
- Show me when I signal

DO:
- Say "sixteen" clearly
- Allow 5 seconds thinking time
- Signal "show me"
- Scan for ones-fingers (card 1)

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- "I will say a number. Hold up 1 if you think it is the first card, 2 for the second, 3 for the third. Sixteen - show me"
- Scan for: 1 finger from 80%+ of the class
PROCEED: If 80%+ show 1, move to You Do.
PIVOT: Most likely misconception - students hold up 2 fingers (they hear "six" and pick "60"). Reteach: place the 16 numeral card alongside a ten frame (10) + ten frame (6) = 16. Place the 60 card alongside 6 ten frames stacked. Re-check with "sixty" then "sixteen" again.

TEACHER NOTES:
This is the lesson's most important CFU. It tests whether students can use the digit position to settle a confusable pair.

WATCH FOR:
- Students who hesitate then copy a neighbour - they are not sure yet; small-group support tomorrow
- Students who hold up fingers fast and confidently - they have place-value sense

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_HINGE_A = `SAY:
- The answer is card 1 - sixteen
- The 1 is in the tens place. The 6 is in the ones place
- Tick or fix on your whiteboard

DO:
- Reveal "Card 1 - sixteen"
- Tap the 1, then the 6 on the displayed card
- Scan the class for the hand signals

TEACHER NOTES:
If many students missed it, replay the tricky-pair I Do once more. Do not move to You Do until 80%+ are confident.

WATCH FOR:
- Students who still seem unsure - book a 5-minute small group during You Do

[Stage 3: CFU Hinge Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Time for independent work
- Take your worksheet
- First - read the numeral
- Next - match it to the picture that shows how many
- Then - write the numeral on the line

DO:
- Distribute the worksheet
- Students work at their desks
- Circulate to spot reversals (e.g. writing 31 for 13)
- For the Extension group, swap their sheet for the Extension when they finish

TEACHER NOTES:
You Do uses different numerals from We Do (worksheet has 11, 17, 22, 28). Section A is to 20, Section B is to 30. Both target SC2 (writing the numeral) and SC3 (matching).

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Work in small group with the teacher. Use plastic numeral cards and the matching ten frame for each item. Read together, then trace the numeral in the worksheet.
- Extra Notes: Sit at the back table with this group during You Do. No separate enabling PDF - the regular worksheet with teacher support is enough.
EXTENDING PROMPT:
- Task: Hand the student the Lesson 1 Extension sheet when they finish the worksheet. The Extension is a number-detective task on tricky pairs like 13/30 and 14/40.
- Extra Notes: Distribute Lesson 1 Extension PDF when ready.

WATCH FOR:
- Students reversing teen digits (13 written as 31) - reteach the "1 first, then the ones" routine
- Students writing the wrong matching numeral - they may be guessing; ask them to count the picture first

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket - one quick question
- I will say a number. Write the numeral on your sheet
- The number is - twenty-three. Write the numeral

DO:
- Display the exit ticket on screen
- Students write the numeral 23 on their sheet
- Collect the slips on the way out

TEACHER NOTES:
Exit ticket targets SC2 (writing the numeral when you hear it). Sort slips into three piles - 23 correct, reversed (32), other. Plan Lesson 2 support accordingly.

WATCH FOR:
- Slips with 32 written - students need more practice writing teens and twenties; small group on Lesson 2 morning
- Slips with 23 written confidently - move to ordering work in Lesson 2 with no extra support

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Lets check our success criteria
- I can read a number out loud - thumbs up, sideways or down
- I can write a number when I hear it - thumbs
- I can match a numeral to a picture - thumbs
- Tell your partner one number that you can now write that you could not write last week

DO:
- Run thumbs check for each SC
- Allow 30 seconds partner talk
- Cold call one student to share

TEACHER NOTES:
The reflection prompt anchors progress in concrete terms. Closing also primes Lesson 2 - ordering these same numerals on a number line.

WATCH FOR:
- Students who show thumbs down on SC2 - small group on Lesson 2 morning
- Students who share confidently - they are ready for ordering work

[General: Closing | VTLM 2.0: Review and Reflect]`;

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres,
    LESSON_TITLE,
    "Lesson 1 of 10 - Numbers to 120",
    "Year 1 Numeracy | AC9M1N01",
    NOTES_TITLE);

  // Slide 2: Teacher Resources (placed immediately after title per megaprompt §44)
  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Slide 3: Launch
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Launch");
    addTitle(s, "What number could it be?", { color: STAGE_COLORS["1"] });

    // Hero digits 2 and 5 - large, side by side
    drawNumeralCard(s, 3.4, 1.8, 1.5, 1.8, "2", { fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });
    drawNumeralCard(s, 5.1, 1.8, 1.5, 1.8, "5", { fill: C.SECONDARY, color: C.WHITE, stroke: C.SECONDARY });

    // Hint card below
    addInstructionCard(s, [
      { role: "header", text: "Tell your partner" },
      { role: "body", text: "What number has a 2 and a 5?" },
    ], { x: 0.5, y: 4.0, w: 9, h: 1.0, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // Slide 4-5: Daily Review 1 - How many counters (7 on a ten frame)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "How many counters?", { color: STAGE_COLORS["1"] });

      // Ten frame with 7 counters - large, centred
      addTensFrame(s, 2.5, 1.7, 5.0, 7, { cellH: 1.0, fillColor: C.PRIMARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR1_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "7 counters", {
        x: 2.5, y: 4.55, w: 5.0, h: 0.6, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 28, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR1_A);
    }
  );

  // Slide 6-7: Daily Review 2 - How many counters (10 full frame)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "How many counters?", { color: STAGE_COLORS["1"] });

      addTensFrame(s, 2.5, 1.7, 5.0, 10, { cellH: 1.0, fillColor: C.PRIMARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR2_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "10 counters - a full ten frame!", {
        x: 1.5, y: 4.55, w: 7.0, h: 0.6, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR2_A);
    }
  );

  // Slide 8-9: Daily Review 3 - How many counters (10 + 4 on double frame)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "How many altogether?", { color: STAGE_COLORS["1"] });

      // Two ten frames side by side, first full, second with 4
      addTensFrame(s, 0.7, 1.9, 4.1, 10, { cellH: 0.82, fillColor: C.PRIMARY });
      addTensFrame(s, 5.2, 1.9, 4.1, 4, { cellH: 0.82, fillColor: C.SECONDARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR3_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "14 counters - 10 and 4 more!", {
        x: 1.5, y: 4.55, w: 7.0, h: 0.6, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_DR3_A);
    }
  );

  // Slide 10: Fluency 1 - Count 1 to 30
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "Count with me: 1 to 30", { color: STAGE_COLORS["1"] });

    // Three rows of 10 numbers each
    const rowY = [1.7, 2.55, 3.4];
    for (let r = 0; r < 3; r += 1) {
      const start = 1 + r * 10;
      const end = start + 9;
      addNumberTrack(s, 0.5, rowY[r], 9, start, end, [], { cellH: 0.7, fontSize: 22 });
    }

    addInstructionCard(s, [
      { role: "header", text: "Point and count" },
      { role: "body", text: "Count from 1 to 30 together, then again faster." },
    ], { x: 0.5, y: 4.35, w: 9, h: 0.75, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FL1);
  })();

  // Slide 11: Fluency 2 - Skip count by 2s to 20
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Fluency");
    addTitle(s, "Skip count by 2s", { color: STAGE_COLORS["1"] });

    // Two rows of 10, highlight evens 2-20
    addNumberTrack(s, 0.5, 2.0, 9, 1, 10, [2, 4, 6, 8, 10], { cellH: 0.8, fontSize: 24 });
    addNumberTrack(s, 0.5, 2.95, 9, 11, 20, [12, 14, 16, 18, 20], { cellH: 0.8, fontSize: 24 });

    addInstructionCard(s, [
      { role: "header", text: "Count the highlighted numbers" },
      { role: "body", text: "2, 4, 6, 8, 10... keep going!" },
    ], { x: 0.5, y: 4.1, w: 9, h: 1.0, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FL2);
  })();

  // Slide 12-13: Fluency 3 - Pattern (18, 19, ?)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Fluency");
      addTitle(s, "What number comes next?", { color: STAGE_COLORS["1"] });

      drawNumeralCard(s, 1.5, 2.0, 1.8, 2.0, "18", { fontSize: 72 });
      drawNumeralCard(s, 4.1, 2.0, 1.8, 2.0, "19", { fontSize: 72 });
      drawNumeralCard(s, 6.7, 2.0, 1.8, 2.0, "?", { fontSize: 72, color: C.ALERT, stroke: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_FL3_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "20 - the new ten!", {
        x: 2.5, y: 4.55, w: 5.0, h: 0.6, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_FL3_A);
    }
  );

  // Slide 14: Vocabulary - numeral
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Word", { color: C.PRIMARY });
    addTitle(s, "Numeral", { color: C.PRIMARY });

    // Big hero numeral
    drawNumeralCard(s, 3.7, 1.7, 2.6, 2.2, "5", { fontSize: 130, fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });

    addInstructionCard(s, [
      { role: "header", text: "A numeral is the way we write a number." },
      { role: "body", text: "The numeral for five is 5." },
      { role: "body", text: "The numeral for twenty-three is 23." },
    ], { x: 0.5, y: 4.1, w: 9, h: 1.0, strip: C.ACCENT });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB1);
  })();

  // Slide 15: LI/SC
  liSlide(pres,
    ["We are learning to read and write numerals to 30."],
    [
      "I can read a number out loud.",
      "I can write a number when I hear it.",
      "I can match a numeral to a picture that shows how many.",
    ],
    NOTES_LI_SC, FOOTER);

  // Slide 16: I Do — Read a numeral (23)
  workedExSlide(pres, 2, "I Do", "How to read a numeral",
    [
      "I read each digit.",
      "The 2 is in the tens place - 2 tens.",
      "The 3 is in the ones place - 3 ones.",
      "2 tens and 3 ones is twenty-three.",
    ],
    NOTES_IDO_READ, FOOTER,
    (slide, lg) => {
      // Right column: big 23 with tens/ones labels
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.PRIMARY });
      // Digit boxes
      drawNumeralCard(slide, lg.rightX + 0.45, lg.panelTopPadded + 0.3, 1.4, 1.6, "2",
        { fontSize: 70, fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });
      drawNumeralCard(slide, lg.rightX + 2.05, lg.panelTopPadded + 0.3, 1.4, 1.6, "3",
        { fontSize: 70, fill: C.SECONDARY, color: C.WHITE, stroke: C.SECONDARY });
      // Labels
      slide.addText("tens", {
        x: lg.rightX + 0.45, y: lg.panelTopPadded + 1.95, w: 1.4, h: 0.3,
        fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      slide.addText("ones", {
        x: lg.rightX + 2.05, y: lg.panelTopPadded + 1.95, w: 1.4, h: 0.3,
        fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, align: "center", margin: 0,
      });

      // Answer banner
      addTextOnShape(slide, "twenty-three", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 2.6, w: lg.rightW - 0.6, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 17: I Do — Write a numeral (17 from hearing "seventeen")
  workedExSlide(pres, 2, "I Do", "How to write a numeral",
    [
      "I hear seventeen.",
      "The teen tells me 1 ten.",
      "I write the 1 first.",
      "Then I write 7 ones.",
      "1 then 7 - seventeen.",
    ],
    NOTES_IDO_WRITE, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.4, { strip: C.SECONDARY });

      // Show 17 written - the 1 first highlighted
      drawNumeralCard(slide, lg.rightX + 0.45, lg.panelTopPadded + 0.3, 1.4, 1.6, "1",
        { fontSize: 70, fill: C.ALERT, color: C.WHITE, stroke: C.ALERT });
      drawNumeralCard(slide, lg.rightX + 2.05, lg.panelTopPadded + 0.3, 1.4, 1.6, "7",
        { fontSize: 70, fill: C.SECONDARY, color: C.WHITE, stroke: C.SECONDARY });

      slide.addText("write FIRST", {
        x: lg.rightX + 0.45, y: lg.panelTopPadded + 1.95, w: 1.4, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true, align: "center", margin: 0,
      });
      slide.addText("then", {
        x: lg.rightX + 2.05, y: lg.panelTopPadded + 1.95, w: 1.4, h: 0.3,
        fontSize: 13, fontFace: FONT_B, color: C.SECONDARY, bold: true, align: "center", margin: 0,
      });

      addTextOnShape(slide, "seventeen", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 2.6, w: lg.rightW - 0.6, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 18-19: CFU 1 - Read the numeral (26)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Read this numeral together", { color: C.ALERT });

      drawNumeralCard(s, 3.5, 1.8, 3.0, 2.4, "26",
        { fontSize: 160, fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });

      addInstructionCard(s, [
        { role: "header", text: "Together when I point" },
        { role: "body", text: "Eyes up. Read the number." },
      ], { x: 0.5, y: 4.4, w: 9, h: 0.7, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "twenty-six", {
        x: 2.5, y: 4.55, w: 5.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 24, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 20-21: We Do - Read these cards (12, 18, 25)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Read these numbers together", { color: STAGE_COLORS["3"] });

      drawNumeralCard(s, 1.0, 1.9, 2.4, 2.0, "12", { fontSize: 90 });
      drawNumeralCard(s, 3.8, 1.9, 2.4, 2.0, "18", { fontSize: 90 });
      drawNumeralCard(s, 6.6, 1.9, 2.4, 2.0, "25", { fontSize: 90 });

      addInstructionCard(s, [
        { role: "header", text: "Read together when I point" },
        { role: "body", text: "Then tell your partner how many counters that would be." },
      ], { x: 0.5, y: 4.1, w: 9, h: 1.0, strip: C.SECONDARY });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "twelve - eighteen - twenty-five", {
        x: 0.5, y: 4.55, w: 9.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_WEDO_A);
    }
  );

  // Slide 22: I Do - Tricky pair 16 vs 60
  workedExSlide(pres, 2, "I Do", "Tricky pair - 16 and 60",
    [
      "16 - the 1 is in the tens place.",
      "1 ten and 6 ones is sixteen.",
      "",
      "60 - the 6 is in the tens place.",
      "6 tens and 0 ones is sixty.",
      "",
      "The digit position tells me the number.",
    ],
    NOTES_IDO_TRICKY, FOOTER,
    (slide, lg) => {
      // Two cards stacked: 16 on top, 60 below
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 1.55, { strip: C.PRIMARY });
      drawNumeralCard(slide, lg.rightX + 0.4, lg.panelTopPadded + 0.2, 1.5, 1.15, "16",
        { fontSize: 60, fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });
      slide.addText("sixteen", {
        x: lg.rightX + 2.0, y: lg.panelTopPadded + 0.45, w: lg.rightW - 2.2, h: 0.7,
        fontSize: 24, fontFace: FONT_H, color: C.PRIMARY, bold: true, valign: "middle", margin: 0,
      });

      addCard(slide, lg.rightX, lg.panelTopPadded + 1.7, lg.rightW, 1.55, { strip: C.SECONDARY });
      drawNumeralCard(slide, lg.rightX + 0.4, lg.panelTopPadded + 1.9, 1.5, 1.15, "60",
        { fontSize: 60, fill: C.SECONDARY, color: C.WHITE, stroke: C.SECONDARY });
      slide.addText("sixty", {
        x: lg.rightX + 2.0, y: lg.panelTopPadded + 2.15, w: lg.rightW - 2.2, h: 0.7,
        fontSize: 24, fontFace: FONT_H, color: C.SECONDARY, bold: true, valign: "middle", margin: 0,
      });
    }
  );

  // Slide 23-24: Hinge CFU - Which card is 16?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU - Hinge", { color: C.ALERT });
      addTitle(s, "Which card is sixteen?", { color: C.ALERT });

      // Three numeral cards: 16, 60, 61
      drawNumeralCard(s, 1.0, 2.0, 2.4, 2.0, "16", { fontSize: 90, fill: C.PRIMARY, color: C.WHITE, stroke: C.PRIMARY });
      drawNumeralCard(s, 3.8, 2.0, 2.4, 2.0, "60", { fontSize: 90, fill: C.SECONDARY, color: C.WHITE, stroke: C.SECONDARY });
      drawNumeralCard(s, 6.6, 2.0, 2.4, 2.0, "61", { fontSize: 90, fill: C.ACCENT, color: C.WHITE, stroke: C.ACCENT });

      // Card labels
      s.addText("Card 1", { x: 1.0, y: 4.05, w: 2.4, h: 0.25, fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });
      s.addText("Card 2", { x: 3.8, y: 4.05, w: 2.4, h: 0.25, fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });
      s.addText("Card 3", { x: 6.6, y: 4.05, w: 2.4, h: 0.25, fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, align: "center", margin: 0 });

      addInstructionCard(s, [
        { role: "header", text: "Hold up 1, 2 or 3 fingers" },
        { role: "body", text: "Which card shows sixteen?" },
      ], { x: 0.5, y: 4.4, w: 9, h: 0.7, strip: C.ALERT });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_HINGE_Q);
      return s;
    },
    (slide) => {
      addTextOnShape(slide, "Card 1 - sixteen!", {
        x: 0.5, y: 4.55, w: 4.0, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
      slide.addNotes(NOTES_HINGE_A);
    }
  );

  // Slide 25: You Do
  workedExSlide(pres, 4, "You Do", "Read, match and write",
    [
      "First - read each numeral.",
      "Next - match it to the picture.",
      "Then - write the numeral on the line.",
      "",
      "Show your reading!",
      "You have 8 minutes.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 2.5, { strip: C.ALERT });
      slide.addText("Remember", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.1, w: lg.rightW - 0.4, h: 0.35,
        fontSize: 18, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Read each numeral out loud.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Count the picture to check.", options: { bullet: true, fontSize: 15, color: C.CHARCOAL, breakLine: true } },
        { text: "Write the 1 first for teens!", options: { bullet: true, fontSize: 15, color: C.ALERT, bold: true } },
      ], {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 0.55, w: lg.rightW - 0.5, h: 1.7,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
    }
  );

  // Slide 26: Exit Ticket
  exitTicketSlide(pres,
    [
      "I say twenty-three. Write the numeral.",
    ],
    NOTES_EXIT, FOOTER);

  // Slide 27: Closing
  closingSlide(pres, {
    reflectionPrompt: "Tell your partner: one number you can now write that you could not write last week.",
    scItems: [
      "I can read a number out loud.",
      "I can write a number when I hear it.",
      "I can match a numeral to a picture that shows how many.",
    ],
    selfAssessment: "Thumbs up, sideways or down for each one",
  }, NOTES_CLOSING);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "G1N_Lesson1_Reading_Writing_Numerals.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Generate PDFs ─────────────────────────────────────────────────────────

  // Worksheet — Read, match and write numerals to 30
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Read, match and write numerals to 30",
      color: C.NAVY,
      lessonInfo: "Lesson 1 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "Read each numeral out loud. Count the ten-frame picture. Write the numeral on the line.", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Section A: Numerals to 20", y, { color: C.NAVY });

    // Draw a row: numeral box, ten frame picture, write line
    function drawRow(doc, label, numeral, filled, y) {
      const rowH = 70;
      const x = 60;
      // Numeral box
      doc.save();
      doc.roundedRect(x, y, 70, rowH, 6).lineWidth(1.5).strokeColor(hex(C.NAVY)).stroke();
      doc.fontSize(40).font("Sans-Bold").fillColor(hex(C.NAVY));
      doc.text(numeral, x, y + 12, { width: 70, align: "center" });
      doc.restore();

      // Ten-frame picture (2 rows of 5 cells)
      const tfX = x + 90;
      const tfY = y + 12;
      const cellSize = 22;
      for (let r = 0; r < 2; r += 1) {
        for (let c = 0; c < 5; c += 1) {
          const idx = r * 5 + c;
          doc.rect(tfX + c * cellSize, tfY + r * cellSize, cellSize, cellSize)
            .lineWidth(1).strokeColor("#444").stroke();
          if (idx < filled) {
            doc.circle(tfX + c * cellSize + cellSize / 2, tfY + r * cellSize + cellSize / 2, cellSize * 0.32)
              .fill(hex(C.PRIMARY));
          }
        }
      }

      // Write line
      const lineX = tfX + 5 * cellSize + 20;
      const lineW = 200;
      doc.fontSize(11).font("Sans-Bold").fillColor("#000");
      doc.text("Write:", lineX, y + 30);
      doc.moveTo(lineX + 50, y + 50).lineTo(lineX + 50 + lineW, y + 50)
        .lineWidth(1.2).strokeColor("#000").stroke();

      return y + rowH + 10;
    }

    function hex(c) { return c.startsWith("#") ? c : "#" + c; }

    y = drawRow(doc, "1", "11", 11, y);
    y = drawRow(doc, "2", "17", 17, y);

    y = addSectionHeading(doc, "Section B: Numerals to 30", y, { color: C.NAVY });
    y = drawRow(doc, "3", "22", 22, y);  // 22 won't fit on 1 ten frame — show 20 in first frame visually
    y = drawRow(doc, "4", "28", 28, y);

    addPdfFooter(doc, "Lesson 1 | Numbers to 120 | Year 1");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Read, match and write numerals to 30 - ANSWERS",
      color: C.NAVY,
      lessonInfo: "Lesson 1 of 10 | Year 1 Numeracy",
    });
    y = addSectionHeading(doc, "Section A: Numerals to 20", y, { color: C.NAVY });
    y = addBodyText(doc, "1. 11 (eleven) - 11 counters on the ten frame", y);
    y = addBodyText(doc, "2. 17 (seventeen) - 17 counters on the ten frame", y);
    y = addSectionHeading(doc, "Section B: Numerals to 30", y, { color: C.NAVY });
    y = addBodyText(doc, "3. 22 (twenty-two) - 22 counters", y);
    y = addBodyText(doc, "4. 28 (twenty-eight) - 28 counters", y);
    y = addBodyText(doc, "Watch for: students writing 71 for 17 or 22 reversed - reteach the 'write the 1 first' routine.", y);
    addPdfFooter(doc, "Lesson 1 | Answer Key | Year 1");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Extension - Number detective task on tricky pairs
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Number detective - tricky pairs!",
      color: C.TEAL,
      lessonInfo: "Lesson 1 of 10 | Year 1 Numeracy",
    });
    y = addTipBox(doc, "These numbers sound a bit alike. Use the digits to be a number detective. Circle the right one.", y, { color: C.TEAL });

    function hex(c) { return c.startsWith("#") ? c : "#" + c; }

    function drawDetectiveRow(doc, label, spoken, options, y) {
      const x = 60;
      doc.fontSize(13).font("Sans-Bold").fillColor("#000");
      doc.text(label + ". I say:  " + spoken, x, y);
      y += 22;
      // Three numeral boxes
      const boxW = 90, boxH = 60, gap = 30;
      options.forEach((num, i) => {
        const cx = x + i * (boxW + gap);
        doc.roundedRect(cx, y, boxW, boxH, 6).lineWidth(1.5).strokeColor(hex(C.NAVY)).stroke();
        doc.fontSize(36).font("Sans-Bold").fillColor(hex(C.NAVY));
        doc.text(num, cx, y + 12, { width: boxW, align: "center" });
      });
      return y + boxH + 20;
    }

    y = addSectionHeading(doc, "Circle the right numeral", y, { color: C.NAVY });
    y = drawDetectiveRow(doc, "1", "sixteen", ["6", "16", "60"], y);
    y = drawDetectiveRow(doc, "2", "thirty", ["3", "13", "30"], y);
    y = drawDetectiveRow(doc, "3", "fifteen", ["5", "15", "50"], y);
    y = drawDetectiveRow(doc, "4", "forty", ["4", "14", "40"], y);

    addPdfFooter(doc, "Lesson 1 | Extension | Year 1");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();
}

build().catch((err) => { console.error(err); process.exit(1); });
