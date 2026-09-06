"use strict";

// Year 1 Numeracy — Doubles facts to double 6
// AC9M1A02 — recognise, continue and create repeating patterns; connect doubles
//            to equal groups and counting (AC9M1N02 add and subtract within 20).
// Manipulative-led: double ten frames, dot/dice cards, counters, fingers.
// Daily Review: Counting on from a number and simple addition within 10.
// Fluency:      Subitising dot patterns to 6.
// Mixed readiness, NOT a revision lesson — build it together.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, hex,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

// Single lesson. Variant 0 (no week given for a standalone lesson).
const T = createTheme("numeracy", "grade1", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, closingSlide,
  workedExSlide, exitTicketSlide,
  addStageBadge,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  addInstructionCard,
  withReveal,
  addTensFrame, addDotCard, addNumberTrack,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const LESSON_TITLE = "Doubles to Double 6";
const FOOTER = "Doubles to Double 6 | Year 1 Numeracy";
const OUT_DIR = "output/G1D_Doubles_to_Double6";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION, "Doubles Practice",
  "Build, count and write doubles from double 1 to double 6.");
const ANSWER_KEY_RES = makeSessionResource(SESSION, "Answer Key",
  "Teacher reference with answers for the Doubles Practice sheet.");
const EXTENDING_RES = makeSessionResource(SESSION, "Doubles Challenge",
  "Early-finisher mat: prove a double two ways and make your own double.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES, EXTENDING_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// ─── Local visual helpers ─────────────────────────────────────────────────────

// Two dot/dice cards with a "+" between them (two groups). Returns geometry.
function drawGroups(slide, n1, n2, opts) {
  const o = opts || {};
  const size = o.size || 1.8;
  const gap = o.gap || 0.85;
  const totalW = size * 2 + gap;
  const x = o.x != null ? o.x : (10 - totalW) / 2;
  const y = o.y != null ? o.y : 1.95;
  const c1 = o.c1 || C.PRIMARY;
  const c2 = o.c2 || C.SECONDARY;
  addDotCard(slide, x, y, size, n1, { dotColor: c1, borderColor: C.CHARCOAL });
  slide.addText("+", {
    x: x + size, y, w: gap, h: size,
    fontSize: o.plusSize || 48, fontFace: FONT_H, color: C.CHARCOAL,
    bold: true, align: "center", valign: "middle", margin: 0,
  });
  addDotCard(slide, x + size + gap, y, size, n2, { dotColor: c2, borderColor: C.CHARCOAL });
  return { x, y, size, gap, totalW, right: x + totalW, midY: y + size / 2 };
}

// Small group label under a card (e.g. "3").
function groupLabel(slide, x, w, y, text, color) {
  slide.addText(String(text), {
    x, y, w, h: 0.34,
    fontSize: 18, fontFace: FONT_B, color: color || C.CHARCOAL,
    bold: true, align: "center", valign: "middle", margin: 0,
  });
}

// A success-coloured answer bar for reveal slides.
function answerBar(slide, text, opts) {
  const o = opts || {};
  addTextOnShape(slide, String(text), {
    x: o.x != null ? o.x : 1.5, y: o.y != null ? o.y : 4.5,
    w: o.w != null ? o.w : 7.0, h: o.h != null ? o.h : 0.6, rectRadius: 0.08,
    fill: { color: o.color || C.SUCCESS },
  }, { fontSize: o.fontSize || 26, fontFace: FONT_H, color: C.WHITE, bold: true });
}

// ─── Teacher Notes ─────────────────────────────────────────────────────────────

const NOTES_TITLE = `SAY:
- Today we are learning all about doubles
- A double is when we have two groups that are exactly the same
- We will use dot cards, ten frames, counters and even our fingers

DO:
- Show the title slide as students arrive on the mat
- Have double ten frames, counters and dot cards ready at the front
- Settle the class with a quick choral count to 10

TEACHER NOTES:
This is a build-it-together lesson, not a revision lesson. Some children will know a few doubles already; many will be meeting the word double for the first time.

WATCH FOR:
- Children who already say doubles fast - they become helpers in We Do
- Children who look unsure - reassure them, we will build every double together

[General: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_RESOURCES = `SAY:
- These are the things we will use today
- We will share counters and a double ten frame with our partner

DO:
- Point to each resource and material on the slide
- Stack counters, dot cards and double ten frames at the front before the lesson
- Keep the Doubles Practice sheet and answer key handy for after We Do

TEACHER NOTES:
Print the Doubles Practice sheet single sided. The Doubles Challenge is for early finishers. Counters and ten frames live on the maths trolley.

WATCH FOR:
- Any missing materials - sort this before the lesson starts, not during

[General: Resources | VTLM 2.0: Planning]`;

const NOTES_DR1_Q = `SAY:
- Daily Review first - this warms up our maths brain
- Look at the number track. Put your finger on 6 in the air
- Start at 6 and count on 3 more. Where do you land? Whisper it to your partner

DO:
- Display the number track with 6 highlighted
- Model touching 6, then tapping on 7, 8, 9 in the air
- Give 15 seconds of partner talk, then signal show me on fingers

TEACHER NOTES:
Daily Review 1 of 3. This reviews counting on, which is prior knowledge that doubles will build on.

WATCH FOR:
- Children who count 6, 7, 8 and land on 8 - they counted the start number; show them to start on the next number
- Children who land on 9 quickly - they count on confidently

[Stage 1: Daily Review 1 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR1_A = `SAY:
- Show me your answer. The answer is 9
- We started at 6, then counted on - 7, 8, 9
- Tick if you got it. Fix it if you did not - fixing is the learning

DO:
- Reveal the answer 9
- Re-count 7, 8, 9 out loud with the class
- Scan whiteboards for 8 or 10

TEACHER NOTES:
Counting on is the bridge to adding two groups, which is what a double is.

WATCH FOR:
- Children who wrote 8 - they included the start number; quick reminder, count the jumps not the start

[Stage 1: Daily Review 1 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_Q = `SAY:
- Next one. Two groups of dots
- How many altogether? Count them all in your head, then tell your partner

DO:
- Display the dot cards 4 and 3
- Give 15 seconds of partner talk
- Signal show me on fingers

TEACHER NOTES:
Daily Review 2 of 3. Simple addition within 10 using two groups of dots, the same picture we will use for doubles today.

WATCH FOR:
- Children who count the first group again - point to one group, then the other
- Children who just know it is 7 - they are subitising and counting on

[Stage 1: Daily Review 2 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR2_A = `SAY:
- The answer is 7. Four and three more is seven
- Tick or fix on your whiteboard

DO:
- Reveal the answer 7
- Count all seven dots together, pointing to each one

TEACHER NOTES:
Notice the two groups are different sizes here. Later we will see doubles, where the groups are the same.

WATCH FOR:
- Children who answer instantly - ready for the main lesson

[Stage 1: Daily Review 2 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_Q = `SAY:
- Last warm up. Two more groups of dots
- How many altogether? Whisper your answer to your partner

DO:
- Display the dot cards 5 and 4
- Give 15 seconds of partner talk
- Signal show me on fingers

TEACHER NOTES:
Daily Review 3 of 3. Adding within 10 again, slightly larger groups.

WATCH FOR:
- Children who count all from 1 - acceptable; nudge to start at 5 and count on
- Children who say 9 fast - strong counting on

[Stage 1: Daily Review 3 | VTLM 2.0: Retention and Recall]`;

const NOTES_DR3_A = `SAY:
- The answer is 9. Five and four more is nine
- Tick or fix. Great warming up

DO:
- Reveal the answer 9
- Count all nine dots together

TEACHER NOTES:
We are ready to move into subitising for fluency, then today's doubles.

WATCH FOR:
- Children who wrote 8 or 10 - a miscount; recount the second group together

[Stage 1: Daily Review 3 Answer | VTLM 2.0: Retention and Recall]`;

const NOTES_FL1_Q = `SAY:
- Fluency time - quick eyes, no counting one by one
- How many dots? Just know it. Get ready to show me on your fingers

DO:
- Flash the dot card showing 4
- Count to 3 in your head, then signal show me

TEACHER NOTES:
Fluency 1 of 3. Subitising dot patterns to 6 - the user fluency focus. Keep it brisk.

WATCH FOR:
- Children who point and count each dot - that is counting, not subitising; flash it again briefly
- Children who say 4 instantly - they are subitising the dice pattern

[Stage 1: Fluency 1 | VTLM 2.0: Automaticity]`;

const NOTES_FL1_A = `SAY:
- It is 4. The corners pattern is always 4

DO:
- Reveal the 4
- Show me thumbs if you just knew it

TEACHER NOTES:
Subitising the 4 dice pattern supports seeing double 2 as two groups of 2 later.

WATCH FOR:
- Children who needed to count - keep flashing dot cards across the week

[Stage 1: Fluency 1 Answer | VTLM 2.0: Automaticity]`;

const NOTES_FL2_Q = `SAY:
- Next flash. Quick eyes - how many dots?
- Get ready to show me

DO:
- Flash the dot card showing 6
- Count to 3 in your head, then signal show me

TEACHER NOTES:
Fluency 2 of 3. Six is the largest subitising target today and links to double 3 (two groups of 3).

WATCH FOR:
- Children who say 5 or 7 - a quick miscount; flash again
- Children who see two columns of 3 - praise that, it is exactly the doubles idea

[Stage 1: Fluency 2 | VTLM 2.0: Automaticity]`;

const NOTES_FL2_A = `SAY:
- It is 6. Two rows of three make six

DO:
- Reveal the 6
- Point to the two rows of three

TEACHER NOTES:
Naming three and three makes six plants the double 3 fact we teach today.

WATCH FOR:
- Children who see it as three and three - they are ready for doubles

[Stage 1: Fluency 2 Answer | VTLM 2.0: Automaticity]`;

const NOTES_FL3_Q = `SAY:
- Last flash. Quick eyes - how many dots?
- Show me when I signal

DO:
- Flash the dot card showing 5
- Count to 3 in your head, then signal show me

TEACHER NOTES:
Fluency 3 of 3. Five as a dice pattern, a key anchor for ten frames coming up.

WATCH FOR:
- Children who say 4 or 6 - flash again slowly
- Children who just know 5 - strong subitising

[Stage 1: Fluency 3 | VTLM 2.0: Automaticity]`;

const NOTES_FL3_A = `SAY:
- It is 5. The dice five - four corners and one in the middle

DO:
- Reveal the 5
- Quick thumbs if you knew it straight away

TEACHER NOTES:
Five recognised quickly helps when we read a full top row on a ten frame.

WATCH FOR:
- Children still counting - more dot card practice this week

[Stage 1: Fluency 3 Answer | VTLM 2.0: Automaticity]`;

const NOTES_LAUNCH = `SAY:
- Hands up everyone. Show me 3 fingers on this hand, and 3 fingers on the other hand
- Look - both hands are the same. Three and three
- How many fingers altogether? Tell your partner
- Some of you may notice the two groups are exactly the same. That is our special new idea today

DO:
- Hold up your own hands showing 3 and 3
- Give 20 seconds partner talk
- Cold call one pair to share, then say we call this a double

TEACHER NOTES:
The launch activates adding two groups from Daily Review and bridges to today: when the two groups are the same, it is a double. Do not name the answer too early - let students find 6.

WATCH FOR:
- Children who show 3 and 2 - gently fix so both hands match
- Children who say six straight away - ask how they knew, listen for three and three

[Stage 1: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LISC = `SAY:
- Read with me - We are learning to find doubles to double 6
- Now our success criteria
- I can make two equal groups
- I can find how many a double makes
- I can say a double fact

DO:
- Choral read the learning intention, then each success criterion
- Point to each line on the slide

TEACHER NOTES:
The first criterion is ultra-achievable - making two equal groups. The middle one is the core target and the exit ticket checks it. The last one stretches to saying the fact, like double 4 is 8.

WATCH FOR:
- Children who look worried at the middle one - reassure them, we are about to learn exactly how

[General: LI/SC | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB = `SAY:
- A new maths word today - double
- A double means two groups that are exactly the same
- Here are two groups of three. They match. That is a double

DO:
- Point to the first group of 3, then the second group of 3
- Say the groups are equal - they are the same
- Have the class say the word double with you

TEACHER NOTES:
Keep this slide to the meaning of double - two equal groups. The total comes in the I Do. Equal simply means the same here.

WATCH FOR:
- Children who think any two groups make a double - stress the groups must be the same

[General: Vocabulary | VTLM 2.0: Vocabulary Build]`;

const NOTES_IDO1 = `SAY:
- Let us make our first double together. Watch how I do it
- I make one group of three, then another group of three. I check - both groups are the same
- Now I count all the dots - one, two, three, four, five, six
- Double 3 is 6. The two equal groups joined make six

DO:
- Point to the first group of 3, then the second group of 3
- Run your finger under both as you count all six dots
- Say double 3 is 6 and have the class echo it

TEACHER NOTES:
This is the core routine for the lesson - make two equal groups, count all, say the double fact. Keep it slow and concrete.

WATCH FOR:
- Children who only count one group - point to both groups before counting
- Children who can already say double 3 is 6 - call on them in We Do

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_IDO2 = `SAY:
- Now a double on our double ten frame. Watch me
- I put 4 counters in this frame, and 4 counters in the other frame. I check - both frames are the same
- I count all - four in here, then five, six, seven, eight
- Double 4 is 8

DO:
- Fill the first ten frame with 4, then the second with 4
- Count all eight counters, touching each one
- Say double 4 is 8 and have the class echo it

TEACHER NOTES:
Same routine, new model. The double ten frame shows the two equal groups clearly. Do not fill whole frames - only 4 in each so the equal groups stay obvious.

WATCH FOR:
- Children who put different amounts in each frame - stop and match them
- Children ready to predict double 5 - hold that thought for We Do

[Stage 2: I Do | VTLM 2.0: Explicit Modelling]`;

const NOTES_CFU1_Q = `SAY:
- Quick check. Two pictures, A and B
- One picture shows a double. One does not
- Hold up A or B. Which one shows two equal groups?

DO:
- Display picture A (3 and 3) and picture B (4 and 2)
- Give 10 seconds thinking time
- Signal show me - hold up A or B fingers

CFU CHECKPOINT:
Technique: Finger Voting
Script:
- "Hold up A if the first picture is the double, B if the second one is. Show me"
- Scan for: most of the class choosing A
PROCEED: If 80 percent or more choose A, move to We Do.
PIVOT: Most likely misconception - children pick B because it also has two groups. Reteach by counting each group aloud - A is three and three, the same; B is four and two, not the same. Re-check with a fresh pair, 2 and 2 against 3 and 1.

TEACHER NOTES:
This checks the heart of the concept - a double needs two equal groups, not just any two groups.

WATCH FOR:
- Children who choose B - they are not yet checking that the groups match

[Stage 2: CFU | VTLM 2.0: Monitor Progress]`;

const NOTES_CFU1_A = `SAY:
- The answer is A. Three and three - the groups are the same
- B is four and two. The groups are not the same, so it is not a double

DO:
- Reveal A is the double
- Count both groups in A to prove they match
- Count B to show four and two are different

TEACHER NOTES:
If many chose B, model one more equal-versus-not-equal pair before We Do.

WATCH FOR:
- Children who now self-correct - they have the idea

[Stage 2: CFU Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_WEDO1_Q = `SAY:
- Your turn with me. Let us build double 2
- Make one group of two counters, then another group of two. Check - are they the same?
- Count all the dots. How many is double 2? Tell your partner

DO:
- Build 2 and 2 on the board as students build with counters
- Move along the rows checking for two equal groups
- Take answers on fingers before the reveal

TEACHER NOTES:
First guided double. Keep the make, check, count all, say it routine visible.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Build the two groups onto a five frame each so the equal groups are framed and easy to count.
- Extra Notes: Sit with this small group and count all together.
EXTENDING PROMPT:
- Task: After double 2, ask the child to predict double 3 and build it to check.
- Extra Notes: They prove their prediction with counters.

WATCH FOR:
- Children who build 2 and 1 - prompt them to match the groups
- Children who count all correctly to 4 - ready to keep going

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO1_A = `SAY:
- Let us check. Two and two - the groups match
- Count all - one, two, three, four. Double 2 is 4
- Tick or fix your work

DO:
- Reveal double 2 is 4
- Count the four counters together

TEACHER NOTES:
Praise children who said the full fact, double 2 is 4.

WATCH FOR:
- Children who said 3 - they likely counted only one extra; recount together

[Stage 3: We Do Answer | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO2_Q = `SAY:
- Next double together, on the double ten frame. Let us build double 5
- Put 5 in this frame and 5 in the other frame. Check - are they the same?
- Count all. How many is double 5? Whisper to your partner

DO:
- Fill 5 and 5 on the board double ten frame as students build
- Check each pair has two matching full top rows
- Take answers before the reveal

TEACHER NOTES:
Second guided double, larger and crossing into the teens. A full top row of five helps children count on from five.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Count all from one, touching each counter, rather than counting on.
- Extra Notes: Stay with this group through the count.
EXTENDING PROMPT:
- Task: Ask the child how they could count without starting at one - count on from five.
- Extra Notes: Listen for five, then six, seven, eight, nine, ten.

WATCH FOR:
- Children who fill the whole frame - remind them only 5 in each
- Children who count on from five - strong strategy

[Stage 3: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_WEDO2_A = `SAY:
- Let us check. Five and five - the frames match
- Count all - that makes ten. Double 5 is 10
- Tick or fix

DO:
- Reveal double 5 is 10
- Count the ten counters, five then five

TEACHER NOTES:
Double 5 is 10 is a useful anchor double for later mental maths.

WATCH FOR:
- Children who said 9 or 11 - a miscount; recount the second frame

[Stage 3: We Do Answer | VTLM 2.0: Scaffold Practice]`;

const NOTES_HINGE_Q = `SAY:
- Big think now. Our largest double today - double 6
- Here are two groups of six. Build it in your head, or with counters
- What is double 6? Write the number on your whiteboard

DO:
- Display two groups of 6
- Give thinking time, allow counters for those who need them
- Signal show me - whiteboards up

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "What is double 6? Write the number. Show me"
- Scan for: the number 12 on most boards
PROCEED: If 80 percent or more write 12, move to You Do.
PIVOT: Most likely misconception - children write 6, they doubled nothing, or miscount past ten. Reteach with the double ten frame, 6 and 6, and count all slowly past ten - eleven, twelve. Re-check with double 6 again on fingers and frames.

TEACHER NOTES:
This is the hinge. It tells you who can count two equal groups all the way past ten.

WATCH FOR:
- Children who write 6 - they kept one group; show both groups again
- Children who write 12 confidently - they are ready for independent work

[Stage 3: CFU Hinge | VTLM 2.0: Monitor Progress]`;

const NOTES_HINGE_A = `SAY:
- The answer is 12. Six and six more makes twelve
- We counted all the way past ten - ten, eleven, twelve
- Tick or fix your whiteboard

DO:
- Reveal double 6 is 12
- Count all twelve dots together, slowly past ten

TEACHER NOTES:
If many missed it, rebuild 6 and 6 on the double ten frame before You Do.

WATCH FOR:
- Children still unsure past ten - book a short small group during You Do

[Stage 3: CFU Hinge Answer | VTLM 2.0: Monitor Progress]`;

const NOTES_YOUDO = `SAY:
- Time for your own work. Take your Doubles Practice sheet
- First - build two equal groups
- Next - count all the dots
- Then - write the double
- Show me double 3 and double 6 on your sheet

DO:
- Hand out the Doubles Practice sheet and counters
- Children work at their desks
- Circulate and check the two groups are equal before they count
- Give early finishers the Doubles Challenge mat

TEACHER NOTES:
You Do uses double 3 and double 6, different from the We Do doubles of 2 and 5. The sheet covers doubles 1 to 6. The middle success criterion, finding how many a double makes, is the focus.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Work with the teacher using counters and a five frame for each group; read, build, then write.
- Extra Notes: No separate sheet needed - the same sheet with support is enough.
EXTENDING PROMPT:
- Task: Hand the Doubles Challenge mat - prove a double two ways and invent your own double.
- Extra Notes: Distribute the Doubles Challenge PDF when ready.

WATCH FOR:
- Children who make unequal groups - prompt them to match before counting
- Children who finish fast and correctly - move them to the Challenge mat

[Stage 4: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EXIT = `SAY:
- Exit ticket - one quick question
- Here are two groups of four. What is double 4?
- Draw it if it helps, then write the number

DO:
- Display the exit ticket
- Children write the answer to double 4
- Collect the tickets on the way out

TEACHER NOTES:
The exit ticket targets the core success criterion - finding how many a double makes. Sort tickets into three piles - 8 correct, 4 (kept one group), other. Plan tomorrow accordingly.

WATCH FOR:
- Tickets with 4 - the child kept one group; revisit two equal groups tomorrow
- Tickets with 8 - secure with finding doubles

[Stage 5: Exit Ticket | VTLM 2.0: Monitor Progress]`;

const NOTES_CLOSING = `SAY:
- Let us check our success criteria
- I can make two equal groups - thumbs up, sideways or down
- I can find how many a double makes - thumbs
- I can say a double fact - thumbs
- Tell your partner one double you can say now, like double 4 is 8

DO:
- Run a thumbs check for each success criterion
- Give 20 seconds partner talk
- Cold call one child to share a double fact

TEACHER NOTES:
Tie the reflection to a concrete double the child can now say. Note any thumbs down on the middle criterion for tomorrow.

WATCH FOR:
- Children with thumbs down on finding a double - small group tomorrow
- Children sharing a fact confidently - ready to extend to doubles past 6

[General: Closing | VTLM 2.0: Review and Reflect]`;

// ─── Build ─────────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres,
    LESSON_TITLE,
    "Two equal groups, counted together",
    "Year 1 Numeracy | Doubles to double 6",
    NOTES_TITLE);

  // Slide 2: Teacher Resources (immediately after title, megaprompt §44)
  addResourceSlide(pres, {
    resources: RESOURCE_ITEMS,
    manipulatives: [
      "Double ten frames (1 per pair)",
      "Counters (about 25 per pair)",
      "Dot / dice cards 1 to 6",
      "Fingers",
    ],
    studentTools: ["Mini-whiteboards and markers"],
    boardSetup: [
      "Two large ten frames drawn side by side",
      "A pile of counters for modelling",
    ],
  }, T, FOOTER, NOTES_RESOURCES);

  // ── Daily Review (counting on / addition within 10) ──
  // Slide 3-4: DR1 — count on 3 from 6
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Start at 6. Count on 3.", { color: STAGE_COLORS["1"] });
      addNumberTrack(s, 0.5, 2.3, 9, 1, 10, [6], { cellH: 0.82, fontSize: 26 });
      addInstructionCard(s, [
        { role: "header", text: "Count on" },
        { role: "body", text: "Start at 6. Count on 3 more. Where do you land?" },
      ], { x: 0.5, y: 3.45, w: 9, h: 1.0, strip: C.ACCENT });
      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR1_Q);
      return s;
    },
    (slide) => {
      answerBar(slide, "9   —   6 then 7, 8, 9", { x: 2.0, y: 4.55, w: 6.0 });
      slide.addNotes(NOTES_DR1_A);
    }
  );

  // Slide 5-6: DR2 — 4 + 3
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "How many altogether?", { color: STAGE_COLORS["1"] });
      drawGroups(s, 4, 3, { y: 1.8, size: 1.6 });
      addInstructionCard(s, [
        { role: "header", text: "Count them all" },
        { role: "body", text: "4 and 3 more. How many altogether?" },
      ], { x: 0.5, y: 3.5, w: 9, h: 0.78, strip: C.ACCENT });
      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR2_Q);
      return s;
    },
    (slide) => {
      answerBar(slide, "7", { x: 4.0, y: 4.55, w: 2.0 });
      slide.addNotes(NOTES_DR2_A);
    }
  );

  // Slide 7-8: DR3 — 5 + 4
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "How many altogether?", { color: STAGE_COLORS["1"] });
      drawGroups(s, 5, 4, { y: 1.8, size: 1.6 });
      addInstructionCard(s, [
        { role: "header", text: "Count them all" },
        { role: "body", text: "5 and 4 more. How many altogether?" },
      ], { x: 0.5, y: 3.5, w: 9, h: 0.78, strip: C.ACCENT });
      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR3_Q);
      return s;
    },
    (slide) => {
      answerBar(slide, "9", { x: 4.0, y: 4.55, w: 2.0 });
      slide.addNotes(NOTES_DR3_A);
    }
  );

  // ── Fluency (subitising dot patterns to 6) ──
  function fluencyFlash(count, qNotes, aNotes, answerText) {
    withReveal(
      () => {
        const s = pres.addSlide();
        addTopBar(s, STAGE_COLORS["1"]);
        addStageBadge(s, 1, "Fluency");
        addTitle(s, "How many dots?", { color: STAGE_COLORS["1"] });
        addDotCard(s, 3.95, 1.5, 2.1, count, { dotColor: C.PRIMARY, borderColor: C.CHARCOAL });
        addInstructionCard(s, [
          { role: "header", text: "Quick eyes!" },
          { role: "body", text: "No counting one by one. Just know it." },
        ], { x: 0.5, y: 3.7, w: 9, h: 0.62, strip: C.ACCENT });
        addFooter(s, FOOTER);
        s.addNotes(qNotes);
        return s;
      },
      (slide) => {
        answerBar(slide, answerText, { x: 4.0, y: 4.5, w: 2.0, h: 0.5 });
        slide.addNotes(aNotes);
      }
    );
  }
  // Slide 9-10: Fluency 1 — 4
  fluencyFlash(4, NOTES_FL1_Q, NOTES_FL1_A, "4");
  // Slide 11-12: Fluency 2 — 6
  fluencyFlash(6, NOTES_FL2_Q, NOTES_FL2_A, "6");
  // Slide 13-14: Fluency 3 — 5
  fluencyFlash(5, NOTES_FL3_Q, NOTES_FL3_A, "5");

  // Slide 15: Launch
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["1"]);
    addStageBadge(s, 1, "Launch");
    addTitle(s, "3 fingers on each hand", { color: STAGE_COLORS["1"] });
    drawGroups(s, 3, 3, { y: 1.85, size: 1.75 });
    addInstructionCard(s, [
      { role: "header", text: "Show your partner" },
      { role: "body", text: "3 and 3. How many fingers altogether?" },
    ], { x: 0.5, y: 3.8, w: 9, h: 0.9, strip: C.ACCENT });
    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // Slide 16: LI/SC
  liSlide(pres,
    ["We are learning to find doubles to double 6."],
    [
      "I can make two equal groups.",
      "I can find how many a double makes.",
      "I can say a double fact.",
    ],
    NOTES_LISC, FOOTER);

  // Slide 17: Vocabulary — double
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Word", { color: C.PRIMARY });
    addTitle(s, "Double", { color: C.PRIMARY });
    const g = drawGroups(s, 3, 3, { y: 1.85, size: 1.75 });
    groupLabel(s, g.x, g.size, g.y + g.size + 0.06, "3", C.PRIMARY);
    groupLabel(s, g.x + g.size + g.gap, g.size, g.y + g.size + 0.06, "3", C.SECONDARY);
    addInstructionCard(s, [
      { role: "header", text: "A double is two groups that are the same." },
      { role: "body", text: "Here are 3 and 3. The groups match." },
    ], { x: 0.5, y: 4.05, w: 9, h: 0.95, strip: C.ACCENT });
    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // Slide 18: I Do 1 — double 3 (dot groups)
  workedExSlide(pres, 2, "I Do", "Double 3",
    [
      "A double is two equal groups.",
      "Here are 3 and 3.",
      "I count all: 6.",
      "Double 3 is 6.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.25, { strip: C.PRIMARY });
      const cardSize = 1.35;
      const y1 = lg.panelTopPadded + 0.35;
      addDotCard(slide, lg.rightX + 0.35, y1, cardSize, 3, { dotColor: C.PRIMARY, borderColor: C.CHARCOAL });
      slide.addText("+", {
        x: lg.rightX + 0.35 + cardSize, y: y1, w: 0.6, h: cardSize,
        fontSize: 34, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      addDotCard(slide, lg.rightX + 0.35 + cardSize + 0.6, y1, cardSize, 3, { dotColor: C.SECONDARY, borderColor: C.CHARCOAL });
      addTextOnShape(slide, "Double 3 is 6", {
        x: lg.rightX + 0.3, y: y1 + cardSize + 0.45, w: lg.rightW - 0.6, h: 0.6, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 19: I Do 2 — double 4 (double ten frame)
  workedExSlide(pres, 2, "I Do", "Double 4",
    [
      "Two equal groups again.",
      "4 in this frame, 4 in that frame.",
      "I count all: 8.",
      "Double 4 is 8.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.5, { strip: C.PRIMARY });
      const frameW = 3.0;
      const frameX = lg.rightX + (lg.rightW - frameW) / 2;
      addTensFrame(slide, frameX, lg.panelTopPadded + 0.28, frameW, 4, { cellH: 0.6, fillColor: C.PRIMARY });
      addTensFrame(slide, frameX, lg.panelTopPadded + 1.60, frameW, 4, { cellH: 0.6, fillColor: C.SECONDARY });
      addTextOnShape(slide, "Double 4 is 8", {
        x: lg.rightX + 0.3, y: lg.panelTopPadded + 2.92, w: lg.rightW - 0.6, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 20-21: CFU 1 — which one is a double?
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "Which one is a double?", { color: C.ALERT });

      // Option A: 3 and 3 (a double)
      addCard(s, 0.6, 1.5, 4.0, 2.0, { strip: C.SECONDARY });
      addDotCard(s, 1.05, 1.78, 1.25, 3, { dotColor: C.PRIMARY, borderColor: C.CHARCOAL });
      slide_plus(s, 2.3, 1.78, 0.5, 1.25);
      addDotCard(s, 2.8, 1.78, 1.25, 3, { dotColor: C.SECONDARY, borderColor: C.CHARCOAL });
      groupLabel(s, 0.6, 4.0, 3.12, "A", C.CHARCOAL);

      // Option B: 4 and 2 (not a double)
      addCard(s, 5.4, 1.5, 4.0, 2.0, { strip: C.PRIMARY });
      addDotCard(s, 5.85, 1.78, 1.25, 4, { dotColor: C.PRIMARY, borderColor: C.CHARCOAL });
      slide_plus(s, 7.1, 1.78, 0.5, 1.25);
      addDotCard(s, 7.6, 1.78, 1.25, 2, { dotColor: C.SECONDARY, borderColor: C.CHARCOAL });
      groupLabel(s, 5.4, 4.0, 3.12, "B", C.CHARCOAL);

      addInstructionCard(s, [
        { role: "header", text: "Hold up A or B" },
        { role: "body", text: "Which one shows two equal groups?" },
      ], { x: 0.5, y: 3.6, w: 9, h: 0.7, strip: C.ALERT });
      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1_Q);
      return s;
    },
    (slide) => {
      answerBar(slide, "A — 3 and 3 are the same!", { x: 0.6, y: 4.45, w: 4.5, h: 0.5, fontSize: 20 });
      slide.addNotes(NOTES_CFU1_A);
    }
  );

  // Slide 22-23: We Do 1 — build double 2
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Build double 2 together", { color: STAGE_COLORS["3"] });
      drawGroups(s, 2, 2, { y: 1.8, size: 1.6 });
      addInstructionCard(s, [
        { role: "header", text: "Build it with counters" },
        { role: "body", text: "2 and 2. Count all. How many is double 2?" },
      ], { x: 0.5, y: 3.5, w: 9, h: 0.78, strip: C.SECONDARY });
      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO1_Q);
      return s;
    },
    (slide) => {
      answerBar(slide, "Double 2 is 4", { x: 2.5, y: 4.5, w: 5.0, h: 0.5 });
      slide.addNotes(NOTES_WEDO1_A);
    }
  );

  // Slide 24-25: We Do 2 — double 5 on double ten frame
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Build double 5 together", { color: STAGE_COLORS["3"] });
      const frameW = 3.0;
      const frameX = (10 - frameW) / 2;
      addTensFrame(s, frameX, 1.4, frameW, 5, { cellH: 0.5, fillColor: C.PRIMARY });
      addTensFrame(s, frameX, 2.5, frameW, 5, { cellH: 0.5, fillColor: C.SECONDARY });
      addInstructionCard(s, [
        { role: "header", text: "Show it on two ten frames" },
        { role: "body", text: "5 and 5. Count all. How many is double 5?" },
      ], { x: 0.5, y: 3.6, w: 9, h: 0.72, strip: C.SECONDARY });
      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO2_Q);
      return s;
    },
    (slide) => {
      answerBar(slide, "Double 5 is 10", { x: 2.5, y: 4.5, w: 5.0, h: 0.5 });
      slide.addNotes(NOTES_WEDO2_A);
    }
  );

  // Slide 26-27: CFU 2 (hinge) — double 6
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU - Hinge", { color: C.ALERT });
      addTitle(s, "What is double 6?", { color: C.ALERT });
      drawGroups(s, 6, 6, { y: 1.75, size: 1.6 });
      addInstructionCard(s, [
        { role: "header", text: "Write the number" },
        { role: "body", text: "6 and 6. Count all. What is double 6?" },
      ], { x: 0.5, y: 3.5, w: 9, h: 0.78, strip: C.ALERT });
      addFooter(s, FOOTER);
      s.addNotes(NOTES_HINGE_Q);
      return s;
    },
    (slide) => {
      answerBar(slide, "Double 6 is 12", { x: 2.5, y: 4.5, w: 5.0, h: 0.5 });
      slide.addNotes(NOTES_HINGE_A);
    }
  );

  // Slide 28: You Do
  workedExSlide(pres, 4, "You Do", "Make doubles",
    [
      "First - build two equal groups.",
      "Next - count all the dots.",
      "Then - write the double.",
      "",
      "Show double 3 and double 6.",
    ],
    NOTES_YOUDO, FOOTER,
    (slide, lg) => {
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, 3.0, { strip: C.ALERT });
      slide.addText("Remember", {
        x: lg.rightX + 0.25, y: lg.panelTopPadded + 0.16, w: lg.rightW - 0.5, h: 0.38,
        fontSize: 19, fontFace: FONT_H, color: C.ALERT, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Make the two groups the same.", options: { bullet: true, fontSize: 14.5, color: C.CHARCOAL, breakLine: true } },
        { text: "Count every dot.", options: { bullet: true, fontSize: 14.5, color: C.CHARCOAL, breakLine: true } },
        { text: "Write the double you made.", options: { bullet: true, fontSize: 14.5, color: C.CHARCOAL } },
      ], {
        x: lg.rightX + 0.35, y: lg.panelTopPadded + 0.58, w: lg.rightW - 0.6, h: 1.05,
        fontFace: FONT_B, margin: 0, valign: "top",
      });
      // Small worked reminder using double 2 (not a You Do task) so the slide
      // carries a visual model without giving away double 3 or double 6.
      const miniSize = 0.78;
      const miniGap = 0.45;
      const miniX = lg.rightX + (lg.rightW - (miniSize * 2 + miniGap)) / 2;
      const miniY = lg.panelTopPadded + 1.8;
      addDotCard(slide, miniX, miniY, miniSize, 2, { dotColor: C.PRIMARY, borderColor: C.CHARCOAL });
      slide.addText("+", {
        x: miniX + miniSize, y: miniY, w: miniGap, h: miniSize,
        fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true, align: "center", valign: "middle", margin: 0,
      });
      addDotCard(slide, miniX + miniSize + miniGap, miniY, miniSize, 2, { dotColor: C.SECONDARY, borderColor: C.CHARCOAL });
      slide.addText("Double 2 is 4", {
        x: lg.rightX + 0.25, y: miniY + miniSize + 0.06, w: lg.rightW - 0.5, h: 0.3,
        fontSize: 15, fontFace: FONT_H, color: C.ALERT, bold: true, align: "center", margin: 0,
      });
    }
  );

  // Slide 29: Exit Ticket
  exitTicketSlide(pres,
    ["Draw two equal groups of 4. What is double 4? Write the number."],
    NOTES_EXIT, FOOTER);

  // Slide 30: Closing
  closingSlide(pres, {
    reflectionPrompt: "Tell your partner one double you can say now, like double 4 is 8.",
    scItems: [
      "I can make two equal groups.",
      "I can find how many a double makes.",
      "I can say a double fact.",
    ],
    selfAssessment: "Thumbs up, sideways or down for each one",
  }, NOTES_CLOSING);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "G1D_Doubles_to_Double6.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Companion PDFs ──────────────────────────────────────────────────────────

  // Draw a large "Double N" row: two draw-boxes (+) and an answer box (=).
  function drawDoubleRow(doc, n, y) {
    const x = 56;
    const boxW = 88, boxH = 66, gap = 24, labelW = 70, writeW = 78;
    doc.fontSize(15).font("Sans-Bold").fillColor("#000000");
    doc.text("Double " + n, x, y + boxH / 2 - 9, { width: labelW });
    let cx = x + labelW + 6;
    doc.roundedRect(cx, y, boxW, boxH, 6).lineWidth(1.4).strokeColor(hex(C.NAVY)).stroke();
    cx += boxW;
    doc.fontSize(26).font("Sans-Bold").fillColor("#000000").text("+", cx, y + boxH / 2 - 15, { width: gap, align: "center" });
    cx += gap;
    doc.roundedRect(cx, y, boxW, boxH, 6).lineWidth(1.4).strokeColor(hex(C.TEAL)).stroke();
    cx += boxW;
    doc.fontSize(26).font("Sans-Bold").fillColor("#000000").text("=", cx, y + boxH / 2 - 15, { width: gap, align: "center" });
    cx += gap;
    doc.roundedRect(cx, y, writeW, boxH, 6).lineWidth(1.6).strokeColor("#000000").stroke();
    return y + boxH + 26;
  }

  // Worksheet — build, count and write doubles
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Build, count and write the double",
      color: C.NAVY,
      lessonInfo: "Year 1 Numeracy | Doubles to double 6",
    });
    y = addTipBox(doc, "Draw the dots in both boxes. Make them the SAME. Count all the dots. Write how many in the last box.", y, { color: C.TEAL });
    y = addSectionHeading(doc, "Make each double", y, { color: C.NAVY });
    y += 6;
    [3, 4, 5, 6].forEach((n) => { y = drawDoubleRow(doc, n, y); });
    addPdfFooter(doc, "Doubles to Double 6 | Year 1 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // Answer Key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Build, count and write the double - ANSWERS",
      color: C.NAVY,
      lessonInfo: "Year 1 Numeracy | Doubles to double 6",
      showNameDate: false,
    });
    y = addSectionHeading(doc, "Answers", y, { color: C.NAVY });
    y = addBodyText(doc, "Double 3 = 6   (3 and 3)", y);
    y = addBodyText(doc, "Double 4 = 8   (4 and 4)", y);
    y = addBodyText(doc, "Double 5 = 10   (5 and 5)", y);
    y = addBodyText(doc, "Double 6 = 12   (6 and 6)", y);
    y += 6;
    y = addBodyText(doc, "Watch for: children who write the same number as one group (for example 3 for double 3). They kept one group instead of joining two equal groups. Rebuild both groups and count all.", y);
    addPdfFooter(doc, "Answer Key | Year 1 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  // Extension — Doubles Challenge mat
  await (async () => {
    const doc = createPdf({ title: EXTENDING_RES.name });
    let y = addPdfHeader(doc, EXTENDING_RES.name, {
      subtitle: "Prove a double, then make your own",
      color: C.TEAL,
      lessonInfo: "Year 1 Numeracy | Doubles to double 6",
    });
    y = addTipBox(doc, "Finished early? Be a doubles expert!", y, { color: C.TEAL });

    y = addSectionHeading(doc, "Show double 4 two ways", y, { color: C.NAVY });
    doc.fontSize(12).font("Sans").fillColor("#000000");
    doc.text("Way 1: draw two groups of dots", 56, y);
    doc.roundedRect(56, y + 18, 220, 90, 6).lineWidth(1.4).strokeColor(hex(C.NAVY)).stroke();
    doc.text("Way 2: draw it on the ten frames", 300, y);
    // two small ten frames stacked
    const tfX = 300, tfY = y + 18, cell = 18;
    for (let f = 0; f < 2; f += 1) {
      for (let r = 0; r < 2; r += 1) {
        for (let c = 0; c < 5; c += 1) {
          doc.rect(tfX + c * cell, tfY + f * (cell * 2 + 8) + r * cell, cell, cell)
            .lineWidth(1).strokeColor("#444444").stroke();
        }
      }
    }
    y = y + 18 + 90 + 24;

    y = addSectionHeading(doc, "Make your own double", y, { color: C.NAVY });
    doc.fontSize(12).font("Sans").fillColor("#000000");
    doc.text("Pick a number. Draw two equal groups. Write the double.", 56, y);
    doc.roundedRect(56, y + 20, 483, 120, 6).lineWidth(1.4).strokeColor(hex(C.TEAL)).stroke();
    doc.fontSize(14).font("Sans-Bold").fillColor("#000000");
    doc.text("Double ____ = ____", 70, y + 150);

    addPdfFooter(doc, "Doubles Challenge | Year 1 Numeracy");
    await writePdf(doc, path.join(OUT_DIR, EXTENDING_RES.fileName));
    console.log("PDF written: " + EXTENDING_RES.fileName);
  })();
}

// Small "+" between two groups on a manual slide.
function slide_plus(slide, x, y, w, h) {
  slide.addText("+", {
    x, y, w, h,
    fontSize: 30, fontFace: FONT_H, color: C.CHARCOAL,
    bold: true, align: "center", valign: "middle", margin: 0,
  });
}

build().catch((err) => { console.error(err); process.exit(1); });
