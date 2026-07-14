"use strict";

// Foundation Numeracy — Lesson 1: Making 10 using ten frames (friends of 10)
// Victorian Curriculum 2.0: Mathematics, Number, Foundation — represent and count
//   quantities to 10; connect number names and quantities; make and name pairs
//   that combine to 10.
// First lesson on friends of 10. Concrete first: count the empty boxes to find how
//   many more make 10, then name the two numbers ("6 and 4 make 10"). Ten frames
//   + counters are the anchor. No + / = symbols on student faces (megaprompt 31a).
// Daily Review: counting collections to 10 and reading numerals.
// Fluency:      subitising with dot cards.
// 45-minute lesson, mixed readiness.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const { composeGlanceNotes } = require("../themes/core/composeNotes");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addWriteLine,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  addTenFramePdf, hex,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "foundation", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, closingSlide,
  workedExSlide, addStageBadge,
  addRevealAnswerBar,
  addTensFrame, addDotCard, addGroupedCounters,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  addInstructionCard,
  withReveal, runSlideDiagnostics,
  STAGE_COLORS,
} = T;

const SESSION = 1;
const FOOTER = "Making 10 | Friends of 10 | Foundation Numeracy";
const OUT_DIR = "output/MK10_Lesson1_Making_10_Ten_Frames";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Make 10 Worksheet",
  "Draw the counters that fill each ten frame to 10, then write how many more.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Make 10 Answer Key",
  "Teacher answers for the Make 10 worksheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

// Counter colour used consistently for all "filled" cells so the frame reads
// the same way on every slide (megaprompt 68j: repeated visual structure).
const COUNTER = C.SECONDARY;

// ─── Teacher notes (Glance Format, megaprompt sections 45-47) ────────────────

const NOTES_TITLE =
  "Open the lesson on the mat. Say the title, then move to Teacher Resources.";

const NOTES_RESOURCES =
  "Prep slide. Print the Make 10 worksheet and answer key. Have a big ten frame " +
  "on the board, tubs of counters and a ten frame each, plus mini-whiteboards ready.";

const NOTES_DR1 = composeGlanceNotes({
  answer: "7 counters.",
  beats: [
    "POINT to the counters. SAY: Daily Review. Let us warm up our counting.",
    "ASK: How many counters? 5 sec, then show me on your fingers. EXPECT: 7 fingers.",
    "SCAN the carpet. 80%+ -> reveal. Less -> touch and count each counter together, re-ask.",
  ],
  trap: "skipping or double-counting a counter. Fix: touch each one once, child recounts.",
  prep: "Prior learning only: counting a collection to 10. Bridges into today's full ten frame of 10.",
  tag: "[Stage 1 | Daily Review | Retention and recall | HITS 6]",
});

const NOTES_DR2 = composeGlanceNotes({
  answer: "nine.",
  beats: [
    "POINT to the numeral. SAY: This is how we write a number. Read it in your head.",
    "ASK: What number is this? 5 sec, choral on my signal. EXPECT: nine.",
    "SCAN for the choral answer. 80%+ -> reveal. Less -> count 1 to 9 together and land on 9, re-ask.",
  ],
  trap: "reading 9 as 6. Fix: point to the tail going down, say nine together, child re-reads.",
  prep: "Prior learning only: reading numerals to 10. Keeps number-name to numeral link warm.",
  tag: "[Stage 1 | Daily Review | Retention and recall | HITS 6]",
});

const NOTES_DR3 = composeGlanceNotes({
  answer: "10 - ten. A full group of ten.",
  beats: [
    "POINT to the counters. SAY: One more count. This time count all the way.",
    "ASK: How many counters? 6 sec, then whisper it to your partner. EXPECT: ten.",
    "SCAN partner talk. 80%+ -> reveal. Less -> count together slowly, tap each counter, re-ask.",
  ],
  trap: "stopping at 9 or losing track past 5. Fix: count the top row 5, then keep going, child recounts.",
  prep: "Prior learning only: counting to 10. Names 10 out loud - the target we build toward today.",
  tag: "[Stage 1 | Daily Review | Retention and recall | HITS 6]",
});

const NOTES_FL1 = composeGlanceNotes({
  answer: "4 dots.",
  beats: [
    "FLASH the dot card. SAY: Fluency. Quick eyes - do not count one by one.",
    "ASK: How many dots? 3 sec, show me on your fingers. EXPECT: 4. ACCEPT: I saw two and two.",
    "SCAN fingers, back row first. 80%+ -> reveal. Less -> name the pattern, re-flash and re-ask.",
  ],
  trap: "counting each dot slowly. Fix: SAY the pattern - two and two is four, child says it back.",
  prep: "Subitising - see the amount at a glance without counting (Number automaticity).",
  tag: "[Stage 1 | Fluency | Mastery and application | HITS 6]",
});

const NOTES_FL2 = composeGlanceNotes({
  answer: "6 dots.",
  beats: [
    "FLASH the dot card. SAY: Next one. Quick eyes.",
    "ASK: How many dots? 3 sec, show me on your fingers. EXPECT: 6. ACCEPT: three and three.",
    "SCAN fingers. 80%+ -> reveal. Less -> cover one column, name three and three, re-flash and re-ask.",
  ],
  trap: "guessing a number near six. Fix: SAY three and three is six, child rebuilds on fingers.",
  prep: "Subitising the dice-six pattern. Builds instant recognition of small amounts.",
  tag: "[Stage 1 | Fluency | Mastery and application | HITS 6]",
});

const NOTES_FL3 = composeGlanceNotes({
  answer: "5 dots.",
  beats: [
    "FLASH the dot card. SAY: Last one. Quick eyes.",
    "ASK: How many dots? 3 sec, show me on your fingers. EXPECT: 5. ACCEPT: four and one.",
    "SCAN fingers. 80%+ -> reveal. Less -> point to the four corners and the middle, re-flash and re-ask.",
  ],
  trap: "showing four (missing the middle dot). Fix: point to the centre dot, child re-shows five.",
  prep: "Subitising five - the anchor amount for one full row of the ten frame.",
  tag: "[Stage 1 | Fluency | Mastery and application | HITS 6]",
});

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "10 - a full ten frame is always 10.",
  beats: [
    "POINT to the full frame. SAY: This is our ten frame. Every box has a counter.",
    "ASK: How many counters when the frame is full? 5 sec, choral. EXPECT: ten.",
    "SAY: Today some boxes will be empty. Our job is to find how many more make 10.",
  ],
  trap: "recounting a frame students already know is full. Fix: SAY a full frame is 10, no need to count.",
  prep: "Launch bridges 'full frame = 10' (prior) to today's 'how many more make 10'. New concept - build together.",
  tag: "[Launch | Attention and knowledge | HITS 2, 3]",
});

const NOTES_LI_SC = composeGlanceNotes({
  beats: [
    "POINT to the learning intention. SAY: Today we learn to make 10 using a ten frame.",
    "SAY: Read the I can statements with me. Choral read, point to each line.",
    "SAY: The first one everyone can do - count the counters in the frame.",
  ],
  prep: "SC1 achievable by all; SC2 is the core target and the exit ticket; SC3 stretches to naming the pair.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_VOCAB = composeGlanceNotes({
  answer: "friends of 10 are two numbers that make 10 together. Here it is 5 and 5.",
  beats: [
    "POINT to the frame. SAY: New maths words - friends of 10.",
    "SAY: The top row is 5. The empty row is 5 more. 5 and 5 make 10 - they are friends of 10.",
    "ASK: Say it with me - 5 and 5 make 10. 5 sec, choral on my signal. EXPECT: 5 and 5 make 10.",
  ],
  prep: "Key language for the unit. One clean symmetric pair (5 and 5) before the I Do uses other pairs.",
  tag: "[Vocabulary | Knowledge and memory | HITS 3]",
});

const NOTES_IDO1 = composeGlanceNotes({
  answer: "4 more. 6 and 4 make 10.",
  beats: [
    "SHOW the frame with 6 counters. SAY: Watch me. First I count how many are here - 6.",
    "MODEL counting the empty boxes. SAY: Now I count the empty boxes - 1, 2, 3, 4. Four more make 10.",
    "SAY: So 6 and 4 make 10. Six and four are friends of 10.",
    "ASK: How many empty boxes? 5 sec, choral. EXPECT: four.",
  ],
  trap: "counting the filled boxes instead of the empty ones. Fix: hand over the counters, count only the gaps, child recounts.",
  stretch: "cover the frame - can you say the friend of 10 for 6 without looking?",
  help: "give the child their own frame and 6 counters, they fill it and count the empties.",
  prep: "First model - meaning before symbols. Count the empty boxes to find how many more make 10.",
  tag: "[Stage 2 | I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_IDO2 = composeGlanceNotes({
  answer: "2 more. 8 and 2 make 10.",
  beats: [
    "SHOW the frame with 8 counters. SAY: Again, watch how I do it. I count what is here - 8.",
    "MODEL counting the gaps. SAY: I count the empty boxes - 1, 2. Two more make 10.",
    "SAY: So 8 and 2 make 10. Eight and two are friends of 10.",
    "ASK: How many more make 10? 5 sec, choral. EXPECT: two.",
  ],
  trap: "saying 8 (the counters seen) as the answer. Fix: point to the two gaps, count them, child re-answers.",
  stretch: "which is bigger - the counters or the empty boxes? Say why.",
  help: "child builds 8 on their own frame, then counts the empty boxes.",
  prep: "Second model, same routine, near-full frame. Repeated structure so students focus on the maths.",
  tag: "[Stage 2 | I Do | Explicit teaching | HITS 3, 4]",
});

const NOTES_CFU1 = composeGlanceNotes({
  answer: "3 more. 7 and 3 make 10.",
  beats: [
    "POINT to the frame with 7 counters. SAY: Quick check. Count the empty boxes.",
    "ASK: How many more make 10? 6 sec, write it on your board. EXPECT: 3. ACCEPT: I counted 3 empty boxes.",
    "SCAN boards, back row first. 80%+ -> reveal. Less -> count the gaps together on the board frame, re-ask.",
  ],
  trap: "writing 7 (counted the counters). Fix: finger on each empty box, count only gaps, child rewrites.",
  stretch: "what is the friend of 10 for 7? Say the pair.",
  help: "point to the first empty box to start the child counting.",
  prep: "Checks the count-the-empty-boxes strategy before guided practice. Decision point.",
  tag: "[Stage 2 | CFU | Supported application | HITS 7, 8]",
});

const NOTES_WEDO1 = composeGlanceNotes({
  answer: "6 more. 4 and 6 make 10.",
  beats: [
    "POINT to the frame with 4 counters. SAY: Your turn with me. Count the empty boxes together.",
    "ASK: How many more make 10? 8 sec, boards up. EXPECT: 6. ACCEPT: 6 empty boxes.",
    "SCAN boards. 80%+ -> reveal. Less -> count the gaps aloud together, tap each one, re-ask.",
  ],
  trap: "losing count over two rows of gaps. Fix: count the gaps in the bottom row, then the top, child recounts.",
  stretch: "say the whole pair - 4 and 6 make 10.",
  help: "count the first two empty boxes with the child, then let them finish.",
  prep: "Guided practice, more empty than full so the count matters. Faded from the I Do.",
  tag: "[Stage 3 | We Do | Supported application | HITS 4, 5]",
});

const NOTES_WEDO2 = composeGlanceNotes({
  answer: "1 more. 9 and 1 make 10.",
  beats: [
    "POINT to the frame with 9 counters. SAY: Together again. Nearly full this time.",
    "ASK: How many more make 10? 6 sec, boards up. EXPECT: 1. ACCEPT: just one empty box.",
    "SCAN boards. 80%+ -> reveal. Less -> point to the single gap, count one, re-ask.",
  ],
  trap: "writing 9. Fix: show the one empty box, SAY only one more, child rewrites 1.",
  stretch: "say the pair - 9 and 1 make 10.",
  help: "guide the child's finger to the single empty box.",
  prep: "Guided practice with a near-ten frame - the friend of 10 is 1. Common and useful.",
  tag: "[Stage 3 | We Do | Supported application | HITS 4, 5]",
});

const NOTES_WEDO3 = composeGlanceNotes({
  answer: "7 more. 3 and 7 make 10.",
  beats: [
    "SAY: Now you build it. Put 3 counters in the top of your own ten frame.",
    "SAY: Fill the rest of your frame to make 10. Count the counters you add.",
    "ASK: How many more did you add? 15 sec, boards up. EXPECT: 7. ACCEPT: 3 and 7 make 10.",
    "SCAN boards and frames. 80%+ -> reveal. Less -> build it together on the board frame, re-ask.",
  ],
  trap: "leaving a gap or overfilling past 10. Fix: count all boxes to 10, no more, child fixes their frame.",
  stretch: "build 10 a different way and say your new friends of 10.",
  help: "hand the child a frame with 3 counters already placed, they fill the rest.",
  prep: "Students use their own frame and counters (kinaesthetic build). Same skill, now hands-on.",
  tag: "[Stage 3 | We Do | Supported application | HITS 4, 5]",
});

const NOTES_CFU2 = composeGlanceNotes({
  answer: "8 more. 2 and 8 make 10.",
  beats: [
    "POINT to the frame with 2 counters. SAY: Hinge check. Only two counters - lots of empty boxes.",
    "ASK: How many more make 10? 8 sec, boards up. EXPECT: 8. ACCEPT: 8 empty boxes.",
    "SCAN boards, back row first. 80%+ -> You Do. Less -> count every empty box together, re-ask.",
  ],
  trap: "writing 2 (counted the counters, not the gaps). Fix: count only the empty boxes, child rewrites 8.",
  stretch: "say the pair - 2 and 8 make 10.",
  help: "count the empty boxes in one row first, then the other.",
  prep: "Hinge before independent work. Few counters, many gaps - the clearest test of the strategy.",
  tag: "[Stage 3 | CFU | Supported application | HITS 7, 8]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "Frame 1 (6): 4 more. Frame 2 (8): 2 more. Frame 3 (5): 5 more.",
  beats: [
    "COLLECT eyes. SAY: On your own now. Draw the counters to fill each frame to 10.",
    "SAY: Count the empty boxes, draw a counter in each, then write how many more.",
    "CIRCULATE, back row first. Watch for children counting the drawn counters, not the gaps.",
  ],
  trap: "writing the number already in the frame. Fix: point to the gaps, count only those, child fixes it.",
  stretch: "the challenge box - make 10 a different way and say your friends of 10.",
  help: "sit with the small group, use real counters on a frame before drawing on paper.",
  prep: "Independent recording on new frames. Worksheet keeps the ten frame model on paper.",
  tag: "[Stage 4 | You Do | Mastery and application | HITS 4, 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "3 more. 7 and 3 make 10.",
  beats: [
    "POINT to the frame with 7 counters. SAY: Last job on your own board. How many more make 10?",
    "ASK: Write your number and hold your board up. 20 sec. EXPECT: 3.",
    "COLLECT boards or note who wrote what.",
  ],
  trap: "answer of 7 (counted counters). Note who for tomorrow's small group.",
  prep: "Assesses SC2 (find how many more make 10). SC number stays in notes, not on the slide.",
  tag: "[Stage 5 | Exit Ticket | Evidence of learning | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - listen for 'the empty boxes show how many more make 10'.",
  beats: [
    "POINT to the I can statements. SAY: Show me thumbs for each one - up, sideways or down.",
    "ASK: How do we find how many more make 10? Turn and tell. 15 sec. EXPECT: count the empty boxes.",
    "SAY: You found the friends of 10 by counting the empty boxes - that is our big idea today.",
  ],
  prep: "Revisits all three success criteria and the count-the-empty-boxes idea. Primes tomorrow's number bonds.",
  tag: "[Closing | Retention and reflection | HITS 9]",
});

// ─── Lesson targets ──────────────────────────────────────────────────────────

const LI = "We are learning to make 10 using a ten frame.";
const SC = [
  "I can count the counters in a ten frame.",
  "I can find how many more counters make 10.",
  "I can say the two numbers that make 10.",
];

// ─── Slide-side helpers ──────────────────────────────────────────────────────

// A big centred ten frame with `filled` counters, sized for the back of the room.
// Cells are square (cellH = cellW) so the counters render as true circles.
function heroTenFrame(slide, filled, opts) {
  const o = opts || {};
  const w = o.w != null ? o.w : 4.6;
  const x = o.x != null ? o.x : (10 - w) / 2;
  const cellH = o.cellH != null ? o.cellH : w / 5;
  addTensFrame(slide, x, o.y != null ? o.y : 1.6, w, filled,
    { cellH, fillColor: COUNTER });
}

// A slim centred routine cue (e.g. "Show me on your board") that leaves the
// bottom of the slide clear for the reveal bar.
function routinePill(slide, text, color) {
  addTextOnShape(slide, text, {
    x: 2.0, y: 3.7, w: 6.0, h: 0.5, rectRadius: 0.1,
    fill: { color: color || C.PRIMARY },
  }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
}

// A big numeral card for the "read the numeral" Daily Review.
function heroNumeral(slide, numeral) {
  slide.addShape("roundRect", {
    x: 3.7, y: 1.6, w: 2.6, h: 2.1, rectRadius: 0.16,
    fill: { color: C.WHITE }, line: { color: C.PRIMARY, width: 3 },
  });
  slide.addText(String(numeral), {
    x: 3.7, y: 1.6, w: 2.6, h: 2.1,
    fontSize: 150, fontFace: FONT_H, color: C.PRIMARY, bold: true,
    align: "center", valign: "middle", margin: 0,
  });
}

// ─── Build ───────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  // Slide 1: Title
  titleSlide(pres, "Making 10",
    "Friends of 10 with ten frames",
    "Foundation Numeracy | Number", NOTES_TITLE);

  // Slide 2: Teacher Resources
  addResourceSlide(pres, {
    resources: RESOURCE_ITEMS,
    manipulatives: [
      "A ten frame for every child",
      "Tubs of counters (about 12 per child)",
    ],
    studentTools: [
      "Mini-whiteboards and markers",
      "Printed Make 10 worksheet",
    ],
    boardSetup: [
      "A large ten frame drawn on the board to fill live",
      "Counters ready to place on the board frame",
    ],
  }, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // ── Daily Review: counting collections to 10 and reading numerals ──────────

  // Slide 3-4: DR1 - count the collection (7)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "How many counters?", { color: STAGE_COLORS["1"] });
      addGroupedCounters(s, 2.35, 2.4, 1, 7, { dot: 0.6, gap: 0.22 });
      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR1);
      return s;
    },
    (slide) => {
      addRevealAnswerBar(slide, "7 counters", { y: 4.3, h: 0.7, fontSize: 30 });
    }
  );

  // Slide 5-6: DR2 - read the numeral (9)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "Read this number", { color: STAGE_COLORS["1"] });
      heroNumeral(s, "9");
      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR2);
      return s;
    },
    (slide) => {
      addRevealAnswerBar(slide, "nine", { y: 4.3, h: 0.7, fontSize: 30 });
    }
  );

  // Slide 7-8: DR3 - count the collection (10)
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Daily Review");
      addTitle(s, "How many counters?", { color: STAGE_COLORS["1"] });
      // Two neat framed groups of 5 read clearly as ten without a ten frame.
      addGroupedCounters(s, 2.35, 2.4, 2, 5, { dot: 0.5, gap: 0.16, groupGap: 0.4 });
      addFooter(s, FOOTER);
      s.addNotes(NOTES_DR3);
      return s;
    },
    (slide) => {
      addRevealAnswerBar(slide, "10 counters - ten", { y: 4.3, h: 0.7, fontSize: 28 });
    }
  );

  // ── Fluency: subitising with dot cards ─────────────────────────────────────

  // Slide 9-10: FL1 - subitise 4
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Fluency");
      addTitle(s, "How many dots?", { color: STAGE_COLORS["1"] });
      addDotCard(s, 3.85, 1.7, 2.3, 4, { dotColor: C.PRIMARY });
      addFooter(s, FOOTER);
      s.addNotes(NOTES_FL1);
      return s;
    },
    (slide) => {
      addRevealAnswerBar(slide, "4 dots", { y: 4.3, h: 0.7, fontSize: 30 });
    }
  );

  // Slide 11-12: FL2 - subitise 6
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Fluency");
      addTitle(s, "How many dots?", { color: STAGE_COLORS["1"] });
      addDotCard(s, 3.85, 1.7, 2.3, 6, { dotColor: C.PRIMARY });
      addFooter(s, FOOTER);
      s.addNotes(NOTES_FL2);
      return s;
    },
    (slide) => {
      addRevealAnswerBar(slide, "6 dots", { y: 4.3, h: 0.7, fontSize: 30 });
    }
  );

  // Slide 13-14: FL3 - subitise 5
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["1"]);
      addStageBadge(s, 1, "Fluency");
      addTitle(s, "How many dots?", { color: STAGE_COLORS["1"] });
      addDotCard(s, 3.85, 1.7, 2.3, 5, { dotColor: C.PRIMARY });
      addFooter(s, FOOTER);
      s.addNotes(NOTES_FL3);
      return s;
    },
    (slide) => {
      addRevealAnswerBar(slide, "5 dots", { y: 4.3, h: 0.7, fontSize: 30 });
    }
  );

  // Slide 15: Launch - a full ten frame is 10
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Launch", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Our ten frame friend", { color: C.PRIMARY });
    heroTenFrame(s, 10, { y: 1.55 });
    addTextOnShape(s, "A full ten frame is 10. Today we make 10.", {
      x: 1.5, y: 3.75, w: 7.0, h: 0.6, rectRadius: 0.1,
      fill: { color: C.PRIMARY },
    }, { fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true });
    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  })();

  // Slide 16: LI / SC
  liSlide(pres, LI, SC, NOTES_LI_SC, FOOTER);

  // Slide 17: Vocabulary - friends of 10
  (() => {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Key Words", { color: C.PRIMARY, w: 1.9 });
    addTitle(s, "Friends of 10", { color: C.PRIMARY });
    heroTenFrame(s, 5, { y: 1.5 });
    addInstructionCard(s, [
      { role: "header", text: "5 and 5 make 10" },
      { role: "body", text: "Two numbers that make 10 are friends of 10." },
    ], { x: 0.5, y: 3.55, w: 9, h: 1.05, strip: C.ACCENT });
    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  })();

  // Slide 18: I Do 1 - 6 and 4 make 10
  workedExSlide(pres, 2, "I Do", "How many more make 10?",
    [
      "6 counters are in the frame.",
      "Count the empty boxes.",
      "1, 2, 3, 4.",
      "6 and 4 make 10.",
    ],
    NOTES_IDO1, FOOTER,
    (slide, lg) => {
      const cardH = SAFE_BOTTOM - lg.panelTopPadded;
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, cardH, { strip: C.PRIMARY });
      slide.addText("6 counters, 4 empty boxes", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      addTensFrame(slide, lg.rightX + 0.3, lg.panelTopPadded + 0.6, lg.rightW - 0.6, 6,
        { cellH: 0.72, fillColor: COUNTER });
      addTextOnShape(slide, "6 and 4 make 10", {
        x: lg.rightX + 0.5, y: lg.panelTopPadded + 2.35, w: lg.rightW - 1.0, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 19: I Do 2 - 8 and 2 make 10
  workedExSlide(pres, 2, "I Do", "How many more make 10?",
    [
      "8 counters are in the frame.",
      "Count the empty boxes.",
      "1, 2.",
      "8 and 2 make 10.",
    ],
    NOTES_IDO2, FOOTER,
    (slide, lg) => {
      const cardH = SAFE_BOTTOM - lg.panelTopPadded;
      addCard(slide, lg.rightX, lg.panelTopPadded, lg.rightW, cardH, { strip: C.PRIMARY });
      slide.addText("8 counters, 2 empty boxes", {
        x: lg.rightX + 0.2, y: lg.panelTopPadded + 0.12, w: lg.rightW - 0.4, h: 0.3,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, align: "center", margin: 0,
      });
      addTensFrame(slide, lg.rightX + 0.3, lg.panelTopPadded + 0.6, lg.rightW - 0.6, 8,
        { cellH: 0.72, fillColor: COUNTER });
      addTextOnShape(slide, "8 and 2 make 10", {
        x: lg.rightX + 0.5, y: lg.panelTopPadded + 2.35, w: lg.rightW - 1.0, h: 0.5, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    }
  );

  // Slide 20-21: CFU 1 - 7 -> 3 more
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "How many more make 10?", { color: C.ALERT });
      addTextOnShape(s, "✓  CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });
      heroTenFrame(s, 7, { y: 1.5 });
      routinePill(s, "Show me on your board", C.ALERT);
      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU1);
      return s;
    },
    (slide) => {
      addRevealAnswerBar(slide, "3 more. 7 and 3 make 10", { y: 4.4, h: 0.66, fontSize: 24 });
    }
  );

  // Slide 22-23: We Do 1 - 4 -> 6 more
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "How many more make 10?", { color: STAGE_COLORS["3"] });
      heroTenFrame(s, 4, { y: 1.5 });
      routinePill(s, "Count the empty boxes with me", STAGE_COLORS["3"]);
      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO1);
      return s;
    },
    (slide) => {
      addRevealAnswerBar(slide, "6 more. 4 and 6 make 10", { y: 4.4, h: 0.66, fontSize: 24 });
    }
  );

  // Slide 24-25: We Do 2 - 9 -> 1 more
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "How many more make 10?", { color: STAGE_COLORS["3"] });
      heroTenFrame(s, 9, { y: 1.5 });
      routinePill(s, "Show me on your board", STAGE_COLORS["3"]);
      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO2);
      return s;
    },
    (slide) => {
      addRevealAnswerBar(slide, "1 more. 9 and 1 make 10", { y: 4.4, h: 0.66, fontSize: 24 });
    }
  );

  // Slide 26-27: We Do 3 - build it yourself, 3 -> 7 more
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, STAGE_COLORS["3"]);
      addStageBadge(s, 3, "We Do");
      addTitle(s, "Build 10 on your frame", { color: STAGE_COLORS["3"] });
      heroTenFrame(s, 3, { y: 1.5 });
      routinePill(s, "Fill your frame to 10. How many more?", STAGE_COLORS["3"]);
      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO3);
      return s;
    },
    (slide) => {
      addRevealAnswerBar(slide, "7 more. 3 and 7 make 10", { y: 4.4, h: 0.66, fontSize: 24 });
    }
  );

  // Slide 28-29: CFU 2 (hinge) - 2 -> 8 more
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.ALERT);
      addBadge(s, "CFU", { color: C.ALERT });
      addTitle(s, "How many more make 10?", { color: C.ALERT });
      addTextOnShape(s, "✓  CHECK", {
        x: 7.8, y: 0.20, w: 1.6, h: 0.42, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
      }, { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true });
      heroTenFrame(s, 2, { y: 1.5 });
      routinePill(s, "Show me on your board", C.ALERT);
      addFooter(s, FOOTER);
      s.addNotes(NOTES_CFU2);
      return s;
    },
    (slide) => {
      addRevealAnswerBar(slide, "8 more. 2 and 8 make 10", { y: 4.4, h: 0.66, fontSize: 24 });
    }
  );

  // Slide 30: You Do
  (() => {
    const s = pres.addSlide();
    addTopBar(s, STAGE_COLORS["4"]);
    addStageBadge(s, 4, "You Do");
    addTitle(s, "Make 10 on your own", { color: STAGE_COLORS["4"] });

    addInstructionCard(s, [
      { role: "header", text: "On your worksheet" },
      { role: "body", text: "First: look at each ten frame." },
      { role: "body", text: "Next: draw counters to make 10." },
      { role: "body", text: "Then: write how many more." },
    ], { x: 0.5, y: CONTENT_TOP, w: 4.6, h: SAFE_BOTTOM - CONTENT_TOP, strip: STAGE_COLORS["4"] });

    addCard(s, 5.3, CONTENT_TOP, 4.2, SAFE_BOTTOM - CONTENT_TOP, { strip: C.SECONDARY });
    s.addText("Count the empty boxes", {
      x: 5.5, y: CONTENT_TOP + 0.14, w: 3.8, h: 0.32,
      fontSize: 15, fontFace: FONT_H, color: C.SECONDARY, bold: true, align: "center", margin: 0,
    });
    addTensFrame(s, 5.9, CONTENT_TOP + 0.7, 3.0, 6, { cellH: 0.6, fillColor: COUNTER });
    s.addText("6 in the frame - draw 4 more", {
      x: 5.5, y: CONTENT_TOP + 2.1, w: 3.8, h: 0.4,
      fontSize: 13, fontFace: FONT_B, color: C.MUTED, italic: true, align: "center", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  // Slide 31: Exit Ticket - 7 -> 3 more (custom so the ten frame stays the hero)
  (() => {
    const s = pres.addSlide();
    s.background = { color: C.BG_CARD };
    s.addShape("rect", { x: 0, y: 0, w: 10, h: 0.06, fill: { color: C.ASSESS } });
    addBadge(s, "Exit Ticket", { color: C.ASSESS, x: 0.5, y: 0.2, w: 2.1 });
    s.addText("How many more make 10?", {
      x: 0.5, y: 0.72, w: 9, h: 0.7,
      fontSize: 30, fontFace: FONT_H, color: C.ASSESS, bold: true, margin: 0,
      fit: "shrink", shrinkText: true,
    });
    heroTenFrame(s, 7, { y: 1.7 });
    addTextOnShape(s, "Write your number and hold up your board", {
      x: 1.3, y: 3.85, w: 7.4, h: 0.6, rectRadius: 0.1,
      fill: { color: C.ASSESS },
    }, { fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true });
    addFooter(s, FOOTER);
    s.addNotes(NOTES_EXIT);
  })();

  // Slide 32: Closing
  closingSlide(pres,
    {
      reflectionPrompt: "Turn and tell your partner: how do we find how many more make 10?",
      scItems: SC,
      selfAssessment: "Thumbs up / sideways / down",
    },
    NOTES_CLOSING);

  // Write PPTX
  fs.mkdirSync(OUT_DIR, { recursive: true });
  const pptxPath = path.join(OUT_DIR, "MK10_Lesson1_Making_10_Ten_Frames.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);

  // ─── Companion PDFs ─────────────────────────────────────────────────────────

  // 1) Make 10 worksheet - draw counters to fill each frame, write how many more.
  await (async () => {
    const doc = createPdf({ title: WORKSHEET_RES.name });
    let y = addPdfHeader(doc, WORKSHEET_RES.name, {
      subtitle: "Fill each ten frame to make 10.",
      color: C.NAVY,
      lessonInfo: "Foundation Numeracy | Making 10",
    });
    y = addTipBox(doc,
      "Draw a counter in each empty box to make 10. Then write how many more.",
      y, { color: C.TEAL });

    // Big cells for young hands, sized so all four frames fit above the footer
    // on one page (megaprompt 68i handwriting space rule + 61a page-fit rule).
    const CELL = 40;

    function drawFrameRow(doc, label, filled, y) {
      y = addSectionHeading(doc, label, y, { color: C.NAVY });
      const bottom = addTenFramePdf(doc, 60, y, filled, {
        cellSize: CELL, fillColor: hex(C.PRIMARY),
      });
      return addWriteLine(doc, "How many more make 10?", bottom + 4);
    }

    y = drawFrameRow(doc, "1. Draw counters to make 10", 6, y);
    y = drawFrameRow(doc, "2. Draw counters to make 10", 8, y);
    y = drawFrameRow(doc, "3. Draw counters to make 10", 5, y);

    y = addSectionHeading(doc, "Challenge: make 10 a different way", y, { color: C.TEAL });
    y = addBodyText(doc, "Draw your own counters to make 10.", y);
    addTenFramePdf(doc, 60, y + 2, 0, { cellSize: 34, fillColor: hex(C.PRIMARY) });

    addPdfFooter(doc, "Session 1 | Make 10 Worksheet | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, WORKSHEET_RES.fileName));
    console.log("PDF written: " + WORKSHEET_RES.fileName);
  })();

  // 2) Answer key
  await (async () => {
    const doc = createPdf({ title: ANSWER_KEY_RES.name });
    let y = addPdfHeader(doc, ANSWER_KEY_RES.name, {
      subtitle: "Answers for the Make 10 worksheet.",
      color: C.NAVY,
      lessonInfo: "Foundation Numeracy | Making 10",
      showNameDate: false,
    });
    y = addSectionHeading(doc, "Answers", y, { color: C.NAVY });
    y = addBodyText(doc, "1. 6 in the frame - draw 4 more. 6 and 4 make 10.", y);
    y = addBodyText(doc, "2. 8 in the frame - draw 2 more. 8 and 2 make 10.", y);
    y = addBodyText(doc, "3. 5 in the frame - draw 5 more. 5 and 5 make 10.", y);
    y = addBodyText(doc, "Challenge: any pair that makes 10 (for example 7 and 3, or 2 and 8).", y);
    y = addTipBox(doc,
      "Watch for: children who write the number already in the frame. Redirect them to count only the empty boxes.",
      y, { color: C.TEAL });
    addPdfFooter(doc, "Session 1 | Answer Key | Foundation Numeracy");
    await writePdf(doc, path.join(OUT_DIR, ANSWER_KEY_RES.fileName));
    console.log("PDF written: " + ANSWER_KEY_RES.fileName);
  })();

  console.log("Lesson build complete.");
}

build().catch((err) => { console.error(err); process.exit(1); });
