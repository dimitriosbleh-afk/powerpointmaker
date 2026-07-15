"use strict";

// Big Ideas (Year 5/6) - Session 2 of 2: Numbers Too Big To Write.
// Student-interest enrichment, not a curriculum unit. Brand new content.
//
// The hook is Graham's number, but the LEARNING is powers of 10: the little
// raised number counts the zeros, and every extra zero is ten times more.
// That is real, assessable Year 5/6 maths, and it is what the exit ticket
// tests. Graham's number is the payoff that makes students want the tool.
//
// The spine is a story about NOTATION outgrowing itself:
//   write the zeros -> too many zeros, use a power -> the power is too big to
//   write, use a tower -> the tower is too tall, use arrows -> Graham's number.
// Anchor: "The little number counts the zeros. One more zero is ten times more."
//
// Content and images guided by the user's supplied "Visualising Large Numbers"
// deck. Facts verified before use: googol was named by Milton Sirotta, aged 9;
// Graham's number was in the 1980 Guinness Book of World Records as the largest
// number used in a maths proof (the title was later retired); its last digits
// are known (it ends in 7) but its first digits are not.
//
// Sessions 1 and 2 are independent topics. This session does not need Session 1.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const {
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
} = require("../themes/pdf_helpers");

const T = createTheme("numeracy", "grade56", 2); // variant 2, shared across both sessions
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  keyWordSlide, exitTicketSlide,
  dailyReviewSlide, fluencySlide,
  addPlaceValueChart,
  addTextOnShape, addCard, addFooter, addTopBar, addTitle, addBadge,
  addInstructionCard, addInstructionalImageCard,
  withReveal, runSlideDiagnostics, addRevealAnswerBar,
  composeGlanceNotes,
} = T;

const SESSION = 2;
const TOTAL = 2;
const FOOTER = `Big Ideas | Session ${SESSION} of ${TOTAL} | Year 5/6 Numeracy`;
const OUT_DIR = "output/BigIdeas_Session2_GrahamsNumber";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const ASSETS = path.join(__dirname, "..", "assets", "bigideas");
const IMG_GRAIN = path.join(ASSETS, "sand-one-grain.png");
const IMG_JIGSAW = path.join(ASSETS, "sand-million-jigsaw.png");
const IMG_EARTH = path.join(ASSETS, "earth-blue-marble.png");
const IMG_STARS = path.join(ASSETS, "stars-field.png");
const IMG_BRAIN = path.join(ASSETS, "brain-glow.png");

const SUPPLIED = "Images and scale facts from the supplied Visualising Large Numbers deck";

const WORKSHEET_RES = makeSessionResource(SESSION,
  "Worksheet",
  "Count the zeros, order the powers of 10, and compare how many times bigger.");
const ANSWER_KEY_RES = makeSessionResource(SESSION,
  "Answer Key",
  "Worked answers and look-fors for the Session 2 worksheet.");
const RESOURCE_ITEMS = [WORKSHEET_RES, ANSWER_KEY_RES];

fs.mkdirSync(RES_DIR, { recursive: true });

/* ------------------------------------------------------------------ */
/*  Local helpers                                                      */
/* ------------------------------------------------------------------ */

// Proper exponent notation: 10 with a real raised power, not "10^6".
// addTextOnShape cannot do this (it takes a plain string), so the chip is
// drawn as a shape plus a two-run text box.
function addPowerChip(slide, x, y, w, h, base, exp, opts) {
  const o = opts || {};
  const fill = o.fill || C.PRIMARY;
  const color = o.color || C.WHITE;
  slide.addShape("roundRect", {
    x, y, w, h, rectRadius: o.rectRadius != null ? o.rectRadius : 0.08,
    fill: { color: fill },
    line: { color: o.line || fill, width: 1 },
  });
  const size = o.fontSize || 20;
  slide.addText([
    { text: String(base), options: { fontSize: size, fontFace: FONT_H, color, bold: true } },
    { text: String(exp), options: { fontSize: size, fontFace: FONT_H, color, bold: true, superscript: true } },
  ], {
    x, y, w, h, align: "center", valign: "middle", margin: 0,
  });
}

// A bottle cap, drawn rather than photographed. The supplied deck's bottle-cap
// photo carries a visible stock-library watermark, so it is not used.
function drawBottleCap(slide, x, y, w, h) {
  const capColor = C.SECONDARY;
  const rimH = h * 0.42;

  // Ridged body first, rim drawn over it.
  slide.addShape("roundRect", {
    x, y: y + rimH * 0.45, w, h: h - rimH * 0.45, rectRadius: 0.07,
    fill: { color: capColor }, line: { color: capColor, width: 1 },
  });
  const ridges = 9;
  const inset = 0.13;
  const step = (w - inset * 2) / (ridges - 1);
  for (let i = 0; i < ridges; i++) {
    slide.addShape("roundRect", {
      x: x + inset + i * step - 0.014, y: y + rimH * 0.95, w: 0.028, h: h * 0.44,
      rectRadius: 0.01,
      fill: { color: C.WHITE }, line: { color: C.WHITE, width: 0.4 },
    });
  }
  // Rim: a pill (roundRect with radius = half height) reads as the circular
  // opening seen at an angle. A real oval would not render in LibreOffice.
  slide.addShape("roundRect", {
    x, y, w, h: rimH, rectRadius: rimH / 2,
    fill: { color: capColor }, line: { color: C.WHITE, width: 1.4 },
  });
  // The water inside the cap.
  slide.addShape("roundRect", {
    x: x + 0.14, y: y + 0.07, w: w - 0.28, h: rimH - 0.14,
    rectRadius: (rimH - 0.14) / 2,
    fill: { color: C.ACCENT }, line: { color: C.ACCENT, width: 0.6 },
  });
}

// Plain power text with no chip behind it.
function addPowerText(slide, x, y, w, h, base, exp, opts) {
  const o = opts || {};
  const size = o.fontSize || 20;
  const color = o.color || C.CHARCOAL;
  slide.addText([
    { text: String(base), options: { fontSize: size, fontFace: FONT_H, color, bold: true } },
    { text: String(exp), options: { fontSize: size, fontFace: FONT_H, color, bold: true, superscript: true } },
  ], {
    x, y, w, h, align: o.align || "center", valign: "middle", margin: 0,
  });
}

// A power tower, drawn as a staircase climbing up and to the right, each level
// smaller than the one below. This is the picture that shows WHY the notation
// had to change: the exponent itself has grown into another number.
//
// yBottom is the BASELINE of the first (largest) level. Boxes advance by their
// full width so successive levels never overlap - a raised exponent reads as
// raised because of the baseline shift, not because the boxes collide.
function drawPowerTower(slide, xLeft, yBottom, levels, opts) {
  const o = opts || {};
  const color = o.color || C.SECONDARY;
  let size = o.fontSize || 30;
  let x = xLeft;
  let baseline = yBottom;
  levels.forEach((level) => {
    const scale = size / 30;
    const w = 0.26 + String(level).length * 0.20 * scale;
    const h = 0.46 * scale + 0.08;
    slide.addText(String(level), {
      x, y: baseline - h, w, h,
      fontSize: size, fontFace: FONT_H, color, bold: true,
      align: "left", valign: "middle", margin: 0,
    });
    x += w;
    baseline -= h * 0.55;
    size = Math.max(10, size * 0.70);
  });
  return { rightX: x, topY: baseline };
}

/* ------------------------------------------------------------------ */
/*  Teacher notes                                                      */
/* ------------------------------------------------------------------ */

const NOTES_TITLE = "Session 2 of 2. Enrichment on student interest. The hook is Graham's number; the learning is powers of 10. Independent of Session 1.";

const NOTES_RESOURCES = composeGlanceNotes({
  answer: "open - teacher preparation slide, not taught",
  beats: [
    "SHOW while students settle. SAY: Whiteboards out. One sheet today, and it comes out at the You Do.",
    "CHECK before the lesson: you can say 7,625,597,484,987 out loud (seven point six trillion) without stumbling.",
  ],
  prep: "Materials: mini-whiteboards, the worksheet. No manipulatives needed. Independent of Session 1, so a student who missed it loses nothing.",
  tag: "[Setup | Planning | HITS 2]",
});

const NOTES_DR = composeGlanceNotes({
  answer: "600. And 4 zeros, which makes 10 000.",
  beats: [
    "POINT to the chart. SAY: Every step left is ten times bigger. That is the whole idea we need today.",
    "ASK: What are the two answers? 40 sec. Write it... chin it... show me, boards up. EXPECT: 600, and 4 zeros.",
    "SCAN boards, back row first. 80%+ -> reveal. Less -> build 6 x 100 on the chart, move the 6 two columns left, re-ask with 6 x 10.",
  ],
  trap: "counting the zeros in the question instead of the answer. Fix: point at the answer on the chart, student recounts the zeros there.",
  prep: "Prior learning: place value and multiplying by 10. Assumption flagged: Year 5s may be newer to this than Year 6s, so the chart stays on screen. Everything today rests on one more zero means ten times more.",
  tag: "[Stage 1 | Daily Review | Retention and recall | HITS 6]",
});

const NOTES_FLUENCY = composeGlanceNotes({
  answer: "80, 800, 8000.",
  beats: [
    "SAY: Fluency. Each one has one more ten in it than the last. Watch the zeros grow.",
    "ASK: What are the answers? 30 sec, boards up on cue. EXPECT: 80, 800, 8000.",
    "SCAN boards. 80%+ -> reveal and move on. Less -> say the chain out loud together, eight tens, eight hundreds, eight thousands, re-ask.",
  ],
  trap: "adding a zero as a rule with no meaning. Fix: say ten times bigger every time, student restates one as a times-ten.",
  prep: "Brisk automaticity (Number). Feeds straight into counting zeros in a power. Whole block under 4 minutes.",
  tag: "[Stage 1 | Fluency | Mastery and application | HITS 6]",
});

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "open - any guess is welcome; the real answer is about 10 with 19 zeros after it",
  beats: [
    "SHOW the grain. SAY: One grain of sand on a fingertip. You could count that. Easy.",
    "ASK: How many grains of sand on every beach on Earth? Have a guess. 30 sec, boards up on cue. EXPECT: anything. ACCEPT: a million, a billion, too many to count.",
    "COLLECT the range without judging. SAY: I can see guesses from a thousand to a squillion, and not one of us can write the real one down easily. That is our problem today.",
  ],
  trap: "refusing to guess because it feels unknowable. Fix: SAY: a rough guess is a real maths move, student commits to any number.",
  prep: "Low-coupling launch: guessing needs no prior knowledge, so every student enters. It creates the need for the tool before the tool arrives. Keep to 4 minutes.",
  tag: "[Launch | Attention, focus and regulation | HITS 3]",
  sources: [SUPPLIED],
});

const NOTES_LI = composeGlanceNotes({
  answer: "open - students read the criteria, no response collected",
  beats: [
    "SAY: By the end you will read numbers so big there is not enough stuff in the universe to write them.",
    "POINT to criterion one. SAY: Everyone gets this one. It is one small number doing one job.",
  ],
  prep: "SC1 is the anchor and is reachable by every student. SC2 is the exit ticket. SC3 is the Graham's number story.",
  tag: "[LI and SC | Planning | HITS 1]",
});

const NOTES_VOCAB = composeGlanceNotes({
  answer: "the little raised number counts how many tens are multiplied, which is how many zeros you write",
  beats: [
    "POINT to the raised 6. SAY: This little number is the whole trick. It counts the tens being multiplied together.",
    "SAY: Six tens multiplied gives you six zeros. So it also counts the zeros. That is the shortcut we will use all lesson.",
    "SAY: Say it with me. Everyone, together, on three. One, two, three... the little number counts the zeros.",
  ],
  trap: "reading 10 to the 6 as 10 x 6 = 60. Fix: write both, 10 x 6 = 60 and 10 to the 6 = 1 000 000, student says which is which.",
  prep: "One word, and it is the anchor for the session. The 10 x 6 = 60 error is the single most common one; name it here, before it costs anyone the exit ticket.",
  tag: "[Vocabulary | Knowledge and memory | SC1 | HITS 6]",
});

const NOTES_IDO_ZEROS = composeGlanceNotes({
  answer: "10 to the 6 is 1 000 000: six tens multiplied, so six zeros",
  beats: [
    "POINT to the top row. SAY: One ten, one zero. Watch what happens as I go down.",
    "MODEL each row. SAY: Two tens, two zeros. Three tens, three zeros. Every single time, the little number tells me exactly how many zeros to write.",
    "POINT to the last row. SAY: So six tens multiplied. I do not have to work it out. I just write a 1 and count six zeros after it.",
    "ASK: How many zeros in 10 to the 9? 15 sec, boards up on cue. EXPECT: 9.",
    "SCAN boards. 80%+ -> next slide. Less -> add the 10 to the 4 row to the ladder together, count the tens aloud, re-ask 10 to the 9.",
  ],
  trap: "writing 10 to the 6 as 100 000 by counting the 1 as a zero. Fix: circle the 1, count only what follows, student recounts.",
  stretch: "work out how many zeros in 10 to the 6 times 10 to the 3, and explain why.",
  help: "write the tens out as 10 x 10 x 10 and count them with a finger before writing zeros.",
  prep: "The anchor made visible: the ladder lets students SEE the rule rather than be told it. The jump from 3 to 6 is deliberate, so they use the rule instead of pattern-matching row to row.",
  tag: "[I Do | Explicit teaching | SC1 | HITS 3, 4]",
});

const NOTES_IDO_LADDER = composeGlanceNotes({
  answer: "there are more stars in the universe than grains of sand on Earth, and more atoms in a bottle cap of water than stars",
  beats: [
    "POINT to the bag. SAY: A million grains of sand, sold as a one million piece jigsaw. Expert level. A million is a bag you can hold.",
    "POINT along. SAY: All the sand on every beach and under every ocean is 19 zeros. Stars in the whole universe, 24 zeros. So there are more stars than there are grains of sand. That is real.",
    "POINT to the last rung. SAY: Now fill a bottle cap with water. The atoms in that are 26 zeros. More atoms in a bottle cap than stars in the universe.",
    "ASK: Is 24 to 26 a small step or a huge one? 20 sec, turn and tell. EXPECT: huge, it is a hundred times. ACCEPT: huge, because each zero is ten times.",
  ],
  trap: "reading 19 to 24 as only five bigger. Fix: five more zeros means ten times, five times over, student says how many times.",
  stretch: "the atoms in the whole Earth are 50 zeros. How many times bigger than a bottle cap?",
  help: "cover the pictures and just read the little numbers in order: 6, 19, 24, 26.",
  prep: "The wow, but the teaching move is that the exponent alone ranks them. Students should be able to order these without understanding astronomy.",
  tag: "[I Do | Explicit teaching | SC2 | HITS 4, 6]",
  sources: [SUPPLIED],
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B - 1000 times bigger, because there are 3 more zeros and each zero means ten times",
  beats: [
    "SAY: 10 to the 6 is a million. 10 to the 9 is a billion. This is the one that matters: not which is bigger, but how MUCH bigger.",
    "ASK: How many times bigger is 10 to the 9 than 10 to the 6? 30 sec think. Write A, B or C, chin it, boards up. EXPECT: B.",
    "SCAN boards, back row first. 80%+ -> cold call one B board: How do you know? Then reveal and move on. Less -> write 1 000 000 and 1 000 000 000 under each other, ring the 3 extra zeros, re-ask.",
    "REVEAL after every board is up. SAY: Three more zeros. Ten, times ten, times ten. A thousand times.",
  ],
  trap: "choosing A, 3 times bigger, by subtracting the little numbers. Fix: the 3 is how many EXTRA TENS, not how many times, student re-states with the zeros ringed.",
  prep: "The decision point of the lesson and the exact structure of the exit ticket. A is the subtract-the-exponents error, C is a guess. Do not release to the You Do until this is secure.",
  tag: "[CFU hinge | Supported application | SC2 | HITS 7, 8]",
});

const NOTES_GOOGOL = composeGlanceNotes({
  answer: "a googol is 10 with 100 zeros, and it is bigger than the number of atoms in the universe",
  beats: [
    "SAY: In 1937 a mathematician asked his nine year old nephew to name a really big number. The boy said googol. That is the actual story, and it is why the search engine is called what it is.",
    "SAY: A googol is 10 with 100 zeros. You could write it on a page in about a minute.",
    "POINT. SAY: But the atoms in the whole universe are only 82 zeros. So if you drew a dot for every atom in existence, you would still not reach a googol dots. You would run out of universe.",
    "ASK: A googol is how many times bigger than the atoms in the universe? 25 sec, turn and tell. EXPECT: 18 more zeros, so huge. ACCEPT: way bigger, because 100 beats 82.",
  ],
  trap: "thinking a googol is unwriteable. Fix: it is easy to write, it is the AMOUNT that is impossible, student says the difference.",
  stretch: "how many zeros between a googol and the atoms in the universe? Say it as times bigger.",
  help: "just compare the two little numbers: 100 and 82. Which is bigger?",
  prep: "The turn: the number is still easy to WRITE but impossible to HAVE. Sets up the moment when writing itself breaks. The nine-year-old is a real fact and it lands hard with this age group.",
  tag: "[We Do | Supported application | SC3 | HITS 6, 7]",
  sources: ["Googol named by Milton Sirotta, aged 9, nephew of mathematician Edward Kasner", SUPPLIED],
});

const NOTES_TOWER = composeGlanceNotes({
  answer: "a googolplex is 10 with a googol zeros, and there are not enough atoms in the universe to write the zeros down",
  beats: [
    "SAY: The same boy named a second number. A googolplex. It is 10 with a GOOGOL zeros after it.",
    "POINT to the tower. SAY: Look what happened to our writing. The little number got so big it needed its own little number. The power grew a power.",
    "SAY: Here is the killer. You cannot write a googolplex out. Not because you would get bored. Because there are not enough atoms in the universe to hold the zeros.",
    "ASK: So why did we need a new way to write? 25 sec, turn and tell. EXPECT: the zeros do not fit, so we stack the powers instead. ACCEPT: because there were too many zeros.",
  ],
  trap: "hearing googolplex as googol plus a bit. Fix: googol is 100 zeros, googolplex is a googol of them, student says which is the number of zeros.",
  stretch: "explain to your partner why a googolplex is not just a googol times a googol.",
  help: "read the tower bottom to top: ten, to the ten, to the hundred.",
  prep: "The first notation break. This is the SC3 idea in one slide: when the number outgrows the writing, mathematicians invent new writing.",
  tag: "[We Do | Supported application | SC3 | HITS 4, 9]",
  sources: [SUPPLIED],
});

const NOTES_PATTERN = composeGlanceNotes({
  answer: "doing powers again and again, which mathematicians write with arrows: 3 arrow arrow 4",
  beats: [
    "POINT to row one. SAY: Adding the same number again and again has a shortcut. We call it multiplying.",
    "POINT to row two. SAY: Multiplying the same number again and again has a shortcut too. We call it a power.",
    "ASK: So what is doing POWERS again and again? There must be a next one. 40 sec, turn and tell, then boards up. EXPECT: another shortcut, a new symbol. ACCEPT: powers of powers.",
    "SCAN boards. 80%+ -> reveal, then cold call one board: Why did there have to be a next row? Less -> re-read rows one and two, ask what changed each time, re-ask.",
    "REVEAL after boards scanned. SAY: Arrows. And every arrow you add is a whole new level, not a bit more.",
  ],
  trap: "expecting the next row to be a slightly bigger number. Fix: point at the jump from row one to row two, it changed the OPERATION, student names what row three changes.",
  stretch: "work out 3 arrow 3 without a calculator, then predict whether 3 arrow arrow 3 has more or fewer than 10 digits.",
  help: "just read rows one and two aloud and say what the shortcut was called each time.",
  prep: "The best thinking moment in the lesson: students predict that a next level must exist before they are told. Real mathematics (hyperoperations), reachable from what they know.",
  tag: "[We Do | Supported application | SC3 | HITS 7, 9]",
});

const NOTES_GRAHAM = composeGlanceNotes({
  answer: "Graham's number is g64: you build g1 with four arrows, then repeat the process 64 times",
  beats: [
    "POINT to the arrows. SAY: Three arrows already gives a tower of threes that is seven trillion levels tall. Four arrows is g one. That is just the FIRST step.",
    "SAY: Then you do it again. And again. Sixty four times. That last one is Graham's number, and Ronald Graham was a real person who needed it as an answer to a real maths problem.",
    "SHOW the brain. SAY: If you tried to hold every digit of it in your head, scientists say there is a limit to how much can fit in a space that size. Your head would collapse into a black hole first.",
    "SAY: And the strangest part. We know how it ENDS. It ends in a 7. Nobody knows what it starts with.",
    "ASK: Why could nobody just write it out? 30 sec, turn and tell. EXPECT: the zeros will not fit, the tower will not fit, there is not enough universe. ACCEPT: it is too big to write.",
  ],
  trap: "thinking Graham's number is infinity. Fix: it is a definite whole number with a last digit, infinity has no last digit, student says which one ends in 7.",
  stretch: "if we know the last digits but not the first, what does that tell you about how it was worked out?",
  help: "count the arrows on the screen, 1, 2, 3, 4, and say what each one does.",
  prep: "The payoff. Every claim here is checked: Guinness 1980 (title later retired), ends in 7 with first digits unknown, and the black hole line is a real consequence of a limit on information in a given space.",
  tag: "[I Do | Mastery and application | SC3 | HITS 3, 6]",
  sources: ["Graham's number: 1980 Guinness Book of World Records, largest number used in a maths proof; last digits known, ends in 7", SUPPLIED],
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "open - the ordering is 10 to the 6, 19, 24, 26, 82, 100",
  beats: [
    "SAY: Sheet out. Section one warms up the counting, then you order the real ones.",
    "SAY: The only tool you need is the little number. Ignore what the things are.",
    "TIME: 12 minutes. CIRCULATE and read section 3 first, the how-many-times question, it is the one the exit ticket needs.",
  ],
  trap: "ordering by how impressive the object sounds instead of by the power. Fix: cover the words with a thumb, read only the little numbers, student re-orders.",
  stretch: "worksheet challenge: write your own scale fact using a power of 10.",
  help: "section 1 rebuilds it: write out the zeros first, then compare the lengths.",
  prep: "Different content from the We Do: they now order and compare independently rather than following a walked ladder. Section 1 re-grounds anyone who needs it.",
  tag: "[You Do | Mastery and application | SC2 | HITS 4, 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "10 to the 26 is bigger. It is 100 times bigger, because it has 2 more zeros and each zero is ten times.",
  beats: [
    "SAY: One question. Four minutes. The how many times part is the bit I am reading for.",
    "COLLECT in books or on a slip. SAY: I want a reason, not just a number.",
    "SORT as you collect. Anyone who wrote 2 times bigger goes in a pile: that is the subtract-the-exponents error from the check.",
  ],
  trap: "answering 2 times bigger. Fix: the 2 is extra zeros, not times, ring the zeros, student recounts as ten times ten.",
  prep: "Assesses SC2. Do not show the SC number to students. This is the same structure as the hinge with new numbers, so it is a fair test, not a new task.",
  tag: "[Exit Ticket | Mastery and application | SC2 | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - self-assessment against the three criteria",
  beats: [
    "POINT to criterion one. ASK: Can you say what the little number tells you? 10 sec, thumbs only, voices off. EXPECT: mostly thumbs up.",
    "READ the thumbs across the room. SAY back what you saw, naming the strategy, not the person.",
    "ASK: What is the biggest number you can now write in five seconds? 20 sec, turn and tell. EXPECT: a power with a huge little number.",
  ],
  prep: "Ties back to the launch: their sand guess is now writeable. Note any thumbs-down on criterion two, that is the exit-ticket skill.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

/* ------------------------------------------------------------------ */
/*  Companion PDFs                                                     */
/* ------------------------------------------------------------------ */

async function writeWorksheet() {
  const doc = createPdf({ title: "Session 2 Worksheet" });
  let y = addPdfHeader(doc, "Numbers Too Big To Write", {
    subtitle: "Powers of 10",
    color: C.PRIMARY,
    showNameDate: true,
  });

  y = addSectionHeading(doc, "1. Count the zeros", y + 4, { color: C.PRIMARY });
  y = addBodyText(doc,
    "The little raised number tells you how many zeros to write. Fill in the table.",
    y, { fontSize: 11 });
  y = drawZeroTablePdf(doc, y + 6);

  y = addSectionHeading(doc, "2. Put them in order", y + 4, { color: C.PRIMARY });
  y = addBodyText(doc,
    "These are all real. Number them 1 to 5, from smallest to biggest. Use only the little number.",
    y, { fontSize: 11 });
  y = drawOrderBoxesPdf(doc, y + 6);

  y = addSectionHeading(doc, "3. How many times bigger?", y + 4, { color: C.SECONDARY });
  y = addBodyText(doc,
    "There are 10 to the 24 stars in the universe. A bottle cap of water holds 10 to the 26 atoms.\n"
    + "Which is bigger, and how many times bigger? How do you know?",
    y, { fontSize: 11 });
  y = addLinedArea(doc, y + 4, 3, { lineSpacing: 24 });

  addTipBox(doc,
    "Challenge: a googol is 10 with 100 zeros. The atoms in the whole universe are 10 with 82 zeros. "
    + "Explain why you could never draw a googol dots, even using every atom there is.",
    y + 6, { color: C.ALERT });

  addPdfFooter(doc, "Big Ideas | Session 2 | Numbers Too Big To Write");
  await writePdf(doc, path.join(RES_DIR, "Session 2 Worksheet.pdf"));
}

// Section 1: a real table with the model row already worked, so the first item
// is doable from the I Do alone.
function drawZeroTablePdf(doc, y) {
  const x = 50;
  const w = 495.28;
  const colW = [140, 150, 205.28];
  const rowH = 28;
  const rows = [
    ["Power of 10", "How many zeros?", "Write the number"],
    ["10 to the 2", "2", "100"],
    ["10 to the 3", "", ""],
    ["10 to the 5", "", ""],
    ["10 to the 7", "", ""],
  ];

  rows.forEach((row, r) => {
    let cx = x;
    row.forEach((cell, c) => {
      const isHeader = r === 0;
      const isModel = r === 1;
      doc.rect(cx, y + r * rowH, colW[c], rowH)
        .fillAndStroke(isHeader ? hexPdf(C.PRIMARY) : (isModel ? "#EEF2F6" : "#FFFFFF"),
          hexPdf(C.MUTED));
      doc.fontSize(isHeader ? 10 : 11)
        .font(isHeader || isModel ? "Sans-Bold" : "Sans")
        .fillColor(isHeader ? "#FFFFFF" : hexPdf(C.CHARCOAL));
      doc.text(cell, cx + 6, y + r * rowH + (rowH - (isHeader ? 10 : 11)) / 2 - 1,
        { width: colW[c] - 12, align: "left", lineBreak: false });
      cx += colW[c];
    });
  });
  doc.fontSize(9).font("Sans-Italic").fillColor(hexPdf(C.MUTED));
  doc.text("The first one is done for you.", x, y + rows.length * rowH + 3, { width: w });
  return y + rows.length * rowH + 12;
}

// Section 2: order boxes. Each item shows the power AND what it counts, so the
// task is ranking by exponent, not general knowledge.
function drawOrderBoxesPdf(doc, y) {
  const items = [
    ["10 to the 24", "stars in the universe"],
    ["10 to the 6", "grains of sand in a small bag"],
    ["10 to the 82", "atoms in the universe"],
    ["10 to the 19", "grains of sand on Earth"],
    ["10 to the 26", "atoms in a bottle cap of water"],
  ];
  const x = 50;
  const rowH = 29;
  items.forEach((it, i) => {
    const ry = y + i * rowH;
    doc.rect(x, ry, 30, 24).stroke(hexPdf(C.CHARCOAL));
    doc.rect(x + 40, ry, 130, 24).fillAndStroke("#EEF2F6", hexPdf(C.MUTED));
    doc.fontSize(11).font("Sans-Bold").fillColor(hexPdf(C.CHARCOAL));
    doc.text(it[0], x + 46, ry + 7, { width: 120, lineBreak: false });
    doc.fontSize(10.5).font("Sans").fillColor(hexPdf(C.CHARCOAL));
    doc.text(it[1], x + 180, ry + 7, { width: 300, lineBreak: false });
  });
  return y + items.length * rowH + 4;
}

function hexPdf(color) {
  return "#" + String(color).replace(/^#/, "");
}

async function writeAnswerKey() {
  const doc = createPdf({ title: "Session 2 Answer Key" });
  let y = addPdfHeader(doc, "Numbers Too Big To Write: Answer Key", {
    subtitle: "Teacher copy. Section 3 is the exit-ticket target.",
    color: C.SECONDARY,
    showNameDate: false,
  });

  y = addSectionHeading(doc, "1. Count the zeros", y + 4, { color: C.SECONDARY });
  y = addBodyText(doc,
    "10 to the 3 = 3 zeros = 1000.  10 to the 5 = 5 zeros = 100 000.  10 to the 7 = 7 zeros = 10 000 000.",
    y, { fontSize: 10.5 });
  y = addBodyText(doc,
    "Watch for students counting the leading 1 as a zero (writing 10 to the 5 as 10 000). Ring the 1 "
    + "and have them count only what follows. Also watch for 10 to the 5 = 50, which is reading the "
    + "power as a multiplication.",
    y, { fontSize: 10.5 });

  y = addSectionHeading(doc, "2. Put them in order", y + 4, { color: C.SECONDARY });
  y = addBodyText(doc,
    "Smallest to biggest: 10 to the 6 (sand in a bag), 10 to the 19 (sand on Earth), 10 to the 24 "
    + "(stars), 10 to the 26 (atoms in a bottle cap), 10 to the 82 (atoms in the universe).",
    y, { fontSize: 10.5 });
  y = addBodyText(doc,
    "The point of this item is that the exponent alone ranks them. A student who orders correctly "
    + "without knowing any astronomy has done exactly what was asked. Two facts usually cause "
    + "argument, and both are true: there are more stars in the universe than grains of sand on "
    + "Earth, and more atoms in one bottle cap of water than there are stars.",
    y, { fontSize: 10.5 });

  y = addSectionHeading(doc, "3. How many times bigger?", y + 4, { color: C.SECONDARY });
  y = addBodyText(doc,
    "10 to the 26 is bigger, and it is 100 times bigger. Look for the reasoning: 2 more zeros, and "
    + "each zero means ten times, so ten times ten = 100 times.",
    y, { fontSize: 10.5 });
  y = addBodyText(doc,
    "Do NOT accept '2 times bigger'. That is the subtract-the-exponents error and it is the main "
    + "thing this lesson is testing. The fix: write both numbers out under each other, ring the two "
    + "extra zeros, and have the student say ten times, then ten times again.",
    y, { fontSize: 10.5 });

  y = addSectionHeading(doc, "Challenge", y + 4, { color: C.ALERT });
  y = addBodyText(doc,
    "A googol has 100 zeros; the atoms in the universe have only 82. That is 18 more zeros, so a "
    + "googol is about a million million million times more than every atom in existence. Even using "
    + "one atom per dot you would run out of universe long before you finished. Strong answers say "
    + "the googol is easy to WRITE but impossible to HAVE - that is the distinction the lesson turns on.",
    y, { fontSize: 10.5 });

  addTipBox(doc,
    "If a student asks whether Graham's number is infinity: no. It is a definite whole number and we "
    + "know it ends in a 7. Infinity has no last digit. That comparison is worth two minutes.",
    y + 6, { color: C.PRIMARY });

  addPdfFooter(doc, "Big Ideas | Session 2 | Answer Key");
  await writePdf(doc, path.join(RES_DIR, "Session 2 Answer Key.pdf"));
}

/* ------------------------------------------------------------------ */
/*  Build                                                              */
/* ------------------------------------------------------------------ */

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  /* 1. Title */
  titleSlide(pres,
    "Numbers Too Big To Write",
    "Powers of 10, and Graham's number",
    "Session 2 of 2 | Year 5/6 Numeracy",
    NOTES_TITLE);

  /* 2. Teacher Resources */
  addResourceSlide(pres, RESOURCE_ITEMS, T, FOOTER, NOTES_RESOURCES);

  /* 3-4. Daily Review + reveal */
  withReveal(
    () => dailyReviewSlide(pres, "Daily Review: Place value",
      ["6 x 100 = ?", "How many zeros does 10 x 10 x 10 x 10 make?"],
      NOTES_DR, FOOTER,
      (s, guide) => {
        addPlaceValueChart(s,
          guide.rightX, guide.panelTopPadded + 0.55,
          ["10 000s", "1000s", "100s", "10s", "1s"],
          ["", "", "", "", ""],
          { totalW: guide.rightW });
      }),
    (s) => {
      addRevealAnswerBar(s, ["600", "4 zeros = 10 000"], { y: 4.30, h: 0.56, fontSize: 18 });
    }
  );

  /* 5-6. Fluency + reveal */
  withReveal(
    () => fluencySlide(pres, "Fluency: Multiplying by 10",
      ["8 x 10", "8 x 100", "8 x 1000"],
      NOTES_FLUENCY, FOOTER),
    (s) => {
      addRevealAnswerBar(s, ["80", "800", "8000"], { y: 4.30, h: 0.56, fontSize: 20 });
    }
  );

  /* 7. Launch */
  contentSlide(pres, "Launch", C.SECONDARY,
    "One grain of sand",
    [
      "You could count this.",
      "Now think bigger.",
      "How many grains on every beach on Earth?",
    ],
    NOTES_LAUNCH, FOOTER,
    (s, guide) => {
      const x = guide.rightX;
      const w = guide.rightW;
      const y = guide.panelTopPadded;
      addInstructionalImageCard(s, IMG_GRAIN, {
        x, y, w, h: 1.86, fit: "crop", strip: C.SECONDARY,
      });
      addInstructionCard(s, [
        { text: "Have a guess", role: "header" },
        { text: "", role: "spacer" },
        { text: "Write it, chin it, show me.", role: "emphasis" },
      ], {
        x, y: y + 2.02, w, h: 1.10,
        strip: C.PRIMARY, fill: C.WHITE,
      });
    });

  /* 8. LI and SC */
  liSlide(pres,
    "I can use powers of 10 to write and compare numbers that are too big to write out.",
    [
      "I can say what the little raised number tells me.",
      "I can say which of two powers of 10 is bigger, and how many times bigger.",
      "I can explain why Graham's number needed a new way to write.",
    ],
    NOTES_LI, FOOTER);

  /* 9. Vocabulary */
  keyWordSlide(pres, {
    word: "Power of 10",
    meaning: "The little raised number counts how many tens are multiplied. That is how many zeros you write.",
    example: "10 to the 6 means six tens multiplied. So it is a 1 with six zeros: 1 000 000.",
    routine: ["Say it", "Count it", "Use it"],
    color: C.PRIMARY,
  }, NOTES_VOCAB, FOOTER);

  /* 10. I Do: the little number counts the zeros */
  contentSlide(pres, "I Do", C.PRIMARY,
    "The little number counts the zeros",
    [
      "Count the tens.",
      "Write that many zeros.",
      "The pattern never breaks.",
    ],
    NOTES_IDO_ZEROS, FOOTER,
    (s, guide) => {
      const x = guide.rightX;
      const w = guide.rightW;
      const y = guide.panelTopPadded;
      addCard(s, x, y, w, 3.14, { strip: C.PRIMARY, fill: C.WHITE });

      // The zeros ladder: the anchor made visible. Students watch the zeros
      // grow in step with the little number, so the rule is seen, not told.
      const rows = [
        { exp: "1", val: "10" },
        { exp: "2", val: "100" },
        { exp: "3", val: "1 000" },
        { exp: "6", val: "1 000 000" },
      ];
      rows.forEach((r, i) => {
        const ry = y + 0.16 + i * 0.66;
        const last = i === rows.length - 1;
        addPowerText(s, x + 0.22, ry, 0.85, 0.54, "10", r.exp, {
          fontSize: last ? 24 : 20,
          color: last ? C.PRIMARY : C.CHARCOAL,
        });
        s.addText("=", {
          x: x + 1.10, y: ry, w: 0.25, h: 0.54,
          fontSize: 16, fontFace: FONT_B, color: C.MUTED,
          align: "center", valign: "middle", margin: 0,
        });
        s.addText(r.val, {
          x: x + 1.42, y: ry, w: 2.66, h: 0.54,
          fontSize: last ? 24 : 20, fontFace: FONT_H,
          color: last ? C.PRIMARY : C.CHARCOAL, bold: true,
          align: "left", valign: "middle", margin: 0,
        });
      });
      s.addText("the little number = how many zeros", {
        x: x + 0.12, y: y + 2.72, w: w - 0.24, h: 0.32,
        fontSize: 12.5, fontFace: FONT_B, color: C.MUTED,
        align: "center", valign: "middle", margin: 0,
      });
    });

  /* 11. I Do: the ladder */
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "I Do", { color: C.PRIMARY });
    addTitle(s, "Read the little number, not the picture");

    const rungs = [
      { exp: "6", label: "sand in a bag", img: IMG_JIGSAW, fit: "contain", color: C.PRIMARY },
      { exp: "19", label: "sand on Earth", img: IMG_EARTH, color: C.SECONDARY },
      { exp: "24", label: "stars in the universe", img: IMG_STARS, color: C.ACCENT },
      { exp: "26", label: "atoms in a bottle cap", draw: drawBottleCap, color: C.ALERT },
    ];

    const cardW = 2.16;
    const gap = 0.25;
    const startX = 0.5;
    const cardY = CONTENT_TOP;
    rungs.forEach((r, i) => {
      const x = startX + i * (cardW + gap);
      addCard(s, x, cardY, cardW, 3.00, { strip: r.color, fill: C.WHITE });
      addPowerChip(s, x + 0.44, cardY + 0.12, 1.28, 0.52, "10", r.exp, {
        fill: r.color, fontSize: 19,
      });
      s.addText(r.label, {
        x: x + 0.08, y: cardY + 0.70, w: cardW - 0.16, h: 0.44,
        fontSize: 11.5, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    });

    // Visuals placed after the cards so they sit on top of the card fill.
    rungs.forEach((r, i) => {
      const x = startX + i * (cardW + gap);
      if (r.draw) {
        addCard(s, x + 0.14, cardY + 1.20, cardW - 0.28, 1.66, { fill: C.BG_LIGHT });
        r.draw(s, x + 0.42, cardY + 1.76, 1.32, 0.62);
      } else {
        addInstructionalImageCard(s, r.img, {
          x: x + 0.14, y: cardY + 1.20, w: cardW - 0.28, h: 1.66,
          // "contain" on the jigsaw: it is a tall photo and its printed
          // "1 000 000 piece" label is the whole point, so it must not crop off.
          fit: r.fit || "crop", fill: C.BG_LIGHT, pad: 0.06,
        });
      }
    });

    addTextOnShape(s, "Every extra zero is ten times more. More stars than grains of sand on Earth.", {
      x: 0.8, y: 4.42, w: 8.4, h: 0.46, rectRadius: 0.08,
      fill: { color: C.BG_LIGHT },
      line: { color: C.PRIMARY, width: 1.2 },
    }, {
      fontSize: 13.5, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_LADDER);
    runSlideDiagnostics(s, pres);
  }

  /* 12-13. CFU hinge + reveal */
  withReveal(
    () => cfuSlide(pres, "Check", "How much bigger?", "Show Me Boards",
      // Kept short on purpose: the million/billion framing is in the SAY beat,
      // and any longer the option C line runs under the reveal bar.
      "How many times bigger is 10 to the 9 than 10 to the 6?\n\nA: 3 times\nB: 1000 times\nC: 300 times",
      NOTES_CFU, FOOTER),
    (s) => {
      addRevealAnswerBar(s, ["B: 1000 times (3 more zeros)"], { y: 4.30, h: 0.56, fontSize: 18 });
    }
  );

  /* 14. We Do: googol */
  contentSlide(pres, "We Do", C.SECONDARY,
    "A googol",
    [
      "10 with 100 zeros.",
      "Named by a boy aged 9.",
      "Easy to write.",
      "Impossible to have.",
    ],
    NOTES_GOOGOL, FOOTER,
    (s, guide) => {
      const x = guide.rightX;
      const w = guide.rightW;
      const y = guide.panelTopPadded;
      addCard(s, x, y, w, 2.90, { strip: C.SECONDARY, fill: C.WHITE });

      addPowerChip(s, x + 1.24, y + 0.16, 1.82, 0.62, "10", "100", {
        fill: C.SECONDARY, fontSize: 22,
      });
      s.addText("a googol", {
        x: x + 0.12, y: y + 0.84, w: w - 0.24, h: 0.30,
        fontSize: 12.5, fontFace: FONT_B, color: C.MUTED,
        align: "center", valign: "middle", margin: 0,
      });

      addPowerChip(s, x + 1.24, y + 1.26, 1.82, 0.62, "10", "82", {
        fill: C.MUTED, fontSize: 22,
      });
      s.addText("every atom in the universe", {
        x: x + 0.12, y: y + 1.94, w: w - 0.24, h: 0.30,
        fontSize: 12.5, fontFace: FONT_B, color: C.MUTED,
        align: "center", valign: "middle", margin: 0,
      });

      addTextOnShape(s, "Not enough atoms to draw a googol dots.", {
        x: x + 0.20, y: y + 2.32, w: w - 0.40, h: 0.40, rectRadius: 0.06,
        fill: { color: C.ALERT },
      }, { fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true });
    });

  /* 15. We Do: googolplex and the tower */
  contentSlide(pres, "We Do", C.SECONDARY,
    "When the zeros run out",
    [
      "A googolplex is 10 with a googol zeros.",
      "The little number grew its own little number.",
      "So we stack them instead.",
    ],
    NOTES_TOWER, FOOTER,
    (s, guide) => {
      const x = guide.rightX;
      const w = guide.rightW;
      const y = guide.panelTopPadded;
      addCard(s, x, y, w, 2.72, { strip: C.SECONDARY, fill: C.WHITE });

      s.addText("a googolplex", {
        x: x + 0.12, y: y + 0.14, w: w - 0.24, h: 0.30,
        fontSize: 12.5, fontFace: FONT_B, color: C.MUTED,
        align: "center", valign: "middle", margin: 0,
      });
      drawPowerTower(s, x + 1.20, y + 1.90, ["10", "10", "100"], {
        fontSize: 34, color: C.SECONDARY,
      });
      addTextOnShape(s, "Too many zeros to write, even with every atom there is.", {
        x: x + 0.20, y: y + 2.14, w: w - 0.40, h: 0.42, rectRadius: 0.06,
        fill: { color: C.ALERT },
      }, { fontSize: 11.5, fontFace: FONT_B, color: C.WHITE, bold: true });
    });

  /* 16-17. We Do: the pattern + reveal */
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.SECONDARY);
      addBadge(s, "We Do", { color: C.SECONDARY });
      addTitle(s, "What comes next?");

      const rows = [
        { doing: "3 + 3 + 3 + 3", name: "adding again and again", short: "3 x 4", color: C.PRIMARY },
        { doing: "3 x 3 x 3 x 3", name: "multiplying again and again", short: "3 to the 4", power: ["3", "4"], color: C.SECONDARY },
      ];
      const rowH = 0.78;
      rows.forEach((r, i) => {
        const ry = CONTENT_TOP + i * (rowH + 0.16);
        addCard(s, 0.5, ry, 9.0, rowH, { strip: r.color, fill: C.WHITE });
        s.addText(r.doing, {
          x: 0.72, y: ry + 0.06, w: 2.5, h: rowH - 0.12,
          fontSize: 19, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
          align: "left", valign: "middle", margin: 0,
        });
        s.addText(r.name, {
          x: 3.30, y: ry + 0.06, w: 3.3, h: rowH - 0.12,
          fontSize: 12.5, fontFace: FONT_B, color: C.MUTED,
          align: "center", valign: "middle", margin: 0,
        });
        if (r.power) {
          // Real exponent notation on the shortcut chip, not "3 to the 4".
          s.addShape("roundRect", {
            x: 6.75, y: ry + 0.16, w: 2.5, h: rowH - 0.32, rectRadius: 0.06,
            fill: { color: r.color }, line: { color: r.color, width: 1 },
          });
          s.addText([
            { text: "so we write ", options: { fontSize: 12.5, fontFace: FONT_B, color: C.WHITE, bold: true } },
            { text: r.power[0], options: { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true } },
            { text: r.power[1], options: { fontSize: 15, fontFace: FONT_H, color: C.WHITE, bold: true, superscript: true } },
          ], {
            x: 6.75, y: ry + 0.16, w: 2.5, h: rowH - 0.32,
            align: "center", valign: "middle", margin: 0,
          });
        } else {
          addTextOnShape(s, `so we write ${r.short}`, {
            x: 6.75, y: ry + 0.16, w: 2.5, h: rowH - 0.32, rectRadius: 0.06,
            fill: { color: r.color },
          }, { fontSize: 12.5, fontFace: FONT_B, color: C.WHITE, bold: true });
        }
      });

      const qy = CONTENT_TOP + 2 * (rowH + 0.16);
      addCard(s, 0.5, qy, 9.0, rowH, { strip: C.ALERT, fill: C.WHITE });
      // Row 3's "doing" is a power tower, so draw one rather than describing it
      // in words. It has to look like the thing the pattern is pointing at.
      drawPowerTower(s, 0.76, qy + 0.70, ["3", "3", "3", "3"], {
        fontSize: 16, color: C.CHARCOAL,
      });
      s.addText("doing POWERS again and again", {
        x: 3.30, y: qy + 0.06, w: 3.3, h: rowH - 0.12,
        fontSize: 12.5, fontFace: FONT_B, color: C.MUTED,
        align: "center", valign: "middle", margin: 0,
      });
      addTextOnShape(s, "so we write ... ?", {
        x: 6.75, y: qy + 0.16, w: 2.5, h: rowH - 0.32, rectRadius: 0.06,
        fill: { color: C.ALERT },
      }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_PATTERN);
      runSlideDiagnostics(s, pres);
      return s;
    },
    (s) => {
      addRevealAnswerBar(s, ["3 ↑↑ 4", "every arrow is a whole new level"],
        { y: 4.30, h: 0.56, fontSize: 17 });
    }
  );

  /* 18. Graham's number */
  contentSlide(pres, "I Do", C.PRIMARY,
    "Graham's number",
    [
      "3 ↑↑↑ 3 is a tower of 3s, 7 trillion levels tall.",
      "3 ↑↑↑↑ 3 is only step 1 of 64.",
      "We know it ends in 7. Nobody knows the start.",
    ],
    NOTES_GRAHAM, FOOTER,
    (s, guide) => {
      const x = guide.rightX;
      const w = guide.rightW;
      const y = guide.panelTopPadded;
      addCard(s, x, y, w, 1.42, { strip: C.PRIMARY, fill: C.WHITE });

      // Real up-arrow notation on the slide face - inventing this symbol IS the
      // idea of the lesson. Teacher notes say "arrow" in words to stay ASCII.
      const steps = [
        { txt: "3 ↑ 3", val: "27" },
        { txt: "3 ↑↑ 3", val: "7 625 597 484 987" },
      ];
      steps.forEach((st, i) => {
        const ry = y + 0.14 + i * 0.58;
        s.addText(st.txt, {
          x: x + 0.16, y: ry, w: 0.94, h: 0.44,
          fontSize: 15, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
          align: "left", valign: "middle", margin: 0,
        });
        s.addText("=", {
          x: x + 1.10, y: ry, w: 0.22, h: 0.44,
          fontSize: 13, fontFace: FONT_B, color: C.MUTED,
          align: "center", valign: "middle", margin: 0,
        });
        s.addText(st.val, {
          x: x + 1.34, y: ry, w: 2.80, h: 0.44,
          fontSize: 13.5, fontFace: FONT_H, color: C.SECONDARY, bold: true,
          align: "left", valign: "middle", margin: 0,
        });
      });

      addInstructionalImageCard(s, IMG_BRAIN, {
        x, y: y + 1.56, w, h: 1.54, fit: "crop", strip: C.ALERT,
        caption: "Hold every digit in your head and it collapses into a black hole.",
      });
    });

  /* 19. You Do */
  contentSlide(pres, "You Do", C.ALERT,
    "Order the universe",
    [
      "First: count the zeros.",
      "Next: order the numbers.",
      "Then: how much bigger?",
    ],
    NOTES_YOUDO, FOOTER,
    (s, guide) => {
      const x = guide.rightX;
      const w = guide.rightW;
      const y = guide.panelTopPadded;
      addInstructionCard(s, [
        { text: "Your only tool", role: "header" },
        { text: "", role: "spacer" },
        { text: "the little number", role: "body" },
        { text: "", role: "spacer" },
        { text: "Ignore what the things are.", role: "emphasis" },
        { text: "", role: "spacer" },
        { text: "12 minutes.", role: "emphasis" },
      ], {
        x, y, w, h: 2.90,
        strip: C.ALERT, fill: C.WHITE,
      });
    });

  /* 20. Exit ticket */
  exitTicketSlide(pres,
    "There are 10 to the 24 stars in the universe.\nA bottle cap of water holds 10 to the 26 atoms.\n\nWhich is bigger, and how many times bigger? How do you know?",
    NOTES_EXIT, FOOTER);

  /* 21. Closing */
  closingSlide(pres, {
    reflectionPrompt: "What is the biggest number you can now write in five seconds?",
    scItems: [
      "I can say what the little raised number tells me.",
      "I can say which of two powers of 10 is bigger, and how many times bigger.",
      "I can explain why Graham's number needed a new way to write.",
    ],
    selfAssessment: "Thumbs up / sideways / down",
    takeaways: [
      "The little number counts the zeros. One more zero is ten times more.",
    ],
  }, NOTES_CLOSING);

  const fileName = path.join(OUT_DIR, "Big Ideas - Session 2 - Numbers Too Big To Write.pptx");
  await pres.writeFile({ fileName });
  console.log("PPTX written to " + fileName);

  await writeWorksheet();
  await writeAnswerKey();
  console.log("PDFs written to " + RES_DIR);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
