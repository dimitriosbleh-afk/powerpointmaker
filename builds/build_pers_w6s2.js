"use strict";

// Persuasive Writing - Term 3 Week 6, Session 2 (Year 5/6 Literacy)
// "Voice, part 1: emotive and precise words"
//
// Victorian Curriculum 2.0: English, Language (Expressing and developing ideas),
// Levels 5-6 - how word choice positions a reader and creates an emotional
// response.
//
// UNIT ANCHOR (locked): Bold Beginning. Mighty Middle. Excellent Ending.
// Voice all the way through. Today works inside the VOICE band.
//
// Lesson shape: example-first with a live upgrade model. Word banks are the
// teacher's own, supplied in the planner.

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme } = require("../themes/factory");
const { composeGlanceNotes } = require("../themes/core/composeNotes");
const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  PAGE, hex,
} = require("../themes/pdf_helpers");

const P = require("./persuasive_lib");

const UNIT_VARIANT = 5;
const T = createTheme("literacy", "grade56", UNIT_VARIANT);
const {
  C, FONT_H, FONT_B, SAFE_BOTTOM, CONTENT_TOP,
  titleSlide, liSlide, keyWordSlide, closingSlide,
  addCard, addFooter, addTextOnShape, addRevealAnswerBar,
  clickBuild, runSlideDiagnostics,
} = T;

const SESSION = 2;
const FOOTER = "Persuasive Writing | Week 6 Session 2 | Year 5/6 Literacy";
const OUT_DIR = "output/Persuasive_W6_S2";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION));

const BANK_RES = makeSessionResource(SESSION,
  "Persuasive Word Bank",
  "The full word bank. Print once, keep in books, use for both weeks.");
const UPGRADE_RES = makeSessionResource(SESSION,
  "Sentence Upgrade",
  "Worked example, then six flat sentences to upgrade using the bank.");
const RESOURCE_ITEMS = [BANK_RES, UPGRADE_RES];
fs.mkdirSync(RES_DIR, { recursive: true });

const FLAT = "The bin is a good idea because wasting paper is bad.";
const UPGRADED = "A scrap paper bin is an outstanding idea, because wasting paper is appalling.";

// ─── Teacher notes ──────────────────────────────────────────────────────────

const NOTES_TITLE =
  "Session 2. One line: last session we built the shape, today we make the words do the work.";

const NOTES_RESOURCES =
  "Prep slide. Print the Persuasive Word Bank once per student - it is the reference for both weeks and is not reprinted.\n" +
  "Print the Sentence Upgrade sheet for today. Whiteboards and markers out.\n" +
  "CATCH-UP: missed Session 1? Hand them the Structure Card, say the anchor phrase, and start them at the launch. Today needs no Session 1 content.";

const NOTES_OVERVIEW =
  "For you, not the class. Skip past it when projecting.\n" +
  "Today sits inside the Voice band of the anchor. Two tools only: precise words and emotive words. Questions and bossy verbs wait for Session 3.\n" +
  "Decision points: the strongest-pull hinge, the boards check in the We Do, and the exit ticket. Between them keep the pace brisk.";

const NOTES_LAUNCH = composeGlanceNotes({
  answer: "the second one - it makes you feel something.",
  beats: [
    [
      "SAY: Two sentences. Same fact. One of them lands harder.",
      "Read both aloud, the second with real feeling.",
    ],
    [
      "ASK: Which one makes you want to do something about it?",
      "15 sec. Fingers at your chest, show me one or two.",
      "EXPECT: two.",
    ],
    [
      "SCAN the fingers, back row first.",
      "80%+ on two -> cold call one: which WORD did that?",
      "Less -> read the first flat, the second angry, re-ask.",
    ],
  ],
  trap: [
    "thinking the longer sentence is automatically stronger.",
    "Fix: point at one word, student names the feeling it makes.",
  ],
  prep: "Low-coupling launch: needs only the ability to compare two sentences, so a student who missed Session 1 starts level.",
  tag: "[Launch | Retention and recall | HITS 6]",
});

const NOTES_LI = composeGlanceNotes({
  beats: [
    "POINT to the intention. SAY: Today the words themselves do the persuading.",
    "SAY: Read the I can statements with me. Everyone, together, on three.",
    "SAY: Everyone leaves able to spot an emotive word.",
  ],
  prep: "SC1 reachable by all; SC2 is the core target the exit ticket assesses; SC3 stretches into explaining the effect.",
  tag: "[LI/SC | Planning made visible | HITS 1]",
});

const NOTES_VOCAB = composeGlanceNotes({
  answer: "emotive words make the reader feel a certain way.",
  beats: [
    [
      "SAY: Emotive words make the reader feel something.",
      "Not just know something. Feel it.",
    ],
    [
      "ASK: Which feels worse, untidy or repulsive?",
      "10 sec. Choral response: Everyone, together, on three.",
      "EXPECT: repulsive.",
    ],
    "SAY: Same mess. Different feeling. That is the writer choosing.",
  ],
  trap: [
    "thinking emotive means angry.",
    "Fix: say courageous and magic, student names those feelings.",
  ],
  prep: "The teacher's own definition from the planner: emotive language makes the reader feel a certain way.",
  sources: "Persuasive word banks supplied by the teacher.",
  tag: "[Key Vocabulary | Knowledge and memory | HITS 6]",
});

const NOTES_IDO = composeGlanceNotes({
  answer: "good becomes outstanding, bad becomes appalling.",
  beats: [
    "SAY: Watch me upgrade this. I am not adding words. I am swapping them.",
    [
      "CLICK 1. SAY: Good becomes outstanding. That is precise.",
      "CLICK 2. SAY: Bad becomes appalling. That one you feel.",
    ],
    [
      "CLICK 3 for the finished sentence. ASK: What changed?",
      "15 sec. Turn and tell. Partner B first.",
      "EXPECT: two words, not the whole sentence.",
    ],
  ],
  trap: [
    "rewriting the whole sentence instead of swapping words.",
    "Fix: cover all but the two swaps, student names them.",
  ],
  stretch: "Find a third word worth upgrading.",
  help: "Give the bank open at the GOOD and BAD columns.",
  prep: "Model the swap, not a rewrite. Students who rewrite lose the point and run out of time.",
  tag: "[I Do | Explicit teaching | SC2 | HITS 4]",
});

const NOTES_EMOTIVE_BANK = composeGlanceNotes({
  answer: "open - any word that names a feeling the reader gets.",
  beats: [
    [
      "SAY: This is your emotive bank. It is on your printed sheet too.",
      "You do not need all of them. You need three you love.",
    ],
    [
      "ASK: Pick one word you would use about wasted paper.",
      "20 sec. Write it... chin it... show me.",
      "EXPECT: disaster, shame, tragic, outrage or similar.",
    ],
    [
      "SCAN boards for a feeling word, not a describing word.",
      "80%+ -> cold call two: what feeling does yours make?",
      "Less -> read three words aloud with feeling, re-ask.",
    ],
  ],
  trap: [
    "picking a word that fits the topic but carries no feeling.",
    "Fix: ask what the reader feels, student swaps the word.",
  ],
  prep: "Do not read the whole bank aloud. Three words chosen well beat forty skimmed.",
  sources: "Persuasive word banks supplied by the teacher.",
  tag: "[I Do | Knowledge and memory | SC1 | HITS 4]",
});

const NOTES_PRECISE_BANK = composeGlanceNotes({
  answer: "open - any upgrade that says HOW good, bad or big.",
  beats: [
    [
      "POINT to the flat words on the left.",
      "SAY: These four words are the laziest words in English.",
    ],
    [
      "ASK: Give me a better word for BAD.",
      "15 sec. Choral response: Everyone, together, on three.",
      "EXPECT: appalling, dreadful, outrageous.",
    ],
    "SAY: Precise words tell the reader HOW bad. Emotive words make them feel it.",
  ],
  trap: [
    "treating precise and emotive as the same tool.",
    "Fix: name one of each from the slide, student says the difference.",
  ],
  prep: "The full ladders for all fifteen flat words are on the printed bank. Four on screen is enough.",
  sources: "Persuasive word banks supplied by the teacher.",
  tag: "[I Do | Knowledge and memory | SC1 | HITS 4]",
});

const NOTES_CFU = composeGlanceNotes({
  answer: "B. It names a feeling and makes the reader care.",
  beats: [
    [
      "SAY: Three sentences, same fact. One pulls hardest.",
      "Do not call out. Boards for this one.",
    ],
    [
      "ASK: Which sentence pulls hardest, A, B or C?",
      "30 sec. Write it... chin it... show me.",
      "EXPECT: B.",
    ],
    [
      "SCAN boards, back row first.",
      "80%+ -> cold call one B: which word did the work? Then reveal.",
      "Less -> read A and B again, ask which made you care, re-ask.",
    ],
    "REVEAL after boards are scanned. SAY: Appalling and tragedy. Two words.",
  ],
  trap: [
    "picking C because it sounds factual and grown up.",
    "Fix: ask what C makes you feel, student re-tests all three.",
  ],
  prep: "The hinge. A is flat, C is neutral fact. Both wrong answers show a student not yet hearing the feeling in a word.",
  tag: "[CFU hinge | Supported application | SC2 | HITS 7, 8]",
});

const NOTES_WEDO = composeGlanceNotes({
  answer: "open - one precise swap and one emotive swap.",
  beats: [
    "SAY: I take BIG. You take BAD. Two swaps only, no rewrites.",
    [
      "ASK: Now swap BAD yourself.",
      "45 sec. Write it... chin it... show me.",
      "EXPECT: an emotive word in place of bad.",
    ],
    [
      "SCAN boards for a swap, not a rewrite.",
      "80%+ -> cold call one strong, one shaky: what feeling?",
      "Less -> rebuild one board using the bank, re-ask.",
    ],
  ],
  trap: [
    "swapping in a word that does not fit.",
    "Fix: read it back aloud, student hears it and fixes it.",
  ],
  stretch: "Swap a third word and say why it is stronger.",
  help: "Point to the bank column, offer two words to pick between.",
  prep: "Same topic as Session 1 on purpose. The topic is familiar so all the thinking goes on word choice.",
  tag: "[We Do | Supported application | SC2 | HITS 4, 5]",
});

const NOTES_YOUDO = composeGlanceNotes({
  answer: "open - each sentence needs at least one precise or emotive swap.",
  beats: [
    [
      "SAY: Three new sentences. Nothing to do with paper.",
      "SAY: Upgrade at least one word in each. Use your bank.",
    ],
    [
      "POINT to the steps. SAY: Read it, find the lazy word, swap it.",
      "TIME: ten minutes on the Sentence Upgrade sheet.",
    ],
    [
      "CIRCULATE. Look for swaps, not rewrites.",
      "COLLECT two strong upgrades to read at the closing.",
    ],
  ],
  trap: [
    "upgrading every word until the sentence stops making sense.",
    "Fix: ask them to read it aloud, student cuts back to two swaps.",
  ],
  stretch: "Write a fourth sentence of your own, then upgrade it.",
  help: "Circle the one word to swap in each sentence before they start.",
  prep: "New topics on purpose, so this is transfer and not a repeat of the We Do. Sheet, not books.",
  tag: "[You Do | Mastery and application | SC2 | HITS 10]",
});

const NOTES_EXIT = composeGlanceNotes({
  answer: "any upgrade of bad to an emotive or precise word that fits.",
  beats: [
    [
      "SAY: One sentence. Upgrade at least one word.",
      "SAY: Underline the word you changed.",
    ],
    [
      "TIME: three minutes. Sheet on the desk when you finish.",
      "COLLECT by walking the rows.",
    ],
    [
      "SORT as you collect: swapped and underlined, or not swapped.",
      "A thick second pile -> reopen the bank at the start of Session 3.",
    ],
  ],
  prep: "Assesses the core target: swap a flat word for a precise or emotive one. The underline makes the evidence scannable.",
  tag: "[Exit Ticket | Mastery and application | SC2 | HITS 8]",
});

const NOTES_CLOSING = composeGlanceNotes({
  answer: "open - most students should show two.",
  beats: [
    "READ the two collected upgrades aloud. SAY: Listen for the swapped word.",
    [
      "ASK: Rate yourself on the three statements.",
      "15 sec. Fingers at your chest. One, two or three. Show me.",
      "EXPECT: most on two.",
    ],
    "SAY: Next session we hook the reader with questions and commands.",
  ],
  prep: "Anyone showing one gets the bank open in front of them next session, at the column they need.",
  tag: "[Closing | Retention and recall | HITS 1, 9]",
});

// ─── Build ──────────────────────────────────────────────────────────────────

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";

  titleSlide(pres, "Words That Persuade",
    "Voice, part 1: emotive and precise words",
    "Week 6 Session 2 | Year 5/6 Literacy", NOTES_TITLE);

  addResourceSlide(pres, RESOURCE_ITEMS, { C, FONT_H, FONT_B }, FOOTER, NOTES_RESOURCES);

  // Teacher overview
  (() => {
    const s = P.customSlide(pres, T, "For the teacher", C.MUTED, "Session 2 at a glance", { badgeW: 2.4 });
    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;

    addCard(s, 0.5, cardY, 4.4, cardH, { strip: C.PRIMARY });
    s.addText("Where this sits", {
      x: 0.7, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Today lives inside the VOICE band of the anchor.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: P.ANCHOR_PHRASE, options: { color: C.PRIMARY, bold: true, breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Two tools only today:", options: { bold: true, breakLine: true } },
      { text: "precise words (how good, how bad)", options: { breakLine: true } },
      { text: "emotive words (what the reader feels)", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Questions, bossy verbs and promises wait for Session 3. Do not preview them.", options: { italic: true, color: C.MUTED } },
    ], {
      x: 0.7, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addCard(s, 5.1, cardY, 4.4, cardH, { strip: C.ACCENT });
    s.addText("How today works", {
      x: 5.3, y: cardY + 0.12, w: 4.0, h: 0.28,
      fontSize: 14, fontFace: FONT_H, color: C.ACCENT, bold: true, margin: 0,
    });
    s.addText([
      { text: "Shape: example-first. One live upgrade model, then two bank tours, then transfer.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Decision points:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "the strongest-pull hinge, the boards check in the We Do, and the exit ticket.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "The trap all lesson:", options: { bold: true, color: C.ACCENT, breakLine: true } },
      { text: "students rewriting whole sentences instead of swapping single words. Keep pulling them back to the swap.", options: { breakLine: true } },
      { text: " ", options: { fontSize: 6, breakLine: true } },
      { text: "Flex: cut one bank tour if the launch runs long.", options: { italic: true, color: C.MUTED } },
    ], {
      x: 5.3, y: cardY + 0.46, w: 4.0, h: cardH - 0.60,
      fontSize: 12.5, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0, paraSpaceAfter: 3,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_OVERVIEW);
    runSlideDiagnostics(s, pres);
  })();

  // Launch - two sentences, same fact
  (() => {
    const s = P.customSlide(pres, T, "Launch", C.SECONDARY, "Same fact. Which one lands?");
    const y0 = CONTENT_TOP;
    const pairs = [
      { key: "1", text: "There is a lot of paper waste in our classrooms." },
      { key: "2", text: "Every day we throw away paper that could have been used again. It is a shameful waste." },
    ];
    const ch = 1.15;
    pairs.forEach((p, i) => {
      const cy = y0 + i * (ch + 0.20);
      s.addShape("roundRect", {
        x: 0.5, y: cy, w: 9, h: ch, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.SECONDARY, width: 1.5 },
      });
      s.addShape("roundRect", {
        x: 0.72, y: cy + 0.33, w: 0.5, h: 0.5, rectRadius: 0.25,
        fill: { color: C.SECONDARY },
      });
      s.addText(p.key, {
        x: 0.72, y: cy + 0.33, w: 0.5, h: 0.5,
        fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(p.text, {
        x: 1.42, y: cy, w: 7.85, h: ch,
        fontSize: 20, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addTextOnShape(s, "Fingers at your chest. Show me one or two.", {
      x: 2.6, y: y0 + 2.72, w: 4.8, h: 0.62, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.SECONDARY, width: 1.5 },
    }, {
      fontSize: 17, fontFace: FONT_B, color: C.SECONDARY, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
    runSlideDiagnostics(s, pres);
  })();

  liSlide(pres,
    "We are learning how word choice makes a reader feel what we want them to feel.",
    [
      "I can find the emotive words in a sentence.",
      "I can swap a flat word for a precise or emotive word.",
      "I can explain what feeling my word choice creates.",
    ],
    NOTES_LI, FOOTER);

  keyWordSlide(pres, {
    word: "emotive",
    meaning: "Emotive words make the reader feel a certain way.",
    example: "Untidy tells you. Repulsive makes you feel it.",
    routine: ["Say it", "Feel it", "Use it"],
    color: C.SECONDARY,
    title: "The word for today",
  }, NOTES_VOCAB, FOOTER);

  // I Do - live upgrade
  (() => {
    const s = P.customSlide(pres, T, "I Do", C.PRIMARY, "Two swaps, not a rewrite");
    const y0 = CONTENT_TOP;

    addCard(s, 0.5, y0, 9, 1.02, { strip: C.MUTED });
    s.addText("Flat", {
      x: 0.75, y: y0 + 0.08, w: 2, h: 0.26,
      fontSize: 14, fontFace: FONT_B, color: C.MUTED, bold: true, margin: 0,
    });
    s.addText([
      { text: "The bin is a ", options: { color: C.CHARCOAL } },
      { text: "good", options: { color: C.ALERT, bold: true } },
      { text: " idea because wasting paper is ", options: { color: C.CHARCOAL } },
      { text: "bad", options: { color: C.ALERT, bold: true } },
      { text: ".", options: { color: C.CHARCOAL } },
    ], {
      x: 0.75, y: y0 + 0.36, w: 8.5, h: 0.56,
      fontSize: 22, fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const swapY = y0 + 1.18;
    clickBuild(s, [
      () => {
        s.addShape("roundRect", {
          x: 0.5, y: swapY, w: 4.4, h: 0.78, rectRadius: 0.08,
          fill: { color: C.SECONDARY },
        });
        s.addText("good  ->  outstanding", {
          x: 0.6, y: swapY, w: 4.2, h: 0.78,
          fontSize: 21, fontFace: FONT_B, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
        s.addText("precise word", {
          x: 0.5, y: swapY + 0.82, w: 4.4, h: 0.30,
          fontSize: 13, fontFace: FONT_B, color: C.SECONDARY, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      },
      () => {
        s.addShape("roundRect", {
          x: 5.1, y: swapY, w: 4.4, h: 0.78, rectRadius: 0.08,
          fill: { color: C.ASSESS },
        });
        s.addText("bad  ->  appalling", {
          x: 5.2, y: swapY, w: 4.2, h: 0.78,
          fontSize: 21, fontFace: FONT_B, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
        s.addText("emotive word", {
          x: 5.1, y: swapY + 0.82, w: 4.4, h: 0.30,
          fontSize: 13, fontFace: FONT_B, color: C.ASSESS, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      },
      () => {
        addRevealAnswerBar(s, [UPGRADED], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 19, label: "Upgraded", color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO);
    runSlideDiagnostics(s, pres);
  })();

  // Emotive bank tour
  (() => {
    const s = P.customSlide(pres, T, "Word Bank", C.ASSESS, "Emotive words: make them feel it");
    const y0 = CONTENT_TOP;
    const picked = [
      "outrage", "disaster", "shame", "tragic", "appalling", "disgust",
      "cruel", "damaging", "urge", "demand", "innocent", "freedom",
      "magnificent", "courageous", "miracle", "phenomenal", "terrified", "wicked",
    ];
    const endY = P.drawWordChips(s, T, {
      x: 0.5, y: y0, w: 9, words: picked, perRow: 6,
      chipH: 0.62, gap: 0.12, fill: C.BG_LIGHT, fontSize: 18,
    });

    addTextOnShape(s, "Pick ONE for wasted paper. Write it... chin it... show me.", {
      x: 0.5, y: endY + 0.24, w: 9, h: 0.66, rectRadius: 0.08,
      fill: { color: C.ASSESS },
    }, {
      fontSize: 18, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    s.addText("The full bank of 44 words is on your printed sheet.", {
      x: 0.5, y: endY + 0.98, w: 9, h: 0.36,
      fontSize: 14, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_EMOTIVE_BANK);
    runSlideDiagnostics(s, pres);
  })();

  // Precise words tour
  (() => {
    const s = P.customSlide(pres, T, "Word Bank", C.SECONDARY, "Precise words: say HOW good, HOW bad");
    const y0 = CONTENT_TOP;
    const ladders = [
      P.PRECISE_WORDS.find((g) => g.flat === "GOOD"),
      P.PRECISE_WORDS.find((g) => g.flat === "BAD"),
      P.PRECISE_WORDS.find((g) => g.flat === "BIG"),
      P.PRECISE_WORDS.find((g) => g.flat === "SCARED"),
    ];
    const rowH = 0.80;
    const gap = 0.14;
    ladders.forEach((group, i) => {
      const ry = y0 + i * (rowH + gap);
      s.addShape("roundRect", {
        x: 0.5, y: ry, w: 1.65, h: rowH, rectRadius: 0.08,
        fill: { color: C.MUTED },
      });
      s.addText(group.flat, {
        x: 0.5, y: ry, w: 1.65, h: rowH,
        fontSize: 19, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText("->", {
        x: 2.22, y: ry, w: 0.42, h: rowH,
        fontSize: 19, fontFace: FONT_B, color: C.SECONDARY, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      P.drawWordChips(s, T, {
        x: 2.72, y: ry + 0.09, w: 6.78,
        words: group.better.slice(0, 4), perRow: 4,
        chipH: rowH - 0.18, gap: 0.10, fill: C.WHITE, fontSize: 16,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_PRECISE_BANK);
    runSlideDiagnostics(s, pres);
  })();

  // CFU hinge
  (() => {
    const s = P.customSlide(pres, T, "CFU", C.ALERT, "Which sentence pulls hardest?");
    const y0 = CONTENT_TOP;

    addTextOnShape(s, "Topic: litter in the school yard", {
      x: 0.5, y: y0, w: 9, h: 0.48, rectRadius: 0.08,
      fill: { color: C.ALERT },
    }, {
      fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const options = [
      { key: "A", text: "There is quite a bit of litter around and it is not very good." },
      { key: "B", text: "Our yard is drowning in litter, and every wrapper is a shameful waste." },
      { key: "C", text: "Litter is found in several areas of the school grounds each week." },
    ];
    const oh = 0.66;
    const og = 0.12;
    options.forEach((opt, i) => {
      const oy = y0 + 0.60 + i * (oh + og);
      s.addShape("roundRect", {
        x: 0.5, y: oy, w: 9, h: oh, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.PRIMARY, width: 1.5 },
      });
      s.addShape("roundRect", {
        x: 0.68, y: oy + 0.11, w: 0.44, h: 0.44, rectRadius: 0.22,
        fill: { color: C.PRIMARY },
      });
      s.addText(opt.key, {
        x: 0.68, y: oy + 0.11, w: 0.44, h: 0.44,
        fontSize: 17, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText(opt.text, {
        x: 1.28, y: oy, w: 7.97, h: oh,
        fontSize: 16.5, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    clickBuild(s, [
      () => {
        addRevealAnswerBar(s, ["B - drowning, shameful, waste"], {
          y: P.REVEAL_Y, h: P.REVEAL_H, fontSize: 22, color: C.SUCCESS,
        });
      },
    ]);

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU);
    runSlideDiagnostics(s, pres);
  })();

  // We Do
  (() => {
    const s = P.customSlide(pres, T, "We Do", C.SUCCESS, "One swap each");
    const y0 = CONTENT_TOP;

    addCard(s, 0.5, y0, 9, 1.15, { strip: C.SUCCESS });
    s.addText("Our sentence", {
      x: 0.75, y: y0 + 0.09, w: 3, h: 0.28,
      fontSize: 14, fontFace: FONT_B, color: C.SUCCESS, bold: true, margin: 0,
    });
    s.addText([
      { text: "Our classroom throws out a ", options: { color: C.CHARCOAL } },
      { text: "big", options: { color: C.ALERT, bold: true } },
      { text: " pile of paper, and that is ", options: { color: C.CHARCOAL } },
      { text: "bad", options: { color: C.ALERT, bold: true } },
      { text: ".", options: { color: C.CHARCOAL } },
    ], {
      x: 0.75, y: y0 + 0.40, w: 8.5, h: 0.66,
      fontSize: 22, fontFace: FONT_B, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const sy = y0 + 1.32;
    addTextOnShape(s, "I take BIG. You take BAD.", {
      x: 0.5, y: sy, w: 4.4, h: 0.82, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    }, {
      fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    addTextOnShape(s, "big  ->  colossal", {
      x: 5.1, y: sy, w: 4.4, h: 0.82, rectRadius: 0.08,
      fill: { color: C.SECONDARY },
    }, {
      fontSize: 20, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addTextOnShape(s, "Your swap for BAD. Write it... chin it... show me.", {
      x: 0.5, y: sy + 1.00, w: 9, h: 0.74, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.SUCCESS, width: 1.5 },
    }, {
      fontSize: 19, fontFace: FONT_B, color: C.SUCCESS, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
    runSlideDiagnostics(s, pres);
  })();

  // You Do
  (() => {
    const s = P.customSlide(pres, T, "You Do", C.ASSESS, "Upgrade three sentences");
    const y0 = CONTENT_TOP;
    const sentences = [
      "School lunches are not very nice.",
      "Too much screen time is bad for kids.",
      "The library is a good place to be.",
    ];
    const ch = 0.72;
    sentences.forEach((sentence, i) => {
      const cy = y0 + i * (ch + 0.14);
      s.addShape("roundRect", {
        x: 0.5, y: cy, w: 6.1, h: ch, rectRadius: 0.08,
        fill: { color: C.WHITE }, line: { color: C.ASSESS, width: 1.5 },
      });
      s.addText(sentence, {
        x: 0.74, y: cy, w: 5.62, h: ch,
        fontSize: 17, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0, fit: "shrink", shrinkText: true,
      });
    });

    addCard(s, 6.8, y0, 2.7, 2.44, { strip: C.ASSESS });
    s.addText("Steps", {
      x: 7.0, y: y0 + 0.10, w: 2.3, h: 0.30,
      fontSize: 15, fontFace: FONT_B, color: C.ASSESS, bold: true, margin: 0,
    });
    ["Read it aloud.", "Find the lazy word.", "Swap it. Underline it."].forEach((step, i) => {
      s.addText(String(i + 1) + ". " + step, {
        x: 7.0, y: y0 + 0.50 + i * 0.62, w: 2.3, h: 0.58,
        fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "middle", margin: 0,
      });
    });

    addTextOnShape(s, "Sentence Upgrade sheet. Ten minutes. Use your bank.", {
      x: 0.5, y: y0 + 2.64, w: 9, h: 0.76, rectRadius: 0.08,
      fill: { color: C.ASSESS },
    }, {
      fontSize: 20, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
    runSlideDiagnostics(s, pres);
  })();

  P.exitTicketPanel(pres, T, {
    topic: "Litter in the yard is bad.",
    task: "Upgrade one word. Underline what you changed.",
    cue: "Precise, or emotive. Your choice.",
    footer: FOOTER,
    notes: NOTES_EXIT,
  });

  closingSlide(pres, {
    reflectionPrompt: "Which word did you swap today that you will use again?",
    scItems: [
      "I can find the emotive words in a sentence.",
      "I can swap a flat word for a precise or emotive word.",
      "I can explain what feeling my word choice creates.",
    ],
    selfAssessment: "Fingers at your chest. One, two or three. Show me.",
    takeaways: ["Precise words say HOW. Emotive words make them FEEL."],
  }, NOTES_CLOSING);

  const pptxPath = path.join(OUT_DIR, "Persuasive Writing W6 S2.pptx");
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to " + pptxPath);
}

// ─── Resource: the full Persuasive Word Bank ────────────────────────────────

async function buildBank() {
  const doc = createPdf({ title: BANK_RES.name });
  let y = addPdfHeader(doc, "Persuasive Word Bank", {
    subtitle: "Keep this for the whole fortnight. You will use it in every session.",
    color: C.PRIMARY,
    lessonInfo: "Year 5/6 Literacy | Persuasive Writing",
    showNameDate: false,
  });

  y = addTipBox(doc,
    "You do not need every word. Find three you love in each list and make them yours.",
    y, { color: C.ACCENT });
  y += 4;

  y = P.addWordChipsPdf(doc, y, "Emotive language: makes the reader feel a certain way",
    P.EMOTIVE_WORDS, { color: C.ASSESS, perRow: 6, chipH: 19, fontSize: 9 });

  y = P.addWordChipsPdf(doc, y, "Rhetorical question starters",
    P.RHETORICAL_STARTERS, { color: C.SECONDARY, perRow: 3, chipH: 20, fontSize: 8.5 });

  y = P.addWordChipsPdf(doc, y, "Bossy verbs: tell the reader what to do",
    P.BOSSY_VERBS, { color: C.PRIMARY, perRow: 3, chipH: 20, fontSize: 9 });

  y = P.addWordChipsPdf(doc, y, "Exaggeration and promises",
    P.EXAGGERATION_PROMISES, { color: C.PRIMARY, perRow: 3, chipH: 20, fontSize: 8.5 });

  y = P.addWordChipsPdf(doc, y, "Low modality: sounds unsure",
    P.MODALITY.low, { color: C.MUTED, perRow: 6, chipH: 19, fontSize: 9 });
  y = P.addWordChipsPdf(doc, y, "Medium modality: sounds likely",
    P.MODALITY.medium, { color: C.SECONDARY, perRow: 6, chipH: 19, fontSize: 9 });
  y = P.addWordChipsPdf(doc, y, "High modality: sounds certain",
    P.MODALITY.high, { color: C.ALERT, perRow: 6, chipH: 19, fontSize: 9 });

  doc.addPage();
  y = PAGE.MARGIN;
  y = addSectionHeading(doc, "Precise words: swap a lazy word for one that says HOW", y,
    { color: C.PRIMARY });
  y += 6;

  P.PRECISE_WORDS.forEach((group) => {
    y = P.addWordChipsPdf(doc, y, group.flat, group.better,
      { color: C.PRIMARY, perRow: 6, chipH: 19, fontSize: 8.8 });
  });

  addPdfFooter(doc, "Persuasive Writing | Year 5/6 Literacy | Diamond Creek East PS");
  await writePdf(doc, path.join(OUT_DIR, BANK_RES.fileName));
  console.log("Wrote " + BANK_RES.name);
}

// ─── Resource: Sentence Upgrade ─────────────────────────────────────────────

async function buildUpgrade() {
  const doc = createPdf({ title: UPGRADE_RES.name });
  let y = addPdfHeader(doc, "Sentence Upgrade", {
    subtitle: "Swap the lazy words. Do not rewrite the sentence.",
    color: C.SECONDARY,
    lessonInfo: "Week 6 Session 2 | Year 5/6 Literacy",
  });

  y = addSectionHeading(doc, "Watch one first", y, { color: C.SECONDARY });
  y += 4;

  // The worked example: drawn as a before/after block so the process is
  // visible on paper exactly as it was modelled on the slide.
  doc.save();
  doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 34, 3)
    .lineWidth(0.8).strokeColor(hex(C.MUTED)).stroke();
  doc.fontSize(9).font("Sans-Bold").fillColor(hex(C.MUTED));
  doc.text("BEFORE", PAGE.MARGIN + 8, y + 6);
  doc.fontSize(11).font("Sans").fillColor("#000000");
  doc.text(FLAT, PAGE.MARGIN + 8, y + 18, { width: PAGE.CONTENT_W - 16 });
  doc.restore();
  y += 40;

  doc.save();
  doc.roundedRect(PAGE.MARGIN, y, PAGE.CONTENT_W, 34, 3).fill(hex(C.SUCCESS));
  doc.fontSize(9).font("Sans-Bold").fillColor("#FFFFFF");
  doc.text("AFTER", PAGE.MARGIN + 8, y + 6);
  doc.fontSize(11).font("Sans-Bold").fillColor("#FFFFFF");
  doc.text(UPGRADED, PAGE.MARGIN + 8, y + 18, { width: PAGE.CONTENT_W - 16 });
  doc.restore();
  y += 40;

  y = addBodyText(doc,
    "Two words changed. Good became outstanding (precise). Bad became appalling (emotive). "
    + "Everything else stayed the same.",
    y, { italic: true });
  y += 6;

  y = addSectionHeading(doc, "Now you. Underline every word you swap.", y, { color: C.SECONDARY });
  y += 4;

  const items = [
    "School lunches are not very nice.",
    "Too much screen time is bad for kids.",
    "The library is a good place to be.",
    "The playground equipment is old and not very safe.",
    "Dropping litter is a bad thing to do.",
    "Reading every night is a good habit.",
  ];

  items.forEach((sentence, i) => {
    doc.save();
    doc.circle(PAGE.MARGIN + 9, y + 8, 9).fill(hex(C.SECONDARY));
    doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(String(i + 1), PAGE.MARGIN + 2, y + 3, { width: 15, align: "center" });
    doc.fontSize(11).font("Sans").fillColor("#000000");
    doc.text(sentence, PAGE.MARGIN + 26, y + 2, { width: PAGE.CONTENT_W - 34 });
    doc.restore();
    y += 24;
    y = addLinedArea(doc, y, 2, { lineSpacing: 26 });
    y += 8;
  });

  y = addTipBox(doc,
    "Challenge: write a seventh sentence of your own about something you would change at school, "
    + "then upgrade it.",
    y, { color: C.ACCENT });

  addPdfFooter(doc, "Persuasive Writing | Week 6 Session 2 | Year 5/6 Literacy");
  await writePdf(doc, path.join(OUT_DIR, UPGRADE_RES.fileName));
  console.log("Wrote " + UPGRADE_RES.name);
}

(async () => {
  await build();
  await buildBank();
  await buildUpgrade();
})().catch((err) => {
  console.error(err);
  process.exit(1);
});
