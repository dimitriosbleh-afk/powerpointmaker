"use strict";

// Tom Unit -- Lesson 22: Chapters 44-47 -- Food Shortage & Second Fleet + KPAS Note-Taking
// Week 5, Lesson 22, Grade 5/6 Literacy
// Reading: Ch 44 (rations cut), Ch 45 (kangaroos at the cove), Ch 46 (isolation), Ch 47 (an emu, the women's ship)
// Sentence-level: KPAS note-taking -- key words, phrases, abbreviations, symbols

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const { createTheme, weekToVariant } = require("../themes/factory");
// Unit cohesion: all Tom lessons use the same variant.
const T = createTheme("literacy", "grade56", weekToVariant(2));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  withReveal,
  titleSlide, liSlide, contentSlide,
  cfuSlide, closingSlide,
  vocabSlide, quoteSlide, modellingSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 22;
const FOOTER = "Chapters 44-47 | Lesson 22 | Week 5 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson22_Food_Shortage_KPAS";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const KPAS_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "KPAS Note-Taking Practice",
  "Student worksheet: convert sentences from Chapters 44-47 into key words, phrases, abbreviations and symbols."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "KPAS Note-Taking Practice Answer Key",
  "Teacher reference: model KPAS conversions for the practice worksheet."
);
const RESOURCE_ITEMS = [KPAS_RESOURCE, ANSWER_KEY_RESOURCE];
const KPAS_PDF_PATH = path.join(OUT_DIR, KPAS_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Lesson 22. Four chapters today: 44, 45, 46 and 47
- The colony struggles with food. Then the second fleet arrives -- but it is not what the colony hoped for
- After reading we practise note-taking using key words, phrases, abbreviations and symbols (KPAS)

DO:
- Display title slide as students settle
- Have copies of the novel ready
- Have the KPAS practice worksheet ready (do not distribute yet)

TEACHER NOTES:
This is a content-dense lesson. Four chapters of reading. You may need to read briskly or split the chapters across two sessions. The sentence-level skill (KPAS) supports note-taking for the rest of the unit -- it will be reused when students plan future writing.

WATCH FOR:
- Students reacting to the Sergeant's comment in Chapter 47 about the women's ship -- pre-frame: "The Sergeant says something dismissive about women in the colony. The author is showing us his frustration, not approving of his view"

[Literacy: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${KPAS_RESOURCE.name} is your practice sheet
- The ${ANSWER_KEY_RESOURCE.name} is for teacher use

DO:
- Print the practice sheet (one per student)
- Print one answer key for teacher use
- Distribute the practice after the I Do / We Do

TEACHER NOTES:
The KPAS skill is a tool students will use again -- when they take notes from non-fiction articles for the conclusion of their information report and beyond. Frame it as a real skill, not just a worksheet.

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands today: reading the chapters and analysing the author's choices, then practising note-taking with KPAS
- Read the success criteria

DO:
- Choral read the LI, then the SCs
- Brief: "Note-taking means writing the IMPORTANT bits in a shorter way -- not the whole sentence"

TEACHER NOTES:
SC1 targets character and authorial-choice analysis. SC2 targets the procedural understanding of KPAS. SC3 targets the application using sentences from Chapters 44-47.

WATCH FOR:
- Students who think note-taking = writing the whole sentence faster -- redirect: "Notes use key words, not full sentences"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_READING = `SAY:
- Four chapters today. I will read aloud
- Four pause points
- The chapters track the colony as food gets scarcer, then a second fleet arrives -- but it is mostly women and the supply ship did not make it

DO:
- 30 seconds for students to find Chapter 44
- Read aloud at a steady pace
- Take pause point 1 at p.237, pause point 2 at p.241, pause point 3 at p.243, pause point 4 at p.250
- Read all four chapters

TEACHER NOTES:
The reading takes 18-22 minutes. These chapters carry quiet weight -- the isolation deepens, then the second fleet's arrival is a complicated relief. The Sergeant's response in Ch 47 needs framing as the author showing his frustration, not endorsing his view.

WATCH FOR:
- Students reacting to the Sergeant's words about the women -- pause and frame
- Students who notice the contrast between the boys' kangaroos / emu and the colony's hunger -- celebrate

[Literacy: Read Aloud | VTLM 2.0: Build Knowledge / Reading]`;

const NOTES_PAUSE1 = `SAY:
- Stop. Tom is thinking about how many days remain until he could be free and board a ship for home
- Turn to your partner: how is the author making you feel about Tom right now? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: sympathy, loneliness, longing; Tom is still counting the days even now; he is homesick despite settling in; the author wants us to remember Tom is far from home

TEACHER NOTES:
This is a quiet but important emotional beat. Tom has been with the Sergeant and Rob for some time, but he still thinks about leaving. The author keeps Tom's homesickness alive even as the colony becomes more familiar.

WATCH FOR:
- Students who only mention the action -- prompt: "How does the AUTHOR want us to feel?"
- Students who connect this to earlier homesickness -- excellent

[Literacy: Pause Point | VTLM 2.0: Comprehension / Author Craft]`;

const NOTES_PAUSE2 = `SAY:
- Stop. The narrator tells us how Tom thinks: "yes, that is how I want to live"
- Turn to your partner: what do you think the author wants us to know about Tom? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: Tom is changing; he is starting to imagine staying; what he wants is shifting from "go home" to "live like this"; the author is showing growth

TEACHER NOTES:
This pause point captures a quiet but big moment. Earlier Tom was counting days to leave. Now he is imagining staying. The author is showing the slow shift in Tom's thinking. Strong students will name the change.

WATCH FOR:
- Students who repeat Tom's thought without naming the change -- prompt: "What did Tom WANT before? Is it different now?"
- Students who name the shift -- celebrate

[Literacy: Pause Point | VTLM 2.0: Comprehension / Character Development]`;

const NOTES_PAUSE3 = `SAY:
- Stop. The narrator tells us: "Now they too watched the harbour, waiting for the sails that never came"
- Turn to your partner: how do things look for Tom, Rob and the Sergeant now? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: bleak, lonely, isolated; even the boys are joining the watching; the disappointment is now shared by everyone; the colony's situation is desperate; the boys have lost some of their optimism

TEACHER NOTES:
"Sails that never came" is a small phrase doing big work. The author connects the boys to the larger colony grief -- they are no longer separate from it. Strong students will name that connection.

WATCH FOR:
- Students who describe the situation literally -- prompt: "Why does the author mention sails that NEVER come?"
- Students who name the emotional connection -- celebrate

[Literacy: Pause Point | VTLM 2.0: Comprehension / Big Idea]`;

const NOTES_PAUSE4 = `SAY:
- Stop. The Sergeant has just learned the second fleet brought women but the supply ship did not make it. He looks at Rob sharply but says nothing
- Turn to your partner: what is going on here? What do you think the author wants us to know? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: the Sergeant is holding back; he is frustrated; his sharp look hints at something he will not say aloud in front of the boys; he is worried about feeding more mouths

TEACHER NOTES:
The Sergeant says nothing but the author tells us he LOOKED sharply. This is showing not telling -- there is something the Sergeant is not saying. Strong students will notice the gap between his look and his silence.

WATCH FOR:
- Students who only describe his look -- prompt: "Why might the Sergeant NOT say what he was thinking?"
- Students who notice the silence is the message -- excellent literary inference

[Literacy: Pause Point | VTLM 2.0: Comprehension / Inference]`;

const NOTES_VOCAB = `SAY:
- One vocabulary word today: meagre
- Meagre means very small or not enough -- often used about food, money or supplies
- A meagre meal is a small meal. A meagre harvest is a small harvest

DO:
- Display the word and the image
- Read the meaning and example aloud
- Quick oral check: "What is the opposite of meagre? Plentiful, generous, large"
- Brief: meagre appears in Chapter 47 when the boys eat the emu

TEACHER NOTES:
Meagre is the explicit vocab word from the supplied unit. Students may also have heard 'meager' in American spellings. Use Australian spelling 'meagre' on the board.

WATCH FOR:
- Students who use the word in their own sentence -- celebrate
- Students unsure of pronunciation -- model: "MEE-ger"

[Literacy: Vocabulary | VTLM 2.0: Build Knowledge / Vocabulary]`;

const NOTES_REVISE = `SAY:
- Quick revision of KPAS -- Key words, Phrases, Abbreviations, Symbols
- Note-taking means writing the IMPORTANT bits in shorter form
- Abbreviations: b/c (because), w/ (with), w/o (without)
- Symbols: = (means / definition), + (and), -> (cause / result), up arrow (more / increase), down arrow (less / decrease), / (full stop or new idea)
- Omit small words like "the", "and", "a" when taking notes

DO:
- Display the KPAS card
- Choral read each abbreviation and symbol
- Quick oral check: "What does b/c mean? What does the down arrow mean?"

TEACHER NOTES:
Keep the recall tight. Most students will have met some of these in earlier years. The key idea is efficiency -- notes are a tool to capture meaning quickly, not to reproduce the sentence.

WATCH FOR:
- Students who think notes need to be neat full sentences -- redirect: "Notes are for YOU. Short is fine"
- Students who already use their own abbreviations -- celebrate and add theirs to the bank

[Literacy: Revise | VTLM 2.0: Retention and Recall]`;

const NOTES_IDO = `SAY:
- Watch me convert one sentence from Chapter 44 into KPAS
- Sentence: "The rations of food in the colony were cut in half because the supply ship was late."
- Step 1: find the key words. WHO/WHAT? rations. WHAT DOING? cut in half. WHY? supply ship late
- Step 2: drop the small words -- the, in, were, was
- Step 3: use abbreviations and symbols
- My notes: "rations cut in half b/c supply ship late"
- I can shorten more: "rations halved b/c supply ship late"

DO:
- Display the original sentence and the KPAS notes side by side
- Talk through each step
- Highlight which words were dropped, which were replaced

TEACHER NOTES:
Show the thinking, not just the answer. Some students will want to keep the whole sentence -- model the deliberate move of dropping the smaller words. Note that "rations halved" replaces "rations cut in half" with a shorter word -- show this as one of several valid choices.

WATCH FOR:
- Students who try to memorise YOUR notes verbatim -- redirect: "Your notes might use different words. That is fine"
- Students who can already do this faster -- ask them to share their notes and discuss the choices

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Together. Mini-whiteboards out
- Sentence from Chapter 46: "The morale in the colony was low because the people felt isolated from the rest of the world."
- Take 60 seconds. Convert this sentence into KPAS notes
- Boards up when called

DO:
- 60 seconds silent work on whiteboards
- Boards up
- Pick 2-3 strong examples to read aloud
- Reveal sample notes: "morale low b/c colony felt isolated"
- Quick discussion: "What did YOU drop? What did you keep?"

TEACHER NOTES:
This is guided practice. Most students should produce workable notes. Common alternatives: "morale down b/c isolated", "colony morale low -- isolation", "low morale b/c far from world". All are acceptable if they capture the key idea concisely.

WATCH FOR:
- Students who copy the whole sentence -- redirect: "Drop the small words"
- Students who leave out the cause -- redirect: "The b/c bit is important"

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Two sets of notes for the same sentence. Which set is BETTER notes?
- Sentence: "Wheat was not growing well near the cove, though it was growing well at a garden further away."
- A: "Wheat was not growing well near the cove but it was growing well at a garden further away."
- B: "wheat not growing near cove / growing well further away"
- Hold up A or B. Three, two, one -- show

DO:
- Display both
- Show Me Boards
- Scan: most students should choose B
- Cold Call 1-2 students: "Why B?"

TEACHER NOTES:
A is barely shortened -- it dropped "though" and added "but" but kept the full sentence. B uses key words, drops small words, and uses the / symbol to signal a new idea. B is the better notes.

WATCH FOR:
- Students who pick A because it is more grammatical -- redirect: "Notes are for YOU. Are they shorter? Do they capture the meaning?"
- Students who pick B and explain -- they have the principle

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Better notes: B
- B uses key words only, drops small words, and uses the / symbol to signal a new idea
- A is too long -- it is still a full sentence, just slightly rearranged
- Aim for B-style notes in your own work

DO:
- Display the reveal banner
- Read B aloud
- Pivot if many picked A: "Compare the LENGTH. Which is shorter? Which captures the same meaning?"

CFU CHECKPOINT:
Technique: Show Me Boards (A or B)
Script:
- "Hold up A or B. Which is the better set of notes?"
- Scan for: most students choose B (>=80%)
PROCEED (>=80%): Most chose B. Release to the worksheet.
PIVOT (<80%): Most likely issue -- students think a note must be a sentence. Reteach: "Notes use key words and symbols. Drop the small words. Use / for new ideas."

TEACHER NOTES:
After reveal, release students to the worksheet. The principle (notes are short, not sentences) is the takeaway.

WATCH FOR:
- Students who self-correct toward B-style notes -- celebrate

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Your turn. The worksheet has six sentences from Chapters 44-47
- For each one: convert it into KPAS notes
- Use the abbreviation and symbol bank at the top
- Drop the small words. Use key words only
- 12 minutes. I will circulate

DO:
- Distribute the worksheet
- Circulate -- prioritise students who looked unsure during the CFU
- Quick conferences: "Read me your notes. Do they capture the key idea?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the enabling version of the worksheet which provides each sentence with the key words already underlined -- students just convert the underlined words into KPAS form
- Extra Notes: This focuses cognitive load on the conversion step rather than on identifying key words

EXTENDING PROMPT:
- Task: After completing the worksheet, take ONE paragraph from any chapter we have read and convert the whole paragraph into KPAS notes. Aim for notes that are HALF the original length
- Extra Notes: Peer-share -- can your partner reconstruct the original meaning from your notes?

TEACHER NOTES:
The 12-minute block is independent practice. The worksheet design (six sentences from the chapters) makes the link between reading and note-taking explicit. Active circulation catches the most common error -- not enough shortening.

WATCH FOR:
- Students who keep the small words -- prompt: "Can you drop 'the' and 'a'?"
- Students who use no symbols at all -- prompt: "Could you use -> here? Or b/c?"
- Students whose notes lose the meaning -- prompt: "Read your notes back. Do they still make sense?"

[Literacy: You Do | VTLM 2.0: Independent Application]`;

const NOTES_CLOSING = `SAY:
- Quick self-check. Show me on your fingers
- SC1: I described what we learned about Tom and the colony from Chapters 44-47 -- 1 to 5
- SC2: I know what KPAS stands for and I can use abbreviations and symbols -- 1 to 5
- SC3: I converted sentences from the chapters into KPAS notes -- 1 to 5
- Then turn to your partner: tell them ONE abbreviation or symbol we used today

DO:
- Run each SC
- Listen in on partner shares -- this tells you who has the terminology
- Wrap up: "Notes are a tool you will use again when we read non-fiction for the next part of our information report"

TEACHER NOTES:
Collect worksheets so you can spot common patterns. If many students kept full sentences, plan a 5-minute reteach next session focusing on dropping small words.

WATCH FOR:
- Students who name multiple abbreviations -- evidence of retention
- Students unsure -- check in privately

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Tom Appleby Ch 44-47 + KPAS Note-Taking -- Lesson 22";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Chapters 44-47",
    "Food Shortage, Second Fleet + KPAS Note-Taking",
    "Lesson 22  |  Week 5  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources (immediately after title)
  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // SLIDE 3 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to analyse the author's choices in Chapters 44-47 and to take notes by identifying key words and phrases and using abbreviations and symbols",
    ],
    [
      "I can describe what we learned about Tom and the colony from Chapters 44 to 47",
      "I can use key words, phrases, abbreviations and symbols to take shorter notes",
      "I can convert sentences from the chapters into KPAS notes",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 4 -- Reading anchor
  contentSlide(
    pres,
    "Read",
    C.PRIMARY,
    "Chapters 44, 45, 46 & 47",
    [
      "Ch 44: Rations cut in half. Convicts and marines still refuse to garden",
      "Ch 45: Convicts' work hours shortened to encourage gardening. Tom sees kangaroos at the cove",
      "Ch 46: Colony morale is low. Isolation deepens. Even the boys feel it",
      "Ch 47: The boys eat an emu. A second fleet arrives -- but the supply ship did not make it",
      "Four pause points along the way",
    ],
    NOTES_READING,
    FOOTER
  );

  // SLIDE 5 -- Pause point 1 (p.237)
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 44  |  Tom counts the days",
    "...to count the days before you can be free and board a ship for home?",
    "p.237",
    "How is the author making you feel about Tom right now?",
    NOTES_PAUSE1,
    FOOTER
  );

  // SLIDE 6 -- Pause point 2 (p.241)
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 45  |  Tom thinks about the cove",
    "...yes, that is how I want to live.",
    "p.241",
    "What do you think the author wants us to know about Tom?",
    NOTES_PAUSE2,
    FOOTER
  );

  // SLIDE 7 -- Pause point 3 (p.243)
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 46  |  Watching the harbour",
    "Now they too watched the harbour, waiting for the sails that never came.",
    "p.243",
    "How do things look for Tom, Rob and the Sergeant now?",
    NOTES_PAUSE3,
    FOOTER
  );

  // SLIDE 8 -- Pause point 4 (p.250)
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 47  |  The Sergeant's sharp look",
    "The sergeant looked at him sharply for a moment, but said nothing.",
    "p.250",
    "What's going on here? What do you think the author wants us to know?",
    NOTES_PAUSE4,
    FOOTER
  );

  // SLIDE 9 -- Vocabulary: meagre
  vocabSlide(
    pres,
    "meagre",
    "adjective",
    "Very small or not enough -- often used about food, money or supplies.",
    "Christmas dinner at the cove was a meagre meal compared to the year before.",
    NOTES_VOCAB,
    FOOTER
  );

  // SLIDE 10 -- Revise: KPAS
  contentSlide(
    pres,
    "Revise",
    C.SECONDARY,
    "KPAS -- Key Words, Phrases, Abbreviations, Symbols",
    [
      "Note-taking = writing the IMPORTANT bits in shorter form",
      "Abbreviations:  b/c (because)  |  w/ (with)  |  w/o (without)",
      "Symbols:  =  (means)  |  +  (and)  |  ->  (cause / result)",
      "More symbols:  up arrow (more / increase)  |  down arrow (less / decrease)  |  /  (new idea)",
      "Drop small words like \"the\", \"and\", \"a\" when taking notes",
    ],
    NOTES_REVISE,
    FOOTER
  );

  // SLIDE 11 -- I Do: model KPAS conversion
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Convert a Sentence into KPAS Notes",
    "Sentence from Ch 44:\n\n\"The rations of food in the\ncolony were cut in half\nbecause the supply ship\nwas late.\"\n\nStep 1: find the key words\n  -> rations\n  -> cut in half\n  -> supply ship late\n\nStep 2: drop small words\nthe, in, of, were, was, the",
    "Step 3: use abbreviations\n  -> b/c  (because)\n\nMy notes:\n\n\"rations cut in half b/c\nsupply ship late\"\n\nEven shorter:\n\n\"rations halved b/c\nsupply ship late\"\n\nNotice: same meaning, fewer words",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 12 -- We Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Together -- Whiteboards", { color: C.SECONDARY });

    addCard(s, 0.5, CONTENT_TOP, 9, 1.45, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("Convert this sentence into KPAS notes (Ch 46):", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.32,
      fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText("\"The morale in the colony was low because the people felt isolated from the rest of the world.\"", {
      x: 0.75, y: CONTENT_TOP + 0.46, w: 8.4, h: 0.90,
      fontSize: 17, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
      fit: "shrink", shrinkText: true, valign: "middle",
    });

    const tipY = CONTENT_TOP + 1.65;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Your job:", {
      x: 0.75, y: tipY + 0.10, w: 4, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Find the KEY words\n- Drop the small words (the, in, was)\n- Use b/c for 'because'\n- 60 seconds. Boards up when called", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: 0.95,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
  }

  // SLIDE 13 + 14 -- CFU: which notes are better? (reveal pair)
  function buildCfuBase() {
    const slide = pres.addSlide();
    addTopBar(slide, C.ALERT);
    addBadge(slide, "CFU", { color: C.ALERT });
    addTitle(slide, "Which Set Is Better Notes?", { color: C.ALERT });

    const stampW = 1.3;
    slide.addShape("roundRect", {
      x: 9.5 - stampW, y: 0.20, w: stampW, h: 0.32, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
    });
    slide.addText("CHECK", {
      x: 9.5 - stampW, y: 0.20, w: stampW, h: 0.32,
      fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    slide.addShape("roundRect", {
      x: 0.5, y: CONTENT_TOP, w: 4.5, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ALERT },
    });
    slide.addText("Show Me Boards: A or B", {
      x: 0.5, y: CONTENT_TOP, w: 4.5, h: 0.40,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Prompt row
    slide.addText("Sentence:  \"Wheat was not growing well near the cove, though it was growing well at a garden further away.\"", {
      x: 0.5, y: CONTENT_TOP + 0.55, w: 9, h: 0.55,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0, valign: "middle",
      fit: "shrink", shrinkText: true,
    });

    const cardY = CONTENT_TOP + 1.20;
    const cardH = 1.05;
    addCard(slide, 0.5, cardY, 9, cardH, { strip: C.MUTED, fill: C.WHITE });
    slide.addText("A", {
      x: 0.7, y: cardY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    slide.addText("\"Wheat was not growing well near the cove but it was growing well at a garden further away.\"", {
      x: 1.2, y: cardY + 0.18, w: 8.0, h: cardH - 0.30,
      fontSize: 14, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const cardBY = cardY + cardH + 0.18;
    const cardBH = SAFE_BOTTOM - cardBY - 0.55;
    addCard(slide, 0.5, cardBY, 9, cardBH, { strip: C.SECONDARY, fill: C.BG_CARD });
    slide.addText("B", {
      x: 0.7, y: cardBY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    slide.addText("\"wheat not growing near cove / growing well further away\"", {
      x: 1.2, y: cardBY + 0.14, w: 8.0, h: cardBH - 0.24,
      fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(slide, FOOTER);
    slide.addNotes(NOTES_CFU_BUILD);
    return slide;
  }

  withReveal(
    buildCfuBase,
    (slide) => {
      const revealY = 4.68;
      addCard(slide, 0.5, revealY, 9, 0.42, { fill: C.SUCCESS });
      slide.addText("Better notes: B  --  key words only, uses / for new idea", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 15 -- You Do: worksheet
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Convert to KPAS Notes");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.85, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Six sentences from Chapters 44-47", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:    Read each sentence aloud quietly\nNext:     Underline the KEY words (who, what doing, where, when, why)\nThen:     Drop the small words (the, and, a, of)\nFinally:  Use abbreviations and symbols where helpful", {
      x: 0.75, y: CONTENT_TOP + 0.46, w: 8.4, h: 1.30,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 2.00;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Time: 12 minutes", {
      x: 0.75, y: tipY + 0.10, w: 4, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- The abbreviation and symbol bank is at the top of your worksheet\n- Read your notes aloud -- do they still capture the meaning?\n- Notes do not need to be full sentences", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: 0.85,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 16 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Tell your partner one abbreviation or symbol we used today.",
      scItems: [
        "I can describe what we learned about Tom and the colony from Chapters 44-47",
        "I can use key words, phrases, abbreviations and symbols to take shorter notes",
        "I can convert sentences from the chapters into KPAS notes",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (got it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: KPAS Practice -------------------------------------------------
  const ws = createPdf({ title: KPAS_RESOURCE.name });
  let wsY = addPdfHeader(ws, "KPAS Note-Taking Practice", {
    color: C.PRIMARY,
    subtitle: "Convert each sentence from Chapters 44-47 into notes",
    lessonInfo: "Lesson 22 | Week 5 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "For each sentence: find the KEY words, drop the small words (the, and, a, of), and use abbreviations and symbols where helpful. Notes do not need to be full sentences.", wsY, { color: C.PRIMARY });

  wsY = addSectionHeading(ws, "Abbreviation & Symbol Bank", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Abbreviations:   b/c = because   |   w/ = with   |   w/o = without", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "Symbols:   =  means / definition   |   +  and   |   ->  cause / result", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "More:   up arrow  more / increase   |   down arrow  less / decrease   |   /  full stop / new idea", wsY, { fontSize: 11 });
  wsY += 8;

  wsY = addSectionHeading(ws, "1.  Sentence from Chapter 44", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "\"The rations of food were cut in half again because the supply ship had not arrived.\"", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "My KPAS notes:", wsY, { fontSize: 10 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;

  wsY = addSectionHeading(ws, "2.  Sentence from Chapter 45", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "\"The convicts' work hours were shortened so they could spend more time caring for their gardens.\"", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "My KPAS notes:", wsY, { fontSize: 10 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;

  wsY = addSectionHeading(ws, "3.  Sentence from Chapter 46", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "\"The morale in the colony was low because everyone felt isolated from the rest of the world.\"", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "My KPAS notes:", wsY, { fontSize: 10 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;

  wsY = addSectionHeading(ws, "4.  Sentence from Chapter 47", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "\"The boys noticed an emu wandering in their garden and Rob shot it so they could eat the meat.\"", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "My KPAS notes:", wsY, { fontSize: 10 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;

  wsY = addSectionHeading(ws, "5.  Sentence from Chapter 47", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "\"The ship that arrived was full of 200 convict women but the supply ship that was meant to come with it did not make it.\"", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "My KPAS notes:", wsY, { fontSize: 10 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;

  wsY = addSectionHeading(ws, "6.  Sentence from Chapter 47", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "\"The Sergeant was upset because there were now more people in the colony to feed but no extra food.\"", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "My KPAS notes:", wsY, { fontSize: 10 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 6;

  wsY = addSectionHeading(ws, "Challenge", wsY, { color: C.ALERT });
  wsY = addBodyText(ws, "Choose any short paragraph from Chapters 44-47. Convert the WHOLE paragraph into KPAS notes. Try to make your notes about HALF the original length.", wsY, { fontSize: 11, italic: true });
  wsY = addLinedArea(ws, wsY, 4, { lineSpacing: 22 });

  addPdfFooter(ws, "Lesson 22 | KPAS Note-Taking Practice");

  // ---- PDF: Answer Key ---------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "KPAS Note-Taking Practice -- Answer Key", {
    color: C.SECONDARY,
    subtitle: "Teacher reference -- many valid versions for each",
    lessonInfo: "Lesson 22 | Week 5 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Many versions of KPAS notes are valid. The principle is: shorter than the original; captures the key idea; uses key words and (where helpful) abbreviations and symbols. Accept any version that meets these tests.", akY, { color: C.SECONDARY });

  akY = addSectionHeading(ak, "1. Rations cut in half (Ch 44)", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sample notes:  \"rations cut in half again b/c supply ship not arrived\"", akY, { fontSize: 11 });
  akY = addBodyText(ak, "Alt:  \"rations halved / supply ship late\"  or  \"food rations down b/c no supply ship\"", akY, { fontSize: 10, italic: true });
  akY += 4;

  akY = addSectionHeading(ak, "2. Work hours shortened (Ch 45)", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sample notes:  \"convicts' work hours shortened -> more time for gardens\"", akY, { fontSize: 11 });
  akY = addBodyText(ak, "Alt:  \"work hours down / time to garden up\"  or  \"shorter work -> more gardening\"", akY, { fontSize: 10, italic: true });
  akY += 4;

  akY = addSectionHeading(ak, "3. Morale low (Ch 46)", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sample notes:  \"colony morale low b/c isolated\"", akY, { fontSize: 11 });
  akY = addBodyText(ak, "Alt:  \"morale down / felt isolated\"  or  \"morale low b/c far from rest of world\"", akY, { fontSize: 10, italic: true });
  akY += 4;

  akY = addSectionHeading(ak, "4. Boys and the emu (Ch 47)", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sample notes:  \"boys saw emu in garden / Rob shot it / they ate meat\"", akY, { fontSize: 11 });
  akY = addBodyText(ak, "Alt:  \"emu in garden -> Rob shot -> meat for food\"", akY, { fontSize: 10, italic: true });
  akY += 4;

  akY = addSectionHeading(ak, "5. Second fleet -- women's ship (Ch 47)", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sample notes:  \"ship arrived = 200 convict women / supply ship did not make it\"", akY, { fontSize: 11 });
  akY = addBodyText(ak, "Alt:  \"ship = 200 women / no supply ship\"  or  \"200 women arrived w/o supply ship\"", akY, { fontSize: 10, italic: true });
  akY += 4;

  akY = addSectionHeading(ak, "6. Sergeant upset (Ch 47)", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sample notes:  \"Sergeant upset b/c more mouths but no extra food\"", akY, { fontSize: 11 });
  akY = addBodyText(ak, "Alt:  \"Sergeant upset / more people + no extra food\"", akY, { fontSize: 10, italic: true });
  akY += 6;

  akY = addSectionHeading(ak, "Common Errors to Watch For", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- Notes that are still full sentences. Reteach: drop small words.", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Notes that drop the cause/result (missing 'b/c' or '->'). Reteach: capture WHY as well as WHAT.", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Notes that are so short the meaning is lost. Reteach: 'read your notes back -- can you reconstruct the meaning?'", akY, { fontSize: 10 });

  addPdfFooter(ak, "Lesson 22 | KPAS Note-Taking Practice Answer Key -- TEACHER USE");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Lesson22.pptx` }),
    writePdf(ws, KPAS_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Lesson22.pptx`);
  console.log("Done: " + KPAS_RESOURCE.name);
  console.log("Done: " + ANSWER_KEY_RESOURCE.name);
}

build().catch(console.error);
