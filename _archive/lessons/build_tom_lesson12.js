"use strict";

// Tom Unit -- Lesson 12: Chapters 25-26 + Single Paragraph Outline (SPO) to summarise text
// Week 3, Lesson 12, Year 5/6 Literacy
// Reading: Chapters 25-26 (Cape storm, Sam's death, sighting land, sailing to Port Jackson)
// Sensitivity: Chapter 25 contains content about death and some adult references.
// Writing: Use a Single Paragraph Outline to summarise main ideas from the chapters.

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const { createTheme, weekToVariant } = require("../themes/factory");
const T = createTheme("literacy", "grade56", weekToVariant(2));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  addInstructionCard,
  withReveal,
  titleSlide, liSlide, contentSlide,
  cfuSlide, closingSlide,
  quoteSlide, modellingSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 12;
const FOOTER = "Chapters 25-26 | Lesson 12 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson12_Cape_to_Port_Jackson";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const SPO_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "SPO Summary Template",
  "Student template: summarise an aspect of Chapters 25-26 using a topic sentence, supporting details and a concluding sentence."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "SPO Summary Model",
  "Teacher reference: model SPO summarising Chapters 25-26 with alternative topic options."
);
const RESOURCE_ITEMS = [SPO_RESOURCE, ANSWER_KEY_RESOURCE];
const SPO_PDF_PATH = path.join(OUT_DIR, SPO_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Two chapters today: 25 and 26
- The voyage is almost over, but the hardest part is still ahead
- After reading we use a Single Paragraph Outline -- an SPO -- to summarise an aspect of what we read

DO:
- Display title slide as students settle
- Have novels and exercise books on desks

SENSITIVITY ADVISORY:
- What it is: Chapter 25 includes Sam's death below deck and some adult references in the description of conditions
- Framing language: "We are going to read about the death of a character we have come to know. Authors include hard moments to show us what people endure. We will read carefully and respectfully"
- Watch for: students who become very quiet, withdraw or are visibly upset
- Protocol: pause if needed; if a student is distressed, allow them to step out with a buddy. Plan to give Sam's death its own quiet moment after the read

TEACHER NOTES:
Lesson 12 of the Tom unit. Chapter 25 is emotionally heavy -- Sam's death follows a storm sequence. Chapter 26 brings the ship to land but they sail away again. Pace the read; pause after Sam's death for a brief moment before continuing.

WATCH FOR:
- Students sensitive to the loss of Sam -- check privately
- Students who notice the bird Tom sees as Sam dies -- excellent symbolic reading

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands. First, what the author shows us about the journey and the characters. Second, summarising what we read using an SPO
- Read each success criterion together

DO:
- Choral read the LI, then each SC
- Brief reminder: "An SPO is a Single Paragraph Outline -- a topic sentence, a few supporting details, and a concluding sentence"

TEACHER NOTES:
SC1 targets reading analysis. SC2 targets vocabulary application. SC3 targets the SPO summarisation skill.

WATCH FOR:
- Students unsure about SPO -- they will see it modelled in the I Do
- Students confident about SPO from earlier lessons -- they can support peers

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_READING = `SAY:
- Reading mode: teacher read aloud
- Two chapters. Listen for how the journey changes -- the storm, then the calm, then the new land
- I will pause twice for us to think together
- After the read, we will summarise what we have learned

DO:
- Read Chapter 25, then Chapter 26
- Plan two pause points: p.126 (Sam's voice fades) and p.132 (the unloading parties)
- After the funeral scene, take a quiet beat before continuing

TEACHER NOTES:
The reading is the largest cognitive load in this lesson. Be alert to student emotional responses around Sam's death. The SPO work afterwards uses material from these chapters.

WATCH FOR:
- Students who go quiet -- check privately after the read
- Students who notice the bird Tom sees as Sam dies -- the author is using imagery; acknowledge

[General: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause here. Sam has been talking to Tom but his voice is fading
- "His voice died away again, into a snore"
- What do we learn from this conversation? [Sam is very ill. He keeps slipping in and out of sleep. The author is showing us Sam fading -- not just sleeping]
- Look at the word "died" used about a voice. What is the author preparing us for? [Foreshadowing -- Sam will not survive the journey]

DO:
- Display the quote
- Think-Pair-Share: 30 seconds think, 30 seconds pair
- Push for the foreshadowing reading -- "died away" before any character actually dies

TEACHER NOTES:
This pause point lets students notice the author's foreshadowing. The verb choice ("died") and the fading metaphor prepare the reader for what comes next.

WATCH FOR:
- Students who notice the word "died" used here -- excellent close reading
- Students worried about Sam -- their feelings are appropriate; the author is doing this on purpose

[General: Pause Point 1 | VTLM 2.0: Higher-Order Questioning]`;

const NOTES_PAUSE2 = `SAY:
- Pause here. The ship has anchored at Port Jackson. The men prepare to disembark
- "...to make the first unloading parties on shore"
- What is going on here? [The First Fleet has reached Australia. The convicts and marines are about to step off the ship for the first time in nearly 8 months. The story is about to enter a completely new world]
- What might happen next? [The colony will begin. Tom's life on land will start. The story shifts from sea to settlement]

DO:
- Display the quote
- Cold Call after thinking time
- Brief context: this is the moment the First Fleet arrived at Sydney Cove in January 1788

TEACHER NOTES:
This pause point connects the novel to history. Students should recognise that this moment is the start of European settlement of Australia -- a turning point in the story and in the history. Be ready to acknowledge that this is also the start of impact on Aboriginal and Torres Strait Islander peoples; this becomes explicit in Chapter 28.

WATCH FOR:
- Students who connect this to Australian history -- excellent
- Students who recognise the historical date -- January 1788

[General: Pause Point 2 | VTLM 2.0: Deep Comprehension]`;

const NOTES_VOCAB = `SAY:
- Three explicit words from today's chapters
- Sporadic -- happening now and then, not regularly. The convicts heard sporadic shouts from the deck during the storm
- Sufficient -- enough. There was not sufficient food for everyone
- Unoccupied -- not lived in or used. The land seemed unoccupied to the captain

DO:
- Read each word and meaning aloud
- Quick oral routine: "Use sufficient in a sentence about the voyage" -- 30 seconds, partner share
- Note pronunciation: spo-RAD-ic; suff-FISH-ent; un-OCC-upied

TEACHER NOTES:
"Unoccupied" carries a weight in this novel. The land was clearly occupied by Aboriginal and Torres Strait Islander peoples; the captain saw it as unoccupied because he did not see European-style settlements. Be ready to discuss this if students raise it -- they will see it more directly in Chapter 28.

WATCH FOR:
- Students who question "unoccupied" -- excellent critical reading; affirm and note the lesson 13 connection
- Students who use the words naturally in discussion -- celebrate

[General: Vocabulary | VTLM 2.0: Vocabulary Instruction]`;

const NOTES_IDO_SPO = `SAY:
- Now we summarise. We use a Single Paragraph Outline -- an SPO -- to capture the main idea of an aspect of the chapters
- Watch me. My topic: the storm at the Cape of Good Hope
- First, my topic sentence. It tells the reader who, what doing, when, where, why or how. "When the ship reached the Cape of Good Hope, it met massive storms that left the convicts trapped below deck for days."
- Now my supporting details. These come from the text:
  - Detail 1: The waves rose impossibly high above the ship
  - Detail 2: The convicts were locked below without food or water
  - Detail 3: The storms knocked the ship about for thirteen days
- Concluding sentence -- wraps up without repeating the topic sentence: "The Cape storm tested every person on board and pushed the convicts to their limits."
- Notice: this is a PLAN. It is short notes and one or two full sentences, not a polished paragraph

DO:
- Display the SPO template on screen
- Build it section by section
- Show students where each detail comes from in the text
- Think aloud: "I check -- do my details actually support my topic sentence? Detail 1 about the waves, yes. Detail 2 about being locked below, yes. Detail 3 about the thirteen days, yes"

TEACHER NOTES:
The I Do models the complete SPO process applied to summarising. The think-aloud about checking that details support the topic sentence is the key metacognitive move.

WATCH FOR:
- Students who want to summarise the whole chapter at once -- redirect: "Pick ONE aspect to summarise"
- Students who include details from the wrong chapter -- gentle redirect

[General: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Together. Topic: the arrival at Port Jackson
- Mini-whiteboards out
- Topic sentence first. Who, what doing, where, when?
- Sample: "On a January day in 1788, the ship anchored at Port Jackson and the First Fleet finally reached its destination." Let me write that. You write your version
- Now three details. What happened at Port Jackson? Use the chapter
- Concluding sentence. How do we wrap up?

DO:
- Build the SPO as a class on the board
- Each student writes their own version on a mini-whiteboard
- Sample details: ship dropped anchor; first unloading parties were sent ashore; tents were set up
- Sample CS: "The voyage was over, but a completely new chapter was about to begin"

TEACHER NOTES:
The We Do is collaborative. The teacher writes a sample and students write their own version. Compare a few student versions. Acceptable variations are wide.

WATCH FOR:
- Students who copy the teacher's sentence verbatim -- prompt: "Try writing it in your own words"
- Students who include details from outside the chapter -- redirect to the text

[General: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU = `SAY:
- Quick check before you write your own
- Look at this topic sentence: "The convicts had a hard time on the ship"
- Is this a strong topic sentence? Thumbs up = yes, thumbs down = no
- Three, two, one, show

DO:
- Use Thumbs Up/Down
- Scan for: thumbs down (it is too vague)
- Cold call: "What would make it stronger?"

CFU CHECKPOINT:
Technique: Thumbs Up / Down

Script:
- "Is this a strong topic sentence? Thumbs up = yes, thumbs down = no"
- Scan for: mostly thumbs down
- Follow up: "What is missing?" [WHO had a hard time, WHEN, WHERE, HOW, WHY -- this sentence has no specifics. A strong TS gives the reader specifics]

PROCEED (>=80%): Most show thumbs down with a reason. Release to write.
PIVOT (<80%): Most likely issue -- students think any sentence about the topic counts as a topic sentence. Reteach: "A topic sentence introduces the MAIN IDEA of a paragraph and gives the reader specifics. 'The convicts had a hard time on the ship' could be the topic sentence for ANY chapter. A good TS should make us go 'oh, this is about THIS specific moment'." Re-check by giving a stronger model: "When the ship hit the Cape of Good Hope, the convicts were locked below for thirteen days without food or water." Better?

TEACHER NOTES:
This CFU targets the difference between a vague topic sentence and a specific one. A strong TS gives the reader who/what/when/where/why/how.

WATCH FOR:
- Students who say it is a strong TS "because it is on topic" -- they need to see that "on topic" is not enough
- Students who can articulate what is missing -- ready to write

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Your turn. On your worksheet, write an SPO about ONE aspect of Chapters 25-26
- Choose your topic. The worksheet lists three options
- Topic sentence first -- who, what doing, when, where, why or how
- Then 3 details from the text. Use your KPAS skills if you want -- short notes, not full sentences
- Concluding sentence -- wrap up without repeating the TS
- 12 minutes. I will circulate

DO:
- Distribute the SPO Summary Template
- The worksheet offers 3 topic choices: the Cape storm, the arrival at Port Jackson, or the captain's decision to sail away from the first land
- Circulate -- check that the topic sentence is specific (not vague), details come from the text, and the CS does not repeat the TS

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide the topic sentence pre-written for the chosen topic. Students write 2-3 supporting details and the concluding sentence
- Extra Notes: These students can refer to the I Do model SPO as a parallel example

EXTENDING PROMPT:
- Task: After completing one SPO, write a SECOND SPO on a different aspect of the chapters. Then turn ONE of the SPOs into a full summary paragraph (3-5 sentences)

TEACHER NOTES:
The 12-minute write is the application of the SPO work. Check that students are choosing a specific aspect, not the whole chapter.

WATCH FOR:
- Students who try to summarise the whole chapter -- redirect: "Pick one moment"
- Students whose details do not match their TS -- the check question: "Does this detail support what your TS is about?"
- Students who finish quickly -- direct to the extending task

[General: You Do | VTLM 2.0: Independent Application]`;

const NOTES_CLOSING = `SAY:
- Quick self-check against the success criteria. Thumbs for each
- Then turn to a partner: tell them ONE thing the author showed us about the journey today

DO:
- Run thumbs check for each SC
- Listen in on partner shares
- Wrap up: "Tomorrow Tom steps off the ship for the first time in nearly eight months"

TEACHER NOTES:
The closing connects today's reading and writing. Collect SPOs for review before the next session.

WATCH FOR:
- Students "thumbs down" on SC3 -- they may need more SPO modelling next session
- Students sharing strong character or journey insights -- excellent

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${SPO_RESOURCE.name} is for writing your SPO summary
- The ${ANSWER_KEY_RESOURCE.name} is for teacher reference

DO:
- Print the SPO template (one per student)
- Print the model (teacher copy only)

TEACHER NOTES:
The student SPOs become formative assessment evidence for SC3. Collect after the lesson.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Chapters 25-26: Cape to Port Jackson -- Lesson 12";

  // Slide 1 -- Title
  titleSlide(
    pres,
    "Chapters 25-26",
    "Cape to Port Jackson",
    "Lesson 12  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // Slide 2 -- LI/SC
  liSlide(
    pres,
    [
      "We are learning to notice what the author shows us about the journey, and to summarise what we read using a Single Paragraph Outline",
    ],
    [
      "I can explain what the author wants us to know about the journey or a character",
      "I can use newly taught vocabulary in my own sentences",
      "I can summarise an aspect of the text using a topic sentence, supporting details and a concluding sentence",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // Slide 3 -- Reading Launch
  contentSlide(
    pres,
    "Teacher Read Aloud",
    C.PRIMARY,
    "Chapters 25 - 26",
    [
      "Reading Mode: Teacher Read Aloud",
      "Chapter 25: the Cape of Good Hope storm and a hard moment for Sam",
      "Chapter 26: the ship reaches land, sails again, and anchors at Port Jackson",
      "Focus: how does the author show us the journey changing?",
    ],
    NOTES_READING,
    FOOTER
  );

  // Slide 4 -- Pause Point 1
  quoteSlide(
    pres,
    "Pause Point 1",
    "Chapter 25 -- p. 126",
    "His voice died away again, into a snore.",
    "p. 126",
    "What do we learn from this moment? Look at the word \"died\" -- what is the author preparing us for?",
    NOTES_PAUSE1,
    FOOTER
  );

  // Slide 5 -- Pause Point 2
  quoteSlide(
    pres,
    "Pause Point 2",
    "Chapter 26 -- p. 132",
    "...to make the first unloading parties on shore.",
    "p. 132",
    "What is going on here? What might happen next?",
    NOTES_PAUSE2,
    FOOTER
  );

  // Slide 6 -- Vocabulary
  contentSlide(
    pres,
    "Vocabulary",
    C.SECONDARY,
    "Three Words from Today's Reading",
    [
      "sporadic -- happening now and then, not regularly (sporadic shouts during the storm)",
      "sufficient -- enough (there was not sufficient food for everyone)",
      "unoccupied -- not lived in or used (the captain thought the land was unoccupied)",
    ],
    NOTES_VOCAB,
    FOOTER
  );

  // Slide 7 -- I Do: Model SPO
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "SPO -- Summarising the Cape Storm",
    "My Topic:\nThe storm at the Cape of\nGood Hope\n\nTopic sentence (TS):\n\"When the ship reached\nthe Cape of Good Hope,\nit met massive storms\nthat left the convicts\ntrapped below deck for\ndays.\"",
    "Supporting details:\n1. Waves rose impossibly\n   high above the ship\n2. Convicts were locked\n   below without food or\n   water\n3. The storms lasted for\n   thirteen days\n\nConcluding sentence (CS):\n\"The Cape storm tested\neveryone on board and\npushed the convicts to\ntheir limits.\"",
    NOTES_IDO_SPO,
    FOOTER
  );

  // Slide 8 -- We Do
  contentSlide(
    pres,
    "We Do",
    C.SECONDARY,
    "Together: Arrival at Port Jackson",
    [
      "Mini-whiteboards out. We build an SPO together",
      "Topic sentence: when, where, what doing? (the ship anchors at Port Jackson)",
      "Three details: anchor dropped, unloading parties, tents set up",
      "Concluding sentence: how do we wrap up? (a new chapter begins)",
      "Compare: how is your TS similar to mine? Different?",
    ],
    NOTES_WEDO,
    FOOTER
  );

  // Slide 9 + 9b -- CFU with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Is This a Strong Topic Sentence?",
      "Thumbs Up = yes  |  Thumbs Down = no",
      "\"The convicts had a hard time on the ship.\"\n\nCould this sentence be the topic sentence for any chapter?\n\nWhat is a strong topic sentence missing here?",
      NOTES_CFU,
      FOOTER
    ),
    (s) => {
      addCard(s, 0.5, SAFE_BOTTOM - 0.95, 9, 0.85, { strip: C.SUCCESS, fill: C.WHITE });
      s.addText("Thumbs Down -- too vague. A strong TS gives specifics: who, what, when, where, why, how", {
        x: 0.75, y: SAFE_BOTTOM - 0.88, w: 8.4, h: 0.70,
        fontSize: 16, fontFace: FONT_B, color: C.PRIMARY, bold: true,
        valign: "middle", margin: 0,
      });
      s.addNotes("SAY:\n- Thumbs Down. The sentence is too vague\n- It could be the topic sentence for almost any chapter -- it does not tell us WHICH hard time we are reading about\n- A strong TS gives the reader specifics: who, what, when, where, why, how\n- Better: 'When the ship hit the Cape of Good Hope, the convicts were locked below for thirteen days'\n\nDO:\n- Reveal the answer\n- Show the comparison between the vague version and a stronger version\n\n[General: CFU Reveal | VTLM 2.0: Formative Feedback]");
    }
  );

  // Slide 10 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write Your SPO");

    addInstructionCard(s, [
      { text: "On your worksheet:", role: "header" },
      { text: "First: Choose ONE topic from the worksheet" },
      { text: "Next: Write a topic sentence with WHO, WHAT DOING, WHEN, WHERE" },
      { text: "Then: Add 3 supporting details from the text" },
      { text: "Last: Write a concluding sentence that wraps up" },
    ], {
      x: 0.5, y: CONTENT_TOP, w: 9, h: 1.85,
      strip: C.PRIMARY, fill: C.WHITE,
      headerColor: C.PRIMARY,
    });

    const tipY = CONTENT_TOP + 2.00;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Remember", {
      x: 0.75, y: tipY + 0.10, w: 4, h: 0.32,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Your TS should give the reader specifics", options: { breakLine: true } },
      { text: "Details come from the text, not from your imagination", options: { breakLine: true } },
      { text: "Your CS should not repeat the TS word for word" },
    ], {
      x: 0.75, y: tipY + 0.46, w: 8.4, h: 0.85,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bullet: true, margin: 0,
      paraSpaceAfter: 2,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // Slide 11 -- Closing
  closingSlide(
    pres,
    "Turn to a partner: tell them ONE thing the author showed us about the journey today.",
    [
      "I can explain what the author wants us to know about the journey or a character",
      "I can use newly taught vocabulary in my own sentences",
      "I can summarise an aspect of the text using a topic sentence, supporting details and a concluding sentence",
    ],
    NOTES_CLOSING
  );

  // Slide 12 -- Resources


  // -----------------------------------------------------------------------
  // Companion PDFs
  // -----------------------------------------------------------------------

  // SPO Template
  const ws = createPdf({ title: SPO_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Single Paragraph Outline -- Summary", {
    color: C.PRIMARY,
    subtitle: "Chapters 25-26: Cape to Port Jackson",
    lessonInfo: "Lesson 12 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "Use this template to summarise ONE aspect of Chapters 25-26. Choose ONE topic from the list, then plan your summary.", wsY, { color: C.PRIMARY });

  wsY = addSectionHeading(ws, "Choose Your Topic (circle one)", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "1. The storm at the Cape of Good Hope", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "2. The arrival at Port Jackson", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "3. The captain's decision to sail away from the first land they reached", wsY, { fontSize: 11 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Topic Sentence (TS)", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "A strong topic sentence tells the reader: who, what doing, when, where, why or how.", wsY, { fontSize: 10, italic: true });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Supporting Details", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Find 3 details from the text that support your topic sentence. Short notes are fine.", wsY, { fontSize: 10, italic: true });
  wsY += 4;
  wsY = addBodyText(ws, "Detail 1:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;
  wsY = addBodyText(ws, "Detail 2:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;
  wsY = addBodyText(ws, "Detail 3:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Concluding Sentence (CS)", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Wrap up your summary. Do NOT repeat your topic sentence word for word.", wsY, { fontSize: 10, italic: true });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 10;

  wsY = addSectionHeading(ws, "Optional Challenge", wsY, { color: C.ALERT });
  wsY = addBodyText(ws, "Turn your SPO into a full summary paragraph (3-5 sentences) on the back of this page.", wsY, { fontSize: 10 });

  addPdfFooter(ws, "Lesson 12 | SPO Summary Template");

  // Answer key / model
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "SPO Summary -- Model Answers", {
    color: C.ALERT,
    subtitle: "Teacher Reference -- Chapters 25-26",
    lessonInfo: "Lesson 12 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Three model SPOs for the three topic options. Student work will vary; assess whether the TS is specific, the details come from the text, and the CS does not repeat the TS.", akY, { color: C.ALERT });

  akY = addSectionHeading(ak, "Topic 1 -- The Cape Storm", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "TS: When the ship reached the Cape of Good Hope, it met massive storms that left the convicts trapped below deck for days.", akY);
  akY = addBodyText(ak, "Detail 1: Waves rose impossibly high above the ship", akY);
  akY = addBodyText(ak, "Detail 2: Convicts were locked below without food or water", akY);
  akY = addBodyText(ak, "Detail 3: The storms lasted for thirteen days", akY);
  akY = addBodyText(ak, "CS: The Cape storm tested everyone on board and pushed the convicts to their limits.", akY);
  akY += 10;

  akY = addSectionHeading(ak, "Topic 2 -- Arrival at Port Jackson", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "TS: After nearly eight months at sea, the ship anchored at Port Jackson, ready to begin a new colony.", akY);
  akY = addBodyText(ak, "Detail 1: The captain decided Port Jackson was a better site than the first land they reached", akY);
  akY = addBodyText(ak, "Detail 2: They sailed just four hours from the first landing to Port Jackson", akY);
  akY = addBodyText(ak, "Detail 3: The ship dropped anchor and the unloading parties were prepared", akY);
  akY = addBodyText(ak, "CS: The voyage was over, but a completely new chapter was about to begin.", akY);
  akY += 10;

  akY = addSectionHeading(ak, "Topic 3 -- Sailing Away from the First Land", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "TS: When the ship first reached land, the captain decided to sail away because the land was too flat for crops.", akY);
  akY = addBodyText(ak, "Detail 1: The crew had been excited to step off the ship after months at sea", akY);
  akY = addBodyText(ak, "Detail 2: The captain inspected the land for five days", akY);
  akY = addBodyText(ak, "Detail 3: He decided Port Jackson, just four hours away, was a better choice", akY);
  akY = addBodyText(ak, "CS: The captain's choice meant another short journey, but a much better place to settle.", akY);
  akY += 10;

  akY = addSectionHeading(ak, "What to Look For", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- Topic sentence is specific (who/what/when/where/why/how), not vague", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- 3 supporting details that come from the text and genuinely support the TS", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Concluding sentence wraps up without repeating the TS word for word", akY, { fontSize: 10 });

  addPdfFooter(ak, "Lesson 12 | SPO Summary Model -- TEACHER COPY");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Lesson12.pptx` }),
    writePdf(ws, SPO_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Lesson12.pptx`);
  console.log("Done: " + SPO_RESOURCE.name);
  console.log("Done: " + ANSWER_KEY_RESOURCE.name);
}

build().catch(console.error);
