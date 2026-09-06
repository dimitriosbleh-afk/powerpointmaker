"use strict";

// Tom Unit -- Lesson 11: Chapters 22-24 + Sentence-level Conjunctions (because/but/so)
// Week 3, Lesson 11, Year 5/6 Literacy
// Reading: Chapters 22-24 (rough seas, doldrums, funeral, Rio, Cape Town, Miss Hildegard)
// Sensitivity: Chapter 22 contains content about death.
// Writing: Use because, but, so to complete comprehension responses.

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

const SESSION_NUMBER = 11;
const FOOTER = "Chapters 22-24 | Lesson 11 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson11_Across_the_Sea";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PRACTICE_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Conjunctions Practice",
  "Student worksheet: complete sentences using because, but and so in response to comprehension questions about Chapters 22-24."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Conjunctions Answer Key",
  "Teacher reference: model answers for the conjunctions practice."
);
const SENTENCE_EXPANSION_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Sentence Expansion Extension",
  "Optional follow-up: expand kernel sentences with when, where, how and why after the because/but/so work."
);
const RESOURCE_ITEMS = [PRACTICE_RESOURCE, ANSWER_KEY_RESOURCE, SENTENCE_EXPANSION_RESOURCE];
const PRACTICE_PDF_PATH = path.join(OUT_DIR, PRACTICE_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
const SENTENCE_EXPANSION_PDF_PATH = path.join(OUT_DIR, SENTENCE_EXPANSION_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Three chapters today: 22, 23, 24
- The journey across the sea continues. You will see how Tom changes as the voyage gets harder
- After reading we go to sentence work using because, but and so

DO:
- Display title slide as students settle
- Have novels and exercise books on desks

SENSITIVITY ADVISORY:
- What it is: Chapter 22 includes the funeral of Rob's mother at sea
- Framing language: "We are going to read about a funeral. Authors include hard moments to help us feel for the characters. We will read carefully and respectfully"
- Watch for: students who become quiet or withdraw -- check in privately after the read
- Protocol: pause if needed; if a student is distressed, allow them to step out with a buddy

TEACHER NOTES:
Lesson 11 of the Tom unit. Chapter 22 contains death (the funeral at sea); Chapter 23 covers the long stop at Rio and Cape Town and Sam's worsening health; Chapter 24 jumps to the older Thomas and his great-granddaughter's teacher. Three chapters is a lot -- pace the read and pause only at the marked points.

WATCH FOR:
- Students sensitive to the funeral content -- check privately after the read
- Students who notice the time jump in Chapter 24 -- excellent comprehension

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands today. First, what the author shows us about the characters. Second, using because, but and so to write clearer answers
- Read each success criterion together

DO:
- Choral read the LI, then each SC
- Brief reminder: "We have used because, but and so before. Today we use them to answer comprehension questions about the text"

TEACHER NOTES:
SC1 targets close reading and authorial choice. SC2 targets vocabulary application. SC3 targets correct conjunction use in written responses.

WATCH FOR:
- Students unsure about "authorial choices" -- restate: "What the author wants us to know"
- Students confident about the conjunctions -- they can support peers in the You Do

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_READING = `SAY:
- Reading mode: teacher read aloud
- Three chapters. Listen for how Tom feels as the journey gets harder
- I will pause twice for us to think together

DO:
- Read Chapter 22, 23, 24 aloud
- Plan two pause points: p.111 (after Rob's funeral) and p.119 (Thomas and Miss Hildegard)
- Brief note: Chapter 24 is a flash-forward to older Thomas. If students seem confused, clarify: "This is the same Tom, much later in his life"

TEACHER NOTES:
The two pause points target two different authorial moves: showing emotion through restraint (Rob's reaction), and revealing values through dialogue (Thomas insisting Millie not be locked in a cupboard). Students who pick up on the emotional restraint and the link back to Tom's own past show strong analytical reading.

WATCH FOR:
- Students moved by the funeral scene -- acknowledge and move on respectfully
- Students confused by the Chapter 24 time jump -- a brief note clarifies

[General: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause here. Rob's mother has just been buried at sea. Tom tries to speak across the decks. Rob looks at him and goes below
- Look at the words: "He nodded once, his face expressionless, then turned and went below"
- What do you think the author wants us to know about Rob? [Rob is grief-stricken but holding it in. Expressionless does not mean he feels nothing -- it means he cannot let it show. Going below is an act of withdrawal, of needing to be alone]

DO:
- Display the quote
- Think-Pair-Share: 20 seconds think, 30 seconds pair, share
- Push for the difference between "no feeling" and "hiding feeling"

TEACHER NOTES:
This pause point is about authorial restraint -- showing grief through what is NOT said or shown. Rob's expressionless face is a deliberate choice that asks the reader to fill in the emotion.

WATCH FOR:
- Students who say "he doesn't care" -- redirect: "If he didn't care, would he go below straight away?"
- Students who connect Rob's silence to Tom's earlier suffering -- excellent

[General: Pause Point 1 | VTLM 2.0: Higher-Order Questioning]`;

const NOTES_PAUSE2 = `SAY:
- Pause here. Thomas is talking to Miss Hildegard about Millie. He says: "But he wasn't having a child of his blood locked in darkness"
- Given what we already know about Tom's past, what do you think is going on here? [Thomas was locked in darkness himself -- in the prison hulks, in the ship's hold. He cannot bear the thought of any child of his family going through that]
- What does this show us about Thomas? [He has not forgotten what happened to him. He is using his power as an elder to protect Millie]

DO:
- Display the quote
- Cold Call after thinking time
- Push for the connection back to Tom's earlier experiences

CFU CHECKPOINT:
Technique: Cold Call (after thinking time)

Script:
- "30 seconds to think. What does this line show us about Thomas?"
- Cold call 2-3 students
- Scan for: connection to Tom's past as a convict, his time in darkness

PROCEED (>=80%): Most students connect this line to Tom's past. Continue.
PIVOT (<80%): Most likely issue -- students miss the link to Tom's own experience. Reteach: "Where in this novel have we seen people locked in darkness? [the prison hulks, the ship's hold]. Now Thomas is the one with the power. What is he choosing to do with it?" Re-check: "Why is this line important?"

TEACHER NOTES:
The pause connects Thomas's adult choices to young Tom's experiences. The author is showing how trauma shapes later decisions. This is a strong example of authorial perspective made evident through dialogue.

WATCH FOR:
- Students who connect to Tom's "courage cloak" metaphor (p.105) -- excellent
- Students who notice Thomas's tone of authority -- strong analytical reading

[General: Pause Point 2 | VTLM 2.0: Deep Comprehension]`;

const NOTES_VOCAB = `SAY:
- Three explicit words from today's chapters. Each one helps us picture Tom's experience
- Unbearable -- so bad you cannot stand it. The smell below deck was unbearable
- Languid -- slow and lazy, no energy. The convicts moved languidly because they had been at sea for so long
- Rationed -- given out in small, controlled amounts. Food and water were rationed during the storms

DO:
- Read each word and meaning aloud
- Quick oral routine: "Use unbearable in a sentence about Tom's voyage" -- 30 seconds, partner share
- Note pronunciation: lan-GWID; RA-shund

TEACHER NOTES:
Three explicit words drawn directly from Chapters 22-23. Keep this brief -- 3 minutes. Students will use these words during the You Do.

WATCH FOR:
- Students who confuse rationed with reasoned -- emphasise pronunciation
- Students who use the words naturally during pause-point discussion -- celebrate

[General: Vocabulary | VTLM 2.0: Vocabulary Instruction]`;

const NOTES_IDO = `SAY:
- Now we move to writing. You know because, but and so. Today we use them to answer comprehension questions
- Comprehension question: Why do you think Rob left the deck without speaking?
- Watch me complete the sentence three different ways
- Because: "Rob left the deck without speaking BECAUSE he could not let his grief show in front of the convicts." Because tells us WHY
- But: "Rob nodded at Tom BUT he could not let himself speak." But tells us a change in direction or contrast
- So: "Rob's mother had just died, SO he went below to be alone with his grief." So tells us what happens as a result
- Notice: each conjunction changes what kind of sentence we are writing -- a reason, a contrast, or a consequence

DO:
- Display the comprehension question and the three model sentences
- Highlight the conjunction in each
- Think aloud: "Each one is a correct answer. They show different parts of the same idea"

TEACHER NOTES:
The I Do shows that conjunctions are not interchangeable. Because = reason, but = contrast, so = consequence. Students often default to one and miss the others.

WATCH FOR:
- Students who think the three sentences mean the same thing -- they need to see that each one tells the reader something different
- Students who already use these conjunctions correctly -- they can support peers

[General: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Together. Comprehension question: How did Tom feel when he heard about Rob's mother?
- We write three answers -- one with because, one with but, one with so
- Mini-whiteboards out. We write together

DO:
- Build the answers as a class on the board
- Sample answers (accept variations):
  - Because: "Tom felt sad BECAUSE he knew what it was like to lose someone he loved"
  - But: "Tom wanted to speak to Rob BUT the deck rules kept them apart"
  - So: "Tom felt sorry for Rob, SO he tried to call across the decks"
- Discuss: "How does the meaning change with each conjunction?"

TEACHER NOTES:
The We Do is collaborative -- the teacher takes the pen and writes with student input. Each conjunction must produce a sensible sentence that answers the question.

WATCH FOR:
- Students who use because correctly but struggle with but -- prompt: "But shows a CHANGE. What changes here?"
- Students who write run-on sentences -- model the comma rule: "BUT and SO usually need a comma before them when joining two ideas"

[General: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU = `SAY:
- Quick check before you write on your own
- Look at this sentence: "Sam was getting more unwell, ___ Tom worried about him"
- Which conjunction fits best -- because, but, or so?
- Show me on your fingers: 1 = because, 2 = but, 3 = so

DO:
- Use Finger Voting (1 / 2 / 3)
- Scan for: 3 (so) -- Tom's worry is the consequence of Sam getting unwell
- Cold call 1-2 students: "Why so?"

CFU CHECKPOINT:
Technique: Finger Voting (1 = because, 2 = but, 3 = so)

Script:
- "Sam was getting more unwell, BLANK Tom worried about him. Which fits? 1, 2 or 3?"
- Scan for: mostly 3
- Follow up: "Why is so the best fit?" [Tom's worry is the result. Sam getting unwell caused the worry]

PROCEED (>=80%): Most show 3 with reasoning. Release to write.
PIVOT (<80%): Most likely issue -- students choose 1 (because) because the structure feels reason-like. Reteach: "Read it as 'BECAUSE Tom worried about him' -- does Tom's worry CAUSE Sam to be unwell? No. So the cause-effect is the other way. Sam unwell -> Tom worried. That is so." Re-check: "Try the sentence with each one. Which one tells us the result?"

TEACHER NOTES:
This CFU checks discrimination between the three conjunctions. The trap is that "because" feels right structurally but the cause-effect runs the wrong way.

WATCH FOR:
- Students who choose 1 (because) -- they need help seeing the cause-effect direction
- Students who choose 3 (so) and explain it correctly -- ready to write independently

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Your turn. On your worksheet, complete each sentence with the right conjunction and finish the sentence
- Use the comprehension questions to guide your answers
- 10 minutes. I will circulate

DO:
- Distribute the Conjunctions Practice worksheet
- Circulate -- check that students are choosing the conjunction that fits the meaning, not just the first one that comes to mind
- Confer with enabling and extending students after the first 3 minutes

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the worksheet but with a sentence bank provided. Students choose the best ending for each sentence stem from a list of 3 options. Reduces the writing demand while keeping the conjunction-meaning focus
- Extra Notes: Refer these students to the I Do model sentences as a reference

EXTENDING PROMPT:
- Task: After completing the worksheet, write 3 of your own comprehension question + answer pairs about Chapters 22-24. Each answer must use a different conjunction (because, but, so). Show how the meaning changes with each choice
- Alternative: hand the student the school Sentence Expansion Practice sheet (Section 1: Identify the Expansions, then Section 2). Frame it as "same family as because/but/so -- both add information to a thin sentence." Note that the sample sentences look ahead to Chapters 36-39

TEACHER NOTES:
The 10-minute write is the application of the conjunction work. Students should be discriminating between the three based on the meaning of their answer. Fast finishers can move to the school Sentence Expansion Practice sheet -- same skill family (growing a sentence) with adverbial prompts (when/where/how/why) instead of conjunctions.

WATCH FOR:
- Students who write run-on sentences -- model the comma rule
- Students who use only because -- prompt: "Try this answer with so. Does it still work?"
- Students using all three accurately -- celebrate and move them to the extending task

[General: You Do | VTLM 2.0: Independent Application]`;

const NOTES_CLOSING = `SAY:
- Quick self-check against the success criteria. Thumbs for each
- Then turn to a partner: tell them ONE thing the author showed us about a character today

DO:
- Run thumbs check for each SC
- Listen in on partner shares -- diagnostic for the next reading lesson
- Wrap up: "Tomorrow we keep reading. Sam is unwell. The Cape of Good Hope is coming"

TEACHER NOTES:
The closing connects today's reading and writing. Students who name a specific authorial choice (Rob's expressionless face, Thomas insisting on protection) show strong analytical thinking.

WATCH FOR:
- Students "thumbs down" on SC3 -- they need more conjunction practice next session
- Students sharing strong character insights -- excellent

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Three resources for this session
- The ${PRACTICE_RESOURCE.name} is for completing comprehension answers using because, but and so
- The ${ANSWER_KEY_RESOURCE.name} is for teacher reference
- The ${SENTENCE_EXPANSION_RESOURCE.name} is an early-finisher or follow-up sheet

DO:
- Print the practice worksheet (one per student)
- Print the answer key (teacher copy only)
- Have the extension sheet ready for fast finishers

TEACHER NOTES:
Students keep their completed worksheet for their writing portfolio. The answer key shows model answers; accept reasonable variations as long as the conjunction matches the meaning.

About the sentence expansion extension: the technique it practises -- expanding a kernel sentence by answering when, where, how and why -- is the same family of moves as adding a because, but or so clause. Both teach students to grow a thin sentence into a richer one. Use it as follows:
- After the You Do conjunctions worksheet is complete, offer the sentence expansion sheet to students who finish early or who would benefit from extra writing reps
- The examples stay within today's voyage context so students can practise the writing move without needing later chapter knowledge
- Section 1 gives prompts; Section 2 asks students to choose their own details

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Chapters 22-24: Across the Sea -- Lesson 11";

  // Slide 1 -- Title
  titleSlide(
    pres,
    "Chapters 22-24",
    "Across the Sea",
    "Lesson 11  |  Year 5/6 Literacy",
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
      "We are learning to notice what the author shows us about a character, and to use because, but and so to write clear comprehension answers",
    ],
    [
      "I can explain what the author wants us to know about a character",
      "I can use newly taught vocabulary in my own sentences",
      "I can complete a comprehension answer using because, but or so to fit the meaning",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // Slide 3 -- Reading Launch
  contentSlide(
    pres,
    "Teacher Read Aloud",
    C.PRIMARY,
    "Chapters 22 - 24",
    [
      "Reading Mode: Teacher Read Aloud",
      "Chapter 22: rough seas, the doldrums, a funeral at sea",
      "Chapter 23: Rio and Cape Town -- Sam grows weaker",
      "Chapter 24: an older Thomas talks with Millie's teacher",
      "Focus: how does the author show us how the characters feel?",
    ],
    NOTES_READING,
    FOOTER
  );

  // Slide 4 -- Pause Point 1
  quoteSlide(
    pres,
    "Pause Point 1",
    "Chapter 22 -- p. 111",
    "He nodded once, his face expressionless, then turned and went below.",
    "p. 111",
    "What do you think the author wants us to know about Rob? Why \"expressionless\" -- not \"sad\"?",
    NOTES_PAUSE1,
    FOOTER
  );

  // Slide 5 -- Pause Point 2
  quoteSlide(
    pres,
    "Pause Point 2",
    "Chapter 24 -- p. 119",
    "But he wasn't having a child of his blood locked in darkness.",
    "p. 119",
    "Given what we know about Tom's past, what is going on here? What does this show us about Thomas?",
    NOTES_PAUSE2,
    FOOTER
  );

  // Slide 6 -- Vocabulary (compact 3-word card)
  contentSlide(
    pres,
    "Vocabulary",
    C.SECONDARY,
    "Three Words from Today's Reading",
    [
      "unbearable -- so bad you cannot stand it (the stench below deck was unbearable)",
      "languid -- slow and lazy, with no energy (the convicts moved languidly after weeks at sea)",
      "rationed -- given out in small, controlled amounts (food and water were rationed during the storms)",
    ],
    NOTES_VOCAB,
    FOOTER
  );

  // Slide 7 -- I Do: model conjunctions on a comprehension question
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "because  |  but  |  so",
    "Comprehension question:\n\nWhy do you think Rob\nleft the deck without\nspeaking?\n\n\nThree conjunctions.\nThree different kinds of\nanswers.",
    "BECAUSE (a reason):\n\"Rob left the deck without speaking BECAUSE he could not let his grief show.\"\n\nBUT (a contrast):\n\"Rob nodded at Tom BUT he could not let himself speak.\"\n\nSO (a consequence):\n\"Rob's mother had just died, SO he went below to be alone.\"",
    NOTES_IDO,
    FOOTER
  );

  // Slide 8 -- We Do
  contentSlide(
    pres,
    "We Do",
    C.SECONDARY,
    "Together: How Did Tom Feel About Rob?",
    [
      "Mini-whiteboards out. We write three answers together",
      "BECAUSE (reason): Tom felt sad BECAUSE...",
      "BUT (contrast): Tom wanted to speak to Rob BUT...",
      "SO (consequence): Tom felt sorry for Rob, SO...",
      "How does each conjunction change what we are saying?",
    ],
    NOTES_WEDO,
    FOOTER
  );

  // Slide 9 + 9b -- CFU with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Which Conjunction Fits?",
      "1 = because  |  2 = but  |  3 = so",
      "\"Sam was getting more unwell, _____ Tom worried about him.\"\n\nFingers: 1 = because, 2 = but, 3 = so.\n\nWhich conjunction fits the meaning best?",
      NOTES_CFU,
      FOOTER
    ),
    (s) => {
      addCard(s, 0.5, SAFE_BOTTOM - 0.95, 9, 0.85, { strip: C.SUCCESS, fill: C.WHITE });
      s.addText("3 = SO   --   Tom's worry is the consequence of Sam being unwell", {
        x: 0.75, y: SAFE_BOTTOM - 0.88, w: 8.4, h: 0.70,
        fontSize: 18, fontFace: FONT_B, color: C.PRIMARY, bold: true,
        valign: "middle", margin: 0,
      });
      s.addNotes("SAY:\n- The answer is SO. Tom's worry is the result of Sam being unwell\n- Test it: 'BECAUSE Tom worried about him' would mean Tom's worry caused Sam to be unwell. That is the wrong direction\n- 'BUT Tom worried about him' breaks the link -- it does not match the meaning\n\nDO:\n- Reveal the answer\n- If the class is split, run the test: try each conjunction in turn and check the meaning\n\n[General: CFU Reveal | VTLM 2.0: Formative Feedback]");
    }
  );

  // Slide 10 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Conjunctions Practice");

    addInstructionCard(s, [
      { text: "On your worksheet:", role: "header" },
      { text: "First: Read each comprehension question" },
      { text: "Next: Choose the conjunction that fits the meaning" },
      { text: "Then: Complete the sentence using your answer" },
    ], {
      x: 0.5, y: CONTENT_TOP, w: 9, h: 1.55,
      strip: C.PRIMARY, fill: C.WHITE,
      headerColor: C.PRIMARY,
    });

    const tipY = CONTENT_TOP + 1.70;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Remember", {
      x: 0.75, y: tipY + 0.10, w: 4, h: 0.32,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "BECAUSE = reason (tells us WHY)", options: { breakLine: true } },
      { text: "BUT = contrast (tells us a CHANGE)", options: { breakLine: true } },
      { text: "SO = consequence (tells us WHAT HAPPENS as a result)", options: { breakLine: true } },
      { text: "Use a comma before BUT and SO when joining two ideas" },
    ], {
      x: 0.75, y: tipY + 0.46, w: 8.4, h: 1.10,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bullet: true, margin: 0,
      paraSpaceAfter: 2,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // Slide 11 -- Closing
  closingSlide(
    pres,
    "Turn to a partner: tell them ONE thing the author showed us about a character today.",
    [
      "I can explain what the author wants us to know about a character",
      "I can use newly taught vocabulary in my own sentences",
      "I can complete a comprehension answer using because, but or so to fit the meaning",
    ],
    NOTES_CLOSING
  );

  // Slide 12 -- Resources


  // -----------------------------------------------------------------------
  // Companion PDFs
  // -----------------------------------------------------------------------

  // Practice worksheet
  const ws = createPdf({ title: PRACTICE_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Conjunctions Practice -- because, but, so", {
    color: C.PRIMARY,
    subtitle: "Chapters 22-24: Across the Sea",
    lessonInfo: "Lesson 11 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "Read each comprehension question. Then complete the sentence using the conjunction in bold. Make sure your answer fits the meaning of the conjunction:", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "BECAUSE = reason (why)    BUT = contrast (a change)    SO = consequence (what happens as a result)", wsY, { fontSize: 11, italic: true });
  wsY += 6;

  wsY = addSectionHeading(ws, "Question 1: Why was the deck empty after the funeral?", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "The deck was empty BECAUSE", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 6;

  wsY = addSectionHeading(ws, "Question 2: How did Rob feel about his mother's death?", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Rob's face was expressionless BUT", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 6;

  wsY = addSectionHeading(ws, "Question 3: Why did Tom try to speak to Rob?", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "Rob had just lost his mother, SO", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 6;

  wsY = addSectionHeading(ws, "Question 4: What did Sam's worsening health show us?", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Sam was getting weaker BECAUSE", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 6;

  wsY = addSectionHeading(ws, "Question 5: Why did Thomas insist Millie not be locked in a cupboard?", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Thomas had been imprisoned as a boy, SO", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Self-Check", wsY, { color: C.ALERT });
  wsY = addBodyText(ws, "- Does each answer fit the meaning of the conjunction?", wsY, { fontSize: 10 });
  wsY = addBodyText(ws, "- Did I add a comma before BUT and SO when joining two ideas?", wsY, { fontSize: 10 });
  wsY = addBodyText(ws, "- Are my answers full sentences (not just one or two words)?", wsY, { fontSize: 10 });

  addPdfFooter(ws, "Lesson 11 | Conjunctions Practice");

  // Answer key
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Conjunctions Practice -- Model Answers", {
    color: C.ALERT,
    subtitle: "Teacher Reference -- Chapters 22-24",
    lessonInfo: "Lesson 11 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Model answers below. Student answers will vary -- accept any answer that (a) fits the conjunction's meaning and (b) is supported by the text.", akY, { color: C.ALERT });

  akY = addSectionHeading(ak, "Question 1 (because = reason)", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "The deck was empty BECAUSE the convicts had been ordered below for the funeral.", akY);
  akY = addBodyText(ak, "Look for: a reason, drawn from the text. Acceptable variations: weather, mourning, ship's rules.", akY, { fontSize: 10, italic: true });
  akY += 8;

  akY = addSectionHeading(ak, "Question 2 (but = contrast)", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "Rob's face was expressionless BUT his grief was clear in the way he turned and went below.", akY);
  akY = addBodyText(ak, "Look for: a contrast between Rob's blank face and his clear inner grief.", akY, { fontSize: 10, italic: true });
  akY += 8;

  akY = addSectionHeading(ak, "Question 3 (so = consequence)", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "Rob had just lost his mother, SO Tom called across the decks to try to comfort him.", akY);
  akY = addBodyText(ak, "Look for: a consequence -- something Tom does or feels as a result of Rob's loss.", akY, { fontSize: 10, italic: true });
  akY += 8;

  akY = addSectionHeading(ak, "Question 4 (because = reason)", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sam was getting weaker BECAUSE the food and water were rationed and the conditions below deck were unbearable.", akY);
  akY = addBodyText(ak, "Look for: a textual reason. Acceptable variations: illness, lack of fresh air, poor diet.", akY, { fontSize: 10, italic: true });
  akY += 8;

  akY = addSectionHeading(ak, "Question 5 (so = consequence)", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "Thomas had been imprisoned as a boy, SO he could not bear the thought of any child of his family being locked in darkness.", akY);
  akY = addBodyText(ak, "Look for: a connection back to Thomas's own experience -- a strong analytical move.", akY, { fontSize: 10, italic: true });
  akY += 12;

  akY = addSectionHeading(ak, "What to Look For", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- Conjunction matches the meaning (reason / contrast / consequence)", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Answer is supported by the text (Chapters 22-24)", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Comma before BUT and SO when joining two ideas", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Answer is a full sentence -- a clause after the conjunction", akY, { fontSize: 10 });

  addPdfFooter(ak, "Lesson 11 | Answer Key -- TEACHER COPY");

  // Sentence expansion extension
  const ext = createPdf({ title: SENTENCE_EXPANSION_RESOURCE.name });
  let extY = addPdfHeader(ext, "Sentence Expansion Extension", {
    color: C.SECONDARY,
    subtitle: "Grow a thin sentence with when, where, how and why details",
    lessonInfo: "Lesson 11 | Year 5/6 Literacy",
    showNameDate: true,
  });

  extY = addTipBox(ext, "Start with the kernel sentence. Add one or two details that answer when, where, how or why. Keep the sentence clear.", extY, { color: C.SECONDARY });
  extY = addBodyText(ext, "Example: Tom waited. -> After the funeral, Tom waited quietly by the rail because he wanted to speak to Rob.", extY, { fontSize: 10, italic: true });
  extY += 6;

  [
    ["Tom called across the deck.", "Add where and why."],
    ["Rob turned away.", "Add how and why."],
    ["The ship drifted.", "Add when and where."],
  ].forEach(([kernel, prompt], index) => {
    extY = addSectionHeading(ext, `Part 1.${index + 1}: ${kernel}`, extY, { color: C.PRIMARY });
    extY = addBodyText(ext, prompt, extY, { fontSize: 10, italic: true });
    extY = addLinedArea(ext, extY, 2, { lineSpacing: 24 });
    extY += 6;
  });

  extY = addSectionHeading(ext, "Part 2: Choose your own details", extY, { color: C.ACCENT });
  extY = addBodyText(ext, "Kernel sentence: Sam coughed.", extY, { fontSize: 11 });
  extY = addBodyText(ext, "Write one expanded sentence. Include at least two useful details.", extY, { fontSize: 10, italic: true });
  extY = addLinedArea(ext, extY, 3, { lineSpacing: 24 });
  extY += 8;
  extY = addSectionHeading(ext, "Self-check", extY, { color: C.ALERT });
  extY = addBodyText(ext, "- Did I add details that answer when, where, how or why?", extY, { fontSize: 10 });
  extY = addBodyText(ext, "- Does the expanded sentence still make sense?", extY, { fontSize: 10 });
  addPdfFooter(ext, "Lesson 11 | Sentence Expansion Extension");

  // Write all
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Lesson11.pptx` }),
    writePdf(ws, PRACTICE_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
    writePdf(ext, SENTENCE_EXPANSION_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Lesson11.pptx`);
  console.log("Done: " + PRACTICE_RESOURCE.name);
  console.log("Done: " + ANSWER_KEY_RESOURCE.name);
  console.log("Done: " + SENTENCE_EXPANSION_RESOURCE.name);
}

build().catch(console.error);
