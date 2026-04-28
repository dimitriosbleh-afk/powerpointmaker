"use strict";

// Tom Unit -- Lesson 16: Chapters 31-33 -- Settling In with the Sergeant + Appositives
// Week 4, Lesson 16, Grade 5/6 Literacy
// Reading: Ch 31 (Tom arrives, digs holes with Rob), Ch 32 (Sergeant returns with rations), Ch 33 (Sergeant plants seeds)
// Sentence-level: Revise and add appositives

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

const SESSION_NUMBER = 16;
const FOOTER = "Chapters 31-33 | Lesson 16 | Week 4 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson16_Settling_In_Appositives";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const APPOS_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Appositives Practice",
  "Student worksheet: identify, match and add appositives to sentences from Chapters 31-33."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Appositives Practice Answer Key",
  "Teacher reference: model answers for the appositive practice worksheet."
);
const RESOURCE_ITEMS = [APPOS_RESOURCE, ANSWER_KEY_RESOURCE];
const APPOS_PDF_PATH = path.join(OUT_DIR, APPOS_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Lesson 16. Three short chapters today: 31, 32 and 33
- Tom arrives at the Sergeant's house. He meets Rob. Life in the colony begins for Tom
- After reading, we will return to appositives -- adding detail to a noun by renaming it

DO:
- Display title slide as students settle
- Have copies of the novel ready
- Have the appositives worksheet ready (do not distribute yet)

TEACHER NOTES:
This lesson combines novel reading (Chapters 31-33) with sentence-level writing (revise and add appositives). The chapters give Tom a new home and a new companion. The sentence work uses content from those chapters as practice material.

WATCH FOR:
- Students who need a recap of the previous chapters -- Tom escaped, was caught, ended up at the colony
- Students eager to read -- the chapters are short, so reading moves quickly

[Literacy: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands: reading the new chapters and analysing characters, then adding appositives in sentences
- Read the success criteria

DO:
- Choral read the LI, then the SCs
- Brief context: "Today is reading and writing. We read the chapters and stop to think about the characters. Then we use sentences from those chapters to practise appositives"

TEACHER NOTES:
SC1 targets character analysis through pause points. SC2 targets the structural understanding of appositives (placement and punctuation). SC3 targets the sentence-level application.

WATCH FOR:
- Students unsure what an appositive is -- reassure: "We will revise it together before we write"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_READING = `SAY:
- Three short chapters today. I will read aloud
- Three pause points. Each one, we stop and think about a character or the author's choice
- Get your novel ready and find Chapter 31

DO:
- Give 30 seconds for students to find the chapter
- Read Chapter 31 aloud at a steady pace
- At each pause point, display the matching slide and run the discussion (Turn & Talk -> Cold Call)
- Continue with Chapter 32, then Chapter 33

TEACHER NOTES:
Reading the three chapters with three pause points takes around 25-30 minutes. Move briskly between pauses. The pause questions are open and discussion-based, not requiring a single right answer.

WATCH FOR:
- Students whose attention drifts -- the pause points re-anchor focus
- Students with rich responses -- celebrate and use to scaffold others

[Literacy: Read Aloud | VTLM 2.0: Build Knowledge / Reading]`;

const NOTES_PAUSE1 = `SAY:
- Stop. We just read about Rob and Tom washing off the dirt and talking about their pasts and futures
- Turn to your partner. What have we learned about Rob and Tom from this conversation? How is the author making you feel about these characters right now? You have 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students for ideas
- Listen for: shared loss / shared loneliness, beginning of friendship, hope, both boys are vulnerable, the author wants us to feel sympathy

TEACHER NOTES:
This pause sets up the friendship between Rob and Tom. The author's choice to put them together in shared tasks (digging, washing) and shared talk (about pasts and futures) signals connection. Students should pick up on the sense of shared situation and growing trust.

WATCH FOR:
- Students who notice the author "letting" them connect -- great authorial-choice thinking
- Students who only summarise the events -- prompt: "How does the author make you FEEL about them?"

[Literacy: Pause Point | VTLM 2.0: Comprehension / Author Craft]`;

const NOTES_PAUSE2 = `SAY:
- Stop. The Sergeant has come back with rations and a fish. He has cooked dinner and shared it with the boys. He even threw the fish head to the hens
- Turn to your partner. How does what we just read add to our understanding of the Sergeant? How do things look for Tom now? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: the Sergeant is generous / kind / paternal; he treats the boys as family; Tom is becoming part of something for the first time; things are looking up for Tom

TEACHER NOTES:
This pause builds the Sergeant's character. He gives Tom his bed when Tom admits he is scared of the dark. The fish-head-to-hens detail shows everyday warmth. Students should notice the quiet care in the Sergeant's actions.

WATCH FOR:
- Students who see the Sergeant only as an authority figure -- redirect to his actions
- Students who notice the small kind gestures -- celebrate the close reading

[Literacy: Pause Point | VTLM 2.0: Comprehension / Character]`;

const NOTES_PAUSE3 = `SAY:
- Stop. Tom is looking at the seedlings the Sergeant has planted. He thinks: "One day, those will be food."
- Turn to your partner. How has the author let you know that something has changed? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: Tom is thinking about the future, he sees himself in this place, he has hope, food will not always be scarce, he is starting to belong

TEACHER NOTES:
This pause is a quiet but powerful change. Tom moves from running from the past to looking at a future. The author signals this through his thoughts -- not a dramatic event, but a shift in perspective. Strong students will notice this is the FIRST time we hear Tom think about the future positively.

WATCH FOR:
- Students who notice Tom's hope -- celebrate the inference
- Students who only describe the plants -- prompt: "What does Tom THINK as he looks at the plants?"

[Literacy: Pause Point | VTLM 2.0: Comprehension / Author Craft]`;

const NOTES_VOCAB = `SAY:
- One vocabulary word from the chapters: prejudice
- Prejudice means thinking badly of someone or a group before you really know them. It is judging without facts
- The chapters mention prejudices when Tom thinks about how the Sergeant is different from what he expected
- Say it with me: prejudice
- Use it: "Tom let go of his prejudices about marines when he met the Sergeant"

DO:
- Display the word
- Choral read the word
- Quick partner activity: "Tell your partner one example of when someone might show prejudice" (30 seconds)

TEACHER NOTES:
"Prejudice" is one of the explicit vocabulary words from this lesson. The other ("blunt") is left as incidental. Keep this brief -- the focus today is reading and sentence work.

WATCH FOR:
- Students confusing "prejudice" with "preference" -- redirect: prejudice is unfair judging, not just liking
- Students who give thoughtful examples -- evidence of understanding

[Literacy: Vocabulary | VTLM 2.0: Vocabulary Development]`;

const NOTES_REVISE = `SAY:
- Quick revision. An appositive is a noun or noun phrase that renames or describes the noun right next to it
- Read the example: "Tom, a young convict, dug holes in the garden"
- The appositive is "a young convict". It renames "Tom" with extra detail
- Notice: it is set off by commas. That is how you know it is non-essential -- the sentence still makes sense without it
- Why use them? They add detail without starting a new sentence. They keep your writing tight

DO:
- Display the example
- Point to the appositive between commas
- Read aloud the sentence WITH and WITHOUT the appositive: "Tom dug holes in the garden" -- still works. So the appositive is non-essential
- Brief check: "What does the appositive tell us about Tom?" (Cold Call)

TEACHER NOTES:
This is REVISION -- students have met appositives before. The 2-3 minute reminder is enough. The key reminders are: renames the noun beside it, often set off by commas, adds detail.

WATCH FOR:
- Students who confuse appositives with relative clauses -- key difference: appositive is a NOUN/NOUN PHRASE; relative clause has a relative pronoun (who, which, that)
- Students who can name the appositive in the example -- they are ready to practise

[Literacy: Revise | VTLM 2.0: Retention and Recall]`;

const NOTES_IDO = `SAY:
- Watch how I add an appositive to a sentence from our chapters
- Original: "Rob helped his father plant the seeds"
- I want to add detail about Rob. Who is Rob? He is the Sergeant's son. So my appositive is "the Sergeant's son"
- Where does it go? Right after Rob -- because that is the noun it describes
- I add commas around it because it is non-essential
- Final: "Rob, the Sergeant's son, helped his father plant the seeds"
- Notice: my sentence still makes sense without the appositive. That is the test

DO:
- Display the original sentence
- Talk through the choice of appositive
- Show the placement (between Rob and the verb)
- Show the commas
- Read aloud the final sentence

TEACHER NOTES:
The I Do uses content from Ch 33 (Sergeant plants seeds; Rob helps). This grounds the grammar work in the novel. The think-aloud is the modelling -- students see HOW you choose what to add and where to place it.

WATCH FOR:
- Students drafting their own appositive ideas already -- good engagement
- Students confused by where to place the appositive -- the rule is "right after the noun it describes"

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Read this sentence on your mini-whiteboard
- "Tom, a boy with no family of his own, settled into life at the cove"
- Underline the appositive. Hold it up
- Three, two, one -- show!

DO:
- Display the sentence
- Use Show Me Boards / mini-whiteboards
- Scan: students should underline "a boy with no family of his own"
- Cold Call 2 students: "Why is that the appositive?"

TEACHER NOTES:
This CFU checks identification of the appositive (SC2). The expected answer: "a boy with no family of his own" -- it renames Tom with extra detail, set off by commas. Students who identify only "a boy" or include the commas in their underline get partial credit.

WATCH FOR:
- Students who underline the wrong section -- they may be confusing with relative clauses or whole noun phrases
- Students who can articulate WHY -- they are ready for the You Do

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- The appositive is "a boy with no family of his own"
- It renames Tom with extra detail
- It sits between the commas -- non-essential

DO:
- Display the reveal
- Point to the appositive between commas
- Pivot if many students missed it: "Look between the commas. What is in there? It renames Tom"

CFU CHECKPOINT:
Technique: Show Me Boards (underline the appositive)
Script:
- "On your whiteboard, underline the appositive in the sentence"
- Scan for: "a boy with no family of his own" underlined (~80%)
PROCEED (>=80%): Most students underlined the correct phrase. Release to You Do.
PIVOT (<80%): Most likely issue -- students underline only the noun (a boy) or include the commas. Reteach: "An appositive is the chunk between the commas that renames the noun. Let us look again." Re-check with a second example.

TEACHER NOTES:
The reveal makes the answer visible. The pivot script handles the common partial-credit responses.

WATCH FOR:
- Students who self-correct after the reveal -- evidence of learning

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Time to practise. Use the appositives worksheet
- Section 1: identify the appositive (5 sentences)
- Section 2: match the appositive to the noun (5 sentences)
- Section 3: add an appositive using the bank (3 sentences)
- Section 4: add an appositive without a bank -- you choose (2 sentences)
- 15 minutes. Check punctuation -- commas around non-essential appositives

DO:
- Distribute the appositives worksheet
- Circulate -- prioritise enabling students first, then extenders
- Quick conferences: "Show me the appositive you added"
- Encourage cross-checking with a partner near the end

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Complete only Section 1 (identify) and Section 3 (add with a bank). Skip Section 4
- Provide the worksheet with sentences pre-marked: highlighted noun phrase shows where the appositive should go
- Extra Notes: These students focus on identification and selection, not generation

EXTENDING PROMPT:
- Task: After completing all sections, write 2 of their own sentences about Chapters 31-33 using appositives. The appositive must add genuine new information (not repeat the noun)
- Extra Notes: Peer-share with another extender

TEACHER NOTES:
The 15 minutes splits roughly 5/5/3/2 minutes per section. Active circulation is the formative assessment. The answer key is for teacher reference; students self-mark or peer-mark at the end.

WATCH FOR:
- Students who add an appositive but forget commas -- direct feedback: "Where do the commas go?"
- Students who add a clause instead of a noun phrase -- redirect: "An appositive is a noun phrase. A clause has a verb"
- Students who use the same appositive twice -- prompt to vary

[Literacy: You Do | VTLM 2.0: Supported Application / Practice]`;

const NOTES_CLOSING = `SAY:
- Quick self-check. Show me on your fingers, 1 to 5
- SC1: I analysed Tom and Rob's characters and the Sergeant from the chapters -- 1 to 5
- SC2: I can identify an appositive in a sentence -- 1 to 5
- SC3: I added an appositive with correct punctuation -- 1 to 5
- One thing you noticed about Tom in these chapters -- tell your partner

DO:
- Run each SC with finger ratings
- Note any patterns -- if many students score themselves low on SC3, plan a quick reteach next session
- Collect worksheets for marking
- Wrap: "Tomorrow we read more chapters. Tom and Rob start to build a chimney"

TEACHER NOTES:
Collecting worksheets gives diagnostic data. Students who score 1-2 on SC3 may need 1:1 conferencing on punctuation around appositives in the next lesson.

WATCH FOR:
- Students confidently sharing observations about Tom -- evidence of comprehension
- Students hesitant to share -- gentle prompt: "Even one detail is enough"

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${APPOS_RESOURCE.name} is your practice sheet -- identify, match and add appositives
- The ${ANSWER_KEY_RESOURCE.name} is for teacher use to mark or for self-marking at the end

DO:
- Print the practice sheet (one per student)
- Print one answer key for teacher use; do not distribute to students before they complete the work

TEACHER NOTES:
The practice sheet uses sentences with content from Chapters 31-33. The answer key includes possible answers; some sections (especially Section 4) have multiple valid answers.

[Literacy: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Tom Appleby Ch 31-33 + Appositives -- Lesson 16";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "Chapters 31-33",
    "Settling In + Adding Appositives",
    "Lesson 16  |  Week 4  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  // =========================================================================
  // SLIDE 2 -- LI / SC
  // =========================================================================
  liSlide(
    pres,
    [
      "We are learning to question the author's choices and analyse character development in Chapters 31-33, and to revise appositives by adding them to sentences with correct punctuation",
    ],
    [
      "I can describe what we learn about Tom, Rob and the Sergeant from the chapters",
      "I can identify the appositive in a sentence",
      "I can add an appositive to a sentence with correct comma punctuation",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Reading anchor
  // =========================================================================
  contentSlide(
    pres,
    "Read",
    C.PRIMARY,
    "Chapters 31, 32 & 33",
    [
      "Ch 31: Tom arrives at the Sergeant's house and meets Rob. They dig holes for crops",
      "Ch 32: Sergeant returns with rations and a fish. They share dinner. Tom sleeps in his bed",
      "Ch 33: Sergeant plants seeds. Tom adjusts and starts to feel at home",
      "Three pause points along the way -- listen for changes in the characters",
      "Find Chapter 31 in your novel",
    ],
    NOTES_READING,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Pause point 1 (p.170)
  // =========================================================================
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 31  |  After Tom & Rob talk",
    "...and began to wash off the dirt.",
    "p.170",
    "What have we learned about Rob and Tom from this conversation? How is the author making you feel about these characters right now?",
    NOTES_PAUSE1,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- Pause point 2 (p.173)
  // =========================================================================
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 32  |  After dinner",
    "The sergeant nodded, and threw the fish head to the hens.",
    "p.173",
    "How does what we just read add to our understanding of the Sergeant? How do things look for Tom now?",
    NOTES_PAUSE2,
    FOOTER
  );

  // =========================================================================
  // SLIDE 6 -- Pause point 3 (p.178)
  // =========================================================================
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 33  |  Tom looks at the seedlings",
    "One day, he thought, those will be food.",
    "p.178",
    "How has the author let you know that something has changed?",
    NOTES_PAUSE3,
    FOOTER
  );

  // =========================================================================
  // SLIDE 7 -- Vocabulary: prejudice
  // =========================================================================
  vocabSlide(
    pres,
    "prejudice",
    "noun",
    "Thinking badly of a person or a group before you really know them. Judging unfairly without good reason.",
    "Tom let go of his prejudices about marines when he saw the Sergeant's quiet kindness.",
    NOTES_VOCAB,
    FOOTER
  );

  // =========================================================================
  // SLIDE 8 -- Revise: appositives
  // =========================================================================
  contentSlide(
    pres,
    "Revise",
    C.SECONDARY,
    "Appositives -- Add Detail to a Noun",
    [
      "An appositive is a noun (or noun phrase) that renames the noun right beside it",
      "Example: \"Tom, a young convict, dug holes in the garden\"",
      "The appositive \"a young convict\" renames Tom with extra detail",
      "Set off non-essential appositives with commas (the sentence still works without it)",
      "Use them to add detail without starting a new sentence",
    ],
    NOTES_REVISE,
    FOOTER
  );

  // =========================================================================
  // SLIDE 9 -- I Do: Model adding an appositive
  // =========================================================================
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Add an Appositive to a Sentence",
    "Original sentence:\n\n\"Rob helped his father plant\nthe seeds.\"\n\nWho is Rob?\n -> the Sergeant's son\n\nWhere does the appositive go?\n -> right after \"Rob\"\n\nDo I need commas?\n -> yes, it is non-essential",
    "Final sentence:\n\n\"Rob, the Sergeant's son, helped his father plant the seeds.\"\n\nThe sentence still makes sense without the appositive. That is the test.\n\nNotice: the appositive renames Rob and tells the reader who he is, all without starting a new sentence.",
    NOTES_IDO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 10 + 11 -- CFU: Identify the appositive (reveal pair)
  // =========================================================================
  function buildCfuBase() {
    const slide = pres.addSlide();
    addTopBar(slide, C.ALERT);
    addBadge(slide, "CFU", { color: C.ALERT });
    addTitle(slide, "Underline the Appositive", { color: C.ALERT });

    const stampW = 1.3;
    slide.addShape("roundRect", {
      x: 9.5 - stampW, y: 0.20, w: stampW, h: 0.32, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
    });
    slide.addText("✓  CHECK", {
      x: 9.5 - stampW, y: 0.20, w: stampW, h: 0.32,
      fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    slide.addShape("roundRect", {
      x: 0.5, y: CONTENT_TOP, w: 2.6, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ALERT },
    });
    slide.addText("Show Me Boards", {
      x: 0.5, y: CONTENT_TOP, w: 2.6, h: 0.40,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const cardY = CONTENT_TOP + 0.55;
    const cardH = 1.95;
    addCard(slide, 0.5, cardY, 9, cardH, { strip: C.SECONDARY, fill: C.WHITE });
    slide.addText("Sentence", {
      x: 0.75, y: cardY + 0.10, w: 4, h: 0.26,
      fontSize: 12, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    slide.addText("\"Tom, a boy with no family of his own, settled into life at the cove.\"", {
      x: 0.75, y: cardY + 0.40, w: 8.5, h: cardH - 0.50,
      fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const promptY = cardY + cardH + 0.18;
    addCard(slide, 0.5, promptY, 9, SAFE_BOTTOM - promptY - 0.55, { strip: C.ACCENT, fill: C.BG_CARD });
    slide.addText("On your whiteboard:  underline the appositive  ->  hold it up.", {
      x: 0.75, y: promptY + 0.14, w: 8.5, h: SAFE_BOTTOM - promptY - 0.85,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
    });

    addFooter(slide, FOOTER);
    slide.addNotes(NOTES_CFU_BUILD);
    return slide;
  }

  withReveal(
    buildCfuBase,
    (slide) => {
      const revealY = 4.90;
      addCard(slide, 0.5, revealY, 9, 0.42, { fill: C.SUCCESS });
      slide.addText("Appositive:  \"a boy with no family of his own\"", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // =========================================================================
  // SLIDE 12 -- You Do: Worksheet practice
  // =========================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Practise Adding Appositives");

    addCard(s, 0.5, CONTENT_TOP, 9, 2.05, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Use the Appositives Practice worksheet", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("Section 1:  Identify the appositive (5 sentences)\nSection 2:  Match the appositive to the noun (5 sentences)\nSection 3:  Add an appositive using the bank (3 sentences)\nSection 4:  Add your own appositive -- no bank (2 sentences)", {
      x: 0.75, y: CONTENT_TOP + 0.46, w: 8.4, h: 1.40,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 2.20;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Time:  15 minutes", {
      x: 0.75, y: tipY + 0.10, w: 4, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Check your commas around non-essential appositives\n- An appositive is a noun phrase, not a clause -- no verb!\n- Try to vary your appositives -- do not repeat the same one", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: 0.85,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // =========================================================================
  // SLIDE 13 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    {
      reflectionPrompt: "Tell your partner one thing you noticed about Tom in these chapters.",
      scItems: [
        "I can describe what we learn about Tom, Rob and the Sergeant from the chapters",
        "I can identify the appositive in a sentence",
        "I can add an appositive to a sentence with correct comma punctuation",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (got it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 14 -- Resources
  // =========================================================================
  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: Appositives Practice ------------------------------------------
  const ws = createPdf({ title: APPOS_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Appositives Practice", {
    color: C.PRIMARY,
    subtitle: "Sentences from Chapters 31-33 of Tom Appleby",
    lessonInfo: "Lesson 16 | Week 4 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "An appositive is a noun (or noun phrase) that renames the noun right beside it. Use commas to set off non-essential appositives.", wsY, { color: C.PRIMARY });

  wsY = addSectionHeading(ws, "Section 1: Identify the Appositive", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Underline the appositive in each sentence.", wsY, { fontSize: 10, italic: true });
  wsY += 4;
  wsY = addBodyText(ws, "1.  Tom, a young convict from England, arrived at the Sergeant's house.", wsY);
  wsY = addBodyText(ws, "2.  Rob, the Sergeant's son, helped him dig holes for crops.", wsY);
  wsY = addBodyText(ws, "3.  The Sergeant, a quiet but firm man, returned with rations for the week.", wsY);
  wsY = addBodyText(ws, "4.  The cove, a place of new beginnings for Tom, became his home.", wsY);
  wsY = addBodyText(ws, "5.  The garden, a small patch beside the house, slowly came to life.", wsY);
  wsY += 8;

  wsY = addSectionHeading(ws, "Section 2: Match the Appositive to the Noun", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Draw a line from each noun on the left to the matching appositive on the right.", wsY, { fontSize: 10, italic: true });
  wsY += 4;
  wsY = addBodyText(ws, "Tom            ----  a kind sergeant who treated the boys like family", wsY);
  wsY = addBodyText(ws, "Rob            ----  a young English convict making a new home", wsY);
  wsY = addBodyText(ws, "The Sergeant   ----  the boy who shared the harbour bath", wsY);
  wsY = addBodyText(ws, "The fish       ----  a place of fresh beginnings far from England", wsY);
  wsY = addBodyText(ws, "The cove       ----  the supper that the Sergeant cooked over the fire", wsY);
  wsY += 8;

  wsY = addSectionHeading(ws, "Section 3: Add an Appositive (with a bank)", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "Choose an appositive from the bank and add it to each sentence. Use commas.", wsY, { fontSize: 10, italic: true });
  wsY = addBodyText(ws, "Bank:  a small wooden cabin  |  the Sergeant's son  |  a fish from the harbour", wsY, { fontSize: 10, bold: true });
  wsY += 4;
  wsY = addBodyText(ws, "1.  Rob _____________________________ helped Tom build the fire.", wsY);
  wsY = addBodyText(ws, "2.  The hut _____________________________ stood near the garden.", wsY);
  wsY = addBodyText(ws, "3.  Their dinner _____________________________ filled the room with steam.", wsY);
  wsY += 8;

  wsY = addSectionHeading(ws, "Section 4: Add Your Own Appositive", wsY, { color: C.ALERT });
  wsY = addBodyText(ws, "Add an appositive of your own to each sentence. Use commas. Make sure it adds new detail.", wsY, { fontSize: 10, italic: true });
  wsY += 4;
  wsY = addBodyText(ws, "1.  The Sergeant ________________________________________ returned at sundown.", wsY);
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 18 });
  wsY = addBodyText(ws, "2.  The garden ________________________________________ began to grow.", wsY);
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 18 });

  addPdfFooter(ws, "Lesson 16 | Appositives Practice");

  // --- PDF 2: Answer Key ----------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Appositives Practice -- Answer Key", {
    color: C.SECONDARY,
    subtitle: "Teacher Reference",
    lessonInfo: "Lesson 16 | Week 4 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Section 4 has many possible answers. Accept any appositive that renames the noun, adds new information, and is correctly punctuated with commas.", akY, { color: C.SECONDARY });

  akY = addSectionHeading(ak, "Section 1: Identify the Appositive", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "1.  a young convict from England", akY);
  akY = addBodyText(ak, "2.  the Sergeant's son", akY);
  akY = addBodyText(ak, "3.  a quiet but firm man", akY);
  akY = addBodyText(ak, "4.  a place of new beginnings for Tom", akY);
  akY = addBodyText(ak, "5.  a small patch beside the house", akY);
  akY += 8;

  akY = addSectionHeading(ak, "Section 2: Match the Appositive to the Noun", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "Tom -> a young English convict making a new home", akY);
  akY = addBodyText(ak, "Rob -> the boy who shared the harbour bath", akY);
  akY = addBodyText(ak, "The Sergeant -> a kind sergeant who treated the boys like family", akY);
  akY = addBodyText(ak, "The fish -> the supper that the Sergeant cooked over the fire", akY);
  akY = addBodyText(ak, "The cove -> a place of fresh beginnings far from England", akY);
  akY += 8;

  akY = addSectionHeading(ak, "Section 3: Add an Appositive (with a bank)", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "1.  Rob, the Sergeant's son, helped Tom build the fire.", akY);
  akY = addBodyText(ak, "2.  The hut, a small wooden cabin, stood near the garden.", akY);
  akY = addBodyText(ak, "3.  Their dinner, a fish from the harbour, filled the room with steam.", akY);
  akY += 8;

  akY = addSectionHeading(ak, "Section 4: Add Your Own Appositive (sample answers)", akY, { color: C.ALERT });
  akY = addBodyText(ak, "1.  The Sergeant, a kind man with weather-worn hands, returned at sundown.", akY);
  akY = addBodyText(ak, "2.  The garden, a small but precious patch of soil, began to grow.", akY);
  akY = addBodyText(ak, "Many other appositives are valid. Check for: noun phrase (no verb), correct placement, and commas.", akY, { fontSize: 10, italic: true });

  addPdfFooter(ak, "Lesson 16 | Appositives Answer Key -- TEACHER USE");

  // --- Write all files ----------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Lesson16.pptx` }),
    writePdf(ws, APPOS_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Lesson16.pptx`);
  console.log("Done: " + APPOS_RESOURCE.name);
  console.log("Done: " + ANSWER_KEY_RESOURCE.name);
}

build().catch(console.error);
