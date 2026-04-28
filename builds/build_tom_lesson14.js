"use strict";

// Tom Unit -- Lesson 14: Chapters 29-30 + Plan Body Paragraph 2 (Why the First Fleet came to Australia)
// Week 3, Lesson 14, Year 5/6 Literacy (Cross-curricular HASS link)
// Reading: Chapters 29-30 (Tom is interrogated, taken into Sgt Stanley's household, walks the harbour)
// Writing: Plan a body paragraph for an information report using a Single Paragraph Outline (SPO).
// Topic: Why the First Fleet came to Australia.

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

const SESSION_NUMBER = 14;
const FOOTER = "Chapters 29-30 + First Fleet | Lesson 14 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson14_Free_and_First_Fleet";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const SPO_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "SPO Plan Body Paragraph 2",
  "Student template: plan a body paragraph for an information report on Why the First Fleet came to Australia."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "SPO Model Body Paragraph 2",
  "Teacher reference: model SPO with mentor topic sentence and concluding sentence options."
);
const RESOURCE_ITEMS = [SPO_RESOURCE, ANSWER_KEY_RESOURCE];
const SPO_PDF_PATH = path.join(OUT_DIR, SPO_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Two chapters today: 29 and 30. Tom's life changes
- After reading we go back to our information report. We plan body paragraph two -- about why the First Fleet came to Australia
- This lesson connects fiction and history -- the story we are reading is set during the events we are reporting on

DO:
- Display title slide as students settle
- Have novels, exercise books and the non-fiction article on desks

TEACHER NOTES:
Lesson 14 of the unit. Two strands today: reading (Chapters 29-30) and text-level writing (planning body paragraph 2 of the information report). Pace the reading tightly so students have time for the SPO planning.

WATCH FOR:
- Students who remember body paragraph 1 from Lesson 10 -- they have a model to work from
- Students unsure why the lesson covers two things -- frame it as "novel and history side by side"

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands. First, what the author shows us about Tom's freedom. Second, planning a body paragraph for our information report
- Read each success criterion together

DO:
- Choral read the LI, then each SC
- Brief reminder: "Body paragraph 1 was about life in 18th century England. Body paragraph 2 is about why the First Fleet came here"

TEACHER NOTES:
SC1 targets reading analysis. SC2 targets the structural understanding (TS + supporting details + CS). SC3 targets the planning skill applied to a new topic.

WATCH FOR:
- Students unsure about the SPO structure -- they will see it modelled
- Students who used SPO in Lessons 9 and 12 -- they can support peers

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_READING = `SAY:
- Reading mode: teacher read aloud
- Two short chapters today. Tom's life changes completely
- I will pause twice for us to think about the author's choices
- After the read, we move to the writing

DO:
- Read Chapter 29 and Chapter 30
- Plan two pause points: p.158 (the moment Sgt Stanley is named) and p.160 (the walk around the harbour)
- Keep the reading tight -- 12 minutes maximum so the writing has time

TEACHER NOTES:
The reading is shorter than recent lessons. Both chapters are pivotal -- Tom moves from convict to assigned servant in a household. The two pause points capture moments of unspoken authorial work: the connection to Rob, and the contrast between expected punishment and walking together as people.

WATCH FOR:
- Students who connect Sgt Stanley to Rob -- excellent chapter-to-chapter linking
- Students moved by Tom's freedom -- acknowledge the emotional shift

[General: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause here. The Major has just announced that Tom will be taken into custody by Sergeant Stanley
- "Tom followed him."
- What have we learned from this conversation? [Tom is no longer a convict in the eyes of the Major. Sgt Stanley is Rob's father -- the boy from the upper deck. Tom is being treated as someone to be trusted]
- What is going on here? [The author is showing us a moment of dignity. Tom does not run, does not resist. He follows. The relationship has shifted]

DO:
- Display the quote
- Think-Pair-Share: 30 seconds think, 30 seconds pair
- Push for the recognition that Sgt Stanley is Rob's father -- the connection back to Chapter 22

TEACHER NOTES:
This pause point develops the recognition theme and connects characters across chapters. The author has been setting up Sgt Stanley as Rob's father; the reveal links Tom and Rob through the household.

WATCH FOR:
- Students who connect Sgt Stanley to Rob -- celebrate
- Students who notice "Tom followed him" as a small but loaded line -- excellent close reading

[General: Pause Point 1 | VTLM 2.0: Higher-Order Questioning]`;

const NOTES_PAUSE2 = `SAY:
- Pause here. Tom and Sgt Stanley leave Sydney Cove. They could have walked anywhere
- "But instead they walked around the harbour."
- What does the author want us to know? [The Sergeant is treating Tom as a person, not as a prisoner. The walk is a small act of trust. They are sharing the new land together]
- What does this show about the Sergeant? [He is choosing kindness. He is acknowledging Tom as worth knowing. He is using his power to give Tom a soft landing into freedom]

DO:
- Display the quote
- Cold Call after thinking time
- Connect back: "How is this different from how Tom started the novel? Where was he? How was he treated?"

CFU CHECKPOINT:
Technique: Cold Call (after thinking time)

Script:
- "30 seconds to think. What does this line show us about the Sergeant and Tom?"
- Cold call 2-3 students
- Scan for: the Sergeant's choice of kindness; Tom's recognition as a person

PROCEED (>=80%): Most students see the Sergeant's choice and the change in Tom's status. Continue.
PIVOT (<80%): Most likely issue -- students miss the symbolic weight of the walk. Reteach: "What did the Sergeant HAVE to do? Take Tom to his quarters. What did the Sergeant CHOOSE to do? Walk around the harbour. The author is showing us a CHOICE." Re-check: "Why does that choice matter for Tom?"

TEACHER NOTES:
This pause point is about authorial choice in showing relationship. The detail of the harbour walk is small but the weight is large.

WATCH FOR:
- Students who name the change in Tom's status -- excellent
- Students who notice "instead" as a key word -- strong reading

[General: Pause Point 2 | VTLM 2.0: Deep Comprehension]`;

const NOTES_BRIDGE = `SAY:
- Now we shift from novel to history. Tom's story is set during the First Fleet
- Our information report has two body paragraphs so far. Body paragraph 1: life in 18th century England. Body paragraph 2: why the First Fleet came to Australia
- Why is this connected? [Tom only ended up here because the British government decided to send a fleet. If we understand WHY they came, we understand more about Tom's story]
- Today you read a non-fiction article about the First Fleet, then plan body paragraph 2 using an SPO

DO:
- Display the bridge slide
- Brief framing: "Novel + history. We have read the story. Now we plan the report"
- Have the non-fiction article ready -- printed or on screen

TEACHER NOTES:
This bridge slide does not need long teaching. It signals the shift from reading to writing and connects the two strands. Keep it under 2 minutes.

WATCH FOR:
- Students who immediately see the connection between novel and history -- celebrate
- Students unsure why they are reading non-fiction -- frame it as evidence for the report

[General: Bridge / Cross-Curricular | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_ARTICLE = `SAY:
- We are going to read or skim a non-fiction article about the First Fleet
- Your teacher has the article ready
- As you read, look for: WHY did the British government send the First Fleet? Use your KPAS skills to take notes
- Look for the key reasons: prisons, transportation, finding new colonies, claiming new land

DO:
- Distribute the non-fiction article (the State Library "First Fleet" sections, the National Museum "Why was a convict colony set up in Australia?" page, or another teacher-selected article)
- 8-10 minutes to read and take KPAS notes
- Circulate -- prompt students to focus on the WHY question

TEACHER NOTES:
The non-fiction reading is the source for the SPO. Suggested teacher-selected articles include the supplied State Library and National Museum links. Teachers may use other school-approved sources. Encourage KPAS notes -- this builds on Lesson 13's skill.

WATCH FOR:
- Students who copy whole sentences instead of taking KPAS notes -- redirect: "You have the symbols. Use them"
- Students who find one reason and stop -- prompt: "Find at least three reasons. The article has more than one"

[General: Reading + Note-Taking | VTLM 2.0: Active Reading]`;

const NOTES_IDO = `SAY:
- Now we plan body paragraph 2. Watch me build an SPO using the article
- My topic: why the First Fleet came to Australia
- Topic sentence: "The First Fleet came to Australia in 1788 to set up a new penal colony where Britain could send its convicted prisoners."
- Now my supporting details. From the article:
  - Detail 1: British prisons were dangerously overcrowded after the Industrial Revolution
  - Detail 2: Britain had previously sent convicts to America, but lost that option after American independence in 1783
  - Detail 3: The British government wanted to claim new land for trade and farming
- Concluding sentence: "These pressures pushed Britain to send the First Fleet halfway around the world to start the colony at Botany Bay."
- Notice: my SPO is built FROM the article. Each detail comes from the source

DO:
- Display the SPO template on screen
- Build it section by section with the article visible
- Show students where each detail comes from in the article
- Think aloud: "I check -- is each detail an answer to WHY they came? Yes. Yes. Yes"

TEACHER NOTES:
The I Do is the most important slide of the lesson. Students see how to lift information from a non-fiction source into an SPO. The think-aloud about checking that details answer the question is key.

WATCH FOR:
- Students who are already drafting in their head -- good
- Students confused about lifting from the article -- the mentor model on the worksheet provides another example

[General: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_YOUDO = `SAY:
- Your turn. On your worksheet, plan body paragraph 2
- Topic: why the First Fleet came to Australia
- Use the article and your KPAS notes
- Write a topic sentence, three supporting details, and a concluding sentence
- 12 minutes. I will circulate

DO:
- Distribute the SPO Plan Body Paragraph 2 worksheet
- For students needing scaffolding, the worksheet provides a starter topic sentence
- Circulate -- check that supporting details come from the article and answer the WHY question

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide the topic sentence pre-written. Students write 3 supporting details from the article and a concluding sentence. Sentence starters provided: "One reason was...", "Another reason was...", "These reasons together..."
- Extra Notes: These students can use a teacher-highlighted version of the article that points to the key reasons

EXTENDING PROMPT:
- Task: Plan TWO body paragraphs -- one on why the First Fleet came (today's topic) and one on what happened when they arrived (drawing on Chapters 27-30 plus the article). Both must use SPO format

TEACHER NOTES:
Students keep their completed SPO for use in the next writing lesson, where they will convert it into a full body paragraph (parallel to Lesson 10's process). Confer actively during the 12 minutes.

WATCH FOR:
- Students who write the article verbatim -- redirect: "Use KPAS, then turn it into your own sentences for the SPO"
- Students whose details do not answer WHY -- redirect: "Does this detail answer the question 'why did they come?'"
- Students who finish quickly -- direct to the extending task

[General: You Do | VTLM 2.0: Independent Application]`;

const NOTES_CLOSING = `SAY:
- Quick self-check against the success criteria. Thumbs for each
- Then turn to a partner: tell them ONE reason why the First Fleet came to Australia. Use your SPO

DO:
- Run thumbs check for each SC
- Listen in on partner shares -- diagnostic for the next writing lesson
- Wrap up: "Next writing lesson, you turn this plan into a full paragraph -- just like body paragraph 1"

TEACHER NOTES:
The closing connects today's planning to the next writing session. Collect SPOs for review before the next session.

WATCH FOR:
- Students "thumbs down" on SC3 -- they may need more SPO modelling
- Students sharing strong reasons with a partner -- their SPO is ready for the write

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${SPO_RESOURCE.name} is for planning body paragraph 2
- The ${ANSWER_KEY_RESOURCE.name} is teacher reference

DO:
- Print the SPO worksheet (one per student)
- Print the model answer (teacher copy only)
- Have the non-fiction article available -- the suggested links from the unit are:
  - State Library NSW: First Fleet (Floating Prisons, Setting Sail)
  - National Museum of Australia: Why was a convict colony set up in Australia?
  - Or another school-approved article on the First Fleet

TEACHER NOTES:
The non-fiction article is teacher-selected. The suggested links from the unit material are listed in the teacher notes. The SPO worksheet does not require any specific article -- students plan from whichever source the teacher provides.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Chapters 29-30 + First Fleet -- Lesson 14";

  // Slide 1 -- Title
  titleSlide(
    pres,
    "Chapters 29-30",
    "Free, and Why They Came",
    "Lesson 14  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  // Slide 2 -- LI/SC
  liSlide(
    pres,
    [
      "We are learning to notice the choices an author makes about a character's freedom, and to plan a body paragraph for an information report using a Single Paragraph Outline",
    ],
    [
      "I can explain what the author wants us to know about Tom's new situation",
      "I can describe the structure of a body paragraph (topic sentence, supporting details, concluding sentence)",
      "I can plan a body paragraph using an SPO drawn from a non-fiction article",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // Slide 3 -- Reading Launch
  contentSlide(
    pres,
    "Teacher Read Aloud",
    C.PRIMARY,
    "Chapters 29 - 30",
    [
      "Reading Mode: Teacher Read Aloud",
      "Chapter 29: Tom's case is heard. Sgt Stanley appears. Tom's status changes",
      "Chapter 30: Tom moves to a new household and walks around the harbour",
      "Focus: how does the author show us Tom's new freedom?",
    ],
    NOTES_READING,
    FOOTER
  );

  // Slide 4 -- Pause Point 1
  quoteSlide(
    pres,
    "Pause Point 1",
    "Chapter 29 -- p. 158",
    "Tom followed him.",
    "p. 158",
    "What have we learned from this conversation? Why is Sgt Stanley significant?",
    NOTES_PAUSE1,
    FOOTER
  );

  // Slide 5 -- Pause Point 2
  quoteSlide(
    pres,
    "Pause Point 2",
    "Chapter 30 -- p. 160",
    "But instead they walked around the harbour.",
    "p. 160",
    "What does the Sergeant's CHOICE show us? How is this different from how Tom started the novel?",
    NOTES_PAUSE2,
    FOOTER
  );

  // Slide 6 -- Bridge: novel to history
  contentSlide(
    pres,
    "Bridge",
    C.SECONDARY,
    "From Novel to History",
    [
      "Tom is in Australia because the British government sent the First Fleet",
      "Body paragraph 1 (Lesson 10): life in 18th century England",
      "Body paragraph 2 (today): why the First Fleet came to Australia",
      "Today: read the non-fiction article, then plan body paragraph 2",
    ],
    NOTES_BRIDGE,
    FOOTER
  );

  // Slide 7 -- Read the article
  contentSlide(
    pres,
    "Read",
    C.PRIMARY,
    "Read the Non-Fiction Article",
    [
      "Your teacher will give you the First Fleet article",
      "As you read, focus on ONE question: WHY did the First Fleet come to Australia?",
      "Use KPAS to take notes -- not full sentences",
      "Look for at least three reasons",
    ],
    NOTES_ARTICLE,
    FOOTER
  );

  // Slide 8 -- I Do: Model SPO
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Plan Body Paragraph 2 -- SPO",
    "My Topic:\nWhy the First Fleet came\nto Australia\n\nTopic sentence (TS):\n\"The First Fleet came\nto Australia in 1788\nto set up a new penal\ncolony where Britain\ncould send its convicted\nprisoners.\"",
    "Supporting details (from the article):\n1. British prisons were\n   dangerously overcrowded\n2. Britain lost America\n   in 1783 and could no\n   longer send convicts there\n3. The British wanted to\n   claim new land for trade\n   and farming\n\nConcluding sentence (CS):\n\"These pressures pushed\nBritain to send the Fleet\nhalfway around the world.\"",
    NOTES_IDO,
    FOOTER
  );

  // Slide 9 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Plan Your Body Paragraph 2");

    addInstructionCard(s, [
      { text: "On your SPO worksheet:", role: "header" },
      { text: "First: Read the topic on the worksheet" },
      { text: "Next: Use the article + your KPAS notes" },
      { text: "Then: Write a topic sentence, 3 details, and a concluding sentence" },
      { text: "Check: do my details ANSWER why they came?" },
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
      { text: "Topic sentence introduces the main idea -- the WHY", options: { breakLine: true } },
      { text: "Each detail must come from the article and answer the question", options: { breakLine: true } },
      { text: "Concluding sentence wraps up without repeating the TS" },
    ], {
      x: 0.75, y: tipY + 0.46, w: 8.4, h: 0.85,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bullet: true, margin: 0,
      paraSpaceAfter: 2,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // Slide 10 -- Closing
  closingSlide(
    pres,
    "Turn to a partner: using your SPO, tell them ONE reason why the First Fleet came to Australia.",
    [
      "I can explain what the author wants us to know about Tom's new situation",
      "I can describe the structure of a body paragraph (topic sentence, supporting details, concluding sentence)",
      "I can plan a body paragraph using an SPO drawn from a non-fiction article",
    ],
    NOTES_CLOSING
  );

  // Slide 11 -- Resources
  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // -----------------------------------------------------------------------
  // Companion PDFs
  // -----------------------------------------------------------------------

  // SPO Template
  const ws = createPdf({ title: SPO_RESOURCE.name });
  let wsY = addPdfHeader(ws, "SPO Plan -- Body Paragraph 2", {
    color: C.PRIMARY,
    subtitle: "Information Report: Why the First Fleet came to Australia",
    lessonInfo: "Lesson 14 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "Use this template to plan body paragraph 2 of your information report. Read the article first; use KPAS to take notes; then build your SPO.", wsY, { color: C.PRIMARY });

  wsY = addSectionHeading(ws, "Topic", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Why the First Fleet came to Australia", wsY, { fontSize: 12 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Topic Sentence (TS)", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Use the starter or write your own. Your TS should introduce the WHY.", wsY, { fontSize: 10, italic: true });
  wsY = addTipBox(ws, "Starter (optional): \"The First Fleet came to Australia in 1788 because...\"", wsY, { color: C.SECONDARY });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Supporting Details", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Find 3 reasons from the article. Notes are fine -- short phrases or KPAS.", wsY, { fontSize: 10, italic: true });
  wsY += 4;
  wsY = addBodyText(ws, "Reason 1:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;
  wsY = addBodyText(ws, "Reason 2:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;
  wsY = addBodyText(ws, "Reason 3:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Concluding Sentence (CS)", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Wrap up your paragraph. Do NOT repeat your TS word-for-word.", wsY, { fontSize: 10, italic: true });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 10;

  wsY = addSectionHeading(ws, "Self-Check", wsY, { color: C.ALERT });
  wsY = addBodyText(ws, "- Does my TS introduce the WHY?", wsY, { fontSize: 10 });
  wsY = addBodyText(ws, "- Does each reason come from the article?", wsY, { fontSize: 10 });
  wsY = addBodyText(ws, "- Does each reason answer the question 'why did they come?'", wsY, { fontSize: 10 });
  wsY = addBodyText(ws, "- Does my CS wrap up without repeating the TS?", wsY, { fontSize: 10 });

  addPdfFooter(ws, "Lesson 14 | SPO Plan -- Keep for Next Writing Lesson");

  // Answer key
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "SPO Model -- Body Paragraph 2", {
    color: C.ALERT,
    subtitle: "Teacher Reference -- Why the First Fleet came to Australia",
    lessonInfo: "Lesson 14 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Model SPO based on the suggested non-fiction articles. Student plans will vary; assess whether reasons come from the article and genuinely answer the WHY question.", akY, { color: C.ALERT });

  akY = addSectionHeading(ak, "Model SPO", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "TS: The First Fleet came to Australia in 1788 to set up a new penal colony where Britain could send its convicted prisoners.", akY);
  akY = addBodyText(ak, "Reason 1: British prisons were dangerously overcrowded after the Industrial Revolution -- so many people were unemployed and turned to petty theft", akY);
  akY = addBodyText(ak, "Reason 2: Britain lost America in 1783 and could no longer send convicts there -- they needed a new place to send them", akY);
  akY = addBodyText(ak, "Reason 3: The British wanted to claim new land for trade and farming, and to expand their empire", akY);
  akY = addBodyText(ak, "CS: These pressures pushed Britain to send the First Fleet halfway around the world to start the colony at Botany Bay.", akY);
  akY += 12;

  akY = addSectionHeading(ak, "Alternative Reasons (from the suggested articles)", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "- Floating prisons (hulks) on the Thames were full and unhygienic", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Petty theft was on the rise due to displacement and unemployment", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- The British government wanted to remove 'undesirables' from cities", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Botany Bay was reported by Captain Cook as suitable for settlement", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Strategic interest in a southern naval base for trade and defence", akY, { fontSize: 10 });
  akY += 12;

  akY = addSectionHeading(ak, "What to Look For", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- Topic sentence introduces the WHY (not just states a fact)", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Each reason is taken from the article (not from imagination)", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Each reason genuinely answers 'why did they come?'", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Concluding sentence wraps up without repeating the TS", akY, { fontSize: 10 });
  akY += 8;

  akY = addSectionHeading(ak, "Suggested Articles", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "State Library of New South Wales: First Fleet (Floating Prisons, Setting Sail)", akY, { fontSize: 10 });
  akY = addBodyText(ak, "National Museum of Australia: Why was a convict colony set up in Australia?", akY, { fontSize: 10 });
  akY = addBodyText(ak, "Or another school-approved article. Highlight a copy for enabling students.", akY, { fontSize: 10, italic: true });

  addPdfFooter(ak, "Lesson 14 | SPO Model -- TEACHER COPY");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Lesson14.pptx` }),
    writePdf(ws, SPO_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Lesson14.pptx`);
  console.log("Done: " + SPO_RESOURCE.name);
  console.log("Done: " + ANSWER_KEY_RESOURCE.name);
}

build().catch(console.error);
