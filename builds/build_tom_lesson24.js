"use strict";

// Tom Unit -- Lesson 24: Chapter 52 (FINAL) + Plan & Write Body Paragraph 4
// Week 6, Lesson 24, Grade 5/6 Literacy
// Reading: Ch 52 -- Thomas tells Millie about the courage cloak; the ghosts are his descendants
// Non-fiction: NMA First Fleet article -- impact on First Nations Australians
// Writing: Plan SPO + write body paragraph 4 of the information report
// Sensitivity: Content covers the impact of British settlement on First Nations Australians

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

const SESSION_NUMBER = 24;
const FOOTER = "Chapter 52 + Body Paragraph 4 | Lesson 24 | Week 6 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson24_Final_Chapter_Body_Paragraph_4";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const SPO_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "SPO and Body Paragraph 4 Template",
  "Student template: plan and write body paragraph 4 on the impact of British settlement on First Nations Australians."
);
const MENTOR_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mentor Paragraph Impact on First Nations",
  "Annotated model body paragraph showing structure, language features and respectful framing."
);
const RESOURCE_ITEMS = [SPO_RESOURCE, MENTOR_RESOURCE];
const SPO_PDF_PATH = path.join(OUT_DIR, SPO_RESOURCE.fileName);
const MENTOR_PDF_PATH = path.join(OUT_DIR, MENTOR_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Lesson 24. A big lesson today
- First: we finish the novel. Chapter 52 is the final chapter
- Then: we use our notes and our research to write body paragraph 4 of our information report
- The topic for body paragraph 4 is the impact of British settlement on First Nations Australians

DO:
- Display title slide as students settle
- Have copies of the novel ready
- Have non-fiction article ready (NMA First Fleet) -- or alternative the school has selected
- Have SPO/Paragraph 4 templates and the mentor paragraph ready (do not distribute yet)

TEACHER NOTES:
This is a milestone lesson. Mark the closing of the novel briefly so it lands -- but do not let nostalgia eat the writing time. The writing block is the most demanding part of the lesson; protect at least 25 minutes for plan + write.

SENSITIVITY ADVISORY:
- What it is: This paragraph covers the impact of British settlement on First Nations Australians, including loss of Country, devastating effects of disease, and forced change.
- Framing language: Use respectful, factual language. The Eora people had cared for their Country for tens of thousands of years before 1788. Settlement was devastating. The First Nations peoples of Australia are still here.
- Watch for: Students who use 'they were wiped out' style language -- redirect to 'devastated' and the fact that First Nations peoples are still here today.
- Protocol: If any student becomes distressed, check in privately. The class teacher knows which students may need extra support.

WATCH FOR:
- Students who want to dwell on the novel ending -- honour briefly, then move on
- Students who struggle to begin the paragraph -- send to the mentor paragraph and the SPO template

[Literacy: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${SPO_RESOURCE.name} is your planner and your writing page in one
- The ${MENTOR_RESOURCE.name} is the model -- annotated for structure and language features

DO:
- Print the SPO/Paragraph template (one per student)
- Print the mentor paragraph (one per student or per pair)
- Distribute the mentor as students enter the writing block

TEACHER NOTES:
The SPO template has the planning grid on page 1 and the writing space on page 2. The mentor paragraph is a careful, respectful model -- use it as the reference for structure and tone.

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_HOOK = `SAY:
- Before we read the final chapter, two quick questions
- Question 1 for mini-whiteboards: where does the novel BEGIN? What is the first big thing that happens to Tom?
- Question 2: what is the courage cloak?
- 60 seconds. Boards up

DO:
- 60 seconds whiteboard work
- Boards up
- Cold call 1-2 for each question
- Confirm: novel begins with Tom's mother dying / Tom being arrested as a thief; courage cloak is the invisible cloak his mother gave him on the day they parted, to remind him of her courage

TEACHER NOTES:
This launch activates the long arc of the novel. Tom started as a small boy in London who lost his mother. He has carried the cloak across the world. Chapter 52 closes that arc. Strong students will name the cloak as a motif that has appeared many times.

WATCH FOR:
- Students who recall the cloak from the very start -- celebrate the connection
- Students who forget the mother detail -- supply gently before the read

[Literacy: Hook | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention
- Two strands: finishing the novel and writing body paragraph 4 on a respectful topic
- Read the success criteria

DO:
- Choral read the LI, then the SCs
- Brief: "Today we draw on both fiction and non-fiction -- the novel for our understanding of Tom, the article for our facts on impact"

TEACHER NOTES:
SC1 is the writing target -- a full paragraph using the structure. SC2 is the language features (appositives, relative clauses, adverbials, third person). SC3 is the respectful framing. All three matter.

WATCH FOR:
- Students who think "I just need to write" -- redirect to SC2 and SC3 also

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_NOVEL = `SAY:
- One word from the novel: generous
- Generous means giving freely -- of time, money, food or kindness
- Thomas reflects on the abundance he has had in his life. He has been generous and others have been generous to him

DO:
- Display the word and meaning
- Read the meaning and example aloud
- Quick oral: "Give me an example of being generous. 20 seconds with your partner"

TEACHER NOTES:
"Generous" appears in the incidental list. Connect it to the closing reflection of the novel -- Tom looking back at his life and counting his blessings. Strong students may also use 'abundance' from the same list.

WATCH FOR:
- Students who give simple examples (sharing food, lending a pencil) -- accept and connect to the novel

[Literacy: Vocabulary | VTLM 2.0: Build Knowledge / Vocabulary]`;

const NOTES_READING = `SAY:
- Chapter 52. This is the last chapter of the book
- I will read aloud. Two pause points
- Listen for the moments where the author closes the story of Tom

DO:
- 30 seconds for students to find Chapter 52
- Read aloud at a calm, steady pace
- Take pause point 1 at p.276, pause point 2 at p.277 (final line of the novel)
- Acknowledge briefly when reading is finished: "We just finished the novel"

TEACHER NOTES:
The reading takes 8-12 minutes. The final chapter is reflective rather than eventful. Thomas (now elderly) tells Millie about the courage cloak. He learns the ghosts are not strangers -- they are his descendants, proud of him. The final image is Thomas looking out at his acres, content. Let the closing land for a moment before moving on.

WATCH FOR:
- Students who are moved by the ending -- briefly acknowledge
- Students who want to discuss it at length -- promise discussion in the closing reflection; protect the writing time

[Literacy: Read Aloud | VTLM 2.0: Build Knowledge / Reading]`;

const NOTES_PAUSE1 = `SAY:
- Stop. Thomas has just told Millie about the courage cloak. He says: "I'll take good care of it."
- Turn to your partner: what is the significance of Thomas telling Millie about the courage cloak? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: he is passing it on; the cloak now belongs to a new generation; it is not just his mother's gift any more -- it is a family heirloom; his story will be remembered; the courage will keep going

TEACHER NOTES:
This is a moment of legacy. The cloak that Tom's mother gave him as a frightened boy now passes to a child of his family. The author is showing us that what Tom carried has lasted across his whole life and into the next generation.

WATCH FOR:
- Students who just describe the action (he gave it to her) -- prompt: "What does the cloak STAND FOR? Why does it matter that he is passing it on?"
- Students who name 'legacy' or 'family' or 'memory' -- excellent

[Literacy: Pause Point | VTLM 2.0: Comprehension / Motif and Theme]`;

const NOTES_PAUSE2 = `SAY:
- Stop. The author's final line: "...gazed out at his acres, and was content."
- Two questions for your partner. 90 seconds
- How does what we just read add to our understanding of Tom?
- How is the author making you feel right now about Tom?

DO:
- 90 seconds Turn & Talk
- Cold Call 2-3 students for the first question, then 2-3 for the second
- Listen for the first: Tom has built a life; he owns land now; he is no longer a convict; the boy who once had nothing now has acres; he has come a long way
- Listen for the second: content; satisfied; happy for him; emotional; the ending feels right; the author has given Tom peace
- Briefly: "That is the last line of the book. Tom is at peace"

TEACHER NOTES:
This is the closing line of the novel. Let it land. The boy who began with nothing -- no family, no home, no freedom -- now owns acres and is content. The author chose this image deliberately. Strong students will name the arc.

WATCH FOR:
- Students who name the journey from convict to landowner -- excellent
- Students who feel sad it is over -- valid; honour briefly

[Literacy: Pause Point | VTLM 2.0: Comprehension / Author's Closing Choice]`;

const NOTES_PIVOT_TO_WRITING = `SAY:
- We just finished the novel. Take a breath
- Now we pivot to the writing for our information report
- Body paragraph 4 is about a topic the novel has touched on but not been the focus: the impact of British settlement on First Nations Australians
- The Eora people had cared for the land around Sydney Cove for tens of thousands of years before 1788
- The settlers' arrival changed that profoundly

DO:
- Read the slide aloud
- Pre-frame the topic: this is a respectful, factual paragraph
- "Our notes from the NMA article will be our source"

TEACHER NOTES:
This is the transition slide. Use it to settle the class after the novel ending and to set up the writing seriously. The framing here matters -- model respectful, factual language now so students hear it.

SENSITIVITY ADVISORY:
- Use respectful language at all times. Avoid 'they were wiped out' or similar.
- Acknowledge: First Nations peoples are still here.
- Acknowledge: the Eora people are the Traditional Custodians of the land Sydney is built on.

WATCH FOR:
- Students who lean into the novel ending and resist the pivot -- promise a closing reflection at the end

[Literacy: Transition | VTLM 2.0: Building Knowledge / Pivot to Writing]`;

const NOTES_NONFICTION = `SAY:
- Quick read of the non-fiction article -- NMA First Fleet, or the alternative your teacher chose
- As I read, jot KPAS notes on impact: what changed for the Eora and other First Nations peoples?
- Listen for: loss of Country; disease; conflict; displacement; the Eora people are still here

DO:
- Project or distribute the article (NMA digital classroom)
- Read aloud or have students read in pairs (5-7 minutes)
- Direct students to take KPAS notes on their template, page 1
- Circulate -- check students are capturing impact (not just dates)

TEACHER NOTES:
The NMA article (https://digital-classroom.nma.gov.au/defining-moments/first-fleet-arrives-sydney-cove) is the suggested text. If unavailable, the teacher may use an alternative article on the same topic. Note-taking here uses the KPAS skill from Lessons 22 and 23 -- this is the payoff for that procedural learning.

WATCH FOR:
- Students who copy full sentences -- redirect: "Use KPAS. Drop the small words"
- Students who note only dates -- redirect: "Notes on IMPACT. What changed for First Nations peoples?"

[Literacy: Read & Note | VTLM 2.0: Build Knowledge / Research]`;

const NOTES_REVISE = `SAY:
- Quick revision of body paragraph structure -- we used this in Lessons 14, 15 and 20
- Topic sentence: introduces the main idea -- when, where, who, what
- Supporting details: expand the topic sentence with facts
- Concluding sentence: adds the FINAL idea -- does not repeat the TS
- Language features for this paragraph: past tense, third person, at least ONE appositive or relative clause, specific vocabulary

DO:
- Display the structure card
- Choral read each section
- Quick oral check: "What is a topic sentence? What is a concluding sentence?"

TEACHER NOTES:
This is recall, not new teaching. Most students will know this structure. Keep it brisk -- the heart of the lesson is the writing block.

WATCH FOR:
- Students who confuse TS and CS -- supply gently
- Students who skip the language features -- redirect to SC2

[Literacy: Revise | VTLM 2.0: Retention and Recall]`;

const NOTES_IDO = `SAY:
- Watch me think about a topic sentence for body paragraph 4
- Topic: the impact of British settlement on First Nations Australians
- My TS draft: "When the British arrived in Sydney Cove in 1788, life changed profoundly for the Eora people and the other First Nations peoples of the surrounding lands."
- Notice: WHEN (1788), WHERE (Sydney Cove), WHO (Eora and other First Nations peoples), WHAT (life changed profoundly)
- Now I show you the full paragraph

DO:
- Display the I Do slide
- Talk through the topic sentence first
- Then read the full paragraph aloud
- Highlight the relative clause, appositive and time adverbial
- Highlight the respectful framing: "had cared for the land for tens of thousands of years", "are still here today"

TEACHER NOTES:
The I Do is the most important model in this lesson. Read the paragraph aloud at a deliberate pace. Pause to point out the language features and the respectful framing. Students should leave this slide knowing: a) what a good TS looks like, b) what the structure should feel like, c) how to write respectfully about a sensitive topic.

WATCH FOR:
- Students who try to memorise the model -- redirect: "Your paragraph will use YOUR notes -- different facts may appear"
- Students who notice the respectful framing -- celebrate

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Two possible topic sentences for body paragraph 4. Which is the stronger TS?
- A: "The British arrived in 1788."
- B: "When the British arrived in Sydney Cove in January 1788, life changed profoundly for the Eora people and the other First Nations peoples of the surrounding lands."
- Hold up A or B. Three, two, one -- show

DO:
- Display both
- Show Me Boards
- Scan: most students should choose B
- Cold Call 1-2 students: "Why B?"

TEACHER NOTES:
A is a bare fact -- it does not signal what the paragraph will discuss. B signals WHEN (January 1788), WHERE (Sydney Cove), WHO (the Eora and other First Nations peoples) and WHAT (life changed profoundly). B is the stronger opener and frames the topic respectfully.

WATCH FOR:
- Students who pick A because it is shorter -- redirect: "Predict the paragraph from A. Now predict from B. Which one sets up the WHOLE paragraph?"
- Students who pick B and articulate WHY -- ready to write

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Stronger TS: B
- B signals when (January 1788), where (Sydney Cove), who (the Eora and other First Nations peoples) and what (life changed profoundly)
- B also names the people respectfully and accurately
- Aim for B-style detail in your own TS

DO:
- Display the reveal banner
- Read B aloud
- Pivot if many picked A: "Predict the paragraph from A. Now from B."

CFU CHECKPOINT:
Technique: Show Me Boards (A or B)
Script:
- "Hold up A or B. Which is the stronger topic sentence?"
- Scan for: most students choose B (>=80%)
PROCEED (>=80%): Most chose B. Release to plan and write.
PIVOT (<80%): Most likely issue -- students think a TS is just the headline fact. Reteach: "A TS sets up the whole paragraph -- when, where, who, what." Re-check with a fresh TS pair if needed.

TEACHER NOTES:
After reveal, release students to write. Time is now the limiting factor -- protect at least 18-22 minutes for plan + write.

WATCH FOR:
- Students who self-correct toward a stronger TS -- celebrate

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Your turn. Use the SPO and Paragraph template
- First: plan your SPO on page 1 -- TS, 3 supporting details, CS
- Next: write your paragraph on page 2
- Then: use ONE appositive or relative clause somewhere in your paragraph
- Use respectful language. The Eora people are still here today
- 22 minutes total -- about 6 to plan, 14 to write, 2 to read back
- I will circulate

DO:
- Distribute the SPO template and the mentor paragraph
- Circulate -- prioritise students who looked unsure during the CFU
- Quick conferences: "Read me your TS. Does it set up the whole paragraph?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use a sentence-frame version of the mentor paragraph with key information left blank. Students fill the blanks using their KPAS notes from the article
- Suggested frame: "When the British arrived in ____ in ____, life changed for the ____ people and the other First Nations peoples of the surrounding lands. The Eora people had cared for this Country for ____. The arrival of the British brought ____. Many First Nations peoples experienced ____ as a result of ____. Despite the impact of settlement, First Nations peoples of Australia ____."
- Extra Notes: These students still cover all three SCs but with reduced cognitive load on sentence construction

EXTENDING PROMPT:
- Task: Add a second sentence that uses an adverbial of time, place or manner. Underline the adverbial
- Extra Notes: Push for specificity -- "for tens of thousands of years", "across the surrounding lands", "with devastating effects"

TEACHER NOTES:
The 22-minute writing block is the heart of the lesson. Active circulation is the formative assessment. Mentor paragraph remains on every desk as the reference.

WATCH FOR:
- Students using 'they' generically -- redirect: "The Eora people. The First Nations peoples"
- Students using 'wiped out' or similar -- redirect: "Use 'devastating' or 'profound impact'. The First Nations peoples are still here today"
- Students copying the mentor verbatim -- prompt: "What did YOUR notes from the article say? Use those facts"

[Literacy: You Do | VTLM 2.0: Supported Application]`;

const NOTES_CLOSING = `SAY:
- Quick self-check against the success criteria
- Then a moment to mark the end of the novel
- Turn to your partner: tell them ONE moment from the novel that has stuck with you. 60 seconds
- One session left in this unit -- the conclusion of our information report

DO:
- Run thumbs / fingers check for each SC
- Listen in on partner shares -- this is the moment to honour the novel ending
- Collect drafts for review -- not for grading, for next-day reteach planning

TEACHER NOTES:
This closing combines a writing self-check and a brief acknowledgement that the novel is finished. Honour both. The reflection moment matters -- students have read 52 chapters together. Collect drafts so you can scan for common weak spots and target them in the conclusion-writing lesson tomorrow.

WATCH FOR:
- Students who name a specific moment -- evidence of engagement with the novel
- Students who feel uncertain about their paragraph -- offer a short small-group time tomorrow

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Tom Appleby Ch 52 + Body Paragraph 4 -- Lesson 24";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Chapter 52 + Body Paragraph 4",
    "The Final Chapter & Writing About Impact",
    "Lesson 24  |  Week 6  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources
  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // SLIDE 3 -- Hook / Launch
  contentSlide(
    pres,
    "Launch",
    C.PRIMARY,
    "Remembering the Novel",
    [
      "Mini-whiteboards out",
      "1: Where does the novel BEGIN? What is the first big thing that happens to Tom?",
      "2: What is the courage cloak?",
      "60 seconds. Boards up when called",
    ],
    NOTES_HOOK,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to read the final chapter of the novel, then use our notes to plan and write body paragraph 4 of our information report on the impact of British settlement on First Nations Australians",
    ],
    [
      "I can plan and write a body paragraph with a topic sentence, supporting details and a concluding sentence",
      "I can use language features of an information report including at least one appositive or relative clause",
      "I can use respectful, accurate language when writing about First Nations peoples",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Vocabulary from novel
  vocabSlide(
    pres,
    "generous",
    "adjective",
    "Giving freely -- of time, money, food or kindness.",
    "Thomas looked back at his life and saw how generous others had been to him.",
    NOTES_VOCAB_NOVEL,
    FOOTER
  );

  // SLIDE 6 -- Reading anchor (Ch 52)
  contentSlide(
    pres,
    "Read",
    C.PRIMARY,
    "Chapter 52  -- The Final Chapter",
    [
      "I will read aloud",
      "Two pause points: p.276 (courage cloak) and p.277 (the final line of the novel)",
      "Listen for the moments where the author closes the story of Tom",
      "We finish the book today",
    ],
    NOTES_READING,
    FOOTER
  );

  // SLIDE 7 -- Pause point 1 (p.276)
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 52  |  Thomas passes on the cloak",
    "I'll take good care of it.",
    "p.276",
    "What is the significance of Thomas telling Millie about the courage cloak?",
    NOTES_PAUSE1,
    FOOTER
  );

  // SLIDE 8 -- Pause point 2 (p.277) -- final line of novel
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 52  |  The final line",
    "...gazed out at his acres, and was content.",
    "p.277",
    "How does what we just read add to our understanding of Tom? How is the author making you feel about Tom now?",
    NOTES_PAUSE2,
    FOOTER
  );

  // SLIDE 9 -- Pivot to writing
  contentSlide(
    pres,
    "Pivot",
    C.SECONDARY,
    "From Novel to Writing",
    [
      "We have just finished the novel",
      "Now we use our writing skills to honour a topic the novel has touched on but not been the focus",
      "Body paragraph 4: the impact of British settlement on First Nations Australians",
      "The Eora people had cared for this Country for tens of thousands of years before 1788",
      "Our notes from the article are our source for this paragraph",
    ],
    NOTES_PIVOT_TO_WRITING,
    FOOTER
  );

  // SLIDE 10 -- Read non-fiction & take notes
  contentSlide(
    pres,
    "Read & Note",
    C.SECONDARY,
    "Non-Fiction: Impact on First Nations Australians",
    [
      "Suggested article: National Museum of Australia -- 'First Fleet Arrives in Sydney Cove'",
      "URL: digital-classroom.nma.gov.au/defining-moments/first-fleet-arrives-sydney-cove",
      "Or use the alternative your teacher has selected",
      "Take KPAS notes on IMPACT: what changed for First Nations peoples?",
      "Listen for: loss of Country, disease, conflict, displacement -- and the fact First Nations peoples are still here today",
    ],
    NOTES_NONFICTION,
    FOOTER
  );

  // SLIDE 11 -- Revise body paragraph structure
  contentSlide(
    pres,
    "Revise",
    C.SECONDARY,
    "Body Paragraph -- Structure & Features",
    [
      "Topic sentence  ->  Supporting details  ->  Concluding sentence",
      "Past tense for historical events (arrived, lived, changed)",
      "Third person (the British, the Eora people, the First Nations peoples)",
      "At least ONE appositive OR relative clause",
      "Specific, respectful vocabulary -- Eora, Country, traditional, devastating, colonisation",
    ],
    NOTES_REVISE,
    FOOTER
  );

  // SLIDE 12 -- I Do: model paragraph
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Model: Impact on First Nations Australians",
    "My SPO plan:\n\nTS:\nWhen British arrived 1788\n-> profound change for\nFirst Nations peoples\n\nDetail 1:\nEora cared for Country\n-- tens of thousands of years\n\nDetail 2:\nLoss of Country -- settlers\ntook land for farming\n\nDetail 3:\nDisease devastating\n-- many lost lives\n\nCS:\nDespite the impact,\nFirst Nations peoples\nare still here today",
    "My finished paragraph:\n\n\"When the British arrived in Sydney Cove in January 1788, life changed profoundly for the Eora people and the other First Nations peoples of the surrounding lands. The Eora, who had cared for their Country for tens of thousands of years, suddenly found their land claimed by settlers for farming and building. The arrival of new diseases, including smallpox, was devastating and caused the deaths of many First Nations people. Conflict over land and resources grew across the colony as more settlers arrived. Despite the profound impact of British settlement, the First Nations peoples of Australia are still here today and their cultures continue.\"",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 13 + 14 -- CFU: stronger TS (reveal pair)
  function buildCfuBase() {
    const slide = pres.addSlide();
    addTopBar(slide, C.ALERT);
    addBadge(slide, "CFU", { color: C.ALERT });
    addTitle(slide, "Which Is the Stronger Topic Sentence?", { color: C.ALERT });

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
      x: 0.5, y: CONTENT_TOP, w: 2.8, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ALERT },
    });
    slide.addText("Show Me Boards: A or B", {
      x: 0.5, y: CONTENT_TOP, w: 2.8, h: 0.40,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const cardY = CONTENT_TOP + 0.55;
    const cardH = 1.00;
    addCard(slide, 0.5, cardY, 9, cardH, { strip: C.MUTED, fill: C.WHITE });
    slide.addText("A", {
      x: 0.7, y: cardY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    slide.addText("\"The British arrived in 1788.\"", {
      x: 1.2, y: cardY + 0.18, w: 8.0, h: cardH - 0.30,
      fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const cardBY = cardY + cardH + 0.18;
    const cardBH = SAFE_BOTTOM - cardBY - 0.55;
    addCard(slide, 0.5, cardBY, 9, cardBH, { strip: C.SECONDARY, fill: C.BG_CARD });
    slide.addText("B", {
      x: 0.7, y: cardBY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    slide.addText("\"When the British arrived in Sydney Cove in January 1788, life changed profoundly for the Eora people and the other First Nations peoples of the surrounding lands.\"", {
      x: 1.2, y: cardBY + 0.14, w: 8.0, h: cardBH - 0.24,
      fontSize: 13, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
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
      slide.addText("Stronger TS: B  --  signals when, where, who and what; names the peoples respectfully", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 15 -- You Do: Plan + Write
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Plan & Write Body Paragraph 4");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.95, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Topic: The impact of British settlement on First Nations Australians", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:    Plan your SPO -- TS, 3 supporting details, CS\nNext:     Write your TS -- B-style detail (when, where, who, what)\nThen:     Expand each detail. Use ONE appositive or relative clause\nFinally:  Write your CS -- the FINAL idea (acknowledge First Nations peoples are still here today)", {
      x: 0.75, y: CONTENT_TOP + 0.46, w: 8.4, h: 1.40,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 2.10;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Writing Time: about 22 minutes", {
      x: 0.75, y: tipY + 0.10, w: 5, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Use your KPAS notes and the mentor paragraph on your desk\n- Respectful language: 'Eora people', 'First Nations peoples', 'devastating', 'profound impact'\n- Avoid: 'wiped out' or similar -- First Nations peoples are still here today", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: 0.85,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 16 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Tell your partner ONE moment from the novel that has stuck with you.",
      scItems: [
        "I can plan and write a body paragraph with TS, supporting details and CS",
        "I can use language features of an information report including at least one appositive or relative clause",
        "I can use respectful, accurate language when writing about First Nations peoples",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (got it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: SPO + Paragraph Template -------------------------------------
  const sp = createPdf({ title: SPO_RESOURCE.name });
  let spY = addPdfHeader(sp, "SPO Plan & Body Paragraph 4", {
    color: C.PRIMARY,
    subtitle: "Topic: The impact of British settlement on First Nations Australians",
    lessonInfo: "Lesson 24 | Week 6 | Year 5/6 Literacy",
    showNameDate: true,
  });

  spY = addTipBox(sp, "Use your KPAS notes from the article. Plan your SPO on this page. Then write your paragraph on the next page. Use respectful, accurate language. First Nations peoples are still here today.", spY, { color: C.PRIMARY });

  spY = addSectionHeading(sp, "1. Topic Sentence (TS)", spY, { color: C.PRIMARY });
  spY = addBodyText(sp, "Hint: signal WHEN (1788), WHERE (Sydney Cove), WHO (the Eora and other First Nations peoples) and WHAT (life changed profoundly).", spY, { fontSize: 10, italic: true });
  spY = addLinedArea(sp, spY, 2, { lineSpacing: 22 });
  spY += 4;

  spY = addSectionHeading(sp, "2. Supporting Detail 1", spY, { color: C.PRIMARY });
  spY = addBodyText(sp, "Hint: connection to Country. How long had First Nations peoples cared for this land?", spY, { fontSize: 10, italic: true });
  spY = addLinedArea(sp, spY, 2, { lineSpacing: 22 });
  spY += 4;

  spY = addSectionHeading(sp, "3. Supporting Detail 2", spY, { color: C.PRIMARY });
  spY = addBodyText(sp, "Hint: what was the impact of the settlers' arrival? (land, farming, building)", spY, { fontSize: 10, italic: true });
  spY = addLinedArea(sp, spY, 2, { lineSpacing: 22 });
  spY += 4;

  spY = addSectionHeading(sp, "4. Supporting Detail 3", spY, { color: C.PRIMARY });
  spY = addBodyText(sp, "Hint: disease or conflict. What were the devastating effects?", spY, { fontSize: 10, italic: true });
  spY = addLinedArea(sp, spY, 2, { lineSpacing: 22 });
  spY += 4;

  spY = addSectionHeading(sp, "5. Concluding Sentence (CS)", spY, { color: C.PRIMARY });
  spY = addBodyText(sp, "Hint: acknowledge that First Nations peoples of Australia are still here today and their cultures continue.", spY, { fontSize: 10, italic: true });
  spY = addLinedArea(sp, spY, 2, { lineSpacing: 22 });

  addPdfFooter(sp, "Lesson 24 | SPO Plan -- Page 1");

  // Page 2: writing space
  sp.addPage();
  let spY2 = addPdfHeader(sp, "Body Paragraph 4 -- Write Here", {
    color: C.PRIMARY,
    subtitle: "Use your SPO plan from page 1",
    lessonInfo: "Lesson 24 | Week 6 | Year 5/6 Literacy",
    showNameDate: false,
  });

  spY2 = addTipBox(sp, "Aim for ONE appositive or relative clause somewhere in your paragraph. Use past tense and third person. Underline the appositive or relative clause when you finish.", spY2, { color: C.SECONDARY });

  spY2 = addLinedArea(sp, spY2, 18, { lineSpacing: 22 });

  addPdfFooter(sp, "Lesson 24 | Body Paragraph 4 -- Page 2");

  // ---- PDF: Mentor Paragraph ---------------------------------------------
  const mp = createPdf({ title: MENTOR_RESOURCE.name });
  let mpY = addPdfHeader(mp, "Mentor Body Paragraph -- Annotated", {
    color: C.PRIMARY,
    subtitle: "Topic: Impact of British settlement on First Nations Australians",
    lessonInfo: "Lesson 24 | Week 6 | Year 5/6 Literacy",
    showNameDate: false,
  });

  mpY = addTipBox(mp, "This is a model body paragraph showing the structure, language features and respectful framing. Use it as a reference -- do not copy it. Your paragraph will use YOUR notes from the article.", mpY, { color: C.PRIMARY });

  mpY = addSectionHeading(mp, "Model Body Paragraph", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "When the British arrived in Sydney Cove in January 1788, life changed profoundly for the Eora people and the other First Nations peoples of the surrounding lands. The Eora, who had cared for their Country for tens of thousands of years, suddenly found their land claimed by settlers for farming and building. The arrival of new diseases, including smallpox, was devastating and caused the deaths of many First Nations people. Conflict over land and resources grew across the colony as more settlers arrived. Despite the profound impact of British settlement, the First Nations peoples of Australia are still here today and their cultures continue.", mpY, { fontSize: 12 });
  mpY += 14;

  mpY = addSectionHeading(mp, "Annotations -- Structure", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "Topic Sentence: signals WHEN (January 1788), WHERE (Sydney Cove), WHO (the Eora and other First Nations peoples) and WHAT (life changed profoundly).", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Supporting Detail 1: The Eora cared for Country for tens of thousands of years -- their relationship to the land predates settlement.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Supporting Detail 2: Loss of land -- settlers claimed it for farming and building.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Supporting Detail 3: Disease, including smallpox, was devastating.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Supporting Detail 4 (optional): Conflict grew as more settlers arrived.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Concluding Sentence: First Nations peoples of Australia are still here today and their cultures continue -- the FINAL idea, not a repeat of the TS.", mpY, { fontSize: 10 });
  mpY += 10;

  mpY = addSectionHeading(mp, "Annotations -- Language Features", mpY, { color: C.SECONDARY });
  mpY = addBodyText(mp, "Relative clause: \"who had cared for their Country for tens of thousands of years\" -- adds detail about the Eora.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Appositive: \"including smallpox\" -- renames \"new diseases\" with a specific example.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Past tense: arrived, changed, had cared, found, was, caused, grew.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Third person: \"the British\", \"the Eora\", \"First Nations peoples\", \"settlers\".", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Adverbial of time: \"in January 1788\", \"for tens of thousands of years\".", mpY, { fontSize: 10 });
  mpY += 10;

  mpY = addSectionHeading(mp, "Annotations -- Respectful Framing", mpY, { color: C.ALERT });
  mpY = addBodyText(mp, "Names the peoples respectfully and accurately: 'the Eora people', 'the First Nations peoples'.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Acknowledges the deep relationship to Country: 'tens of thousands of years'.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Uses accurate impact language: 'devastating', 'profound impact' -- avoids 'wiped out' or similar.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Closes by acknowledging continuity: 'are still here today and their cultures continue'.", mpY, { fontSize: 10 });

  addPdfFooter(mp, "Lesson 24 | Mentor Paragraph -- TEACHER AND STUDENT REFERENCE");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Lesson24.pptx` }),
    writePdf(sp, SPO_PDF_PATH),
    writePdf(mp, MENTOR_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Lesson24.pptx`);
  console.log("Done: " + SPO_RESOURCE.name);
  console.log("Done: " + MENTOR_RESOURCE.name);
}

build().catch(console.error);
