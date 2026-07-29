"use strict";

// Tom Unit -- Lesson 19: Chapters 40-41 -- Disease & Settlement + Plan Body Paragraph 3
// Week 4, Lesson 19, Grade 5/6 Literacy
// Reading: Ch 40 (Tom finds a dead Eora boy), Ch 41 (settlement struggles)
// Writing: Plan body paragraph 3 (SPO) -- Arrival in Sydney Cove and Early Settlement
// Sensitivity: Ch 40-41 contain content about widespread death of Aboriginal peoples

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
  quoteSlide, modellingSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 19;
const FOOTER = "Chapters 40-41 + Writing | Lesson 19 | Week 4 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson19_Disease_and_Settlement_SPO";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const SPO_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "SPO Planning Template",
  "Student template: plan body paragraph 3 (Arrival in Sydney Cove) using a single paragraph outline."
);
const MODEL_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "SPO Model Answer",
  "Teacher reference: model SPO plan for body paragraph 3 with alternative supporting details."
);
const RESOURCE_ITEMS = [SPO_RESOURCE, MODEL_RESOURCE];
const SPO_PDF_PATH = path.join(OUT_DIR, SPO_RESOURCE.fileName);
const MODEL_PDF_PATH = path.join(OUT_DIR, MODEL_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Lesson 19. Chapters 40 and 41 today
- These chapters describe a hard time -- disease has spread and many Eora people are dying. We will read with care
- After reading, we shift to writing. Today we plan body paragraph 3 of our information report. Topic: Arrival in Sydney Cove

DO:
- Display title slide as students settle
- Have copies of the novel ready
- Have the SPO planning template ready (do not distribute yet)
- Have the supplementary non-fiction texts ready

TEACHER NOTES:
This is a content-dense and emotionally heavy lesson. Chapters 40-41 deal with the smallpox outbreak and the deaths of many Aboriginal peoples. The user-supplied unit plan flags this as sensitive content. Please honour that. The writing strand mirrors lesson 9's SPO planning, scaled for body paragraph 3 (Arrival in Sydney Cove and Early Settlement).

SENSITIVITY ADVISORY:
- What it is: Chapters 40-41 describe widespread death of Aboriginal peoples from a disease (presumed smallpox) following British arrival. The text includes references to bodies and sores.
- Framing language: "Today the chapters are serious. Many Aboriginal people died from a sickness after the British arrived. We will read carefully and respectfully. If you need a moment, let me know."
- Watch for: Aboriginal and Torres Strait Islander students -- be aware of cultural sensitivity around death and naming. Check in privately before the lesson.
- Protocol: Pause if needed. You may summarise heavy passages rather than read aloud. Acknowledge the harm of colonisation in your own framing. Do not skip the content -- the impact on Eora people is part of this history.

WATCH FOR:
- Any student showing distress
- Aboriginal and Torres Strait Islander students -- offer to meet with their family, school AEW or wellbeing team if needed

[Literacy: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands: reading the chapters, then planning the next body paragraph of our information report
- Read the success criteria

DO:
- Choral read the LI, then the SCs
- Brief: "We are over halfway through our information report. Today we plan body paragraph 3"

TEACHER NOTES:
SC1 targets reading comprehension and respectful engagement with hard content. SC2 targets the structural understanding (TS + supporting details + CS). SC3 targets the planning application -- creating an SPO for body paragraph 3.

WATCH FOR:
- Students unsure about which body paragraph this is -- briefly orient: "Body 1 was 18th century England. Body 2 was Why the First Fleet came. Body 3 is Arrival in Sydney Cove"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_READING = `SAY:
- Two chapters today. Both deal with disease in the colony
- A note before we read. After the British arrived, many Aboriginal people died from a sickness. The colony assumed it was smallpox. We do not know exactly how it spread. The chapters describe this. We will read with care
- Two pause points

DO:
- 30 seconds for students to find Chapter 40
- Read aloud at a steady, respectful pace
- Consider summarising the most graphic descriptions rather than reading every detail
- Take pause point 1 at p.214, pause point 2 at p.216
- Then read Chapter 41

TEACHER NOTES:
The reading takes 15-20 minutes. Acknowledge in your own words that this is a difficult chapter in Australian history. Do not skip the content -- the impact on Eora people is part of the truth. Use respectful framing throughout.

WATCH FOR:
- Aboriginal and Torres Strait Islander students -- offer space, do not single them out for response
- All students reacting -- pause and acknowledge

[Literacy: Read Aloud | VTLM 2.0: Build Knowledge / Reading]`;

const NOTES_PAUSE1 = `SAY:
- Stop. Tom is looking at the dead Eora boy on the shore. He thinks about how he might be remembered -- and the chapter mentions the empty huts of the people who once lived along the harbour
- Turn to your partner: What is this all about? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: many Eora people have died; the disease is widespread; Tom is realising the colony is built on land where many people lived; the British presence is connected to the deaths; this is a tragedy

TEACHER NOTES:
This is a significant moment. Strong students will recognise that the empty huts and lost people are a direct consequence of British arrival. Frame this with care -- this is the historical reality of colonisation. Some students may struggle with the weight of this; that is appropriate.

WATCH FOR:
- Students naming the connection between British arrival and the deaths -- celebrate respectfully
- Students who only describe what Tom sees -- prompt: "Why is the author showing us the empty huts?"

[Literacy: Pause Point | VTLM 2.0: Comprehension / Author Craft]`;

const NOTES_PAUSE2 = `SAY:
- Stop. The boys keep watching over the sheep and the precious hens. Despite all the chaos, life in the cove continues
- Turn to your partner: What does the author want us to know here? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: life goes on alongside terrible events; the boys are still doing their everyday work; the colony is fragile; small acts of caring matter; the contrast between everyday work and disaster outside

TEACHER NOTES:
The juxtaposition is the key. The British colony is doing everyday tasks (tending the garden, watching the animals) while around them people are dying. The author sets these side by side. Strong students will notice this contrast.

WATCH FOR:
- Students who name the contrast -- celebrate the literary thinking
- Students who only describe the boys' actions -- prompt: "What is happening AROUND them at the same time?"

[Literacy: Pause Point | VTLM 2.0: Comprehension / Big Idea]`;

const NOTES_BACKGROUND = `SAY:
- We now move to writing. Body paragraph 3. Topic: Arrival in Sydney Cove and Early Settlement
- Before we plan, we read. Knowing the facts is what makes a strong information report
- Listen as the article is read aloud (or read in pairs). Take notes
- Listen for: when did the First Fleet arrive? What did they find? What was life like for convicts? What happened to the Eora people?

DO:
- Display or distribute the supplementary article (NMA "First Fleet Arrives in Sydney Cove" snapshot, or MHNSW "What was life in early Sydney like for convicts?")
- Read aloud or read in pairs (8-10 minutes)
- Pause briefly to ask: "What is one thing you have learned?"
- Have students take point-form notes
- After reading, give 2 minutes for partners to share top three facts

TEACHER NOTES:
The non-fiction text builds shared background knowledge. The user-supplied article suggestions are placeholders -- teachers may select alternative reading materials. Frame the reading respectfully -- the early settlement story has multiple perspectives, and the impact on the Eora people is part of the truth.

WATCH FOR:
- Students copying full sentences -- prompt: "Notes are key words"
- Students with no notes -- pair with a peer or scaffold orally

[Literacy: Build Background | VTLM 2.0: Build Knowledge]`;

const NOTES_REVISE = `SAY:
- Quick reminder before we plan
- A body paragraph has THREE parts: topic sentence, supporting details, concluding sentence
- Topic sentence introduces the main idea -- it tells the reader what the paragraph is about
- Supporting details give the facts that expand the topic sentence
- Concluding sentence wraps up without repeating
- For our paragraph, topic sentence should answer: WHAT was the arrival in Sydney Cove like?

DO:
- Display the structure
- Reference last lesson's mentor paragraph if visible
- Keep this brief: 2 minutes maximum

TEACHER NOTES:
Students have done this twice before. The reminder is intentionally short. The new layer is the new topic (Arrival in Sydney Cove) and the planning -- not the writing today.

WATCH FOR:
- Students who jump to writing -- redirect: "Today is planning. Tomorrow we write"

[Literacy: Revise | VTLM 2.0: Retention and Recall]`;

const NOTES_IDO = `SAY:
- Watch how I plan body paragraph 3 using my notes
- TS: "When the First Fleet arrived in Sydney Cove in January 1788, the British settlers faced enormous challenges in setting up a colony on Eora land."
- This TS tells the reader WHEN, WHERE and WHAT (the challenges). It also acknowledges Eora land
- Supporting detail 1: The British arrived on Eora Country -- this had been Aboriginal land for tens of thousands of years
- Supporting detail 2: The land was difficult for British farming -- soil was thin, wheat did not grow well near the cove
- Supporting detail 3: Convicts and marines struggled with food shortages, poor housing and disease
- CS: "These early years were marked by hardship for the British and devastation for the Eora people."
- Notice my CS does not just repeat the TS. It adds the FINAL idea -- the dual impact

DO:
- Display the SPO being built section by section
- Talk through each choice
- Show the model SPO planning template (resource 2)
- Pause: "Notice the topic sentence does NOT say everything. It just sets up what is coming"

TEACHER NOTES:
The I Do explicitly demonstrates SPO planning for body paragraph 3. The TS is structured as a "when/where/what" sentence. The supporting details are three angles: Eora Country, the land, the British struggle. The CS names the dual impact. This approach acknowledges Aboriginal history respectfully while honouring the historical content.

WATCH FOR:
- Students drafting their own ideas already
- Students confused by the dual impact in the CS -- this is a depth criterion; supported students may simplify

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Two topic sentences. Which one is stronger for body paragraph 3 about Arrival in Sydney Cove?
- A: "The First Fleet arrived in Sydney Cove."
- B: "When the First Fleet arrived in Sydney Cove in January 1788, the British settlers faced enormous challenges in setting up a colony on Eora land."
- Hold up A or B
- Three, two, one -- show!

DO:
- Display both
- Show Me Boards
- Scan: most students should choose B
- Cold Call 1-2 students: "Why B?"

TEACHER NOTES:
B sets up the paragraph by signalling WHEN, WHERE, WHO and WHAT (the challenges). It also acknowledges Eora land. A is just a fact -- it does not set up a paragraph. Students who pick A think simpler = better. The reteach: a TS sets up the WHOLE paragraph.

WATCH FOR:
- Students who pick A -- redirect: "What does B tell you that A does not?"
- Students who articulate WHY B -- they are ready to plan

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Stronger TS: B
- B tells the reader WHEN (January 1788), WHERE (Sydney Cove), WHO (British settlers) and WHAT (challenges) -- and acknowledges Eora land
- A is just a fact. It does not set up a paragraph
- When you write your own TS, aim for B-style detail

DO:
- Display the reveal
- Read B aloud
- Pivot if many missed it: "What does B let you predict about the rest of the paragraph? What does A let you predict?"

CFU CHECKPOINT:
Technique: Show Me Boards (A or B)
Script:
- "Hold up A or B. Which is the stronger topic sentence?"
- Scan for: most students choose B (~80%)
PROCEED (>=80%): Most chose B. Release to plan.
PIVOT (<80%): Most likely issue -- students think a TS just states a fact. Reteach: "A topic sentence sets up the WHOLE paragraph. It tells the reader what is coming." Re-check by asking students to predict what the paragraph will cover from B alone.

TEACHER NOTES:
After reveal, release students to plan their SPO. The B-style TS structure (when, where, who, what) is the takeaway.

WATCH FOR:
- Students who self-correct toward a stronger TS -- celebrate

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Time to plan. Use the SPO Planning Template
- First: write your topic sentence. Aim for B-style detail -- when, where, who and what
- Next: list 3 supporting details from your notes. Use point form
- Then: draft your concluding sentence -- not a copy of the TS
- 12 minutes. We will write the full paragraph in our next session

DO:
- Distribute the SPO Planning Template
- Circulate -- prioritise enabling students first
- Quick conferences: "Read me your TS. Does it set up the whole paragraph?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide a TS and CS already written. Students fill in only the 3 supporting details from their notes
- Suggested TS: "When the First Fleet arrived in Sydney Cove in 1788, the British settlers and the Eora people both faced major changes."
- Suggested CS: "The early years of the colony were difficult for everyone living near Sydney Cove."
- Extra Notes: These students focus on selecting and recording supporting details -- the most cognitively demanding task at SC2 level

EXTENDING PROMPT:
- Task: After completing the SPO, write the OPENING sentence of body paragraph 3 in full (not just the TS notes). The sentence must use a relative clause OR an appositive
- Extra Notes: Peer-share the opening sentence

TEACHER NOTES:
The 12-minute planning block is the heart of this lesson. Active circulation is the formative assessment. Students collect their SPOs to use in the next session, when they will write the full body paragraph.

WATCH FOR:
- Students who copy the model's TS verbatim -- prompt: "What is YOUR angle on this paragraph?"
- Students who write the full paragraph instead of an SPO -- redirect: "Today is planning. Just point form for the details"
- Students whose CS just repeats the TS -- prompt: "How does the paragraph END? What is the final idea?"

[Literacy: You Do -- Plan | VTLM 2.0: Supported Application / Planning]`;

const NOTES_CLOSING = `SAY:
- Quick self-check. Show me on your fingers
- SC1: I described what we learned about Tom and the colony from Chapters 40-41 -- 1 to 5
- SC2: I understand the structure of a body paragraph -- 1 to 5
- SC3: I planned my SPO for body paragraph 3 -- 1 to 5
- One thing you want to research more before writing -- tell your partner

DO:
- Run each SC
- Note any patterns -- if many score low on SC3, plan a check-in next session
- Collect SPOs for review -- they will be used in the next session
- Wrap: "Next session we write body paragraph 3. Bring your SPO"

TEACHER NOTES:
Collecting SPOs is essential -- they are the input for the next writing session. The reflection prompt invites students to identify their own gaps without judgement.

WATCH FOR:
- Students naming specific research gaps -- evidence of metacognition
- Students who feel unsure -- gentle reassurance: "Your SPO is the seed. We grow it next session"

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${SPO_RESOURCE.name} is your planning sheet
- The ${MODEL_RESOURCE.name} is for teacher reference -- can be shared with students who need a model

DO:
- Print the planning template (one per student)
- Print one model answer for teacher use
- Some teachers may share the model with students working at the enabling level

TEACHER NOTES:
The model SPO can be displayed during the I Do or shared as a reference during planning. Some teachers prefer to keep it on the teacher's desk and use only when a student is genuinely stuck.

[Literacy: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Tom Appleby Ch 40-41 + Plan Body Paragraph 3 -- Lesson 19";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Chapters 40-41",
    "Disease, Settlement + Plan Body Paragraph 3",
    "Lesson 19  |  Week 4  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // SLIDE 2 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to read Chapters 40-41 with care, and to plan body paragraph 3 of our information report about arrival in Sydney Cove using a single paragraph outline (SPO)",
    ],
    [
      "I can describe respectfully what we learned from Chapters 40 and 41",
      "I can explain the structure of a body paragraph (TS, supporting details, CS)",
      "I can plan body paragraph 3 using an SPO drawn from non-fiction reading",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 3 -- Reading anchor
  contentSlide(
    pres,
    "Read",
    C.PRIMARY,
    "Chapters 40 & 41",
    [
      "Ch 40: Tom finds a dead Eora boy. The Sergeant explains a sickness has spread through the colony",
      "Ch 41: Disease continues. The colony struggles -- food is short, soil is thin, supply ship is late",
      "Note: these chapters describe the deaths of many Aboriginal people. We read with care",
      "Two pause points along the way",
      "Find Chapter 40 in your novel",
    ],
    NOTES_READING,
    FOOTER
  );

  // SLIDE 4 -- Pause point 1 (p.214)
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 40  |  Tom looks at the harbour",
    "...and huts turning back to soil?",
    "p.214",
    "What is this all about?",
    NOTES_PAUSE1,
    FOOTER
  );

  // SLIDE 5 -- Pause point 2 (p.216)
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 41  |  The boys keep watch",
    "...the sheep and precious hens, and kept watch.",
    "p.216",
    "What do you think the author wants us to know?",
    NOTES_PAUSE2,
    FOOTER
  );

  // SLIDE 6 -- Pivot to writing: Build Background
  contentSlide(
    pres,
    "Read & Note",
    C.SECONDARY,
    "Build Background: Arrival in Sydney Cove",
    [
      "Listen as the non-fiction article is read aloud (or read in pairs)",
      "Take notes in point form -- key words and short phrases",
      "Listen for: when did the First Fleet arrive? What did they find? What was life like for convicts? What happened to the Eora people?",
      "Sources: NMA \"First Fleet Arrives in Sydney Cove\" or MHNSW \"What was life in early Sydney like for convicts?\"",
      "Update your notes -- you will need them in the next step",
    ],
    NOTES_BACKGROUND,
    FOOTER
  );

  // SLIDE 7 -- Brief revise: body paragraph structure
  contentSlide(
    pres,
    "Revise",
    C.SECONDARY,
    "Body Paragraph -- Quick Reminder",
    [
      "Topic sentence:  introduces the main idea -- tells the reader what the paragraph is about",
      "Supporting details:  the facts that expand the topic sentence",
      "Concluding sentence:  wraps up the paragraph without repeating the TS",
      "Today is PLANNING. We use an SPO -- single paragraph outline -- before we write",
    ],
    NOTES_REVISE,
    FOOTER
  );

  // SLIDE 8 -- I Do: model SPO
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Plan Body Paragraph 3 -- Arrival in Sydney Cove",
    "TS:\n\"When the First Fleet arrived in Sydney Cove in January 1788, the British settlers faced enormous challenges in setting up a colony on Eora land.\"\n\nSupporting details (point form):\n- Arrived on Eora Country -- Aboriginal land for tens of thousands of years\n- Land difficult for farming -- thin soil, wheat did not grow well\n- Convicts and marines struggled -- food shortages, poor housing, disease",
    "CS:\n\"These early years were marked by hardship for the British and devastation for the Eora people.\"\n\nNotice:\n- TS signals when, where, who and what\n- 3 supporting details, each from a different angle\n- CS adds the final idea (dual impact) without repeating the TS\n- This is a PLAN. Tomorrow we will write the full sentences",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 9 + 10 -- CFU: stronger topic sentence (reveal pair)
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
    slide.addText("✓  CHECK", {
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
    const cardH = 1.10;
    addCard(slide, 0.5, cardY, 9, cardH, { strip: C.MUTED, fill: C.WHITE });
    slide.addText("A", {
      x: 0.7, y: cardY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    slide.addText("\"The First Fleet arrived in Sydney Cove.\"", {
      x: 1.2, y: cardY + 0.18, w: 8.0, h: cardH - 0.30,
      fontSize: 17, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const cardBY = cardY + cardH + 0.18;
    const cardBH = SAFE_BOTTOM - cardBY - 0.55;
    addCard(slide, 0.5, cardBY, 9, cardBH, { strip: C.SECONDARY, fill: C.BG_CARD });
    slide.addText("B", {
      x: 0.7, y: cardBY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    slide.addText("\"When the First Fleet arrived in Sydney Cove in January 1788, the British settlers faced enormous challenges in setting up a colony on Eora land.\"", {
      x: 1.2, y: cardBY + 0.14, w: 8.0, h: cardBH - 0.24,
      fontSize: 14, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
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
      slide.addText("Stronger TS: B  --  signals when, where, who and what", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 11 -- You Do: Plan SPO
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Plan Your SPO -- Body Paragraph 3");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.85, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Topic:  Arrival in Sydney Cove and Early Settlement", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   Write your TS -- aim for B-style detail (when, where, who, what)\nNext:   List 3 supporting details from your notes (point form)\nThen:   Draft your CS -- not a copy of the TS\nFinish:  Read your SPO aloud -- does it tell a clear story?", {
      x: 0.75, y: CONTENT_TOP + 0.46, w: 8.4, h: 1.30,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 2.00;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Time:  12 minutes", {
      x: 0.75, y: tipY + 0.10, w: 4, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Today is PLANNING. We write the full paragraph next session\n- Use point form for the details -- not full sentences yet\n- Acknowledge Eora land in your TS or details where relevant", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: 0.85,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 12 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Tell your partner one thing you want to research more before writing.",
      scItems: [
        "I can describe respectfully what we learned from Chapters 40 and 41",
        "I can explain the structure of a body paragraph",
        "I can plan body paragraph 3 using an SPO drawn from non-fiction reading",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (got it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // SLIDE 13 -- Resources


  // ---- PDF: SPO Planning Template -----------------------------------------
  const ws = createPdf({ title: SPO_RESOURCE.name });
  let wsY = addPdfHeader(ws, "SPO Planning Template -- Body Paragraph 3", {
    color: C.PRIMARY,
    subtitle: "Topic: Arrival in Sydney Cove and Early Settlement",
    lessonInfo: "Lesson 19 | Week 4 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "Today is PLANNING -- not the full paragraph. Use point form for supporting details. Aim for a TS that tells the reader when, where, who and what.", wsY, { color: C.PRIMARY });

  wsY = addSectionHeading(ws, "Topic Sentence", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Write a topic sentence that signals WHEN, WHERE, WHO and WHAT (the challenges).", wsY, { fontSize: 10, italic: true });
  wsY = addLinedArea(ws, wsY, 3, { lineSpacing: 22 });
  wsY += 6;

  wsY = addSectionHeading(ws, "Supporting Details (point form)", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "List 3 supporting details from your notes. Use point form -- not full sentences yet.", wsY, { fontSize: 10, italic: true });
  wsY += 4;
  wsY = addBodyText(ws, "Detail 1:", wsY, { bold: true });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 20 });
  wsY = addBodyText(ws, "Detail 2:", wsY, { bold: true });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 20 });
  wsY = addBodyText(ws, "Detail 3:", wsY, { bold: true });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 20 });
  wsY += 6;

  wsY = addSectionHeading(ws, "Concluding Sentence", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "Wrap up the paragraph -- without repeating the TS.", wsY, { fontSize: 10, italic: true });
  wsY = addLinedArea(ws, wsY, 3, { lineSpacing: 22 });
  wsY += 6;

  wsY = addSectionHeading(ws, "Notes Bank (optional)", wsY, { color: C.ALERT });
  wsY = addBodyText(ws, "Use this space to jot down extra facts from the article that you might use.", wsY, { fontSize: 10, italic: true });
  wsY = addLinedArea(ws, wsY, 4, { lineSpacing: 20 });

  addPdfFooter(ws, "Lesson 19 | SPO Planning Template");

  // ---- PDF: SPO Model Answer ----------------------------------------------
  const ma = createPdf({ title: MODEL_RESOURCE.name });
  let maY = addPdfHeader(ma, "SPO Model Answer -- Body Paragraph 3", {
    color: C.SECONDARY,
    subtitle: "Teacher Reference: Arrival in Sydney Cove",
    lessonInfo: "Lesson 19 | Week 4 | Year 5/6 Literacy",
    showNameDate: false,
  });

  maY = addTipBox(ma, "This is one model SPO. Many alternative supporting details are valid -- accept any that draw from the supplied non-fiction articles and acknowledge the impact on Eora people.", maY, { color: C.SECONDARY });

  maY = addSectionHeading(ma, "Model Topic Sentence", maY, { color: C.PRIMARY });
  maY = addBodyText(ma, "When the First Fleet arrived in Sydney Cove in January 1788, the British settlers faced enormous challenges in setting up a colony on Eora land.", maY, { fontSize: 12 });
  maY += 8;

  maY = addSectionHeading(ma, "Model Supporting Details", maY, { color: C.SECONDARY });
  maY = addBodyText(ma, "Detail 1:  Arrived on Eora Country -- Aboriginal land for tens of thousands of years", maY);
  maY = addBodyText(ma, "Detail 2:  Land was difficult for British farming -- thin soil, wheat did not grow well", maY);
  maY = addBodyText(ma, "Detail 3:  Convicts and marines struggled -- food shortages, poor housing, disease", maY);
  maY += 8;

  maY = addSectionHeading(ma, "Model Concluding Sentence", maY, { color: C.ACCENT });
  maY = addBodyText(ma, "These early years were marked by hardship for the British and devastation for the Eora people.", maY, { fontSize: 12 });
  maY += 10;

  maY = addSectionHeading(ma, "Alternative Supporting Details (sample)", maY, { color: C.ALERT });
  maY = addBodyText(ma, "- Most convicts and marines lived in tents at first; permanent buildings took months", maY);
  maY = addBodyText(ma, "- Sergeant George Phillip was the first Governor; he led the early settlement", maY);
  maY = addBodyText(ma, "- Within a year, disease (presumed smallpox) had killed many Eora people in the area", maY);
  maY = addBodyText(ma, "- The supply ships were often late, leaving the colony with low food rations for long stretches", maY);

  addPdfFooter(ma, "Lesson 19 | SPO Model Answer -- TEACHER USE");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Lesson19.pptx` }),
    writePdf(ws, SPO_PDF_PATH),
    writePdf(ma, MODEL_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Lesson19.pptx`);
  console.log("Done: " + SPO_RESOURCE.name);
  console.log("Done: " + MODEL_RESOURCE.name);
}

build().catch(console.error);
