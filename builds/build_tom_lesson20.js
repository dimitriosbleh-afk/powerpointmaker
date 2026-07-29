"use strict";

// Tom Unit -- Lesson 20: Write Body Paragraph 3 -- Arrival in Sydney Cove
// Week 4 / Week 5 transition, Lesson 20, Grade 5/6 Literacy
// Writing: Take the SPO from Lesson 19 and write the full body paragraph 3 of the information report.
// Topic: Arrival in Sydney Cove and Early Settlement
// No new novel chapters today -- this is a full writing session.

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
  modellingSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 20;
const FOOTER = "Write Body Paragraph 3 | Lesson 20 | Week 5 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson20_Write_Body_Paragraph_Sydney_Cove";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const MENTOR_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mentor Paragraph Sydney Cove",
  "Annotated model body paragraph for Arrival in Sydney Cove and Early Settlement -- shows structure and language features.",
);
const CHECKLIST_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Body Paragraph Editing Checklist",
  "Student checklist for proofreading, revising and editing the body paragraph.",
);
const RESOURCE_ITEMS = [MENTOR_RESOURCE, CHECKLIST_RESOURCE];
const MENTOR_PDF_PATH = path.join(OUT_DIR, MENTOR_RESOURCE.fileName);
const CHECKLIST_PDF_PATH = path.join(OUT_DIR, CHECKLIST_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Lesson 20. Today we WRITE body paragraph 3 of our information report
- Topic: Arrival in Sydney Cove and Early Settlement
- You will use the SPO you planned last session

DO:
- Display title slide as students settle
- Have students' SPO planning sheets (from Lesson 19) on desks
- Have the mentor paragraph and editing checklist ready (do not distribute yet)
- Have the supplementary non-fiction articles available for reference

TEACHER NOTES:
This session mirrors Lesson 15 (writing body paragraph 2) in structure. The cognitive lift is lower because the SPO is already done. The new layer is language features -- students aim to include one appositive or relative clause in their paragraph. This is also where reading is connected to writing: students may notice the dual perspective from Chapters 40-41 in their CS.

WATCH FOR:
- Students who arrive without an SPO -- offer the model SPO from the previous session for them to use as a starting frame, then complete a quick SPO with them at small group

[Literacy: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${MENTOR_RESOURCE.name} is the annotated model body paragraph for this topic
- The ${CHECKLIST_RESOURCE.name} is what you use during the edit phase

DO:
- Print the mentor paragraph (one per student or one per pair)
- Print the editing checklist (one per student)
- Distribute the mentor during the I Do or before You Do; checklist comes out for the edit phase
- Have students' SPOs from Lesson 19 within reach

TEACHER NOTES:
The mentor and the SPO from last session are both reference points during writing. Some teachers display the mentor on screen during writing; others print it. Either works.

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_LAUNCH = `SAY:
- Quick recall before we write. On your whiteboard, write the three parts of a body paragraph
- 30 seconds. Best you can
- Some of you may remember the names; others may describe what each part does -- that is fine too

DO:
- 30 seconds silent recall on whiteboards
- Walk and scan: who has all three? who has one or two?
- Reveal on the next slide

TEACHER NOTES:
Active recall before the LI/SC. The three parts are topic sentence, supporting details, concluding sentence -- met in Lessons 9, 10, 14, 15 and 19. This launch reactivates that prior knowledge before students apply it today.

WATCH FOR:
- Students with all three -- they are ready to write
- Students with one or two -- the revise slide will rebuild the rest

[Literacy: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Today we write body paragraph 3 of our information report. Topic: Arrival in Sydney Cove
- Read the success criteria

DO:
- Choral read the LI, then the SCs
- Brief: "Today is writing. We are using the SPO you planned last session and turning it into a real paragraph"

TEACHER NOTES:
SC1 targets the structure (TS, supporting details, CS). SC2 targets language features (past tense, third person, plus at least one appositive or relative clause). SC3 targets the edit phase -- proofreading and revising for clarity.

WATCH FOR:
- Students who feel underprepared because their SPO is rough -- reassure: "The SPO is a starting point. We can update it as we write"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_BACKGROUND = `SAY:
- Quick refresh of the non-fiction reading from last session
- I will read two short extracts aloud while you look at your notes
- Update your notes if you hear something useful

DO:
- Display or briefly re-read 2-3 short sections of the non-fiction articles used last session (NMA First Fleet snapshot or MHNSW article)
- Pause to ask: "Did anything new stand out? Update your notes"
- 5 minutes maximum

TEACHER NOTES:
This is a refresher, not a re-read. The aim is to surface any facts students may have missed. Some students will have light notes; this is the moment to thicken them before writing.

WATCH FOR:
- Students whose notes have only the obvious facts (e.g. "1788") -- prompt them to add a fact about the Eora people or the convicts' first months
- Students whose notes are very full -- celebrate, and let them help a partner

[Literacy: Build Background | VTLM 2.0: Build Knowledge]`;

const NOTES_REVISE = `SAY:
- Quick revision of the structure and features
- Three parts: topic sentence, supporting details, concluding sentence
- Language features to remember: past tense, third person, specific vocabulary
- Aim for ONE appositive or ONE relative clause in your paragraph

DO:
- Display the structure card
- Quick partner check: "Tell your partner ONE language feature of an information report"
- Keep this brief: 2 minutes maximum

TEACHER NOTES:
The unit reuses appositives (Lesson 16) and relative clauses (Lesson 17) as enrichment. Even one of either is enough at the writing stage; students do not need to use both.

WATCH FOR:
- Students unsure about past tense for a historical topic -- redirect: "Historical events use past tense (arrived, were, struggled). The report voice stays factual"
- Students who think first person ("I think") is fine -- redirect to third person (the British, the convicts, the Eora people)

[Literacy: Revise | VTLM 2.0: Retention and Recall]`;

const NOTES_IDO = `SAY:
- Watch how I take my SPO from last session and turn it into a real body paragraph
- Read my SPO plan first. Notice it is just notes -- not full sentences
- Now watch as I write the paragraph
- Listen for the language features as I go: past tense, third person, an appositive and a relative clause

DO:
- Display the SPO on the left, the finished paragraph on the right (mentor)
- Read both aloud
- Highlight the appositive and the relative clause when you read the paragraph
- Pause: "Notice the TS sets up when, where and what. Notice the CS adds the final idea -- it does not just repeat the TS"

TEACHER NOTES:
The mentor paragraph models the structure and the language features in one place. Highlight the two enrichment moves (appositive, relative clause) so students see them in context. The dual-impact CS is a depth move -- not all students will reach this, but it is in the model.

WATCH FOR:
- Students drafting their own paragraph already -- celebrate the keen start
- Students who think they must use this exact paragraph -- redirect: "This is a model. Your facts will be different"

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Two topic sentences. Which one is stronger for body paragraph 3 about Arrival in Sydney Cove?
- A: "Sydney Cove was the place where the First Fleet arrived."
- B: "When the First Fleet sailed into Sydney Cove in January 1788, the British settlers and the Eora people both faced major changes."
- Hold up A or B. Three, two, one -- show

DO:
- Display both sentences
- Show Me Boards
- Scan: most students should choose B
- Cold Call 1-2 students: "Why B?"

TEACHER NOTES:
B signals when, where and what (the major changes for both groups). A is just a fact. Strong students will also notice that B acknowledges both perspectives -- British and Eora. That is the depth move.

WATCH FOR:
- Students who pick A because it is shorter -- redirect: "Shorter is not always better. Does A set up the whole paragraph?"
- Students who articulate WHY B sets up the paragraph -- ready to write

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Stronger TS: B
- B signals when (January 1788), where (Sydney Cove), who (the British and the Eora people) and what (major changes)
- A is just a fact. It does not set up the paragraph
- Aim for B-style detail in your own opener

DO:
- Display the reveal banner
- Read B aloud
- Pivot if many missed it: "Predict the paragraph from B. Now predict the paragraph from A. Which gives you more?"

CFU CHECKPOINT:
Technique: Show Me Boards (A or B)
Script:
- "Hold up A or B. Which is the stronger topic sentence?"
- Scan for: most students choose B (>=80%)
PROCEED (>=80%): Most chose B. Release to write.
PIVOT (<80%): Most likely issue -- students think a TS just states a fact. Reteach: "A TS sets up the whole paragraph -- it signals when, where, who and what." Re-check by asking students to predict the paragraph content from B alone.

TEACHER NOTES:
After reveal, release students to write. The B-style TS structure (when, where, who, what) is the takeaway.

WATCH FOR:
- Students who self-correct toward a stronger TS -- celebrate

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Your turn. Use your SPO from last session
- First: update your notes if you heard anything new in the refresher
- Next: write your TS. Aim for B-style detail -- when, where, who, what
- Then: expand each supporting detail into a full sentence. Try ONE appositive or relative clause
- Finally: write your CS. Do not repeat the TS -- add the final idea
- 18 minutes writing. I will circulate

DO:
- Distribute the mentor paragraph (one per student or per pair) -- keep on desk for reference
- Circulate -- prioritise enabling students first
- Quick conferences: "Read me your TS. Does it set up the whole paragraph?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use a sentence-frame version of the mentor paragraph with key information left blank. Students fill in the blanks using their SPO facts
- Suggested frame: "When the First Fleet arrived in Sydney Cove in ____, the British settlers faced ____. The British arrived on ____ Country, which had been ____. The land was difficult for ____ because ____. Convicts and marines struggled with ____. ____."
- Extra Notes: These students still cover all three SCs but with reduced cognitive load on sentence construction

EXTENDING PROMPT:
- Task: After writing the paragraph, add a SECOND appositive or relative clause to a different sentence. Underline both
- Extra Notes: Push specificity -- use proper nouns (e.g. Sergeant Stanley, Sydney Cove, January 1788, the Eora people) for richer detail

TEACHER NOTES:
The 18-minute block is the heart of the lesson. Active circulation is the formative assessment. Most students should finish a draft inside the block; some will still be editing. The edit phase is the next slide.

WATCH FOR:
- Students copying the mentor paragraph verbatim -- prompt: "What are YOUR three supporting details? Use your SPO"
- Students who skip the CS -- redirect: "The CS wraps up. What is the FINAL idea?"
- Students using first person ("I think the convicts...") -- redirect to third person

[Literacy: You Do -- Write | VTLM 2.0: Supported Application]`;

const NOTES_EDIT = `SAY:
- Time to edit. Use the editing checklist
- Read your paragraph aloud quietly -- does it sound right?
- Tick each item on the checklist as you check it
- Then swap with a partner. Give one specific positive comment and one suggestion

DO:
- Distribute the editing checklist (one per student)
- 8 minutes for the edit
- 4 minutes for the partner swap
- Circulate -- prompt students who finish quickly to redraft a weak sentence

TEACHER NOTES:
The checklist makes editing concrete -- it is not "fix what feels wrong". The partner swap surfaces sentences that do not flow even if the writer thinks they do. Both moves build editing as a skill, not just a check.

WATCH FOR:
- Students who tick everything without checking -- prompt: "Read me the sentence where you used your appositive"
- Students who give vague feedback ("it was good") -- model a specific comment: "I liked your TS because it told me when AND why"

[Literacy: Edit | VTLM 2.0: Reflection and Refinement]`;

const NOTES_CLOSING = `SAY:
- Quick self-check against the success criteria
- Then turn to your partner: tell them ONE thing you are proud of in your paragraph
- One more session of writing left in this report -- the conclusion is next

DO:
- Run thumbs / fingers check for each SC
- Listen in on partner shares -- note who feels strong and who feels unsure
- Collect drafts for review -- not for grading, for next-day reteach planning

TEACHER NOTES:
The reflection invites students to claim ownership of a specific sentence or move. Collect drafts so you can scan for the most common weak spot and target it tomorrow.

WATCH FOR:
- Students who name a specific sentence they are proud of -- evidence of self-awareness
- Students who feel unsure across all three SCs -- check in privately and plan small-group support

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Write Body Paragraph 3 -- Arrival in Sydney Cove -- Lesson 20";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Write Body Paragraph 3",
    "Arrival in Sydney Cove & Early Settlement",
    "Lesson 20  |  Week 5  |  Year 5/6 Literacy",
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

  // SLIDE 3 -- Launch (active recall)
  contentSlide(
    pres,
    "Launch",
    C.PRIMARY,
    "The Three Parts of a Body Paragraph",
    [
      "Mini-whiteboards out",
      "30 seconds: write the three parts of a body paragraph",
      "Names if you remember -- descriptions are fine too",
      "We will reveal together",
    ],
    NOTES_LAUNCH,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to write body paragraph 3 of our information report about arrival in Sydney Cove using our SPO and language features of an information report",
    ],
    [
      "I can write a body paragraph with a topic sentence, supporting details and a concluding sentence",
      "I can use language features of an information report including at least one appositive or relative clause",
      "I can proofread, revise and edit my paragraph for clarity and cohesion",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Build Background refresher
  contentSlide(
    pres,
    "Read & Note",
    C.SECONDARY,
    "Refresher: Arrival in Sydney Cove",
    [
      "I will re-read 2-3 short sections from our non-fiction articles",
      "Listen for: when did the First Fleet arrive? What did they find? What happened to the Eora people?",
      "Update your notes if you hear something new",
      "Sources: NMA \"First Fleet Arrives in Sydney Cove\" or MHNSW article on early Sydney for convicts",
    ],
    NOTES_BACKGROUND,
    FOOTER
  );

  // SLIDE 6 -- Revise: structure & features
  contentSlide(
    pres,
    "Revise",
    C.SECONDARY,
    "Body Paragraph -- Structure & Features",
    [
      "Topic sentence  ->  Supporting details  ->  Concluding sentence",
      "Past tense for historical events (arrived, struggled, lived)",
      "Third person (the British, the convicts, the Eora people)",
      "Aim for ONE appositive OR relative clause",
      "Use specific vocabulary -- Sydney Cove, Eora Country, January 1788, ration",
    ],
    NOTES_REVISE,
    FOOTER
  );

  // SLIDE 7 -- I Do: From SPO to Paragraph
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "From SPO to Body Paragraph 3",
    "My SPO plan:\n\nTS:\nWhen First Fleet arrived 1788\n-> challenges on Eora land\n\nDetail 1:\nEora Country -- tens of thousands of years\n\nDetail 2:\nLand hard for British farming\n-- thin soil, wheat failed\n\nDetail 3:\nConvicts & marines struggled\n-- food, housing, disease\n\nCS:\nHardship for British,\ndevastation for Eora people",
    "My finished paragraph:\n\n\"When the First Fleet arrived in Sydney Cove in January 1788, the British settlers, who had sailed for eight months from England, faced enormous challenges in setting up a colony on Eora land. The cove sat on the traditional Country of the Eora people, who had cared for that land for tens of thousands of years. The soil near the cove was thin and unsuitable for British wheat, and gardens further away grew better but were hard to defend. Convicts and marines, many of whom had never farmed, struggled with food shortages, poor housing and the spread of disease. These early years were marked by hardship for the British and devastation for the Eora people.\"",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 8 + 9 -- CFU: stronger topic sentence (reveal pair)
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
    const cardH = 1.10;
    addCard(slide, 0.5, cardY, 9, cardH, { strip: C.MUTED, fill: C.WHITE });
    slide.addText("A", {
      x: 0.7, y: cardY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    slide.addText("\"Sydney Cove was the place where the First Fleet arrived.\"", {
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
    slide.addText("\"When the First Fleet sailed into Sydney Cove in January 1788, the British settlers and the Eora people both faced major changes.\"", {
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

  // SLIDE 10 -- You Do: Write
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write Body Paragraph 3");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.85, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Topic: Arrival in Sydney Cove and Early Settlement", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:   Update your SPO with any new notes from the refresher\nNext:    Write your TS -- B-style detail (when, where, who, what)\nThen:    Expand each supporting detail. Use ONE appositive or relative clause\nFinally: Write your CS -- add the final idea, do not repeat the TS", {
      x: 0.75, y: CONTENT_TOP + 0.46, w: 8.4, h: 1.30,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 2.00;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Writing Time: 18 minutes", {
      x: 0.75, y: tipY + 0.10, w: 4, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Use your SPO and your notes -- not invented facts\n- The mentor paragraph is on your desk if you need a model\n- Vocabulary to spell carefully: Sydney Cove, Eora, convicts, marines, colony", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: 0.85,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 11 -- Edit & Proofread
  contentSlide(
    pres,
    "Edit",
    C.ACCENT,
    "Proofread, Revise, Edit",
    [
      "Read your paragraph aloud quietly -- does it sound right?",
      "Use the editing checklist -- tick each item as you check it",
      "Check the features: at least one appositive or relative clause",
      "Check the commas: non-essential information needs commas around it",
      "Swap with a partner -- one specific positive, one suggestion",
    ],
    NOTES_EDIT,
    FOOTER
  );

  // SLIDE 12 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Tell your partner one sentence in your paragraph you are proud of.",
      scItems: [
        "I can write a body paragraph with TS, supporting details and CS",
        "I can use language features of an information report including an appositive or relative clause",
        "I can proofread, revise and edit my paragraph",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (got it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Mentor Paragraph ----------------------------------------------
  const mp = createPdf({ title: MENTOR_RESOURCE.name });
  let mpY = addPdfHeader(mp, "Mentor Body Paragraph -- Annotated", {
    color: C.PRIMARY,
    subtitle: "Topic: Arrival in Sydney Cove and Early Settlement",
    lessonInfo: "Lesson 20 | Week 5 | Year 5/6 Literacy",
    showNameDate: false,
  });

  mpY = addTipBox(mp, "This is a model body paragraph showing the structure and language features of an information report. Use it as a reference when writing your own paragraph.", mpY, { color: C.PRIMARY });

  mpY = addSectionHeading(mp, "Model Body Paragraph", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "When the First Fleet arrived in Sydney Cove in January 1788, the British settlers, who had sailed for eight months from England, faced enormous challenges in setting up a colony on Eora land. The cove sat on the traditional Country of the Eora people, who had cared for that land for tens of thousands of years. The soil near the cove was thin and unsuitable for British wheat, and gardens further away grew better but were hard to defend. Convicts and marines, many of whom had never farmed, struggled with food shortages, poor housing and the spread of disease. These early years were marked by hardship for the British and devastation for the Eora people.", mpY, { fontSize: 12 });
  mpY += 14;

  mpY = addSectionHeading(mp, "Annotations -- Structure", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "Topic Sentence: \"When the First Fleet arrived in Sydney Cove in January 1788, the British settlers... faced enormous challenges in setting up a colony on Eora land.\" -- signals when, where, who and what.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Supporting Detail 1: The cove was Eora Country -- cared for over tens of thousands of years.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Supporting Detail 2: The land was difficult for British farming -- thin soil, wheat failed near the cove.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Supporting Detail 3: Convicts and marines struggled -- food shortages, poor housing, disease.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Concluding Sentence: \"These early years were marked by hardship for the British and devastation for the Eora people.\" -- adds the FINAL idea (dual impact) without repeating the TS.", mpY, { fontSize: 10 });
  mpY += 10;

  mpY = addSectionHeading(mp, "Annotations -- Language Features", mpY, { color: C.SECONDARY });
  mpY = addBodyText(mp, "Relative clause: \"who had sailed for eight months from England\" -- adds detail about the British settlers.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Relative clause: \"who had cared for that land for tens of thousands of years\" -- adds detail about the Eora people.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Appositive: \"many of whom had never farmed\" -- renames \"Convicts and marines\" with extra detail.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Past tense throughout (arrived, faced, sat, was, struggled, were) -- correct for a historical information report.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Third person: \"the British settlers\", \"the Eora people\", \"convicts and marines\" -- no I, we, you.", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Specific vocabulary: \"Sydney Cove\", \"Eora\", \"convicts\", \"marines\", \"colony\", \"traditional Country\".", mpY, { fontSize: 10 });

  addPdfFooter(mp, "Lesson 20 | Mentor Body Paragraph -- TEACHER AND STUDENT REFERENCE");

  // ---- PDF: Editing Checklist ---------------------------------------------
  const cl = createPdf({ title: CHECKLIST_RESOURCE.name });
  let clY = addPdfHeader(cl, "Body Paragraph Editing Checklist", {
    color: C.PRIMARY,
    subtitle: "Information Report: Arrival in Sydney Cove",
    lessonInfo: "Lesson 20 | Week 5 | Year 5/6 Literacy",
    showNameDate: true,
  });

  clY = addTipBox(cl, "Use this checklist to proofread and edit your body paragraph. Tick each item as you check it. Then swap with a partner for peer feedback.", clY, { color: C.PRIMARY });

  clY = addSectionHeading(cl, "Structure", clY, { color: C.PRIMARY });
  clY = addBodyText(cl, "__ My topic sentence signals WHEN, WHERE, WHO and WHAT", clY);
  clY = addBodyText(cl, "__ I have at least 3 supporting detail sentences that expand the topic sentence", clY);
  clY = addBodyText(cl, "__ My supporting details come from my notes or the article (factual, not invented)", clY);
  clY = addBodyText(cl, "__ My concluding sentence adds the final idea without repeating the topic sentence", clY);
  clY += 8;

  clY = addSectionHeading(cl, "Language Features", clY, { color: C.SECONDARY });
  clY = addBodyText(cl, "__ I used past tense consistently (arrived, was, struggled -- not is, are)", clY);
  clY = addBodyText(cl, "__ I used third person (the British, convicts, the Eora people -- not I, we, you)", clY);
  clY = addBodyText(cl, "__ I included at least one appositive OR relative clause to add detail", clY);
  clY = addBodyText(cl, "__ I used specific vocabulary (Sydney Cove, Eora, convicts, marines, colony)", clY);
  clY += 8;

  clY = addSectionHeading(cl, "Spelling, Grammar, Punctuation", clY, { color: C.ACCENT });
  clY = addBodyText(cl, "__ Sydney Cove, Eora, convicts, marines -- spelt correctly", clY);
  clY = addBodyText(cl, "__ Commas around non-essential relative clauses (e.g. \", who had sailed for eight months,\")", clY);
  clY = addBodyText(cl, "__ Commas around appositives (e.g. \", many of whom had never farmed,\")", clY);
  clY = addBodyText(cl, "__ Every sentence starts with a capital letter and ends with a full stop", clY);
  clY += 8;

  clY = addSectionHeading(cl, "Cohesion", clY, { color: C.ALERT });
  clY = addBodyText(cl, "__ Each sentence connects logically to the one before it", clY);
  clY = addBodyText(cl, "__ I read it aloud and it flows smoothly", clY);
  clY = addBodyText(cl, "__ My paragraph makes sense to a reader who has not read the article", clY);
  clY += 12;

  clY = addSectionHeading(cl, "Peer Feedback", clY, { color: C.PRIMARY });
  clY = addBodyText(cl, "Partner's name: ____________________", clY);
  clY += 4;
  clY = addBodyText(cl, "One specific positive comment:", clY, { fontSize: 10 });
  clY = addLinedArea(cl, clY, 2, { lineSpacing: 22 });
  clY += 4;
  clY = addBodyText(cl, "One suggestion for improvement:", clY, { fontSize: 10 });
  clY = addLinedArea(cl, clY, 2, { lineSpacing: 22 });

  addPdfFooter(cl, "Lesson 20 | Editing Checklist");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Lesson20.pptx` }),
    writePdf(mp, MENTOR_PDF_PATH),
    writePdf(cl, CHECKLIST_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Lesson20.pptx`);
  console.log("Done: " + MENTOR_RESOURCE.name);
  console.log("Done: " + CHECKLIST_RESOURCE.name);
}

build().catch(console.error);
