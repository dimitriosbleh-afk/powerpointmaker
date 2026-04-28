"use strict";

// Tom Unit -- Lesson 15: Write Body Paragraph 2 -- Why the First Fleet Came to Australia
// Week 3 final session, Grade 5/6 Literacy
// No new chapters -- dedicated text-level writing session
// Writing: Read non-fiction, take notes, write second body paragraph from SPO, edit

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const { createTheme, weekToVariant } = require("../themes/factory");
// Unit cohesion: all Tom lessons use the same variant (matches sessions 6-10).
const T = createTheme("literacy", "grade56", weekToVariant(2));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  addTopBar, addBadge, addTitle, addCard, addInstructionCard, addFooter,
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

const SESSION_NUMBER = 15;
const FOOTER = "Information Report | Lesson 15 | Week 3 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson15_Body_Paragraph_First_Fleet";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const MENTOR_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mentor Body Paragraph",
  "Teacher and student reference: annotated model body paragraph on why the First Fleet came to Australia."
);
const CHECKLIST_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Editing Checklist",
  "Student checklist: proofread and edit body paragraph for structure, language features, spelling and cohesion."
);
const RESOURCE_ITEMS = [MENTOR_RESOURCE, CHECKLIST_RESOURCE];
const MENTOR_PDF_PATH = path.join(OUT_DIR, MENTOR_RESOURCE.fileName);
const CHECKLIST_PDF_PATH = path.join(OUT_DIR, CHECKLIST_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Lesson 15. No new novel chapters today
- This is a dedicated writing session. Today you write the second body paragraph of your information report
- The topic of this paragraph is: Why the First Fleet came to Australia
- Have your SPO from earlier sessions ready

DO:
- Display title slide as students settle
- Check every student has their information report draft and SPO. Have spares ready
- Have the supplementary reading material ready to display or distribute

TEACHER NOTES:
This session continues the information report writing strand. Students have already written body paragraph 1 (about 18th century England) in Session 10. Today they research and write body paragraph 2 about why the First Fleet came to Australia. The supplementary texts listed in the unit plan support background-knowledge building.

WATCH FOR:
- Students who have lost or forgotten their SPO -- have spare blank templates ready
- Students who have not yet finished body paragraph 1 -- they can complete it during writing time

[Literacy: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Read the success criteria
- This lesson has three big parts: read and take notes, then write, then edit

DO:
- Choral read the LI, then the SCs
- Brief check: "Hold up your SPO if you have it ready" [scan]
- Reassure: "If your SPO is not finished, that is okay. We will update it today using new information"

TEACHER NOTES:
SC1 targets the research and note-taking move. SC2 targets the structural writing (TS + supporting + CS) with language features. SC3 targets the editing move. The progression goes: gather information, then write, then refine.

WATCH FOR:
- Students confident and eager to write -- they can start drafting during the I Do
- Students unsure about the topic -- the article will give them what they need

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_BACKGROUND = `SAY:
- Before we write, we read. Knowing the facts is what makes a strong information report
- I will read the article aloud. You will take notes on key reasons the First Fleet came to Australia
- Use point form. Just key words and short phrases
- Listen for: problems in England, what was happening with prisons, who was being sent and why

DO:
- Display or distribute the supplementary article (State Library NSW or National Museum of Australia article from the unit plan)
- Read aloud or have students read in pairs (10 minutes)
- Pause every 1-2 paragraphs and ask: "What is one key reason you heard?" [Cold Call]
- Have students jot notes as they listen / read
- After reading: "Turn to your partner. Share your top three reasons" [2 minutes]

TEACHER NOTES:
Use one of the supplementary texts listed in the unit plan: First Fleet (State Library NSW Floating Prisons / Setting Sail sections) or "Why was a convict colony set up in Australia?" (National Museum of Australia). Teachers may select alternative reading material. The note-taking is the bridge between reading and writing -- students need raw material before they can draft.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide a graphic organiser with three boxes labelled: "Problems in England", "Prison problems", "Why Australia?". Students fill each box with notes from the article
- Extra Notes: These students may benefit from rereading specific paragraphs in pairs

EXTENDING PROMPT:
- Task: After taking notes, students identify TWO causes and ONE consequence from the article. They use this cause-consequence framing in their topic sentence
- Extra Notes: This deepens the analysis without adding more reading

WATCH FOR:
- Students copying full sentences -- redirect: "Notes are key words and short phrases, not full sentences"
- Students with no notes after 5 minutes -- prompt: "Tell me one thing you remember from the article"

[Literacy: Build Background | VTLM 2.0: Build Knowledge / Activate Prior Knowledge]`;

const NOTES_REVISION = `SAY:
- Quick reminder. A body paragraph is not a list of facts. It is structured
- Topic sentence introduces the main idea. Supporting details give the facts. Concluding sentence wraps up
- Language features we have learned: appositives from Lesson 16-style work, relative clauses from earlier work
- For a HISTORICAL information report, past tense is acceptable. Information reports can use past tense when describing past events
- Aim to include at least one appositive or relative clause in your paragraph

DO:
- Display the structure and features overview
- Point to last lesson's mentor text if available
- Keep this brief: 2-3 minutes maximum

TEACHER NOTES:
The revision is intentionally lean. Students have done structural work before. The key reminder is that historical information reports can use past tense -- this is a common student question. The other key prompt is to use at least one taught language feature.

WATCH FOR:
- Students who try to use every feature -- redirect: "Pick one or two and use them well"
- Students who default to listing facts in simple sentences -- prompt them to expand using a relative clause or appositive

[Literacy: Revision | VTLM 2.0: Retention and Recall]`;

const NOTES_IDO = `SAY:
- Watch how I turn my notes into a body paragraph
- My SPO plan is on the left. My finished paragraph is on the right
- Topic sentence: "Several social and economic problems in 18th century England led the British government to establish a convict colony in Australia."
- Notice: that topic sentence tells the reader WHY -- the main idea of the paragraph
- Detail one: I expand my "Industrial Revolution" note: "The Industrial Revolution, which transformed the economy, left thousands of workers unemployed and homeless."
- See that? "which transformed the economy" -- a relative clause adding detail
- Detail two: "Cities such as London became overcrowded, with rising crime and growing numbers of petty thieves filling the prisons."
- Detail three: "Prisons, hulks of dilapidated ships moored on the Thames, could no longer hold all the convicts."
- See "hulks of dilapidated ships moored on the Thames" -- that is an appositive renaming "Prisons"
- Concluding sentence: "Faced with these mounting pressures, the government chose to send the First Fleet to establish a penal colony far from home."

DO:
- Display the SPO + paragraph side by side
- Point to the appositive and the relative clause as you read them
- Think aloud: "I am building each detail from my notes, and weaving in features as I go"

TEACHER NOTES:
The I Do explicitly demonstrates expanding notes into a body paragraph with embedded language features. The mentor PDF gives students another reference. Highlight the cohesion -- each sentence flows logically into the next.

WATCH FOR:
- Students mentally drafting already -- good engagement
- Students confused by the mention of past tense -- reassure that this is correct for historical reports

[Literacy: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Two topic sentences. Which one is stronger for our paragraph about WHY the First Fleet came?
- Read sentence A. Read sentence B
- Hold up A or B on your mini-whiteboard
- Three, two, one -- show!

DO:
- Display both sentences clearly
- Use Show Me Boards
- Scan: most students should choose B
- Call on 2 students: "Why did you choose B?"

TEACHER NOTES:
This CFU directly checks SC2 (topic sentence quality). A is vague and uninformative. B identifies the cause-consequence link and signals "WHY" -- which is the paragraph's purpose. Students who choose A often think simpler = better. The reteach is: a topic sentence must signal what the paragraph is about, with enough detail to set up the supporting sentences.

WATCH FOR:
- Students who pick A "because it is shorter" -- redirect: shorter is not better when it teaches the reader nothing
- Students who can articulate WHY B is better -- they are ready to write

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- The stronger topic sentence is B
- B tells us WHY -- it names the cause (problems in England) and the consequence (a colony in Australia)
- A just says "lots of reasons" -- that does not teach the reader anything yet

DO:
- Display the reveal
- Read B aloud
- Point out: "cause and consequence in one sentence -- that is what a strong TS does"
- Pivot if many chose A: "Look again at sentence A. What is it actually telling you about?" [Likely answer: nothing specific] "Now look at B. What does it tell you?" [Likely: the cause and what happened next]

CFU CHECKPOINT:
Technique: Show Me Boards (A vs B)
Script:
- "Hold up A or B. Which is the stronger topic sentence?"
- Scan for: most students choose B (~80%)
PROCEED (>=80%): Most chose B and can name a feature. Release to write.
PIVOT (<80%): Most likely issue -- students think shorter = clearer = better. Reteach: read both aloud back-to-back. "Which one tells you what the paragraph will be about?" Re-check: "Which one would help a reader who knows nothing about the First Fleet?"

TEACHER NOTES:
This is the reveal half of the CFU pair. After reveal, release students to write.

WATCH FOR:
- Students who now want to redraft their TS -- celebrate, they are using the model

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Time to write your body paragraph. Use your notes and SPO
- First: update your SPO with any new information from the article
- Next: write your topic sentence -- aim for "WHY" framing like our model
- Then: expand each supporting detail into a full sentence. Try one appositive or relative clause
- Finally: write your concluding sentence
- You have 15 minutes to draft. Then we edit
- The mentor paragraph is on your desk if you need a reference

DO:
- Release students to write
- Distribute the mentor body paragraph (one per pair or per student)
- Circulate -- prioritise enabling students first, then extenders
- Quick conference questions: "What is your topic sentence?" "Where will you use your appositive?"
- Confer for 30-60 seconds per student

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide the topic sentence written out: "Several social and economic problems in 18th century England led the British government to set up a colony in Australia."
- Sentence starters for details: "One problem was...", "Another reason was...", "In addition..."
- Concluding sentence starter: "Together, these problems led to..."
- Extra Notes: These students focus their effort on supporting details and concluding sentence

EXTENDING PROMPT:
- Task: After writing body paragraph 2, draft the OPENING sentence of body paragraph 3 (Arrival in Sydney Cove) -- preview thinking for next lesson
- Extra Notes: Must use a relative clause or an appositive in the new sentence

TEACHER NOTES:
The 15-minute writing block is the heart of this lesson. Active circulation is the formative assessment. Students who finish early should self-edit using the checklist before starting the extending task.

WATCH FOR:
- Students copying their notes word-for-word -- prompt: "Notes are seeds. Now grow them into full sentences"
- Students missing the concluding sentence -- prompt: "How do you wrap this up without repeating your TS?"
- Students using both an appositive and a relative clause -- celebrate publicly

[Literacy: You Do -- Independent Writing | VTLM 2.0: Supported Application]`;

const NOTES_EDIT = `SAY:
- Drafting time is up. Time to edit
- Use the editing checklist. Read through your paragraph and check each item
- Read your paragraph aloud quietly to yourself. Your ear catches what your eye misses
- Pay attention to commas around appositives and relative clauses
- After 4 minutes, swap with a partner. Give one positive comment and one specific suggestion

DO:
- Distribute the editing checklist (one per student)
- 4 minutes self-edit
- 3 minutes peer swap
- Final 2 minutes: students make changes based on feedback

TEACHER NOTES:
The edit phase is non-negotiable. Reading aloud is the single most effective edit move for upper primary. The peer swap creates a second pair of eyes and reinforces the success criteria as a class standard.

WATCH FOR:
- Students who say "it is fine" without reading -- insist: "Read it aloud. Do not skip"
- Students giving vague peer feedback ("good job") -- model specific feedback aloud: "Your topic sentence sets up the WHY clearly. Could detail two use an appositive?"

[Literacy: Edit -- Refining | VTLM 2.0: Monitor Progress and Feedback]`;

const NOTES_CLOSING = `SAY:
- Quick self-check. Show me on your fingers
- SC1: I took notes from the article -- 1 to 5
- SC2: I wrote a body paragraph with TS, supporting details and CS, with at least one feature -- 1 to 5
- SC3: I edited my paragraph -- 1 to 5
- One thing you are proud of in your paragraph -- tell your partner

DO:
- Run each SC with a finger rating
- Record any patterns -- if many students score themselves low on SC2, plan reteaching for the next writing session
- Collect drafts for teacher feedback before the next writing lesson
- Wrap: "You are over halfway through your information report. Next session we plan body paragraph 3 about arrival in Sydney Cove"

TEACHER NOTES:
Collecting drafts gives diagnostic data. Students who self-rate 1-2 on any SC need targeted small-group support next session. The final reflection prompt (one thing you are proud of) builds confidence and ownership.

WATCH FOR:
- Students who name a specific feature in their writing -- evidence of metacognition
- Students who hesitate to share -- gentle prompt: "Even one good sentence is something to be proud of"

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${MENTOR_RESOURCE.name} is the model paragraph with annotations -- use it as a reference while writing
- The ${CHECKLIST_RESOURCE.name} is what you use during the edit phase

DO:
- Print the mentor paragraph (one per pair or one per student)
- Print the editing checklist (one per student)
- Both used during the lesson; not take-home

TEACHER NOTES:
The mentor paragraph can be displayed or printed. Some teachers prefer to display during I Do then distribute during writing. The checklist comes out for the edit phase.

[Literacy: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Information Report Body Paragraph 2 -- Lesson 15";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "Information Report",
    "Write Body Paragraph 2: Why the First Fleet Came",
    "Lesson 15  |  Week 3  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  // =========================================================================
  // SLIDE 2 -- LI / SC
  // =========================================================================
  liSlide(
    pres,
    [
      "We are learning to write a body paragraph for our information report about why the First Fleet came to Australia, drawing on a non-fiction article and editing for clarity",
    ],
    [
      "I can take notes from a non-fiction article to support my body paragraph",
      "I can write a body paragraph with a topic sentence, supporting details and a concluding sentence using language features of an information report",
      "I can proofread, revise and edit my paragraph to refine language and ensure cohesion",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Build Background: Read & Take Notes
  // =========================================================================
  contentSlide(
    pres,
    "Read & Note",
    C.SECONDARY,
    "Build Background: Why the First Fleet?",
    [
      "Listen as the article is read aloud (or read in pairs)",
      "Take notes in point form -- key words, not full sentences",
      "Listen for: problems in England, prison problems, why Australia",
      "Update your SPO with any new information",
      "Sources: First Fleet (State Library NSW) or NMA First Fleet article",
    ],
    NOTES_BACKGROUND,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Quick Revision: Body Paragraph Structure
  // =========================================================================
  contentSlide(
    pres,
    "Revise",
    C.SECONDARY,
    "Body Paragraph -- Structure & Features",
    [
      "Topic sentence -> Supporting details -> Concluding sentence",
      "Use past tense for historical information reports (this is correct)",
      "Use third person -- they, the British, convicts, the government",
      "Aim to use one appositive OR relative clause to add detail",
      "Use specific vocabulary from the article -- e.g. penal colony, transportation, Industrial Revolution",
    ],
    NOTES_REVISION,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- I Do: Mentor Paragraph
  // =========================================================================
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "From Notes to Body Paragraph",
    "My SPO plan:\n\nTS:\nProblems in England led to a\nconvict colony in Australia\n\nDetail 1:\nIndustrial Revolution -> jobs lost\n\nDetail 2:\nCities overcrowded, more crime\n\nDetail 3:\nPrisons full, hulks too\n\nCS:\nGovernment sent the First Fleet",
    "My finished paragraph:\n\n\"Several social and economic problems in 18th century England led the British government to establish a convict colony in Australia. The Industrial Revolution, which transformed the economy, left thousands of workers unemployed and homeless. Cities such as London became overcrowded, with rising crime and growing numbers of petty thieves filling the prisons. Prisons, hulks of dilapidated ships moored on the Thames, could no longer hold all the convicts. Faced with these mounting pressures, the government chose to send the First Fleet to establish a penal colony far from home.\"",
    NOTES_IDO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 6 + 7 -- CFU: Better topic sentence (reveal pair)
  // =========================================================================
  function buildCfuBase() {
    const slide = pres.addSlide();
    addTopBar(slide, C.ALERT);
    addBadge(slide, "CFU", { color: C.ALERT });
    addTitle(slide, "Which Is the Stronger Topic Sentence?", { color: C.ALERT });

    // CHECK stamp top right
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

    // Technique pill
    slide.addShape("roundRect", {
      x: 0.5, y: CONTENT_TOP, w: 2.8, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ALERT },
    });
    slide.addText("Show Me Boards: A or B", {
      x: 0.5, y: CONTENT_TOP, w: 2.8, h: 0.40,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Sentence A card
    const cardY = CONTENT_TOP + 0.55;
    const cardH = 1.45;
    addCard(slide, 0.5, cardY, 9, cardH, { strip: C.MUTED, fill: C.WHITE });
    slide.addText("A", {
      x: 0.7, y: cardY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    slide.addText("\"There were lots of reasons people came to Australia.\"", {
      x: 1.2, y: cardY + 0.18, w: 8.0, h: cardH - 0.30,
      fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Sentence B card
    const cardBY = cardY + cardH + 0.18;
    const cardBH = SAFE_BOTTOM - cardBY - 0.55;
    addCard(slide, 0.5, cardBY, 9, cardBH, { strip: C.SECONDARY, fill: C.BG_CARD });
    slide.addText("B", {
      x: 0.7, y: cardBY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    slide.addText("\"Several social and economic problems in 18th century England led the British government to establish a convict colony in Australia.\"", {
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
      // Add a reveal banner at the bottom
      const revealY = 4.90;
      addCard(slide, 0.5, revealY, 9, 0.42, { fill: C.SUCCESS });
      slide.addText("Stronger TS: B  --  it tells the reader WHY", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // =========================================================================
  // SLIDE 8 -- You Do: Write Your Body Paragraph
  // =========================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write Your Body Paragraph");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.85, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Topic: Why the First Fleet came to Australia", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:  Update your SPO with any new information from the article\nNext:   Write your topic sentence -- aim for a WHY-style opener\nThen:   Expand each supporting detail. Try one appositive or relative clause\nFinally: Write your concluding sentence", {
      x: 0.75, y: CONTENT_TOP + 0.46, w: 8.4, h: 1.30,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 2.00;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Writing Time: 15 minutes", {
      x: 0.75, y: tipY + 0.10, w: 4, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Use your notes from the article -- not made-up facts\n- The mentor paragraph is on your desk if you need a model\n- Vocabulary to spell carefully: Industrial Revolution, penal colony, convicts", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: 0.85,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // =========================================================================
  // SLIDE 9 -- Edit & Proofread
  // =========================================================================
  contentSlide(
    pres,
    "Edit",
    C.ACCENT,
    "Proofread, Revise, Edit",
    [
      "Read your paragraph aloud quietly -- does it sound right?",
      "Check structure: topic sentence, supporting details, concluding sentence",
      "Check features: at least one appositive or relative clause used",
      "Check punctuation: commas around non-essential clauses",
      "Swap with a partner -- one positive comment, one specific suggestion",
    ],
    NOTES_EDIT,
    FOOTER
  );

  // =========================================================================
  // SLIDE 10 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    {
      reflectionPrompt: "Tell your partner one thing you are proud of in your paragraph.",
      scItems: [
        "I can take notes from a non-fiction article to support my body paragraph",
        "I can write a body paragraph with TS, supporting details and CS using language features",
        "I can proofread, revise and edit my paragraph",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (got it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 11 -- Resources
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

  // --- PDF 1: Mentor Body Paragraph -----------------------------------------
  const mp = createPdf({ title: MENTOR_RESOURCE.name });
  let mpY = addPdfHeader(mp, "Mentor Body Paragraph -- Annotated", {
    color: C.PRIMARY,
    subtitle: "Why the First Fleet Came to Australia",
    lessonInfo: "Lesson 15 | Week 3 | Year 5/6 Literacy",
    showNameDate: false,
  });

  mpY = addTipBox(mp, "This is a model body paragraph showing the structure and language features of an information report. Use it as a reference when writing your own paragraph.", mpY, { color: C.PRIMARY });

  mpY = addSectionHeading(mp, "Model Body Paragraph", mpY, { color: C.PRIMARY });

  mpY = addBodyText(mp, "Several social and economic problems in 18th century England led the British government to establish a convict colony in Australia. The Industrial Revolution, which transformed the economy, left thousands of workers unemployed and homeless. Cities such as London became overcrowded, with rising crime and growing numbers of petty thieves filling the prisons. Prisons, hulks of dilapidated ships moored on the Thames, could no longer hold all the convicts. Faced with these mounting pressures, the government chose to send the First Fleet to establish a penal colony far from home.", mpY, { fontSize: 12 });
  mpY += 14;

  mpY = addSectionHeading(mp, "Annotations -- Structure", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "Topic Sentence: \"Several social and economic problems in 18th century England led the British government to establish a convict colony in Australia.\" -- introduces the main idea (cause -> consequence)", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Supporting Detail 1: Industrial Revolution -> unemployment and displacement", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Supporting Detail 2: Overcrowded cities and rising crime", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Supporting Detail 3: Prisons (and hulks) full to capacity", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Concluding Sentence: \"Faced with these mounting pressures, the government chose to send the First Fleet...\" -- wraps up without repeating the TS", mpY, { fontSize: 10 });
  mpY += 10;

  mpY = addSectionHeading(mp, "Annotations -- Language Features", mpY, { color: C.SECONDARY });
  mpY = addBodyText(mp, "Relative clause: \"which transformed the economy\" (non-essential, commas needed) -- adds detail about the Industrial Revolution", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Appositive: \"hulks of dilapidated ships moored on the Thames\" -- renames \"Prisons\" with extra detail", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Past tense throughout (led, transformed, became, could) -- correct for historical information reports", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Third person: \"the British government\", \"workers\", \"the convicts\" (no I, we, you)", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Specific vocabulary: \"Industrial Revolution\", \"petty thieves\", \"penal colony\", \"hulks\"", mpY, { fontSize: 10 });

  addPdfFooter(mp, "Lesson 15 | Mentor Body Paragraph -- TEACHER AND STUDENT REFERENCE");

  // --- PDF 2: Editing Checklist ---------------------------------------------
  const cl = createPdf({ title: CHECKLIST_RESOURCE.name });
  let clY = addPdfHeader(cl, "Body Paragraph Editing Checklist", {
    color: C.PRIMARY,
    subtitle: "Information Report: Why the First Fleet Came",
    lessonInfo: "Lesson 15 | Week 3 | Year 5/6 Literacy",
    showNameDate: true,
  });

  clY = addTipBox(cl, "Use this checklist to proofread and edit your body paragraph. Tick each item as you check it. Then swap with a partner for peer feedback.", clY, { color: C.PRIMARY });

  clY = addSectionHeading(cl, "Structure", clY, { color: C.PRIMARY });
  clY = addBodyText(cl, "__ My topic sentence introduces WHY the First Fleet came to Australia", clY);
  clY = addBodyText(cl, "__ I have at least 3 supporting detail sentences that expand the topic sentence", clY);
  clY = addBodyText(cl, "__ My supporting details come from the article (factual, not made up)", clY);
  clY = addBodyText(cl, "__ My concluding sentence wraps up without repeating the topic sentence", clY);
  clY += 8;

  clY = addSectionHeading(cl, "Language Features", clY, { color: C.SECONDARY });
  clY = addBodyText(cl, "__ I used past tense consistently (was, were, led, became -- not is, are)", clY);
  clY = addBodyText(cl, "__ I used third person (the British, convicts, the government -- not I, we, you)", clY);
  clY = addBodyText(cl, "__ I included at least one appositive OR relative clause to add detail", clY);
  clY = addBodyText(cl, "__ I used specific vocabulary from the article (e.g. penal colony, Industrial Revolution)", clY);
  clY += 8;

  clY = addSectionHeading(cl, "Spelling, Grammar, Punctuation", clY, { color: C.ACCENT });
  clY = addBodyText(cl, "__ Industrial Revolution, penal colony, convicts -- spelt correctly", clY);
  clY = addBodyText(cl, "__ Commas around non-essential relative clauses (e.g. \", which transformed the economy,\")", clY);
  clY = addBodyText(cl, "__ Commas around appositives (e.g. \"Prisons, hulks of dilapidated ships,\")", clY);
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
  clY = addBodyText(cl, "One positive comment:", clY, { fontSize: 10 });
  clY = addLinedArea(cl, clY, 2, { lineSpacing: 22 });
  clY += 4;
  clY = addBodyText(cl, "One suggestion for improvement:", clY, { fontSize: 10 });
  clY = addLinedArea(cl, clY, 2, { lineSpacing: 22 });

  addPdfFooter(cl, "Lesson 15 | Editing Checklist");

  // --- Write all files ----------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Lesson15.pptx` }),
    writePdf(mp, MENTOR_PDF_PATH),
    writePdf(cl, CHECKLIST_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Lesson15.pptx`);
  console.log("Done: " + MENTOR_RESOURCE.name);
  console.log("Done: " + CHECKLIST_RESOURCE.name);
}

build().catch(console.error);
