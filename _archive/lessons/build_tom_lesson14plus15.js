"use strict";

// Tom Unit -- Combined Session 4: Chapters 29-30 + Plan & Write Body Paragraph 2
// (Merges the content of original lessons 14 and 15 into one extended session.)
// Week 3, Session 4, Year 5/6 Literacy (cross-curricular HASS link)
// Reading: Chapters 29-30 (Tom is interrogated, taken into Sgt Stanley's household, walks the harbour)
// Writing: Plan and write body paragraph 2 of the information report on "Why the First Fleet came to Australia".
// v10.3: Teacher Resources slide appears as slide 2 (immediately after the title slide).
//        A multi-session teacher overview is included to set the writing structure for the longer writing arc.

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const { createTheme, weekToVariant } = require("../themes/factory");
// Unit cohesion: all Tom Week 3 lessons use the same variant.
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
const FOOTER = "Chapters 29-30 + Body Paragraph 2 | Session 4 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Session4_Free_and_Write_Report";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const SPO_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "SPO Plan Body Paragraph",
  "Student template: plan body paragraph 2 of the information report (Why the First Fleet came to Australia)."
);
const MENTOR_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mentor Body Paragraph",
  "Annotated model body paragraph showing structure and language features. Use as a reference while writing."
);
const CHECKLIST_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Editing Checklist",
  "Student checklist: proofread and edit for structure, language features, spelling and cohesion."
);
const RESOURCE_ITEMS = [SPO_RESOURCE, MENTOR_RESOURCE, CHECKLIST_RESOURCE];
const SPO_PDF_PATH = path.join(OUT_DIR, SPO_RESOURCE.fileName);
const MENTOR_PDF_PATH = path.join(OUT_DIR, MENTOR_RESOURCE.fileName);
const CHECKLIST_PDF_PATH = path.join(OUT_DIR, CHECKLIST_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back. This is our final session of the week
- Two short chapters today: 29 and 30. Tom's situation changes a lot
- Then we move to writing. You will plan AND write body paragraph 2 of our information report

DO:
- Display title slide as students settle
- Have novels, exercise books, the non-fiction article, and SPO sheets on desks

TEACHER NOTES:
This is the combined plan-and-write session for body paragraph 2. The original two-session pacing has been compressed into one extended block (approx 80-90 min total). Pace the reading tightly so writing has time. Body paragraph 1 was written earlier in the unit.

WATCH FOR:
- Students who remember body paragraph 1 -- they have a model to work from
- Students who need a refresher on SPO -- they will see one modelled

[Literacy: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_RESOURCES = `SAY:
- Three resources for today
- The ${SPO_RESOURCE.name} is for planning the paragraph
- The ${MENTOR_RESOURCE.name} is the model paragraph to reference while writing
- The ${CHECKLIST_RESOURCE.name} is for proofreading at the end

DO:
- Print the SPO template (one per student)
- Print the mentor paragraph (one per pair or per student)
- Print the editing checklist (one per student)
- Have the non-fiction article ready -- either the State Library NSW "First Fleet" sections or the NMA "Why was a convict colony set up in Australia?" page, or another school-approved article

TEACHER NOTES:
This is an extended session combining planning AND writing. All three resources are used today. Students keep their SPO and editing checklist for the writing portfolio. The mentor paragraph stays in the classroom set for later writing lessons.

[General: Resources | VTLM 2.0: Student Resources]`;

const NOTES_OVERVIEW = `SAY:
- Quick overview of what we are doing today
- The information report has three parts so far: introduction (done), body paragraph 1 (done in Session 6 of this unit), body paragraph 2 (today)
- Body paragraph 1 was about life in 18th century England
- Body paragraph 2 is about WHY the First Fleet came to Australia
- The writing structure is the same as paragraph 1 -- topic sentence, supporting details, concluding sentence
- We will read, plan, write and edit in one session today

DO:
- Display the overview slide
- Frame: "Plan + write in one go. We will be efficient with time"
- Keep this slide to 2-3 minutes maximum

TEACHER NOTES:
This is the teacher-facing multi-session overview (v10.3 section 68b). It belongs near the start so the teacher and students know the writing arc. Body paragraph structure is fixed; reasons are chosen by the student from the article.

WATCH FOR:
- Students unsure why two strands are taught together -- frame as "novel and history side by side"
- Students worried about time -- reassure that the SPO does most of the heavy lifting before writing

[Literacy: Multi-Session Overview | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands today. First, what the author shows us about Tom's freedom. Second, plan and write body paragraph 2
- Read each success criterion together

DO:
- Choral read the LI, then each SC
- Brief reminder: "Body paragraph 1 was about life in England. Body paragraph 2 is about why they came here"

TEACHER NOTES:
SC1 covers the planning skill (SPO from non-fiction). SC2 is the core writing target -- the exit ticket targets this. SC3 is the stretch -- using a language feature and editing for cohesion.

WATCH FOR:
- Students confident about SPO structure (from earlier sessions) -- they can support peers
- Students who hesitate on "language feature" -- they will see appositives and relative clauses modelled

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_READING = `SAY:
- Reading mode: teacher read aloud
- Two short chapters today. Tom's life changes completely
- I will pause twice for us to think about the author's choices
- After the read, we move to the writing -- keep your pencils ready

DO:
- Read Chapter 29 and Chapter 30
- Plan two pause points: p.158 (the moment Sgt Stanley is named) and p.160 (the walk around the harbour)
- Keep the reading tight -- 15 minutes maximum so the writing has time

TEACHER NOTES:
The reading is shorter than previous sessions. Both chapters are pivotal -- Tom moves from convict to assigned servant in a household. The two pause points capture moments of unspoken authorial work: the connection to Rob, and the contrast between expected punishment and walking together as people.

WATCH FOR:
- Students who connect Sgt Stanley to Rob -- excellent chapter-to-chapter linking
- Students moved by Tom's freedom -- acknowledge the emotional shift

[Literacy: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

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

[Literacy: Pause Point 1 | VTLM 2.0: Higher-Order Questioning]`;

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

PROCEED (>=80%): Most students see the Sergeant's choice and the change in Tom's status. Continue to the bridge.
PIVOT (<80%): Most likely issue -- students miss the symbolic weight of the walk. Reteach: "What did the Sergeant HAVE to do? Take Tom to his quarters. What did the Sergeant CHOOSE to do? Walk around the harbour. The author is showing us a CHOICE." Re-check: "Why does that choice matter for Tom?"

TEACHER NOTES:
This pause point is about authorial choice in showing relationship. The detail of the harbour walk is small but the weight is large.

WATCH FOR:
- Students who name the change in Tom's status -- excellent
- Students who notice "instead" as a key word -- strong reading

[Literacy: Pause Point 2 | VTLM 2.0: Deep Comprehension]`;

const NOTES_BRIDGE = `SAY:
- Now we shift from novel to history. Tom's story is set during the First Fleet
- Our information report has two body paragraphs. Body paragraph 1: life in 18th century England. Body paragraph 2: why the First Fleet came to Australia
- Why is this connected? [Tom only ended up here because the British government decided to send a fleet. If we understand WHY they came, we understand more about Tom's story]
- We will read the non-fiction article, plan with an SPO, then write the paragraph

DO:
- Display the bridge slide
- Brief framing: "Novel + history. We have read the story. Now we plan and write the report"
- Have the non-fiction article ready -- printed or on screen

TEACHER NOTES:
This bridge slide does not need long teaching. It signals the shift from reading to writing and connects the two strands. Keep it under 2 minutes.

WATCH FOR:
- Students who immediately see the connection between novel and history -- celebrate
- Students unsure why they are reading non-fiction -- frame it as evidence for the report

[Literacy: Bridge / Cross-Curricular | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_ARTICLE = `SAY:
- We are going to read a non-fiction article about the First Fleet
- Your teacher has the article ready
- As you read, look for: WHY did the British government send the First Fleet? Use your KPAS skills to take notes
- Look for the key reasons: prisons, transportation after losing America, finding new colonies, claiming new land

DO:
- Distribute the non-fiction article (suggested: State Library NSW First Fleet sections, or NMA "Why was a convict colony set up in Australia?", or another school-approved article)
- 8-10 minutes to read and take KPAS notes
- Circulate -- prompt students to focus on the WHY question

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide a graphic organiser with three boxes labelled "Problems in England", "Prison problems", "Why Australia?". Students fill each box with notes from the article
- Extra Notes: These students may benefit from rereading specific paragraphs in pairs

EXTENDING PROMPT:
- Task: After taking notes, students identify TWO causes and ONE consequence from the article. They use this cause-consequence framing in their topic sentence

TEACHER NOTES:
The non-fiction reading is the source for the SPO. Suggested articles from the unit plan: State Library NSW First Fleet (Floating Prisons, Setting Sail) or NMA "Why was a convict colony set up in Australia?". Teachers may use other school-approved sources. Encourage KPAS notes -- this builds on the Session 3 note-taking work.

WATCH FOR:
- Students who copy whole sentences instead of taking KPAS notes -- redirect: "You have the symbols. Use them"
- Students who find one reason and stop -- prompt: "Find at least three reasons. The article has more than one"

[Literacy: Reading + Note-Taking | VTLM 2.0: Active Reading]`;

const NOTES_IDO = `SAY:
- Watch how I turn notes into a full body paragraph
- My SPO plan is on the left. My finished paragraph is on the right
- Topic sentence: "Several social and economic problems in 18th century England led the British government to establish a convict colony in Australia."
- Notice: that topic sentence tells the reader WHY -- the main idea of the paragraph
- Detail one: "The Industrial Revolution, which transformed the economy, left thousands of workers unemployed and homeless."
- See that? "which transformed the economy" -- that is a relative clause adding detail to "Industrial Revolution"
- Detail two: "Cities such as London became overcrowded, with rising crime and growing numbers of petty thieves filling the prisons."
- Detail three: "Prisons, hulks of dilapidated ships moored on the Thames, could no longer hold all the convicts."
- See "hulks of dilapidated ships moored on the Thames" -- that is an appositive renaming "Prisons"
- Concluding sentence: "Faced with these mounting pressures, the government chose to send the First Fleet to establish a penal colony far from home."

DO:
- Display the SPO + paragraph side by side
- Point to the relative clause and the appositive as you read them
- Think aloud: "I am building each detail from my notes, and weaving in features as I go"
- Show students where each detail comes from in the article

TEACHER NOTES:
This is the most important slide of the lesson. Students see how to lift information from a non-fiction source into an SPO AND then turn the SPO into a paragraph with embedded language features. Two language features are demonstrated: a non-essential relative clause and an appositive renaming. Both use commas.

WATCH FOR:
- Students already mentally drafting -- good engagement
- Students confused about lifting from the article -- the mentor paragraph PDF gives another reference

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
This CFU directly checks SC2 (topic sentence quality). A is vague and uninformative. B identifies the cause-consequence link and signals "WHY" -- which is the paragraph's purpose. Students who choose A often think shorter = better. The reteach is: a topic sentence must signal what the paragraph is about, with enough detail to set up the supporting sentences.

WATCH FOR:
- Students who pick A "because it is shorter" -- redirect: shorter is not better when it teaches the reader nothing
- Students who can articulate WHY B is better -- they are ready to write

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- The stronger topic sentence is B
- B tells us WHY -- it names the cause (problems in England) and the consequence (a colony in Australia)
- A just says "lots of reasons" -- that does not teach the reader anything yet

DO:
- Reveal the answer
- Read B aloud
- Point out: "cause and consequence in one sentence -- that is what a strong TS does"
- Pivot if many chose A: "Look again at sentence A. What is it actually telling you about?" [Likely: nothing specific] "Now look at B. What does it tell you?" [Likely: the cause and what happened next]

CFU CHECKPOINT:
Technique: Show Me Boards (A vs B)

Script:
- "Hold up A or B. Which is the stronger topic sentence?"
- Scan for: most students choose B (~80%)
- Follow up: "Why is B better?"

PROCEED (>=80%): Most chose B and can name a feature. Release to plan and write.
PIVOT (<80%): Most likely issue -- students think shorter = clearer = better. Reteach: read both aloud back-to-back. "Which one tells you what the paragraph will be about?" Re-check: "Which one would help a reader who knows nothing about the First Fleet?"

TEACHER NOTES:
After reveal, release students to the You Do plan and write.

WATCH FOR:
- Students who now want to redraft a TS -- celebrate, they are using the model

[Literacy: CFU Reveal | VTLM 2.0: Formative Feedback]`;

const NOTES_YOUDO = `SAY:
- Your turn. Plan and write body paragraph 2
- First, finish your SPO using the article and your KPAS notes
- Then turn the SPO into a full paragraph -- topic sentence, three supporting details, concluding sentence
- Try to include at least one appositive or relative clause -- look at the mentor paragraph for examples
- You have about 25 minutes. I will circulate

DO:
- Release students -- 10 minutes plan, 15 minutes write
- Distribute the SPO template and the mentor paragraph
- Circulate -- prioritise enabling students first, then extenders
- Quick conference questions: "What is your topic sentence?" "Where will you use your appositive?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide the topic sentence pre-written: "Several social and economic problems in 18th century England led the British government to set up a colony in Australia."
- Sentence starters for details: "One problem was...", "Another reason was...", "In addition..."
- Concluding sentence starter: "Together, these problems led to..."
- Extra Notes: These students focus their effort on supporting details and concluding sentence

EXTENDING PROMPT:
- Task: After writing body paragraph 2, draft the OPENING sentence of body paragraph 3 (Arrival in Sydney Cove) -- preview thinking for next week
- Extra Notes: Must use a relative clause or an appositive in the new sentence

TEACHER NOTES:
The 25-minute combined plan + write block is the heart of this lesson. Active circulation is the formative assessment. Students who finish early should self-edit using the checklist before starting the extending task.

WATCH FOR:
- Students copying notes verbatim -- prompt: "Notes are seeds. Now grow them into full sentences"
- Students missing the concluding sentence -- prompt: "How do you wrap this up without repeating your TS?"
- Students using both an appositive and a relative clause -- celebrate publicly

[Literacy: You Do -- Plan and Write | VTLM 2.0: Independent Application]`;

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
- Tell your partner ONE thing you are proud of in your paragraph

DO:
- Run thumbs check for each SC
- Listen in on partner shares
- Collect drafts for teacher feedback before next session
- Wrap up: "You are now writing a real information report from real research. Next week we move to the next part"

TEACHER NOTES:
Collecting drafts gives diagnostic data. Students who self-rate low on SC2 need targeted support next session. The final reflection prompt (one thing you are proud of) builds confidence and ownership.

WATCH FOR:
- Students who name a specific feature in their writing -- evidence of metacognition
- Students who hesitate to share -- gentle prompt: "Even one good sentence is something to be proud of"

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Chapters 29-30 + Body Paragraph 2 -- Session 4";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "Chapters 29-30",
    "Free, and Why They Came + Write Body Paragraph 2",
    "Session 4  |  Week 3  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  // =========================================================================
  // SLIDE 2 -- Teacher Resources (v10.3 position: immediately after title)
  // =========================================================================
  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // =========================================================================
  // SLIDE 3 -- Multi-Session Overview (teacher-facing, v10.3 section 68b)
  // =========================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "Teacher Overview", { color: C.SECONDARY, w: 2.2 });
    addTitle(s, "Information Report -- Writing Arc");

    // Section: writing structure
    const cardY = CONTENT_TOP;
    addCard(s, 0.5, cardY, 4.4, 3.40, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Writing structure this unit", {
      x: 0.70, y: cardY + 0.10, w: 4.0, h: 0.32,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Introduction (earlier)", options: { breakLine: true, bullet: true } },
      { text: "Body 1: 18th century England (earlier)", options: { breakLine: true, bullet: true } },
      { text: "Body 2: Why the First Fleet came (TODAY)", options: { breakLine: true, bullet: true, bold: true } },
      { text: "Body 3+: Arrival and settlement (later)", options: { breakLine: true, bullet: true } },
      { text: "Conclusion (later)", options: { bullet: true } },
    ], {
      x: 0.70, y: cardY + 0.50, w: 4.0, h: 2.80,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      paraSpaceAfter: 4,
    });

    // Section: today's focus
    addCard(s, 5.10, cardY, 4.4, 3.40, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Today's explicit focus", {
      x: 5.30, y: cardY + 0.10, w: 4.0, h: 0.32,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Read Chapters 29-30 (~15 min)", options: { breakLine: true, bullet: true } },
      { text: "Read the non-fiction article (~10 min)", options: { breakLine: true, bullet: true } },
      { text: "Plan body paragraph 2 with SPO (~10 min)", options: { breakLine: true, bullet: true } },
      { text: "Write the body paragraph (~15 min)", options: { breakLine: true, bullet: true, bold: true } },
      { text: "Edit using checklist (~10 min)", options: { bullet: true } },
    ], {
      x: 5.30, y: cardY + 0.50, w: 4.0, h: 2.80,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      paraSpaceAfter: 4,
    });

    // Footer banner: expected output
    const bandY = cardY + 3.55;
    addCard(s, 0.5, bandY, 9, 0.45, { fill: C.PRIMARY });
    s.addText("By the end: every student has drafted ONE body paragraph (TS + 3 details + CS) with at least one appositive or relative clause.", {
      x: 0.5, y: bandY, w: 9, h: 0.45,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_OVERVIEW);
  }

  // =========================================================================
  // SLIDE 4 -- LI / SC
  // =========================================================================
  liSlide(
    pres,
    [
      "We are learning to notice the choices an author makes about Tom's freedom, and to plan and write a body paragraph for our information report on why the First Fleet came to Australia",
    ],
    [
      "I can take notes from the article and use them to plan a body paragraph with an SPO",
      "I can write a body paragraph with a topic sentence, supporting details and a concluding sentence",
      "I can include at least one appositive or relative clause and edit my paragraph for cohesion",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- Reading Launch
  // =========================================================================
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

  // =========================================================================
  // SLIDE 6 -- Pause Point 1
  // =========================================================================
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

  // =========================================================================
  // SLIDE 7 -- Pause Point 2
  // =========================================================================
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

  // =========================================================================
  // SLIDE 8 -- Bridge: novel to history
  // =========================================================================
  contentSlide(
    pres,
    "Bridge",
    C.SECONDARY,
    "From Novel to History",
    [
      "Tom is in Australia because the British government sent the First Fleet",
      "Body paragraph 1 (earlier): life in 18th century England",
      "Body paragraph 2 (today): why the First Fleet came to Australia",
      "Today: read the article, plan with an SPO, write the paragraph, edit",
    ],
    NOTES_BRIDGE,
    FOOTER
  );

  // =========================================================================
  // SLIDE 9 -- Read the article
  // =========================================================================
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

  // =========================================================================
  // SLIDE 10 -- I Do: SPO + Mentor Paragraph
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
  // SLIDE 11 + 12 -- CFU with reveal (Stronger Topic Sentence)
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
    slide.addText("CHECK", {
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
    const cardH = 1.40;
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
    const cardBH = 4.10 - cardBY;
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
      const revealY = 4.30;
      addCard(slide, 0.5, revealY, 9, 0.55, { fill: C.SUCCESS });
      slide.addText("Stronger TS: B  --  it tells the reader WHY (cause + consequence)", {
        x: 0.5, y: revealY, w: 9, h: 0.55,
        fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // =========================================================================
  // SLIDE 13 -- You Do: Plan and Write
  // =========================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Plan and Write Body Paragraph 2");

    addInstructionCard(s, [
      { text: "Topic: Why the First Fleet came to Australia", role: "header" },
      { text: "First: Finish your SPO using your notes from the article" },
      { text: "Next: Write your topic sentence -- aim for WHY framing" },
      { text: "Then: Expand each detail. Try one appositive or relative clause" },
      { text: "Finally: Write your concluding sentence and reread" },
    ], {
      x: 0.5, y: CONTENT_TOP, w: 9, h: 1.95,
      strip: C.PRIMARY, fill: C.WHITE,
      headerColor: C.PRIMARY,
    });

    const tipY = CONTENT_TOP + 2.10;
    const tipH = SAFE_BOTTOM - tipY - 0.20;
    addCard(s, 0.5, tipY, 9, tipH, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Writing Time: about 25 minutes (plan + write)", {
      x: 0.75, y: tipY + 0.10, w: 8.0, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Topic sentence introduces the WHY (cause + consequence)", options: { breakLine: true, bullet: true } },
      { text: "Each detail must come from the article and answer WHY they came", options: { breakLine: true, bullet: true } },
      { text: "Use the mentor paragraph if you need a model", options: { bullet: true } },
    ], {
      x: 0.75, y: tipY + 0.46, w: 8.0, h: tipH - 0.56,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
      paraSpaceAfter: 4,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // =========================================================================
  // SLIDE 14 -- Edit & Proofread
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
  // SLIDE 15 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    {
      reflectionPrompt: "Tell your partner one thing you are proud of in your paragraph.",
      scItems: [
        "I can take notes from the article and use them to plan a body paragraph with an SPO",
        "I can write a body paragraph with a topic sentence, supporting details and a concluding sentence",
        "I can include at least one appositive or relative clause and edit my paragraph for cohesion",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (got it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // -----------------------------------------------------------------------
  // Companion PDFs
  // -----------------------------------------------------------------------

  // --- PDF 1: SPO Plan ----------------------------------------------------
  const sp = createPdf({ title: SPO_RESOURCE.name });
  let spY = addPdfHeader(sp, "SPO Plan -- Body Paragraph 2", {
    color: C.PRIMARY,
    subtitle: "Information Report: Why the First Fleet came to Australia",
    lessonInfo: "Session 4 | Week 3 | Year 5/6 Literacy",
    showNameDate: true,
  });

  spY = addTipBox(sp, "Use this template to plan body paragraph 2. Read the article first; use KPAS to take notes; then build your SPO. After planning, write the paragraph in your writing book using the mentor paragraph as a model.", spY, { color: C.PRIMARY });

  spY = addSectionHeading(sp, "Topic", spY, { color: C.SECONDARY });
  spY = addBodyText(sp, "Why the First Fleet came to Australia", spY, { fontSize: 12 });
  spY += 8;

  spY = addSectionHeading(sp, "Topic Sentence (TS)", spY, { color: C.PRIMARY });
  spY = addBodyText(sp, "Your TS should introduce the WHY (cause + consequence).", spY, { fontSize: 10, italic: true });
  spY = addTipBox(sp, "Starter (optional): \"Several social and economic problems in 18th century England led the British government to...\"", spY, { color: C.SECONDARY });
  spY = addLinedArea(sp, spY, 2, { lineSpacing: 24 });
  spY += 8;

  spY = addSectionHeading(sp, "Supporting Details", spY, { color: C.PRIMARY });
  spY = addBodyText(sp, "Find 3 reasons from the article. Notes are fine -- short phrases or KPAS.", spY, { fontSize: 10, italic: true });
  spY += 4;
  spY = addBodyText(sp, "Reason 1:", spY, { fontSize: 11 });
  spY = addLinedArea(sp, spY, 2, { lineSpacing: 22 });
  spY += 4;
  spY = addBodyText(sp, "Reason 2:", spY, { fontSize: 11 });
  spY = addLinedArea(sp, spY, 2, { lineSpacing: 22 });
  spY += 4;
  spY = addBodyText(sp, "Reason 3:", spY, { fontSize: 11 });
  spY = addLinedArea(sp, spY, 2, { lineSpacing: 22 });
  spY += 8;

  spY = addSectionHeading(sp, "Concluding Sentence (CS)", spY, { color: C.PRIMARY });
  spY = addBodyText(sp, "Wrap up your paragraph. Do NOT repeat your TS word-for-word.", spY, { fontSize: 10, italic: true });
  spY = addLinedArea(sp, spY, 2, { lineSpacing: 24 });
  spY += 10;

  spY = addSectionHeading(sp, "Self-Check Before Writing", spY, { color: C.ALERT });
  spY = addBodyText(sp, "- Does my TS introduce the WHY?", spY, { fontSize: 10 });
  spY = addBodyText(sp, "- Does each reason come from the article?", spY, { fontSize: 10 });
  spY = addBodyText(sp, "- Does each reason answer the question 'why did they come?'", spY, { fontSize: 10 });
  spY = addBodyText(sp, "- Does my CS wrap up without repeating the TS?", spY, { fontSize: 10 });

  addPdfFooter(sp, "Session 4 | SPO Plan -- Body Paragraph 2");

  // --- PDF 2: Mentor Paragraph --------------------------------------------
  const mp = createPdf({ title: MENTOR_RESOURCE.name });
  let mpY = addPdfHeader(mp, "Mentor Body Paragraph -- Annotated", {
    color: C.PRIMARY,
    subtitle: "Why the First Fleet Came to Australia",
    lessonInfo: "Session 4 | Week 3 | Year 5/6 Literacy",
    showNameDate: false,
  });

  mpY = addTipBox(mp, "Use this model body paragraph as a reference while writing. Look at the structure, the language features and the vocabulary.", mpY, { color: C.PRIMARY });

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

  addPdfFooter(mp, "Session 4 | Mentor Body Paragraph -- STUDENT REFERENCE");

  // --- PDF 3: Editing Checklist -------------------------------------------
  const cl = createPdf({ title: CHECKLIST_RESOURCE.name });
  let clY = addPdfHeader(cl, "Body Paragraph Editing Checklist", {
    color: C.PRIMARY,
    subtitle: "Information Report: Why the First Fleet Came",
    lessonInfo: "Session 4 | Week 3 | Year 5/6 Literacy",
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

  addPdfFooter(cl, "Session 4 | Editing Checklist");

  // --- Write all ----------------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Session4.pptx` }),
    writePdf(sp, SPO_PDF_PATH),
    writePdf(mp, MENTOR_PDF_PATH),
    writePdf(cl, CHECKLIST_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Session4.pptx`);
  console.log("Done: " + SPO_RESOURCE.name);
  console.log("Done: " + MENTOR_RESOURCE.name);
  console.log("Done: " + CHECKLIST_RESOURCE.name);
}

build().catch(console.error);
