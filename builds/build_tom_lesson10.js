"use strict";

// Tom Unit -- Lesson 10: Write Body Paragraph for Information Report
// Week 2/3, Lesson 10, Year 5/6 Literacy
// No new chapters -- dedicated text-level writing lesson.
// Topic: 18th Century England.

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
  modellingSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 10;
const FOOTER = "Information Report | Lesson 10 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson10_Body_Paragraph_18C";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const MENTOR_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mentor Body Paragraph",
  "Annotated mentor paragraph showing the structure and language features of an information report."
);
const CHECKLIST_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Editing Checklist",
  "Student checklist for proofreading, revising and editing a body paragraph."
);
const RESOURCE_ITEMS = [MENTOR_RESOURCE, CHECKLIST_RESOURCE];
const MENTOR_PDF_PATH = path.join(OUT_DIR, MENTOR_RESOURCE.fileName);
const CHECKLIST_PDF_PATH = path.join(OUT_DIR, CHECKLIST_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today is a writing lesson. No new chapters. We are turning yesterday's plan into a written body paragraph
- Topic: 18th Century England, the same topic from your Single Paragraph Outline
- Have your SPO from last lesson on your desk

DO:
- Display title slide as students settle
- Check every student has their SPO. Have spare blank templates ready
- Have the non-fiction article available for reference

TEACHER NOTES:
This is a text-level writing lesson tied to the information report unit. Students convert yesterday's SPO into a full body paragraph, then proofread and edit. The mentor paragraph models the language features expected.

WATCH FOR:
- Students without a plan -- give them 5 minutes with a spare template before the main write
- Students who are eager to start writing -- they can begin during the revision slide

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Today we go from plan to paragraph, then we polish
- Read the success criteria together. Each starts with "I can"

DO:
- Choral read the LI, then each SC
- Quick check: "Hold up your SPO if you have it ready" [scan]

TEACHER NOTES:
SC1 targets structural output (TS + supporting details + CS). SC2 targets language features. SC3 targets editing. The criteria appear as a single unlabelled list per the megaprompt.

WATCH FOR:
- Students confident and ready -- they can draft during the revision slide
- Students without a plan -- redirect to a spare SPO template

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_REVISE = `SAY:
- Quick revision. The body paragraph in an information report classifies or describes one aspect of your topic
- Three parts: a topic sentence, supporting details, and a concluding sentence
- Some of you may remember our work on appositives and relative clauses. You can use them today to add detail
- If this feels new, that is okay. Aim for the structure first, then add a feature you have learned

DO:
- Display the structure and features overview
- Keep this brief: 2 to 3 minutes
- Point to the mentor paragraph -- show the topic sentence, the details and the concluding sentence

TEACHER NOTES:
The revision links sentence-level skills (appositives, relative clauses) to text-level writing. Students do not need to use every feature. Aim for one appositive or relative clause per paragraph.

WATCH FOR:
- Students overwhelmed by the feature list -- reassure: clear paragraph first, then features
- Students who already know the structure -- they can begin drafting

[General: Review | VTLM 2.0: Retention and Recall]`;

const NOTES_MENTOR = `SAY:
- This is a mentor paragraph about 18th Century England. A model that shows what we are aiming for
- Notice the topic sentence -- it tells the reader what the paragraph is about
- The middle sentences expand on that idea using facts from the article
- The last sentence wraps up without repeating the topic sentence

DO:
- Read the mentor paragraph aloud once
- Then read it again, pausing on each highlighted feature
- Point out the appositive and the relative clause
- Distribute the printed mentor paragraph if students need a desk copy

TEACHER NOTES:
The mentor paragraph is the visual anchor for this lesson. Use it to show structure and features in context, not as a list of rules. Students can refer to it during the write.

WATCH FOR:
- Students who notice features without prompting -- celebrate publicly
- Students unsure where the topic sentence is -- ask: "Which sentence tells us the BIG idea?"

[General: Mentor Text | VTLM 2.0: Worked Example / Modelling]`;

const NOTES_IDO = `SAY:
- Watch me turn my SPO into a written paragraph
- Topic sentence first: "Life in 18th century England was shaped by the dramatic changes of the Industrial Revolution."
- Now I expand each detail. Detail one becomes: "Factories, which replaced traditional craftsmen, transformed the way people worked." That is a relative clause -- "which replaced traditional craftsmen"
- Detail two: "Cities grew rapidly as workers, many of them former farm labourers, moved to find employment." That is an appositive -- "many of them former farm labourers"
- Detail three: "Living conditions in these growing cities were often poor, with overcrowding, pollution and disease widespread among the working class."
- Concluding sentence: "These dramatic changes transformed English society, creating both new opportunities and significant hardship."
- Notice -- my plan was the skeleton. The writing is the flesh

DO:
- Build the paragraph sentence by sentence on the screen or board
- Highlight the appositive and the relative clause as you add them
- Think aloud: "Does this sentence connect to the one before it? Yes. That is cohesion"
- Display the full finished paragraph at the end

TEACHER NOTES:
This I Do explicitly demonstrates converting an SPO into a paragraph. Key modelling moves: expanding notes into full sentences, weaving in features, checking cohesion. Use the mentor paragraph as a parallel reference if helpful.

WATCH FOR:
- Students mentally drafting already -- good, they are engaged
- Students confused by the expansion process -- the mentor PDF provides another model

[General: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU = `SAY:
- Quick check before you write. I am going to show you two sentences. Which one uses information report features more effectively?
- Sentence A or Sentence B
- Thumbs up for A, thumbs down for B
- Three, two, one -- show

DO:
- Use Thumbs Up/Down
- Scan for: mostly thumbs down (B is more effective)
- Cold call 2-3 students: "Name ONE feature that makes B better"

CFU CHECKPOINT:
Technique: Thumbs Up/Down

Script:
- "Which sentence is more effective for an information report? Thumbs up = A, thumbs down = B"
- Scan for: mostly thumbs down
- Follow up: "Name ONE feature that makes B better" [Cold Call 2-3 students]

PROCEED (>=80%): Most show thumbs down and can name a feature. Release to write.
PIVOT (<80%): Most likely issue -- students choose A because it feels simpler. Reteach: "Sentence A gives almost no information. Sentence B tells us WHO worked (children as young as five), WHAT the factories were like (dominated the city skyline), and uses one rich sentence instead of two. An information report needs to inform." Re-check: "Which sentence teaches the reader more about 18th century England?"

TEACHER NOTES:
This CFU checks understanding of effective information-report writing before students write. The comparison directly targets SC2 (language features).

WATCH FOR:
- Students who choose A "because it is easier" -- redirect: the goal is to inform
- Students who can name features in B -- they are ready to write

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Sentence B is the stronger choice
- B uses an appositive ("many of them children as young as five"), a relative clause ("that dominated the city skyline"), and one rich sentence instead of two short ones
- A is clear but it tells us almost nothing. An information report has to inform

DO:
- Reveal Sentence B as the answer
- Highlight the appositive and the relative clause
- Quick prompt: "Use this kind of sentence in your own paragraph"

TEACHER NOTES:
The reveal closes the loop on the CFU and primes students to apply the same moves in their independent write.

[General: CFU Reveal | VTLM 2.0: Formative Feedback]`;

const NOTES_WRITE = `SAY:
- Time to write. Open your SPO from last lesson
- First: write your topic sentence
- Next: expand each supporting detail into a full sentence. Try to include at least one appositive or relative clause
- Then: write your concluding sentence
- You have 15 minutes to draft. Then we edit

DO:
- Release students to write
- Circulate actively -- this is your conferencing time
- Priority checks: structure correct? details from the article? at least one appositive or relative clause attempted?
- Confer with enabling students first, then extending students

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide the topic sentence written out in full. Students write 2 supporting details using their SPO notes, then a concluding sentence. Provide sentence starters: "One major change was...", "As a result of this,...", "In conclusion,..."
- Extra Notes: Refer these students to the mentor paragraph for structural guidance

EXTENDING PROMPT:
- Task: After completing the first body paragraph, draft a SECOND body paragraph on a different aspect of 18th century England (e.g., education, social class, child labour). Must include at least one appositive AND one relative clause

TEACHER NOTES:
The 15-minute write is the core of this lesson. Circulate and confer -- formative assessment in real time. Students who finish early should proofread using the editing checklist before starting the extending task.

WATCH FOR:
- Students who copy SPO notes verbatim into the paragraph -- remind: "Plan is the skeleton. Now write full sentences"
- Students who forget the concluding sentence -- prompt: "How do you wrap up?"
- Students using language features naturally -- celebrate: "You used a relative clause there"

[General: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EDIT = `SAY:
- Drafting is done. Now we edit
- Use the editing checklist. Read each item and tick when done
- Read your paragraph aloud quietly to yourself. Does it flow? Does it sound right?
- Check commas around appositives and relative clauses
- After 3 minutes, swap with a partner for one positive and one suggestion

DO:
- Distribute the editing checklist if not already on desks
- 3 minutes self-edit, then 2 minutes peer swap
- Final 1-2 minutes for any final changes

TEACHER NOTES:
The edit phase is non-negotiable. Students must read and revise their own work before it is considered complete. Peer swap builds collaborative editing habits.

WATCH FOR:
- Students who say "it is fine" without reading -- insist they read aloud
- Students giving vague peer feedback -- model specific feedback: "Your topic sentence clearly introduces the topic. Could you add an appositive in detail 2?"

[General: You Do -- Edit | VTLM 2.0: Monitor Progress and Feedback]`;

const NOTES_CLOSING = `SAY:
- Quick self-check against the success criteria. Thumbs up, sideways or down for each
- Then a quick write: one thing you are proud of in your paragraph, and one thing you want to improve next time

DO:
- Run thumbs check for each SC
- The quick-write is a self-assessment tool -- collect for teacher review
- Wrap up: "Well done. You started with a plan and finished with a polished paragraph"

TEACHER NOTES:
Collecting the quick-write provides diagnostic data for the next writing lesson. Students "thumbs down" on SC2 may need more sentence-level modelling next time.

WATCH FOR:
- Students who name specific improvements -- valuable metacognition
- Collect all body paragraphs for assessment and feedback before the next writing lesson

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${MENTOR_RESOURCE.name} is a model paragraph showing the structure and features of an information report
- The ${CHECKLIST_RESOURCE.name} is for proofreading and editing your paragraph

DO:
- Print the mentor paragraph (one per student or display on screen)
- Print the editing checklist (one per student)
- Both are used during the lesson

TEACHER NOTES:
The mentor paragraph supports the I Do and the independent write. The checklist supports the edit phase.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Write Body Paragraph -- Lesson 10";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "Information Report",
    "Write Your Body Paragraph",
    "Lesson 10  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  // =========================================================================
  // SLIDE 2 -- LI / SC
  // =========================================================================
  liSlide(
    pres,
    [
      "We are learning to write a body paragraph for an information report by turning a plan into polished writing with appropriate language features",
    ],
    [
      "I can write a body paragraph that includes a topic sentence, supporting details and a concluding sentence",
      "I can use language features of an information report in my writing",
      "I can proofread, revise and edit my paragraph to improve it",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Quick Revision
  // =========================================================================
  contentSlide(
    pres,
    "Quick Revision",
    C.SECONDARY,
    "Body Paragraph -- Structure and Features",
    [
      "Structure: Topic Sentence -> Supporting Details -> Concluding Sentence",
      "Present tense and third person (people, workers, they, their)",
      "Specific vocabulary -- not vague words like \"stuff\" or \"things\"",
      "Appositives -- add detail about a noun",
      "Relative clauses -- add information using who, which, that",
      "Aim: include at least one appositive or relative clause",
    ],
    NOTES_REVISE,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Mentor Paragraph (visual anchor)
  // =========================================================================
  modellingSlide(
    pres,
    "Mentor Text",
    "Mentor Body Paragraph -- 18th Century England",
    "What to notice:\n\n* Topic sentence introduces\n  the main idea\n\n* Supporting details expand\n  with facts from the article\n\n* Concluding sentence wraps\n  up without repeating the TS\n\n* Look for an appositive and\n  a relative clause",
    "Life in 18th century England was shaped by the dramatic changes of the Industrial Revolution. Factories, which replaced traditional craftsmen, transformed the way people worked across the country. Cities grew rapidly as workers, many of them former farm labourers, moved to find employment in the new industries. Living conditions in these growing cities were often poor, with overcrowding, pollution and disease widespread among the working class. These dramatic changes transformed English society, creating both new opportunities and significant hardship for ordinary people.",
    NOTES_MENTOR,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- I Do: Plan -> Paragraph
  // =========================================================================
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "From Plan to Paragraph",
    "My SPO plan:\n\nTS: Industrial Revolution\n     shaped life\n\nDetail 1: Factories replaced\n     craftsmen\n\nDetail 2: Cities grew, people\n     moved for jobs\n\nDetail 3: Poor living conditions\n\nCS: New opportunities + hardship",
    "I expand each note into a full sentence and weave in language features as I go.\n\n\"Life in 18th century England was shaped by the dramatic changes of the Industrial Revolution. Factories, which replaced traditional craftsmen, transformed the way people worked. Cities grew rapidly as workers, many of them former farm labourers, moved for new jobs. Living conditions were often poor, with overcrowding, pollution and disease. These changes created both new opportunities and significant hardship.\"",
    NOTES_IDO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 6 + 6b -- CFU with Reveal
  // =========================================================================
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Which Sentence Is More Effective?",
      "Thumbs Up = A   |   Thumbs Down = B",
      "A: \"People worked in factories. The factories were big.\"\n\nB: \"Workers, many of them children as young as five, laboured in factories that dominated the city skyline.\"\n\nWhich one informs the reader more clearly about life in 18th century England?",
      NOTES_CFU,
      FOOTER
    ),
    (s) => {
      addCard(s, 0.5, SAFE_BOTTOM - 0.95, 9, 0.85, { strip: C.SUCCESS, fill: C.WHITE });
      s.addText("Sentence B -- richer detail, an appositive and a relative clause", {
        x: 0.75, y: SAFE_BOTTOM - 0.88, w: 8.4, h: 0.70,
        fontSize: 18, fontFace: FONT_B, color: C.PRIMARY, bold: true,
        valign: "middle", margin: 0,
      });
      s.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // =========================================================================
  // SLIDE 7 -- You Do: Write Your Paragraph
  // =========================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write Your Body Paragraph");

    addInstructionCard(s, [
      { text: "Using your SPO plan:", role: "header" },
      { text: "First: Write your topic sentence" },
      { text: "Next: Expand each detail into a full sentence" },
      { text: "Try to include at least one appositive or relative clause" },
      { text: "Then: Write your concluding sentence" },
    ], {
      x: 0.5, y: CONTENT_TOP, w: 9, h: 1.7,
      strip: C.PRIMARY, fill: C.WHITE,
      headerColor: C.PRIMARY,
    });

    const tipY = CONTENT_TOP + 1.85;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Writing Time: 15 minutes", {
      x: 0.75, y: tipY + 0.10, w: 4.5, h: 0.32,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Your SPO is the skeleton. Your writing is the flesh", options: { breakLine: true } },
      { text: "Expand notes into full, detailed sentences", options: { breakLine: true } },
      { text: "Check: does each sentence connect to the one before it?", options: { breakLine: true } },
      { text: "Use the mentor paragraph if you need a model" },
    ], {
      x: 0.75, y: tipY + 0.46, w: 8.4, h: 0.95,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bullet: true, margin: 0,
      paraSpaceAfter: 2,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WRITE);
  }

  // =========================================================================
  // SLIDE 8 -- Edit and Proofread
  // =========================================================================
  contentSlide(
    pres,
    "Edit",
    C.ACCENT,
    "Proofread, Revise, Edit",
    [
      "Read your paragraph aloud quietly -- does it flow?",
      "Use the editing checklist for structure, features, spelling and grammar",
      "Check commas around appositives and relative clauses",
      "Swap with a partner: one positive, one suggestion",
      "Make final changes based on feedback",
    ],
    NOTES_EDIT,
    FOOTER
  );

  // =========================================================================
  // SLIDE 9 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    "Quick write: one thing you are proud of in your paragraph, and one thing you want to improve next time.",
    [
      "I can write a body paragraph that includes a topic sentence, supporting details and a concluding sentence",
      "I can use language features of an information report in my writing",
      "I can proofread, revise and edit my paragraph to improve it",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 10 -- Resources
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

  // --- PDF 1: Mentor Body Paragraph (annotated) ---------------------------
  const mp = createPdf({ title: MENTOR_RESOURCE.name });
  let mpY = addPdfHeader(mp, "Mentor Body Paragraph -- Annotated", {
    color: C.SECONDARY,
    subtitle: "Information Report: 18th Century England",
    lessonInfo: "Lesson 10 | Year 5/6 Literacy",
    showNameDate: false,
  });

  mpY = addTipBox(mp, "This is a model body paragraph. Use it as a reference when writing your own. The annotations show the structure and features of an information report.", mpY, { color: C.SECONDARY });

  mpY = addSectionHeading(mp, "Model Body Paragraph", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "Life in 18th century England was shaped by the dramatic changes of the Industrial Revolution. Factories, which replaced traditional craftsmen, transformed the way people worked across the country. Cities grew rapidly as workers, many of them former farm labourers, moved to find employment in the new industries. Living conditions in these growing cities were often poor, with overcrowding, pollution and disease widespread among the working class. These dramatic changes transformed English society, creating both new opportunities and significant hardship for ordinary people.", mpY, { fontSize: 12 });
  mpY += 14;

  mpY = addSectionHeading(mp, "Structure", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "Topic Sentence: \"Life in 18th century England was shaped by the dramatic changes of the Industrial Revolution.\" -- introduces the main idea (the impact of the Industrial Revolution)", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Supporting Detail 1: Factories replaced craftsmen (evidence of change)", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Supporting Detail 2: Cities grew rapidly (consequence of change)", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Supporting Detail 3: Poor living conditions (impact of change)", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Concluding Sentence: \"These dramatic changes transformed English society...\" -- wraps up without repeating the topic sentence", mpY, { fontSize: 10 });
  mpY += 10;

  mpY = addSectionHeading(mp, "Language Features", mpY, { color: C.SECONDARY });
  mpY = addBodyText(mp, "Relative clause: \"which replaced traditional craftsmen\" (non-essential, between commas)", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Appositive: \"many of them former farm labourers\" (extra detail about the workers, between commas)", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Third person: \"people\", \"workers\", \"the working class\" -- no I, we, you", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Specific vocabulary: \"Industrial Revolution\", \"craftsmen\", \"employment\", \"overcrowding\"", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Note on tense: this paragraph uses past tense because it describes historical events. For a current-topic information report, use present tense.", mpY, { fontSize: 10, italic: true });

  addPdfFooter(mp, "Lesson 10 | Mentor Body Paragraph");

  // --- PDF 2: Editing Checklist -------------------------------------------
  const cl = createPdf({ title: CHECKLIST_RESOURCE.name });
  let clY = addPdfHeader(cl, "Body Paragraph Editing Checklist", {
    color: C.PRIMARY,
    subtitle: "Information Report: 18th Century England",
    lessonInfo: "Lesson 10 | Year 5/6 Literacy",
    showNameDate: true,
  });

  clY = addTipBox(cl, "Use this checklist to proofread and edit your paragraph. Tick each item as you check it. Then swap with a partner for peer feedback.", clY, { color: C.PRIMARY });

  clY = addSectionHeading(cl, "Structure", clY, { color: C.PRIMARY });
  clY = addBodyText(cl, "__ My paragraph has a clear topic sentence that introduces the main idea", clY);
  clY = addBodyText(cl, "__ I have at least 3 supporting detail sentences that expand the topic sentence", clY);
  clY = addBodyText(cl, "__ My supporting details come from the article (factual, not made up)", clY);
  clY = addBodyText(cl, "__ My concluding sentence wraps up without repeating the topic sentence", clY);
  clY += 8;

  clY = addSectionHeading(cl, "Language Features", clY, { color: C.SECONDARY });
  clY = addBodyText(cl, "__ I used third person (people, they, workers -- not I, we, you)", clY);
  clY = addBodyText(cl, "__ I included at least one appositive OR relative clause", clY);
  clY = addBodyText(cl, "__ I used specific vocabulary (not vague words like \"stuff\" or \"things\")", clY);
  clY += 8;

  clY = addSectionHeading(cl, "Spelling, Grammar, Punctuation", clY, { color: C.ACCENT });
  clY = addBodyText(cl, "__ I checked spelling of key vocabulary (Industrial Revolution, factories, craftsmen)", clY);
  clY = addBodyText(cl, "__ Commas are correct around appositives and non-essential relative clauses", clY);
  clY = addBodyText(cl, "__ Every sentence starts with a capital letter and ends with a full stop", clY);
  clY += 8;

  clY = addSectionHeading(cl, "Cohesion", clY, { color: C.ALERT });
  clY = addBodyText(cl, "__ Each sentence connects logically to the one before it", clY);
  clY = addBodyText(cl, "__ I read it aloud and it flows smoothly", clY);
  clY = addBodyText(cl, "__ My paragraph makes sense to a reader who has not seen my plan", clY);
  clY += 12;

  clY = addSectionHeading(cl, "Peer Feedback", clY, { color: C.PRIMARY });
  clY = addBodyText(cl, "Partner's name: ____________________", clY);
  clY += 4;
  clY = addBodyText(cl, "One positive comment:", clY, { fontSize: 10 });
  clY = addLinedArea(cl, clY, 2, { lineSpacing: 22 });
  clY += 4;
  clY = addBodyText(cl, "One suggestion for improvement:", clY, { fontSize: 10 });
  clY = addLinedArea(cl, clY, 2, { lineSpacing: 22 });

  addPdfFooter(cl, "Lesson 10 | Editing Checklist");

  // --- Write all files -----------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Lesson10.pptx` }),
    writePdf(mp, MENTOR_PDF_PATH),
    writePdf(cl, CHECKLIST_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Lesson10.pptx`);
  console.log("Done: " + MENTOR_RESOURCE.name);
  console.log("Done: " + CHECKLIST_RESOURCE.name);
}

build().catch(console.error);
