"use strict";

// Tom Unit -- Session 10: Write Body Paragraph for Information Report
// Week 2, Session 10, Grade 5/6 Literacy
// No new chapters -- dedicated writing session
// Writing: Write body paragraph from SPO, identify language features, proofread/edit

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const { createTheme, weekToVariant } = require("../themes/factory");
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

const SESSION_NUMBER = 10;
const FOOTER = "Information Report | Session 10 | Week 2 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Session10_Body_Paragraph";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const CHECKLIST_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Editing Checklist",
  "Student checklist: proofread and edit body paragraph for language features, spelling, grammar and cohesion."
);
const MENTOR_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mentor Body Paragraph",
  "Teacher model: annotated body paragraph showing language features of an information report."
);
const RESOURCE_ITEMS = [CHECKLIST_RESOURCE, MENTOR_RESOURCE];
const CHECKLIST_PDF_PATH = path.join(OUT_DIR, CHECKLIST_RESOURCE.fileName);
const MENTOR_PDF_PATH = path.join(OUT_DIR, MENTOR_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Session 10 -- our final session for this week. No new chapters today. This is a dedicated writing session
- Today you are writing the first body paragraph of your information report using your SPO plan from last session
- Have your SPO template ready on your desk

DO:
- Display title slide as students settle
- Check every student has their SPO from Session 9. Have blank spares for students who forgot
- Have the non-fiction article available for students who want to add to their plans

TEACHER NOTES:
Session 10 of 10. This is a writing-focused session with no new novel content. Students convert their SPO plan into a full body paragraph, then proofread and edit using a checklist. The mentor paragraph models the language features expected.

WATCH FOR:
- Students who do not have their SPO -- provide a spare template and give 5 minutes to plan before writing
- Students who are ready to write immediately -- let them start while others catch up

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Today is about writing and polishing. We go from plan to paragraph
- Read the success criteria. SC1 is about writing the paragraph with the right structure. SC2 is about using the language features of an information report. SC3 is about editing your own work

DO:
- Choral read the LI, then the SCs
- Brief check: "Who has their SPO ready? Hold it up" [scan]

TEACHER NOTES:
SC1 targets the structural output (TS + supporting details + CS). SC2 targets language features (present tense, third person, noun groups, appositives, relative clauses). SC3 targets the editing/proofreading skill.

WATCH FOR:
- Students confident and ready to write -- they can begin during the revision slide
- Students without a plan -- they need planning time before writing

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_REVISION = `SAY:
- Quick revision. The purpose of a body paragraph in an information report is to classify or describe one aspect of your topic
- Structure: topic sentence, supporting detail sentences, concluding sentence
- Language features we have been learning this unit that you can use: appositives from Session 6, relative clauses from Session 7
- Other features: present tense, third person pronouns, adjectives for description, subheadings, technical vocabulary
- You do not need to use EVERY feature in one paragraph. But aim to include at least an appositive or a relative clause to show what you have learned

DO:
- Display the structure and features overview
- Point to the mentor text body paragraph -- show where specific features appear
- Keep this brief: 3 minutes maximum. Students are eager to write

TEACHER NOTES:
The revision connects the sentence-level skills from Sessions 6 and 7 to the text-level writing. Students should see that appositives and relative clauses are tools for enriching their information report writing, not isolated grammar exercises.

WATCH FOR:
- Students who do not see the connection between sentence-level and text-level work -- make it explicit: "Remember our appositives? You can use them in your information report to add detail"
- Students who feel overwhelmed by the feature list -- reassure: "Focus on writing a clear paragraph first. Then look for places to add features"

[General: Review | VTLM 2.0: Retention and Recall]`;

const NOTES_IDO_WRITE = `SAY:
- Watch me turn my SPO into a written paragraph
- I start with my topic sentence. I wrote: "Life in 18th century England was shaped by the dramatic changes of the Industrial Revolution"
- Now I expand my first detail: "Factories, which replaced traditional craftsmen, transformed the way people worked." See that? I used a relative clause: "which replaced traditional craftsmen"
- Second detail: "Cities grew rapidly as workers, many of them former farm labourers, moved to find employment in the new industries." I used an appositive: "many of them former farm labourers"
- Third detail: "Living conditions in these rapidly growing cities were often poor, with overcrowding, pollution and disease becoming widespread"
- Concluding sentence: "These dramatic changes transformed English society, creating both new opportunities and significant hardship for ordinary people"
- Notice: my paragraph follows the SPO plan exactly. The plan is the skeleton. The writing is the flesh

DO:
- Display the model paragraph being built step by step
- Highlight the appositive and relative clause as they are added
- Think aloud: "I check -- does each sentence connect to the one before it? Yes. That is cohesion"
- Display the full finished paragraph

TEACHER NOTES:
The I Do explicitly demonstrates converting an SPO into a paragraph. The key modelling moves: expanding notes into full sentences, weaving in language features (appositive, relative clause), checking cohesion between sentences.

WATCH FOR:
- Students who are mentally drafting already -- good: they are engaged
- Students who look confused about the expansion process -- the mentor PDF provides another model

[General: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU = `SAY:
- Quick check before you write. I am going to show you two sentences. Which one uses a language feature of an information report more effectively?
- Sentence A: "People worked in factories. The factories were big"
- Sentence B: "Workers, many of them children as young as five, laboured in factories that dominated the city skyline"
- Thumbs up for A, thumbs down for B
- Three, two, one -- show! [Scan]

DO:
- Use Thumbs Up/Down
- Scan for: thumbs down (B is more effective)
- Brief discussion: "What makes B better?" [More detail, an appositive, a relative clause, precise vocabulary, one sentence instead of two]

CFU CHECKPOINT:
Technique: Thumbs Up/Down

Script:
- "Which sentence is more effective for an information report? Thumbs up = A, thumbs down = B"
- Scan for: mostly thumbs down
- Follow up: "Name ONE feature that makes B better" [Cold Call 2-3 students]

PROCEED (>=80%): Most show thumbs down and can name a feature. Release to write.
PIVOT (<80%): Most likely issue -- students choose A because it is simpler and they think "clear = good." Reteach: "Clear writing is important. But Sentence A gives almost no information. Sentence B tells us WHO worked (children as young as five), WHAT the factories were like (dominated the skyline), and does it in one sentence. An information report needs to INFORM -- B does that, A does not." Re-check: "Which sentence teaches the reader more about 18th century England?"

TEACHER NOTES:
This CFU checks that students understand what effective information report writing looks like before they begin their own paragraph. The comparison directly targets SC2 (language features).

WATCH FOR:
- Students who choose A "because it is easier to read" -- redirect: the goal is to inform, not just be simple
- Students who can name features in B -- they are ready to write

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_WRITE = `SAY:
- Time to write. Open your SPO template from last session. You are turning your plan into a full body paragraph
- First: write your topic sentence
- Next: expand each supporting detail into a full sentence. Try to include at least one appositive or relative clause
- Then: write your concluding sentence
- You have 15 minutes to draft. Then we will edit
- If you need to update your SPO with new information from the article, do that first

DO:
- Release students to write
- Circulate actively -- this is the key conferencing time
- Priority checks: Is the structure correct (TS + details + CS)? Are details from the article? Is there at least one appositive or relative clause attempt?
- Confer with enabling students first, then extending students

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide the topic sentence written out in full. Students write 2 supporting detail sentences using their SPO notes, then a concluding sentence. Sentence starters provided: "One major change was...", "As a result of this,...", "In conclusion,..."
- Extra Notes: These students can refer to the mentor paragraph for structural guidance

EXTENDING PROMPT:
- Task: After completing the first body paragraph, draft a SECOND body paragraph on a different aspect of 18th century England (e.g., education, social class, child labour). Write independently with their own topic sentence, supporting details, and concluding sentence. Must include at least one appositive AND one relative clause

TEACHER NOTES:
The 15-minute writing block is the core of this session. Circulate and confer -- this is formative assessment in real time. Students who finish early should proofread using the editing checklist before starting the extending task.

WATCH FOR:
- Students who write their SPO notes as the paragraph (no expansion) -- remind: "Your SPO is the plan. Now write full sentences"
- Students who forget the concluding sentence -- remind: "Your paragraph needs an ending. How do you wrap up?"
- Students who include language features naturally -- celebrate: "You used a relative clause there. Well done"

[General: You Do | VTLM 2.0: Supported Application]`;

const NOTES_EDIT = `SAY:
- Writing time is up. Now we edit
- Use the editing checklist on your worksheet. Read through your paragraph and check each item
- Read your paragraph aloud quietly to yourself. Does it sound right? Does it flow?
- Check spelling, grammar, and punctuation. Especially check your commas around appositives and relative clauses
- You have 5 minutes to edit. Then swap with a partner for peer feedback

DO:
- Distribute the Session 10 Editing Checklist (if not already distributed)
- Students read and self-edit using the checklist
- After 3 minutes: "Swap paragraphs with your partner. Read theirs and give ONE piece of positive feedback and ONE suggestion"
- After 5 minutes total: "Make any final changes based on your partner's feedback"

TEACHER NOTES:
The edit phase is non-negotiable -- students must read and revise their own work before it is considered complete. The peer swap adds a second pair of eyes and builds collaborative editing habits.

WATCH FOR:
- Students who say "it is fine" without reading -- insist: "Read it aloud. Your ear catches what your eye misses"
- Students who give vague peer feedback ("good job") -- model specific feedback: "Your topic sentence clearly introduces the topic. Can you add an appositive to sentence 3?"

[General: You Do -- Edit | VTLM 2.0: Monitor Progress and Feedback]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: writing a body paragraph with TS, supporting details, and CS -- thumbs? [scan]
- SC2: using at least one language feature of an information report -- thumbs? [scan]
- SC3: proofreading and editing your paragraph -- thumbs? [scan]
- Quick write: on a sticky note or the back of your checklist, write one thing you are proud of in your paragraph and one thing you want to improve next time

DO:
- Run through each SC with thumbs check
- The quick write is a self-assessment tool -- collect for teacher review
- Wrap up the week: "That is our final session for Week 2. You have read 11 chapters, learned appositives, relative clauses, KPAS note-taking, and started your information report. Well done"

TEACHER NOTES:
The closing wraps up the entire week. Collecting the quick-write sticky notes provides diagnostic data for planning Week 3. Students who are "thumbs down" on SC2 may need additional sentence-level modelling in the next week.

WATCH FOR:
- Students who are proud of their work -- celebrate publicly
- Students who identify specific improvements -- this metacognition is valuable
- Collect all body paragraphs for teacher assessment and feedback before Session 11

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${CHECKLIST_RESOURCE.name} is for proofreading and editing your body paragraph
- The ${MENTOR_RESOURCE.name} is a model paragraph showing the language features of an information report

DO:
- Print the editing checklist (one per student)
- Print the mentor paragraph (one per student or display on screen)
- Both are used during the lesson, not as take-home resources

TEACHER NOTES:
The editing checklist is used during the edit phase. The mentor paragraph can be displayed during I Do or given to students as a reference during writing.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Write Body Paragraph -- Session 10";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "Information Report",
    "Write Your Body Paragraph",
    "Session 10  |  Week 2  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // =========================================================================
  // SLIDE 2 -- LI / SC
  // =========================================================================
  liSlide(
    pres,
    [
      "We are learning to write a body paragraph for an information report by converting our plan into polished writing with appropriate language features",
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
  // SLIDE 3 -- Quick Revision: Structure and Features
  // =========================================================================
  contentSlide(
    pres,
    "Quick Revision",
    C.SECONDARY,
    "Body Paragraph -- Structure and Features",
    [
      "Structure: Topic Sentence -> Supporting Details -> Concluding Sentence",
      "Present tense -- information reports describe facts as they are",
      "Third person pronouns -- they, their, people, workers, children",
      "Appositives -- add detail about a noun (Session 6)",
      "Relative clauses -- add information using who, which, that (Session 7)",
      "Aim: include at least one appositive or relative clause in your paragraph",
    ],
    NOTES_REVISION,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- I Do: Model Writing
  // =========================================================================
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "From Plan to Paragraph",
    "My SPO plan:\nTS: Industrial Revolution shaped life\nDetail 1: Factories replaced craftsmen\nDetail 2: Cities grew, people moved\nDetail 3: Poor living conditions\nCS: New opportunities + hardship\n\nI expand each note into a full\nsentence, weaving in language\nfeatures as I go.",
    "My written paragraph:\n\n\"Life in 18th century England was shaped by the dramatic changes of the Industrial Revolution. Factories, which replaced traditional craftsmen, transformed the way people worked. Cities grew rapidly as workers, many of them former farm labourers, moved to find employment. Living conditions were often poor, with overcrowding, pollution and disease. These dramatic changes created both new opportunities and significant hardship.\"",
    NOTES_IDO_WRITE,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- CFU: Which sentence is more effective?
  // =========================================================================
  cfuSlide(
    pres,
    "CFU",
    "Which Is More Effective?",
    "Thumbs Up / Down",
    "A: \"People worked in factories. The factories were big.\"\n\nB: \"Workers, many of them children as young as five, laboured in factories that dominated the city skyline.\"\n\nThumbs UP for A, thumbs DOWN for B.\n\nWhich sentence is more effective for an information report?",
    NOTES_CFU,
    FOOTER
  );

  // =========================================================================
  // SLIDE 6 -- You Do: Write Your Paragraph
  // =========================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Write Your Body Paragraph");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.6, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Using your SPO plan from Session 9:", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.28,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0,
    });
    s.addText("First: Write your topic sentence\nNext: Expand each supporting detail into a full sentence\n         Try to include at least one appositive or relative clause\nThen: Write your concluding sentence", {
      x: 0.75, y: CONTENT_TOP + 0.44, w: 8.4, h: 1.00,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 1.74;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Writing Time: 15 minutes", {
      x: 0.75, y: tipY + 0.08, w: 3, h: 0.28,
      fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Your SPO is the skeleton. Your writing is the flesh\n- Expand notes into full, detailed sentences\n- Check: does each sentence connect to the one before it?\n- Use the mentor paragraph as a model if you need help", {
      x: 0.75, y: tipY + 0.38, w: 8.4, h: 0.80,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WRITE);
  }

  // =========================================================================
  // SLIDE 7 -- Edit and Proofread
  // =========================================================================
  contentSlide(
    pres,
    "Edit",
    C.ACCENT,
    "Proofread, Revise, Edit",
    [
      "Read your paragraph aloud quietly -- does it sound right?",
      "Use the editing checklist to check structure, language features, spelling and grammar",
      "Check commas around appositives and relative clauses",
      "Swap with a partner: give one positive comment and one suggestion",
      "Make final changes based on feedback",
    ],
    NOTES_EDIT,
    FOOTER
  );

  // =========================================================================
  // SLIDE 8 -- Closing
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
  // SLIDE 9 -- Resources
  // =========================================================================


  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: Editing Checklist ---------------------------------------------
  const cl = createPdf({ title: CHECKLIST_RESOURCE.name });
  let clY = addPdfHeader(cl, "Body Paragraph Editing Checklist", {
    color: C.PRIMARY,
    subtitle: "Information Report: 18th Century England",
    lessonInfo: "Session 10 | Week 2 | Year 5/6 Literacy",
    showNameDate: true,
  });

  clY = addTipBox(cl, "Use this checklist to proofread and edit your body paragraph. Check each item and tick when done. Then swap with a partner for peer feedback.", clY, { color: C.PRIMARY });

  clY = addSectionHeading(cl, "Structure", clY, { color: C.PRIMARY });
  clY = addBodyText(cl, "__ My paragraph has a clear topic sentence that introduces the main idea", clY);
  clY = addBodyText(cl, "__ I have at least 3 supporting detail sentences that expand on my topic sentence", clY);
  clY = addBodyText(cl, "__ My supporting details come from the non-fiction article (factual, not made up)", clY);
  clY = addBodyText(cl, "__ My concluding sentence wraps up without repeating the topic sentence", clY);
  clY += 8;

  clY = addSectionHeading(cl, "Language Features", clY, { color: C.SECONDARY });
  clY = addBodyText(cl, "__ I used present tense (is, are, has, makes -- not was, were, had, made)", clY);
  clY = addBodyText(cl, "__ I used third person pronouns (they, their, people -- not I, we, you)", clY);
  clY = addBodyText(cl, "__ I included at least one appositive OR relative clause to add detail", clY);
  clY = addBodyText(cl, "__ I used specific vocabulary (not vague words like \"stuff\", \"things\", \"a lot\")", clY);
  clY += 8;

  clY = addSectionHeading(cl, "Spelling, Grammar, Punctuation", clY, { color: C.ACCENT });
  clY = addBodyText(cl, "__ I checked spelling of key vocabulary (Industrial Revolution, factories, craftsmen, etc.)", clY);
  clY = addBodyText(cl, "__ Commas are correct around appositives (extra info between commas)", clY);
  clY = addBodyText(cl, "__ Commas are correct around non-essential relative clauses", clY);
  clY = addBodyText(cl, "__ No commas around essential relative clauses (tells us WHICH one)", clY);
  clY = addBodyText(cl, "__ Every sentence starts with a capital letter and ends with a full stop", clY);
  clY += 8;

  clY = addSectionHeading(cl, "Cohesion", clY, { color: C.ALERT });
  clY = addBodyText(cl, "__ Each sentence connects logically to the one before it", clY);
  clY = addBodyText(cl, "__ I read my paragraph aloud and it flows smoothly", clY);
  clY = addBodyText(cl, "__ My paragraph makes sense to someone who has not read my SPO plan", clY);
  clY += 12;

  clY = addSectionHeading(cl, "Peer Feedback", clY, { color: C.PRIMARY });
  clY = addBodyText(cl, "Partner's name: ____________________", clY);
  clY += 4;
  clY = addBodyText(cl, "One positive comment:", clY, { fontSize: 10 });
  clY = addLinedArea(cl, clY, 2, { lineSpacing: 22 });
  clY += 4;
  clY = addBodyText(cl, "One suggestion for improvement:", clY, { fontSize: 10 });
  clY = addLinedArea(cl, clY, 2, { lineSpacing: 22 });

  addPdfFooter(cl, "Session 10 | Editing Checklist");

  // --- PDF 2: Mentor Body Paragraph -----------------------------------------
  const mp = createPdf({ title: MENTOR_RESOURCE.name });
  let mpY = addPdfHeader(mp, "Mentor Body Paragraph -- Annotated", {
    color: C.SECONDARY,
    subtitle: "Information Report: 18th Century England",
    lessonInfo: "Session 10 | Week 2 | Year 5/6 Literacy",
    showNameDate: false,
  });

  mpY = addTipBox(mp, "This is a model body paragraph showing the structure and language features of an information report. Use it as a reference when writing your own paragraph.", mpY, { color: C.SECONDARY });

  mpY = addSectionHeading(mp, "Model Body Paragraph", mpY, { color: C.PRIMARY });

  mpY = addBodyText(mp, "Life in 18th century England was shaped by the dramatic changes of the Industrial Revolution. Factories, which replaced traditional craftsmen, transformed the way people worked across the country. Cities grew rapidly as workers, many of them former farm labourers, moved to find employment in the new industries. Living conditions in these rapidly growing cities were often poor, with overcrowding, pollution and disease becoming widespread among the working class. These dramatic changes transformed English society, creating both new opportunities and significant hardship for ordinary people.", mpY, { fontSize: 12 });
  mpY += 14;

  mpY = addSectionHeading(mp, "Annotations -- Structure", mpY, { color: C.PRIMARY });
  mpY = addBodyText(mp, "Topic Sentence: \"Life in 18th century England was shaped by the dramatic changes of the Industrial Revolution.\" -- introduces the main idea (the Industrial Revolution's impact)", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Supporting Detail 1: Factories replaced craftsmen (evidence of change)", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Supporting Detail 2: Cities grew rapidly (consequence of change)", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Supporting Detail 3: Poor living conditions (impact of change)", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Concluding Sentence: \"These dramatic changes transformed English society...\" -- wraps up without repeating the TS", mpY, { fontSize: 10 });
  mpY += 10;

  mpY = addSectionHeading(mp, "Annotations -- Language Features", mpY, { color: C.SECONDARY });
  mpY = addBodyText(mp, "Relative clause: \"which replaced traditional craftsmen\" (non-essential, commas needed)", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Appositive: \"many of them former farm labourers\" (adds detail about the workers)", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Present tense would be used if describing current facts; this paragraph uses past tense because it describes historical events (acceptable for historical reports)", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Third person: \"people\", \"workers\", \"the working class\" (no I, we, you)", mpY, { fontSize: 10 });
  mpY = addBodyText(mp, "Specific vocabulary: \"Industrial Revolution\", \"craftsmen\", \"employment\", \"overcrowding\"", mpY, { fontSize: 10 });

  addPdfFooter(mp, "Session 10 | Mentor Body Paragraph -- TEACHER AND STUDENT REFERENCE");

  // --- Write all files ----------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Session10.pptx` }),
    writePdf(cl, CHECKLIST_PDF_PATH),
    writePdf(mp, MENTOR_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Session10.pptx`);
  console.log("Done: " + CHECKLIST_RESOURCE.name);
  console.log("Done: " + MENTOR_RESOURCE.name);
}

build().catch(console.error);
