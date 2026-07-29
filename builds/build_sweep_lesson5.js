"use strict";

// Sweep Unit -- Lesson 5: Write Info Report Introduction
// Week 1, Session 5, Grade 5/6 Literacy
// Writing-focused lesson: write the introduction of an information report
// Cross-curricular: HASS (Convict Settlement in Australia)

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const T = createTheme("literacy", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  addTextOnShape, addInstructionCard, withReveal,
  titleSlide, liSlide, contentSlide,
  cfuSlide, closingSlide,
  vocabSlide, quoteSlide, modellingSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  addWriteLine,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 5;
const FOOTER = "Sweep | Lesson 5 | Week 1 | Year 5/6 Literacy";
const OUT_DIR = "output/Sweep_Lesson5_Info_Report_Intro";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const WRITING_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Introduction Writing Sheet",
  "Student writing sheet: draft, revise, and edit the introduction for the information report."
);
const CHECKLIST_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Editing Checklist",
  "Student self-assessment checklist for proofreading and editing the introduction."
);
const RESOURCE_ITEMS = [WRITING_RESOURCE, CHECKLIST_RESOURCE];
const WRITING_PDF_PATH = path.join(OUT_DIR, WRITING_RESOURCE.fileName);
const CHECKLIST_PDF_PATH = path.join(OUT_DIR, CHECKLIST_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Our final lesson for Week 1. Today is a writing lesson
- We are writing the introduction for our information report on Convict Settlement in Australia
- Last lesson you planned your title and General statement, and took notes from the article. Today you will draft the full introduction and then edit it

DO:
- Display title slide as students settle
- Ask students to have their Session 4 Planning Sheet on their desks
- Have the non-fiction articles available for reference

TEACHER NOTES:
Session 5 is a writing-focused lesson with no new novel reading. Students build on the planning work from Lesson 4. The lesson follows a write-revise-edit cycle. The non-fiction articles (British Convicts to Australia, First Fleet resources) should be accessible for reference during writing.

WATCH FOR:
- Students who did not complete their planning sheet -- they will need extra support at the start
- Students who are ready to write immediately -- let them begin once the structure is revised

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Today is all about writing. We are drafting and editing the introduction for our information report
- Read the success criteria. SC1 is about using the GST format we learned last lesson. SC2 is about using the language features of an information report. SC3 is about proofreading and editing your work

DO:
- Choral read the LI, then each SC
- Briefly preview: "Today you will write, then revise, then edit. Three stages"

TEACHER NOTES:
SC1 targets the GST structure (from Lesson 4). SC2 targets language features (present tense, third person, precise vocabulary, noun groups). SC3 targets the editing process. SC1 is the floor (using the structure). SC2 extends to language precision. SC3 is the quality-assurance skill.

WATCH FOR:
- Students who think editing means rewriting everything -- "Editing is targeted. Fix spelling, check grammar, improve one or two word choices"
- Students who skip straight to writing without reviewing their plan -- "Start by reviewing your plan from last lesson"

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_REVISE_STRUCTURE = `SAY:
- Before we write, let's quickly revise what an information report looks like. We covered this last lesson
- Purpose: classify or describe factual information
- Structure: Title, Introduction, Body Paragraphs, Conclusion
- Our focus today: the Introduction, using the GST format
- G: General statement -- big picture context. S: Specific statement -- zooms in on the topic. T: Thesis statement -- tells the reader what the report will discuss
- Key features to use in your writing: present tense, third person pronouns, precise vocabulary, noun groups that add detail

DO:
- Display the structure reminder briefly
- Point to each feature
- Emphasise: "These features need to appear in YOUR introduction today"

TEACHER NOTES:
This is a brief revision slide, not a full reteach. Students learned the structure in Lesson 4. The purpose is to reactivate the GST framework and the key features before they write. Keep this under 3 minutes.

WATCH FOR:
- Students who have forgotten the GST format -- direct them to review their planning sheet from Lesson 4
- Students who are confident and ready to write -- let them begin during the We Do if they are clearly ready

[General: Revision | VTLM 2.0: Retention and Recall]`;

const NOTES_NONFICTION = `SAY:
- Before writing, let's build more background knowledge. We have a second non-fiction article today about the First Fleet arriving in Sydney Cove
- As I read key sections, add to your notes from last lesson. Focus on facts that could strengthen your introduction
- Listen for: when the First Fleet sailed, how many ships, who was on board, why they came, what they found when they arrived

DO:
- Display or distribute the second non-fiction article
- Read key sections aloud (or have a strong reader read)
- Pause at key facts: "Write that one down -- it could go in your S or T statement"
- Allow 3-4 minutes for reading and note-taking

TEACHER NOTES:
This article adds depth to the planning work from Lesson 4. Students should be adding to their existing notes, not starting fresh. The First Fleet article provides specific detail (dates, numbers, names) that strengthens the Specific and Thesis statements. Keep this to 5 minutes maximum -- the bulk of the lesson is writing.

WATCH FOR:
- Students who are copying sentences from the article -- "Write the fact in YOUR words. Do not copy the author's sentence"
- Students who find strong specific details (e.g. 11 ships, 1,500 people) -- these strengthen the S statement

[General: Reading for Information | VTLM 2.0: Building Knowledge]`;

const NOTES_IDO = `SAY:
- Watch me write my introduction. I am using my plan from last lesson and the notes I just added
- First, I check my plan. My G statement: "In the late 18th century, Britain faced a crisis in its overcrowded prisons." Good -- but I can make it stronger. Let me add a noun group: "In the late 18th century, Britain faced a growing crisis in its dangerously overcrowded prisons and a society riddled with poverty"
- See what I did? I added "growing" and "dangerously overcrowded" and extended with "a society riddled with poverty." More detail, more precision
- Now my S statement. I check: is it in present tense? No -- let me adjust. "The British government establishes a penal colony in New South Wales, transporting thousands of convicted criminals across the world to the distant shores of Australia"
- My T statement: "This report examines life in 18th century England, the purpose and voyage of the First Fleet, the arrival in Sydney Cove, daily life for convicts in early Sydney, and the devastating impact on First Nations Australians"
- Final check: is it third person? Yes. Present tense? Yes. Precise vocabulary? "Penal colony," "convicted criminals," "devastating impact." Yes

DO:
- Display the planning sheet on the left, the written introduction on the right
- Model revising G aloud -- show the before and after
- Model checking S for tense
- Model checking T for specificity
- Think aloud at every decision point

MISCONCEPTIONS:
- Misconception: An introduction just states the topic: "This report is about convicts"
  Why: Students treat the introduction as a label rather than a structured entry point
  Impact: Introductions lack context, detail, and reader engagement
  Quick correction: "An introduction does three jobs: sets the scene (G), identifies the specific topic (S), and maps the report (T). 'This report is about convicts' does only part of one job. Where is the context? Where is the roadmap?"

TEACHER NOTES:
The I Do models the REVISION process as much as the writing process. Students see a plan transformed into polished writing through deliberate word choice improvements. The think-aloud focuses on the decision points: adding noun groups, checking tense, ensuring specificity. This is the key teaching moment for SC2 (language features).

WATCH FOR:
- Students who see the revision and think they need to rewrite their entire plan -- "Small changes make big differences. You do not need to start from scratch"
- Students who notice the noun group additions -- affirm: "That is exactly the kind of detail that makes information reports precise"

[General: I Do -- Writing Introduction | VTLM 2.0: Explicit Modelling]`;

// The CFU carries its own notes. It previously reused NOTES_WEDO, which left
// two consecutive slides with byte-identical notes - the teacher clicks on and
// the notes never change.
const NOTES_CFU_G = `SAY:
- Two General statements on the board. One is much stronger than the other
- Read both to yourself. Decide which one does the job better
- Everyone votes. Show me 1 finger or 2 fingers on my signal

DO:
- Give 15 seconds of silent thinking before the vote
- "Fingers at your chest... show me." Scan the room before revealing
- Expect 2. Cold call one student: "What makes Statement 2 stronger?"
- Reveal the answer bar only after every hand is up

CFU CHECKPOINT:
Technique: Finger Voting

Script:
- "Show me: 1 or 2" [Expect 2]
- "Why is Statement 2 stronger?" [specific time period, precise vocabulary, noun groups, context]

PROCEED:
- 80%+ voting 2 -> move to independent writing

PIVOT:
- Most likely misconception: thinking longer automatically means stronger
- Reteach: underline the precise words in Statement 2, not its length
- Re-check: read a third short-but-precise statement and re-vote`;

const NOTES_WEDO = `SAY:
- Let's work on your General statements together before you write independently
- Take out your planning sheet from last lesson. Look at your G statement
- With your partner, read your G statement aloud. Check: does it set the big picture context? Is it in present tense? Does it use precise vocabulary?
- Now revise it. Can you add a noun group to make it more detailed? Can you make the vocabulary more precise?
- You have 2 minutes to revise your G statement with your partner

DO:
- Give 2 minutes for partner revision of G statements
- Circulate and check for present tense and context
- Take 2-3 revised G statements and discuss as a class
- Model improving one weak example live: "Britain had problems" becomes "In the late 1700s, Britain grappled with overflowing prisons and rampant poverty that left thousands of its citizens desperate"

CFU CHECKPOINT:
Technique: Choral Response

Script:
- "I am going to read two G statements. Vote for the stronger one by showing 1 finger or 2 fingers"
- "Statement 1: 'In the old days, Britain had lots of convicts.'"
- "Statement 2: 'In the late 18th century, Britain faced a dire crisis as its overcrowded prisons could no longer contain the growing number of convicted criminals.'"
- "Show me: 1 or 2" [Expect 2]
- "Why is Statement 2 stronger?" [It has specific time period, precise vocabulary, noun groups, context]

PROCEED (>=80%): Students identify Statement 2 and explain why. Move to You Do.
PIVOT (<80%): Most likely issue -- students are unsure what makes a G statement "strong." Reteach: "Three tests for a strong G statement: (1) Does it tell me WHEN? (2) Does it tell me the SITUATION? (3) Does it use precise vocabulary, not vague words like 'lots' or 'things'? Check your own G statement against these three tests." Re-check: "Read your G statement to your partner. Do they hear when, situation, and precise vocabulary?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the sentence frame to write the G statement: "In the late 1700s, Britain struggled with ___ which led to ___." Focus on completing the frame with factual information from the article
- Extra Notes: These students should write G and S only. The T statement can be drafted with teacher support

EXTENDING PROMPT:
- Task: Draft the complete GST introduction, then revise it to include at least two noun groups and one adverbial phrase. Check all three sentences for present tense and third person

TEACHER NOTES:
The We Do focuses on revising G statements only -- isolating one component for quality. The full writing happens in You Do. The comparison CFU makes the quality criteria concrete: students can SEE the difference between a weak and strong G statement.

WATCH FOR:
- Students who use vague language ("a long time ago," "lots of") -- push for precision
- Students who add strong noun groups during revision -- affirm and share

[General: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_YOUDO = `SAY:
- Now write your full introduction on the writing sheet
- FIRST: Review your plan from last lesson. Revise your G, S, and T using what we practised today
- NEXT: Write the introduction in full sentences -- G, then S, then T. Check for present tense, third person, and precise vocabulary
- THEN: Use the editing checklist to proofread and edit. Read your introduction aloud quietly to check it flows

DO:
- Distribute the Session 5 Introduction Writing Sheet and the Editing Checklist
- Students should have their Session 4 Planning Sheet for reference
- Circulate -- check for GST structure, present tense, third person
- After 8 minutes: "You should be starting your editing now"
- Conference with students who are stuck on S or T

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Write G and S statements using the sentence frames on the writing sheet. The T statement can be completed orally with teacher support and then written down
- Extra Notes: Focus on getting the ideas down first. Editing for features comes after the draft is complete

EXTENDING PROMPT:
- Task: After writing and editing the introduction, write the first body paragraph about life in 18th century England. Use a subheading, topic sentence, and 2-3 supporting details from the non-fiction articles

TEACHER NOTES:
The You Do is the main writing block. Students draft the full GST introduction and then self-edit using the checklist. The editing checklist targets the specific features taught in this unit: GST structure, present tense, third person, noun groups, precise vocabulary. The extending prompt invites students to begin the first body paragraph.

WATCH FOR:
- Students who write in past tense -- "Information reports use present tense. Check each verb"
- Students who skip the editing checklist -- "The checklist is part of the task. It makes your writing stronger"
- Students who produce strong, well-structured introductions -- share with the class as mentor texts

[General: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: I can write an introduction using GST format -- thumbs? [scan]
- SC2: I can use language features of an information report in my introduction -- thumbs? [scan]
- SC3: I can proofread and edit my introduction to improve language and correctness -- thumbs? [scan]
- Turn and talk: Read your introduction to your partner. Give them one specific piece of feedback: what is the strongest part of their introduction?

DO:
- Run through each SC with thumbs check
- The turn-and-talk provides peer feedback -- positive, targeted
- Preview: "Next week we continue the Sweep novel and write the body paragraphs for our information report"

TEACHER NOTES:
The closing combines self-assessment with peer feedback. Positive peer feedback at this stage builds confidence for the longer writing tasks ahead. Students showing thumbs-down on SC1 may need a small-group reteach on GST before the body paragraph writing begins.

WATCH FOR:
- Students who give vague feedback ("it's good") -- redirect: "Tell them specifically: which sentence is strongest? Why?"
- Students thumbs-down on SC2 -- the language features need more practice. Build this into future writing lessons

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${WRITING_RESOURCE.name} has structured space for your GST introduction draft
- The ${CHECKLIST_RESOURCE.name} helps you proofread and edit

DO:
- Print both resources before the lesson (one per student)
- Students should keep their completed introductions for next week

TEACHER NOTES:
The writing sheet and checklist are designed to be used together. The checklist targets the specific features of information report writing taught in this unit.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Sweep - Lesson 5 - Write Info Report Introduction";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "Sweep",
    "Writing -- Information Report Introduction",
    "Lesson 5  |  Week 1  |  Year 5/6 Literacy",
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
  // SLIDE 2 -- LI/SC
  // =========================================================================
  liSlide(
    pres,
    [
      "We are learning to write and edit the introduction of an information report using the GST format and key language features",
    ],
    [
      "I can write an introduction for an information report using GST format",
      "I can use language features of an information report such as present tense, third person, and precise vocabulary",
      "I can proofread and edit my introduction to improve language, spelling, and cohesion",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Revise Structure
  // =========================================================================
  contentSlide(
    pres,
    "Revision",
    C.SECONDARY,
    "Information Report -- Quick Revision",
    [
      "Purpose: classify or describe factual information",
      "Structure: Title -> Introduction (GST) -> Body Paragraphs -> Conclusion",
      "Today's focus: writing the Introduction using GST",
      "Key features: present tense, third person, precise vocabulary, noun groups",
    ],
    NOTES_REVISE_STRUCTURE,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Non-Fiction Reading
  // =========================================================================
  contentSlide(
    pres,
    "Build Knowledge",
    C.PRIMARY,
    "Non-Fiction Article -- First Fleet",
    [
      "Listen and add to your notes from last lesson",
      "Focus on facts that strengthen your introduction:",
      "When did the First Fleet sail? How many ships?",
      "Who was on board? Why did they come?",
      "What did they find when they arrived?",
    ],
    NOTES_NONFICTION,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- I Do: Write Introduction
  // =========================================================================
  modellingSlide(
    pres,
    "I Do",
    "Writing the Introduction",
    "From plan to polished writing:\n\n1. Check your plan\n2. Revise for precision\n3. Add noun groups\n4. Check: present tense?\n5. Check: third person?\n6. Check: precise vocabulary?",
    "G: \"In the late 18th century,\nBritain faced a growing crisis\nin its dangerously overcrowded\nprisons and a society riddled\nwith poverty.\"\n\nS: \"The British government\nestablishes a penal colony in\nNew South Wales, transporting\nthousands of convicted criminals\nacross the world.\"\n\nT: \"This report examines 18th\ncentury England, the First Fleet,\narrival in Sydney Cove, convict\nlife, and the impact on First\nNations Australians.\"",
    NOTES_IDO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 6 -- We Do: Revise G Statements
  // =========================================================================
  contentSlide(
    pres,
    "We Do",
    C.SUCCESS,
    "Revise Your General Statement",
    [
      "Take out your planning sheet from Lesson 4",
      "Read your G statement aloud to your partner",
      "Check: does it set context? Present tense? Precise vocabulary?",
      "Revise: can you add a noun group to make it more detailed?",
      "You have 2 minutes to revise together",
    ],
    NOTES_WEDO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 7 -- CFU: Compare G Statements
  // =========================================================================
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Which General Statement Is Stronger?",
      "Finger Voting -- Show 1 or 2",
      "Statement 1:\n\"In the old days, Britain had lots of convicts.\"\n\nStatement 2:\n\"In the late 18th century, Britain faced a dire crisis as its overcrowded prisons could no longer contain the growing number of convicted criminals.\"",
      NOTES_CFU_G,
      FOOTER
    ),
    (slide) => {
      const ansY = 4.20;
      slide.addShape("roundRect", {
        x: 0.5, y: ansY, w: 9, h: 0.75, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      });

      slide.addText("Statement 2 is stronger: specific time period, precise vocabulary, detailed noun groups, clear context", {
        x: 0.7, y: ansY + 0.06, w: 8.6, h: 0.60,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_G);
    }
  );

  // =========================================================================
  // SLIDE 8 -- You Do: Write Introduction
  // =========================================================================
  contentSlide(
    pres,
    "You Do",
    C.ACCENT,
    "Write Your Introduction",
    [
      "FIRST: Review your plan from Lesson 4. Revise G, S, and T",
      "NEXT: Write the full introduction on the writing sheet. Check for present tense, third person, precise vocabulary",
      "THEN: Use the editing checklist to proofread and edit. Read aloud to check flow",
    ],
    NOTES_YOUDO,
    FOOTER,
    (slide, layoutGuide) => {
      const tipY = layoutGuide.panelTopPadded + 0.1;
      addCard(slide, layoutGuide.rightX, tipY, layoutGuide.rightW, 2.3, {
        strip: C.SECONDARY, fill: C.BG_CARD,
      });
      slide.addText("Feature Check", {
        x: layoutGuide.rightX + 0.15, y: tipY + 0.08, w: 3.8, h: 0.28,
        fontSize: 11, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Present tense (examines, not examined)", options: { bullet: true, breakLine: true, fontSize: 12, color: C.CHARCOAL } },
        { text: "Third person (the government, not we)", options: { bullet: true, breakLine: true, fontSize: 12, color: C.CHARCOAL } },
        { text: "Precise vocabulary (penal colony, not jail)", options: { bullet: true, breakLine: true, fontSize: 12, color: C.CHARCOAL } },
        { text: "Noun groups that add detail", options: { bullet: true, fontSize: 12, color: C.CHARCOAL } },
      ], {
        x: layoutGuide.rightX + 0.15, y: tipY + 0.44, w: 3.8, h: 1.6,
        fontFace: FONT_B, valign: "top", margin: 0, paraSpaceAfter: 5,
      });
    }
  );

  // =========================================================================
  // SLIDE 9 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    "Read your introduction to your partner. Give them one specific piece of feedback: what is the strongest part of their introduction?",
    [
      "I can write an introduction for an information report using GST format",
      "I can use language features of an information report such as present tense, third person, and precise vocabulary",
      "I can proofread and edit my introduction to improve language, spelling, and cohesion",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 10 -- Resources
  // =========================================================================


  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: Introduction Writing Sheet ------------------------------------
  const ws = createPdf({ title: WRITING_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Information Report Introduction", {
    color: C.NAVY,
    subtitle: "Convict Settlement in Australia",
    lessonInfo: "Sweep | Lesson 5 | Week 1 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "Use your planning sheet from Lesson 4 and your notes from the non-fiction articles.\nYour introduction has three parts: G (General), S (Specific), T (Thesis).\nCheck: present tense, third person, precise vocabulary, noun groups.", wsY, { color: C.NAVY });

  // Title
  wsY = addSectionHeading(ws, "Title of Your Information Report", wsY, { color: C.NAVY });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 30 });
  wsY += 8;

  // G Statement
  wsY = addSectionHeading(ws, "G: General Statement", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Sets the big picture context. Tells the reader about the world that led to this topic.", wsY, { fontSize: 9, italic: true });
  wsY = addBodyText(ws, "Frame (optional): \"In the late ___ century, Britain ___\"", wsY, { fontSize: 9, italic: true });
  wsY = addLinedArea(ws, wsY, 3, { lineSpacing: 26 });
  wsY += 6;

  // S Statement
  wsY = addSectionHeading(ws, "S: Specific Statement", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Zooms in on the topic directly. Uses key terms and addresses the subject.", wsY, { fontSize: 9, italic: true });
  wsY = addBodyText(ws, "Frame (optional): \"The British government ___\"", wsY, { fontSize: 9, italic: true });
  wsY = addLinedArea(ws, wsY, 3, { lineSpacing: 26 });
  wsY += 6;

  // T Statement
  wsY = addSectionHeading(ws, "T: Thesis Statement", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Tells the reader what the body paragraphs will discuss. The roadmap for the report.", wsY, { fontSize: 9, italic: true });
  wsY = addBodyText(ws, "Frame (optional): \"This report examines ___\"", wsY, { fontSize: 9, italic: true });
  wsY = addLinedArea(ws, wsY, 3, { lineSpacing: 26 });
  wsY += 8;

  // Final draft space
  wsY = addSectionHeading(ws, "Final Draft -- Write your full introduction below", wsY, { color: C.ACCENT });
  wsY = addLinedArea(ws, wsY, 8, { lineSpacing: 26 });

  addPdfFooter(ws, "Sweep | Lesson 5 | Introduction Writing Sheet -- Keep for next week");

  // --- PDF 2: Editing Checklist ---------------------------------------------
  const cl = createPdf({ title: CHECKLIST_RESOURCE.name });
  let clY = addPdfHeader(cl, "Editing Checklist", {
    color: C.ALERT,
    subtitle: "Information Report Introduction",
    lessonInfo: "Sweep | Lesson 5 | Week 1 | Year 5/6 Literacy",
    showNameDate: true,
  });

  clY = addTipBox(cl, "Read your introduction aloud quietly. Then check each item below.\nTick the box if your introduction meets the criterion. If not, revise before submitting.", clY, { color: C.ALERT });

  clY = addSectionHeading(cl, "Structure (GST)", clY, { color: C.NAVY });
  const structureItems = [
    "My introduction has a General statement that sets the big picture context",
    "My introduction has a Specific statement that zooms in on the topic using key terms",
    "My introduction has a Thesis statement that tells the reader what the report will discuss",
    "My G, S, and T flow logically -- each sentence leads naturally to the next",
  ];
  for (const item of structureItems) {
    clY = addBodyText(cl, `[ ]  ${item}`, clY, { fontSize: 10 });
    clY += 2;
  }
  clY += 4;

  clY = addSectionHeading(cl, "Language Features", clY, { color: C.SECONDARY });
  const featureItems = [
    "I have used present tense (examines, establishes, faces -- not examined, established, faced)",
    "I have used third person pronouns (the government, Britain, convicts -- not we, I, you)",
    "I have used precise vocabulary (penal colony, transported, convicted -- not jail, sent, bad)",
    "I have used at least one noun group to add detail (e.g. 'dangerously overcrowded prisons')",
    "I have used at least one adverbial to add context (e.g. 'In the late 18th century')",
  ];
  for (const item of featureItems) {
    clY = addBodyText(cl, `[ ]  ${item}`, clY, { fontSize: 10 });
    clY += 2;
  }
  clY += 4;

  clY = addSectionHeading(cl, "Proofreading", clY, { color: C.ACCENT });
  const proofItems = [
    "I have checked spelling of key words (Australia, colony, government, criminals)",
    "I have checked capital letters (proper nouns: Britain, Australia, New South Wales, First Fleet)",
    "I have checked punctuation (full stops, commas in lists)",
    "I have read my introduction aloud to check it sounds right",
  ];
  for (const item of proofItems) {
    clY = addBodyText(cl, `[ ]  ${item}`, clY, { fontSize: 10 });
    clY += 2;
  }
  clY += 6;

  clY = addSectionHeading(cl, "Peer Feedback", clY, { color: C.NAVY });
  clY = addBodyText(cl, "Partner's name: ___________________", clY, { fontSize: 10 });
  clY += 2;
  clY = addBodyText(cl, "The strongest part of my partner's introduction is:", clY, { fontSize: 10 });
  clY = addLinedArea(cl, clY, 2, { lineSpacing: 26 });
  clY += 2;
  clY = addBodyText(cl, "One suggestion for improvement:", clY, { fontSize: 10 });
  clY = addLinedArea(cl, clY, 2, { lineSpacing: 26 });

  addPdfFooter(cl, "Sweep | Lesson 5 | Editing Checklist");

  // --- Write all files ------------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Sweep_Lesson5.pptx` }),
    writePdf(ws, WRITING_PDF_PATH),
    writePdf(cl, CHECKLIST_PDF_PATH),
  ]);

  console.log(`PPTX written to ${OUT_DIR}/Sweep_Lesson5.pptx`);
  console.log(`Done: ${WRITING_RESOURCE.name}`);
  console.log(`Done: ${CHECKLIST_RESOURCE.name}`);
}

build().catch(console.error);
