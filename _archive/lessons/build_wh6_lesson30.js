"use strict";

// War Horse Unit -- Lesson 30: Concluding Paragraph + Edit Full Persuasive Text
// Week 6, Session 7, Grade 5/6 Literacy
// Indigenous Australian WW1 reading + concluding paragraph (TSG) + full text editing

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  addTextOnShape, withReveal,
  titleSlide, liSlide, contentSlide,
  cfuSlide, taskSlide, closingSlide,
  vocabSlide, quoteSlide, modellingSlide,
} = require("../themes/wh6_helpers");

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  addStepInstructions,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 7;
const FOOTER = "War Horse | Lesson 30 | Week 6 | Year 5/6 Literacy";
const OUT_DIR = "output/WH6_Lesson30_Concluding_Paragraph";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLANNER_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Concluding Paragraph Planner",
  "Planning template for the concluding paragraph using the TSG structure."
);
const CHECKLIST_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Editing Checklist",
  "Student checklist for editing the complete persuasive text."
);
const RESOURCE_ITEMS = [PLANNER_RESOURCE, CHECKLIST_RESOURCE];
const PLANNER_PDF_PATH = path.join(OUT_DIR, PLANNER_RESOURCE.fileName);
const CHECKLIST_PDF_PATH = path.join(OUT_DIR, CHECKLIST_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Final writing session for our persuasive text today
- First we read about Indigenous Australians in World War 1
- Then we write our concluding paragraph and edit our complete persuasive text

DO:
- Display title slide as students settle
- Have the Britannica Kids article ready on devices or printed
- Students should have all their previous writing drafts (introduction, body paragraphs 1 and 2)

TEACHER NOTES:
Session 7, the final session of the unit. This lesson combines the Indigenous Australian WW1 reading from L30 with the concluding paragraph and full-text editing. Students complete their persuasive text by the end of this session. The cultural sensitivity advisory must be read before the lesson begins.

WATCH FOR:
- Students who do not have their previous drafts -- they will need them for the editing phase
- Students who are ready to finish -- this is the culmination of several sessions of writing

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Today we are completing our persuasive text with a concluding paragraph, then editing the whole thing
- Read the success criteria. SC1 is about knowing what a concluding paragraph does. SC2 is about writing one. SC3 is about editing your complete text

DO:
- Choral read the LI, then the SCs
- Emphasise: "By the end of today, your persuasive text will be complete -- introduction, two body paragraphs, and a conclusion"

TEACHER NOTES:
SC1 is the foundation -- naming the features of a concluding paragraph. SC2 is the core target -- actually writing one. SC3 is the depth -- editing the full text with attention to language choices, structure, and effect. The editing phase is where students revise across all four components.

WATCH FOR:
- Students who are anxious about having to edit everything -- reassure: "We will use a checklist to make editing manageable"
- Students who are excited to finish -- channel that energy into quality, not speed

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_SENSITIVITY = `SAY:
- Before we begin reading, an important acknowledgement
- Today's reading explores Aboriginal and Torres Strait Islander histories during World War 1
- We acknowledge Aboriginal and Torres Strait Islander Peoples as the First Nations of Australia
- Some material may contain words or descriptions that reflect the views of the time period, which may not be considered appropriate today
- Aboriginal and Torres Strait Islander Peoples should be aware that some material may contain images, voices, and names of deceased persons
- If anything in the reading is difficult or uncomfortable, that is okay -- you can speak to me privately

DO:
- Read the acknowledgement with sincerity and gravity
- Pause briefly after delivering it
- Check in visually with any Aboriginal or Torres Strait Islander students in the class

SENSITIVITY ADVISORY:
- What it is: Content about Aboriginal and Torres Strait Islander peoples' experiences during WW1, including discrimination in enlistment and lack of acknowledgement of service
- Framing language: "This reading shows us part of Australia's history that is important and sometimes painful. We read it with respect and with the understanding that these are real people's stories"
- Watch for: Students who become uncomfortable, particularly Aboriginal and Torres Strait Islander students. Students who make insensitive comments about historical treatment
- Protocol: Address insensitive comments immediately and calmly. If a student is distressed, acknowledge privately and offer to step outside for a moment. Follow school cultural safety protocols

TEACHER NOTES:
The cultural sensitivity framing is from the lesson plan and must be delivered before the reading begins. This is not optional. The reading about Indigenous Australians in WW1 includes content about discrimination and lack of recognition -- handle with respect and care.

WATCH FOR:
- Students who have questions about why Indigenous Australians were treated differently -- address honestly and age-appropriately
- Students who connect this to contemporary issues of recognition -- these connections are valid and valuable

[General: Cultural Acknowledgement | VTLM 2.0: Culturally Responsive Teaching]`;

const NOTES_READING = `SAY:
- We are reading from Britannica Kids: Australia in World War 1
- Today's sections focus on Indigenous Australians and key battles
- Sections: Role of Indigenous Australians, Key Campaigns and Battles, Gallipoli Campaign, Anzac Troops, The Western Front
- We will also watch the ANZACs: Indigenous Australians video
- Take notes as you read -- this is your last chance to gather evidence for your persuasive text

DO:
- Direct students to the article on devices or distribute printed copies
- Play the Indigenous Australians video from the article
- Select readers for the text sections
- Monitor note-taking -- students should be looking for evidence they can use

TEACHER NOTES:
This reading broadens students' understanding of Australia's WW1 experience to include Indigenous perspectives. The content about discrimination in enlistment and lack of post-war recognition is historically important. Incidental vocabulary: peninsula, campaigns, enlistment, discrimination, acknowledged, empire.

SOURCES:
Australia in World War 1 - Britannica Kids
ANZACs: Indigenous Australians (embedded video)

WATCH FOR:
- Students who are surprised by the discrimination faced by Indigenous soldiers -- this is an important learning moment
- Students who take notes connecting this reading to their persuasive text -- they are thinking like writers

[General: Reading | VTLM 2.0: Structured Reading Practice]`;

const NOTES_MACRO_REVISION = `SAY:
- Quick revision of persuasive text macrostructure before we write our conclusion
- Read from slide: Introduction, Body Paragraphs, Concluding Paragraph
- You have written your introduction and two body paragraphs. Today: the concluding paragraph
- Purpose of a concluding paragraph: to sum up the arguments made and restate the author's position
- Features -- TSG structure: Thesis statement (restate your position), Specific statement (your strongest point), General statement (a big-picture takeaway or call to action)

DO:
- Display the macrostructure with the concluding paragraph highlighted
- Point to each TSG element as you explain it
- Emphasise: the conclusion does NOT introduce new evidence. It wraps up what you have already argued

CFU CHECKPOINT:
Technique: Choral Response

Script:
- "What is the purpose of a concluding paragraph?" [Sum up the arguments and restate the position]
- "What are the three parts of TSG?" [Thesis, Specific, General]
- "Does the conclusion introduce new evidence?" [No -- it wraps up existing arguments]
- Scan for: confident, accurate choral responses

PROCEED (>=80%): Strong choral responses. Move to modelling.
PIVOT (<80%): Most likely issue -- students confuse the conclusion with another body paragraph. Reteach: "A body paragraph makes an argument with evidence. A conclusion LOOKS BACK at the arguments you already made. Think of it as the closing speech in a court case -- you are not presenting new evidence, you are reminding the jury what you have already proved." Re-check: "What is the ONE thing a conclusion must NOT do?" [Introduce new evidence]

TEACHER NOTES:
The TSG structure is new -- Thesis, Specific, General. This is a clear, memorable framework for conclusions. The revision of macrostructure connects back to Session 4 where students first saw the persuasive text structure.

WATCH FOR:
- Students who think the conclusion is just repeating the introduction -- "The conclusion restates your position but in a stronger, more confident way. You have PROVEN your point now"
- Students who want to add new arguments -- "Save new ideas for a different text. The conclusion wraps up this one"

[General: I Do -- Revision | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_MODEL = `SAY:
- Watch me plan and write a concluding paragraph. I am using the TSG structure
- Thesis statement -- restate my position: "The contribution of horses to the war effort was undeniably significant and must be remembered."
- Specific statement -- my strongest point: "From transporting troops and supplies across impossible terrain to serving as ambulance carriers and companions for soldiers in their darkest hours, horses proved themselves essential at every stage of the conflict."
- General statement -- big picture: "Remembering the sacrifice of these animals is not just a matter of history -- it is a matter of justice. Their service deserves the same recognition we give to the soldiers who fought beside them."
- Notice: I restated my position (thesis), highlighted my strongest argument (specific), then zoomed out to a big-picture call to action (general)
- Notice the language: high modality ("undeniably," "must"), nominalisation ("contribution"), tier 2/3 vocabulary ("conflict," "sacrifice," "recognition")

DO:
- Display the model concluding paragraph with TSG labels
- Think aloud at each decision: "For my thesis, I am restating my position but with more confidence now because I have proved it"
- Point out the persuasive language techniques across the paragraph

TEACHER NOTES:
The modelled conclusion demonstrates all three persuasive language tools students have learned: modality (Session 4), nominalisations (Session 4), and tier 2/3 vocabulary (Session 6). The call to action in the general statement is a deliberate rhetorical choice -- it leaves the reader with something to think about or do.

WATCH FOR:
- Students who notice the persuasive language techniques being combined -- affirm: "That is exactly what strong persuasive writing does"
- Students who think the conclusion needs to be long -- "A strong conclusion can be short. Quality over quantity"

[General: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU = `SAY:
- Quick check. I will show you a concluding sentence. Tell me which TSG element it is
- "Animals were essential to the success of the war effort and their contribution must never be forgotten."
- On your whiteboards: write T, S, or G

DO:
- Display the sentence
- Allow 20 seconds for whiteboard responses
- Signal: boards up

CFU CHECKPOINT:
Technique: Show Me Boards

Script:
- "Write T for thesis, S for specific, or G for general. Boards up in 20 seconds."
- Scan for: T (this is a thesis statement -- it restates the position)
- Follow up: "Why is this a thesis and not a general statement?" [It states the POSITION directly -- what the author believes. A general statement would zoom out to a bigger idea or call to action]

PROCEED (>=80%): Most show T with clear reasoning. Move to We Do.
PIVOT (<80%): Most likely issue -- students confuse thesis with general because both sound broad. Reteach: "The thesis restates YOUR position -- what YOU are arguing. The general statement zooms OUT to something bigger than your essay. Thesis: 'Animals were essential.' General: 'Every nation should build memorials to the animals that served.' See the difference? The thesis is about your argument. The general is about the world." Re-check: "Is this T, S, or G: 'We owe it to future generations to remember these sacrifices'?" [G -- general, zooms out to future generations]

TEACHER NOTES:
This CFU checks TSG identification before students apply it. The follow-up question tests whether students understand the distinction between thesis (restating position) and general (big picture). This is the key conceptual threshold.

WATCH FOR:
- Students who write S -- they may be confusing thesis with specific. The thesis does not name a PARTICULAR argument, it restates the OVERALL position
- Students who are confident -- they are ready for independent planning

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_WEDO = `SAY:
- Now plan your concluding paragraph together. Look at your introduction and body paragraphs
- Your thesis statement: restate your position. How would you say your main argument in one strong sentence?
- Your specific statement: what was your STRONGEST argument? The one with the best evidence?
- Your general statement: zoom out. What is the big takeaway? Why should the reader care beyond your essay?
- Discuss with your partner: what will your TSG look like?
- Fill in the concluding paragraph planner

DO:
- Distribute the Session 7 Concluding Paragraph Planner
- Give 2 minutes for partner discussion
- Circulate -- check that thesis statements restate the position (not introduce new arguments)
- After 4 minutes: "You should have all three elements drafted"

CFU CHECKPOINT:
Technique: Cold Call

Script:
- After 3 minutes, cold call 2-3 students: "Read me your thesis statement."
- Listen for: a clear restatement of their position with confident language
- If a thesis introduces new evidence: "That sounds like a body paragraph point. Your thesis restates what you BELIEVE, not new proof"

PROCEED (>=80%): Most thesis statements clearly restate the position. Students continue planning.
PIVOT (<80%): Most likely issue -- students are writing new body paragraph arguments instead of restating their position. Reteach: "Think back to your introduction. What did you say your position was? Now say it again, but with more confidence. You have PROVED it now." Re-check: cold call 2 more students.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide the planner with sentence starters for each TSG element. T: "It is clear that..." S: "The strongest evidence for this is..." G: "This matters because..."
- Extra Notes: Students can look at their introduction to help restate their position

EXTENDING PROMPT:
- Task: After completing the planner, write the concluding paragraph AND add a rhetorical question as a final sentence that challenges the reader to take action or think differently

TEACHER NOTES:
The We Do is guided planning for the concluding paragraph. Students should have their previous writing in front of them -- the conclusion must connect to what they have already written, not stand alone.

WATCH FOR:
- Students whose general statement is too vague ("War is bad") -- push: "Can you make that more specific to YOUR argument about animals?"
- Students who are unsure what "zoom out" means -- "Think: what should the reader DO or THINK after reading your essay?"

[General: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_YOUDO = `SAY:
- Two tasks now. First: write your concluding paragraph from your planner
- Then the big finish: reread your COMPLETE persuasive text -- introduction, body paragraph 1, body paragraph 2, and conclusion
- Use the editing checklist to review your writing
- First: write your concluding paragraph using TSG
- Next: put all four parts together and reread from the beginning
- Then: use the editing checklist to check and fix your work

DO:
- Students write their concluding paragraph first (5-7 minutes)
- Then students arrange all parts of their persuasive text in order
- Distribute the Session 7 Editing Checklist
- Circulate -- the editing phase is where the real improvement happens
- After 10 minutes of editing: "Make your final changes"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Write the concluding paragraph using the sentence starters from the planner. For editing, focus on only the first three items on the checklist (structure, topic sentences, concluding sentences)
- Extra Notes: These students may benefit from reading their text aloud to a partner to hear errors

EXTENDING PROMPT:
- Task: After completing the full edit, write a brief self-assessment: "What is the strongest part of my persuasive text? What would I improve if I had more time?" Identify one example each of modality, nominalisation, and tier 2/3 vocabulary in their final text

TEACHER NOTES:
The You Do has two phases: writing the conclusion, then editing the full text. The editing checklist structures the review so students do not feel overwhelmed by rereading everything. This is the culmination of the persuasive writing unit -- students should leave with a complete, edited text.

WATCH FOR:
- Students who rush the editing -- "The editing is where good writing becomes great writing. Take your time"
- Students who only check spelling -- "The checklist covers structure, language, and argument. Spelling is just one part"
- Students who are proud of their finished text -- celebrate this. They have written a complete persuasive text over several sessions

[General: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: purpose and features of a concluding paragraph -- thumbs? [scan]
- SC2: planning and writing a concluding paragraph using TSG -- thumbs? [scan]
- SC3: editing your complete persuasive text -- thumbs? [scan]
- You have written a complete persuasive text: introduction, two body paragraphs, and a conclusion. That is a serious achievement
- Turn and talk: What is one thing you are most proud of in your persuasive text? What persuasive technique did you use best?

DO:
- Run through each SC with thumbs check
- Acknowledge the achievement -- this has been a multi-session writing project
- Collect all completed persuasive texts
- If time: 1-2 volunteers share their favourite sentence from their text

TEACHER NOTES:
The closing celebrates the completion of a sustained writing project. Students have learned and applied modality, nominalisations, tier 2/3 vocabulary, SPO planning, and TSG conclusions across four writing sessions. The turn-and-talk invites metacognition about their own growth as persuasive writers.

WATCH FOR:
- Students who are genuinely proud -- this should be the majority. Affirm their effort
- Students who feel their text is not good enough -- reassure: "Every draft can be improved. The fact that you FINISHED a complete persuasive text is the achievement"
- Students who can name specific persuasive techniques they used -- excellent metacognitive awareness

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two resources for today
- The ${PLANNER_RESOURCE.name} helps you plan your TSG concluding paragraph
- The ${CHECKLIST_RESOURCE.name} guides you through editing your complete persuasive text

DO:
- Print the planner before the lesson (one per student)
- Print the checklist before the lesson (one per student)
- Click any resource card to open the PDF

TEACHER NOTES:
The editing checklist is designed to be reusable for future persuasive writing tasks. Students may keep it in their writing folders as a reference tool.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "War Horse - Lesson 30 - Concluding Paragraph + Edit";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "War Horse",
    "Concluding Paragraph + Edit",
    "Lesson 30  |  Week 6  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  const theme = { C, FONT_H, FONT_B, addTopBar, addTitle, addFooter, addCard };
  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    theme,
    FOOTER,
    NOTES_RESOURCES
  );

  // =========================================================================
  // SLIDE 2 -- LI/SC
  // =========================================================================
  liSlide(
    pres,
    [
      "We are learning to plan and write a concluding paragraph for a persuasive text and to edit our complete persuasive writing",
    ],
    [
      "I can identify the purpose and features of a concluding paragraph in a persuasive text",
      "I can plan and write a concluding paragraph that restates my position and sums up my arguments",
      "I can edit my complete persuasive text for clarity, language choices, and effect",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Cultural Sensitivity Acknowledgement
  // =========================================================================
  contentSlide(
    pres,
    "Acknowledgement",
    C.PRIMARY,
    "Cultural Sensitivity",
    [
      "Today's reading explores Aboriginal and Torres Strait Islander histories during World War 1",
      "We acknowledge Aboriginal and Torres Strait Islander Peoples as the First Nations of Australia",
      "Some material may reflect views of the time period that are not considered appropriate today",
      "Aboriginal and Torres Strait Islander Peoples should be aware that material may contain images, voices, and names of deceased persons",
    ],
    NOTES_SENSITIVITY,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Reading Launch
  // =========================================================================
  contentSlide(
    pres,
    "Read Aloud",
    C.SECONDARY,
    "Indigenous Australians in World War 1",
    [
      "Britannica Kids: Australia in World War 1",
      "Sections: Role of Indigenous Australians, Key Campaigns and Battles",
      "Gallipoli Campaign, Anzac Troops, The Western Front",
      "Watch the ANZACs: Indigenous Australians video",
      "Take notes: what evidence could you use in your persuasive text?",
    ],
    NOTES_READING,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- Persuasive Macrostructure + Concluding Paragraph Features
  // =========================================================================
  modellingSlide(
    pres,
    "I Do",
    "The Concluding Paragraph -- TSG",
    "Persuasive text macrostructure:\n1. Introduction\n2. Body Paragraphs (done!)\n3. Concluding Paragraph (today)\n\nPurpose: sum up the arguments and restate the author's position\n\nThe conclusion does NOT introduce new evidence",
    "TSG Structure:\n\nT - Thesis Statement\nRestate your position with confidence\n\nS - Specific Statement\nHighlight your strongest argument\n\nG - General Statement\nBig-picture takeaway or call to action",
    NOTES_MACRO_REVISION,
    FOOTER
  );

  // =========================================================================
  // SLIDE 6 -- Teacher Models Concluding Paragraph
  // =========================================================================
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Model Concluding Paragraph",
    [
      "T: \"The contribution of horses to the war effort was undeniably significant and must be remembered.\"",
      "S: \"From transporting troops across impossible terrain to serving as ambulance carriers and companions, horses proved themselves essential at every stage of the conflict.\"",
      "G: \"Remembering the sacrifice of these animals is not just a matter of history -- it is a matter of justice. Their service deserves the same recognition we give to the soldiers who fought beside them.\"",
    ],
    NOTES_MODEL,
    FOOTER
  );

  // =========================================================================
  // SLIDES 7-8 -- CFU (withReveal)
  // =========================================================================
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Thesis, Specific, or General?",
      "Show Me Boards",
      "\"Animals were essential to the success of the war effort and their contribution must never be forgotten.\"\n\nOn your whiteboard: write T, S, or G.",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      const ansY = 4.10;
      slide.addShape("roundRect", {
        x: 0.5, y: ansY, w: 9, h: 0.85, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      });

      slide.addShape("roundRect", {
        x: 0.7, y: ansY + 0.10, w: 1.4, h: 0.28, rectRadius: 0.08,
        fill: { color: C.WHITE },
      });
      slide.addText("Answer", {
        x: 0.7, y: ansY + 0.10, w: 1.4, h: 0.28,
        fontSize: 10, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      slide.addText("T -- Thesis. It restates the POSITION directly. A general statement would zoom out to something bigger than the essay.", {
        x: 2.3, y: ansY + 0.08, w: 7.0, h: 0.68,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU);
    }
  );

  // =========================================================================
  // SLIDE 9 -- We Do: Plan Concluding Paragraph
  // =========================================================================
  contentSlide(
    pres,
    "We Do",
    C.SUCCESS,
    "Plan Your Concluding Paragraph",
    [
      "Look at your introduction and body paragraphs",
      "T -- Thesis: restate your position in one strong sentence",
      "S -- Specific: what was your STRONGEST argument with the best evidence?",
      "G -- General: zoom out. Why should the reader care? What should they do or think?",
      "Discuss with your partner, then fill in the concluding paragraph planner",
    ],
    NOTES_WEDO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 10 -- You Do
  // =========================================================================
  taskSlide(
    pres,
    "You Do",
    "Write, Compile, and Edit",
    [
      { label: "FIRST", instruction: "Write your concluding paragraph from your TSG planner" },
      { label: "NEXT", instruction: "Put all parts together: introduction + body paragraph 1 + body paragraph 2 + conclusion. Reread from the start" },
      { label: "THEN", instruction: "Use the editing checklist to review and improve your complete persuasive text" },
    ],
    NOTES_YOUDO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 11 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    "What is one thing you are most proud of in your persuasive text? What persuasive technique did you use best?",
    [
      "I can identify the purpose and features of a concluding paragraph in a persuasive text",
      "I can plan and write a concluding paragraph that restates my position and sums up my arguments",
      "I can edit my complete persuasive text for clarity, language choices, and effect",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 12 -- Resources
  // =========================================================================

  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: Concluding Paragraph Planner ----------------------------------
  const plan = createPdf({ title: PLANNER_RESOURCE.name });
  let planY = addPdfHeader(plan, "Concluding Paragraph Planner", {
    color: C.PRIMARY,
    subtitle: "TSG Structure",
    lessonInfo: "War Horse | Lesson 30 | Week 6 | Year 5/6 Literacy",
    showNameDate: true,
  });

  planY = addTipBox(plan, "Your concluding paragraph has THREE parts:\n- T (Thesis): Restate your position with confidence\n- S (Specific): Highlight your strongest argument\n- G (General): Big-picture takeaway or call to action\n\nThe conclusion does NOT introduce new evidence. It wraps up what you have already argued.", planY, { color: C.PRIMARY });

  planY = addSectionHeading(plan, "My Position (from my introduction):", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "Look at your introduction. Write your position here as a reminder.", planY, { fontSize: 9, italic: true });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 26 });
  planY += 8;

  planY = addSectionHeading(plan, "T -- Thesis Statement:", planY, { color: C.PRIMARY });
  planY = addBodyText(plan, "Restate your position. Use high modality language (must, undeniably, clearly). Sound confident -- you have PROVED your point.", planY, { fontSize: 9, italic: true });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 26 });
  planY += 8;

  planY = addSectionHeading(plan, "S -- Specific Statement:", planY, { color: C.SECONDARY });
  planY = addBodyText(plan, "Which was your STRONGEST argument? Summarise it in one powerful sentence.", planY, { fontSize: 9, italic: true });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 26 });
  planY += 8;

  planY = addSectionHeading(plan, "G -- General Statement:", planY, { color: C.ACCENT });
  planY = addBodyText(plan, "Zoom out. Why does this matter beyond your essay? What should the reader think, feel, or do?", planY, { fontSize: 9, italic: true });
  planY = addLinedArea(plan, planY, 2, { lineSpacing: 26 });
  planY += 12;

  planY = addSectionHeading(plan, "My Concluding Paragraph:", planY, { color: C.PRIMARY, fontSize: 14 });
  planY = addBodyText(plan, "Write your complete concluding paragraph here.", planY, { fontSize: 9, italic: true });
  planY = addLinedArea(plan, planY, 8, { lineSpacing: 26 });

  addPdfFooter(plan, "War Horse | Lesson 30 | Concluding Paragraph Planner");

  // --- PDF 2: Editing Checklist --------------------------------------------
  const edit = createPdf({ title: CHECKLIST_RESOURCE.name });
  let editY = addPdfHeader(edit, "Persuasive Text Editing Checklist", {
    color: C.ACCENT,
    subtitle: "Review Your Complete Text",
    lessonInfo: "War Horse | Lesson 30 | Week 6 | Year 5/6 Literacy",
    showNameDate: true,
  });

  editY = addTipBox(edit, "Read your complete persuasive text from beginning to end. Use this checklist to review each aspect of your writing. Tick each box when you have checked it, and fix any issues you find.", editY, { color: C.ACCENT });

  editY = addSectionHeading(edit, "Structure", editY, { color: C.PRIMARY });
  editY = addBodyText(edit, "[ ]  My text has all four parts: introduction, body paragraph 1, body paragraph 2, conclusion", editY);
  editY = addBodyText(edit, "[ ]  Each body paragraph focuses on ONE clear argument", editY);
  editY = addBodyText(edit, "[ ]  My body paragraphs have different arguments (not the same point repeated)", editY);
  editY = addBodyText(edit, "[ ]  My conclusion restates my position and does not introduce new evidence", editY);
  editY += 8;

  editY = addSectionHeading(edit, "Topic Sentences and Concluding Sentences", editY, { color: C.SECONDARY });
  editY = addBodyText(edit, "[ ]  Each body paragraph starts with a clear, interesting topic sentence", editY);
  editY = addBodyText(edit, "[ ]  Each body paragraph ends with a concluding sentence that links back to the argument", editY);
  editY = addBodyText(edit, "[ ]  My topic sentences make the reader want to keep reading", editY);
  editY += 8;

  editY = addSectionHeading(edit, "Persuasive Language", editY, { color: C.ACCENT });
  editY = addBodyText(edit, "[ ]  I have used high modality language (must, essential, undoubtedly, clearly)", editY);
  editY = addBodyText(edit, "[ ]  I have used at least one nominalisation (e.g., importance, contribution, destruction)", editY);
  editY = addBodyText(edit, "[ ]  I have used tier 2 or tier 3 vocabulary (e.g., significant, casualties, sacrifice)", editY);
  editY = addBodyText(edit, "[ ]  My language sounds confident and authoritative, not uncertain", editY);
  editY += 8;

  editY = addSectionHeading(edit, "Sentences and Spelling", editY, { color: C.PRIMARY });
  editY = addBodyText(edit, "[ ]  I have used a variety of sentence types (statements, questions, exclamations)", editY);
  editY = addBodyText(edit, "[ ]  I have used conjunctions to build compound and complex sentences", editY);
  editY = addBodyText(edit, "[ ]  I have checked my spelling, especially of key vocabulary words", editY);
  editY = addBodyText(edit, "[ ]  I have checked my punctuation (capital letters, full stops, commas)", editY);
  editY += 8;

  editY = addSectionHeading(edit, "Overall Effect", editY, { color: C.SECONDARY });
  editY = addBodyText(edit, "[ ]  My text would convince a reader who disagrees with me", editY);
  editY = addBodyText(edit, "[ ]  My conclusion leaves the reader with something to think about", editY);
  editY = addBodyText(edit, "[ ]  I am proud of this piece of writing", editY);
  editY += 12;

  editY = addSectionHeading(edit, "My Reflection:", editY, { color: C.ACCENT });
  editY = addBodyText(edit, "One thing I am proud of in my persuasive text:", editY, { fontSize: 10, italic: true });
  editY = addLinedArea(edit, editY, 2, { lineSpacing: 26 });
  editY += 4;
  editY = addBodyText(edit, "One thing I would improve if I had more time:", editY, { fontSize: 10, italic: true });
  editY = addLinedArea(edit, editY, 2, { lineSpacing: 26 });

  addPdfFooter(edit, "War Horse | Lesson 30 | Editing Checklist");

  // --- Write all files ------------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/WH6_Lesson30.pptx` }),
    writePdf(plan, PLANNER_PDF_PATH),
    writePdf(edit, CHECKLIST_PDF_PATH),
  ]);

  console.log(`PPTX written to ${OUT_DIR}/WH6_Lesson30.pptx`);
  console.log(`Done: ${PLANNER_RESOURCE.name}`);
  console.log(`Done: ${CHECKLIST_RESOURCE.name}`);
}

build().catch(console.error);
