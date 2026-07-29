"use strict";

// Tom Unit -- Session 9: Chapters 20-21 -- Setting Sail + Information Report Planning
// Week 2, Session 9, Grade 5/6 Literacy
// Reading: Chapters 20-21 (Scarborough sets sail, Tom meets Rob)
// Writing: Plan a body paragraph for information report (SPO) about 18th century England
// Cross-curricular: links to HASS/History

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
  vocabSlide, quoteSlide, modellingSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addTwoColumnOrganiser,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 9;
const FOOTER = "Chapters 20-21 | Session 9 | Week 2 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Session9_Setting_Sail";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const SPO_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "SPO Planning Template",
  "Student template: plan a body paragraph using a single paragraph outline (topic sentence, supporting details, concluding sentence)."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "SPO Model Answer",
  "Teacher reference: model SPO plan with alternative supporting details."
);
const RESOURCE_ITEMS = [SPO_RESOURCE, ANSWER_KEY_RESOURCE];
const SPO_PDF_PATH = path.join(OUT_DIR, SPO_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Session 9. Two chapters to read today: 20 and 21. The Scarborough finally sets sail
- After reading, we shift to our information report. Today we plan the first body paragraph about 18th century England
- This connects our novel to real history -- Tom's story is set during the First Fleet era

DO:
- Display title slide as students settle
- Have copies of the novel and the non-fiction article on desks

TEACHER NOTES:
Session 9 of 10. This is the pivot session -- reading continues but the writing focus shifts to text-level work (information report). Students read Chapters 20-21 then use the non-fiction article about 18th century England to plan a body paragraph. The KPAS skills from Session 8 support the note-taking from the article.

WATCH FOR:
- Students who need a recap of Chapter 19 -- Thomas reflected on his life, we learned about his family and career
- Students ready for the cross-curricular connection -- the novel gives context to the historical content

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands: reading the final novel chapters for this week, and planning the first body paragraph of our information report
- Read the success criteria. SC1 is about what we learn from the reading. SC2 is about understanding body paragraph structure. SC3 is the planning itself

DO:
- Choral read the LI, then the SCs
- Brief context: "Our information report is about life in 18th century England. The novel gives us one perspective -- a fictional one. The non-fiction article gives us the facts"

TEACHER NOTES:
SC1 targets the reading analysis (authorial intent with Rob's character). SC2 is the structural knowledge (body paragraph = TS + supporting details + CS). SC3 is the application (planning using SPO). The progression moves from reading comprehension to structural understanding to independent planning.

WATCH FOR:
- Students unsure about what an information report is -- brief reminder: "A factual text that classifies or describes information about a topic"
- Students who remember SPO from Session 3 -- they can support peers

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_READING = `SAY:
- Chapters 20 and 21. Short chapters today. The Scarborough sets sail with the First Fleet
- I will read aloud. Two pause points
- Focus: a new character appears -- a boy on the upper deck. What does the author want us to know about him?
- Find Chapter 20 now

DO:
- Give students 30 seconds to find the chapter
- Teacher reads aloud
- Plan for 2 pause points: p.102 and p.106

TEACHER NOTES:
Chapter 20 introduces Rob, a free boy on the upper deck. The contrast between Tom (convict below) and Rob (passenger above) is the key authorial choice. Chapter 21 is brief -- the ship anchors at Santa Cruz and Tom learns Rob's name.

WATCH FOR:
- Students who notice the physical separation (below deck vs upper deck) as metaphor for social class -- excellent
- Students curious about Santa Cruz and the voyage route -- channel into the information report

[General: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause here. "The horde of convicts limped obediently down the hatch once more"
- Ask: What is the big idea here? [The convicts are controlled, obedient, broken. They "limped" -- they are physically damaged. They go "down" again -- back into darkness. The word "horde" dehumanises them -- they are not individuals, they are a mass]
- Ask: What specific words tell us how the author views their situation? [Limped -- injured. Obediently -- no resistance left. Horde -- dehumanised. Down the hatch -- confined]

DO:
- Display the quote
- Cold Call for initial interpretation
- Push for close reading of individual word choices
- Finger vote: "Does the author want us to feel sorry for the convicts? 1 = yes, 2 = no, 3 = something more complicated" [Aim for 3]

CFU CHECKPOINT:
Technique: Finger Voting (1/2/3)

Script:
- "Fingers up: 1 = the author wants sympathy, 2 = no sympathy intended, 3 = it is more complicated than that"
- Scan for: mostly 3s
- Follow up: "Why is it complicated?" [The author shows them as both victimised AND dehumanised. We feel for them but the language strips their individuality]

PROCEED (>=80%): Most show 3 -- they see the complexity. Continue reading.
PIVOT (<80%): Most likely issue -- students choose 1 (simple sympathy) without seeing the dehumanising language. Reteach: "The author could have written 'The prisoners walked back below deck.' Instead they wrote 'horde' and 'limped obediently.' Why change from 'prisoners' to 'horde'? What does 'horde' make them sound like?" [Animals, a swarm, a mass -- not individual people.] Re-check: "So is the author JUST making us feel sorry, or also showing us something about how the system treats people?"

TEACHER NOTES:
This pause point develops close-reading skills. The word choices are deliberate and reveal both the convicts' suffering and the system's dehumanisation of them.

WATCH FOR:
- Students who pick up on "horde" -- excellent vocabulary awareness from a character of the system
- Students who connect "down" to the recurring descent motif -- link to Session 8's "Down, down, down"

[General: Pause Point 1 | VTLM 2.0: Higher-Order Questioning]`;

const NOTES_PAUSE2 = `SAY:
- Pause here. "But now he knew that the boy's name was Rob"
- Ask: What do you think the author wants us to know about Rob? [Rob is important -- the author has been building up to this moment. Tom has been watching Rob from below deck. Knowing his name makes Rob real, makes a connection possible. The author is setting up a relationship between the two boys]
- Ask: Why is the name significant? Think about what a name does in this novel. [Tom barely has an identity in the system -- he is a number, a convict, a "horde." Learning Rob's name is a human connection. Names matter because the system tries to erase them]

DO:
- Display the quote
- Think-Pair-Share: 20 seconds think, 30 seconds pair, share
- Draw the contrast: Tom below deck, Rob above. Convict and free passenger. Yet the author brings them together through a name

TEACHER NOTES:
This pause point is about authorial foreshadowing and the significance of human connection. Rob will become important in later chapters. The name revelation is a deliberate narrative choice.

WATCH FOR:
- Students who predict Rob will become Tom's friend -- acknowledge: "The author is certainly setting that up"
- Students who notice the deck separation as social metaphor -- strong analytical thinking

[General: Pause Point 2 | VTLM 2.0: Deep Comprehension]`;

const NOTES_INFO_REPORT = `SAY:
- Now we shift to our information report. We have been reading a novel set in the late 1700s. Today we connect that fiction to real history
- Our information report topic: 18th century England
- Very briefly: an information report classifies or describes factual information. Structure: title, introduction, body paragraphs, conclusion
- Today we focus on ONE body paragraph. A body paragraph has three parts: a topic sentence, supporting details, and a concluding sentence
- Topic sentence: introduces the main idea -- tells the reader what this paragraph is about
- Supporting details: sentences that expand on the topic sentence with specific information
- Concluding sentence: wraps up the paragraph and reminds the reader of the main idea without repeating the topic sentence

DO:
- Display the structure overview
- Refer to the mentor text if available -- point to the body paragraph structure
- Keep this brief -- students covered this in Session 4. This is a quick revision, not a full reteach

TEACHER NOTES:
Students were introduced to information report structure in Session 4 and planned a title/introduction. This session focuses on the first body paragraph. The revision should be quick -- 2-3 minutes maximum.

WATCH FOR:
- Students who remember the structure from Session 4 -- they can summarise for the class
- Students who are unsure about topic sentences -- the I Do model will clarify

[General: Review | VTLM 2.0: Retention and Recall]`;

const NOTES_IDO_SPO = `SAY:
- Watch me plan a body paragraph using a single paragraph outline -- an SPO
- My paragraph topic: life in 18th century England
- First, my topic sentence. This tells the reader what the paragraph is about. "Life in 18th century England was shaped by the dramatic changes of the Industrial Revolution"
- Now, supporting details. I look at my non-fiction article and my notes. What key facts support my topic sentence?
  - Detail 1: Factories replaced craftsmen, changing how people worked
  - Detail 2: Cities grew rapidly as people moved from rural areas for factory jobs
  - Detail 3: Living conditions were often poor -- overcrowding, pollution, disease
- Finally, my concluding sentence. This wraps up without repeating the topic sentence: "These changes transformed English society, creating both new opportunities and significant hardship"
- Notice: my SPO is a plan, not a finished paragraph. Each detail is a phrase or short sentence, not a polished piece of writing

DO:
- Display the SPO template on screen
- Model filling in each section step by step
- Think aloud: "I check -- do my supporting details actually support my topic sentence? Detail about factories -- yes, that is about the Industrial Revolution. Detail about cities -- yes, that is a consequence. Detail about living conditions -- yes, that is an impact"
- Show the non-fiction article as the source for the supporting details

TEACHER NOTES:
The I Do models the complete SPO process. The think-aloud about checking that details support the topic sentence is the key metacognitive move -- students often include interesting but irrelevant details.

WATCH FOR:
- Students who want to write full paragraphs instead of planning -- redirect: "Today is planning. Tomorrow is writing"
- Students who struggle to find supporting details -- show them how to scan the article for key facts

[General: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Now you will plan your own body paragraph using the SPO template on your worksheet
- Your topic: 18th century England -- same as mine, but YOUR supporting details
- I have provided a topic sentence to get you started. You will generate your own supporting details and concluding sentence
- Use the non-fiction article as your source. Use your KPAS skills from last session to find the key information
- Students who feel confident may write their own topic sentence as well
- You have 10 minutes. I will circulate to check your plans

DO:
- Distribute the Session 9 SPO Planning Template
- The template has the topic sentence pre-filled for most students
- Circulate -- check that supporting details come from the article and genuinely support the topic sentence
- Confer with students working independently on their own topic sentence

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide the topic sentence AND the first supporting detail pre-filled. Students find and write 2 more supporting details and the concluding sentence. The article should have key sections highlighted for these students
- Extra Notes: These students can use their KPAS notes from Session 8 as a scaffold

EXTENDING PROMPT:
- Task: Plan TWO body paragraphs using SPO -- one about living conditions and one about education or work in 18th century England. Write both topic sentences independently. Each paragraph must draw on different sections of the non-fiction article

TEACHER NOTES:
The differentiation for enabling students includes a pre-filled topic sentence and first detail, reducing the cognitive load. Extending students plan two paragraphs independently, demonstrating they can structure and source content across multiple body paragraphs.

WATCH FOR:
- Students who copy sentences from the article instead of noting key ideas -- remind: "Use your KPAS skills. Notes, not copies"
- Students whose supporting details do not connect to the topic sentence -- the check question: "Does this detail tell us about the Industrial Revolution's impact?"
- Students finishing quickly -- direct to extending task

[General: We Do / You Do | VTLM 2.0: Guided + Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: explaining what the author wants us to know about Rob -- thumbs? [scan]
- SC2: understanding body paragraph structure (TS, supporting details, CS) -- thumbs? [scan]
- SC3: planning your own body paragraph using an SPO -- thumbs? [scan]
- Turn and talk: what is the most interesting fact you found in the non-fiction article about 18th century England?

DO:
- Run through each SC with thumbs check
- The turn-and-talk connects the reading to the writing -- students share facts that will appear in their reports
- Preview: "Next session you will write the body paragraph from your plan. Bring your SPO to the next lesson"

TEACHER NOTES:
The closing bridges today's planning to tomorrow's writing. Students should leave with a complete SPO ready to convert into a full paragraph.

WATCH FOR:
- Students whose SPO is incomplete -- they may need time at the start of Session 10 to finish
- Students "thumbs down" on SC2 -- the structure (TS + details + CS) may need a brief re-model at the start of Session 10

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${SPO_RESOURCE.name} is for planning your body paragraph
- The ${ANSWER_KEY_RESOURCE.name} is for teacher reference with the model SPO

DO:
- Print the SPO template before the lesson (one per student)
- Print the model answer (teacher copy only)
- Students keep their completed SPO for use in Session 10

TEACHER NOTES:
Students must bring their completed SPO to Session 10, where they will convert it into a full paragraph. If students lose their template, have spares available.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Chapters 20-21: Setting Sail -- Session 9";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "Chapters 20-21",
    "Setting Sail",
    "Session 9  |  Week 2  |  Year 5/6 Literacy",
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
      "We are learning to analyse how an author introduces a new character, and to plan a body paragraph for an information report using a single paragraph outline",
    ],
    [
      "I can explain what the author wants us to know about a new character",
      "I can describe the structure of a body paragraph (topic sentence, supporting details, concluding sentence)",
      "I can plan a body paragraph using a single paragraph outline",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Reading Launch
  // =========================================================================
  contentSlide(
    pres,
    "Teacher Read Aloud",
    C.PRIMARY,
    "Chapters 20-21",
    [
      "Reading Mode: Teacher Read Aloud",
      "Chapter 20: The Scarborough sets sail with the First Fleet",
      "Chapter 21: The ship anchors at Santa Cruz -- Tom learns the boy's name",
      "Focus: A new character appears on the upper deck. What does the author want us to know?",
    ],
    NOTES_READING,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Pause Point 1: "The horde of convicts"
  // =========================================================================
  quoteSlide(
    pres,
    "Pause Point 1",
    "Chapter 20 -- p. 102",
    "The horde of convicts limped obediently down the hatch once more.",
    "p. 102",
    "What is the big idea? Look closely at the words: 'horde', 'limped', 'obediently'. What is the author showing us?",
    NOTES_PAUSE1,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- Pause Point 2: "his name was Rob"
  // =========================================================================
  quoteSlide(
    pres,
    "Pause Point 2",
    "Chapter 21 -- p. 106",
    "But now he knew that the boy's name was Rob.",
    "p. 106",
    "What do you think the author wants us to know about Rob? Why is knowing his name significant?",
    NOTES_PAUSE2,
    FOOTER
  );

  // =========================================================================
  // SLIDE 6 -- Information Report: Body Paragraph Structure
  // =========================================================================
  contentSlide(
    pres,
    "Information Report",
    C.SECONDARY,
    "Body Paragraph Structure",
    [
      "A body paragraph classifies or describes one aspect of your topic",
      "Topic Sentence (TS) -- introduces the main idea of the paragraph",
      "Supporting Details -- sentences that expand on the topic sentence with evidence and examples",
      "Concluding Sentence (CS) -- wraps up the paragraph and reminds the reader of the main idea",
      "Today: plan your first body paragraph about 18th century England",
    ],
    NOTES_INFO_REPORT,
    FOOTER
  );

  // =========================================================================
  // SLIDE 7 -- I Do: Model SPO
  // =========================================================================
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Single Paragraph Outline (SPO)",
    "Topic: 18th Century England\n\nTopic Sentence:\n\"Life in 18th century England was\nshaped by the dramatic changes\nof the Industrial Revolution.\"\n\nCheck: Does my TS tell the reader\nwhat this paragraph is about?\nYes -- the Industrial Revolution's\nimpact on life.",
    "Supporting Details:\n1. Factories replaced craftsmen,\n   changing how people worked\n2. Cities grew rapidly as people\n   moved for factory jobs\n3. Living conditions were often poor --\n   overcrowding, pollution, disease\n\nConcluding Sentence:\n\"These changes transformed English\nsociety, creating both new\nopportunities and significant\nhardship.\"",
    NOTES_IDO_SPO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 8 -- We Do / You Do: Plan Your SPO
  // =========================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do / You Do", { color: C.SECONDARY, w: 2.0 });
    addTitle(s, "Plan Your Body Paragraph");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.4, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("On your SPO template:", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.28,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0,
    });
    s.addText("First: Read the topic sentence provided (or write your own)\nNext: Find 3 supporting details from the non-fiction article\nThen: Write a concluding sentence that wraps up the paragraph", {
      x: 0.75, y: CONTENT_TOP + 0.44, w: 8.4, h: 0.80,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 1.54;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Remember", {
      x: 0.75, y: tipY + 0.08, w: 3, h: 0.28,
      fontSize: 11, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Supporting details must SUPPORT your topic sentence\n- Use the non-fiction article as your source -- use KPAS to find key facts\n- Your concluding sentence should NOT repeat your topic sentence\n- This is a PLAN, not a finished paragraph", {
      x: 0.75, y: tipY + 0.38, w: 8.4, h: 0.90,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
  }

  // =========================================================================
  // SLIDE 9 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    "What is the most interesting fact you found in the non-fiction article about 18th century England? Tell your partner.",
    [
      "I can explain what the author wants us to know about a new character",
      "I can describe the structure of a body paragraph (topic sentence, supporting details, concluding sentence)",
      "I can plan a body paragraph using a single paragraph outline",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 10 -- Resources
  // =========================================================================


  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: SPO Planning Template -----------------------------------------
  const ws = createPdf({ title: SPO_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Single Paragraph Outline (SPO) -- Body Paragraph", {
    color: C.PRIMARY,
    subtitle: "Information Report: 18th Century England",
    lessonInfo: "Session 9 | Week 2 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "Use this template to plan your first body paragraph. A body paragraph has: a Topic Sentence (TS) that introduces the main idea, Supporting Details that expand on the TS with evidence, and a Concluding Sentence (CS) that wraps up.", wsY, { color: C.PRIMARY });

  wsY = addSectionHeading(ws, "Topic Sentence (TS)", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "The topic sentence tells the reader what this paragraph is about. Use the provided sentence or write your own:", wsY, { fontSize: 10, italic: true });
  wsY += 4;
  wsY = addTipBox(ws, "Provided TS: \"Life in 18th century England was shaped by the dramatic changes of the Industrial Revolution.\"", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Your topic sentence (if writing your own):", wsY, { fontSize: 10 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Supporting Details", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Find 3 key facts from the non-fiction article that SUPPORT your topic sentence. Write them as notes or short phrases.", wsY, { fontSize: 10, italic: true });
  wsY += 4;
  wsY = addBodyText(ws, "Detail 1:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 6;
  wsY = addBodyText(ws, "Detail 2:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 6;
  wsY = addBodyText(ws, "Detail 3:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Concluding Sentence (CS)", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Wrap up your paragraph. Remind the reader of the main idea WITHOUT repeating your topic sentence.", wsY, { fontSize: 10, italic: true });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 10;

  wsY = addSectionHeading(ws, "Self-Check", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "- Do my supporting details actually support my topic sentence?", wsY, { fontSize: 10 });
  wsY = addBodyText(ws, "- Did I use information from the non-fiction article (not made up)?", wsY, { fontSize: 10 });
  wsY = addBodyText(ws, "- Does my concluding sentence wrap up without repeating the TS?", wsY, { fontSize: 10 });
  wsY = addBodyText(ws, "- Do I have enough detail to write a full paragraph in Session 10?", wsY, { fontSize: 10 });

  addPdfFooter(ws, "Session 9 | SPO Planning Template -- Keep for Session 10");

  // --- PDF 2: Model Answer --------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "SPO Model Answer -- Body Paragraph", {
    color: C.ALERT,
    subtitle: "Teacher Reference -- 18th Century England",
    lessonInfo: "Session 9 | Week 2 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "This is the teacher's model SPO demonstrated during I Do. Student plans will vary -- assess whether supporting details genuinely connect to the topic sentence and come from the article.", akY, { color: C.ALERT });

  akY = addSectionHeading(ak, "Model SPO", akY, { color: C.PRIMARY });

  akY = addBodyText(ak, "Topic Sentence: \"Life in 18th century England was shaped by the dramatic changes of the Industrial Revolution.\"", akY);
  akY += 4;
  akY = addBodyText(ak, "Detail 1: Factories replaced craftsmen, changing how people worked -- machinery took over manual labour", akY);
  akY = addBodyText(ak, "Detail 2: Cities grew rapidly as people moved from rural areas for factory jobs -- urbanisation", akY);
  akY = addBodyText(ak, "Detail 3: Living conditions were often poor -- overcrowding, lack of hygiene, pollution, disease", akY);
  akY += 4;
  akY = addBodyText(ak, "Concluding Sentence: \"These changes transformed English society, creating both new opportunities and significant hardship.\"", akY);
  akY += 14;

  akY = addSectionHeading(ak, "Alternative Supporting Details (from the article)", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "- Children as young as 5 worked in factories and mines (child labour)", akY);
  akY = addBodyText(ak, "- Education was limited -- charity schools and dame schools for the poor; private tutors and accomplishments for the wealthy", akY);
  akY = addBodyText(ak, "- The gap between rich and poor widened -- mansions vs slums", akY);
  akY = addBodyText(ak, "- Malnutrition and disease were common among the working class", akY);
  akY = addBodyText(ak, "- Orphans were sent to workhouses or became chimney sweeps (connects to the novel)", akY);
  akY += 10;

  akY = addSectionHeading(ak, "What to Look For in Student Plans", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- Topic sentence clearly introduces what the paragraph is about", akY);
  akY = addBodyText(ak, "- Supporting details are factual (from the article, not invented)", akY);
  akY = addBodyText(ak, "- Supporting details genuinely support the topic sentence (not tangential)", akY);
  akY = addBodyText(ak, "- Concluding sentence wraps up without repeating the TS word-for-word", akY);
  akY = addBodyText(ak, "- Plan has enough substance to generate a full paragraph in Session 10", akY);

  addPdfFooter(ak, "Session 9 | SPO Model Answer -- TEACHER COPY");

  // --- Write all files ----------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Session9.pptx` }),
    writePdf(ws, SPO_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Session9.pptx`);
  console.log("Done: " + SPO_RESOURCE.name);
  console.log("Done: " + ANSWER_KEY_RESOURCE.name);
}

build().catch(console.error);
