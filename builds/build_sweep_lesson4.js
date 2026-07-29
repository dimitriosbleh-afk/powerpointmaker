"use strict";

// Sweep Unit -- Lesson 4: Chapters 9-10 + Info Report Planning
// Week 1, Session 4, Grade 5/6 Literacy
// Novel: "Sweep" by Jonathan Auxier
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

const SESSION_NUMBER = 4;
const FOOTER = "Sweep | Lesson 4 | Week 1 | Year 5/6 Literacy";
const OUT_DIR = "output/Sweep_Lesson4_Chapters9_10";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLANNING_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Info Report Planning Sheet",
  "Student planning sheet: GST introduction framework and note-taking organiser for information report."
);
const RESOURCE_ITEMS = [PLANNING_RESOURCE];
const PLANNING_PDF_PATH = path.join(OUT_DIR, PLANNING_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Lesson 4 today. We read our final novel chapters for this week -- Chapters 9 and 10
- Then something new: we begin planning an information report. This is a brand new text type for this term
- The information report connects to our novel. Tom lives in 18th century England. People like Tom were sometimes transported to Australia as convicts. Our report will be about Convict Settlement in Australia

DO:
- Display title slide as students settle
- Have copies of Sweep bookmarked at Chapter 9
- Have the non-fiction article ready to display or distribute

TEACHER NOTES:
Session 4 bridges the novel study and the cross-curricular writing. Chapters 9-10 focus on Jem and Tom's growing bond and their plan to escape. The writing shift to information reports is significant -- students are moving from narrative response to expository writing. The convict settlement topic connects thematically to the novel (18th century England, poverty, injustice).

WATCH FOR:
- Students who see the connection between Tom's England and convict transportation -- excellent historical thinking
- Students who are anxious about a new text type -- reassure: "We will learn this together, step by step"

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands: finishing our novel chapters for the week and beginning to plan an information report
- Read the success criteria. SC1 is about understanding what Tom and Jem's plan tells us about them as characters. SC2 is about understanding the purpose and structure of an information report. SC3 is about planning an introduction using the GST format
- GST stands for General statement, Specific statement, Thesis statement. We will learn what each one means

DO:
- Choral read the LI, then each SC
- Briefly preview: "An information report is factual writing that classifies and describes. We will learn its structure today"

TEACHER NOTES:
SC1 targets reading comprehension through character analysis. SC2 introduces the information report as a text type. SC3 targets the specific skill of planning an introduction using GST. SC1 is the floor (character inference). SC2 extends to text-type knowledge. SC3 is the application skill.

WATCH FOR:
- Students who think information reports are just lists of facts -- "It is organised, structured writing with a clear purpose"
- Students unfamiliar with the term "thesis statement" -- "It tells the reader what the report will discuss"

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_READING = `SAY:
- Chapters 9 and 10 -- the last two novel chapters this week
- Chapter 9: Jem shows Tom something secret. A hidden stash of goods in the cellar. Jem has a plan for escape
- Chapter 10: Tom and Jem talk about their families and how they ended up with Master Jack. They plan for the future
- Reading mode: student read aloud. Three pause points
- These chapters are quieter than 7-8 but important. The boys are not just surviving now -- they are planning

DO:
- Give students 30 seconds to find Chapter 9
- Select readers
- These chapters are dialogue-heavy -- choose readers who can handle conversational pace

TEACHER NOTES:
Chapters 9-10 develop the Tom-Jem relationship and introduce the escape plot. The stolen goods create moral complexity -- Jem steals to survive. The family backstory scenes add emotional depth. These chapters slow the pace after the intensity of the chimney rescue.

WATCH FOR:
- Students who judge Jem for stealing -- the moral complexity is worth discussing: "Is stealing wrong if you are stealing to escape slavery?"
- Students who notice the shift in Tom's character -- he is becoming more proactive, less passive

[General: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause. "...and felt the hate and hope blossom in his heart." Page 47
- What have we learned from this conversation? [Jem has been stealing small items from houses they clean and hiding them. He plans to sell them one day to buy his freedom. Tom feels both hatred for their situation and hope that escape is possible]
- Ask: The author pairs two opposite emotions -- hate and hope. Why? [They exist together. Tom hates what his life has become AND hopes for something better. Neither cancels the other out. This is realistic emotional complexity]

DO:
- Display the quote
- Cold Call 3 students
- Draw attention to the pairing of "hate and hope" -- alliteration AND emotional contrast

TEACHER NOTES:
This is a character-defining moment. Tom moves from passive endurance to active hope. The hate-hope pairing is a sophisticated authorial choice that Year 5/6 students can appreciate. The alliteration draws attention to the contrast.

WATCH FOR:
- Students who focus only on hate or only on hope -- push for both: "They exist together. That is what makes Tom human"
- Students who connect hope to the courage cloak -- both are coping mechanisms Tom creates

[General: Pause Point 1 | VTLM 2.0: Literary Analysis]`;

const NOTES_PAUSE2 = `SAY:
- Pause. "'...ye'd never have risked yer neck fer me back in the chimley.'" Page 50
- What have we learned from this conversation? [Jem acknowledges that Tom saved his life. Their bond has deepened. Jem trusts Tom enough to share his escape plan. The friendship is real -- not just convenience]
- Ask: How has Tom's relationship with Jem changed since Chapter 6? [In Ch6 Tom was a stranger entering Jem's world. Now they are partners. The chimney rescue built trust. Jem shows Tom his secret stash -- that is a huge act of trust]

DO:
- Display the quote
- Think-Pair-Share: how has the Tom-Jem relationship developed?
- Push for evidence from across the chapters, not just this page

CFU CHECKPOINT:
Technique: Think-Pair-Share

Script:
- "Think for 30 seconds: how has the relationship between Tom and Jem changed from Chapter 6 to now?"
- "Share with your partner. Use evidence from at least two different chapters"
- "I am selecting pairs. [Name], what evidence did your pair find?"
- Scan for: students tracking character development across chapters, not just describing the current scene

PROCEED (>=80%): Most pairs cite evidence from multiple chapters showing the relationship's growth. Continue reading.
PIVOT (<80%): Most likely issue -- students describe the current friendship without tracking how it developed. Reteach: "Let me show you the arc. Ch6: Tom arrives, Jem is just another boy in the cellar. Ch7: Jem teaches Tom about chimney work. Ch8: Tom saves Jem's life. Ch9: Jem shares his most dangerous secret. See the progression? Each event builds more trust." Re-check: "Name one event that changed their relationship."

TEACHER NOTES:
Character development tracking across chapters is a key analytical skill for Year 5/6. Students should be learning to see characters as dynamic, not static. The Tom-Jem arc is a strong example.

WATCH FOR:
- Students who only describe the current scene -- redirect to the arc across chapters
- Students who notice the dialect in Jem's speech ("yer," "chimley") -- worth noting as characterisation through dialogue

[General: Pause Point 2 | VTLM 2.0: Character Analysis]`;

const NOTES_PAUSE3 = `SAY:
- Final pause. "...and carried them wrapped in rags against his body up the chimney." Page 52
- How do things look for Tom and Jem now? [They have a plan. Jem steals small items and hides them. When they have enough, they will sell them and escape. There is real hope now -- but also real danger. If Master Jack finds the stash, the consequences would be severe]
- Ask: Is what Jem is doing right or wrong? Think carefully before you answer. [This is morally complex. Jem steals from wealthy homes to escape what is essentially child slavery. The morality depends on the context]

DO:
- Display the quote
- Let students debate briefly -- this is a genuine ethical question
- Do not resolve it -- let the ambiguity sit: "Sometimes stories ask questions that do not have simple answers"

TEACHER NOTES:
This closes the reading section for the week. The moral question about Jem's stealing is a strong discussion point that connects to the convict settlement topic -- many convicts were transported for theft driven by poverty. This bridge between the novel and the information report topic is worth flagging but not over-explaining.

WATCH FOR:
- Students who take absolute positions (stealing is always wrong / always okay) -- push for nuance
- Students who connect Jem's situation to convict transportation -- excellent cross-curricular thinking

[General: Pause Point 3 | VTLM 2.0: Critical Thinking]`;

const NOTES_IR_STRUCTURE = `SAY:
- Now our writing focus shifts. We are starting something new: an information report
- An information report classifies or describes factual information about a topic. It is not a story -- it is structured, factual writing
- The structure has four parts: Title, Introduction, Body Paragraphs, and Conclusion
- Key features: written in present tense, third person, subheadings, precise vocabulary, action verbs and relating verbs
- Our topic: Convict Settlement in Australia. This connects to our novel -- Tom's England is the same England that sent convicts to Australia

DO:
- Display the structure overview
- Point to each part and explain briefly
- Show the mentor text (if distributing copies, do so now)
- Emphasise: "We will learn this one piece at a time. Today we focus on the introduction"

TEACHER NOTES:
This is the first exposure to the information report text type this term. The structure overview is presented on the slide for reference but the lesson focuses specifically on the introduction. The connection to the novel grounds the topic in something students care about. The mentor text provides a concrete model to reference.

WATCH FOR:
- Students overwhelmed by the full structure -- reassure: "We are only writing the introduction today. The rest comes in future lessons"
- Students who notice the difference from narrative writing -- affirm: "Exactly. Different purpose, different structure"

[General: I Do -- Info Report Structure | VTLM 2.0: Explicit Teaching]`;

const NOTES_GST_IDO = `SAY:
- The introduction to an information report follows a pattern called GST
- G: General statement -- introduces the topic broadly. It gives context. It is the big picture
- S: Specific statement -- narrows in on the topic. It uses key terms and addresses the topic directly
- T: Thesis statement -- tells the reader what the body paragraphs will discuss. It is the roadmap for the report
- Watch me model. I am writing a title and introduction for our topic: Convict Settlement in Australia
- Title: "Convict Settlement in Australia: From British Prisons to a New Colony"
- G: "In the late 18th century, Britain faced a crisis in its overcrowded prisons and needed a new solution for dealing with convicted criminals."
- S: "The British government established a penal colony in New South Wales, transporting thousands of convicts across the world to the shores of Australia."
- T: "This report will examine life in 18th century England, the purpose of the First Fleet, the arrival in Sydney Cove, convict life in early Sydney, and the impact on First Nations Australians."

DO:
- Display the GST framework on the left
- Model writing each part step by step on the right
- Think aloud: "My G gives the big picture. My S zooms in. My T tells the reader exactly what is coming"
- Point out: the T statement lists what each body paragraph will cover

TEACHER NOTES:
The GST framework gives students a concrete structure for writing introductions. The model uses the exact topic from the unit plan. The thesis statement lists the five body paragraph topics that will be covered across the unit. Students do not need to write all body paragraphs -- only the introduction this week.

WATCH FOR:
- Students who confuse G and S -- "G is the wide view. S zooms in. Think of a camera: G is the wide shot, S is the close-up"
- Students who write a T that is too vague ("This report is about convicts") -- "List what you will discuss specifically"

[General: I Do -- GST Introduction | VTLM 2.0: Explicit Modelling]`;

const NOTES_WEDO = `SAY:
- Let's plan together. We are going to write a General statement as a class
- The topic is Convict Settlement in Australia. What is the BIG PICTURE context?
- Think about: when did this happen? Where? Why was it needed? What was the situation?
- With your partner, draft a General statement on your whiteboard. 45 seconds

DO:
- Display the prompt
- Give 45 seconds for partner drafting
- Take 3 responses
- Guide the class toward a strong G statement
- Model revising a weak example: "Convicts went to Australia" becomes "In the late 1700s, Britain struggled with overflowing prisons and harsh poverty, leading the government to search for a drastic solution"

CFU CHECKPOINT:
Technique: Show Me Boards

Script:
- "Write a General statement for the topic 'Convict Settlement in Australia' on your whiteboard. 45 seconds. Show me"
- Scan for: sentences that provide CONTEXT (when, why the situation existed) rather than jumping straight to the specific topic
- "A General statement sets the scene. It should NOT mention convict settlement yet -- save that for S. It should tell the reader about the WORLD that led to this event"

PROCEED (>=80%): Most boards show a contextual statement about 18th century Britain. Move to You Do.
PIVOT (<80%): Most likely issue -- students jump to the specific topic in their G statement ("The First Fleet sailed to Australia"). Reteach: "That is your S statement, not your G. G is WIDER. Think bigger. Why did the First Fleet happen at all? What was happening in England that made them send people to the other side of the world? THAT is your G." Re-check: "Rewrite. Start with 'In the late 1700s, Britain...' What was happening?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the sentence frame: "In the ___ century, Britain faced a problem because ___." Focus on stating ONE contextual fact about the time period
- Extra Notes: Students can refer to the non-fiction article for key facts

EXTENDING PROMPT:
- Task: Draft the complete GST introduction (all three sentences). Then read it aloud to check flow and coherence. Does G lead naturally to S? Does S lead to T?

TEACHER NOTES:
The We Do focuses on just the G statement to isolate the skill. The full GST comes in Lesson 5. Students need to understand the zoom-in logic: G (wide context) -> S (specific topic) -> T (roadmap). The most common error is making G too specific.

WATCH FOR:
- Students who write S instead of G -- the zoom distinction is the key concept
- Students who write strong contextual statements -- share as models

[General: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_YOUDO = `SAY:
- On the planning sheet, you have space to start your own information report plan
- FIRST: Write a title for your information report on Convict Settlement in Australia
- NEXT: Draft your General statement -- remember, this sets the big picture context
- THEN: Begin noting key facts from the non-fiction article that might support your introduction
- We will write the full introduction next lesson. Today is about planning and note-taking

DO:
- Distribute the Session 4 Info Report Planning Sheet
- Display or distribute the non-fiction article for reference
- Circulate -- check that titles are specific and General statements provide context
- After 5 minutes: "You should have a title and be working on your General statement"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Write only the title using the frame: "Convict Settlement in Australia: ___." Then write one fact from the article in your own words
- Extra Notes: Help these students read the article by highlighting key sentences together first

EXTENDING PROMPT:
- Task: Draft the complete GST introduction. Then write three key facts from the article, one for each potential body paragraph topic

TEACHER NOTES:
The You Do is a planning and note-taking session. Students are not writing the full introduction yet -- that is Lesson 5. The focus is on the title, the General statement, and beginning to gather facts from the non-fiction article. This sets up Lesson 5's writing session.

WATCH FOR:
- Students who try to write the entire report -- redirect: "Today is planning only. We write the introduction next lesson"
- Students who struggle with the article -- pair them with a stronger reader for note-taking support

[General: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: I can explain what Tom and Jem's escape plan reveals about their characters -- thumbs? [scan]
- SC2: I can describe the purpose and structure of an information report -- thumbs? [scan]
- SC3: I can plan an introduction using the GST format -- thumbs? [scan]
- Turn and talk: What connection can you see between Tom's story in the novel and the topic of convict settlement in Australia?

DO:
- Run through each SC with thumbs check
- The turn-and-talk invites cross-curricular connection
- Preview: "Next lesson we write the full introduction for our information report"

TEACHER NOTES:
The closing bridges the novel study and the information report. Students who can articulate the connection (18th century England, poverty, injustice, transportation) are demonstrating sophisticated cross-curricular thinking. Students showing thumbs-down on SC3 will get a second exposure to GST in Lesson 5.

WATCH FOR:
- Students who see the poverty-injustice-transportation thread -- excellent historical thinking
- Students thumbs-down on SC2 or SC3 -- information reports are new, revisited fully in Lesson 5

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- One resource today -- the planning sheet for your information report
- You will use this in Lesson 5 as well, so keep it safe

DO:
- Print the planning sheet before the lesson (one per student)
- Click the resource card to open the PDF

TEACHER NOTES:
The planning sheet carries across to Lesson 5. Students should bring their completed planning sheet to the next lesson.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Sweep - Lesson 4 - Chapters 9-10 + Info Report Planning";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "Sweep",
    "Chapters 9-10 -- Plans and Promises",
    "Lesson 4  |  Week 1  |  Year 5/6 Literacy",
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
      "We are learning to analyse character development through dialogue, and to understand the purpose and structure of an information report",
    ],
    [
      "I can explain what Tom and Jem's escape plan reveals about their characters",
      "I can describe the purpose and structure of an information report",
      "I can plan an introduction for an information report using GST format",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Reading Launch
  // =========================================================================
  contentSlide(
    pres,
    "Read Aloud",
    C.PRIMARY,
    "Chapters 9-10",
    [
      "Reading Mode: Student read aloud",
      "Ch 9: Jem reveals his secret stash and a plan for escape",
      "Ch 10: Tom and Jem share their stories and plan for the future",
      "Focus: what do their plans and conversations reveal about who they are?",
    ],
    NOTES_READING,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Pause Point 1 (p.47)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 1", "Chapter 9 -- p. 47",
    "...and felt the hate and hope blossom in his heart.",
    "p. 47",
    "The author pairs hate and hope. Why do both emotions exist together?",
    NOTES_PAUSE1, FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- Pause Point 2 (p.50)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 2", "Chapter 10 -- p. 50",
    "'...ye'd never have risked yer neck fer me back in the chimley.'",
    "p. 50",
    "How has the relationship between Tom and Jem changed since Chapter 6?",
    NOTES_PAUSE2, FOOTER
  );

  // =========================================================================
  // SLIDE 6 -- Pause Point 3 (p.52)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 3", "Chapter 10 -- p. 52",
    "...and carried them wrapped in rags against his body up the chimney.",
    "p. 52",
    "How do things look for Tom and Jem now? Is what Jem is doing right or wrong?",
    NOTES_PAUSE3, FOOTER
  );

  // =========================================================================
  // SLIDE 7 -- Info Report Structure
  // =========================================================================
  contentSlide(
    pres,
    "New Text Type",
    C.SECONDARY,
    "Information Report -- Structure",
    [
      "Purpose: classify or describe factual information about a topic",
      "Title -- tells the reader what the report is about",
      "Introduction -- introduces the topic (uses GST format)",
      "Body Paragraphs -- each covers one aspect of the topic",
      "Conclusion -- summarises the main points",
      "Features: present tense, third person, subheadings, precise vocabulary",
    ],
    NOTES_IR_STRUCTURE,
    FOOTER
  );

  // =========================================================================
  // SLIDE 8 -- I Do: GST Introduction
  // =========================================================================
  modellingSlide(
    pres,
    "I Do",
    "Planning an Introduction -- GST",
    "G: General Statement\nIntroduces the topic broadly.\nGives context -- the big picture.\n\nS: Specific Statement\nNarrows to the topic directly.\nUses key terms.\n\nT: Thesis Statement\nTells the reader what the body\nparagraphs will discuss.\nThe roadmap for the report.",
    "Topic: Convict Settlement\nin Australia\n\nG: \"In the late 18th century,\nBritain faced a crisis in its\novercrowded prisons.\"\n\nS: \"The government established\na penal colony in New South\nWales, transporting thousands\nof convicts to Australia.\"\n\nT: \"This report will examine\n18th century England, the First\nFleet, arrival in Sydney Cove,\nconvict life, and the impact on\nFirst Nations Australians.\"",
    NOTES_GST_IDO,
    FOOTER
  );

  // =========================================================================
  // SLIDES 9-10 -- We Do: General Statement (withReveal)
  // =========================================================================
  withReveal(
    () => contentSlide(
      pres,
      "We Do",
      C.SUCCESS,
      "Draft a General Statement",
      [
        "Topic: Convict Settlement in Australia",
        "The General statement sets the BIG PICTURE context",
        "Think: When? Where? Why was it needed? What was the situation?",
        "Draft a General statement on your whiteboard -- 45 seconds",
      ],
      NOTES_WEDO,
      FOOTER
    ),
    (slide) => {
      const ansY = 3.90;
      slide.addShape("roundRect", {
        x: 0.5, y: ansY, w: 9, h: 1.05, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      });

      slide.addShape("roundRect", {
        x: 0.7, y: ansY + 0.08, w: 1.5, h: 0.28, rectRadius: 0.08,
        fill: { color: C.WHITE },
      });
      slide.addText("Model G", {
        x: 0.7, y: ansY + 0.08, w: 1.5, h: 0.28,
        fontSize: 10, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      slide.addText("\"In the late 1700s, Britain struggled with overflowing prisons and widespread poverty, leading the government to search for a drastic new solution for its convicted criminals.\"", {
        x: 2.3, y: ansY + 0.06, w: 7.0, h: 0.90,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_WEDO);
    }
  );

  // =========================================================================
  // SLIDE 11 -- You Do: Planning
  // =========================================================================
  contentSlide(
    pres,
    "You Do",
    C.ACCENT,
    "Plan Your Information Report",
    [
      "FIRST: Write a title for your information report",
      "NEXT: Draft your General statement -- set the big picture context",
      "THEN: Note key facts from the non-fiction article for your introduction",
    ],
    NOTES_YOUDO,
    FOOTER,
    (slide, layoutGuide) => {
      const tipY = layoutGuide.panelTopPadded + 0.1;
      addCard(slide, layoutGuide.rightX, tipY, layoutGuide.rightW, 2.2, {
        strip: C.PRIMARY, fill: C.BG_CARD,
      });
      slide.addText("GST Reminder", {
        x: layoutGuide.rightX + 0.15, y: tipY + 0.08, w: 3.8, h: 0.28,
        fontSize: 11, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
      });
      slide.addText([
        { text: "G = Big picture context (wide shot)", options: { bullet: true, breakLine: true, fontSize: 12, color: C.CHARCOAL } },
        { text: "S = Zoom in on the topic (close-up)", options: { bullet: true, breakLine: true, fontSize: 12, color: C.CHARCOAL } },
        { text: "T = Roadmap -- what the report covers", options: { bullet: true, breakLine: true, fontSize: 12, color: C.CHARCOAL } },
        { text: "Today: Title + G + fact notes", options: { bullet: true, fontSize: 12, color: C.MUTED, italic: true } },
      ], {
        x: layoutGuide.rightX + 0.15, y: tipY + 0.44, w: 3.8, h: 1.5,
        fontFace: FONT_B, valign: "top", margin: 0, paraSpaceAfter: 5,
      });
    }
  );

  // =========================================================================
  // SLIDE 12 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    "What connection can you see between Tom's story in the novel and the topic of convict settlement in Australia?",
    [
      "I can explain what Tom and Jem's escape plan reveals about their characters",
      "I can describe the purpose and structure of an information report",
      "I can plan an introduction for an information report using GST format",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 13 -- Resources
  // =========================================================================


  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: Info Report Planning Sheet ------------------------------------
  const ps = createPdf({ title: PLANNING_RESOURCE.name });
  let psY = addPdfHeader(ps, "Information Report Planning Sheet", {
    color: C.NAVY,
    subtitle: "Convict Settlement in Australia",
    lessonInfo: "Sweep | Lesson 4 | Week 1 | Year 5/6 Literacy",
    showNameDate: true,
  });

  psY = addTipBox(ps, "An information report classifies or describes factual information.\nStructure: Title -> Introduction (GST) -> Body Paragraphs -> Conclusion\nFeatures: present tense, third person, subheadings, precise vocabulary", psY, { color: C.NAVY });

  // Title
  psY = addSectionHeading(ps, "Title", psY, { color: C.NAVY });
  psY = addBodyText(ps, "Write a clear, specific title for your information report:", psY, { fontSize: 10 });
  psY = addLinedArea(ps, psY, 1, { lineSpacing: 30 });
  psY += 6;

  // Introduction - GST
  psY = addSectionHeading(ps, "Introduction -- GST Format", psY, { color: C.SECONDARY });

  psY = addBodyText(ps, "G: General Statement (big picture context -- when, where, why)", psY, { fontSize: 10, bold: true });
  psY = addBodyText(ps, "Frame: \"In the ___ century, Britain ___\"", psY, { fontSize: 9, italic: true });
  psY = addLinedArea(ps, psY, 2, { lineSpacing: 26 });
  psY += 4;

  psY = addBodyText(ps, "S: Specific Statement (zoom in on the topic, use key terms)", psY, { fontSize: 10, bold: true });
  psY = addBodyText(ps, "Frame: \"The British government ___\"", psY, { fontSize: 9, italic: true });
  psY = addLinedArea(ps, psY, 2, { lineSpacing: 26 });
  psY += 4;

  psY = addBodyText(ps, "T: Thesis Statement (what the body paragraphs will discuss)", psY, { fontSize: 10, bold: true });
  psY = addBodyText(ps, "Frame: \"This report will examine ___\"", psY, { fontSize: 9, italic: true });
  psY = addLinedArea(ps, psY, 2, { lineSpacing: 26 });
  psY += 8;

  // Key Facts from Non-Fiction Article
  psY = addSectionHeading(ps, "Key Facts from the Non-Fiction Article", psY, { color: C.ACCENT });
  psY = addBodyText(ps, "Write key facts in your own words. These will support your introduction and body paragraphs.", psY, { fontSize: 9, italic: true });
  psY += 2;
  for (let i = 1; i <= 5; i++) {
    psY = addWriteLine(ps, `Fact ${i}:`, psY, {});
    psY += 4;
  }

  addPdfFooter(ps, "Sweep | Lesson 4 | Info Report Planning Sheet -- Keep for Lesson 5");

  // --- Write all files ------------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Sweep_Lesson4.pptx` }),
    writePdf(ps, PLANNING_PDF_PATH),
  ]);

  console.log(`PPTX written to ${OUT_DIR}/Sweep_Lesson4.pptx`);
  console.log(`Done: ${PLANNING_RESOURCE.name}`);
}

build().catch(console.error);
