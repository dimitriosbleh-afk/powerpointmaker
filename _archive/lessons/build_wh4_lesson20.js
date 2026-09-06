"use strict";

// War Horse Unit -- Lesson 20: Plan and Write an Introductory Paragraph
// Week 4, Session 5, Grade 5/6 Literacy
// Supplementary text: War Horse Facts → GST structure → Model intro → Students write

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  withReveal,
  titleSlide, liSlide, contentSlide,
  cfuSlide, taskSlide, closingSlide,
  modellingSlide,
} = require("../themes/wh4_helpers");

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 5;
const FOOTER = "War Horse | Lesson 20 of 25 | Week 4 | Year 5/6 Literacy";
const OUT_DIR = "output/WH4_Lesson20_Introductory_Paragraph";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLANNING_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "GST Planning Template",
  "Planning template: General, Specific, Thesis structure for introductory paragraph. Includes sentence starters."
);
const ENABLING_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Enabling Scaffold",
  "Scaffolded planning template with sentence starters pre-filled for each GST component."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Answer Key",
  "Teacher reference: model introductory paragraphs (for and against) with GST components annotated."
);
const RESOURCE_ITEMS = [PLANNING_RESOURCE, ENABLING_RESOURCE, ANSWER_KEY_RESOURCE];
const PLANNING_PDF_PATH = path.join(OUT_DIR, PLANNING_RESOURCE.fileName);
const ENABLING_PDF_PATH = path.join(OUT_DIR, ENABLING_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today we move from planning to writing. We are going to focus on the introductory paragraph of our persuasive text
- Last lesson we brainstormed arguments for and against using horses in WWI. Today we learn how to open our persuasive text powerfully

DO:
- Display title slide as students settle
- Have the War Horse Facts article ready for the reading section

TEACHER NOTES:
Lesson 20 of 25, fifth and final session of Week 4. This lesson bridges planning (Lesson 19) and extended writing (Lessons 21+). The introductory paragraph is the critical first piece students write.

WATCH FOR:
- Students who were absent for Lesson 19 -- they need a brief catch-up on the topic and their position

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Today we are zooming in on ONE section of the persuasive text -- the introduction
- Read the success criteria together. SC1 is about understanding the structure. SC2 is about planning. SC3 is about actually writing
- Ask: Which SC feels most achievable right now? [Take 2-3 responses]

DO:
- Choral read the LI, then the SCs
- Emphasise the progression: understand the structure, plan your intro, write your intro

TEACHER NOTES:
The three SC are deliberately progressive -- structural knowledge, planning, then writing. SC1 is the floor (every student should know GST by end of lesson). SC3 is the stretch (writing with rhetorical devices).

WATCH FOR:
- Students who are anxious about writing -- reassure: "We plan first, then we write. And I model every step before you try"

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_READING = `SAY:
- Before we start writing, let's build more knowledge about our topic
- Today's text is "War Horse Facts" from The Brooke charity. It gives us important details about how horses were used in both World Wars
- Reading mode is Student Read Aloud. I will select readers
- As you read, take notes on any facts that could strengthen your persuasive argument

DO:
- Select readers for the article (change every fact or short section)
- If the videos are available, play them after the reading
- Circulate during note-taking: check students are identifying facts relevant to their chosen position

TEACHER NOTES:
This supplementary text provides additional factual ammunition for the persuasive writing. Students should be actively connecting facts to their arguments from Lesson 19.

WATCH FOR:
- Students not taking notes -- prompt: "Write down one fact that supports YOUR side of the argument"
- Students who find facts that challenge their position -- great: "That is what the other side would say. Knowing their arguments makes yours stronger"

[General: Reading Launch | VTLM 2.0: Research and Note-Taking]`;

const NOTES_REVISION = `SAY:
- Quick revision before we start writing. What is our persuasive writing topic? [Should horses have been used in WWI?]
- Last lesson we brainstormed arguments for and against. Who can give me one argument for? [Take 1-2] One against? [Take 1-2]
- We also looked at the macrostructure: Introduction, Body Paragraphs, Conclusion. Today we focus on the Introduction

DO:
- Quick oral recall -- keep this to 2 minutes maximum
- Use Cold Call to check recall from Lesson 19
- Display the topic prominently before moving to the GST structure

TEACHER NOTES:
Brief spaced retrieval from Lesson 19. Students need to reactivate their planning before writing.

WATCH FOR:
- Students who can't recall any arguments -- pair them with a partner to review their planning organiser from Lesson 19

[General: Review | VTLM 2.0: Retention and Recall]`;

const NOTES_GST = `SAY:
- The introductory paragraph of a persuasive text has three parts. We call it GST
- G -- General Statement: a broad opening about the topic that hooks the reader. It does NOT state your opinion yet
- S -- Specific Statement: narrow down to the specific issue you are writing about
- T -- Thesis Statement: clearly state YOUR position -- what you believe and why
- Think of it as a funnel: start wide, then narrow down to your point

DO:
- Display the GST structure clearly
- Use the funnel analogy -- point to each section getting more specific
- Read the example sentences for each component on the right

TEACHER NOTES:
GST is the foundational structure for persuasive introductions. The funnel metaphor helps students understand the narrowing from general to specific to personal position.

WATCH FOR:
- Students who want to state their opinion in the General Statement -- redirect: "Save your opinion for the T. The G is about the big picture"
- Students confused about the difference between S and T -- S is the specific topic, T is YOUR position on it

[General: I Do -- Text Structure | VTLM 2.0: Explicit Teaching]`;

const NOTES_RHETORICAL = `SAY:
- Two powerful tools for your introduction: rhetorical questions and emotive language
- A rhetorical question makes the reader THINK without expecting an answer. It draws them in
- Example: "How could we send thousands of innocent animals to die in a war they could never understand?"
- Emotive language makes the reader FEEL. Choose words that trigger emotion
- Example: "Innocent horses were exploited, injured, and abandoned in the horrific conditions of World War 1"
- Notice: "exploited" is stronger than "used," "horrific" is stronger than "bad"

DO:
- Read each example with expression -- model the emotional weight
- Ask: "How does the rhetorical question make you feel compared to the statement?"
- Point to the word choices that carry emotional weight

TEACHER NOTES:
These two features are taught explicitly now because they appear in the teacher's model paragraph. Students need to recognise them before they can use them.

WATCH FOR:
- Students who write actual questions expecting answers instead of rhetorical questions -- clarify: "A rhetorical question makes a point. You don't want an answer"
- Students who overload with emotive language -- model: "One or two strong emotional words per sentence is enough"

[General: I Do -- Language Features | VTLM 2.0: Explicit Teaching]`;

const NOTES_MODEL = `SAY:
- Watch me write a complete introductory paragraph. I am arguing FOR -- that horses should have been used
- First, my General Statement: "Throughout history, animals have played a vital role in human conflict, serving alongside soldiers in ways that shaped the outcome of wars." That is broad -- about animals in war generally
- Next, my Specific Statement: "During World War 1, horses were among the most relied-upon animals, used for transport, cavalry, and communication across the Western Front." Now I have narrowed to horses in WWI specifically
- Finally, my Thesis Statement: "While the cost to these animals was undeniably high, horses were an essential part of the war effort, and their use was justified by the lives they helped save." There it is -- my position, clearly stated
- Notice: I used "undeniably" (high modality) and "essential" (strong vocabulary). I acknowledged the other side ("the cost was high") before stating my position

DO:
- Write or display the paragraph as you narrate each sentence
- Point to the GST labels as you write each component
- Think aloud about word choices: why "vital" not "important," why "undeniably" not "very"
- Highlight the concession ("the cost was high") -- this is a sophisticated move

TEACHER NOTES:
The teacher models arguing FOR throughout the unit. This gives students a consistent reference point. The model includes a concession (acknowledging the counter-argument) which is a higher-order persuasive move students can aspire to.

MISCONCEPTIONS:
- Misconception: The thesis must be the first sentence
  Why: Students are used to "topic sentences" at the start of paragraphs
  Impact: They skip the funnel structure and jump straight to their opinion, losing the persuasive build-up
  Quick correction: "The thesis comes LAST in the introduction. Build up to it -- hook first, then narrow, then state your case"

WATCH FOR:
- Students who want to write the same paragraph as the teacher -- redirect: "You will write your OWN position. Mine is a model for the structure and language, not the content"
- Students who don't notice the language features -- pause and highlight: "What kind of language is 'undeniably'?"

[General: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU = `SAY:
- Quick check before you plan your own. I am going to read you a sentence. You tell me whether it is the G, the S, or the T
- Listen: "The use of horses in World War 1 was a cruel and unnecessary decision that led to the deaths of millions of animals"
- Is that the General statement, the Specific statement, or the Thesis statement?
- Hold up: 1 finger for G, 2 for S, 3 for T
- Three, two, one -- show me! [Scan]

DO:
- Use Finger Voting
- Scan for: 3 fingers (Thesis -- it states a clear position with emotive language)
- If mostly correct, move to We Do planning. If not, reteach

CFU CHECKPOINT:
Technique: Finger Voting

Script:
- "Hold up 1 for General, 2 for Specific, 3 for Thesis"
- Read the sentence, then count down
- Scan for: mostly 3 fingers (Thesis Statement)

PROCEED (>=80%): Most show 3 fingers. Move to We Do.
PIVOT (<80%): Most likely issue -- students confuse Thesis with Specific because both are about horses in WWI. Reteach: "The Specific statement gives INFORMATION about the topic. The Thesis states your OPINION. Listen for opinion words: 'cruel,' 'unnecessary' -- those are the writer's judgement. That makes it a Thesis." Re-check: "Horses were among the most relied-upon animals during WWI." G, S, or T? [S -- information, no opinion]

TEACHER NOTES:
The chosen sentence contains strong opinion language ("cruel," "unnecessary") which should make it identifiable as a Thesis. The most common confusion is S vs T.

WATCH FOR:
- Students who hold up 2 (Specific) -- they may be hearing "horses in WWI" and not noticing the opinion language
- Students who hold up 1 (General) -- they may not understand what "general" means in this context

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- The answer is T -- Thesis Statement. This sentence states a clear POSITION: that using horses was "cruel and unnecessary"
- The giveaway is the opinion language: "cruel," "unnecessary." A General or Specific statement gives information. A Thesis gives your position
- Remember the funnel: G is wide, S narrows, T is your point

DO:
- Reveal the answer
- Brief reinforcement, then transition to planning

[General: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_WEDO = `SAY:
- Now let's plan your introductory paragraph together. Use the GST Planning Template
- First, write your topic at the top
- For your G -- General Statement -- think broadly. What can you say about animals and war in general? Talk with your partner for 30 seconds, then I will take ideas
- [Take 2-3 ideas] Good. Now narrow down for your S -- Specific Statement. What specifically about horses in WWI? 30 seconds with your partner
- [Take 2-3 ideas] Now your T -- Thesis Statement. State YOUR position clearly. Remember: opinion language
- Try to include a rhetorical question somewhere -- it could be your opening hook instead of a general statement

DO:
- Distribute the Session 5 GST Planning Template
- Time each component: 30 seconds pair talk + 30 seconds writing per section
- Cold Call after each section to share strong examples
- Circulate to check plans are on track

CFU CHECKPOINT:
Technique: Cold Call (paired oral planning)

Script:
- After the T section, cold call 3 pairs: "Read me your thesis statement"
- Listen for: clear position + opinion language
- "Does that sentence tell me what you BELIEVE? Can I tell which side you are on?"

PROCEED (>=80%): Most thesis statements clearly state a position. Move to You Do.
PIVOT (<80%): Most likely issue -- thesis statements that describe instead of argue. Reteach: "Add an opinion word. Instead of 'Horses were used in WWI,' say 'Horses SHOULD NEVER have been used in WWI because...' The 'should never' is your position." Re-check by cold calling 2 more pairs.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the Session 5 Enabling Scaffold which has sentence starters pre-filled for each GST component. Students complete the sentences rather than writing from scratch
- Extra Notes: Pair enabling students with a confident partner for the planning phase

EXTENDING PROMPT:
- Task: Include a concession in the thesis statement -- acknowledge the OTHER side before stating your position (e.g., "While some argue that..., it is clear that...")
- Extra Notes: Point to the teacher's model as an example of a concession

TEACHER NOTES:
Collaborative planning with targeted pair talk ensures all students have a plan before writing independently. The enabling scaffold reduces the blank-page barrier.

WATCH FOR:
- Students whose G is too specific (already about horses) -- push: "Go wider. What about animals and war in general?"
- Students whose T sounds like an S (informational, no opinion) -- add: "What do you BELIEVE? Add that word"

[General: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Here are some strong examples of each component
- Notice how the G starts broad, the S narrows, and the T states a clear position
- If your plan looks different, that is fine -- there are many ways to write a strong introduction. What matters is the funnel structure

DO:
- Display the model examples
- Briefly highlight what makes each one effective
- Transition: "Now you are going to use your plan to write a complete introductory paragraph"

[General: We Do Reveal | VTLM 2.0: Guided Practice]`;

const NOTES_YOUDO = `SAY:
- Time to write. Use your GST plan to write a complete introductory paragraph
- First: write your General Statement -- broad, hooks the reader
- Next: write your Specific Statement -- narrow to horses in WWI
- Then: write your Thesis Statement -- state your position clearly with strong language
- Try to include at least one rhetorical question OR one example of emotive language
- You have 10 minutes. Write in your booklet. When you finish, read it back and check: can the reader tell which side you are on?

DO:
- Set a 10-minute timer
- Circulate -- prioritise students who struggled during the We Do
- Check: Does each paragraph follow the GST structure? Is the position clear?

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Students use the Enabling Scaffold to complete sentence starters for each GST component, then copy their completed sentences into a paragraph in their booklet
- Extra Notes: Allow enabling students to write 2 sentences (G+T or S+T) rather than requiring all 3

EXTENDING PROMPT:
- Task: Write a SECOND introductory paragraph arguing the OPPOSITE position. Then compare: which introduction is more persuasive and why?

TEACHER NOTES:
Independent writing with the GST plan as scaffold. The review step ("read it back") builds self-monitoring habits.

WATCH FOR:
- Students who write one long sentence instead of the funnel structure -- show them where to put full stops
- Students who forget the thesis -- their paragraph describes but doesn't argue
- Students who finish early -- direct to review: "Read it aloud to your partner. Can they tell your position?"

[General: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Let's review our success criteria. SC1: three parts of a persuasive introduction -- what are they? [G, S, T]
- Thumbs check for SC1. [scan] SC2 -- you planned your intro? [scan] SC3 -- you wrote it with rhetorical question or emotive language? [scan]
- Turn and talk: Read your thesis statement to your partner. Your partner tells you whether your position is crystal clear

DO:
- Quick oral recall of GST before thumbs check
- The turn-and-talk focuses specifically on the thesis statement quality
- Preview: "Next week we move into writing body paragraphs -- the arguments with evidence"

TEACHER NOTES:
The turn-and-talk provides peer feedback on the thesis, which is the hardest component. Students who are "down" on SC1 need the GST structure reference from Lesson 19.

WATCH FOR:
- Students "thumbs down" on SC3 -- they may not have finished writing. Check if it is a time issue or a skill issue
- Partners who just say "yes, it is clear" without really checking -- push: "What SIDE are they on? How do you know?"

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Three printable resources for today
- The ${PLANNING_RESOURCE.name} is your planning tool for the We Do and You Do
- The ${ENABLING_RESOURCE.name} has sentence starters to help you get started
- The ${ANSWER_KEY_RESOURCE.name} is for teacher reference -- model paragraphs for both positions

DO:
- Print the GST Planning Template before the lesson (one per student)
- Print the Enabling Scaffold (for identified students)
- Print the Answer Key (teacher copy only)
- Click any resource card to open the PDF

TEACHER NOTES:
The enabling scaffold should be discreetly provided to students who need it, not distributed to the whole class. The answer key shows both positions so the teacher can support students arguing either side.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "War Horse - Lesson 20 - Writing a Persuasive Introduction";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "War Horse",
    "Writing a Persuasive Introduction",
    "Lesson 20  |  Week 4  |  Year 5/6 Literacy",
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
  // SLIDE 2 -- Learning Intention & Success Criteria
  // =========================================================================
  liSlide(
    pres,
    [
      "We are learning to plan and write an introductory paragraph for a persuasive text using the GST structure and rhetorical devices",
    ],
    [
      "I can identify the three parts of a persuasive introduction: General, Specific, and Thesis",
      "I can plan an introductory paragraph that clearly states my position on the topic",
      "I can write an introductory paragraph using a rhetorical question or emotive language",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Reading Launch: War Horse Facts
  // =========================================================================
  contentSlide(
    pres,
    "Read Aloud",
    C.PRIMARY,
    "War Horse Facts",
    [
      "Reading Mode: Student Read Aloud",
      "Source: The Brooke -- War Horse Facts",
      "Significant details about how horses were used in both World Wars",
      "As you read, take notes on facts that strengthen YOUR argument",
    ],
    NOTES_READING,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Quick Revision
  // =========================================================================
  contentSlide(
    pres,
    "Review",
    C.SECONDARY,
    "Quick Revision: Our Persuasive Topic",
    [
      "Topic: Should horses have been used in World War 1?",
      "Last lesson we brainstormed arguments FOR and AGAINST",
      "We learned the macrostructure: Introduction, Body Paragraphs, Conclusion",
      "Today we zoom in on the INTRODUCTION",
    ],
    NOTES_REVISION,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- GST Structure (I Do)
  // =========================================================================
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Introductory Paragraph: GST Structure",
    [
      "G -- General Statement: Broad opening about the topic (no opinion yet)",
      "S -- Specific Statement: Narrow to the specific issue",
      "T -- Thesis Statement: State YOUR position clearly",
    ],
    NOTES_GST,
    FOOTER,
    (slide, layoutGuide) => {
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW || 4.3;
      const topY = layoutGuide.panelTopPadded;

      // Funnel visual -- three blocks narrowing
      const gW = rw;
      const sW = rw * 0.75;
      const tW = rw * 0.50;

      // G block (full width)
      addCard(slide, rx, topY, gW, 0.9, { fill: C.PRIMARY });
      slide.addText("G -- GENERAL\n\"Throughout history, animals have played a vital role in human conflict...\"", {
        x: rx + 0.1, y: topY + 0.08, w: gW - 0.2, h: 0.74,
        fontSize: 10, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });

      // S block (narrower)
      const sX = rx + (gW - sW) / 2;
      addCard(slide, sX, topY + 1.02, sW, 0.9, { fill: C.SECONDARY });
      slide.addText("S -- SPECIFIC\n\"During WWI, horses were used for transport, cavalry, and communication...\"", {
        x: sX + 0.1, y: topY + 1.10, w: sW - 0.2, h: 0.74,
        fontSize: 10, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });

      // T block (narrowest)
      const tX = rx + (gW - tW) / 2;
      addCard(slide, tX, topY + 2.04, tW, 0.9, { fill: C.ACCENT });
      slide.addText("T -- THESIS\n\"Horses were essential and their use was justified...\"", {
        x: tX + 0.1, y: topY + 2.12, w: tW - 0.2, h: 0.74,
        fontSize: 10, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });

      // Funnel label
      slide.addText("Funnel: broad to specific to YOUR position", {
        x: rx, y: topY + 3.06, w: rw, h: 0.26,
        fontSize: 10, fontFace: FONT_B, color: C.MUTED, italic: true, align: "center", margin: 0,
      });
    }
  );

  // =========================================================================
  // SLIDE 6 -- Rhetorical Questions & Emotive Language
  // =========================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "I Do", { color: C.PRIMARY });
    addTitle(s, "Rhetorical Questions & Emotive Language");

    // Rhetorical Questions card
    const rqY = CONTENT_TOP;
    addCard(s, 0.5, rqY, 9, 1.7, { strip: C.PRIMARY, fill: C.WHITE });
    s.addShape("roundRect", {
      x: 0.7, y: rqY + 0.12, w: 2.6, h: 0.30, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    });
    s.addText("Rhetorical Question", {
      x: 0.7, y: rqY + 0.12, w: 2.6, h: 0.30,
      fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("A question that makes the reader think -- no answer expected", {
      x: 3.5, y: rqY + 0.12, w: 5, h: 0.30,
      fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
    });
    s.addText("\"How could we send thousands of innocent animals to die in a war they could never understand?\"", {
      x: 0.75, y: rqY + 0.60, w: 8.4, h: 0.90,
      fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
      valign: "middle",
    });

    // Emotive Language card
    const elY = rqY + 1.84;
    addCard(s, 0.5, elY, 9, 1.7, { strip: C.ACCENT, fill: C.WHITE });
    s.addShape("roundRect", {
      x: 0.7, y: elY + 0.12, w: 2.4, h: 0.30, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    });
    s.addText("Emotive Language", {
      x: 0.7, y: elY + 0.12, w: 2.4, h: 0.30,
      fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("Words chosen to make the reader FEEL strong emotions", {
      x: 3.3, y: elY + 0.12, w: 5, h: 0.30,
      fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
    });
    s.addText("\"Innocent horses were exploited, injured, and abandoned in the horrific conditions of World War 1.\"", {
      x: 0.75, y: elY + 0.60, w: 8.4, h: 0.90,
      fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
      valign: "middle",
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_RHETORICAL);
  }

  // =========================================================================
  // SLIDE 7 -- Teacher Models Introductory Paragraph (I Do)
  // =========================================================================
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Model Introductory Paragraph",
    "Teacher's position: FOR\n\nG: \"Throughout history, animals have played a vital role in human conflict, serving alongside soldiers in ways that shaped the outcome of wars.\"\n\nS: \"During World War 1, horses were among the most relied-upon animals, used for transport, cavalry, and communication across the Western Front.\"\n\nT: \"While the cost to these animals was undeniably high, horses were an essential part of the war effort, and their use was justified by the lives they helped save.\"",
    null,
    NOTES_MODEL,
    FOOTER
  );

  // =========================================================================
  // SLIDES 8-9 -- CFU: Identify GST (withReveal)
  // =========================================================================
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "G, S, or T?",
      "Finger Voting",
      "Listen to this sentence:\n\n\"The use of horses in World War 1 was a cruel and unnecessary decision that led to the deaths of millions of animals.\"\n\nHold up:\n1 finger = General Statement\n2 fingers = Specific Statement\n3 fingers = Thesis Statement",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      const ansY = 4.0;
      slide.addShape("roundRect", {
        x: 0.5, y: ansY, w: 9, h: 1.0, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      });
      slide.addShape("roundRect", {
        x: 0.7, y: ansY + 0.10, w: 2.6, h: 0.28, rectRadius: 0.08,
        fill: { color: C.WHITE },
      });
      slide.addText("THESIS STATEMENT", {
        x: 0.7, y: ansY + 0.10, w: 2.6, h: 0.28,
        fontSize: 12, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addText("This sentence states a clear POSITION: \"cruel and unnecessary\" are opinion words. A General or Specific statement gives information without judgement. The Thesis states what the writer BELIEVES.", {
        x: 3.5, y: ansY + 0.08, w: 5.8, h: 0.84,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // =========================================================================
  // SLIDES 10-11 -- We Do: Plan Your Introductory Paragraph (withReveal)
  // =========================================================================
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.SECONDARY);
      addBadge(s, "We Do", { color: C.SECONDARY });
      addTitle(s, "Plan Your Introductory Paragraph");

      // GST planning prompts
      const sections = [
        { label: "G -- General Statement", prompt: "What can you say broadly about animals and war? Write a sentence that hooks the reader.", color: C.PRIMARY },
        { label: "S -- Specific Statement", prompt: "Narrow down: What specifically about horses in World War 1?", color: C.SECONDARY },
        { label: "T -- Thesis Statement", prompt: "State YOUR position. Which side are you on? Use strong opinion language.", color: C.ACCENT },
      ];

      sections.forEach((sec, i) => {
        const sy = CONTENT_TOP + i * 1.22;
        addCard(s, 0.5, sy, 9, 1.1, { strip: sec.color, fill: C.WHITE });
        s.addShape("roundRect", {
          x: 0.7, y: sy + 0.12, w: 2.8, h: 0.30, rectRadius: 0.08,
          fill: { color: sec.color },
        });
        s.addText(sec.label, {
          x: 0.7, y: sy + 0.12, w: 2.8, h: 0.30,
          fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
        s.addText(sec.prompt, {
          x: 0.75, y: sy + 0.52, w: 8.4, h: 0.48,
          fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO);
      return s;
    },
    (slide) => {
      // Reveal: example completions
      const examples = [
        { y: CONTENT_TOP + 0.52, text: "e.g., \"For centuries, animals have been caught up in human conflicts, often paying the highest price.\"" },
        { y: CONTENT_TOP + 1.74, text: "e.g., \"In World War 1, over one million horses served on the Western Front, facing conditions no animal should endure.\"" },
        { y: CONTENT_TOP + 2.96, text: "e.g., \"The use of horses in this war was a devastating act of cruelty that can never be justified.\"" },
      ];
      examples.forEach((ex) => {
        slide.addShape("roundRect", {
          x: 0.7, y: ex.y, w: 8.5, h: 0.48, rectRadius: 0.06,
          fill: { color: C.BG_LIGHT },
        });
        slide.addText(ex.text, {
          x: 0.8, y: ex.y + 0.02, w: 8.2, h: 0.44,
          fontSize: 11.5, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
        });
      });
      slide.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // =========================================================================
  // SLIDE 12 -- You Do: Write Your Introductory Paragraph
  // =========================================================================
  taskSlide(
    pres,
    "You Do",
    "Write Your Introductory Paragraph",
    [
      { label: "First", instruction: "Write your General Statement -- broad, hooks the reader, no opinion yet." },
      { label: "Next", instruction: "Write your Specific Statement -- narrow to horses in WWI." },
      { label: "Then", instruction: "Write your Thesis Statement -- state YOUR position clearly with strong language. Try to include a rhetorical question or emotive language." },
    ],
    NOTES_YOUDO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 13 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    "Read your thesis statement to your partner. Your partner tells you: can they tell which side you are on? Is the language strong enough to persuade?",
    [
      "I can identify the three parts of a persuasive introduction: General, Specific, and Thesis",
      "I can plan an introductory paragraph that clearly states my position",
      "I can write an introductory paragraph using rhetorical questions or emotive language",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 14 -- Resources
  // =========================================================================

  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: GST Planning Template ----------------------------------------
  const pt = createPdf({ title: PLANNING_RESOURCE.name });
  let ptY = addPdfHeader(pt, "GST Planning Template", {
    color: C.PRIMARY,
    subtitle: "Plan your introductory paragraph for: Should horses have been used in World War 1?",
    lessonInfo: "War Horse | Lesson 20 | Week 4 | Year 5/6 Literacy",
    showNameDate: true,
  });

  ptY = addTipBox(pt, "Use this template to plan your introductory paragraph. Each section builds on the one before -- start broad (General), narrow down (Specific), then state your position (Thesis).", ptY, { color: C.PRIMARY });

  ptY = addSectionHeading(pt, "My Position (circle one):    FOR    /    AGAINST", ptY, { color: C.PRIMARY });
  ptY += 8;

  ptY = addSectionHeading(pt, "G -- General Statement", ptY, { color: C.PRIMARY });
  ptY = addBodyText(pt, "Write a broad opening sentence about animals in war. Hook the reader. Do NOT state your opinion yet.", ptY, { fontSize: 10, italic: true });
  ptY = addLinedArea(pt, ptY, 3, { lineSpacing: 26 });
  ptY += 10;

  ptY = addSectionHeading(pt, "S -- Specific Statement", ptY, { color: C.SECONDARY });
  ptY = addBodyText(pt, "Narrow down to horses in World War 1 specifically. Give information about the topic.", ptY, { fontSize: 10, italic: true });
  ptY = addLinedArea(pt, ptY, 3, { lineSpacing: 26 });
  ptY += 10;

  ptY = addSectionHeading(pt, "T -- Thesis Statement", ptY, { color: C.ACCENT });
  ptY = addBodyText(pt, "State YOUR position clearly. Use strong opinion language (e.g., \"essential,\" \"cruel,\" \"unjustifiable\"). This is where the reader learns what YOU believe.", ptY, { fontSize: 10, italic: true });
  ptY = addLinedArea(pt, ptY, 3, { lineSpacing: 26 });
  ptY += 10;

  ptY = addSectionHeading(pt, "Language Features Checklist", ptY, { color: C.ACCENT });
  ptY = addBodyText(pt, "Tick the features you used in your introduction:", ptY, { fontSize: 10 });
  ptY = addBodyText(pt, "[ ] Rhetorical question", ptY, { fontSize: 11 });
  ptY = addBodyText(pt, "[ ] Emotive language", ptY, { fontSize: 11 });
  ptY = addBodyText(pt, "[ ] High modality language", ptY, { fontSize: 11 });
  ptY = addBodyText(pt, "[ ] Tier 2 or 3 vocabulary", ptY, { fontSize: 11 });

  addPdfFooter(pt, "War Horse | Lesson 20 | GST Planning Template");

  // --- PDF 2: Enabling Scaffold -------------------------------------------
  const es = createPdf({ title: ENABLING_RESOURCE.name });
  let esY = addPdfHeader(es, "Introductory Paragraph -- Enabling Scaffold", {
    color: C.SECONDARY,
    subtitle: "Sentence starters to help you plan and write your introduction",
    lessonInfo: "War Horse | Lesson 20 | Week 4 | Year 5/6 Literacy",
    showNameDate: true,
  });

  esY = addTipBox(es, "Complete each sentence starter below. Then copy your completed sentences into a paragraph in your booklet.", esY, { color: C.SECONDARY });

  esY = addSectionHeading(es, "My Position (circle one):    FOR    /    AGAINST", esY, { color: C.PRIMARY });
  esY += 6;

  esY = addSectionHeading(es, "G -- General Statement", esY, { color: C.PRIMARY });
  esY = addBodyText(es, "Complete this sentence:", esY, { fontSize: 10, italic: true });
  esY = addBodyText(es, "\"Throughout history, animals have been used in wars because...\"", esY, { fontSize: 12 });
  esY = addLinedArea(es, esY, 2, { lineSpacing: 26 });
  esY += 8;

  esY = addSectionHeading(es, "S -- Specific Statement", esY, { color: C.SECONDARY });
  esY = addBodyText(es, "Complete this sentence:", esY, { fontSize: 10, italic: true });
  esY = addBodyText(es, "\"During World War 1, horses were used to...\"", esY, { fontSize: 12 });
  esY = addLinedArea(es, esY, 2, { lineSpacing: 26 });
  esY += 8;

  esY = addSectionHeading(es, "T -- Thesis Statement (choose FOR or AGAINST)", esY, { color: C.ACCENT });
  esY = addBodyText(es, "If you are arguing FOR, complete:", esY, { fontSize: 10, italic: true });
  esY = addBodyText(es, "\"I believe horses needed to be used in World War 1 because...\"", esY, { fontSize: 12 });
  esY = addLinedArea(es, esY, 2, { lineSpacing: 26 });
  esY += 4;
  esY = addBodyText(es, "If you are arguing AGAINST, complete:", esY, { fontSize: 10, italic: true });
  esY = addBodyText(es, "\"I believe horses should never have been used in World War 1 because...\"", esY, { fontSize: 12 });
  esY = addLinedArea(es, esY, 2, { lineSpacing: 26 });
  esY += 10;

  esY = addTipBox(es, "Now copy your three sentences into your booklet as one paragraph. Read it back: can you tell which side you are on?", esY, { color: C.ACCENT });

  addPdfFooter(es, "War Horse | Lesson 20 | Enabling Scaffold");

  // --- PDF 3: Answer Key ---------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Introductory Paragraph -- Answer Key", {
    color: C.ALERT,
    subtitle: "Teacher reference: model introductory paragraphs (FOR and AGAINST)",
    lessonInfo: "War Horse | Lesson 20 | Week 4 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "These are model answers showing both positions. Use them to support students arguing either side. Each paragraph follows the GST structure.", akY, { color: C.ALERT });

  akY = addSectionHeading(ak, "Model 1: FOR (Horses should have been used)", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "[G] Throughout history, animals have played a vital role in human conflict, serving alongside soldiers in ways that shaped the outcome of wars.", akY);
  akY = addBodyText(ak, "[S] During World War 1, horses were among the most relied-upon animals, used for transport, cavalry, and communication across the Western Front.", akY);
  akY = addBodyText(ak, "[T] While the cost to these animals was undeniably high, horses were an essential part of the war effort, and their use was justified by the lives they helped save.", akY);
  akY += 6;
  akY = addBodyText(ak, "Features used: high modality (\"undeniably,\" \"essential\"), concession (\"the cost was high\"), Tier 2 vocabulary (\"vital,\" \"relied-upon\").", akY, { fontSize: 10, italic: true });
  akY += 12;

  akY = addSectionHeading(ak, "Model 2: AGAINST (Horses should NOT have been used)", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "[G] How could we justify sending millions of innocent animals into a conflict they could never understand?", akY);
  akY = addBodyText(ak, "[S] During World War 1, over one million horses served on the Western Front, and the vast majority never returned home.", akY);
  akY = addBodyText(ak, "[T] The use of horses in this war was a devastating act of cruelty -- no military advantage could ever justify the suffering inflicted on these animals.", akY);
  akY += 6;
  akY = addBodyText(ak, "Features used: rhetorical question (opening), emotive language (\"innocent,\" \"devastating,\" \"cruelty\"), high modality (\"never,\" \"no... could ever\"), Tier 2 vocabulary (\"inflicted\").", akY, { fontSize: 10, italic: true });
  akY += 14;

  akY = addSectionHeading(ak, "What to Look For in Student Work", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- G is broad (about animals/war in general), not already about horses in WWI", akY);
  akY = addBodyText(ak, "- S narrows to horses in WWI specifically, gives factual information", akY);
  akY = addBodyText(ak, "- T clearly states a position with opinion language (not just information)", akY);
  akY = addBodyText(ak, "- At least one language feature is attempted (rhetorical question, emotive language, or high modality)", akY);
  akY = addBodyText(ak, "- The funnel structure is visible: broad to narrow to specific position", akY);

  addPdfFooter(ak, "War Horse | Lesson 20 | Answer Key -- TEACHER COPY");

  // --- Write all files ----------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/WH4_Lesson20.pptx` }),
    writePdf(pt, PLANNING_PDF_PATH),
    writePdf(es, ENABLING_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  const pptxPath = `${OUT_DIR}/WH4_Lesson20.pptx`;
  console.log("PPTX written to " + pptxPath);
  console.log(`Done: ${PLANNING_RESOURCE.name}.pdf`);
  console.log(`Done: ${ENABLING_RESOURCE.name}.pdf`);
  console.log(`Done: ${ANSWER_KEY_RESOURCE.name}.pdf`);
}

build().catch(console.error);
