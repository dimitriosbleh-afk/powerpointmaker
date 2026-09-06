"use strict";

// War Horse Unit -- Lesson 19: Plan a Persuasive Text
// Week 4, Session 4, Grade 5/6 Literacy
// Supplementary text: Animals in WWI → Brainstorm → Persuasive text structure + features

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
  addTwoColumnOrganiser,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 4;
const FOOTER = "War Horse | Lesson 19 of 25 | Week 4 | Year 5/6 Literacy";
const OUT_DIR = "output/WH4_Lesson19_Plan_Persuasive_Text";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PLANNING_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Persuasive Planning Organiser",
  "Graphic organiser: For/Against arguments + position statement. Students brainstorm and record their chosen side."
);
const STRUCTURE_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Persuasive Text Structure Reference",
  "Student reference: macrostructure of a persuasive text + key language features with definitions and examples."
);
const RESOURCE_ITEMS = [PLANNING_RESOURCE, STRUCTURE_RESOURCE];
const PLANNING_PDF_PATH = path.join(OUT_DIR, PLANNING_RESOURCE.fileName);
const STRUCTURE_PDF_PATH = path.join(OUT_DIR, STRUCTURE_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today we are moving from reading War Horse to writing about it
- We have been following Joey's story and learning about animals in World War 1. Now we are going to use what we know to plan a persuasive text
- The big question: Should horses have been used in World War 1?

DO:
- Display title slide as students settle
- Have the BBC Bitesize article ready on screen or printed for student access

TEACHER NOTES:
Lesson 19 of 25, fourth session of Week 4. This lesson transitions from novel reading to persuasive writing. The supplementary text provides factual context about animals in WWI that feeds directly into the persuasive planning task.

WATCH FOR:
- Students who are unclear about the shift from narrative to persuasive mode -- frame it clearly: "We are still learning from War Horse, but now we are writing our own opinion"

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Today is about planning, not writing yet. We are building up our knowledge and arguments
- Read the success criteria together. Ask: Which one do you think will be the most challenging? [Take 2-3 responses]

DO:
- Choral read the LI, then the SCs
- Emphasise: today is a planning lesson -- the writing comes in the next sessions

TEACHER NOTES:
SC1 grounds arguments in content knowledge. SC2 ensures students understand the writing form. SC3 moves them toward commitment and planning.

WATCH FOR:
- Students who think "persuasive" means "right answer" -- clarify: you can argue either side
- Students who are eager to start writing -- redirect: "Great enthusiasm, but strong writing starts with strong planning"

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_READING = `SAY:
- Before we plan our persuasive text, we need more information about how animals were used in World War 1
- We are reading a BBC Bitesize article: "Why were animals useful in World War 1?"
- Reading mode is Student Read Aloud -- I will select readers
- As you read, take notes in your booklet. Write down facts that could support either side of the argument

DO:
- Display the reading task and focus question
- Select readers for the article (change every paragraph)
- If the video is available, play it after reading the text
- Circulate during note-taking: check students are recording useful facts, not copying sentences

TEACHER NOTES:
The supplementary text provides factual grounding for the persuasive task. Students need content knowledge before they can argue a position. Note-taking during reading builds research skills.

WATCH FOR:
- Students who are reading but not taking notes -- prompt: "What is one fact you could use in an argument?"
- Students who copy whole sentences instead of note-taking -- model: "Write the key idea in your own words"

[General: Reading Launch | VTLM 2.0: Structured Reading / Research Skills]`;

const NOTES_TOPIC = `SAY:
- Now we know more about animals in the war. Here is the question we will be writing about
- Read from slide: "Should horses have been used in World War 1?"
- This is a genuine debate. There are strong arguments on both sides
- Some of you may already have an opinion -- that is great. But before you commit to a side, let's look at arguments for AND against

DO:
- Display the topic clearly
- Let the question land -- give 5 seconds of silence before continuing
- Emphasise: even if you have an opinion, exploring both sides makes your argument stronger

TEACHER NOTES:
Presenting both sides before students choose a position is deliberate. Understanding the counterargument strengthens persuasive writing because the writer can anticipate and rebut objections.

WATCH FOR:
- Students who immediately declare a side -- acknowledge but redirect: "Hold that thought -- let's build the full picture first"

[General: I Do -- Topic Introduction | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_BRAINSTORM = `SAY:
- Let's brainstorm arguments for both sides together
- For: What reasons might someone give for using horses in WWI? [Take 3-4 responses]
- Against: What reasons might someone give against using horses? [Take 3-4 responses]
- Some examples if needed -- horses were needed for transport and mounted infantry, but there was massive loss of animal life and many couldn't be brought home

DO:
- Display the For/Against framework
- Use Cold Call to gather student ideas first, then fill in any gaps with the example arguments
- Record student contributions visually on the board or slide if possible

CFU CHECKPOINT:
Technique: Cold Call

Script:
- "Give me one argument FOR using horses in World War 1. [Name], what have you got?"
- "Now one argument AGAINST. [Name]?"
- Scan for: students generating genuine arguments, not just restating facts
- "Can anyone give me an argument we haven't heard yet?"

PROCEED (>=80%): Students can generate at least one argument per side. Move to macrostructure.
PIVOT (<80%): Most likely issue -- students confuse facts with arguments. Reteach: "A fact tells us what happened. An argument tells us what SHOULD have happened. 'Horses carried ammunition' is a fact. 'Horses were essential because no vehicles could cross the terrain' is an argument." Re-check: "Turn that fact into an argument."

TEACHER NOTES:
Cold Call ensures broad participation. Brainstorming both sides first prevents students from locking into a position before they have considered the full picture.

WATCH FOR:
- Students who can only argue one side -- push: "What would someone on the other side say?"
- Students who confuse facts with arguments -- model the difference explicitly

[General: We Do -- Brainstorm | VTLM 2.0: Collaborative Discussion]`;

const NOTES_STRUCTURE = `SAY:
- Now let's look at how a persuasive text is built. Every persuasive text follows a macrostructure
- Introduction: hook the reader, state your position
- Body paragraphs: each one presents a separate argument with evidence and explanation
- Conclusion: restate your position strongly, call to action
- Think of it like building a case in court. You open with your position, present your evidence, then close with a strong final statement

DO:
- Walk through each section of the structure on the left
- Point to the structural diagram on the right as you explain each part
- Keep this brisk -- it is a preview, not deep teaching. Deeper work on each section comes in following lessons

TEACHER NOTES:
The macrostructure is introduced here as an overview. Lessons 20+ will teach each section in detail. Students need to see the whole shape before zooming in.

WATCH FOR:
- Students who want to start writing immediately -- redirect: "We need to understand the whole structure first"
- Students who seem overwhelmed by the structure -- reassure: "We will work on one section at a time in coming lessons"

[General: I Do -- Text Structure | VTLM 2.0: Explicit Teaching]`;

const NOTES_FEATURES = `SAY:
- Persuasive texts use specific language features to convince the reader. Here are five key ones
- Rhetorical questions: questions you don't expect an answer to -- they make the reader think. "How could we possibly justify such suffering?"
- Emotive language: words chosen to make the reader FEEL something -- "innocent animals were forced into a horrific war"
- High modality: strong, definite language -- "must," "certainly," "undoubtedly," "essential"
- Nominalisations: turning actions into things -- "destroy" becomes "destruction," "suffer" becomes "suffering"
- Tier 2 and 3 vocabulary: precise, powerful words -- "devastating" instead of "bad," "exploitation" instead of "use"

DO:
- Display each feature with its example
- Read each example aloud with expression to show how they sound in a persuasive text
- Quick oral check: "Give me a thumbs up if you have heard of rhetorical questions before" [scan]

TEACHER NOTES:
This is a preview, not deep teaching. Each feature will be taught explicitly in later lessons (Lesson 20 focuses on rhetorical questions and emotive language). Students need to see the toolkit before choosing which tools to use.

WATCH FOR:
- Students who are unfamiliar with nominalisations -- this is the hardest concept and will be taught in depth later
- Students who can identify the features but not explain why they are persuasive -- model: "Emotive language works because it makes the reader feel what you feel"

[General: I Do -- Language Features | VTLM 2.0: Explicit Teaching]`;

const NOTES_CFU = `SAY:
- Quick check. I am going to describe a section of a persuasive text. You tell me which part of the macrostructure it belongs to
- "The writer presents one reason why their position is correct, gives evidence from research, and explains how the evidence supports their point." Which section is this?
- Hold up 1 finger for Introduction, 2 for Body Paragraph, 3 for Conclusion
- Three, two, one -- show me! [Scan fingers]

DO:
- Use Finger Voting -- students hold up fingers on your count
- Scan for: 2 fingers (Body Paragraph)
- If mostly correct, move to You Do. If not, reteach the distinction

CFU CHECKPOINT:
Technique: Finger Voting

Script:
- "Hold up 1 finger for Introduction, 2 for Body Paragraph, 3 for Conclusion"
- "I describe: The writer presents one reason with evidence and explanation. Which section? Three, two, one -- show me!"
- Scan for: mostly 2 fingers (Body Paragraph)

PROCEED (>=80%): Most show 2 fingers. Move to You Do.
PIVOT (<80%): Most likely issue -- students confuse body paragraph with introduction because both mention the topic. Reteach: "The introduction tells the reader your POSITION. A body paragraph gives ONE REASON with EVIDENCE. The introduction is the promise; the body paragraphs are the proof." Re-check: "The writer hooks the reader and states their opinion clearly. Which section?" [1 -- Introduction]

TEACHER NOTES:
Finger Voting is used here for quick universal response. The question tests structural understanding, not content knowledge.

WATCH FOR:
- Students who hold up 3 (Conclusion) -- they may think "explains" means "wraps up"
- Students who don't commit to an answer -- prompt: "Best guess -- no wrong answers here, I need to see where you are at"

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- The answer is Body Paragraph. That is where you present one argument, back it up with evidence, and explain how the evidence supports your point
- The introduction hooks and states your position. The conclusion restates and pushes for action. The body paragraphs do the heavy lifting

DO:
- Reveal the answer
- Brief acknowledgement, then transition to You Do

[General: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Time to make your choice and start planning
- First: decide which side you want to argue -- for or against horses being used in WWI
- Next: write your position clearly at the top of your planning organiser
- Then: brainstorm at least three arguments that support your position. Use facts from today's reading and our class discussion

DO:
- Distribute the Session 4 Persuasive Planning Organiser
- Set a timer: 8-10 minutes for planning
- Circulate -- check students have committed to a side and are generating arguments, not just copying from the board

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Students who struggle to generate arguments use the planning organiser with the pre-filled examples to prompt their thinking. They choose 2 arguments from the examples provided and add 1 of their own
- Extra Notes: Pair enabling students with a confident partner for the first 2 minutes to talk through ideas before writing

EXTENDING PROMPT:
- Task: Plan arguments for BOTH sides of the debate. For each argument, note what evidence they would use and anticipate what the opposing side would say in response (counter-argument)
- Extra Notes: This prepares them for a more sophisticated persuasive text that acknowledges and rebuts counterarguments

TEACHER NOTES:
Independent planning time. The enabling scaffold on the organiser provides pre-filled examples to lower the barrier. Extending students work on counter-arguments, which is a higher-order persuasive skill.

WATCH FOR:
- Students who can't choose a side -- scaffold: "Which side did you feel more strongly about during our brainstorm? Go with that"
- Students who write only one argument -- push: "You need at least three. What else did we discuss?"
- Students who copy arguments verbatim from the board -- prompt: "Put it in your own words. How would YOU say it?"

[General: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Let's review our success criteria. SC1: arguments for and against -- thumbs up, sideways, or down? [scan]
- SC2: macrostructure and language features of a persuasive text -- thumbs? [scan]
- SC3: your own position with initial arguments -- thumbs? [scan]
- Turn and talk: Tell your partner which side you chose and your strongest argument. Your partner's job is to challenge you with a counter-argument

DO:
- Run through each SC with thumbs check
- The turn-and-talk is deliberately provocative -- partners challenge each other's arguments
- Preview: "Next lesson we start writing. We will focus on the introductory paragraph"

TEACHER NOTES:
The turn-and-talk builds argumentation skills. Being challenged by a partner forces students to think about the strength of their reasoning.

WATCH FOR:
- Students who are "thumbs down" on SC2 -- they may need the Persuasive Text Structure Reference to review before next lesson
- Partners who agree instead of challenging -- redirect: "Your job is to find a weakness in their argument"

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two printable resources for today's lesson
- The ${PLANNING_RESOURCE.name} is your planning tool for the You Do -- use it to record your position and arguments
- The ${STRUCTURE_RESOURCE.name} is a keepsake -- it shows the macrostructure and language features we covered today

DO:
- Print the planning organiser before the lesson (one per student)
- Print the structure reference (one per student -- they keep this for the whole unit)
- Click any resource card to open the PDF

TEACHER NOTES:
The structure reference stays with students throughout the persuasive writing unit. The planning organiser feeds directly into Lessons 20-25.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "War Horse - Lesson 19 - Plan a Persuasive Text";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "War Horse",
    "Planning a Persuasive Text",
    "Lesson 19  |  Week 4  |  Year 5/6 Literacy",
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
      "We are learning to research a topic and plan the structure and arguments for a persuasive text",
    ],
    [
      "I can identify arguments for and against using horses in World War 1",
      "I can explain the macrostructure and key language features of a persuasive text",
      "I can plan my own position with initial arguments for a persuasive writing task",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Reading Launch: Supplementary Text
  // =========================================================================
  contentSlide(
    pres,
    "Read Aloud",
    C.PRIMARY,
    "Did Animals Need to Be Used During World War 1?",
    [
      "Reading Mode: Student Read Aloud",
      "Source: BBC Bitesize -- Why were animals useful in World War 1?",
      "Read the text and watch the accompanying video",
      "As you read, take notes: What facts could support an argument FOR or AGAINST using animals in war?",
    ],
    NOTES_READING,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Persuasive Writing Topic
  // =========================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, C.ACCENT);
    addBadge(s, "Writing Focus", { color: C.ACCENT, w: 1.8 });
    addTitle(s, "Our Persuasive Writing Topic");

    // Central question card
    const qY = CONTENT_TOP + 0.3;
    addCard(s, 1.0, qY, 8, 1.6, { fill: C.PRIMARY });
    s.addText("Should horses have been used\nin World War 1?", {
      x: 1.2, y: qY + 0.15, w: 7.6, h: 1.3,
      fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Context note
    const noteY = qY + 1.8;
    s.addText("You will argue FOR or AGAINST this question. Today we plan. Writing starts next lesson.", {
      x: 1.0, y: noteY, w: 8, h: 0.5,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_TOPIC);
  }

  // =========================================================================
  // SLIDE 5 -- Brainstorm: For and Against
  // =========================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Brainstorm: For and Against");

    // FOR card
    const cardY = CONTENT_TOP;
    addCard(s, 0.5, cardY, 4.3, 3.6, { strip: C.PRIMARY, fill: C.WHITE });
    s.addShape("roundRect", {
      x: 0.7, y: cardY + 0.12, w: 1.4, h: 0.32, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    });
    s.addText("FOR", {
      x: 0.7, y: cardY + 0.12, w: 1.4, h: 0.32,
      fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText([
      { text: "Horses should have been used:", options: { bold: true, breakLine: true, fontSize: 13, color: C.CHARCOAL } },
      { text: "Needed to transport weapons and ammunition", options: { bullet: true, breakLine: true, fontSize: 12.5, color: C.CHARCOAL } },
      { text: "Crucial for mounted infantry", options: { bullet: true, breakLine: true, fontSize: 12.5, color: C.CHARCOAL } },
      { text: "No motorised vehicles could handle the terrain", options: { bullet: true, fontSize: 12.5, color: C.CHARCOAL } },
    ], {
      x: 0.75, y: cardY + 0.56, w: 3.8, h: 2.9,
      fontFace: FONT_B, valign: "top", margin: 0, paraSpaceAfter: 4,
    });

    // AGAINST card
    addCard(s, 5.2, cardY, 4.3, 3.6, { strip: C.ACCENT, fill: C.WHITE });
    s.addShape("roundRect", {
      x: 5.4, y: cardY + 0.12, w: 1.8, h: 0.32, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    });
    s.addText("AGAINST", {
      x: 5.4, y: cardY + 0.12, w: 1.8, h: 0.32,
      fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText([
      { text: "Horses should NOT have been used:", options: { bold: true, breakLine: true, fontSize: 13, color: C.CHARCOAL } },
      { text: "Additional unnecessary loss of life", options: { bullet: true, breakLine: true, fontSize: 12.5, color: C.CHARCOAL } },
      { text: "Inability to transport them home after the war", options: { bullet: true, breakLine: true, fontSize: 12.5, color: C.CHARCOAL } },
      { text: "Horses suffered in conditions they could not understand", options: { bullet: true, fontSize: 12.5, color: C.CHARCOAL } },
    ], {
      x: 5.45, y: cardY + 0.56, w: 3.8, h: 2.9,
      fontFace: FONT_B, valign: "top", margin: 0, paraSpaceAfter: 4,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_BRAINSTORM);
  }

  // =========================================================================
  // SLIDE 6 -- Macrostructure of a Persuasive Text
  // =========================================================================
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Macrostructure of a Persuasive Text",
    [
      "INTRODUCTION: Hook the reader, introduce the topic, state your position",
      "BODY PARAGRAPHS: Each paragraph presents one argument + evidence + explanation",
      "CONCLUSION: Restate your position strongly, call the reader to action",
    ],
    NOTES_STRUCTURE,
    FOOTER,
    (slide, layoutGuide) => {
      // Visual diagram on the right showing the structure
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW || 4.3;
      const topY = layoutGuide.panelTopPadded;

      // Introduction block
      addCard(slide, rx, topY, rw, 0.8, { fill: C.PRIMARY });
      slide.addText("INTRODUCTION\nHook + Position", {
        x: rx + 0.1, y: topY + 0.1, w: rw - 0.2, h: 0.6,
        fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Arrow
      slide.addText("v", {
        x: rx + rw / 2 - 0.2, y: topY + 0.82, w: 0.4, h: 0.22,
        fontSize: 14, fontFace: FONT_B, color: C.MUTED, align: "center", margin: 0,
      });

      // Body block
      const bodyY = topY + 1.06;
      addCard(slide, rx, bodyY, rw, 1.4, { fill: C.SECONDARY });
      slide.addText("BODY PARAGRAPHS\nArgument 1 + Evidence\nArgument 2 + Evidence\nArgument 3 + Evidence", {
        x: rx + 0.1, y: bodyY + 0.1, w: rw - 0.2, h: 1.2,
        fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Arrow
      slide.addText("v", {
        x: rx + rw / 2 - 0.2, y: bodyY + 1.42, w: 0.4, h: 0.22,
        fontSize: 14, fontFace: FONT_B, color: C.MUTED, align: "center", margin: 0,
      });

      // Conclusion block
      const concY = bodyY + 1.66;
      addCard(slide, rx, concY, rw, 0.8, { fill: C.ACCENT });
      slide.addText("CONCLUSION\nRestate + Call to Action", {
        x: rx + 0.1, y: concY + 0.1, w: rw - 0.2, h: 0.6,
        fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // =========================================================================
  // SLIDE 7 -- Key Language Features
  // =========================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "I Do", { color: C.PRIMARY });
    addTitle(s, "Key Language Features of Persuasive Writing");

    const features = [
      { name: "Rhetorical questions", ex: "\"How could we justify such suffering?\"", color: C.PRIMARY },
      { name: "Emotive language", ex: "\"Innocent animals were forced into a horrific war\"", color: C.SECONDARY },
      { name: "High modality", ex: "\"must,\" \"certainly,\" \"undoubtedly,\" \"essential\"", color: C.ACCENT },
      { name: "Nominalisations", ex: "\"destruction\" (not destroy), \"suffering\" (not suffer)", color: C.PRIMARY },
      { name: "Tier 2 & 3 vocabulary", ex: "\"devastating\" instead of \"bad,\" \"exploitation\" instead of \"use\"", color: C.SECONDARY },
    ];

    const cardH = 0.62;
    const gap = 0.10;
    features.forEach((f, i) => {
      const fy = CONTENT_TOP + i * (cardH + gap);
      if (fy + cardH > SAFE_BOTTOM) return;
      addCard(s, 0.5, fy, 9, cardH, { strip: f.color, fill: C.WHITE });
      s.addText(f.name, {
        x: 0.75, y: fy + 0.06, w: 3.2, h: 0.24,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0,
      });
      s.addText(f.ex, {
        x: 0.75, y: fy + 0.30, w: 8.4, h: 0.26,
        fontSize: 11.5, fontFace: FONT_B, color: C.MUTED, italic: true, margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_FEATURES);
  }

  // =========================================================================
  // SLIDES 8-9 -- CFU (withReveal)
  // =========================================================================
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Which Section Is This?",
      "Finger Voting",
      "\"The writer presents one reason why their position is correct, gives evidence from research, and explains how the evidence supports their point.\"\n\nHold up:\n1 finger = Introduction\n2 fingers = Body Paragraph\n3 fingers = Conclusion",
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
        x: 0.7, y: ansY + 0.10, w: 2.2, h: 0.28, rectRadius: 0.08,
        fill: { color: C.WHITE },
      });
      slide.addText("BODY PARAGRAPH", {
        x: 0.7, y: ansY + 0.10, w: 2.2, h: 0.28,
        fontSize: 12, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addText("The body paragraph is where the heavy lifting happens. Each paragraph presents ONE argument, backs it up with EVIDENCE, and EXPLAINS how the evidence supports your position.", {
        x: 3.1, y: ansY + 0.08, w: 6.2, h: 0.84,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // =========================================================================
  // SLIDE 10 -- You Do: Choose Position and Plan
  // =========================================================================
  taskSlide(
    pres,
    "You Do",
    "Choose Your Position and Plan",
    [
      { label: "First", instruction: "Decide: Will you argue FOR or AGAINST horses being used in WWI? Write your position clearly." },
      { label: "Next", instruction: "Brainstorm at least 3 arguments that support your position. Use facts from today's reading and our discussion." },
      { label: "Then", instruction: "For each argument, note what evidence you could use to back it up." },
    ],
    NOTES_YOUDO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 11 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    "Tell your partner which side you chose and your strongest argument. Your partner's job is to challenge you with a counter-argument.",
    [
      "I can identify arguments for and against using horses in World War 1",
      "I can explain the macrostructure and key language features of a persuasive text",
      "I can plan my own position with initial arguments for a persuasive writing task",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 12 -- Resources
  // =========================================================================

  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: Persuasive Planning Organiser --------------------------------
  const po = createPdf({ title: PLANNING_RESOURCE.name });
  let poY = addPdfHeader(po, "Persuasive Planning Organiser", {
    color: C.PRIMARY,
    subtitle: "Should horses have been used in World War 1?",
    lessonInfo: "War Horse | Lesson 19 | Week 4 | Year 5/6 Literacy",
    showNameDate: true,
  });

  poY = addTipBox(po, "Choose your position: FOR or AGAINST. Then brainstorm arguments for your chosen side. Use facts from the BBC Bitesize article and our class discussion.", poY, { color: C.PRIMARY });

  // Position statement
  poY = addSectionHeading(po, "My Position", poY, { color: C.PRIMARY });
  poY = addBodyText(po, "I believe horses (should / should not) have been used in World War 1 because:", poY);
  poY = addLinedArea(po, poY, 2, { lineSpacing: 26 });
  poY += 8;

  // For/Against organiser
  poY = addTwoColumnOrganiser(po, "ARGUMENTS FOR", "ARGUMENTS AGAINST", poY, {
    color: C.PRIMARY,
    rows: 4,
    rowH: 55,
    leftContent: [
      "e.g., Horses were essential for transporting weapons across terrain vehicles could not handle",
    ],
    rightContent: [
      "e.g., Using horses led to massive unnecessary animal suffering and death",
    ],
  });
  poY += 10;

  // My strongest arguments
  poY = addSectionHeading(po, "My Three Strongest Arguments (for my chosen side)", poY, { color: C.ACCENT });
  for (let i = 1; i <= 3; i++) {
    poY = addBodyText(po, `Argument ${i}:`, poY, { fontSize: 11 });
    poY = addLinedArea(po, poY, 2, { lineSpacing: 24 });
    poY = addBodyText(po, "Evidence I could use:", poY, { fontSize: 10, italic: true });
    poY = addLinedArea(po, poY, 1, { lineSpacing: 24 });
    poY += 4;
  }

  addPdfFooter(po, "War Horse | Lesson 19 | Persuasive Planning Organiser");

  // --- PDF 2: Persuasive Text Structure Reference --------------------------
  const sr = createPdf({ title: STRUCTURE_RESOURCE.name });
  let srY = addPdfHeader(sr, "Persuasive Text Structure Reference", {
    color: C.ACCENT,
    subtitle: "Macrostructure and key language features",
    lessonInfo: "War Horse | Lesson 19 | Year 5/6 Literacy",
    showNameDate: false,
  });

  srY = addTipBox(sr, "Keep this reference sheet for the whole persuasive writing unit. Use it to check your structure and language choices.", srY, { color: C.ACCENT });

  // Macrostructure
  srY = addSectionHeading(sr, "Macrostructure of a Persuasive Text", srY, { color: C.PRIMARY });

  srY = addBodyText(sr, "1. INTRODUCTION", srY, { fontSize: 12 });
  srY = addBodyText(sr, "   - Hook the reader (rhetorical question, surprising fact, or bold statement)", srY, { fontSize: 10 });
  srY = addBodyText(sr, "   - Introduce the topic clearly", srY, { fontSize: 10 });
  srY = addBodyText(sr, "   - State your position (thesis statement)", srY, { fontSize: 10 });
  srY += 4;

  srY = addBodyText(sr, "2. BODY PARAGRAPHS (2-3 paragraphs)", srY, { fontSize: 12 });
  srY = addBodyText(sr, "   - Each paragraph presents ONE argument", srY, { fontSize: 10 });
  srY = addBodyText(sr, "   - Back up each argument with EVIDENCE (facts, statistics, examples)", srY, { fontSize: 10 });
  srY = addBodyText(sr, "   - EXPLAIN how the evidence supports your position", srY, { fontSize: 10 });
  srY += 4;

  srY = addBodyText(sr, "3. CONCLUSION", srY, { fontSize: 12 });
  srY = addBodyText(sr, "   - Restate your position strongly", srY, { fontSize: 10 });
  srY = addBodyText(sr, "   - Summarise your strongest argument", srY, { fontSize: 10 });
  srY = addBodyText(sr, "   - Call the reader to action or leave them thinking", srY, { fontSize: 10 });
  srY += 10;

  // Language features
  srY = addSectionHeading(sr, "Key Language Features", srY, { color: C.SECONDARY });

  const features = [
    {
      name: "Rhetorical Questions",
      def: "Questions you ask to make the reader think, not to get an answer.",
      example: "\"How could we possibly justify sending innocent animals into a war they could never understand?\"",
    },
    {
      name: "Emotive Language",
      def: "Words chosen to make the reader feel strong emotions.",
      example: "\"Innocent horses were forced to endure horrific conditions.\" (\"innocent,\" \"forced,\" \"horrific\" trigger emotion)",
    },
    {
      name: "High Modality Language",
      def: "Strong, definite words that express certainty.",
      example: "\"must,\" \"certainly,\" \"undoubtedly,\" \"essential,\" \"always,\" \"never\"",
    },
    {
      name: "Nominalisations",
      def: "Turning verbs (actions) into nouns (things). Makes writing sound more formal and authoritative.",
      example: "\"destroy\" -> \"destruction\" | \"suffer\" -> \"suffering\" | \"decide\" -> \"decision\"",
    },
    {
      name: "Tier 2 and 3 Vocabulary",
      def: "Precise, powerful words that show expertise and elevate the writing.",
      example: "\"devastating\" (not \"bad\") | \"exploitation\" (not \"use\") | \"catastrophic\" (not \"terrible\")",
    },
  ];

  for (const f of features) {
    srY = addBodyText(sr, f.name, srY, { fontSize: 12 });
    srY = addBodyText(sr, f.def, srY, { fontSize: 10 });
    srY = addBodyText(sr, "Example: " + f.example, srY, { fontSize: 10, italic: true });
    srY += 6;
  }

  addPdfFooter(sr, "War Horse | Lesson 19 | Persuasive Text Structure Reference");

  // --- Write all files ----------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/WH4_Lesson19.pptx` }),
    writePdf(po, PLANNING_PDF_PATH),
    writePdf(sr, STRUCTURE_PDF_PATH),
  ]);

  const pptxPath = `${OUT_DIR}/WH4_Lesson19.pptx`;
  console.log("PPTX written to " + pptxPath);
  console.log(`Done: ${PLANNING_RESOURCE.name}.pdf`);
  console.log(`Done: ${STRUCTURE_RESOURCE.name}.pdf`);
}

build().catch(console.error);
