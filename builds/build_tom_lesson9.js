"use strict";

// Tom Appleby, Convict Boy -- Lesson 9: Chapters 20-21 -- Setting Sail
// Year 5/6 Literacy | Novel study + Text-level writing
// Reading: Chapters 20-21 (Scarborough sets sail, Tom sees Rob)
// Writing (text-level): Plan a body paragraph for an information report using SPO
//   Topic: 18th Century England Life (informed by non-fiction article)
// Unit: Information Report -- Convict Settlement in Australia (Body Paragraph 1)

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
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 9;
const FOOTER = "Chapters 20-21 | Lesson 9 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson9_Setting_Sail";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const SPO_TEMPLATE = makeSessionResource(
  SESSION_NUMBER,
  "Body Paragraph SPO Template",
  "Student planning template: single paragraph outline for 18th Century England body paragraph."
);
const MODEL_SPO = makeSessionResource(
  SESSION_NUMBER,
  "Body Paragraph SPO Model",
  "Teacher reference: completed SPO model showing topic sentence, supporting detail and concluding sentence."
);
const RESOURCE_ITEMS = [SPO_TEMPLATE, MODEL_SPO];
const SPO_TEMPLATE_PATH = path.join(OUT_DIR, SPO_TEMPLATE.fileName);
const MODEL_SPO_PATH = path.join(OUT_DIR, MODEL_SPO.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Two short chapters today -- 20 and 21. The Scarborough finally sets sail, and Tom meets a new boy named Rob
- Then we shift gears: we begin planning our first body paragraph for the information report on Convict Settlement in Australia
- Today's body paragraph is about life in 18th Century England -- the world Tom was leaving behind

DO:
- Display title slide as students settle
- Have novels on desks, bookmarked at Chapter 20
- Have the non-fiction article about 18th Century England ready to distribute (teacher-provided)

TEACHER NOTES:
This is a pivot lesson -- reading continues but the writing focus shifts from sentence-level to text-level. The note-taking skills from Lesson 8 support the reading of the non-fiction article.

WATCH FOR:
- Students who want to dwell on Jem's death from earlier chapters -- keep the focus forward
- Students who remember the SPO from earlier lessons -- they can support peers today

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_LI_SC = `SAY:
- Two strands today: reading Chapters 20-21 briefly, and planning the first body paragraph of our information report
- Read the success criteria from the slide
- SC1 is the floor -- everyone can write at least one key fact about 18th Century England. SC2 is the core -- planning a clean SPO. SC3 is depth -- choosing the strongest details

DO:
- Choral read LI and SCs
- Emphasise SC2 as the main exit target -- every student should have a completed SPO by the end of the lesson

TEACHER NOTES:
SC2 is the write-up target. SPO must have a topic sentence, supporting detail (at least 3 points), and a concluding sentence. SC3 is about detail quality -- are the supporting points the strongest ones available?

WATCH FOR:
- Students expecting a full paragraph today -- clarify: we plan the SPO today, we write the paragraph next lesson or for homework
- Students unclear on what SPO means -- see the I Do

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_READING = `SAY:
- Chapters 20 and 21. Short chapters. I will read aloud and we will pause once
- Focus: what do we learn about Tom's new situation and the other characters around him?
- Find Chapter 20 now

DO:
- Give 30 seconds to find the chapter
- Read aloud at a steady pace -- these chapters move quickly
- One pause point: p. 102

TEACHER NOTES:
These chapters establish Rob as a character and the new community of the ship. We keep the reading brief today because the writing task takes the second half of the lesson.

WATCH FOR:
- Students interested in Rob as a contrast to Tom -- flag this for future lessons
- Students getting hung up on the ship terminology -- skim past these

[Literacy: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE = `SAY:
- Pause here. "The horde of convicts limped obediently down the hatch once more."
- Ask: what is the big idea in this sentence? [The convicts, including Tom, are tired, beaten down, and no longer resisting. They follow orders like a group of exhausted animals]
- Ask: "horde" and "limped obediently" -- what do these word choices tell us? [Horde = a large mass, loses individual identity. Limped obediently = they are hurt, they have stopped fighting]
- Ask: how does this compare with Tom at the start of the novel? [Earlier Tom was angry, resistant, full of plans. Now he just follows the group]

DO:
- Display the quote
- Ten seconds of think time
- Cold Call 3 students
- Connect to earlier Tom -- emphasise the shift

CFU CHECKPOINT:
Technique: Cold Call
Script:
- "What has changed about Tom since the start of the novel?"
- "Why has the author called the prisoners a 'horde'?"
- Scan for: students identifying the shift in Tom's spirit AND the word-choice effect
PROCEED:
- Most students identify the shift and the craft choice. Move to the writing focus.
PIVOT:
- Most likely issue -- students see no change
- Reteach: "At the start, Tom was arguing, hoarding, planning. Now he 'limped obediently'. One word -- obediently -- what has happened to Tom's spirit?"
- Re-check: "What word from earlier in the novel is the OPPOSITE of obediently?" [resolved, determined, defiant]

TEACHER NOTES:
This pause point links character change to word choice. It is the final pause point before we shift to text-level writing.

WATCH FOR:
- Students who call the shift 'sad' but cannot say why -- push for the specific word evidence
- Students who connect to the recurring "down" motif from Lesson 8 -- excellent continuity

[Literacy: Pause Point | VTLM 2.0: Higher-Order Questioning]`;

const NOTES_ARTICLE = `SAY:
- Now we shift. For the rest of the lesson, we are preparing to write a body paragraph about 18th Century England
- You have a short non-fiction article. Read it with a partner. Use your note-taking skills from last lesson -- find keywords about life in 18th Century England
- Pay special attention to: rich and poor, the Industrial Revolution, children's lives, schools
- You have 8 minutes

DO:
- Distribute the non-fiction article
- Circulate -- check that students are noting key facts, not copying sentences
- Remind of KPAS shortcuts from Lesson 8 -- keywords only

TEACHER NOTES:
The non-fiction article is teacher-provided. The one suggested in planning is: https://blogs.ancestry.com/cm/what-was-it-like-to-live-in-18th-century-england/ -- or a classroom-appropriate equivalent. Alternative articles are acceptable if they cover similar ground: rich/poor divide, Industrial Revolution, children's lives.

WATCH FOR:
- Students overwhelmed by the text -- pair them with a stronger reader
- Students who find the reading easy -- direct them to record MORE notes, with specific examples

[Literacy: Reading for Writing | VTLM 2.0: Research and Note-Taking]`;

const NOTES_IDO_STRUCTURE = `SAY:
- Before we plan our body paragraph, we revise the structure. An information report has four parts
- A TITLE tells readers what the report is about
- An INTRODUCTION introduces the topic -- we wrote this in an earlier lesson
- BODY PARAGRAPHS describe specific aspects of the topic -- today we plan the first one
- A CONCLUSION summarises
- Each body paragraph has its own mini-structure. That is what we focus on today
- A body paragraph has a TOPIC SENTENCE, SUPPORTING DETAIL, and a CONCLUDING SENTENCE. That is the SPO -- Single Paragraph Outline

DO:
- Display the four-part structure
- Emphasise we are planning ONE body paragraph today
- Link to the mentor text the students saw earlier in the unit

TEACHER NOTES:
Keep this brief -- 2-3 minutes. Students saw the structure in an earlier lesson. This is revision, not first teaching.

WATCH FOR:
- Students confusing body paragraph structure with whole-report structure -- clarify
- Students familiar with SPO from earlier work -- they can lead the next slide

[Literacy: I Do -- Revising Structure | VTLM 2.0: Explicit Teaching]`;

const NOTES_IDO_SPO = `SAY:
- Inside a body paragraph, we use the SPO -- Single Paragraph Outline
- TOPIC SENTENCE (TS): introduces the main idea. Often tells the reader who/what, doing what, where, when, how or why
- SUPPORTING DETAIL: expands the main idea with facts, examples or descriptions
- CONCLUDING SENTENCE (CS): reminds the reader of the main idea without repeating the topic sentence
- Watch me plan a body paragraph on 18th Century England
- TS: "Life in 18th Century England was very different depending on whether you were rich or poor"
- Supporting Detail 1: Rich people lived luxuriously in mansions with servants
- Supporting Detail 2: Poor people often worked long hours in factories during the Industrial Revolution
- Supporting Detail 3: Many poor children, including orphans, worked from a young age
- CS: "These deep differences between rich and poor shaped every part of daily life"

DO:
- Display the modelled SPO
- Point to each part as you read it
- Think aloud: "My topic sentence tells the reader this paragraph is about rich vs poor. My details give facts. My concluding sentence closes the idea"

TEACHER NOTES:
This is the mentor model. Students will plan their own SPO next. Emphasise the topic sentence sets the direction -- every supporting detail must serve it.

WATCH FOR:
- Students whose TS is too broad ("England was different") -- push for a specific angle
- Students whose CS just repeats the TS -- it should summarise, not parrot

[Literacy: I Do -- Modelled SPO | VTLM 2.0: Worked Example]`;

const NOTES_CFU = `SAY:
- Quick check. I will show you three sentences. One is the topic sentence, one is a supporting detail, and one is the concluding sentence
- On your whiteboard, write T, SD or CS next to each letter
- A, B, C -- label them
- You have 30 seconds

DO:
- Display the three sentences labelled A, B, C
- Use Show Me Boards -- count down, then scan

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Label each sentence: T, SD or CS. Write A/B/C and the label. 30 seconds"
- Count down, scan
- Look for: A = SD, B = T, C = CS
PROCEED:
- Most boards correctly label. Transition to You Do.
PIVOT:
- Most likely issue -- students mix up TS and CS because both summarise
- Reteach: "The TOPIC sentence INTRODUCES. The CONCLUDING sentence CLOSES. The TS says 'here is what this paragraph is about'. The CS says 'here is what I just showed you'"
- Re-check: "Which one would you expect to start the paragraph?"

TEACHER NOTES:
Show Me Boards reveal the whole class at once. Students should recognise the TS by the introducing move, the CS by the closing move, and the SD by the specific detail.

WATCH FOR:
- Students confusing TS and CS -- this is the common confusion
- Students who correctly identify all three -- they are ready to plan their own

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- A was the supporting detail -- "Many children worked in factories from age six." It is a specific fact
- B was the topic sentence -- "Life for poor children in 18th Century England was harsh and dangerous." It introduces the paragraph
- C was the concluding sentence -- "These harsh conditions shaped the childhoods of thousands of young people." It reminds the reader of the main idea without repeating

DO:
- Reveal each answer
- Highlight the introducing vs closing move
- Transition to You Do

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Now you plan your own SPO for a body paragraph about 18th Century England
- On your template, you already have a topic sentence: "Life in 18th Century England was very different depending on whether you were rich or poor"
- You need to add at least THREE supporting details using your notes from the article
- Then write a concluding sentence that closes the paragraph without repeating the topic sentence
- If you are ready for an extension, cross out the provided topic sentence and write your own about a different angle on 18th Century England
- You have 15 minutes

DO:
- Distribute the SPO Template
- Circulate -- check that each supporting detail actually supports the topic sentence
- Prompt: "Does this detail serve your topic sentence? If not, swap it"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: use the provided topic sentence AND the provided concluding sentence. Only fill in three supporting details from your notes
- Extra Notes: refer to your KPAS notes from the article. Each supporting detail should be one short sentence
EXTENDING PROMPT:
- Task: write your own topic sentence on a different angle (e.g., children's work, the Industrial Revolution, schools, housing). Then plan three supporting details and a concluding sentence. If you finish, start drafting the full paragraph in sentences

TEACHER NOTES:
The template does the structural work. Students focus on content choice. The exit target is a complete SPO with at least 3 supporting details.

WATCH FOR:
- Students whose supporting details do not match their topic sentence -- reteach: every detail must serve the topic
- Students whose CS just repeats the TS -- push for a closing move, not a repeat
- Students finishing quickly and with quality -- direct to the extension

[Literacy: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: I have written at least one fact about 18th Century England -- thumbs?
- SC2: I have a complete SPO with a topic sentence, at least 3 supporting details, and a concluding sentence -- thumbs?
- SC3: My supporting details are the strongest ones from my notes -- thumbs?
- Turn and Tell: share your topic sentence with your partner. Check -- does it set a clear direction for your paragraph?

DO:
- Thumbs check each SC
- Collect completed SPOs at the door -- these become the plan for writing the full paragraph next lesson or at home
- Students keep their template for reference

TEACHER NOTES:
Completed SPOs are the evidence of learning for SC2. Students who have not finished may need a small-group follow-up next lesson. Students will convert SPO to full paragraph next lesson.

WATCH FOR:
- Students thumbs-down on SC2 -- investigate: missing TS, missing SD count, or missing CS? Different follow-ups
- Students thumbs-up on all three with quality work -- they can start drafting the paragraph at home

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two printable resources today
- The ${SPO_TEMPLATE.name} is for you to plan your body paragraph
- The ${MODEL_SPO.name} is for teacher reference -- shows the mentor model

DO:
- Print the SPO template before the lesson (one per student)
- Print the model (teacher copy only)

TEACHER NOTES:
The SPO template has a provided topic sentence for supported students. Extension students can replace it. The model SPO shows the structure and quality expected.

[Literacy: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Chapters 20-21: Setting Sail -- Lesson 9";

  // Slide 1 -- Title
  titleSlide(
    pres,
    "Chapters 20-21",
    "Setting Sail + Body Paragraph Planning",
    "Lesson 9  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  // Slide 2 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to plan a body paragraph for an information report, drawing on research notes to write a topic sentence, supporting detail and a concluding sentence",
    ],
    [
      "I can write at least one key fact about 18th Century England",
      "I can plan a single paragraph outline with a topic sentence, supporting detail and concluding sentence",
      "I can choose the strongest supporting details from my research notes",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // Slide 3 -- Reading Launch
  contentSlide(
    pres,
    "Teacher Read Aloud",
    C.PRIMARY,
    "Chapters 20-21",
    [
      "Reading Mode: Teacher Read Aloud",
      "Chapter 20: The Scarborough sets sail as part of the First Fleet",
      "Chapter 21: The ship anchors at Santa Cruz -- Tom learns the new boy's name is Rob",
      "Focus: what has changed about Tom? Who is Rob?",
    ],
    NOTES_READING,
    FOOTER
  );

  // Slide 4 -- Pause Point: horde of convicts
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 20 -- p. 102",
    "The horde of convicts limped obediently down the hatch once more.",
    "p. 102",
    "What has changed about Tom since the start of the novel? What do the words 'horde' and 'obediently' tell us?",
    NOTES_PAUSE,
    FOOTER
  );

  // Slide 5 -- Non-fiction article / research
  contentSlide(
    pres,
    "Research",
    C.SECONDARY,
    "18th Century England -- Non-Fiction Article",
    [
      "Read the article with a partner (teacher-provided)",
      "Use your KPAS notes skill from Lesson 8 -- keywords only",
      "Focus on: rich and poor, the Industrial Revolution, children's lives",
      "You have 8 minutes",
    ],
    NOTES_ARTICLE,
    FOOTER
  );

  // Slide 6 -- I Do: Information Report Structure (brief)
  modellingSlide(
    pres,
    "I Do -- Recap",
    "Information Report Structure",
    "Structure of the whole report:\n\n- TITLE\n- INTRODUCTION (done in earlier lesson)\n- BODY PARAGRAPHS (today!)\n- CONCLUSION (later lesson)\n\nToday we plan the first body paragraph:\n\n18th Century England Life",
    "Structure of one body paragraph:\n\n- Topic Sentence (TS)\n  introduces the main idea\n\n- Supporting Detail (SD)\n  gives facts, examples, descriptions\n\n- Concluding Sentence (CS)\n  closes the paragraph\n\nThat is the SPO -- Single Paragraph Outline.",
    NOTES_IDO_STRUCTURE,
    FOOTER
  );

  // Slide 7 -- I Do: Modelled SPO
  modellingSlide(
    pres,
    "I Do -- Model SPO",
    "Planning: 18th Century England",
    "TOPIC SENTENCE:\nLife in 18th Century England was very different depending on whether you were rich or poor.\n\nSUPPORTING DETAIL 1:\nRich people lived luxuriously in mansions with servants.\n\nSUPPORTING DETAIL 2:\nPoor people often worked long hours in factories during the Industrial Revolution.\n\nSUPPORTING DETAIL 3:\nMany poor children, including orphans, worked from a young age.",
    "CONCLUDING SENTENCE:\nThese deep differences between rich and poor shaped every part of daily life.\n\nNotice:\n\n- TS sets a specific angle (rich vs poor), not a vague overview\n- Each SD serves the TS\n- CS summarises without repeating the TS word-for-word",
    NOTES_IDO_SPO,
    FOOTER
  );

  // Slides 8-9 -- CFU with reveal: label TS / SD / CS
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "TS, SD or CS?",
      "Show Me Boards",
      "Label each sentence T (Topic), SD (Supporting Detail) or CS (Concluding).\n\nA: Many children worked in factories from age six.\nB: Life for poor children in 18th Century England was harsh and dangerous.\nC: These harsh conditions shaped the childhoods of thousands of young people.",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      const ansY = 4.25;
      slide.addShape("roundRect", {
        x: 0.5, y: ansY, w: 9, h: 0.82, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      });
      slide.addText("A = SD       B = T       C = CS", {
        x: 0.5, y: ansY + 0.06, w: 9, h: 0.32,
        fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addText("B introduces, A is a specific fact, C closes the idea.", {
        x: 0.5, y: ansY + 0.42, w: 9, h: 0.32,
        fontSize: 12, fontFace: FONT_B, color: C.WHITE,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // Slide 10 -- You Do
  modellingSlide(
    pres,
    "You Do",
    "Plan Your SPO",
    "Use your SPO Template and your article notes.\n\nFirst: Your topic sentence is provided. Read it carefully.\n\nNext: Write at least THREE supporting details from your notes. Each detail must support the topic sentence.\n\nThen: Write a concluding sentence that closes without repeating the topic sentence.\n\nExtension: Replace the topic sentence with your own angle, then rebuild the SPO.",
    "Checklist before you finish:\n\n- Does each SD truly support my TS?\n- Are my SDs the strongest from my notes?\n- Does my CS close the idea without repeating the TS?\n- Is my SPO ready to become a full paragraph next lesson?\n\nYou have 15 minutes.",
    NOTES_YOUDO,
    FOOTER
  );

  // Slide 11 -- Closing
  closingSlide(
    pres,
    "Share your topic sentence with your partner. Does it set a clear direction for your paragraph?",
    [
      "I can write at least one key fact about 18th Century England",
      "I can plan a complete SPO (TS, 3+ SDs, CS)",
      "I can choose the strongest supporting details from my notes",
    ],
    NOTES_CLOSING
  );

  // Slide 12 -- Resources
  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // ---------------------------------------------------------------------------
  // PDF 1 -- SPO Template
  // ---------------------------------------------------------------------------
  const ws = createPdf({ title: SPO_TEMPLATE.name });
  let wsY = addPdfHeader(ws, "Body Paragraph SPO -- 18th Century England", {
    color: C.PRIMARY,
    subtitle: "Information Report: Convict Settlement in Australia",
    lessonInfo: "Lesson 9 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(
    ws,
    "An SPO is a Single Paragraph Outline. It plans one body paragraph. Three parts: a Topic Sentence (TS), Supporting Details (SD), a Concluding Sentence (CS). Each detail must support the topic sentence.",
    wsY,
    { color: C.PRIMARY }
  );

  wsY = addSectionHeading(ws, "Topic Sentence", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Use this topic sentence OR write your own (extension only):", wsY);
  wsY += 4;
  wsY = addTipBox(
    ws,
    "Life in 18th Century England was very different depending on whether you were rich or poor.",
    wsY,
    { color: C.SECONDARY }
  );
  wsY = addBodyText(ws, "My topic sentence:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Supporting Details (at least 3)", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "Use your notes from the non-fiction article. Each detail should be one short sentence.", wsY);
  wsY += 4;

  wsY = addBodyText(ws, "SD 1:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "SD 2:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "SD 3:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "SD 4 (optional):", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Concluding Sentence", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Close the paragraph. Remind the reader of the main idea WITHOUT repeating the topic sentence word-for-word.", wsY);
  wsY += 4;
  wsY = addBodyText(ws, "My concluding sentence:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 10;

  wsY = addTipBox(
    ws,
    "Check before you finish:  Does each SD truly support my TS?  |  Are my SDs the strongest from my notes?  |  Does my CS close without repeating?",
    wsY,
    { color: C.SECONDARY }
  );

  addPdfFooter(ws, "Lesson 9 | 18th Century England | SPO Planning Template");
  await writePdf(ws, SPO_TEMPLATE_PATH);

  // ---------------------------------------------------------------------------
  // PDF 2 -- Model SPO
  // ---------------------------------------------------------------------------
  const ak = createPdf({ title: MODEL_SPO.name });
  let akY = addPdfHeader(ak, "Body Paragraph SPO -- Model", {
    color: C.PRIMARY,
    subtitle: "Teacher Reference",
    lessonInfo: "Lesson 9 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(
    ak,
    "Marking note: accept any SPO where every supporting detail serves the topic sentence, and the concluding sentence closes without repeating. Expect some student detail to focus on rich/poor, some on work, some on children. All valid.",
    akY,
    { color: C.ALERT }
  );

  akY = addSectionHeading(ak, "Model SPO -- 18th Century England", akY, { color: C.PRIMARY });
  akY += 2;
  akY = addBodyText(ak, "Topic Sentence:", akY, { fontSize: 11, color: C.PRIMARY });
  akY = addBodyText(ak, "Life in 18th Century England was very different depending on whether you were rich or poor.", akY, { fontSize: 11 });
  akY += 8;

  akY = addBodyText(ak, "Supporting Detail 1:", akY, { fontSize: 11, color: C.ACCENT });
  akY = addBodyText(ak, "Rich people lived luxuriously in mansions with many servants.", akY, { fontSize: 11 });
  akY += 6;
  akY = addBodyText(ak, "Supporting Detail 2:", akY, { fontSize: 11, color: C.ACCENT });
  akY = addBodyText(ak, "Poor people often worked long hours in factories during the Industrial Revolution.", akY, { fontSize: 11 });
  akY += 6;
  akY = addBodyText(ak, "Supporting Detail 3:", akY, { fontSize: 11, color: C.ACCENT });
  akY = addBodyText(ak, "Many poor children, including orphans, worked from a young age and did not go to school.", akY, { fontSize: 11 });
  akY += 8;

  akY = addBodyText(ak, "Concluding Sentence:", akY, { fontSize: 11, color: C.PRIMARY });
  akY = addBodyText(ak, "These deep differences between rich and poor shaped every part of daily life in 18th Century England.", akY, { fontSize: 11 });
  akY += 10;

  akY = addSectionHeading(ak, "Marking Notes", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "- TS must have a clear angle (not 'England was different' -- but 'rich vs poor', 'work', 'children', etc.)", akY);
  akY = addBodyText(ak, "- At least 3 SDs; each must directly support the TS. Off-topic details (e.g., royalty when TS is about poor children) lose the mark", akY);
  akY = addBodyText(ak, "- CS closes the idea. It should not word-for-word repeat the TS", akY);
  akY = addBodyText(ak, "- Extension students may write their own TS on a different angle -- accept any valid angle (work, school, housing, food, children)", akY);
  akY += 10;

  akY = addSectionHeading(ak, "Common Student Errors", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- TS too broad ('Life was different') -- push for a specific angle", akY);
  akY = addBodyText(ak, "- CS repeats TS -- reteach closing move vs introducing move", akY);
  akY = addBodyText(ak, "- SD drifts off topic -- every SD must match the TS angle", akY);
  akY = addBodyText(ak, "- SDs are opinions rather than facts from the article -- redirect to notes", akY);

  addPdfFooter(ak, "Lesson 9 | SPO Model -- TEACHER COPY");
  await writePdf(ak, MODEL_SPO_PATH);

  // ---------------------------------------------------------------------------
  // Write PPTX
  // ---------------------------------------------------------------------------
  const outName = path.join(OUT_DIR, "Tom_Lesson9_Chapters_20-21.pptx");
  fs.mkdirSync(OUT_DIR, { recursive: true });
  await pres.writeFile({ fileName: outName });
  console.log("PPTX written to " + outName);
  console.log("PDF written to " + SPO_TEMPLATE_PATH);
  console.log("PDF written to " + MODEL_SPO_PATH);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
