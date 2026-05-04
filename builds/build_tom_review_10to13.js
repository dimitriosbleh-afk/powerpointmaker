"use strict";

// Tom Unit -- Review of Lessons 10-13
// Year 5/6 Literacy. Sits at the end of the Information Report + Sentence-Level
// Skills sequence (Lessons 10-13). Consolidates four prior skills:
//   L10: Body paragraph for an information report (TS + supporting details + CS)
//   L11: Conjunctions because / but / so to fit meaning in comprehension answers
//   L12: Single Paragraph Outline (SPO) to summarise text
//   L13: Note taking with KPAS (Keywords, Phrases, Abbreviations, Symbols)
// No new chapters. No new content. Mixed-skills review with one consolidated practice sheet.

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

const SESSION_NUMBER = 14;
const FOOTER = "Lessons 10-13 Review | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Review_Lessons10to13";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PRACTICE_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mixed Skills Review",
  "Mixed-skills practice covering body paragraph structure, because/but/so, SPO and KPAS.",
  {
    name: "Lessons 10 to 13 Review Practice",
    fileName: path.posix.join(getSessionResourceFolder(SESSION_NUMBER), "Lessons 10 to 13 Review Practice.pdf"),
  }
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mixed Skills Review Answer Key",
  "Teacher reference: model answers across all four review sections.",
  {
    name: "Lessons 10 to 13 Review Answer Key",
    fileName: path.posix.join(getSessionResourceFolder(SESSION_NUMBER), "Lessons 10 to 13 Review Answer Key.pdf"),
  }
);
const RESOURCE_ITEMS = [PRACTICE_RESOURCE, ANSWER_KEY_RESOURCE];
const PRACTICE_PDF_PATH = path.join(OUT_DIR, PRACTICE_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today is a review of the past four lessons. No new chapters
- We will revisit four skills: body paragraph structure, because/but/so, SPO, and KPAS notes
- Bring your previous work -- your SPOs, your KPAS notes and your body paragraph

DO:
- Display title slide as students settle
- Have novels and exercise books on desks
- Have students' own SPOs and KPAS notes from the past week within reach

TEACHER NOTES:
This is a consolidation lesson. The aim is to help students see how the four skills connect: SPO -> body paragraph; because/but/so for comprehension answers; KPAS for fast note taking. One mixed-skills worksheet supports independent practice.

WATCH FOR:
- Students who cannot find earlier work -- pair them with a peer for the recall
- Students who are confident -- they can take the lead in partner work

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LAUNCH = `SAY:
- Quick recall. On your whiteboard, write the four skills we have built across the past four lessons
- 60 seconds. Best you can. Use the names from the lessons if you remember them
- Some of you may remember all four. If you remember one or two, that is fine

DO:
- 60 seconds silent recall on whiteboards
- Walk and scan: who has all four? who has two or three? who is stuck?
- Reveal the four skills on the next slide

TEACHER NOTES:
Active recall before the LI/SC slide. The four skills are body paragraph (Lesson 10), because/but/so (Lesson 11), SPO (Lesson 12), KPAS (Lesson 13). Use the launch as a quick formative scan.

WATCH FOR:
- Students with all four -- excellent retention
- Students with one or two -- this lesson will rebuild the others
- Students with none -- pair with a confident peer for the rest of the lesson

[General: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Today is review. We are not learning anything brand new -- we are tightening up four skills we have already met
- Read each success criterion together

DO:
- Choral read the LI, then each SC
- Brief reminder: review lesson, mixed practice at the end

TEACHER NOTES:
SC1 targets information report structure (Lesson 10). SC2 targets sentence-level skills (Lesson 11 conjunctions plus Lesson 12 SPO). SC3 targets KPAS note taking (Lesson 13). The criteria appear as a single unlabelled list.

WATCH FOR:
- Students who already feel confident on all three -- they can challenge themselves on the optional extension
- Students unsure about one strand -- they will see it modelled and supported

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_RECALL_BODYPARA = `SAY:
- Lesson 10 review. Body paragraph for an information report
- Three parts: topic sentence, supporting details, concluding sentence
- The TS introduces the main idea. The supporting details expand on it. The CS wraps it up without repeating the TS

DO:
- Display the structure card
- Quick partner check: "Tell your partner ONE thing each part of a body paragraph does"

TEACHER NOTES:
Keep the recall tight. Do NOT reteach -- this is a refresher. If students are shaky, the You Do mixed practice will catch any gaps.

WATCH FOR:
- Students who confuse TS and CS -- prompt: "TS comes first; CS comes last; CS does not repeat the TS word for word"

[General: Recall | VTLM 2.0: Retention and Recall]`;

const NOTES_RECALL_CONJUNCTIONS = `SAY:
- Lesson 11 review. Three conjunctions for comprehension answers
- BECAUSE = a reason. Tells us WHY
- BUT = a contrast. Tells us a CHANGE in direction
- SO = a consequence. Tells us WHAT HAPPENS as a result

DO:
- Display the conjunctions card
- Quick oral check: "Give me a 'because' sentence about Tom in one breath"

TEACHER NOTES:
Refresh, not reteach. Remind students about the comma rule (comma before BUT and SO when joining two ideas) but keep it brief.

WATCH FOR:
- Students who try to use 'because' as a contrast -- redirect: "Because is about WHY"
- Students who use SO accurately -- celebrate

[General: Recall | VTLM 2.0: Retention and Recall]`;

const NOTES_RECALL_SPO = `SAY:
- Lesson 12 review. Single Paragraph Outline -- a fast plan for one paragraph
- Three parts of an SPO: topic sentence, supporting details, concluding sentence
- The SPO is the PLAN. The body paragraph is the WRITTEN VERSION of that plan

DO:
- Display the SPO card
- Hold up your SPO from Lesson 12 -- this is the same shape we will use today

TEACHER NOTES:
Make the explicit link: SPO (plan) -> body paragraph (write). Lessons 10 and 12 share the same structure -- this is intentional.

WATCH FOR:
- Students who treat SPO and body paragraph as different things -- the structure is the same; the SPO is the plan version

[General: Recall | VTLM 2.0: Retention and Recall]`;

const NOTES_RECALL_KPAS = `SAY:
- Lesson 13 review. KPAS = Keywords, Phrases, Abbreviations, Symbols
- We use KPAS to take fast notes -- to capture meaning without writing every word
- Symbols we know: b/c, w/, w/o, =, +, ->, up arrow, down arrow, /

DO:
- Display the KPAS reference
- Quick choral read of the symbols

TEACHER NOTES:
Refresh the symbol bank. This is the most recent skill, so most students will recall it. The mixed practice will check that the symbols are being used in context.

WATCH FOR:
- Students who confuse the right arrow (cause) with the up arrow (trend) -- the difference is on the new symbols card from Lesson 13

[General: Recall | VTLM 2.0: Retention and Recall]`;

const NOTES_IDO = `SAY:
- Watch me. I am going to take ONE moment from the past week of reading and use ALL FOUR skills on it
- The moment: Sam grows weaker during the storms
- First the SPO -- the plan
- Then KPAS notes from one supporting detail
- Then a comprehension answer using BECAUSE
- Last, I'll hint at the body paragraph that the SPO would become

DO:
- Display the I Do model
- Think aloud through each move
- Highlight: same content, four different ways to capture or use it

TEACHER NOTES:
The I Do shows how the four skills connect. The same chapter content can be planned (SPO), noted (KPAS), discussed in a comprehension answer (because/but/so) and written up (body paragraph). The integration is the point.

WATCH FOR:
- Students who say "we already did this" -- yes, that is the goal of review
- Students who notice the connection across skills -- excellent

[General: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Together. Mini-whiteboards out
- New moment: when the ship reached Port Jackson
- Round 1: write a topic sentence (TS) with WHO, WHAT DOING, WHEN
- Round 2: complete this comprehension answer with BECAUSE: "The captain decided to sail away from the first land BECAUSE..."
- Round 3: convert this sentence to KPAS: "The ship anchored at Port Jackson on a calm afternoon."

DO:
- Run three short rounds (60 seconds each)
- Hold up boards after each round
- Pick 1-2 strong examples to show the class

CFU CHECKPOINT:
Technique: Show Me Boards (three rounds)

Script:
- Round 1: "Topic sentence about Port Jackson arrival. Who, what doing, when. Boards up in 60 seconds"
- Scan for: a sentence that names WHO (the ship / the convicts / the captain), WHAT DOING (anchored / arrived / disembarked) and WHEN
- Round 2: "Complete: 'The captain decided to sail away from the first land BECAUSE...' Boards up"
- Scan for: a textual reason (the land was too flat for crops, or similar)
- Round 3: "Convert: 'The ship anchored at Port Jackson on a calm afternoon.' Boards up"
- Scan for: keywords kept (ship anchored, Port Jackson, calm afternoon); small words dropped (the, on, a)

PROCEED (>=80% on each round): Continue to the next round, then release to You Do.
PIVOT (<80% on any round):
- TS pivot: most likely issue -- TS is too vague. Reteach by example: "The ship arrived" is not enough. Add WHEN, WHERE, WHO. Re-check on a new prompt
- BECAUSE pivot: most likely issue -- the answer is not from the text. Reteach: "Because needs a REASON, and your reason has to come from the chapter, not your imagination"
- KPAS pivot: most likely issue -- students keep small words. Reteach by example: cross out 'the', 'on', 'a'; circle the keywords; rewrite

TEACHER NOTES:
Three short rounds rebuild the three sentence-level skills in one go. The body paragraph skill comes back in the You Do, where students draft from an SPO.

WATCH FOR:
- Students whose TS still says "There was a hard time" -- redirect to specifics
- Students whose KPAS keeps small words -- redirect to keywords only

[General: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU = `SAY:
- Quick check before you write. I will show you a sentence. Hold up fingers
- 1 = Topic Sentence (TS)
- 2 = Supporting Detail
- 3 = Concluding Sentence (CS)

DO:
- Read the sentence on the screen
- Use Finger Voting (1 / 2 / 3)
- Cold call after the vote: "Why did you choose that?"

CFU CHECKPOINT:
Technique: Finger Voting (1 / 2 / 3)

Script:
- "The journey of the First Fleet was long, dangerous and changed many lives forever."
- "Is this a TS, a supporting detail, or a CS? Fingers up"
- Scan for: split between 1 and 3

PROCEED (>=80% on either 1 or 3 with strong reasoning): Reveal and discuss why both are defensible.
PIVOT (<80%): Most likely issue -- students cannot tell TS from CS because both can sound general. Reteach: "TS opens the paragraph and tells you what is coming. CS closes the paragraph and reminds you of the main idea -- but it does not repeat the TS word for word." Re-check: "If you saw this sentence at the START of a paragraph, would it work? At the END?"

TEACHER NOTES:
This CFU targets the TS vs CS distinction, which is the most common confusion in Lesson 10. The reveal acknowledges that the sentence could function as either, depending on position -- the point is for students to articulate WHY.

WATCH FOR:
- Students who answer 2 -- prompt them: "Is this a SPECIFIC fact, or a BIG IDEA?"
- Students who answer 1 with reasoning -- ready to write
- Students who answer 3 with reasoning -- also ready to write

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Your turn. The review worksheet has four short sections, one for each skill from the past week
- 15 minutes. Work through all four sections
- I will circulate

DO:
- Distribute the Lessons 10-13 Review Practice
- Circulate and confer
- Have the answer key out for quick checks (teacher copy only)

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the worksheet, but with the section headings highlighted and one example completed in each section as a model
- Extra Notes: These students can refer to the I Do model on screen for parallel structure

EXTENDING PROMPT:
- Task: After completing the worksheet, write a short paragraph (5-6 sentences) that combines TWO skills: a body paragraph (TS + 3 supporting details + CS) about the journey, written from your SPO, with at least one comprehension-style sentence using BECAUSE, BUT or SO
- Extra Notes: Then convert one of the sentences to KPAS notes

TEACHER NOTES:
The four sections target the four skills directly. Students should be able to complete them in 15 minutes. If you see one section consistently weak across the class, that is the target for next-day reteaching.

WATCH FOR:
- Students who finish quickly -- direct them to the extension
- Students who get stuck on KPAS section -- the symbol bank is at the top of the worksheet
- Students who get stuck on the body paragraph section -- the SPO frame is on the worksheet too

[General: You Do | VTLM 2.0: Independent Application]`;

const NOTES_CLOSING = `SAY:
- Quick self-check against the success criteria. Thumbs for each
- Then turn to a partner: tell them WHICH of the four skills you feel strongest on, and WHICH you want more practice on
- Tomorrow we move forward into new chapters

DO:
- Run thumbs check for each SC
- Listen in on partner shares -- this tells you where to spend reteach time
- Wrap up: "These four skills come back across the unit. We will keep practising them"

TEACHER NOTES:
Use the partner share as a quick survey. If most students name the same skill as 'want more practice', plan a 10-minute warm-up on it tomorrow.

WATCH FOR:
- Students who name a different skill from their actual performance on the worksheet -- check in privately
- Students who feel strong across all four -- celebrate

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The Lessons 10-13 Review Practice is one mixed-skills worksheet covering all four skills
- The Lessons 10-13 Review Answer Key is teacher reference

DO:
- Print the practice worksheet (one per student)
- Print the answer key (teacher copy only)
- Have students' own previous work (SPOs, KPAS notes, body paragraph from Lesson 10) within reach for the launch

TEACHER NOTES:
The single worksheet keeps cognitive load low. Differentiation is built in -- enabling students get a partially-completed version as the model; extending students complete a connected paragraph after the worksheet.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Lessons 10-13 Review -- Information Report and Sentence Skills";

  // Slide 1 -- Title
  titleSlide(
    pres,
    "Lessons 10-13 Review",
    "Information Reports & Sentence Skills",
    "Year 5/6 Literacy",
    NOTES_TITLE
  );

  // Slide 2 -- Resources (immediately after title per megaprompt rule 19)
  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // Slide 3 -- Launch (active recall)
  contentSlide(
    pres,
    "Launch",
    C.PRIMARY,
    "What Are the Four Skills?",
    [
      "Mini-whiteboards out",
      "60 seconds: write the four skills we have built across the past four lessons",
      "Best you can. Names if you remember; descriptions are fine too",
      "We will reveal the four together on the next slides",
    ],
    NOTES_LAUNCH,
    FOOTER
  );

  // Slide 4 -- LI/SC
  liSlide(
    pres,
    [
      "We are reviewing the four skills from the past week: body paragraph structure, choosing because/but/so, summarising with an SPO, and taking quick notes with KPAS",
    ],
    [
      "I can name the parts of a body paragraph and explain what each part does",
      "I can use the right conjunction (because, but, so) and write a clear topic sentence for an SPO",
      "I can convert a sentence into KPAS notes that another reader can understand",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // Slide 5 -- Recall: Body Paragraph (Lesson 10)
  contentSlide(
    pres,
    "Recall -- Lesson 10",
    C.PRIMARY,
    "Body Paragraph for an Information Report",
    [
      "Topic Sentence (TS) -- introduces the main idea: who, what doing, when, where, why, how",
      "Supporting Details -- sentences that expand on the main idea with facts from research",
      "Concluding Sentence (CS) -- wraps the paragraph up without repeating the TS",
      "Language features: present tense, third person, nouns and noun groups, adverbials",
    ],
    NOTES_RECALL_BODYPARA,
    FOOTER
  );

  // Slide 6 -- Recall: Conjunctions (Lesson 11)
  contentSlide(
    pres,
    "Recall -- Lesson 11",
    C.SECONDARY,
    "because  |  but  |  so",
    [
      "BECAUSE = reason -- tells us WHY",
      "BUT = contrast -- tells us a CHANGE in direction",
      "SO = consequence -- tells us WHAT HAPPENS as a result",
      "Use a comma before BUT and SO when joining two ideas",
    ],
    NOTES_RECALL_CONJUNCTIONS,
    FOOTER
  );

  // Slide 7 -- Recall: SPO (Lesson 12)
  contentSlide(
    pres,
    "Recall -- Lesson 12",
    C.ACCENT,
    "Single Paragraph Outline (SPO)",
    [
      "An SPO is the PLAN for one paragraph",
      "Same shape as the body paragraph: TS + supporting details + CS",
      "Use the SPO before you write -- it saves time and keeps your paragraph on topic",
      "Update your SPO if you find new information while you write",
    ],
    NOTES_RECALL_SPO,
    FOOTER
  );

  // Slide 8 -- Recall: KPAS (Lesson 13)
  contentSlide(
    pres,
    "Recall -- Lesson 13",
    C.PRIMARY,
    "KPAS -- Note Taking Shorthand",
    [
      "K = Keywords  |  P = Phrases  |  A = Abbreviations  |  S = Symbols",
      "Symbols: b/c, w/, w/o, =, +, -> (cause), up arrow (more), down arrow (less), / (break)",
      "Drop little words: the, and, a, on, of, was",
      "Test: read your notes back -- can you recover the meaning?",
    ],
    NOTES_RECALL_KPAS,
    FOOTER
  );

  // Slide 9 -- I Do: model the four skills on one moment
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "One Moment, Four Skills",
    "Moment from the text:\n\nSam grew weaker during\nthe storms because food\nand water were rationed\nand the conditions below\ndeck were unbearable.\n\nWatch me use FOUR\nskills on this moment.",
    "1) SPO (plan):\nTS: \"Sam's health declined\nduring the long sea voyage.\"\nDetails: rationed food / locked\nbelow deck / illness spread\n\n2) KPAS notes:\n\"Sam health down arrow b/c\nfood + water rationed / locked\nbelow / unbearable conditions\"\n\n3) BECAUSE answer:\n\"Sam grew weaker BECAUSE\nthe convicts were rationed and\nthe air below was unbearable.\"",
    NOTES_IDO,
    FOOTER
  );

  // Slide 10 -- We Do
  contentSlide(
    pres,
    "We Do",
    C.SECONDARY,
    "Together: Three Quick Rounds on Port Jackson",
    [
      "Mini-whiteboards out -- three short rounds",
      "Round 1: Topic sentence about the ship arriving at Port Jackson (WHO, WHAT DOING, WHEN)",
      "Round 2: Complete -- \"The captain decided to sail away from the first land BECAUSE...\"",
      "Round 3: Convert to KPAS -- \"The ship anchored at Port Jackson on a calm afternoon.\"",
    ],
    NOTES_WEDO,
    FOOTER
  );

  // Slide 11 + 11b -- CFU with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "TS, Supporting Detail, or CS?",
      "1 = Topic Sentence  |  2 = Supporting Detail  |  3 = Concluding Sentence",
      "\"The journey of the First Fleet was long, dangerous and changed many lives forever.\"\n\nFingers: 1, 2 or 3?",
      NOTES_CFU,
      FOOTER
    ),
    (s) => {
      addCard(s, 0.5, SAFE_BOTTOM - 0.95, 9, 0.85, { strip: C.SUCCESS, fill: C.WHITE });
      s.addText("1 OR 3 (with reasoning) -- it is a BIG IDEA. Position decides whether it opens or closes the paragraph", {
        x: 0.75, y: SAFE_BOTTOM - 0.88, w: 8.4, h: 0.70,
        fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true,
        valign: "middle", margin: 0,
      });
      s.addNotes("SAY:\n- Either 1 or 3 is defensible. The key is the reasoning\n- This is a big idea about the First Fleet -- not a specific detail\n- If it OPENS the paragraph, it is a TS\n- If it CLOSES the paragraph after specific details, it is a CS\n- The trap is choosing 2 -- this sentence is too BROAD to be a supporting detail\n\nDO:\n- Reveal the answer\n- If most chose 2, run the test: 'Is this a fact about the First Fleet, or a big idea?'\n- If most chose 1 or 3, celebrate the reasoning\n\n[General: CFU Reveal | VTLM 2.0: Formative Feedback]");
    }
  );

  // Slide 12 -- You Do (instruction card + reference)
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Lessons 10-13 Mixed Practice");

    addInstructionCard(s, [
      { text: "On your worksheet:", role: "header" },
      { text: "First: Read each section heading -- it tells you which skill" },
      { text: "Next: Work through the four sections in order" },
      { text: "Then: Use the symbol reference at the top for the KPAS section" },
      { text: "Last: If you finish, try the optional extension at the bottom" },
    ], {
      x: 0.5, y: CONTENT_TOP, w: 9, h: 1.85,
      strip: C.PRIMARY, fill: C.WHITE,
      headerColor: C.PRIMARY,
    });

    const tipY = CONTENT_TOP + 2.00;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Symbol Quick Reference", {
      x: 0.75, y: tipY + 0.10, w: 5, h: 0.32,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "b/c, w/, w/o, =, +, ->, up arrow, down arrow, /", options: { breakLine: true } },
      { text: "BECAUSE = reason   |   BUT = contrast   |   SO = consequence", options: { breakLine: true } },
      { text: "Body paragraph and SPO share the same shape: TS + supporting details + CS" },
    ], {
      x: 0.75, y: tipY + 0.46, w: 8.4, h: 1.10,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bullet: true, margin: 0,
      paraSpaceAfter: 2,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // Slide 13 -- Closing
  closingSlide(
    pres,
    "Turn to a partner: tell them WHICH of the four skills you feel strongest on, and WHICH you want more practice on.",
    [
      "I can name the parts of a body paragraph and explain what each part does",
      "I can use the right conjunction (because, but, so) and write a clear topic sentence for an SPO",
      "I can convert a sentence into KPAS notes that another reader can understand",
    ],
    NOTES_CLOSING
  );

  // -----------------------------------------------------------------------
  // Companion PDFs
  // -----------------------------------------------------------------------

  // Practice
  const ws = createPdf({ title: PRACTICE_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Lessons 10-13 Review Practice", {
    color: C.PRIMARY,
    subtitle: "Mixed Skills: Body Paragraph, because/but/so, SPO, KPAS",
    lessonInfo: "Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "Four short sections, one for each skill from the past week. Work through them in order. Use the symbol reference below for the KPAS section. If you finish, try the optional extension at the bottom.", wsY, { color: C.PRIMARY });

  wsY = addSectionHeading(ws, "Symbol & Conjunction Reference", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "b/c = because    w/ = with    w/o = without    = means    + and    -> causes / results in", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "up arrow = more / increase    down arrow = less / decrease    / = full stop or new idea", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "BECAUSE = reason (why)    BUT = contrast (a change)    SO = consequence (what happens)", wsY, { fontSize: 11 });
  wsY += 8;

  // Section 1: Body paragraph parts (Lesson 10)
  wsY = addSectionHeading(ws, "Section 1 -- Body Paragraph Parts (Lesson 10)", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Read the paragraph below. Label each sentence in the margin: TS, SD or CS.", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "(1) Life in 18th Century England was very hard for poor families. (2) Many adults could not find work after the Industrial Revolution put them out of jobs. (3) Children as young as five were sent to work in factories. (4) Food was scarce and many families lived with hunger and disease. (5) For most poor families in 18th Century England, daily life was a struggle to survive.", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "Sentence 1 = ____    Sentence 2 = ____    Sentence 3 = ____    Sentence 4 = ____    Sentence 5 = ____", wsY, { fontSize: 11 });
  wsY += 10;

  // Section 2: Conjunctions (Lesson 11)
  wsY = addSectionHeading(ws, "Section 2 -- Choose the Conjunction (Lesson 11)", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Complete each sentence using BECAUSE, BUT or SO. Choose the one that fits the meaning.", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "a) Tom wanted to speak to Rob, ____ Rob had just lost his mother and went below in silence.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY = addBodyText(ws, "b) Sam was getting weaker every day, ____ Tom worried about him constantly.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY += 6;

  // Section 3: SPO (Lesson 12)
  wsY = addSectionHeading(ws, "Section 3 -- Write a Topic Sentence for an SPO (Lesson 12)", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "Main idea: \"The journey across the sea was very hard for the convicts.\"", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "Write a TS that gives the reader specifics (WHO, WHAT DOING, WHEN, WHERE):", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY = addBodyText(ws, "List ONE supporting detail from the text:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY += 6;

  // Section 4: KPAS (Lesson 13)
  wsY = addSectionHeading(ws, "Section 4 -- Convert to KPAS (Lesson 13)", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Convert this sentence into KPAS notes:", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "\"During the storms at the Cape of Good Hope, the convicts were locked below deck for thirteen days without enough food or water.\"", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "KPAS:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 8;

  // Optional extension
  wsY = addSectionHeading(ws, "Optional Extension", wsY, { color: C.ALERT });
  wsY = addBodyText(ws, "Use your TS from Section 3 to write a short body paragraph (4-5 sentences). Add at least one comprehension-style sentence using BECAUSE, BUT or SO inside the paragraph.", wsY, { fontSize: 10 });
  wsY = addLinedArea(ws, wsY, 8, { lineSpacing: 22 });

  addPdfFooter(ws, "Lessons 10-13 Review Practice");

  // Answer key
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Lessons 10-13 Review -- Model Answers", {
    color: C.ALERT,
    subtitle: "Teacher Reference -- Mixed Skills",
    lessonInfo: "Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Model answers below. Student answers in Sections 2, 3 and the extension will vary -- accept any answer that fits the prompt and is supported by the text.", akY, { color: C.ALERT });

  akY = addSectionHeading(ak, "Section 1 -- Body Paragraph Parts", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sentence 1 = TS  |  Sentences 2, 3, 4 = SD  |  Sentence 5 = CS", akY);
  akY = addBodyText(ak, "TS opens with the main idea (hard life for poor families). SDs add specific facts (jobs lost, child labour, hunger and disease). CS wraps up with 'struggle to survive' without repeating the TS.", akY, { fontSize: 10, italic: true });
  akY += 8;

  akY = addSectionHeading(ak, "Section 2 -- Choose the Conjunction", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "a) BUT (the contrast: Tom wanted to speak vs Rob went below in silence)", akY);
  akY = addBodyText(ak, "b) SO (the consequence: Sam getting weaker -> Tom's constant worry)", akY);
  akY = addBodyText(ak, "Note: 'because' would invert the direction -- Tom's worry would have to cause Sam's illness, which is wrong.", akY, { fontSize: 10, italic: true });
  akY += 8;

  akY = addSectionHeading(ak, "Section 3 -- Topic Sentence (SPO)", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "Model TS: \"The eight-month sea voyage from England to Botany Bay tested the convicts in many ways.\"", akY);
  akY = addBodyText(ak, "Or: \"Throughout the long journey to Australia, the convicts faced storms, illness and harsh conditions below deck.\"", akY);
  akY = addBodyText(ak, "Look for: WHO (convicts), WHAT DOING (faced / endured), WHEN/WHERE (the journey / on the ship). Reject vague TS like 'It was hard'.", akY, { fontSize: 10, italic: true });
  akY = addBodyText(ak, "Model supporting detail: 'The convicts were locked below deck for thirteen days during the storms at the Cape of Good Hope.'", akY);
  akY += 8;

  akY = addSectionHeading(ak, "Section 4 -- Convert to KPAS", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Original: 'During the storms at the Cape of Good Hope, the convicts were locked below deck for thirteen days without enough food or water.'", akY, { fontSize: 10, italic: true });
  akY = addBodyText(ak, "Model KPAS: \"Cape storms / convicts locked below 13 days / w/o sufficient food + water\"", akY);
  akY = addBodyText(ak, "Look for: small words dropped (the, were, for, or); keywords kept (Cape, storms, convicts, locked below, 13 days, food, water); use of w/o; use of /.", akY, { fontSize: 10, italic: true });
  akY += 12;

  akY = addSectionHeading(ak, "What to Look For (across all sections)", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- Section 1: TS opens, CS closes; SDs sit between with specific facts", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Section 2: conjunction matches the meaning (reason / contrast / consequence)", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Section 3: TS gives specifics, not generic statements", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Section 4: notes are short but the meaning is recoverable; new symbols used where they fit", akY, { fontSize: 10 });

  addPdfFooter(ak, "Review -- Answer Key -- TEACHER COPY");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Review_Lessons10to13.pptx` }),
    writePdf(ws, PRACTICE_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Review_Lessons10to13.pptx`);
  console.log("Done: " + PRACTICE_RESOURCE.name);
  console.log("Done: " + ANSWER_KEY_RESOURCE.name);
}

build().catch(console.error);
