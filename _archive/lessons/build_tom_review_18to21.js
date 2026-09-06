"use strict";

// Tom Unit -- Review of Lessons 18-21
// Year 5/6 Literacy. Sits at the end of the Body Paragraph 3 + Sentence-Level Skills sequence:
//   L18: Sentence Expansion (who / what doing / when / where / why)
//   L19: Plan body paragraph 3 with an SPO (Arrival in Sydney Cove)
//   L20: Write body paragraph 3 (language features)
//   L21: Subordinating conjunctions (dependent + independent clause = complex sentence)
// No new chapters. Mixed-skills review with one consolidated practice sheet.

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

const SESSION_NUMBER = 22;
const FOOTER = "Lessons 18-21 Review | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Review_Lessons18to21";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PRACTICE_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mixed Skills Review",
  "Mixed-skills practice covering sentence expansion, body paragraph structure, language features and subordinating conjunctions.",
  {
    name: "Lessons 18 to 21 Review Practice",
    fileName: path.posix.join(getSessionResourceFolder(SESSION_NUMBER), "Lessons 18 to 21 Review Practice.pdf"),
  }
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mixed Skills Review Answer Key",
  "Teacher reference: model answers across all four review sections.",
  {
    name: "Lessons 18 to 21 Review Answer Key",
    fileName: path.posix.join(getSessionResourceFolder(SESSION_NUMBER), "Lessons 18 to 21 Review Answer Key.pdf"),
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
- We will revisit three skills: sentence expansion, body paragraph 3 (plan + write) and subordinating conjunctions
- Bring your body paragraph 3 from Lesson 20 -- we will use it

DO:
- Display title slide as students settle
- Have novels and exercise books on desks
- Have students' written body paragraphs from Lesson 20 within reach

TEACHER NOTES:
This is a consolidation lesson. The aim is to help students see how the four lessons fit together: sentence-level skill (expansion, complex sentences) builds the sentences inside the bigger structure (body paragraph). One mixed-skills worksheet supports independent practice.

WATCH FOR:
- Students who cannot find their body paragraph -- pair them with a peer or provide a sample paragraph from the mentor (Lesson 20 PDF)
- Students who are confident -- they can take the lead in partner work

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The Lessons 18-21 Review Practice is one mixed-skills worksheet
- The Lessons 18-21 Review Answer Key is teacher reference

DO:
- Print the practice worksheet (one per student)
- Print the answer key (teacher copy only)
- Have students' own body paragraphs (from Lesson 20) within reach for the launch and the extension

TEACHER NOTES:
The single worksheet keeps cognitive load low. Differentiation is built in -- enabling students get a partially completed version; extending students rewrite a sentence from their own body paragraph using the new sentence-level moves.

[General: Resources | VTLM 2.0: Student Resources]`;

const NOTES_LAUNCH = `SAY:
- Quick recall. On your whiteboard, name what we have learned in the past four lessons
- 60 seconds. Best you can. Names if you remember; descriptions are fine too
- Some of you may remember all three skills. If you remember one or two, that is fine

DO:
- 60 seconds silent recall on whiteboards
- Walk and scan: who has all three? who has one or two? who is stuck?
- Reveal on the next slides

TEACHER NOTES:
Active recall before the LI/SC slide. The three skills are sentence expansion (L18), body paragraph plan + write (L19-20), subordinating conjunctions (L21). Use the launch as a quick formative scan.

WATCH FOR:
- Students with all three -- excellent retention
- Students with one or two -- this lesson will rebuild the others
- Students with none -- pair with a confident peer for the rest of the lesson

[General: Launch | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Today is review. We are not learning anything brand new -- we are tightening three skills we have already met
- Read each success criterion together

DO:
- Choral read the LI, then each SC
- Brief reminder: review lesson, mixed practice at the end

TEACHER NOTES:
SC1 targets sentence expansion (Lesson 18). SC2 targets body paragraph structure and writing (Lessons 19-20). SC3 targets subordinating conjunctions and complex sentences (Lesson 21).

WATCH FOR:
- Students who feel confident on all three -- they can challenge themselves on the optional extension
- Students unsure about one strand -- they will see it modelled and supported

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_RECALL_EXPAND = `SAY:
- Lesson 18 review. Sentence expansion
- We start with a short kernel sentence -- the bare bones
- We add detail: WHO, WHAT DOING, WHEN, WHERE, WHY, HOW
- If the WHEN detail comes first, we use a comma after it
- We do this to give the reader a clearer picture

DO:
- Display the expansion card
- Quick oral check: "Take 'The boys waited.' What WHEN detail could you add?"
- Possible: "On Christmas morning, the boys waited."

TEACHER NOTES:
Keep the recall tight. The key idea is that expansion adds detail for the reader -- it is not about making sentences longer for the sake of it. Adverbials (when, where, how) are placed thoughtfully.

WATCH FOR:
- Students who default to adjectives only -- redirect to adverbials (when, where, why, how)
- Students who add detail but lose clarity -- prompt: "Read it back. Is it clearer or just longer?"

[General: Recall | VTLM 2.0: Retention and Recall]`;

const NOTES_RECALL_BODYPARA = `SAY:
- Lessons 19 and 20 review. Body paragraph for an information report
- Three parts: topic sentence, supporting details, concluding sentence
- The TS introduces the main idea -- it tells the reader when, where, who and what
- The supporting details expand on the main idea with facts
- The CS adds the final idea -- it does not just repeat the TS
- We also use language features: past tense, third person, plus one appositive or relative clause

DO:
- Display the structure card
- Quick partner check: "Tell your partner ONE thing each part of a body paragraph does"

TEACHER NOTES:
Keep the recall tight. Do NOT reteach -- this is a refresher. If students are shaky, the You Do mixed practice will catch any gaps.

WATCH FOR:
- Students who confuse TS and CS -- prompt: "TS opens; CS closes; CS does not repeat the TS"

[General: Recall | VTLM 2.0: Retention and Recall]`;

const NOTES_RECALL_SUBCONJ = `SAY:
- Lesson 21 review. Subordinating conjunctions and complex sentences
- A dependent clause cannot stand alone -- it needs an independent clause to be a full sentence
- The joining word is the subordinating conjunction -- even though, since, after, because, when, while, although, if
- If the dependent clause comes first, we use a comma after it
- The two clauses together make a complex sentence

DO:
- Display the structure card
- Quick oral check: "Take 'Because the rations were low,' -- what independent clause could finish it?"
- Possible: "...the colony struggled to feed everyone."

TEACHER NOTES:
Refresh the rule -- independent must stand alone. Avoid introducing new terminology. Subordinating conjunctions, dependent and independent are enough.

WATCH FOR:
- Students who add a fragment (starts with "and" or "but") -- redirect: "Read it alone -- does it stand alone?"
- Students who write a comma splice -- redirect to the conjunction list

[General: Recall | VTLM 2.0: Retention and Recall]`;

const NOTES_IDO = `SAY:
- Watch me. I am going to take ONE plain sentence about the cove and use ALL THREE skills on it
- Plain sentence: "The boys met an Eora woman."
- First: SENTENCE EXPANSION -- I add when, where and how
- Next: I use it as a TS for a body paragraph and plan three supporting details
- Last: I turn it into a COMPLEX SENTENCE using a subordinating conjunction
- Same content, three different ways to grow it

DO:
- Display the I Do model
- Think aloud through each move
- Highlight: same plain sentence, but it grows in different shapes

TEACHER NOTES:
The I Do shows how the three skills connect. The same kernel can be expanded with detail (L18), positioned as a TS in a paragraph (L19-20) or made into a complex sentence (L21). The integration is the point.

WATCH FOR:
- Students who say "we already did this" -- yes, that is the goal of review
- Students who notice the same kernel growing three ways -- excellent

[General: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Together. Mini-whiteboards out
- Three short rounds. Each round uses a different skill on a sentence about Chapter 42 or 43
- Round 1: EXPAND -- "The Sergeant returned." Add WHEN and WHERE detail
- Round 2: COMPLEX SENTENCE -- finish "Although the food was scarce,"
- Round 3: TOPIC SENTENCE for a body paragraph about "How the Eora woman helped the boys"

DO:
- Run three short rounds (60 seconds each)
- Hold up boards after each round
- Pick 1-2 strong examples to show the class

CFU CHECKPOINT:
Technique: Show Me Boards (three rounds)

Script:
- Round 1: "Expand 'The Sergeant returned.' Add WHEN and WHERE. Boards up in 60 seconds"
- Scan for: an adverbial of when (e.g. "Late that evening") and an adverbial of where (e.g. "to the cove")
- Round 2: "Finish 'Although the food was scarce,' with a full independent clause. Boards up"
- Scan for: an independent clause that stands alone (e.g. "the boys still helped tend the garden")
- Round 3: "Write a topic sentence for a body paragraph about 'how the Eora woman helped the boys.' Boards up"
- Scan for: a TS that signals WHEN, WHERE, WHO and WHAT (e.g. "During their fishing trip on Sunday, the boys learned new ways to find food from an Eora woman who shared her knowledge generously.")

PROCEED (>=80% on each round): Continue to the next round, then release to You Do.
PIVOT (<80% on any round):
- Expansion pivot: most likely issue -- adjectives only, no adverbials. Reteach: "Adverbials answer when, where, how. Add ONE of those"
- Complex sentence pivot: most likely issue -- the ending is a fragment. Reteach: "Read it alone -- does it make sense?"
- TS pivot: most likely issue -- TS is too vague. Reteach by example: "Add WHEN and WHERE to the front. Tell the reader more"

TEACHER NOTES:
Three short rounds rebuild three skills in one go. The body paragraph (full structure) skill comes back in the You Do, where students apply it to a Chapter 42-43 topic.

WATCH FOR:
- Students whose expansion just adds adjectives -- redirect to adverbials
- Students whose complex sentence ending is a fragment -- redirect to placement
- Students whose TS is just a fact -- redirect to specifics

[General: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU = `SAY:
- Quick check before you write. I will show you a sentence. Hold up fingers
- 1 = Sentence Expansion  |  2 = Complex Sentence
- Look carefully -- which technique was used?

DO:
- Read the sentence on the screen
- Use Finger Voting (1 / 2)
- Cold call after the vote: "Why did you choose that?"

CFU CHECKPOINT:
Technique: Finger Voting (1 / 2)

Script:
- Sentence: "Even though the food was scarce, the colony still tried to celebrate Christmas at the cove."
- "Is this expansion or a complex sentence? Fingers up"
- Scan for: most students choose 2

PROCEED (>=80% choose 2): Reveal and discuss -- "Even though" is a subordinating conjunction. There are two clauses joined together
PIVOT (<80%): Most likely issue -- students confuse "extra detail" with "complex sentence". Reteach: "Expansion adds detail to ONE clause. A complex sentence has TWO clauses joined by a subordinating conjunction"

TEACHER NOTES:
This CFU targets the difference between expansion (one clause, more detail) and a complex sentence (two clauses). The reveal acknowledges that both make sentences longer, but the structure is different.

WATCH FOR:
- Students who answer 1 -- prompt them: "Is there one part or two parts? Can you split it?"
- Students who answer 2 with reasoning -- ready to write
- Students who notice "Even though" -- excellent

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Your turn. The review worksheet has three sections, one for each skill we have built
- 15 minutes. Work through all three sections
- I will circulate

DO:
- Distribute the Lessons 18-21 Review Practice
- Circulate and confer
- Have the answer key out for quick checks (teacher copy only)

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the worksheet, but with one example completed in each section as a model
- Extra Notes: These students can refer to the I Do model on screen for parallel structure

EXTENDING PROMPT:
- Task: After completing the worksheet, take your own body paragraph from Lesson 20 and rewrite TWO of its sentences -- one expanded with more adverbial detail, one rewritten as a complex sentence. Underline the change
- Extra Notes: Push specificity -- use proper nouns and details from Chapters 36-43

TEACHER NOTES:
The three sections target the three skills directly. Students should be able to complete them in 15 minutes. If you see one section consistently weak across the class, that is the target for next-day reteaching.

WATCH FOR:
- Students who finish quickly -- direct them to the extension
- Students who get stuck on expansion -- the adverbial bank is at the top of the worksheet
- Students who get stuck on complex sentences -- the conjunction bank is at the top of the worksheet

[General: You Do | VTLM 2.0: Independent Application]`;

const NOTES_CLOSING = `SAY:
- Quick self-check against the success criteria. Thumbs for each
- Then turn to a partner: tell them WHICH of the three skills you feel strongest on, and WHICH you want more practice on
- Tomorrow we move forward into new chapters

DO:
- Run thumbs check for each SC
- Listen in on partner shares -- this tells you where to spend reteach time
- Wrap up: "These three skills come back across the unit. We will keep practising them"

TEACHER NOTES:
Use the partner share as a quick survey. If most students name the same skill as 'want more practice', plan a 10-minute warm-up on it tomorrow.

WATCH FOR:
- Students who name a different skill from their actual performance on the worksheet -- check in privately
- Students who feel strong across all three -- celebrate

[General: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Lessons 18-21 Review -- Sentence and Paragraph Skills";

  // Slide 1 -- Title
  titleSlide(
    pres,
    "Lessons 18-21 Review",
    "Sentence Expansion, Body Paragraph & Complex Sentences",
    "Year 5/6 Literacy",
    NOTES_TITLE
  );

  // Slide 2 -- Resources (immediately after title)
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
    "What Have We Learned This Week?",
    [
      "Mini-whiteboards out",
      "60 seconds: name the skills we have built across the past four lessons",
      "Best you can. Names if you remember; descriptions are fine",
      "We will reveal together",
    ],
    NOTES_LAUNCH,
    FOOTER
  );

  // Slide 4 -- LI/SC
  liSlide(
    pres,
    [
      "We are reviewing the three skills from the past week: sentence expansion, body paragraph structure and writing, and complex sentences using subordinating conjunctions",
    ],
    [
      "I can expand a sentence by adding adverbial detail (when, where, how)",
      "I can name the parts of a body paragraph and explain its language features",
      "I can write a complex sentence using a subordinating conjunction",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // Slide 5 -- Recall: Sentence Expansion (L18)
  contentSlide(
    pres,
    "Recall -- Lesson 18",
    C.PRIMARY,
    "Sentence Expansion",
    [
      "Start with a kernel sentence -- the bare bones",
      "Add detail: WHO, WHAT DOING, WHEN, WHERE, WHY, HOW",
      "Adverbials answer when, where, how -- e.g. \"on Christmas morning\", \"at the cove\", \"quietly\"",
      "If WHEN comes first, use a comma after it -- e.g. \"On Christmas morning, the boys waited.\"",
    ],
    NOTES_RECALL_EXPAND,
    FOOTER
  );

  // Slide 6 -- Recall: Body Paragraph (L19-20)
  contentSlide(
    pres,
    "Recall -- Lessons 19-20",
    C.SECONDARY,
    "Body Paragraph for an Information Report",
    [
      "Topic Sentence  ->  Supporting Details  ->  Concluding Sentence",
      "TS signals when, where, who, what. SDs add facts. CS adds the final idea (not a repeat)",
      "Features: past tense, third person, ONE appositive or relative clause",
      "Lesson 19 = PLAN with an SPO. Lesson 20 = WRITE the paragraph",
    ],
    NOTES_RECALL_BODYPARA,
    FOOTER
  );

  // Slide 7 -- Recall: Subordinating Conjunctions (L21)
  contentSlide(
    pres,
    "Recall -- Lesson 21",
    C.ACCENT,
    "Complex Sentences",
    [
      "Dependent clause cannot stand alone. Independent clause stands alone",
      "Joining word: subordinating conjunction -- even though, since, after, because, when, while, although, if",
      "Dependent clause first?  Use a comma after it",
      "Example: \"Although the food was scarce, the colony still tried to celebrate Christmas.\"",
    ],
    NOTES_RECALL_SUBCONJ,
    FOOTER
  );

  // Slide 8 -- I Do: model
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "One Sentence, Three Skills",
    "Plain sentence:\n\n\"The boys met an Eora woman.\"\n\nWatch me grow this sentence\nin THREE different ways\nusing the three skills\nfrom this week.",
    "1) Sentence expansion:\n\"On a fishing trip on Sunday,\nthe boys met an Eora woman\nat the rocky shore.\"\n\n2) Used as a topic sentence:\n\"During a fishing trip on a Sunday,\nthe boys met an Eora woman who\nshared her food knowledge generously.\"\n(then 3 supporting details + CS)\n\n3) Complex sentence:\n\"After the boys met an Eora woman,\nthey began to see their new home\nin a different way.\"",
    NOTES_IDO,
    FOOTER
  );

  // Slide 9 -- We Do
  contentSlide(
    pres,
    "We Do",
    C.SECONDARY,
    "Together: Three Quick Rounds",
    [
      "Mini-whiteboards out -- three short rounds",
      "Round 1: EXPAND \"The Sergeant returned.\" Add WHEN and WHERE detail",
      "Round 2: COMPLEX SENTENCE -- finish \"Although the food was scarce,\"",
      "Round 3: Write a TS for a body paragraph about: \"How the Eora woman helped the boys\"",
    ],
    NOTES_WEDO,
    FOOTER
  );

  // Slide 10 + 11 -- CFU with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Expansion or Complex Sentence?",
      "1 = Sentence Expansion  |  2 = Complex Sentence",
      "\"Even though the food was scarce, the colony still tried to celebrate Christmas at the cove.\"\n\nIs this an EXPANDED sentence or a COMPLEX sentence? Fingers: 1 or 2?",
      NOTES_CFU,
      FOOTER
    ),
    (s) => {
      addCard(s, 0.5, SAFE_BOTTOM - 0.95, 9, 0.85, { strip: C.SUCCESS, fill: C.WHITE });
      s.addText("2 -- COMPLEX SENTENCE. \"Even though\" joins TWO clauses, each with its own subject and verb.", {
        x: 0.75, y: SAFE_BOTTOM - 0.88, w: 8.4, h: 0.70,
        fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true,
        valign: "middle", margin: 0,
      });
      s.addNotes(`SAY:
- The answer is 2 -- complex sentence
- "Even though" is a subordinating conjunction. It joins two clauses
- The first clause cannot stand alone. The second clause can
- Sentence expansion adds detail to ONE clause. A complex sentence has TWO clauses joined together

DO:
- Reveal the answer
- If most chose 1, run the test: "Can you split it into two sentences that each make sense?"
- If most chose 2, celebrate the reasoning

[General: CFU Reveal | VTLM 2.0: Formative Feedback]`);
    }
  );

  // Slide 12 -- You Do (instruction card + reference)
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Lessons 18-21 Mixed Practice");

    addInstructionCard(s, [
      { text: "On your worksheet:", role: "header" },
      { text: "First: Read each section heading -- it tells you which skill" },
      { text: "Next: Work through the three sections in order" },
      { text: "Then: Use the adverbial and conjunction banks at the top" },
      { text: "Last: If you finish, try the optional extension at the bottom" },
    ], {
      x: 0.5, y: CONTENT_TOP, w: 9, h: 1.85,
      strip: C.PRIMARY, fill: C.WHITE,
      headerColor: C.PRIMARY,
    });

    const tipY = CONTENT_TOP + 2.00;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Quick Reference", {
      x: 0.75, y: tipY + 0.10, w: 5, h: 0.32,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText([
      { text: "Adverbials answer WHEN, WHERE or HOW", options: { breakLine: true } },
      { text: "Body paragraph = TS + supporting details + CS", options: { breakLine: true } },
      { text: "Complex sentence = dependent clause + independent clause + subordinating conjunction" },
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
    {
      reflectionPrompt: "Tell a partner WHICH skill you feel strongest on and WHICH you want more practice on.",
      scItems: [
        "I can expand a sentence by adding adverbial detail",
        "I can name the parts of a body paragraph and explain its language features",
        "I can write a complex sentence using a subordinating conjunction",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (got it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // -----------------------------------------------------------------------
  // Companion PDFs
  // -----------------------------------------------------------------------

  // Practice
  const ws = createPdf({ title: PRACTICE_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Lessons 18-21 Review Practice", {
    color: C.PRIMARY,
    subtitle: "Mixed Skills: Sentence Expansion, Body Paragraph, Complex Sentences",
    lessonInfo: "Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "Three short sections, one for each skill from the past week. Work through them in order. Use the quick reference below for help. If you finish, try the optional extension at the bottom.", wsY, { color: C.PRIMARY });

  wsY = addSectionHeading(ws, "Quick Reference", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Adverbials -- when (in 1788, on Sunday); where (at the cove, near the gardens); how (quickly, carefully)", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "Body paragraph parts: TS (topic sentence)  |  SD (supporting details)  |  CS (concluding sentence)", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "Subordinating conjunctions: even though, since, after, because, when, while, although, if", wsY, { fontSize: 11 });
  wsY += 8;

  // Section 1: Sentence Expansion (L18)
  wsY = addSectionHeading(ws, "Section 1 -- Sentence Expansion (Lesson 18)", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Expand each kernel sentence. Add a WHEN, WHERE or HOW detail (or all three).", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "a) The boys waited.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY = addBodyText(ws, "b) The Sergeant brought food.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY = addBodyText(ws, "c) Tom watched the harbour.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY += 6;

  // Section 2: Body Paragraph parts and features (L19-20)
  wsY = addSectionHeading(ws, "Section 2 -- Body Paragraph Spot Check (Lessons 19-20)", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Read the paragraph below. Label each sentence in the margin: TS, SD or CS.", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "(1) When the First Fleet arrived in Sydney Cove in January 1788, the British settlers faced enormous challenges in setting up a colony on Eora land. (2) The cove sat on the traditional Country of the Eora people, who had cared for that land for tens of thousands of years. (3) The soil near the cove was thin and unsuitable for British wheat. (4) Convicts and marines struggled with food shortages, poor housing and disease. (5) These early years were marked by hardship for the British and devastation for the Eora people.", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "Sentence 1 = ____    Sentence 2 = ____    Sentence 3 = ____    Sentence 4 = ____    Sentence 5 = ____", wsY, { fontSize: 11 });
  wsY += 6;

  wsY = addBodyText(ws, "Find ONE example of each language feature in the paragraph above.", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "a) Past tense verb:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY = addBodyText(ws, "b) Third person reference:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY = addBodyText(ws, "c) Adverbial (when, where or how):", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY += 6;

  // Section 3: Complex Sentences (L21)
  wsY = addSectionHeading(ws, "Section 3 -- Complex Sentences (Lesson 21)", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "Finish each dependent clause with a full independent clause. Use content from Chapters 36-43.", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "a) Because the food rations were low,", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY = addBodyText(ws, "b) When the supply ship did not arrive,", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY = addBodyText(ws, "c) Although the colony struggled,", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY += 8;

  // Optional extension
  wsY = addSectionHeading(ws, "Optional Extension", wsY, { color: C.ALERT });
  wsY = addBodyText(ws, "Take your own body paragraph from Lesson 20. Rewrite TWO of its sentences -- one EXPANDED with new adverbial detail, one rewritten as a COMPLEX SENTENCE with a subordinating conjunction. Underline both changes.", wsY, { fontSize: 10 });
  wsY = addLinedArea(ws, wsY, 6, { lineSpacing: 22 });

  addPdfFooter(ws, "Lessons 18-21 Review Practice");

  // Answer key
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Lessons 18-21 Review -- Model Answers", {
    color: C.ALERT,
    subtitle: "Teacher Reference -- Mixed Skills",
    lessonInfo: "Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Model answers below. Student answers in Sections 1 and 3 will vary -- accept any answer that fits the prompt and is grammatically correct.", akY, { color: C.ALERT });

  akY = addSectionHeading(ak, "Section 1 -- Sentence Expansion", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "a) Model: \"On Christmas morning, the boys waited quietly outside the Sergeant's hut.\"", akY);
  akY = addBodyText(ak, "b) Model: \"At noon, the Sergeant brought home a small parcel of dried meat.\"", akY);
  akY = addBodyText(ak, "c) Model: \"From the doorway, Tom watched the harbour anxiously for the supply ship.\"", akY);
  akY = addBodyText(ak, "Look for: at least one adverbial (when, where or how); correct punctuation if WHEN comes first; the kernel meaning preserved.", akY, { fontSize: 10, italic: true });
  akY += 8;

  akY = addSectionHeading(ak, "Section 2 -- Body Paragraph Spot Check", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "Sentence 1 = TS  |  Sentences 2, 3, 4 = SD  |  Sentence 5 = CS", akY);
  akY = addBodyText(ak, "TS signals when (January 1788), where (Sydney Cove), who (British settlers) and what (challenges). SDs add specific facts (Eora Country, thin soil, food shortages). CS adds the FINAL idea (dual impact) without repeating the TS.", akY, { fontSize: 10, italic: true });
  akY += 6;
  akY = addBodyText(ak, "a) Past tense verb -- e.g. \"arrived\", \"faced\", \"sat\", \"was\", \"struggled\", \"were marked\".", akY);
  akY = addBodyText(ak, "b) Third person reference -- e.g. \"the British settlers\", \"the Eora people\", \"convicts and marines\".", akY);
  akY = addBodyText(ak, "c) Adverbial -- \"in Sydney Cove\" (where), \"in January 1788\" (when), \"near the cove\" (where), \"for tens of thousands of years\" (how long / when).", akY);
  akY = addBodyText(ak, "Accept any reasonable example pulled directly from the paragraph.", akY, { fontSize: 10, italic: true });
  akY += 8;

  akY = addSectionHeading(ak, "Section 3 -- Complex Sentences", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "a) Model: \"Because the food rations were low, the colony cut every meal in half.\"", akY);
  akY = addBodyText(ak, "b) Model: \"When the supply ship did not arrive, the Sergeant grew quiet and worried.\"", akY);
  akY = addBodyText(ak, "c) Model: \"Although the colony struggled, the boys still kept watch over the gardens.\"", akY);
  akY = addBodyText(ak, "Look for: an independent clause that stands alone; a comma after the dependent clause; content drawn from Chapters 36-43.", akY, { fontSize: 10, italic: true });
  akY += 12;

  akY = addSectionHeading(ak, "What to Look For (across all sections)", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- Section 1: an adverbial of when, where or how; comma after WHEN if it comes first", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Section 2: TS opens with the WHEN/WHERE; SDs sit between; CS adds a final idea", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Section 3: independent clause stands alone; no fragments starting with 'and' or 'but'", akY, { fontSize: 10 });

  addPdfFooter(ak, "Review -- Answer Key -- TEACHER COPY");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Review_Lessons18to21.pptx` }),
    writePdf(ws, PRACTICE_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Review_Lessons18to21.pptx`);
  console.log("Done: " + PRACTICE_RESOURCE.name);
  console.log("Done: " + ANSWER_KEY_RESOURCE.name);
}

build().catch(console.error);
