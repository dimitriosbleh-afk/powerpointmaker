"use strict";

// Tom Unit -- Review of Lessons 14-17
// Year 5/6 Literacy. Sits at the end of the Information Report Body Paragraph 2 +
// Sentence-Level Skills sequence (Lessons 14-17). Consolidates four prior skills:
//   L14: Plan body paragraph 2 with an SPO (Why the First Fleet came to Australia)
//   L15: Write body paragraph 2 from the SPO (language features of an info report)
//   L16: Appositives -- adding noun-phrase detail beside a noun
//   L17: Relative clauses -- adding who / which / that / where / when / why detail
// No new chapters. No new content. Mixed-skills review with one consolidated practice sheet.

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const { createTheme, weekToVariant } = require("../themes/factory");
// Unit cohesion: all Tom lessons use the same variant (matches sessions 6-17).
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

const SESSION_NUMBER = 18;
const FOOTER = "Lessons 14-17 Review | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Review_Lessons14to17";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PRACTICE_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mixed Skills Review",
  "Mixed-skills practice covering body paragraph structure, info-report language features, appositives and relative clauses.",
  {
    name: "Lessons 14 to 17 Review Practice",
    fileName: path.posix.join(getSessionResourceFolder(SESSION_NUMBER), "Lessons 14 to 17 Review Practice.pdf"),
  }
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Mixed Skills Review Answer Key",
  "Teacher reference: model answers across all four review sections.",
  {
    name: "Lessons 14 to 17 Review Answer Key",
    fileName: path.posix.join(getSessionResourceFolder(SESSION_NUMBER), "Lessons 14 to 17 Review Answer Key.pdf"),
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
- We will revisit four skills: body paragraph structure, info-report language features, appositives and relative clauses
- Bring your previous work -- your SPO from Lesson 14 and your written body paragraph from Lesson 15

DO:
- Display title slide as students settle
- Have novels and exercise books on desks
- Have students' SPOs and written body paragraphs from the past week within reach

TEACHER NOTES:
This is a consolidation lesson. The aim is to help students see how the four skills connect: SPO -> body paragraph -> enriched with appositives and relative clauses. One mixed-skills worksheet supports independent practice.

WATCH FOR:
- Students who cannot find their SPO or body paragraph -- pair them with a peer for the recall
- Students who are confident -- they can take the lead in partner work

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The Lessons 14-17 Review Practice is one mixed-skills worksheet covering all four skills
- The Lessons 14-17 Review Answer Key is teacher reference

DO:
- Print the practice worksheet (one per student)
- Print the answer key (teacher copy only)
- Have students' own previous work (SPO from L14, body paragraph from L15) within reach for the launch

TEACHER NOTES:
The single worksheet keeps cognitive load low. Differentiation is built in -- enabling students get a partially-completed version as the model; extending students rewrite a body paragraph from earlier in the week using the new sentence-level moves.

[General: Resources | VTLM 2.0: Student Resources]`;

const NOTES_LAUNCH = `SAY:
- Quick recall. On your whiteboard, write the four skills we have built across the past four lessons
- 60 seconds. Best you can. Names if you remember them; descriptions are fine too
- Some of you may remember all four. If you remember one or two, that is fine

DO:
- 60 seconds silent recall on whiteboards
- Walk and scan: who has all four? who has two or three? who is stuck?
- Reveal the four skills on the next slides

TEACHER NOTES:
Active recall before the LI/SC slide. The four skills are body paragraph plan / write (Lessons 14-15), appositives (Lesson 16), relative clauses (Lesson 17). Use the launch as a quick formative scan.

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
SC1 targets information report structure (Lessons 14-15). SC2 targets appositives (Lesson 16). SC3 targets relative clauses (Lesson 17). The criteria appear as a single unlabelled list.

WATCH FOR:
- Students who already feel confident on all three -- they can challenge themselves on the optional extension
- Students unsure about one strand -- they will see it modelled and supported

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_RECALL_BODYPARA = `SAY:
- Lessons 14 and 15 review. Body paragraph for an information report
- Three parts: topic sentence, supporting details, concluding sentence
- The TS introduces the main idea. The supporting details expand on it. The CS wraps it up without repeating the TS
- Lesson 14 was the PLAN. Lesson 15 was the WRITE

DO:
- Display the structure card
- Quick partner check: "Tell your partner ONE thing each part of a body paragraph does"

TEACHER NOTES:
Keep the recall tight. Do NOT reteach -- this is a refresher. If students are shaky, the You Do mixed practice will catch any gaps.

WATCH FOR:
- Students who confuse TS and CS -- prompt: "TS comes first; CS comes last; CS does not repeat the TS word for word"

[General: Recall | VTLM 2.0: Retention and Recall]`;

const NOTES_RECALL_FEATURES = `SAY:
- Lesson 15 review. Language features of a body paragraph in an information report
- Present tense -- "convicts WERE sent" stays in past for the historical topic, but the report VOICE stays factual
- Third person pronouns -- they, them, their (not I or you)
- Nouns and noun groups, with adjectives that describe, classify, compare or quantify
- Adverbials -- where, when, how (e.g. "after 1788", "in Botany Bay", "at great cost")
- Subheadings for emphasis when the report has many sections

DO:
- Display the features card
- Quick oral check: "Give me a noun group with one adjective from the topic of the First Fleet"
- Possible answers: "the overcrowded prisons", "the long voyage", "the British government"

TEACHER NOTES:
Refresh, not reteach. The point is for students to recognise these features when they see them. They do not need to name them all to use them well.

WATCH FOR:
- Students who slip into first person ("I think the prisons were full") -- redirect to third person
- Students who give a noun group with multiple adjectives -- celebrate

[General: Recall | VTLM 2.0: Retention and Recall]`;

const NOTES_RECALL_APPOS = `SAY:
- Lesson 16 review. An appositive is a noun or noun group that sits NEXT TO another noun and renames or describes it
- We use commas around the appositive when it adds extra information
- Example: "Sergeant Stanley, Rob's father, was a kind man." The bit between commas is the appositive
- Without the appositive: "Sergeant Stanley was a kind man." The sentence still works -- the appositive just adds detail

DO:
- Display the appositive card
- Quick oral check: "Add an appositive to: 'Tom worked in the garden.' What could you add about Tom?"
- Possible: "Tom, a young convict, worked in the garden."

TEACHER NOTES:
Refresh the comma rule. Most students will recall the structure; the worksheet will check that they can add an appositive that genuinely adds information.

WATCH FOR:
- Students who forget the second comma -- common error. Both commas are needed
- Students who add an appositive that just repeats the noun -- prompt: "Add NEW information"

[General: Recall | VTLM 2.0: Retention and Recall]`;

const NOTES_RECALL_RELCLAUSE = `SAY:
- Lesson 17 review. A relative clause is a group of words that gives extra information about a noun
- It starts with a relative pronoun (who, whom, which, whose, that) or a relative adverb (when, where, why)
- It sits IMMEDIATELY AFTER the noun it describes
- Example: "The seeds, which the Sergeant had planted, began to grow." The clause between commas describes "the seeds"

DO:
- Display the relative clause card
- Quick oral check: "Add a 'who' clause to: 'Rob warned the men.' Who is Rob?"
- Possible: "Rob, who had returned home, warned the men."

TEACHER NOTES:
Refresh the placement rule -- relative clauses sit right after the noun they describe. Keep the essential vs non-essential distinction brief: if the meaning of the sentence depends on the clause, no commas; if it is extra information, use commas.

WATCH FOR:
- Students who place the clause too far from the noun -- prompt: "Move the clause RIGHT NEXT TO the thing it describes"
- Students who confuse appositives and relative clauses -- both add detail; appositives use a noun group, relative clauses use a clause with a verb

[General: Recall | VTLM 2.0: Retention and Recall]`;

const NOTES_IDO = `SAY:
- Watch me. I am going to take ONE plain sentence about the First Fleet and use ALL FOUR skills on it
- Plain sentence: "The convicts arrived at Botany Bay."
- First the body paragraph -- I'll position this sentence as a topic sentence and add two supporting details and a CS
- Then I'll add an appositive to one of the sentences
- Last I'll add a relative clause to another sentence to thicken the description
- Notice: same content, four different ways to capture or use it

DO:
- Display the I Do model
- Think aloud through each move
- Highlight: same plain sentence, but the writing gets richer at each step

TEACHER NOTES:
The I Do shows how the four skills connect. The same kernel sentence can be planned (TS), expanded (supporting details + CS), and enriched (appositives + relative clauses). The integration is the point.

WATCH FOR:
- Students who say "we already did this" -- yes, that is the goal of review
- Students who notice the connection across skills -- excellent

[General: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Together. Mini-whiteboards out
- Three short rounds. Same Sergeant Stanley sentence each round, but we are doing different jobs on it
- Round 1: Add an APPOSITIVE to: "Sergeant Stanley returned with the rations."
- Round 2: Add a RELATIVE CLAUSE to: "The garden produced fresh vegetables."
- Round 3: Write a TOPIC SENTENCE for a body paragraph about: "Why fresh food mattered to the colony."

DO:
- Run three short rounds (60 seconds each)
- Hold up boards after each round
- Pick 1-2 strong examples to show the class

CFU CHECKPOINT:
Technique: Show Me Boards (three rounds)

Script:
- Round 1: "Add an appositive to 'Sergeant Stanley returned with the rations.' Boards up in 60 seconds"
- Scan for: a noun group between commas right after Stanley (e.g. "Sergeant Stanley, the boys' guardian, returned with the rations.")
- Round 2: "Add a relative clause to 'The garden produced fresh vegetables.' Boards up"
- Scan for: a relative clause that sits right next to "the garden" or "fresh vegetables" (e.g. "The garden, which the boys had planted, produced fresh vegetables.")
- Round 3: "Write a topic sentence for a body paragraph about 'why fresh food mattered to the colony.' Boards up"
- Scan for: a sentence that names WHO (the colony / the convicts), WHAT DOING (depended on / needed) and WHY (to prevent disease / scurvy)

PROCEED (>=80% on each round): Continue to the next round, then release to You Do.
PIVOT (<80% on any round):
- Appositive pivot: most likely issue -- students forget the second comma. Reteach: "Both commas are needed. The appositive sits between them"
- Relative clause pivot: most likely issue -- the clause does not sit next to the noun. Reteach: "Move it right next to the thing it describes"
- TS pivot: most likely issue -- TS is too vague. Reteach by example: "Fresh food was important" is not enough. Add the WHY

TEACHER NOTES:
Three short rounds rebuild three of the four skills in one go. The body paragraph (full structure) skill comes back in the You Do, where students draft from scratch using all four moves.

WATCH FOR:
- Students whose appositive only repeats the noun -- redirect to NEW information
- Students whose relative clause is far from the noun -- redirect to placement
- Students whose TS still says "Fresh food was important" -- redirect to specifics

[General: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU = `SAY:
- Quick check before you write. I will show you a sentence with extra detail added. Hold up fingers
- 1 = Appositive (a renaming noun group, between commas)
- 2 = Relative Clause (a clause that starts with who/which/that/where/when/why)

DO:
- Read the sentence on the screen
- Use Finger Voting (1 / 2)
- Cold call after the vote: "Why did you choose that?"

CFU CHECKPOINT:
Technique: Finger Voting (1 / 2)

Script:
- "Tom, who had been locked in a dark room, told the Major the truth."
- "Is the underlined part an appositive or a relative clause? Fingers up"
- Scan for: most students choose 2

PROCEED (>=80% choose 2): Reveal and discuss why -- "who had been locked" starts with WHO and contains a verb
PIVOT (<80%): Most likely issue -- students confuse the two. Reteach: "An appositive is a NOUN GROUP. A relative clause has a VERB and starts with who/whom/which/whose/that/when/where/why. This one has 'had been locked' -- that is a verb"

TEACHER NOTES:
This CFU targets the appositive vs relative clause distinction, which is the most common confusion across Lessons 16 and 17. The reveal acknowledges that both add detail but the structure is different.

WATCH FOR:
- Students who answer 1 -- prompt them: "Is there a verb in the underlined part?"
- Students who answer 2 with reasoning -- ready to write
- Students who notice the comma rule applies to both -- excellent

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Your turn. The review worksheet has four short sections, one for each skill from the past week
- 15 minutes. Work through all four sections
- I will circulate

DO:
- Distribute the Lessons 14-17 Review Practice
- Circulate and confer
- Have the answer key out for quick checks (teacher copy only)

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the worksheet, but with the section headings highlighted and one example completed in each section as a model
- Extra Notes: These students can refer to the I Do model on screen for parallel structure

EXTENDING PROMPT:
- Task: After completing the worksheet, take your own body paragraph from Lesson 15 and rewrite TWO of its sentences -- one with an added appositive, one with an added relative clause. Underline both
- Extra Notes: Push the language features check too -- third person, present-tense report voice, noun groups with adjectives

TEACHER NOTES:
The four sections target the four skills directly. Students should be able to complete them in 15 minutes. If you see one section consistently weak across the class, that is the target for next-day reteaching.

WATCH FOR:
- Students who finish quickly -- direct them to the extension
- Students who get stuck on appositives -- the model on the worksheet shows the comma rule
- Students who get stuck on relative clauses -- the relative pronoun bank is at the top of the worksheet

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

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Lessons 14-17 Review -- Information Report and Sentence Skills";

  // Slide 1 -- Title
  titleSlide(
    pres,
    "Lessons 14-17 Review",
    "Information Report & Sentence Skills",
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
      "We are reviewing the four skills from the past week: structure of a body paragraph in an information report, language features of an information report, adding appositives and adding relative clauses",
    ],
    [
      "I can name the parts of a body paragraph and the language features of an information report",
      "I can add an appositive to a sentence using commas correctly",
      "I can add a relative clause that sits next to the noun it describes",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // Slide 5 -- Recall: Body Paragraph (Lessons 14-15)
  contentSlide(
    pres,
    "Recall -- Lessons 14-15",
    C.PRIMARY,
    "Body Paragraph for an Information Report",
    [
      "Topic Sentence (TS) -- introduces the main idea: who, what doing, when, where, why, how",
      "Supporting Details -- sentences that expand on the main idea with facts from research",
      "Concluding Sentence (CS) -- wraps up the paragraph without repeating the TS",
      "Lesson 14 = PLAN with an SPO. Lesson 15 = WRITE the paragraph",
    ],
    NOTES_RECALL_BODYPARA,
    FOOTER
  );

  // Slide 6 -- Recall: Language Features (Lesson 15)
  contentSlide(
    pres,
    "Recall -- Lesson 15",
    C.SECONDARY,
    "Language Features of an Information Report",
    [
      "Third person pronouns -- they, them, their (not I or you)",
      "Nouns and noun groups -- adjectives that describe, classify, compare or quantify",
      "Adverbials -- where, when, how (e.g. \"in Botany Bay\", \"after 1788\", \"at great cost\")",
      "Subheadings for emphasis when the report has many sections",
    ],
    NOTES_RECALL_FEATURES,
    FOOTER
  );

  // Slide 7 -- Recall: Appositives (Lesson 16)
  contentSlide(
    pres,
    "Recall -- Lesson 16",
    C.ACCENT,
    "Appositives",
    [
      "An appositive is a NOUN or NOUN GROUP that sits next to another noun and renames or describes it",
      "Use commas around the appositive when it adds extra information",
      "Example: \"Sergeant Stanley, Rob's father, was a kind man.\"",
      "Both commas are needed -- one to open the appositive and one to close it",
    ],
    NOTES_RECALL_APPOS,
    FOOTER
  );

  // Slide 8 -- Recall: Relative Clauses (Lesson 17)
  contentSlide(
    pres,
    "Recall -- Lesson 17",
    C.PRIMARY,
    "Relative Clauses",
    [
      "A relative clause is a group of words (with a verb) that gives extra information about a noun",
      "It starts with: who, whom, which, whose, that, when, where, why",
      "It sits RIGHT AFTER the noun it describes",
      "Example: \"The seeds, which the Sergeant had planted, began to grow.\"",
    ],
    NOTES_RECALL_RELCLAUSE,
    FOOTER
  );

  // Slide 9 -- I Do: model the four skills on one moment
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "One Sentence, Four Skills",
    "Plain sentence:\n\n\"The convicts arrived at\nBotany Bay.\"\n\nWatch me use all FOUR\nskills to grow this into\na rich body paragraph\nfor an information report.",
    "1) Body paragraph (TS):\n\"In January 1788, the convicts\nof the First Fleet arrived at\nBotany Bay after an exhausting\neight-month sea voyage.\"\n\n2) Add an APPOSITIVE:\n\"The convicts, mostly poor\nfamilies and petty thieves,\nstepped onto the new land.\"\n\n3) Add a RELATIVE CLAUSE:\n\"Botany Bay, which Captain\nCook had recommended, was\nnot the rich farmland the\nBritish had hoped for.\"",
    NOTES_IDO,
    FOOTER
  );

  // Slide 10 -- We Do
  contentSlide(
    pres,
    "We Do",
    C.SECONDARY,
    "Together: Three Quick Rounds",
    [
      "Mini-whiteboards out -- three short rounds",
      "Round 1: Add an APPOSITIVE to -- \"Sergeant Stanley returned with the rations.\"",
      "Round 2: Add a RELATIVE CLAUSE to -- \"The garden produced fresh vegetables.\"",
      "Round 3: Write a TOPIC SENTENCE for a body paragraph about -- \"Why fresh food mattered to the colony.\"",
    ],
    NOTES_WEDO,
    FOOTER
  );

  // Slide 11 + 11b -- CFU with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Appositive or Relative Clause?",
      "1 = Appositive  |  2 = Relative Clause",
      "\"Tom, who had been locked in a dark room, told the Major the truth.\"\n\nIs the part between commas an APPOSITIVE or a RELATIVE CLAUSE? Fingers: 1 or 2?",
      NOTES_CFU,
      FOOTER
    ),
    (s) => {
      addCard(s, 0.5, SAFE_BOTTOM - 0.95, 9, 0.85, { strip: C.SUCCESS, fill: C.WHITE });
      s.addText("2 -- RELATIVE CLAUSE. It starts with WHO and contains a verb (\"had been locked\")", {
        x: 0.75, y: SAFE_BOTTOM - 0.88, w: 8.4, h: 0.70,
        fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true,
        valign: "middle", margin: 0,
      });
      s.addNotes("SAY:\n- The answer is 2 -- relative clause\n- It starts with WHO and contains a VERB (had been locked)\n- An appositive would be a NOUN GROUP, like \"Tom, a young convict, told the Major the truth\"\n- Both add extra information; both use commas; the difference is structure\n\nDO:\n- Reveal the answer\n- If most chose 1, run the test: 'Is there a verb in the underlined part?'\n- If most chose 2, celebrate the reasoning\n\n[General: CFU Reveal | VTLM 2.0: Formative Feedback]");
    }
  );

  // Slide 12 -- You Do (instruction card + reference)
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Lessons 14-17 Mixed Practice");

    addInstructionCard(s, [
      { text: "On your worksheet:", role: "header" },
      { text: "First: Read each section heading -- it tells you which skill" },
      { text: "Next: Work through the four sections in order" },
      { text: "Then: Use the relative pronoun bank at the top for Section 4" },
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
      { text: "Body paragraph: TS + supporting details + CS", options: { breakLine: true } },
      { text: "Info report features: third person, noun groups, adverbials", options: { breakLine: true } },
      { text: "Appositive = noun group between commas. Relative clause = who, whom, which, whose, that, when, where, why" },
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
      "I can name the parts of a body paragraph and the language features of an information report",
      "I can add an appositive to a sentence using commas correctly",
      "I can add a relative clause that sits next to the noun it describes",
    ],
    NOTES_CLOSING
  );

  // -----------------------------------------------------------------------
  // Companion PDFs
  // -----------------------------------------------------------------------

  // Practice
  const ws = createPdf({ title: PRACTICE_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Lessons 14-17 Review Practice", {
    color: C.PRIMARY,
    subtitle: "Mixed Skills: Body Paragraph, Language Features, Appositives, Relative Clauses",
    lessonInfo: "Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "Four short sections, one for each skill from the past week. Work through them in order. Use the quick reference below for Sections 3 and 4. If you finish, try the optional extension at the bottom.", wsY, { color: C.PRIMARY });

  wsY = addSectionHeading(ws, "Quick Reference", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Body paragraph parts: TS (topic sentence)  |  SD (supporting details)  |  CS (concluding sentence)", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "Info report features: third person (they, them, their); noun groups with adjectives; adverbials of where / when / how", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "Appositive = a noun or noun group, set off by commas, that renames the noun beside it", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "Relative pronouns / adverbs: who, whom, which, whose, that, when, where, why", wsY, { fontSize: 11 });
  wsY += 8;

  // Section 1: Body paragraph parts (Lessons 14-15)
  wsY = addSectionHeading(ws, "Section 1 -- Body Paragraph Parts (Lessons 14-15)", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Read the paragraph below. Label each sentence in the margin: TS, SD or CS.", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "(1) The First Fleet sailed from Britain in 1787 because the British government needed somewhere new to send convicted prisoners. (2) Prisons in Britain had become dangerously overcrowded after the Industrial Revolution put many people out of work. (3) Britain had previously sent convicts to America, but the loss of the American colonies in 1783 closed that option. (4) The British also wanted to claim new land for trade and farming. (5) Together, these pressures pushed Britain to send the First Fleet halfway around the world to start a new colony at Botany Bay.", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "Sentence 1 = ____    Sentence 2 = ____    Sentence 3 = ____    Sentence 4 = ____    Sentence 5 = ____", wsY, { fontSize: 11 });
  wsY += 10;

  // Section 2: Language features (Lesson 15)
  wsY = addSectionHeading(ws, "Section 2 -- Spot the Language Features (Lesson 15)", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "In the paragraph above, find ONE example of each feature. Write each example on the line.", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "a) Third person pronoun:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY = addBodyText(ws, "b) A noun group with an adjective:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY = addBodyText(ws, "c) An adverbial (where, when or how):", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY += 6;

  // Section 3: Appositives (Lesson 16)
  wsY = addSectionHeading(ws, "Section 3 -- Add an Appositive (Lesson 16)", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "Add an appositive to each sentence. Remember: noun group between commas, adding NEW information.", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "a) Tom worked alongside Rob in the garden.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY = addBodyText(ws, "b) The Sergeant brought home the rations for the week.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY += 6;

  // Section 4: Relative Clauses (Lesson 17)
  wsY = addSectionHeading(ws, "Section 4 -- Add a Relative Clause (Lesson 17)", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Add a relative clause to each sentence. Use the relative pronoun bank above. Place it RIGHT AFTER the noun it describes.", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "a) The boys built a chimney.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY = addBodyText(ws, "b) Two convicts tried to steal from the garden.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 22 });
  wsY += 8;

  // Optional extension
  wsY = addSectionHeading(ws, "Optional Extension", wsY, { color: C.ALERT });
  wsY = addBodyText(ws, "Take your own body paragraph from Lesson 15. Rewrite TWO of its sentences -- one with an added appositive, one with an added relative clause. Underline both additions.", wsY, { fontSize: 10 });
  wsY = addLinedArea(ws, wsY, 6, { lineSpacing: 22 });

  addPdfFooter(ws, "Lessons 14-17 Review Practice");

  // Answer key
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Lessons 14-17 Review -- Model Answers", {
    color: C.ALERT,
    subtitle: "Teacher Reference -- Mixed Skills",
    lessonInfo: "Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Model answers below. Student answers in Sections 2, 3 and 4 will vary -- accept any answer that fits the prompt and is grammatically correct.", akY, { color: C.ALERT });

  akY = addSectionHeading(ak, "Section 1 -- Body Paragraph Parts", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sentence 1 = TS  |  Sentences 2, 3, 4 = SD  |  Sentence 5 = CS", akY);
  akY = addBodyText(ak, "TS introduces the main idea (Britain needed somewhere new for prisoners). SDs add specific reasons (overcrowded prisons, loss of America, new land for trade). CS wraps up with 'these pressures pushed Britain' without repeating the TS.", akY, { fontSize: 10, italic: true });
  akY += 8;

  akY = addSectionHeading(ak, "Section 2 -- Language Features", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "a) Third person pronoun -- examples: \"the British government\" (third person noun group); other valid pronoun examples include any third person reference within the paragraph", akY);
  akY = addBodyText(ak, "b) Noun group with adjective -- e.g. \"convicted prisoners\", \"dangerously overcrowded prisons\", \"the American colonies\", \"a new colony\"", akY);
  akY = addBodyText(ak, "c) Adverbial -- e.g. \"in 1787\" (when), \"halfway around the world\" (where / how), \"after the Industrial Revolution\" (when), \"at Botany Bay\" (where)", akY);
  akY = addBodyText(ak, "Accept any reasonable example pulled from the paragraph that fits the feature.", akY, { fontSize: 10, italic: true });
  akY += 8;

  akY = addSectionHeading(ak, "Section 3 -- Add an Appositive", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "a) Model: \"Tom, a young convict from England, worked alongside Rob in the garden.\"", akY);
  akY = addBodyText(ak, "b) Model: \"The Sergeant, a kind and fair man, brought home the rations for the week.\"", akY);
  akY = addBodyText(ak, "Look for: a noun group between commas; both commas present; new information about the noun (not just a repeat).", akY, { fontSize: 10, italic: true });
  akY += 8;

  akY = addSectionHeading(ak, "Section 4 -- Add a Relative Clause", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "a) Model: \"The boys, who had asked for the Sergeant's blessing, built a chimney.\"", akY);
  akY = addBodyText(ak, "b) Model: \"Two convicts, who were desperate for fresh food, tried to steal from the garden.\"", akY);
  akY = addBodyText(ak, "Look for: a relative pronoun (who / whom / which / whose / that) or relative adverb (when / where / why); the clause sits right after the noun; commas used when the clause adds extra information.", akY, { fontSize: 10, italic: true });
  akY += 12;

  akY = addSectionHeading(ak, "What to Look For (across all sections)", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- Section 1: TS opens, CS closes; SDs sit between with specific facts", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Section 2: students name examples directly from the paragraph", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Section 3: appositive between commas, with new information", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Section 4: relative clause sits right after the noun, uses a relative pronoun / adverb", akY, { fontSize: 10 });

  addPdfFooter(ak, "Review -- Answer Key -- TEACHER COPY");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Review_Lessons14to17.pptx` }),
    writePdf(ws, PRACTICE_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Review_Lessons14to17.pptx`);
  console.log("Done: " + PRACTICE_RESOURCE.name);
  console.log("Done: " + ANSWER_KEY_RESOURCE.name);
}

build().catch(console.error);
