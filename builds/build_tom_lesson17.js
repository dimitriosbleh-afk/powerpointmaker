"use strict";

// Tom Unit -- Lesson 17: Chapters 34-35 -- Building & Threats + Relative Clauses
// Week 4, Lesson 17, Grade 5/6 Literacy
// Reading: Ch 34 (Tom & Rob build a chimney), Ch 35 (Winter, disease, garden raid)
// Sentence-level: Identify and add relative (embedded) clauses
// Sensitivity: Ch 35 contains skinning/preparing a dead animal -- flagged in notes

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const { createTheme, weekToVariant } = require("../themes/factory");
// Unit cohesion: all Tom lessons use the same variant.
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

const SESSION_NUMBER = 17;
const FOOTER = "Chapters 34-35 | Lesson 17 | Week 4 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson17_Building_and_Threats_Relative_Clauses";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const RC_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Relative Clauses Practice",
  "Student worksheet: identify and add relative clauses (essential and non-essential) to sentences from Chapters 34-35."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Relative Clauses Practice Answer Key",
  "Teacher reference: model answers for relative clauses practice."
);
const RESOURCE_ITEMS = [RC_RESOURCE, ANSWER_KEY_RESOURCE];
const RC_PDF_PATH = path.join(OUT_DIR, RC_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Lesson 17. Chapters 34 and 35 today
- The chimney is built. Then winter comes, and with it disease and trouble in the colony
- After reading we return to relative clauses -- adding extra information using words like who, which, that

DO:
- Display title slide as students settle
- Have copies of the novel ready
- Have the relative clauses worksheet ready (do not distribute yet)

TEACHER NOTES:
This lesson combines novel reading (Chapters 34-35) with sentence-level writing (relative clauses). Chapter 35 contains content about skinning and preparing a dead animal -- this is brief and not graphic, but flag it for sensitive students before reading. The garden-raid scene shows Tom being attacked, which is intense. Pause for discussion if needed.

SENSITIVITY ADVISORY:
- What it is: Chapter 35 includes a brief passage about skinning a dead animal. There is also a scene where Tom is attacked by two men.
- Framing language: "There is a short part where the Sergeant prepares a kangaroo for cooking. There is also a scene where Tom is attacked by two convicts. Let me know if you need a moment."
- Watch for: students reacting to the violence, discomfort with the animal-preparation passage
- Protocol: Pause if needed. Students may step out briefly. Read the attack quickly without dwelling on detail; debrief after.

WATCH FOR:
- Students recapping previous chapters (Tom is settling in with Rob and the Sergeant)

[Literacy: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands again: reading the new chapters, then practising relative clauses
- Read the success criteria

DO:
- Choral read the LI, then the SCs
- Brief: "Today you will identify relative clauses and add them with correct punctuation"

TEACHER NOTES:
SC1 targets character and authorial-choice analysis. SC2 targets identifying a relative clause and the relative pronoun that signals it. SC3 targets adding a relative clause with correct comma use (essential vs non-essential).

WATCH FOR:
- Students who ask what makes a relative clause "essential" -- this is taught in the revise slide

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_READING = `SAY:
- Two chapters today. I will read aloud
- Three pause points along the way
- A note before we start: Chapter 35 has a brief section about preparing a dead animal, and a scene where Tom is attacked. Let me know if you need to step out
- Find Chapter 34 in your novel

DO:
- 30 seconds for students to find the chapter
- Read Chapter 34 aloud at a steady pace
- Pause point 1 at the end of Ch 34 / start of Ch 35
- Continue reading Ch 35 -- read briskly through the animal-preparation section
- Pause for the attack scene; some teachers may briefly summarise rather than read aloud, depending on the class
- Pause point 2 after the attack
- Pause point 3 at the end of Ch 35

TEACHER NOTES:
The reading takes 20-25 minutes including pauses. The attack scene is the most intense moment -- a teacher with sensitive students may choose to summarise rather than read it aloud. Keep the pause point discussions short (60 seconds each).

WATCH FOR:
- Students reacting to the attack -- normalise: "It is a tense moment in the story; the author wants us to feel that"
- Students who lose track of who is who -- reorient briefly

[Literacy: Read Aloud | VTLM 2.0: Build Knowledge / Reading]`;

const NOTES_PAUSE1 = `SAY:
- Stop. Tom and Rob have built a chimney for the house. The author writes that the autumn sun was "lending it some brightness"
- Turn to your partner: What is the significance of Rob and Tom working together to build a chimney, given Tom's past? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: Tom finally has a home / family / something he built / a sense of permanence; the chimney is symbolic of belonging; the brightness reflects Tom's growing hope

TEACHER NOTES:
This is a strong moment for inferential thinking. Tom went from no home, to a prison hulk, to the colony, and now is helping BUILD his home. The chimney symbolises stability and shared work with the family he has found.

WATCH FOR:
- Students who connect to Tom's earlier homelessness -- great inference
- Students who only describe the building -- prompt: "Why is it special that TOM is building this?"

[Literacy: Pause Point | VTLM 2.0: Comprehension / Inference]`;

const NOTES_PAUSE2 = `SAY:
- Stop. The convicts came to steal from the Sergeant's garden. Tom tried to confront them. Rob arrived just in time
- Rob says: "We had better see what harm's been done before Da comes home"
- Turn to your partner: What have we learned about Rob and Tom from this situation? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: Rob is loyal / brave / protective of Tom; Tom is no longer alone -- he has someone in his corner; their friendship is now deep enough to face danger together; Rob acts like a brother

TEACHER NOTES:
This moment cements the friendship. Rob does not just defend the garden -- he protects Tom. Notice the language Rob uses: "Da" (his father), "we", "before Da comes home" -- assuming Tom is part of "we", part of the household.

WATCH FOR:
- Students who notice "we" and "Da" -- celebrate the close-reading work
- Students who summarise events -- prompt: "What does this show about their friendship?"

[Literacy: Pause Point | VTLM 2.0: Comprehension / Character]`;

const NOTES_PAUSE3 = `SAY:
- Stop. The boys have decided what to do with the food they have left after the raid. Rob says: "We'll have the rest in the morning"
- Turn to your partner: What is the big idea here? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: hunger is becoming a real problem; the boys are acting like adults / making careful choices; the colony is still struggling to feed itself; even after attack and theft, life continues

TEACHER NOTES:
The "big idea" question is open. Strong responses connect this small choice to the bigger picture: the colony is fragile, food is precious, the boys are stepping up. Some students will see the everyday quality of survival.

WATCH FOR:
- Students who connect to colony-wide food shortages -- excellent contextual reading
- Students who read it only as "they will eat tomorrow" -- prompt: "Why does the author show us this small choice?"

[Literacy: Pause Point | VTLM 2.0: Comprehension / Big Idea]`;

const NOTES_VOCAB = `SAY:
- One vocabulary word from Chapter 35: wielded
- Wielded means held or used, especially a weapon or a tool with skill or force
- The convicts wielded sticks when they attacked Tom
- Say it with me: wielded
- Use it: "The Sergeant wielded his hoe with steady, careful strokes"

DO:
- Display the word
- Choral read
- Quick partner activity: "Tell your partner one thing a person might wield" (a tool, a sword, a pen). 30 seconds

TEACHER NOTES:
"Wielded" is one of the explicit vocabulary words. The other ("scrutinised") is left as incidental for this lesson. Keep this brief.

WATCH FOR:
- Students who give creative examples -- evidence of grasping the meaning
- Students who think wielded = held (lightly) -- nudge: "It is held with skill or force"

[Literacy: Vocabulary | VTLM 2.0: Vocabulary Development]`;

const NOTES_REVISE = `SAY:
- Quick revision. A relative clause is a chunk of words that adds extra information about a noun
- It begins with a relative pronoun: who, whom, which, whose, that. Or a relative adverb: when, where, why
- It sits right after the noun it describes
- Two types:
- Essential (no commas) -- tells the reader WHICH ONE: "The chimney that Tom and Rob built rose from the roof"
- Non-essential (with commas) -- gives EXTRA information: "Rob, who arrived home just in time, drove the men away"
- Test: if you can remove the clause and the sentence still identifies the same noun, it is non-essential -> use commas

DO:
- Display both example sentences
- Read aloud both -- contrast the two
- Show the comma test: "The chimney rose from the roof" -- but WHICH chimney? We need the clause -> essential -> no commas
- Show: "Rob drove the men away" -- still clear who Rob is -> the clause is extra -> commas

TEACHER NOTES:
This revise covers a lot. Keep the two contrasting examples on screen. The comma rule is the trickiest part for upper primary -- the test "can I remove it and still know who/what?" is the key heuristic.

WATCH FOR:
- Students who try to memorise without understanding -- redirect to the comma test
- Students who confuse relative clauses with appositives -- key difference: a relative clause has a verb (who, which, that + verb); an appositive is a noun phrase

[Literacy: Revise | VTLM 2.0: Retention and Recall]`;

const NOTES_IDO = `SAY:
- Watch how I add a relative clause to a sentence about our chapter
- Original: "The men attacked Tom"
- Question: WHICH men? I want to tell the reader. So I add an essential relative clause
- I use "who" because it is people. The clause: "who tried to steal from the garden"
- Where does it go? Right after "men"
- Final: "The men who tried to steal from the garden attacked Tom"
- Notice: no commas. Why? It is essential -- it tells us WHICH men. Without it, the reader does not know who attacked Tom

DO:
- Display original
- Talk through the choice: which one? -> essential -> no commas
- Show the placement (right after "men")
- Read aloud the final sentence
- Then briefly contrast with a non-essential example: "Rob, who arrived home just in time, drove the men away"
- Point: "Here we already know who Rob is. The clause is extra. So we use commas"

TEACHER NOTES:
The I Do explicitly teaches the comma decision through TWO examples -- one essential, one non-essential. This is the trickiest part for students. Use content from Ch 35 (the attack, Rob's return).

WATCH FOR:
- Students drafting their own clauses already
- Students confused by the comma rule -- promise that the CFU and worksheet will give more practice

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check on punctuation. Two sentences. Which one is correctly punctuated?
- A: "Rob who arrived home just in time drove the convicts away"
- B: "Rob, who arrived home just in time, drove the convicts away"
- Hold up A or B
- Three, two, one -- show!

DO:
- Display both sentences
- Show Me Boards
- Scan: most students should choose B
- Cold Call 1-2 students: "Why B?"

TEACHER NOTES:
This CFU directly checks SC3 -- punctuation around non-essential relative clauses. The reasoning: "who arrived home just in time" is non-essential because we already know who Rob is. Therefore commas. Some students will choose A because they think shorter / no commas = simpler -- redirect to the test.

WATCH FOR:
- Students who pick A "because there are fewer commas" -- redirect: "Is the clause essential? No. Then commas"
- Students who can articulate WHY B -- they are ready for the worksheet

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- The correctly punctuated sentence is B
- Why? "Who arrived home just in time" is extra information -- non-essential
- We already know who Rob is. So the clause is set off with commas
- Apply this rule when you write your own sentences

DO:
- Display the reveal
- Read B aloud, pausing slightly at the commas
- Pivot if many missed it: "What does the clause add? Is it telling us WHICH Rob, or just extra detail? Extra detail -> commas"

CFU CHECKPOINT:
Technique: Show Me Boards (A or B)
Script:
- "A or B -- which one is correctly punctuated?"
- Scan for: most students choose B (~80%)
PROCEED (>=80%): Most chose B and can explain. Release to You Do.
PIVOT (<80%): Most likely issue -- students do not see the clause as non-essential. Reteach: "Read the sentence WITHOUT the clause. Does it still tell you who? Yes. Then the clause is extra. Extra means commas." Re-check with one new sentence.

TEACHER NOTES:
The reveal shows the answer with commas highlighted. After reveal, release students to the worksheet.

WATCH FOR:
- Students who immediately want to redo their writing -- celebrate the metacognition

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Time to practise. Use the relative clauses worksheet
- Section 1: identify the relative clause and the relative pronoun (5 sentences)
- Section 2: add a relative clause to answer "Which one?" (essential -- no commas)
- Section 3: add a relative clause to give EXTRA information (non-essential -- use commas)
- Section 4: punctuate the sentences correctly
- 15 minutes. Use the comma test for every non-essential example

DO:
- Distribute the relative clauses worksheet
- Circulate -- prioritise enabling students first
- Quick conferences: "Show me where your relative clause goes. Is it essential or non-essential?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Complete only Section 1 (identify) and Section 2 (essential -- no commas). Skip Sections 3 and 4
- Provide a list of relative pronouns at the top (who, which, that, whose, when, where)
- Extra Notes: These students focus on identification and one type of relative clause

EXTENDING PROMPT:
- Task: After completing all sections, write 2 of their own sentences about Chapters 34-35 -- one with an essential clause, one with a non-essential clause. Annotate the punctuation choice
- Extra Notes: Peer-share the explanations

TEACHER NOTES:
Section 4 is the trickiest -- students must DECIDE whether each clause is essential or non-essential. Active circulation is essential to catch comma errors before they are repeated.

WATCH FOR:
- Students who use commas everywhere (over-using) -- redirect to the test: "Read it without the clause. Still clear?"
- Students who use no commas anywhere (under-using) -- the same redirect
- Students who choose the wrong relative pronoun (e.g. "which" for a person) -- gentle correction: "who" for people, "which/that" for things

[Literacy: You Do | VTLM 2.0: Supported Application / Practice]`;

const NOTES_CLOSING = `SAY:
- Quick self-check. Show me on your fingers
- SC1: I described what we learned about Tom and Rob from Chapters 34 and 35 -- 1 to 5
- SC2: I can identify a relative clause and its relative pronoun -- 1 to 5
- SC3: I can add a relative clause and use commas correctly -- 1 to 5
- One word that describes Tom and Rob's friendship right now -- tell your partner

DO:
- Run each SC
- Note any patterns -- if many score low on SC3, plan a quick comma reteach next session
- Collect worksheets for marking
- Wrap: "Tomorrow: Christmas in the colony, and a serious moment. Get ready"

TEACHER NOTES:
The reflection prompt "one word for the friendship" lets students share without needing complex articulation. This builds confidence and varies the closing.

WATCH FOR:
- Students naming "loyal", "brotherly", "tested", "strong" -- evidence of strong inference

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${RC_RESOURCE.name} is your practice sheet for relative clauses
- The ${ANSWER_KEY_RESOURCE.name} is for teacher use to mark or for whole-class review at the end

DO:
- Print the practice sheet (one per student)
- Print one answer key for teacher use

TEACHER NOTES:
Section 4 is the trickiest part of the worksheet. Plan to walk through one example as a class if many students struggle during circulation.

[Literacy: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Tom Appleby Ch 34-35 + Relative Clauses -- Lesson 17";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Chapters 34-35",
    "Building, Threats + Relative Clauses",
    "Lesson 17  |  Week 4  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  // SLIDE 2 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to question the author's choices in Chapters 34-35 and to identify and add relative (embedded) clauses with correct punctuation",
    ],
    [
      "I can describe what we learn about Tom and Rob from Chapters 34 and 35",
      "I can identify the relative clause and the relative pronoun in a sentence",
      "I can add a relative clause to a sentence and punctuate it correctly",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 3 -- Reading anchor
  contentSlide(
    pres,
    "Read",
    C.PRIMARY,
    "Chapters 34 & 35",
    [
      "Ch 34: Tom and Rob build a chimney for the house",
      "Ch 35: Winter brings disease. Two convicts try to steal from the garden",
      "Three pause points along the way",
      "Note: Ch 35 has a brief section about preparing a dead animal, and a scene where Tom is attacked",
      "Find Chapter 34 in your novel",
    ],
    NOTES_READING,
    FOOTER
  );

  // SLIDE 4 -- Pause point 1
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 34  |  After the chimney is built",
    "...as though the autumn sun was lending it some brightness.",
    "p.183",
    "What is the significance of Rob and Tom working together to build a chimney, given Tom's past?",
    NOTES_PAUSE1,
    FOOTER
  );

  // SLIDE 5 -- Pause point 2
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 35  |  After the attack, Rob arrives",
    "'We'd better see what harm's been done before Da comes home.'",
    "p.191",
    "What have we learned about Rob and Tom from this situation?",
    NOTES_PAUSE2,
    FOOTER
  );

  // SLIDE 6 -- Pause point 3
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 35  |  After the boys decide about food",
    "'We'll have the rest in the morning.'",
    "p.192",
    "What's the big idea?",
    NOTES_PAUSE3,
    FOOTER
  );

  // SLIDE 7 -- Vocabulary: wielded
  vocabSlide(
    pres,
    "wielded",
    "verb",
    "Held or used something (often a weapon or a tool) with skill or force.",
    "The convicts wielded sticks as they attacked Tom near the garden.",
    NOTES_VOCAB,
    FOOTER
  );

  // SLIDE 8 -- Revise: relative clauses
  contentSlide(
    pres,
    "Revise",
    C.SECONDARY,
    "Relative Clauses -- Add Detail with Who, Which, That",
    [
      "A relative clause adds extra information about a noun",
      "Begins with: who, whom, which, whose, that  (or when, where, why)",
      "Sits right after the noun it describes",
      "Essential (no commas):  \"The chimney that Tom and Rob built rose from the roof\"",
      "Non-essential (with commas):  \"Rob, who arrived home just in time, drove the men away\"",
      "Test: remove the clause -- does the sentence still name the same thing? Yes -> commas",
    ],
    NOTES_REVISE,
    FOOTER
  );

  // SLIDE 9 -- I Do: model adding a relative clause
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Add a Relative Clause to a Sentence",
    "Original sentence:\n\n\"The men attacked Tom.\"\n\nWHICH men?\n -> who tried to steal from the garden\n\nWhere does it go?\n -> right after \"men\"\n\nDo I need commas?\n -> NO, it is essential",
    "Final sentence:\n\n\"The men who tried to steal from the garden attacked Tom.\"\n\nNo commas -- the clause tells the reader WHICH men.\n\nContrast (non-essential):\n\n\"Rob, who arrived home just in time, drove the men away.\"\n\nCommas -- we already know who Rob is, so the clause is extra.",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 10 + 11 -- CFU: which is correctly punctuated? (reveal pair)
  function buildCfuBase() {
    const slide = pres.addSlide();
    addTopBar(slide, C.ALERT);
    addBadge(slide, "CFU", { color: C.ALERT });
    addTitle(slide, "Which Is Correctly Punctuated?", { color: C.ALERT });

    const stampW = 1.3;
    slide.addShape("roundRect", {
      x: 9.5 - stampW, y: 0.20, w: stampW, h: 0.32, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
    });
    slide.addText("✓  CHECK", {
      x: 9.5 - stampW, y: 0.20, w: stampW, h: 0.32,
      fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    slide.addShape("roundRect", {
      x: 0.5, y: CONTENT_TOP, w: 2.8, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ALERT },
    });
    slide.addText("Show Me Boards: A or B", {
      x: 0.5, y: CONTENT_TOP, w: 2.8, h: 0.40,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const cardY = CONTENT_TOP + 0.55;
    const cardH = 1.45;
    addCard(slide, 0.5, cardY, 9, cardH, { strip: C.MUTED, fill: C.WHITE });
    slide.addText("A", {
      x: 0.7, y: cardY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    slide.addText("\"Rob who arrived home just in time drove the convicts away.\"", {
      x: 1.2, y: cardY + 0.18, w: 8.0, h: cardH - 0.30,
      fontSize: 17, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const cardBY = cardY + cardH + 0.18;
    const cardBH = SAFE_BOTTOM - cardBY - 0.55;
    addCard(slide, 0.5, cardBY, 9, cardBH, { strip: C.SECONDARY, fill: C.BG_CARD });
    slide.addText("B", {
      x: 0.7, y: cardBY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    slide.addText("\"Rob, who arrived home just in time, drove the convicts away.\"", {
      x: 1.2, y: cardBY + 0.14, w: 8.0, h: cardBH - 0.24,
      fontSize: 17, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(slide, FOOTER);
    slide.addNotes(NOTES_CFU_BUILD);
    return slide;
  }

  withReveal(
    buildCfuBase,
    (slide) => {
      const revealY = 4.90;
      addCard(slide, 0.5, revealY, 9, 0.42, { fill: C.SUCCESS });
      slide.addText("Correct: B  --  non-essential clause needs commas", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 12 -- You Do: worksheet
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Practise Relative Clauses");

    addCard(s, 0.5, CONTENT_TOP, 9, 2.10, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Use the Relative Clauses Practice worksheet", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("Section 1:  Identify the clause and the relative pronoun (5 sentences)\nSection 2:  Add a clause to answer \"Which one?\" (essential, no commas)\nSection 3:  Add a clause that gives EXTRA information (non-essential, commas)\nSection 4:  Punctuate the sentences correctly", {
      x: 0.75, y: CONTENT_TOP + 0.46, w: 8.4, h: 1.45,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 2.25;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Time:  15 minutes", {
      x: 0.75, y: tipY + 0.10, w: 4, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Use the comma test:  remove the clause -- does the sentence still name the same noun?  Yes -> commas\n- Use \"who\" for people, \"which/that\" for things\n- The clause sits right after the noun it describes", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: 0.85,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 13 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Tell your partner one word that describes Tom and Rob's friendship right now.",
      scItems: [
        "I can describe what we learn about Tom and Rob from Chapters 34 and 35",
        "I can identify the relative clause and the relative pronoun in a sentence",
        "I can add a relative clause to a sentence and punctuate it correctly",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (got it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // SLIDE 14 -- Resources
  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // ---- PDF: Relative Clauses Practice -------------------------------------
  const ws = createPdf({ title: RC_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Relative Clauses Practice", {
    color: C.PRIMARY,
    subtitle: "Sentences from Chapters 34-35 of Tom Appleby",
    lessonInfo: "Lesson 17 | Week 4 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "Relative pronouns: who, whom, which, whose, that. Relative adverbs: when, where, why. Use commas around non-essential clauses.", wsY, { color: C.PRIMARY });

  wsY = addSectionHeading(ws, "Section 1: Identify the Clause and Pronoun", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Underline the relative clause. Circle the relative pronoun.", wsY, { fontSize: 10, italic: true });
  wsY += 4;
  wsY = addBodyText(ws, "1.  The chimney that Tom and Rob built rose proudly from the roof.", wsY);
  wsY = addBodyText(ws, "2.  Rob, who arrived home just in time, drove the convicts away.", wsY);
  wsY = addBodyText(ws, "3.  The men who tried to steal from the garden attacked Tom.", wsY);
  wsY = addBodyText(ws, "4.  The Sergeant, whose garden had been raided, watched the boys carefully.", wsY);
  wsY = addBodyText(ws, "5.  The cottage, where the boys lived together, faced the harbour.", wsY);
  wsY += 8;

  wsY = addSectionHeading(ws, "Section 2: Add a Clause to Answer \"Which One?\"  (essential, no commas)", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Add a relative clause that tells the reader WHICH ONE. No commas.", wsY, { fontSize: 10, italic: true });
  wsY += 4;
  wsY = addBodyText(ws, "1.  The chimney __________________________________________ rose proudly from the roof.", wsY);
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 18 });
  wsY = addBodyText(ws, "2.  The garden __________________________________________ slowly came back to life.", wsY);
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 18 });
  wsY = addBodyText(ws, "3.  The fish __________________________________________ filled the pot with steam.", wsY);
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 18 });
  wsY += 6;

  wsY = addSectionHeading(ws, "Section 3: Add a Clause with EXTRA Information  (non-essential, commas)", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "Add a relative clause that gives EXTRA detail. Use commas.", wsY, { fontSize: 10, italic: true });
  wsY += 4;
  wsY = addBodyText(ws, "1.  Tom ___________________________________________ stood his ground at the garden gate.", wsY);
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 18 });
  wsY = addBodyText(ws, "2.  Rob ___________________________________________ shouted at the men to leave.", wsY);
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 18 });
  wsY += 6;

  wsY = addSectionHeading(ws, "Section 4: Punctuate the Sentences", wsY, { color: C.ALERT });
  wsY = addBodyText(ws, "Decide if the clause is essential or non-essential. Add commas only where needed. Rewrite each.", wsY, { fontSize: 10, italic: true });
  wsY += 4;
  wsY = addBodyText(ws, "1.  The chimney that the boys built kept the cottage warm.", wsY);
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 18 });
  wsY = addBodyText(ws, "2.  The Sergeant who had been gone all day returned with a fish.", wsY);
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 18 });
  wsY = addBodyText(ws, "3.  The men who attacked Tom ran into the woods.", wsY);
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 18 });

  addPdfFooter(ws, "Lesson 17 | Relative Clauses Practice");

  // ---- PDF: Answer Key ----------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Relative Clauses Practice -- Answer Key", {
    color: C.SECONDARY,
    subtitle: "Teacher Reference",
    lessonInfo: "Lesson 17 | Week 4 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Sections 2 and 3 have many possible answers. Accept any clause with the correct relative pronoun, correct placement and correct comma use.", akY, { color: C.SECONDARY });

  akY = addSectionHeading(ak, "Section 1", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "1.  Clause: that Tom and Rob built  |  Pronoun: that", akY);
  akY = addBodyText(ak, "2.  Clause: who arrived home just in time  |  Pronoun: who", akY);
  akY = addBodyText(ak, "3.  Clause: who tried to steal from the garden  |  Pronoun: who", akY);
  akY = addBodyText(ak, "4.  Clause: whose garden had been raided  |  Pronoun: whose", akY);
  akY = addBodyText(ak, "5.  Clause: where the boys lived together  |  Adverb: where", akY);
  akY += 8;

  akY = addSectionHeading(ak, "Section 2 (sample answers -- essential, no commas)", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "1.  The chimney that Tom and Rob had built rose proudly from the roof.", akY);
  akY = addBodyText(ak, "2.  The garden that the Sergeant had planted slowly came back to life.", akY);
  akY = addBodyText(ak, "3.  The fish that the Sergeant brought home filled the pot with steam.", akY);
  akY += 8;

  akY = addSectionHeading(ak, "Section 3 (sample answers -- non-essential, commas)", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "1.  Tom, who had no family of his own, stood his ground at the garden gate.", akY);
  akY = addBodyText(ak, "2.  Rob, whose loyalty to Tom was now clear, shouted at the men to leave.", akY);
  akY += 8;

  akY = addSectionHeading(ak, "Section 4 (correctly punctuated)", akY, { color: C.ALERT });
  akY = addBodyText(ak, "1.  The chimney that the boys built kept the cottage warm.   (essential -- no commas)", akY);
  akY = addBodyText(ak, "2.  The Sergeant, who had been gone all day, returned with a fish.   (non-essential -- commas)", akY);
  akY = addBodyText(ak, "3.  The men who attacked Tom ran into the woods.   (essential -- no commas)", akY);

  addPdfFooter(ak, "Lesson 17 | Relative Clauses Answer Key -- TEACHER USE");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Lesson17.pptx` }),
    writePdf(ws, RC_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Lesson17.pptx`);
  console.log("Done: " + RC_RESOURCE.name);
  console.log("Done: " + ANSWER_KEY_RESOURCE.name);
}

build().catch(console.error);
