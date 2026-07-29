"use strict";

// Tom Unit -- Lesson 21: Chapters 42-43 -- Eora Woman & Second Christmas + Subordinating Conjunctions
// Week 5, Lesson 21, Grade 5/6 Literacy
// Reading: Ch 42 (boys meet an Eora woman who shows them food), Ch 43 (second Christmas, food scarcity)
// Sentence-level: Subordinating conjunctions -- complete dependent clauses with independent clauses
// Sensitivity: Ch 42 contains some brief adult content -- flagged in notes

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

const SESSION_NUMBER = 21;
const FOOTER = "Chapters 42-43 | Lesson 21 | Week 5 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson21_Eora_Woman_Subordinating_Conjunctions";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const SC_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Subordinating Conjunctions Practice",
  "Student worksheet: complete dependent clauses with independent clauses using content from Chapters 42-43."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Subordinating Conjunctions Practice Answer Key",
  "Teacher reference: model completions for the practice worksheet."
);
const RESOURCE_ITEMS = [SC_RESOURCE, ANSWER_KEY_RESOURCE];
const SC_PDF_PATH = path.join(OUT_DIR, SC_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Lesson 21. Two chapters today: 42 and 43
- The boys go fishing and meet an Eora woman who shows them how to find food. Then it is the second Christmas at the cove
- After reading we return to sentence work -- subordinating conjunctions

DO:
- Display title slide as students settle
- Have copies of the novel ready
- Have the subordinating conjunctions worksheet ready (do not distribute yet)

TEACHER NOTES:
This lesson reconnects reading with sentence-level skill. Chapter 42 is a high-warmth chapter where the Eora woman shares food knowledge with the boys -- a deliberate counterpoint to earlier chapters about isolation and scarcity. The novel uses period-typical wording ("Indian woman"); in your own framing, prefer "Eora woman" when discussing the chapter while quoting the novel exactly.

SENSITIVITY ADVISORY:
- What it is: Chapter 42 contains some brief adult content -- the woman is shown crying earlier in the chapter, and there is a brief reference that may need framing.
- Framing language: "There is a short moment in Chapter 42 that hints at something sad in this woman's past. We will read it without dwelling. Let me know if you need a moment."
- Watch for: students who may ask follow-up questions about what the text implies
- Protocol: Read the passage at pace. If a student asks a follow-up, answer briefly and respectfully then move on.

WATCH FOR:
- Students unsure about how the woman is described -- pre-frame the difference between the novel's period language and respectful modern language (Eora)

[Literacy: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${SC_RESOURCE.name} is your practice worksheet
- The ${ANSWER_KEY_RESOURCE.name} is for teacher reference

DO:
- Print the practice worksheet (one per student)
- Print one answer key for teacher use
- Distribute the practice after the I Do / We Do

TEACHER NOTES:
The worksheet is the You Do task. The answer key gives sample completions for each dependent clause and lists valid alternatives so the teacher can score generously.

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands again: reading the chapters, then writing with subordinating conjunctions
- Read the success criteria

DO:
- Choral read the LI, then the SCs
- Brief: "Today you will finish off some half-sentences. The start of each one is the dependent clause -- you finish it with the independent clause that makes it whole"

TEACHER NOTES:
SC1 targets character and authorial-choice analysis. SC2 targets the structural understanding -- dependent clause + independent clause = complex sentence. SC3 targets the application using content from the chapters.

WATCH FOR:
- Students who think "dependent" means weak or unimportant -- redirect: "A dependent clause cannot stand alone. It needs an independent clause to complete it"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_READING = `SAY:
- Two chapters today. I will read aloud
- Three pause points
- A note before we start: Chapter 42 has a moment that hints at something sad in the Eora woman's past. We read it without dwelling. Let me know if you need to step out

DO:
- 30 seconds for students to find Chapter 42
- Read aloud at a steady pace
- Take pause point 1 at p.220, pause point 2 at p.228, pause point 3 at p.234
- Then read Chapter 43

TEACHER NOTES:
The reading takes 15-20 minutes. Both chapters carry quiet emotional weight. Chapter 42 is mostly warm -- the Eora woman teaches the boys to find food. Chapter 43 returns to the colony's isolation and the Sergeant's loneliness.

WATCH FOR:
- Students reacting to the woman's grief -- pause and acknowledge
- Students who connect the woman's knowledge to what we read in Chapters 40-41 -- celebrate the connection

[Literacy: Read Aloud | VTLM 2.0: Build Knowledge / Reading]`;

const NOTES_PAUSE1 = `SAY:
- Stop. The Eora woman is crying. The narrator wonders: "Had they died, while she survived?"
- Turn to your partner: how is the author making you FEEL about the woman right now? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: empathy, sorrow, awareness of loss; she is a survivor; the author wants us to see her as a person, not a stranger; the question links back to the deaths in Chapters 40-41

TEACHER NOTES:
This is one of the most important authorial moves in this part of the novel. The narrator's question forces the reader to humanise her. Strong students will connect "they died while she survived" to the disease deaths in Chapters 40-41 -- the smallpox outbreak. Celebrate that connection.

WATCH FOR:
- Students who only describe the action (she is crying) -- prompt: "How does the AUTHOR want us to feel?"
- Students who connect this scene to the disease chapter -- excellent literary thinking

[Literacy: Pause Point | VTLM 2.0: Comprehension / Author Craft]`;

const NOTES_PAUSE2 = `SAY:
- Stop. The narrator tells us: "Tom knew the answer to both questions"
- Turn to your partner: what do you think the author wants us to know here? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: Tom has grown -- he understands now; he can read the woman's situation without needing it explained; the two boys are different now; Tom is no longer the same boy who arrived

TEACHER NOTES:
This is a character-development moment. Tom's silent understanding shows his growth. Earlier in the novel Tom would not have grasped the nuance; here he does. The author is showing us, not telling us, that Tom has changed.

WATCH FOR:
- Students who repeat what Tom knows but do not name the change -- prompt: "Would the OLD Tom have known this?"
- Students who name Tom's growth -- celebrate the inference

[Literacy: Pause Point | VTLM 2.0: Comprehension / Character Development]`;

const NOTES_PAUSE3 = `SAY:
- Stop. The Sergeant says about news from home: "And that's something that no fishing net or garden can give us"
- Turn to your partner: what have we learned from this conversation? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: the colony is still isolated; the Sergeant feels lonely; food is not the only thing people need; connection and news matter; the Sergeant is more honest with the boys now

TEACHER NOTES:
The conversation is about more than supplies. The Sergeant is naming a different kind of hunger -- for connection and home. Strong students will pick up on the emotional layer. The juxtaposition with the warm Chapter 42 (the woman feeds them) is intentional.

WATCH FOR:
- Students who only mention supplies -- prompt: "What is the Sergeant REALLY saying he is missing?"
- Students who name the emotional dimension -- excellent

[Literacy: Pause Point | VTLM 2.0: Comprehension / Big Idea]`;

const NOTES_REVISE = `SAY:
- Quick revision before we practise
- A dependent clause cannot stand on its own. It needs an independent clause to make a whole sentence
- A subordinating conjunction is the word that joins the two clauses
- Examples of subordinating conjunctions: even though, since, after, because, when, while, although, if
- Today we have THREE dependent clauses about Chapter 42. You finish each one

DO:
- Display the revise card
- Choral read the three example subordinating conjunctions
- Quick partner check: "Tell your partner one subordinating conjunction"

TEACHER NOTES:
Students may have met simple/compound/complex sentences in earlier years. This is the revision layer for complex sentences. Keep the vocabulary tight -- dependent, independent, subordinating conjunction. Avoid introducing more terminology today.

WATCH FOR:
- Students who think "subordinating" means "less important" -- redirect: "It means it depends on the other clause"
- Students who add a coordinating conjunction (and, but) by mistake -- redirect: "Subordinating conjunctions are different -- they make one clause depend on the other"

[Literacy: Revise | VTLM 2.0: Retention and Recall]`;

const NOTES_IDO = `SAY:
- Watch me. I have a dependent clause: "Even though the Indian woman was crying,"
- The comma signals the dependent clause is finished -- now I need to add the independent clause
- My finish: "she still wanted to interact with the boys"
- Whole sentence: "Even though the Indian woman was crying, she still wanted to interact with the boys"
- Notice: the independent clause stands on its own. The two clauses together make a complex sentence

DO:
- Write the original dependent clause on the board
- Add the independent clause underneath
- Read the whole sentence aloud
- Highlight: "Even though" is the subordinating conjunction. The comma sits after the dependent clause

TEACHER NOTES:
The example uses content from Chapter 42 -- the Eora woman's interaction with the boys. Model thinking aloud about WHY the chosen ending fits the chapter content. Students who simply finish the sentence without reference to the text are missing the link.

WATCH FOR:
- Students who finish with content that contradicts the chapter (e.g. "she stopped talking to them") -- redirect to chapter content
- Students who write a fragment -- redirect: "Read it back. Does the second part stand on its own?"

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Together. Mini-whiteboards out
- Dependent clause: "Since the Indian woman knew how to find food,"
- Add the independent clause that finishes the sentence. Use what you know from Chapter 42
- 60 seconds. Best you can

DO:
- 60 seconds silent work on whiteboards
- Boards up
- Pick 2-3 strong examples to read aloud
- Reveal the model ending: "she was able to teach the boys some new skills"
- Quick discussion: "What other endings would work?"

TEACHER NOTES:
This is guided practice. Most students should manage a workable ending. Common alternative endings: "she shared her knowledge with them", "the boys returned home well-fed", "they were no longer hungry". All are acceptable if they form a complete independent clause and reflect Chapter 42 content.

WATCH FOR:
- Students whose ending does not stand alone -- redirect: "Read your ending without the first part. Does it make sense by itself?"
- Students who echo the dependent clause -- redirect: "Add NEW information. What HAPPENED because of that?"

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Which option below is a CORRECT independent clause to finish this dependent clause?
- "After the boys encounter the Indian woman,"
- A: "and they were happy"
- B: "they wonder if her life has changed now that they are on her land"
- Hold up A or B. Three, two, one -- show

DO:
- Display both
- Show Me Boards
- Scan: most students should choose B
- Cold Call 1-2 students: "Why B?"

TEACHER NOTES:
A starts with "and" -- it is a fragment. It cannot stand alone. B is a full independent clause -- it has a subject (they), a verb (wonder) and makes sense on its own. The structural rule is the key, not just the content.

WATCH FOR:
- Students who pick A because it "sounds right" -- redirect: "Read A by itself. Does 'and they were happy' make sense on its own?"
- Students who pick B and explain -- they have the structural rule

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Correct independent clause: B
- A starts with "and" -- that makes it a fragment, not an independent clause. It cannot stand alone
- B is a full sentence on its own: "they wonder if her life has changed now that they are on her land"
- Whole complex sentence: "After the boys encounter the Indian woman, they wonder if her life has changed now that they are on her land"

DO:
- Display the reveal banner
- Read the whole complex sentence aloud
- Pivot if many picked A: "Read A by itself. Does it stand alone? No -- and that is the rule"

CFU CHECKPOINT:
Technique: Show Me Boards (A or B)
Script:
- "Hold up A or B. Which is a correct independent clause to finish the sentence?"
- Scan for: most students choose B (>=80%)
PROCEED (>=80%): Most chose B. Release to the worksheet.
PIVOT (<80%): Most likely issue -- students think any phrase after a comma works. Reteach: "Independent clause = a full sentence on its own. Read it alone. If it does not make sense alone, it is not independent."

TEACHER NOTES:
After reveal, release students to the worksheet. The structural rule (independent = stands alone) is the takeaway.

WATCH FOR:
- Students who self-correct toward B -- they have the rule

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Your turn. The worksheet has six dependent clauses for you to complete
- All use content from Chapters 42 and 43
- For each one: read the dependent clause, then add a full independent clause that makes a complex sentence
- Make sure your ending stands alone if you read it by itself
- 12 minutes. I will circulate

DO:
- Distribute the practice worksheet
- Circulate -- prioritise students who looked unsure during the CFU
- Quick conferences: "Read me your sentence. Does the second part stand alone?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the enabling version of the worksheet which provides two possible endings for each dependent clause -- students choose the better one and explain why
- Extra Notes: This focuses cognitive load on the structural rule (does it stand alone?) rather than on generation

EXTENDING PROMPT:
- Task: After completing the worksheet, write THREE of your own dependent clauses using subordinating conjunctions (although, while, because, when) and a different independent clause for each one. Underline the subordinating conjunction
- Extra Notes: Use content from any chapter we have read

TEACHER NOTES:
The 12-minute block is independent practice. The worksheet design (six items, content from the chapters) makes the link between reading and writing explicit. Active circulation catches structural errors (fragments) before they become habits.

WATCH FOR:
- Students whose endings echo the dependent clause -- redirect to NEW information
- Students who use a comma splice (no subordinating conjunction) -- redirect to the conjunction list
- Students who write fragments -- redirect: "Read it back without the first part"

[Literacy: You Do | VTLM 2.0: Independent Application]`;

const NOTES_CLOSING = `SAY:
- Quick self-check. Show me on your fingers
- SC1: I described what we learned from Chapters 42-43 -- 1 to 5
- SC2: I understand what a dependent clause and an independent clause are -- 1 to 5
- SC3: I completed dependent clauses with independent clauses -- 1 to 5
- Then turn to your partner: tell them ONE subordinating conjunction we used today

DO:
- Run each SC
- Listen in on partner shares -- this tells you who has the terminology
- Wrap up: "These complex sentences will sit nicely inside our writing -- you can use them in your information report"

TEACHER NOTES:
The closing also acts as a quick survey. If most students name only "even though" or only "since", flag the others for next-day warm-up.

WATCH FOR:
- Students who name multiple subordinating conjunctions -- evidence of retention
- Students unsure -- check in privately

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Tom Appleby Ch 42-43 + Subordinating Conjunctions -- Lesson 21";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Chapters 42-43",
    "Eora Woman, Second Christmas + Subordinating Conjunctions",
    "Lesson 21  |  Week 5  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources (immediately after title)
  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // SLIDE 3 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to analyse Chapters 42-43 and to complete complex sentences by adding an independent clause to a dependent clause using subordinating conjunctions",
    ],
    [
      "I can describe what we learned from Chapters 42 and 43",
      "I can identify a dependent clause, an independent clause and a subordinating conjunction",
      "I can complete a dependent clause with an independent clause to make a complex sentence",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 4 -- Reading anchor
  contentSlide(
    pres,
    "Read",
    C.PRIMARY,
    "Chapters 42 & 43",
    [
      "Ch 42: The boys go fishing and meet an Eora woman who shares food knowledge",
      "Ch 43: A second Christmas at the cove -- food is scarce, the Sergeant is lonely",
      "Three pause points along the way",
      "Note: Ch 42 has a brief moment that hints at something sad -- we read it respectfully",
      "Find Chapter 42 in your novel",
    ],
    NOTES_READING,
    FOOTER
  );

  // SLIDE 5 -- Pause point 1 (p.220)
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 42  |  The Eora woman is crying",
    "Had they died, while she survived?",
    "p.220",
    "How is the author making you feel right now about the woman?",
    NOTES_PAUSE1,
    FOOTER
  );

  // SLIDE 6 -- Pause point 2 (p.228)
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 42  |  After the boys return",
    "...Tom knew the answer to both questions.",
    "p.228",
    "What do you think the author wants us to know?",
    NOTES_PAUSE2,
    FOOTER
  );

  // SLIDE 7 -- Pause point 3 (p.234)
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 43  |  The Sergeant speaks",
    "And that's something that no fishing net or garden can give us.",
    "p.234",
    "What have we learned from this conversation?",
    NOTES_PAUSE3,
    FOOTER
  );

  // SLIDE 8 -- Revise
  contentSlide(
    pres,
    "Revise",
    C.SECONDARY,
    "Dependent + Independent = Complex Sentence",
    [
      "Dependent clause:  cannot stand alone -- needs an independent clause to be a sentence",
      "Independent clause:  stands alone as a full sentence",
      "Subordinating conjunction:  the joining word -- even though, since, after, because, when, while, although, if",
      "When the dependent clause comes first, use a comma after it",
    ],
    NOTES_REVISE,
    FOOTER
  );

  // SLIDE 9 -- I Do: model
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Complete the Sentence",
    "Dependent clause:\n\n\"Even though the Indian woman was crying,\"\n\nThe subordinating conjunction is:\n\"Even though\"\n\nThe comma signals:\nthe dependent clause is finished",
    "My independent clause:\n\n\"she still wanted to interact with the boys\"\n\nWhole complex sentence:\n\n\"Even though the Indian woman was crying, she still wanted to interact with the boys.\"\n\nNotice:\n- The second part stands alone\n- Together the two clauses make a complex sentence",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 10 -- We Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Together -- Whiteboards", { color: C.SECONDARY });

    addCard(s, 0.5, CONTENT_TOP, 9, 1.30, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("Finish this complex sentence:", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.32,
      fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText("\"Since the Indian woman knew how to find food,  ____________________________________ .\"", {
      x: 0.75, y: CONTENT_TOP + 0.46, w: 8.4, h: 0.74,
      fontSize: 18, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
      fit: "shrink", shrinkText: true, valign: "middle",
    });

    const tipY = CONTENT_TOP + 1.50;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Your job:", {
      x: 0.75, y: tipY + 0.10, w: 4, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Add the INDEPENDENT clause -- a full sentence on its own\n- Use content from Chapter 42\n- 60 seconds. Best you can. Boards up when called", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: 1.10,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
  }

  // SLIDE 11 + 12 -- CFU: which is an independent clause? (reveal pair)
  function buildCfuBase() {
    const slide = pres.addSlide();
    addTopBar(slide, C.ALERT);
    addBadge(slide, "CFU", { color: C.ALERT });
    addTitle(slide, "Which One Stands Alone?", { color: C.ALERT });

    const stampW = 1.3;
    slide.addShape("roundRect", {
      x: 9.5 - stampW, y: 0.20, w: stampW, h: 0.32, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
    });
    slide.addText("CHECK", {
      x: 9.5 - stampW, y: 0.20, w: stampW, h: 0.32,
      fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    slide.addShape("roundRect", {
      x: 0.5, y: CONTENT_TOP, w: 4.5, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ALERT },
    });
    slide.addText("Show Me Boards: A or B", {
      x: 0.5, y: CONTENT_TOP, w: 4.5, h: 0.40,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Prompt row
    slide.addText("Dependent clause:  \"After the boys encounter the Indian woman,\"", {
      x: 0.5, y: CONTENT_TOP + 0.55, w: 9, h: 0.40,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0, valign: "middle",
    });

    const cardY = CONTENT_TOP + 1.05;
    const cardH = 1.10;
    addCard(slide, 0.5, cardY, 9, cardH, { strip: C.MUTED, fill: C.WHITE });
    slide.addText("A", {
      x: 0.7, y: cardY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    slide.addText("\"and they were happy\"", {
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
    slide.addText("\"they wonder if her life has changed now that they are on her land\"", {
      x: 1.2, y: cardBY + 0.14, w: 8.0, h: cardBH - 0.24,
      fontSize: 15, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(slide, FOOTER);
    slide.addNotes(NOTES_CFU_BUILD);
    return slide;
  }

  withReveal(
    buildCfuBase,
    (slide) => {
      const revealY = 4.68;
      addCard(slide, 0.5, revealY, 9, 0.42, { fill: C.SUCCESS });
      slide.addText("Correct independent clause: B  --  it stands alone", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 13 -- You Do: worksheet
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Complete the Complex Sentences");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.85, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Six dependent clauses from Chapters 42-43", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:    Read the dependent clause aloud quietly\nNext:     Add a full independent clause that makes a complex sentence\nThen:     Check -- does your ending stand alone if you read it by itself?\nFinally:  Use content from Chapters 42-43 in your endings", {
      x: 0.75, y: CONTENT_TOP + 0.46, w: 8.4, h: 1.30,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 2.00;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Time: 12 minutes", {
      x: 0.75, y: tipY + 0.10, w: 4, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Read your ending aloud without the dependent clause -- does it stand alone?\n- Avoid starting your ending with 'and' or 'but' (those make fragments)\n- Use details from the chapters, not invented content", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: 0.85,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 14 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Tell your partner one subordinating conjunction we used today.",
      scItems: [
        "I can describe what we learned from Chapters 42 and 43",
        "I can identify a dependent clause, an independent clause and a subordinating conjunction",
        "I can complete a dependent clause with an independent clause to make a complex sentence",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (got it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: Subordinating Conjunctions Practice ---------------------------
  const ws = createPdf({ title: SC_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Subordinating Conjunctions Practice", {
    color: C.PRIMARY,
    subtitle: "Complete the complex sentences using content from Chapters 42-43",
    lessonInfo: "Lesson 21 | Week 5 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "Each dependent clause below cannot stand alone. Add an independent clause to make a complex sentence. Use content from Chapters 42-43. Read your ending aloud by itself -- if it stands alone as a sentence, it is an independent clause.", wsY, { color: C.PRIMARY });

  wsY = addSectionHeading(ws, "Subordinating Conjunctions Bank", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "even though     |     since     |     after     |     because     |     when     |     while     |     although     |     if", wsY, { fontSize: 11 });
  wsY += 6;

  wsY = addSectionHeading(ws, "1.  Even though the Indian woman was crying,", wsY, { color: C.PRIMARY });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;

  wsY = addSectionHeading(ws, "2.  Since the Indian woman knew how to find food,", wsY, { color: C.PRIMARY });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;

  wsY = addSectionHeading(ws, "3.  After the boys encounter the Indian woman,", wsY, { color: C.PRIMARY });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;

  wsY = addSectionHeading(ws, "4.  Because the food rations were scarce,", wsY, { color: C.PRIMARY });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;

  wsY = addSectionHeading(ws, "5.  When the Sergeant talked about news from home,", wsY, { color: C.PRIMARY });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;

  wsY = addSectionHeading(ws, "6.  Although it was a second Christmas at the cove,", wsY, { color: C.PRIMARY });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 6;

  wsY = addSectionHeading(ws, "Challenge", wsY, { color: C.ALERT });
  wsY = addBodyText(ws, "Write your own complex sentence using a subordinating conjunction and content from any chapter we have read. Underline the subordinating conjunction.", wsY, { fontSize: 11, italic: true });
  wsY = addLinedArea(ws, wsY, 3, { lineSpacing: 22 });

  addPdfFooter(ws, "Lesson 21 | Subordinating Conjunctions Practice");

  // ---- PDF: Answer Key ----------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Subordinating Conjunctions Practice -- Answer Key", {
    color: C.SECONDARY,
    subtitle: "Teacher reference -- accept any independent clause that stands alone and reflects chapter content",
    lessonInfo: "Lesson 21 | Week 5 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Multiple endings are valid. The structural rule is: the ending must stand alone as a full sentence (independent clause). The content rule is: the ending must reflect Chapters 42-43.", akY, { color: C.SECONDARY });

  akY = addSectionHeading(ak, "1. Even though the Indian woman was crying,", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sample: she still wanted to interact with the boys.", akY, { fontSize: 11 });
  akY = addBodyText(ak, "Alt: she offered them oysters / she kept showing them where to find food.", akY, { fontSize: 10, italic: true });
  akY += 4;

  akY = addSectionHeading(ak, "2. Since the Indian woman knew how to find food,", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sample: she was able to teach the boys some new skills.", akY, { fontSize: 11 });
  akY = addBodyText(ak, "Alt: the boys returned home with food / she shared her knowledge with them.", akY, { fontSize: 10, italic: true });
  akY += 4;

  akY = addSectionHeading(ak, "3. After the boys encounter the Indian woman,", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sample: they wonder if her life has changed now that they are on her land.", akY, { fontSize: 11 });
  akY = addBodyText(ak, "Alt: they brought birds back for the Sergeant / they understood more about the people who lived here first.", akY, { fontSize: 10, italic: true });
  akY += 4;

  akY = addSectionHeading(ak, "4. Because the food rations were scarce,", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sample: the colony struggled to feed everyone.", akY, { fontSize: 11 });
  akY = addBodyText(ak, "Alt: people were always thinking about what they could eat / the Sergeant rationed every meal carefully.", akY, { fontSize: 10, italic: true });
  akY += 4;

  akY = addSectionHeading(ak, "5. When the Sergeant talked about news from home,", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sample: the boys could hear that he missed England.", akY, { fontSize: 11 });
  akY = addBodyText(ak, "Alt: he looked sad and longed for the supply ship / he reminded everyone how isolated the colony felt.", akY, { fontSize: 10, italic: true });
  akY += 4;

  akY = addSectionHeading(ak, "6. Although it was a second Christmas at the cove,", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sample: the boys still received new clothes from the Sergeant.", akY, { fontSize: 11 });
  akY = addBodyText(ak, "Alt: the meal was smaller than last year / the colony tried to celebrate despite the lack of food.", akY, { fontSize: 10, italic: true });
  akY += 6;

  akY = addSectionHeading(ak, "Common Errors to Watch For", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- Endings that start with 'and' or 'but' (fragments). Reteach: independent clause must stand alone.", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Endings that echo the dependent clause (no new information). Prompt for NEW content.", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Endings that contradict chapter content. Redirect to the novel.", akY, { fontSize: 10 });

  addPdfFooter(ak, "Lesson 21 | Subordinating Conjunctions Practice Answer Key -- TEACHER USE");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Lesson21.pptx` }),
    writePdf(ws, SC_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Lesson21.pptx`);
  console.log("Done: " + SC_RESOURCE.name);
  console.log("Done: " + ANSWER_KEY_RESOURCE.name);
}

build().catch(console.error);
