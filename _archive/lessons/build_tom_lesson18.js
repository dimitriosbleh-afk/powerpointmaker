"use strict";

// Tom Unit -- Lesson 18: Chapters 36-39 -- Christmas, Hanging + Sentence Expansion
// Week 4, Lesson 18, Grade 5/6 Literacy
// Reading: Ch 36 (resentment, Indians), Ch 37 (Christmas), Ch 38 (hanging), Ch 39 (Thomas's birthday)
// Sentence-level: Sentence Expansion using adverbials and clauses
// Sensitivity: Ch 38 contains detailed content about hanging -- flagged in notes

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

const SESSION_NUMBER = 18;
const FOOTER = "Chapters 36-39 | Lesson 18 | Week 4 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson18_Christmas_Hanging_Sentence_Expansion";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const SE_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Sentence Expansion Practice",
  "Student worksheet: expand kernel sentences from Chapters 36-39 using who, what doing, where, when and why detail."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Sentence Expansion Practice Answer Key",
  "Teacher reference: model expansions for the practice worksheet."
);
const RESOURCE_ITEMS = [SE_RESOURCE, ANSWER_KEY_RESOURCE];
const SE_PDF_PATH = path.join(OUT_DIR, SE_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Lesson 18. Four chapters today: 36, 37, 38 and 39
- Christmas comes to the cove. Then a serious moment: seven marines are hanged for stealing food
- After reading, we will practise sentence expansion -- adding detail to short sentences

DO:
- Display title slide as students settle
- Have copies of the novel ready
- Have the sentence expansion worksheet ready

TEACHER NOTES:
This is a content-dense lesson. Four chapters is a lot of reading -- you may need to read briskly or split chapters across two sessions. Chapter 38 contains detailed content about hanging. Flag this for sensitive students before reading. Consider summarising the hanging description briefly rather than reading it all aloud.

SENSITIVITY ADVISORY:
- What it is: Chapter 38 describes a public hanging of seven marines. It is not gratuitous but it is detailed.
- Framing language: "Today the chapter has a serious part. Seven marines are punished by hanging for stealing food. Let me know if you need to step out. We will keep it respectful."
- Watch for: students who are visibly upset
- Protocol: Pause if needed. You may summarise rather than read aloud. Debrief gently.

WATCH FOR:
- Students reacting to the hanging -- normalise: "Tom is shocked too. The author wants us to feel the weight of this"

[Literacy: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands today: reading the chapters and analysing the author's choices, then expanding sentences with detail
- Read the success criteria

DO:
- Choral read the LI, then the SCs
- Brief: "Sentence expansion means starting with a short sentence and adding detail -- who, where, when, why -- so the reader can picture more"

TEACHER NOTES:
SC1 targets character analysis through pause points. SC2 targets understanding of adverbials and expansion. SC3 is the application -- students expand sentences using content from the chapters.

WATCH FOR:
- Students who think expansion = adding adjectives only -- redirect: "Adverbials answer when, where, why, how"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_READING = `SAY:
- Four chapters today. I will read aloud
- Three pause points -- we stop and think
- A note before we start: Chapter 38 has a serious moment -- seven marines are hanged for stealing food. Let me know if you need to step out
- Find Chapter 36 in your novel

DO:
- 30 seconds for students to find the chapter
- Read at a steady pace
- Consider summarising the hanging description in Ch 38 rather than reading every detail
- Take pause points at the listed pages

TEACHER NOTES:
A 4-chapter reading is at the upper limit of what fits in 60 minutes alongside sentence work. Watch the clock. If short on time, read Ch 36-37 today and Ch 38-39 with another short session, OR summarise Ch 39 (Thomas's birthday) since it returns to the modern frame story rather than advancing Tom's plot. The third pause point (p.208) is the strongest big-idea moment -- protect time for it.

WATCH FOR:
- Students visibly upset by Ch 38 -- pause and acknowledge
- Students who lose track of who Thomas (modern frame) is vs Tom -- briefly orient

[Literacy: Read Aloud | VTLM 2.0: Build Knowledge / Reading]`;

const NOTES_PAUSE1 = `SAY:
- Stop. We just heard Tom feeling some resentment. He thinks Rob has more than he does -- a father, freedom -- and "there was no point dreaming of any different"
- Turn to your partner: How does what we just read add to our understanding of Tom? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: Tom is human / has complicated feelings; even after settling in, he carries old wounds; he envies Rob; he is realistic / accepting; he suppresses his own dreams

TEACHER NOTES:
This is a moment of psychological complexity. Tom has a home and a family, but he still feels the loss of his own father. Strong students will see that contentment and grief can co-exist.

WATCH FOR:
- Students who notice Tom's mixed feelings -- celebrate the inference
- Students who say Tom is ungrateful -- redirect: "Is grief about freedom or family ungratefulness?"

[Literacy: Pause Point | VTLM 2.0: Comprehension / Inference]`;

const NOTES_PAUSE2 = `SAY:
- Stop. Seven marines have been hanged. The Sergeant says, "Come on. Let us go home."
- Turn to your partner: What is going on here? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: the Sergeant is exhausted / shaken / wants to protect the boys; the public hanging is meant to scare people but it has shaken everyone; the colony is harsh and food is precious enough to die for; the boys have witnessed something that will mark them

TEACHER NOTES:
This is the moment after the hanging. The brief, weary line "Let us go home" carries weight. Strong students will notice the Sergeant's protective instinct -- he wants to take the boys away from the public spectacle.

WATCH FOR:
- Students who only describe the hanging -- prompt: "Why does the Sergeant just say 'Let us go home'?"
- Students who connect to the food shortage -- excellent contextual reading

[Literacy: Pause Point | VTLM 2.0: Comprehension / Author Craft]`;

const NOTES_PAUSE3 = `SAY:
- Stop. The Sergeant has just warned the boys that food stores are running low. He says: "Aye, and the hens and sheep as well"
- Turn to your partner: What is the author saying here? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: the colony is in real trouble; the hens and sheep -- which are family animals -- might have to be eaten; even the most precious things are at risk; survival is hard; this links to why the marines stole food

TEACHER NOTES:
This pause connects everything. The marines stole food. They were hanged. Now the Sergeant is telling the boys their own livestock might have to be killed for food. The cycle of scarcity is brutal and clear. Some students will see the author is showing the colony has limits.

WATCH FOR:
- Students who connect to Ch 38 (the hangings) -- excellent thinking
- Students who only mention the animals -- prompt: "What does this tell us about the colony?"

[Literacy: Pause Point | VTLM 2.0: Comprehension / Big Idea]`;

const NOTES_VOCAB = `SAY:
- One vocabulary word: festive
- Festive means cheerful, celebratory -- the way you feel at a party or a special meal
- Chapter 39 describes Thomas's birthday in the modern story -- a festive meal with food crammed on the table
- Say it with me: festive
- Use it: "The cottage felt festive on Christmas morning, with small gifts and the smell of cooking"

DO:
- Display the word
- Choral read
- Quick partner activity: "Tell your partner one festive event you have been to" (30 seconds)

TEACHER NOTES:
"Festive" is one of the explicit vocabulary words. Notice the contrast: Ch 37 (Christmas at the cove -- modest and tender) vs Ch 39 (Thomas's birthday -- lavish, modern). The word "festive" connects both, with very different meanings of plenty.

WATCH FOR:
- Students who give simple examples (birthday party, Christmas) -- accept these
- Students who notice the contrast -- celebrate the cross-chapter thinking

[Literacy: Vocabulary | VTLM 2.0: Vocabulary Development]`;

const NOTES_REVISE = `SAY:
- Quick revision. We expand sentences to give the reader more information
- Start with a short kernel sentence: "The Sergeant returned"
- Add detail by answering questions: who, what doing, where, when, why, how
- These extra details are called adverbials -- they tell us when, where or how
- Position matters. Adverbials can go at the start, middle or end
- If the adverbial is at the start (e.g. "On a cold afternoon,...") use a comma after it

DO:
- Display the kernel sentence
- Display the expansion questions on screen
- Show one quick example: "On a cold afternoon, the Sergeant returned wearily to the cottage with rations for the week"
- Point to each detail: "On a cold afternoon" (when) | "wearily" (how) | "to the cottage" (where) | "with rations for the week" (why)

TEACHER NOTES:
This revision combines the concept of sentence expansion with the formal language of adverbials. Keep the explanation tight -- 2-3 minutes maximum. The trick is the comma rule for sentence-initial adverbials.

WATCH FOR:
- Students who add only adjectives (e.g. "the tired Sergeant") -- this is fine but not the focus today
- Students who place adverbials in the wrong position -- gentle redirect

[Literacy: Revise | VTLM 2.0: Retention and Recall]`;

const NOTES_IDO = `SAY:
- Watch how I expand a kernel sentence
- Original: "The boys dug"
- Who? -- the boys (already there)
- What doing? -- dug
- Where? -- in the garden
- When? -- on a sunny morning
- Why? -- because they were planting seeds
- I will start the sentence with the WHEN detail. That means I add a comma after it
- Final: "On a sunny morning, the boys dug in the garden because they were planting seeds"
- Notice: I did not need to use ALL the questions. Pick the ones that add the most useful detail

DO:
- Display the kernel
- Build the expansion live -- one detail at a time
- Show the comma after the sentence-initial adverbial
- Read aloud the final sentence
- Brief comparison: "Original tells us almost nothing. Expanded gives the reader a clear picture"

TEACHER NOTES:
The I Do uses content from earlier chapters (digging holes for crops, Ch 31-33) to keep the focus on the current grammar skill rather than the heavy Ch 38 content. The key teaching points: pick the most useful details, watch the comma rule for sentence-initial adverbials.

WATCH FOR:
- Students who try to use every question -- redirect: "Which detail adds the most picture?"
- Students who notice the comma -- celebrate the close attention

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Two expanded versions of the same kernel. Which one is more effective?
- Kernel: "Tom watched"
- A: "Tom watched the men"
- B: "From the back of the crowd, Tom watched silently as the marines climbed the gallows"
- Hold up A or B
- Three, two, one -- show!

DO:
- Display both
- Show Me Boards
- Scan: most students should choose B
- Cold Call 1-2 students: "Why B?"

TEACHER NOTES:
B uses three different expansions: where (from the back of the crowd), how (silently), what (as the marines climbed the gallows). Together they paint a clear picture. A only adds an object. Students who pick A may be thinking simpler = clearer -- redirect to picture-painting.

WATCH FOR:
- Students who pick A -- redirect to "Which one helps you picture the scene?"
- Students who articulate which expansions B uses -- they are ready

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- B is more effective
- B uses three expansions: WHERE (from the back of the crowd), HOW (silently), WHAT WAS HAPPENING (as the marines climbed the gallows)
- Together, you can SEE the moment
- Use this in your own expansion -- pick a few different details, not just one

DO:
- Display the reveal
- Annotate the three expansions verbally
- Pivot if many missed it: "If I just say 'Tom watched the men,' what do you SEE? Almost nothing. Now read B and tell me what you see"

CFU CHECKPOINT:
Technique: Show Me Boards (A or B)
Script:
- "Hold up A or B. Which expansion is more effective?"
- Scan for: most students choose B (~80%)
PROCEED (>=80%): Most chose B and can name an expansion. Release to You Do.
PIVOT (<80%): Most likely issue -- students think simpler = better. Reteach: read both aloud. "Which one paints a picture?" Re-check: ask students to name ONE detail from B.

TEACHER NOTES:
After reveal, release students to the worksheet. The three-expansion principle is the takeaway.

WATCH FOR:
- Students who immediately want to revisit their drafts -- celebrate

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Time to expand. Use the Sentence Expansion worksheet
- Section 1: identify the expansions in a sentence
- Section 2: expand a kernel using prompts (who, where, when)
- Section 3: expand a kernel WITHOUT prompts -- you choose
- For the LAST section: start your sentence with a "when" detail. Use a comma after it
- 15 minutes. Read your expansions aloud -- do they paint a picture?

DO:
- Distribute the worksheet
- Circulate -- prioritise enabling students first
- Quick conferences: "Which expansion did you choose? Why?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Complete only Sections 1 and 2 (with prompts). Skip Section 3
- Provide expansion stems on the worksheet (e.g. "On... morning,", "in the...", "because...")
- Extra Notes: These students focus on placement and one type of expansion at a time

EXTENDING PROMPT:
- Task: After completing all sections, write 2 of their own expanded sentences about Chapters 36-39 -- one starting with a "when" detail (use a comma), one starting with the noun
- Extra Notes: Peer-share and identify each other's expansions

TEACHER NOTES:
The 15-minute block splits roughly: Section 1 (5 min), Section 2 (5 min), Section 3 (5 min). Active circulation is the formative assessment. The comma rule is the trickiest skill -- check during conferencing.

WATCH FOR:
- Students who put a comma everywhere -- redirect to the sentence-initial rule
- Students who only add one detail -- prompt: "What else could you add? Where? When?"
- Students writing complete picture-painting expansions -- celebrate

[Literacy: You Do | VTLM 2.0: Supported Application / Practice]`;

const NOTES_CLOSING = `SAY:
- Quick self-check. Show me on your fingers
- SC1: I described what we learned about Tom from Chapters 36-39 -- 1 to 5
- SC2: I can identify expansions in a sentence -- 1 to 5
- SC3: I can expand a kernel sentence using detail -- 1 to 5
- One word about Chapter 38 -- tell your partner

DO:
- Run each SC
- Note any patterns -- if many score low on SC3, plan a quick reteach next session
- Collect worksheets for marking
- Wrap: "Tomorrow we read the final chapters of this section -- and start planning our third body paragraph for the information report"

TEACHER NOTES:
The "one word about Chapter 38" prompt invites brief, manageable reflection without forcing students to articulate complex feelings about the hanging. Many will say "sad", "hard", "shocking", "unfair" -- accept all.

WATCH FOR:
- Students naming powerful single words -- evidence of emotional engagement with the text

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${SE_RESOURCE.name} is your practice sheet
- The ${ANSWER_KEY_RESOURCE.name} is for teacher use to mark or for whole-class review

DO:
- Print the practice sheet (one per student)
- Print one answer key for teacher use

TEACHER NOTES:
Many sections have multiple valid answers. The answer key shows samples; accept any expansion that adds genuine detail and uses correct comma placement.

[Literacy: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Tom Appleby Ch 36-39 + Sentence Expansion -- Lesson 18";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Chapters 36-39",
    "Christmas, Hanging + Sentence Expansion",
    "Lesson 18  |  Week 4  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // SLIDE 2 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to question the author's choices in Chapters 36-39 and to expand kernel sentences with adverbial detail using correct punctuation",
    ],
    [
      "I can describe what we learn about Tom from the events of Chapters 36 to 39",
      "I can identify the expansions (who, what doing, where, when, why) in a sentence",
      "I can expand a kernel sentence with detail and use commas correctly",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 3 -- Reading anchor
  contentSlide(
    pres,
    "Read",
    C.PRIMARY,
    "Chapters 36, 37, 38 & 39",
    [
      "Ch 36: Tom feels resentment toward Rob. Indians attack the brickwork",
      "Ch 37: Christmas at the cove. Sergeant gives gifts and remembers his late wife",
      "Ch 38: Seven marines are hanged for stealing food. The Sergeant warns about food stores",
      "Ch 39: Thomas's birthday in the modern story -- a lavish meal",
      "Note: Ch 38 contains detailed content about hanging",
    ],
    NOTES_READING,
    FOOTER
  );

  // SLIDE 4 -- Pause point 1 (p.195)
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 36  |  Tom thinks about Rob",
    "And there was no point dreaming of any different.",
    "p.195",
    "How does what we just read add to our understanding of Tom?",
    NOTES_PAUSE1,
    FOOTER
  );

  // SLIDE 5 -- Pause point 2 (p.198)
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 38  |  After the hanging",
    "'Come on,' he said wearily. 'Let's go home.'",
    "p.198",
    "What's going on here?",
    NOTES_PAUSE2,
    FOOTER
  );

  // SLIDE 6 -- Pause point 3 (p.208)
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 38  |  The Sergeant warns about food",
    "...Aye, and the hens and sheep as well.",
    "p.208",
    "What is the author saying here?",
    NOTES_PAUSE3,
    FOOTER
  );

  // SLIDE 7 -- Vocabulary: festive
  vocabSlide(
    pres,
    "festive",
    "adjective",
    "Cheerful and celebratory -- the way a place feels at a party or a special meal.",
    "The cottage felt festive on Christmas morning, with small gifts and the smell of cooking.",
    NOTES_VOCAB,
    FOOTER
  );

  // SLIDE 8 -- Revise: sentence expansion
  contentSlide(
    pres,
    "Revise",
    C.SECONDARY,
    "Sentence Expansion -- Add Detail with Adverbials",
    [
      "Start with a kernel: a short subject + verb sentence (e.g. \"The Sergeant returned\")",
      "Expand by answering: who, what doing, where, when, why, how",
      "Adverbials tell us when, where or how -- they can go at the start, middle or end",
      "Sentence-initial adverbial?  Add a comma after it",
      "Example:  \"On a cold afternoon, the Sergeant returned wearily to the cottage with rations\"",
    ],
    NOTES_REVISE,
    FOOTER
  );

  // SLIDE 9 -- I Do: model expanding a sentence
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Expand a Kernel Sentence",
    "Kernel:\n\n\"The boys dug.\"\n\nQuestions to ask:\n  -> Who?  the boys\n  -> What doing?  dug\n  -> Where?  in the garden\n  -> When?  on a sunny morning\n  -> Why?  because they were planting seeds\n\nI will start with the WHEN detail.\n -> use a comma after it",
    "Final sentence:\n\n\"On a sunny morning, the boys dug in the garden because they were planting seeds.\"\n\nNotice the comma after the sentence-initial adverbial.\n\nI did not need to use every question. I picked the details that paint the clearest picture.",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 10 + 11 -- CFU: which expansion is more effective? (reveal pair)
  function buildCfuBase() {
    const slide = pres.addSlide();
    addTopBar(slide, C.ALERT);
    addBadge(slide, "CFU", { color: C.ALERT });
    addTitle(slide, "Which Expansion Is More Effective?", { color: C.ALERT });

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

    // Kernel banner
    const kY = CONTENT_TOP + 0.55;
    addCard(slide, 0.5, kY, 9, 0.45, { fill: C.SECONDARY });
    slide.addText("Kernel:  \"Tom watched.\"", {
      x: 0.5, y: kY, w: 9, h: 0.45,
      fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    // Sentence A card
    const cardY = kY + 0.55;
    const cardH = 1.05;
    addCard(slide, 0.5, cardY, 9, cardH, { strip: C.MUTED, fill: C.WHITE });
    slide.addText("A", {
      x: 0.7, y: cardY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    slide.addText("\"Tom watched the men.\"", {
      x: 1.2, y: cardY + 0.18, w: 8.0, h: cardH - 0.30,
      fontSize: 17, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Sentence B card
    const cardBY = cardY + cardH + 0.18;
    const cardBH = SAFE_BOTTOM - cardBY - 0.55;
    addCard(slide, 0.5, cardBY, 9, cardBH, { strip: C.SECONDARY, fill: C.BG_CARD });
    slide.addText("B", {
      x: 0.7, y: cardBY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    slide.addText("\"From the back of the crowd, Tom watched silently as the marines climbed the gallows.\"", {
      x: 1.2, y: cardBY + 0.14, w: 8.0, h: cardBH - 0.24,
      fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
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
      slide.addText("More effective: B  --  uses where, how and what was happening", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 15, fontFace: FONT_B, color: C.WHITE, bold: true,
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
    addTitle(s, "Practise Sentence Expansion");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.85, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Use the Sentence Expansion Practice worksheet", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("Section 1:  Identify the expansions in each sentence (5 sentences)\nSection 2:  Expand each kernel using prompts (who, where, when)  (3 kernels)\nSection 3:  Expand a kernel WITHOUT prompts -- you choose  (2 kernels)\nFinish: write one sentence that begins with a \"when\" detail. Use a comma.", {
      x: 0.75, y: CONTENT_TOP + 0.46, w: 8.4, h: 1.30,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 2.00;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Time:  15 minutes", {
      x: 0.75, y: tipY + 0.10, w: 4, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Read your expansion aloud -- does it paint a picture?\n- Pick the details that add the most. You do not need every question\n- Sentence-initial adverbial?  Add a comma after it", {
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
      reflectionPrompt: "Tell your partner one word that describes Chapter 38.",
      scItems: [
        "I can describe what we learn about Tom from Chapters 36 to 39",
        "I can identify the expansions in a sentence",
        "I can expand a kernel sentence with detail and use commas correctly",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (got it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // SLIDE 14 -- Resources


  // ---- PDF: Sentence Expansion Practice -----------------------------------
  const ws = createPdf({ title: SE_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Sentence Expansion Practice", {
    color: C.PRIMARY,
    subtitle: "Sentences from Chapters 36-39 of Tom Appleby",
    lessonInfo: "Lesson 18 | Week 4 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "Expand kernel sentences by answering: who, what doing, where, when, why, how. Use a comma after a sentence-initial adverbial.", wsY, { color: C.PRIMARY });

  wsY = addSectionHeading(ws, "Section 1: Identify the Expansions", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Underline the expansions (where / when / how / why) in each sentence.", wsY, { fontSize: 10, italic: true });
  wsY += 4;
  wsY = addBodyText(ws, "1.  On Christmas morning, the Sergeant carved the meat slowly at the cottage table.", wsY);
  wsY = addBodyText(ws, "2.  After the hanging, Tom walked silently down the hill behind the Sergeant.", wsY);
  wsY = addBodyText(ws, "3.  In the garden, Rob planted the seeds carefully because the soil was thin.", wsY);
  wsY = addBodyText(ws, "4.  When winter arrived, the colony struggled to find fresh food.", wsY);
  wsY = addBodyText(ws, "5.  At the brickwork, the men worked from dawn until dusk to finish the wall.", wsY);
  wsY += 8;

  wsY = addSectionHeading(ws, "Section 2: Expand Using Prompts", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Expand each kernel by answering the prompts. Watch your commas.", wsY, { fontSize: 10, italic: true });
  wsY += 4;
  wsY = addBodyText(ws, "1.  Kernel: \"The Sergeant returned.\"   |   Prompts: when?  where?  why?", wsY);
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 20 });
  wsY = addBodyText(ws, "2.  Kernel: \"The boys watched.\"   |   Prompts: where?  how?  what was happening?", wsY);
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 20 });
  wsY = addBodyText(ws, "3.  Kernel: \"Rob smiled.\"   |   Prompts: when?  why?  who at?", wsY);
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 20 });
  wsY += 6;

  wsY = addSectionHeading(ws, "Section 3: Expand Without Prompts", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "Expand each kernel using your own choices. Pick the details that paint the clearest picture.", wsY, { fontSize: 10, italic: true });
  wsY += 4;
  wsY = addBodyText(ws, "1.  Kernel: \"The marines climbed.\"", wsY);
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 20 });
  wsY = addBodyText(ws, "2.  Kernel: \"The cottage was quiet.\"", wsY);
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 20 });
  wsY += 6;

  wsY = addSectionHeading(ws, "Finish: Sentence-Initial Adverbial", wsY, { color: C.ALERT });
  wsY = addBodyText(ws, "Write one sentence about Chapters 36-39 that BEGINS with a \"when\" detail. Use a comma after it.", wsY, { fontSize: 10, italic: true });
  wsY += 4;
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });

  addPdfFooter(ws, "Lesson 18 | Sentence Expansion Practice");

  // ---- PDF: Answer Key ----------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Sentence Expansion Practice -- Answer Key", {
    color: C.SECONDARY,
    subtitle: "Teacher Reference",
    lessonInfo: "Lesson 18 | Week 4 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Sections 2 and 3 have many possible answers. Accept any expansion that adds genuine detail and uses commas correctly.", akY, { color: C.SECONDARY });

  akY = addSectionHeading(ak, "Section 1 (expansions to underline)", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "1.  WHEN: On Christmas morning  |  HOW: slowly  |  WHERE: at the cottage table", akY);
  akY = addBodyText(ak, "2.  WHEN: After the hanging  |  HOW: silently  |  WHERE: down the hill  |  WHO: behind the Sergeant", akY);
  akY = addBodyText(ak, "3.  WHERE: In the garden  |  HOW: carefully  |  WHY: because the soil was thin", akY);
  akY = addBodyText(ak, "4.  WHEN: When winter arrived  |  WHY: to find fresh food", akY);
  akY = addBodyText(ak, "5.  WHERE: At the brickwork  |  WHEN: from dawn until dusk  |  WHY: to finish the wall", akY);
  akY += 8;

  akY = addSectionHeading(ak, "Section 2 (sample answers)", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "1.  Late in the afternoon, the Sergeant returned to the cottage with the week's rations.", akY);
  akY = addBodyText(ak, "2.  From the back of the crowd, the boys watched in silence as the marines climbed the gallows.", akY);
  akY = addBodyText(ak, "3.  On Christmas morning, Rob smiled at his father over the carved meat because the gift had been kind.", akY);
  akY += 8;

  akY = addSectionHeading(ak, "Section 3 (sample answers)", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "1.  In the heavy summer heat, the marines climbed the gallows step by reluctant step.", akY);
  akY = addBodyText(ak, "2.  After the boys had gone to sleep, the cottage was quiet under a thin moon.", akY);
  akY += 8;

  akY = addSectionHeading(ak, "Finish (sample)", akY, { color: C.ALERT });
  akY = addBodyText(ak, "On a cold December morning, Tom and Rob walked back to the cottage in heavy silence.", akY);

  addPdfFooter(ak, "Lesson 18 | Sentence Expansion Answer Key -- TEACHER USE");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Lesson18.pptx` }),
    writePdf(ws, SE_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Lesson18.pptx`);
  console.log("Done: " + SE_RESOURCE.name);
  console.log("Done: " + ANSWER_KEY_RESOURCE.name);
}

build().catch(console.error);
