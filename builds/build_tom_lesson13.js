"use strict";

// Tom Unit -- Lesson 13: Chapters 27-28 + Note Taking (KPAS)
// Week 3, Lesson 13, Year 5/6 Literacy
// Reading: Chapters 27-28 (Tom finally disembarks, swamp work, first encounters)
// Sensitivity: Chapter 28 contains content about death and violence; first encounters
// with Aboriginal and Torres Strait Islander people are referenced via the historical
// term "Indians", which is from the source text and is discussed as an authorial choice.
// Writing: Note taking -- keywords, phrases, abbreviations and symbols (KPAS).

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
  quoteSlide, modellingSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 13;
const FOOTER = "Chapters 27-28 | Lesson 13 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson13_New_Land";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const PRACTICE_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "KPAS Practice",
  "Student worksheet: convert sentences from Chapters 27-28 into keywords, phrases, abbreviations and symbols."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "KPAS Answer Key",
  "Teacher reference: model KPAS conversions for the practice sentences."
);
const RESOURCE_ITEMS = [PRACTICE_RESOURCE, ANSWER_KEY_RESOURCE];
const PRACTICE_PDF_PATH = path.join(OUT_DIR, PRACTICE_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Two chapters today: 27 and 28
- Tom finally steps off the ship and starts work in the new land
- After reading we go back to KPAS -- our note-taking shorthand. Today we add three new symbols

DO:
- Display title slide as students settle
- Have novels and exercise books on desks

SENSITIVITY ADVISORY:
- What it is: Chapter 28 includes death and violence at the swamp. The chapters refer to Aboriginal and Torres Strait Islander peoples using the historical term "Indians", which is the wording in the novel
- Framing language: "The author uses the word 'Indians' the way English settlers wrote at the time. We are going to talk about why she made that choice and what it shows us about the settlers, not about Aboriginal and Torres Strait Islander peoples. Today they are called the First Peoples or Aboriginal and Torres Strait Islander peoples"
- Watch for: Aboriginal and Torres Strait Islander students in the room -- check before the lesson, give them the opportunity to step out, and invite their voice without putting anyone on the spot
- Protocol: keep the discussion focused on the author's choice and the historical record; do not invite role-play or speculation about the violence

TEACHER NOTES:
Lesson 13. Chapter 27 is the disembarking and the new colony. Chapter 28 is the swamp work and first violent encounter. The novel uses "Indians" because that is what English settlers wrote at the time -- this is a real historical record of perspective, not an endorsement. Discuss the author's choice carefully and respectfully.

WATCH FOR:
- Students with Aboriginal and Torres Strait Islander heritage -- check in before; offer choice
- Students who challenge the term -- excellent; this is the discussion the author wants

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands today. First, what the author shows us about the settlers and the new land. Second, using KPAS to take quick, useful notes
- Read each success criterion together

DO:
- Choral read the LI, then each SC
- Brief reminder: KPAS = Keywords, Phrases, Abbreviations, Symbols -- a way to capture meaning without writing every word

TEACHER NOTES:
SC1 targets close reading and authorial choice. SC2 targets vocabulary application. SC3 targets the KPAS skill -- specifically converting full sentences into shorthand notes.

WATCH FOR:
- Students unsure about KPAS -- they will see it modelled
- Students who used KPAS in Lesson 8 -- they can support peers

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_READING = `SAY:
- Reading mode: teacher read aloud
- Two chapters. Listen for two things: how does Tom feel about the new land, and how do the settlers see the people who already live here?
- I will pause twice for us to think together
- After the read, we work on KPAS

DO:
- Read Chapter 27 and Chapter 28 aloud
- Plan two pause points: p.134 (Tom's first impression of the land) and p.136 (the use of "Indians")
- Brief contextual note: Chapter 28 includes a violent moment at the swamp. Read it carefully, then pause for breath before continuing

TEACHER NOTES:
The two pause points target two different aspects of authorial perspective. Pause 1 is Tom's interior view of the new land. Pause 2 is the lens of the colonial narrator and how the author chooses to present that lens.

WATCH FOR:
- Students who flinch at the violence -- acknowledge briefly, move on respectfully
- Students who notice the difference between Tom's view and the marines' view -- excellent

[General: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause here. Tom is finally on land after eleven days of waiting offshore
- "...the thought came to him: this land is beautiful"
- Why does Tom need the courage cloak metaphor here? [The land is unfamiliar. Beautiful, but strange. He has been through a hard journey and is now in a place he has never seen]
- What does the author want us to know? [Tom can see beauty even after suffering. He is curious about the new land. He is not the broken boy from the start of the novel]

DO:
- Display the quote
- Think-Pair-Share: 30 seconds think, 30 seconds pair
- Connect back: "Remember the courage cloak from Chapter 22? What is it doing here?"

TEACHER NOTES:
This pause point develops the courage cloak motif and shows Tom's resilience. The author is using the same image across chapters to track Tom's growth.

WATCH FOR:
- Students who connect to the courage cloak earlier -- excellent motif tracking
- Students who notice "beautiful" as a contrast to the journey -- strong reading

[General: Pause Point 1 | VTLM 2.0: Higher-Order Questioning]`;

const NOTES_PAUSE2 = `SAY:
- Pause here. The marines find a canoe and weapons that belong to the local people
- The author calls them "Indians". Who are the people the author is referring to? [Aboriginal and Torres Strait Islander peoples -- the First Peoples of this land. The settlers used the word "Indians" because they thought of the lands they sailed to as the East Indies and West Indies. It is a historical word, not the right word now]
- Why might the author use "Indians" instead of "Aboriginal people"? [She is showing us how the settlers thought, what words they used. The author wants us to see the world through the settlers' eyes -- not because they were right, but because that perspective is part of the history]
- Who are the "Indians" really? [The First Peoples of this land, who had been here for over 60,000 years. They had names for themselves, languages and Country -- the settlers did not learn or use those names]

DO:
- Display the quote
- Cold Call after thinking time
- Push for the difference between the author's choice (showing the colonial perspective) and the modern reader's awareness (the people had their own names)

CFU CHECKPOINT:
Technique: Cold Call (after thinking time)

Script:
- "Why might the author use 'Indians' instead of 'Aboriginal people'? Who is she letting us see through?"
- Cold call 2-3 students after 30 seconds of think time
- Scan for: she is showing the settlers' perspective; "Indians" is the colonial word; modern readers know the people had their own names and languages

PROCEED (>=80%): Most students recognise that the author is using the settlers' word to show us their perspective. Continue.
PIVOT (<80%): Most likely issue -- students think the author is calling the people the wrong name on purpose to be unkind. Reteach: "The author is putting us inside the settlers' heads. The settlers thought 'Indians' because they had no other word. The author is telling us what it was really like, not what should have been said. We can see the gap -- between what the settlers saw and who the people really were." Re-check: "What does the author want us to notice when she uses that word?"

TEACHER NOTES:
This is a delicate but important pause point. The author's choice is a deliberate window onto the settler perspective; the modern reader sees the gap and the harm. Frame the discussion as analysis of authorial choice, not as judgment of any group of people.

WATCH FOR:
- Students who name First Nations groups they know (e.g., Eora, Gadigal) -- celebrate this knowledge; it is the right knowledge to bring to this moment
- Students troubled by the term -- their feelings are appropriate. Affirm: "It is not the right word. The author is showing us what the settlers said"

[General: Pause Point 2 | VTLM 2.0: Critical Analysis]`;

const NOTES_VOCAB = `SAY:
- Three explicit words from today's chapters
- Revelry -- loud, joyful celebration. The marines' revelry on shore included singing and shouting
- Taunt -- to tease or insult someone. The convicts taunted each other to keep their spirits up
- Ignorant -- not knowing something, especially something you should know. The captain was ignorant of the people who already lived here

DO:
- Read each word and meaning aloud
- Quick oral routine: "Use ignorant in a sentence about the captain" -- 30 seconds, partner share
- Note pronunciation: REV-el-ree; TAWNT; IG-nor-ant

TEACHER NOTES:
"Ignorant" is a useful word for the discussion of Pause Point 2 -- the captain's ignorance of the First Peoples. Encourage that connection.

WATCH FOR:
- Students who connect "ignorant" to the captain's blindness to the people already on the land -- excellent
- Students who use the words naturally in discussion -- celebrate

[General: Vocabulary | VTLM 2.0: Vocabulary Instruction]`;

const NOTES_KPAS_REVIEW = `SAY:
- Quick revision. KPAS = Keywords, Phrases, Abbreviations, Symbols. A way to take fast notes
- We learned these abbreviations and symbols in Lesson 8
- Remember: when we take notes we leave out little words like "the", "and", "a"
- Today we add three new symbols. Up arrow, down arrow, and forward slash. Watch the next slide

DO:
- Display the existing KPAS reference
- Choral read: "b/c means because; w/ means with; w/o means without; equals means definition; plus means and; arrow means causes or results in"
- Keep this brief -- 90 seconds

TEACHER NOTES:
This is revision, not new teaching. Keep it tight. Students should be able to read these symbols quickly.

WATCH FOR:
- Students who do not recall b/c -- a brief reteach during the I Do
- Students who already know all six -- they can show peers

[General: Review | VTLM 2.0: Retention and Recall]`;

const NOTES_NEW_SYMBOLS = `SAY:
- Three new symbols today
- Up arrow -- means "more" or "an increase". The number of convicts UP ARROW means the number went up
- Down arrow -- means "less" or "a decrease". Sam's strength DOWN ARROW means his strength went down
- Forward slash -- can be a full stop, a comma, or signals a new idea. Use it instead of writing a full stop in your notes

DO:
- Display the three symbols on screen
- Demonstrate each with a quick spoken sentence: "ship anchored / convicts disembarked / new colony begins"
- Note: forward slash speeds up notes; do NOT use it in formal writing

TEACHER NOTES:
The three new symbols extend KPAS for trend (up/down) and structure (slash for breaks). Keep examples short.

WATCH FOR:
- Students who confuse up arrow with the existing right arrow (causes) -- the difference: right arrow is cause-effect, up arrow is amount/trend
- Students who try to use the slash in formal sentences -- redirect: "Slash is for notes, not for paragraphs"

[General: New Symbols | VTLM 2.0: Explicit Vocabulary / Symbol Instruction]`;

const NOTES_IDO = `SAY:
- Watch me convert a sentence into KPAS
- Sentence: "After eleven days waiting offshore, Tom was finally allowed to step off the ship and the marines were excited because they could begin building the new colony."
- First I find the keywords. Who? Tom. What doing? Step off ship. When? After 11 days. Why? Marines wanted to begin building
- Now I convert: "11 days / Tom off ship / marines excited b/c new colony begins"
- Notice: I dropped "after", "was finally allowed to", "the", "they could". I kept the people, the actions and the cause-effect

DO:
- Display the original sentence and the KPAS version side by side
- Highlight the keywords in the original
- Think aloud: "I check -- can I read my notes back and remember the meaning? Yes. That is the test"

TEACHER NOTES:
The I Do shows the full conversion process: identify keywords, drop the small words, use abbreviations and symbols. The check question -- "Can I read my notes and recover the meaning?" -- is the heart of the skill.

WATCH FOR:
- Students who keep too many words -- prompt: "What can you cut?"
- Students who drop too many words -- prompt: "Can you still read this back?"

[General: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU = `SAY:
- Quick check before you write. Look at the sentence on the screen
- "The number of convicts on shore went up as more boats arrived from the ship."
- Which symbol best captures "the number went up"?
- Show me on your fingers: 1 = arrow up, 2 = arrow down, 3 = right arrow

DO:
- Use Finger Voting (1 / 2 / 3)
- Scan for: 1 (up arrow)
- Cold call: "Why up arrow?"

CFU CHECKPOINT:
Technique: Finger Voting (1 / 2 / 3)

Script:
- "The number of convicts on shore went up. Which symbol fits? 1 = up arrow, 2 = down arrow, 3 = right arrow"
- Scan for: mostly 1
- Follow up: "Why up arrow?" [Up arrow shows an increase in amount; right arrow shows cause-effect, which is different]

PROCEED (>=80%): Most show 1 with reasoning. Release to write.
PIVOT (<80%): Most likely issue -- students choose 3 (right arrow) because they remember it as "results in". Reteach: "Right arrow shows cause-effect: A causes B. Up arrow shows the AMOUNT or TREND: more, an increase. The number going up is a TREND, not a cause-effect." Re-check: "Try this: 'Storm right arrow ship damaged'. Does that work? Yes -- cause and effect. Now try: 'Number of convicts up arrow'. Does that work? Yes -- amount."

TEACHER NOTES:
This CFU targets the difference between the two arrow symbols. The trap is over-using right arrow because it is the most familiar.

WATCH FOR:
- Students who choose 3 -- they need to see the cause-effect vs trend distinction
- Students who choose 1 and explain it -- ready to write

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Your turn. On your worksheet, convert the sentences into KPAS
- Use the symbol reference if you need it
- 10 minutes. I will circulate

DO:
- Distribute the KPAS Practice worksheet
- Circulate -- check that students are using the new symbols correctly and that their notes are still readable
- Confer with enabling and extending students

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the worksheet, but with key words highlighted in each sentence and a symbol reference at the top of the page. Students convert the highlighted keywords using the reference
- Extra Notes: These students can refer to the I Do model as a parallel example

EXTENDING PROMPT:
- Task: After completing the practice sentences, choose a paragraph from Chapter 27 or 28 and convert it to KPAS notes (no longer than 5 lines of notes for the paragraph). Then read the notes back and check the meaning is recoverable

TEACHER NOTES:
The 10-minute write is the application of the KPAS work. Check that students are using ALL the symbols, not just one or two.

WATCH FOR:
- Students who write full sentences -- redirect: "Drop the small words"
- Students whose notes are unreadable -- redirect: "Can you read this back? If not, keep the keyword"
- Students using the new symbols accurately -- celebrate

[General: You Do | VTLM 2.0: Independent Application]`;

const NOTES_CLOSING = `SAY:
- Quick self-check against the success criteria. Thumbs for each
- Then turn to a partner: read your KPAS notes for one sentence and see if your partner can rebuild the original meaning

DO:
- Run thumbs check for each SC
- Listen in on partner shares -- this is a quality check on KPAS
- Wrap up: "Tomorrow Tom finds out what happens after the violence at the swamp"

TEACHER NOTES:
The partner check is a great formative assessment for KPAS quality. If a partner cannot recover the meaning, the notes need more keywords.

WATCH FOR:
- Students whose notes are too sparse -- they need more keywords
- Students whose notes are too long -- they need fewer small words

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${PRACTICE_RESOURCE.name} is for converting sentences into KPAS
- The ${ANSWER_KEY_RESOURCE.name} is teacher reference

DO:
- Print the practice worksheet (one per student)
- Print the answer key (teacher copy only)

TEACHER NOTES:
KPAS practice should continue across sessions. Consider revisiting in Lesson 14 when students take notes from a non-fiction article.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Chapters 27-28: New Land -- Lesson 13";

  // Slide 1 -- Title
  titleSlide(
    pres,
    "Chapters 27-28",
    "New Land",
    "Lesson 13  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  // Slide 2 -- LI/SC
  liSlide(
    pres,
    [
      "We are learning to notice the choices an author makes when describing a new land, and to take quick, accurate notes using keywords, phrases, abbreviations and symbols",
    ],
    [
      "I can explain the choices an author makes when describing the settlers and the new land",
      "I can use newly taught vocabulary in my own sentences",
      "I can convert a sentence into keywords, phrases, abbreviations and symbols (KPAS)",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // Slide 3 -- Reading Launch
  contentSlide(
    pres,
    "Teacher Read Aloud",
    C.PRIMARY,
    "Chapters 27 - 28",
    [
      "Reading Mode: Teacher Read Aloud",
      "Chapter 27: the convicts disembark and the new colony begins",
      "Chapter 28: Tom works in the swamp; the settlers meet the people who already live here",
      "Focus: how does the author show us the settlers? How does she show us the land?",
    ],
    NOTES_READING,
    FOOTER
  );

  // Slide 4 -- Pause Point 1
  quoteSlide(
    pres,
    "Pause Point 1",
    "Chapter 27 -- p. 134",
    "...the thought came to him: this land is beautiful.",
    "p. 134",
    "Why does Tom need the courage cloak here? What does the author want us to know about Tom?",
    NOTES_PAUSE1,
    FOOTER
  );

  // Slide 5 -- Pause Point 2 (sensitive)
  quoteSlide(
    pres,
    "Pause Point 2",
    "Chapter 28 -- p. 136",
    "And all the while the village of tents grew.",
    "p. 136",
    "The author calls them \"Indians\". Why does she choose that word? Whose perspective are we seeing?",
    NOTES_PAUSE2,
    FOOTER
  );

  // Slide 6 -- Vocabulary
  contentSlide(
    pres,
    "Vocabulary",
    C.SECONDARY,
    "Three Words from Today's Reading",
    [
      "revelry -- loud, joyful celebration (the marines' revelry on shore)",
      "taunt -- to tease or insult someone (the convicts taunted each other to keep their spirits up)",
      "ignorant -- not knowing something you should know (the captain was ignorant of the people already living here)",
    ],
    NOTES_VOCAB,
    FOOTER
  );

  // Slide 7 -- KPAS revision
  contentSlide(
    pres,
    "KPAS",
    C.PRIMARY,
    "Quick Revision -- Symbols We Already Know",
    [
      "b/c   =   because",
      "w/    =   with",
      "w/o   =   without",
      "=     =   means / definition",
      "+     =   and",
      "->    =   causes, results in",
    ],
    NOTES_KPAS_REVIEW,
    FOOTER
  );

  // Slide 8 -- New symbols
  contentSlide(
    pres,
    "New Symbols",
    C.ACCENT,
    "Three New Symbols Today",
    [
      "Up arrow   =   more, an increase (population up arrow)",
      "Down arrow   =   less, a decrease (food down arrow)",
      "Forward slash /   =   full stop / comma / new idea",
      "Up and down arrows show TREND -- the amount goes up or down",
      "Right arrow (->) shows CAUSE -- one thing leads to another",
    ],
    NOTES_NEW_SYMBOLS,
    FOOTER
  );

  // Slide 9 -- I Do: model conversion
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Convert a Sentence to KPAS",
    "Original sentence:\n\n\"After eleven days\nwaiting offshore, Tom\nwas finally allowed to\nstep off the ship and\nthe marines were\nexcited because they\ncould begin building\nthe new colony.\"",
    "Find the keywords: WHO, WHAT DOING, WHEN, WHY.\n\nKPAS version:\n\n\"11 days / Tom off ship / marines excited b/c new colony begins\"\n\nCheck: can I read this back and recover the meaning? Yes.",
    NOTES_IDO,
    FOOTER
  );

  // Slide 10 + 10b -- CFU with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Which Symbol Fits?",
      "1 = up arrow  |  2 = down arrow  |  3 = right arrow",
      "\"The number of convicts on shore went up as more boats arrived from the ship.\"\n\nWhich symbol best captures \"went up\"?\n\nFingers: 1, 2 or 3?",
      NOTES_CFU,
      FOOTER
    ),
    (s) => {
      addCard(s, 0.5, SAFE_BOTTOM - 0.95, 9, 0.85, { strip: C.SUCCESS, fill: C.WHITE });
      s.addText("1 = UP ARROW   --   shows an increase in amount or trend", {
        x: 0.75, y: SAFE_BOTTOM - 0.88, w: 8.4, h: 0.70,
        fontSize: 18, fontFace: FONT_B, color: C.PRIMARY, bold: true,
        valign: "middle", margin: 0,
      });
      s.addNotes("SAY:\n- Up arrow. The number went UP -- that is a trend or amount\n- Right arrow shows cause: A leads to B\n- Quick check: 'storm right arrow ship damaged' -- that is cause-effect\n- 'Convicts up arrow' -- that is a trend\n\nDO:\n- Reveal the answer\n- Show the difference between trend (up/down arrow) and cause (right arrow)\n\n[General: CFU Reveal | VTLM 2.0: Formative Feedback]");
    }
  );

  // Slide 11 -- You Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Convert Sentences to KPAS");

    addInstructionCard(s, [
      { text: "On your worksheet:", role: "header" },
      { text: "First: Read each sentence carefully" },
      { text: "Next: Find the keywords (who, what doing, when, why)" },
      { text: "Then: Convert using abbreviations and symbols" },
      { text: "Last: Check -- can you read your notes back?" },
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
      { text: "b/c, w/, w/o, =, +, ->  (already know)", options: { breakLine: true } },
      { text: "up arrow = more  |  down arrow = less  |  / = full stop or new idea (NEW)" },
    ], {
      x: 0.75, y: tipY + 0.46, w: 8.4, h: 0.85,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, bullet: true, margin: 0,
      paraSpaceAfter: 2,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // Slide 12 -- Closing
  closingSlide(
    pres,
    "Read your KPAS notes for one sentence to a partner. Can your partner rebuild the original meaning?",
    [
      "I can explain the choices an author makes when describing the settlers and the new land",
      "I can use newly taught vocabulary in my own sentences",
      "I can convert a sentence into keywords, phrases, abbreviations and symbols (KPAS)",
    ],
    NOTES_CLOSING
  );

  // Slide 13 -- Resources
  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // -----------------------------------------------------------------------
  // Companion PDFs
  // -----------------------------------------------------------------------

  // Practice
  const ws = createPdf({ title: PRACTICE_RESOURCE.name });
  let wsY = addPdfHeader(ws, "KPAS Practice -- Note Taking", {
    color: C.PRIMARY,
    subtitle: "Chapters 27-28: New Land",
    lessonInfo: "Lesson 13 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "Convert each sentence below into KPAS notes. Use the symbol reference. Drop little words like 'the', 'and', 'a'. Keep keywords -- who, what doing, when, where, why, how. Test by reading your notes back.", wsY, { color: C.PRIMARY });

  wsY = addSectionHeading(ws, "Symbol Reference", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "b/c = because    w/ = with    w/o = without    = means    + and    -> causes / results in", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "up arrow = more / increase    down arrow = less / decrease    / = full stop or new idea", wsY, { fontSize: 11 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Sentence 1", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "After eleven days waiting offshore, Tom was finally allowed to step off the ship and the marines were excited because they could begin building the new colony.", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "KPAS:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 6;

  wsY = addSectionHeading(ws, "Sentence 2", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Tom was sent to work in the swamp because the rushes there could be cut down and dried out for use as roofing.", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "KPAS:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 6;

  wsY = addSectionHeading(ws, "Sentence 3", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "As the days passed, the number of tents in the new colony went up and the food supplies went down.", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "KPAS:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 6;

  wsY = addSectionHeading(ws, "Sentence 4", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Despite the strangeness of sleeping on ground that didn't lurch and sway, Tom slept through the night.", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "KPAS:", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Self-Check", wsY, { color: C.ALERT });
  wsY = addBodyText(ws, "- Did I drop the small words like 'the', 'and', 'a'?", wsY, { fontSize: 10 });
  wsY = addBodyText(ws, "- Did I keep the keywords (who, what doing, when, where, why, how)?", wsY, { fontSize: 10 });
  wsY = addBodyText(ws, "- Did I use the new symbols where they fit (up arrow, down arrow, /)?", wsY, { fontSize: 10 });
  wsY = addBodyText(ws, "- Can I read my notes back and recover the original meaning?", wsY, { fontSize: 10 });

  addPdfFooter(ws, "Lesson 13 | KPAS Practice");

  // Answer key
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "KPAS Practice -- Model Answers", {
    color: C.ALERT,
    subtitle: "Teacher Reference -- Chapters 27-28",
    lessonInfo: "Lesson 13 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Model KPAS conversions. Student work will vary; assess whether keywords are preserved, the meaning is recoverable, and the new symbols are used appropriately.", akY, { color: C.ALERT });

  akY = addSectionHeading(ak, "Sentence 1", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Original: After eleven days waiting offshore, Tom was finally allowed to step off the ship and the marines were excited because they could begin building the new colony.", akY, { fontSize: 10, italic: true });
  akY = addBodyText(ak, "Model KPAS: 11 days / Tom off ship / marines excited b/c new colony begins", akY);
  akY += 6;

  akY = addSectionHeading(ak, "Sentence 2", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "Original: Tom was sent to work in the swamp because the rushes there could be cut down and dried out for use as roofing.", akY, { fontSize: 10, italic: true });
  akY = addBodyText(ak, "Model KPAS: Tom -> swamp / cut rushes / dried for roofing", akY);
  akY += 6;

  akY = addSectionHeading(ak, "Sentence 3", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "Original: As the days passed, the number of tents in the new colony went up and the food supplies went down.", akY, { fontSize: 10, italic: true });
  akY = addBodyText(ak, "Model KPAS: tents up arrow / food down arrow", akY);
  akY += 6;

  akY = addSectionHeading(ak, "Sentence 4", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Original: Despite the strangeness of sleeping on ground that didn't lurch and sway, Tom slept through the night.", akY, { fontSize: 10, italic: true });
  akY = addBodyText(ak, "Model KPAS: ground new / no lurching / Tom slept through night", akY);
  akY += 12;

  akY = addSectionHeading(ak, "What to Look For", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- Small words dropped (the, and, a)", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Keywords preserved (who, what doing, when, where, why)", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- New symbols used where they fit (up/down arrow for trend; / for breaks)", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Notes are short but meaning is recoverable", akY, { fontSize: 10 });

  addPdfFooter(ak, "Lesson 13 | KPAS Answer Key -- TEACHER COPY");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Lesson13.pptx` }),
    writePdf(ws, PRACTICE_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Lesson13.pptx`);
  console.log("Done: " + PRACTICE_RESOURCE.name);
  console.log("Done: " + ANSWER_KEY_RESOURCE.name);
}

build().catch(console.error);
