"use strict";

// Tom Appleby, Convict Boy -- Lesson 6: Chapters 11-14 -- Escape and Arrest
// Year 5/6 Literacy | Novel study
// Reading: Character analysis across four chapters (Jem's fear, Jem's death, Thomas/Millie, Tom's arrest)
// Writing (sentence-level): Identify and add appositives to sentences
// Sensitivity: Chapter 12 contains content about death (Jem dies in chimney collapse)

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

const SESSION_NUMBER = 6;
const FOOTER = "Chapters 11-14 | Lesson 6 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson6_Escape_and_Arrest";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const WORKSHEET_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Appositives Worksheet",
  "Student worksheet: identify and add appositives using content from Chapters 11-14."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Appositives Answer Key",
  "Teacher reference: model appositives and marking guidance."
);
const RESOURCE_ITEMS = [WORKSHEET_RESOURCE, ANSWER_KEY_RESOURCE];
const WORKSHEET_PDF_PATH = path.join(OUT_DIR, WORKSHEET_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- We are continuing our novel today, picking up at Chapter 11
- Four chapters to read -- 11, 12, 13 and 14. Big things happen for Tom
- Our writing focus is appositives -- extra detail about a noun, set off by commas

DO:
- Display title slide as students settle
- Have novels on desks, bookmarked at Chapter 11

TEACHER NOTES:
Chapters 11-14 cover Jem's second chimney incident, the escape plan, Jem's death, and Tom's arrest. Chapter 12 needs a sensitivity advisory. The appositive writing connects to the reading objective of expanding description.

WATCH FOR:
- Students who need a quick reminder of where we left off in Chapter 10
- Students processing content from the previous week

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_LI_SC = `SAY:
- Today has two strands: reading Chapters 11-14 for character development, and writing with appositives
- Read the learning intention and the three success criteria from the slide
- SC1 is our floor -- everyone can point to what a character says or does. SC2 is the core -- tracking how the author develops Tom across chapters. SC3 is the writing target -- adding an appositive

DO:
- Choral read the LI and each SC
- Do not over-explain appositives yet -- the I Do covers this

TEACHER NOTES:
SC1 is achievable for every student. SC2 is the core assessment target and is checked at the exit. SC3 is the writing application -- adding appositives.

WATCH FOR:
- Students who look blank at "appositives" -- reassure: "We will build this together step by step"
- Students already familiar with appositives -- they can support peers later

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_HOARD = `SAY:
- First word: hoard. Choral read: hoard
- Hoard is a secret store of valuable things, or to collect and hide things away
- In the novel, the boys have been hoarding stolen items for their escape
- Quick check: if someone has a hoard of sweets, have they shared them or hidden them? [Hidden]

DO:
- Display word, choral read, define, give example
- Show three images -- coins hidden under a floorboard, someone sharing food, a treasure chest -- ask: which show a hoard? [coins and treasure chest]
- Cold Call 2 students: "What might someone hoard, and why?"

TEACHER NOTES:
Pre-teaching "hoard" before reading supports comprehension of the boys' escape preparations. The word recurs around Tom's theft.

WATCH FOR:
- Students who confuse hoard with horde (a large group) -- clarify spelling and meaning
- Students who only see hoard as negative -- expand: squirrels hoard nuts, people hoard memories

[Literacy: Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_VOCAB_OCCUPIED = `SAY:
- Second word: occupied. Choral read: occupied
- Occupied means being used or taken up by someone or something. A room is occupied when someone is in it. A person is occupied when they are busy
- In the novel, Tom enters a bedroom and finds it occupied -- someone is sleeping in the bed
- Turn and Tell: name a time you found something was occupied -- a seat, a bathroom, a computer [20 seconds]

DO:
- Display word, choral read, define, example
- Turn and Tell: 20 seconds, then Cold Call 2 pairs
- Thumbs Up or Down: "If a classroom is empty, is it occupied?" [Thumbs down]

TEACHER NOTES:
"Occupied" is versatile. Teaching both the physical and abstract meanings builds transfer.

WATCH FOR:
- Students who confuse occupied with preoccupied -- acknowledge the link but distinguish: occupied = busy or in use; preoccupied = distracted by worry
- Students giving room examples only -- push for the "busy" meaning too

[Literacy: Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_VOCAB_RESOLVED = `SAY:
- Third word: resolved. Choral read: resolved
- Resolved means firmly decided or determined. Stronger than simply "decided"
- After something terrible happens in Chapter 12, Tom resolves to still try to escape. Despite everything, he is determined
- Ask: what is the difference between "decided" and "resolved"? [Resolved is stronger -- mind made up, will not change]

DO:
- Display word, choral read, define, example
- Cold Call: "Can you think of a time you resolved to do something hard?" [Take 2-3 responses]
- Sentence completion on whiteboards: "After failing the test, she resolved to ___"

TEACHER NOTES:
"Resolved" carries emotional weight -- Tom's determination despite grief. This word connects directly to character analysis in the pause points.

WATCH FOR:
- Students who think resolved only means "solved" -- clarify: here it means determined, not solved
- Students who give weak examples -- push for genuine determination, not casual decisions

[Literacy: Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_READING = `SAY:
- Chapters 11 through 14 today. I will read aloud and we will pause at key moments
- Focus: how does Tom respond to what happens? What do his actions and decisions tell us about his character?
- Chapter 12 has some sad content -- a character we know dies. If it feels uncomfortable, that is okay. We will talk about it together
- Find Chapter 11 now

DO:
- Give 30 seconds to find Chapter 11
- Teacher reads aloud to control pacing around the sensitive content
- Plan to pause at three key moments: p.57, p.64, and p.67

SENSITIVITY ADVISORY:
- What it is: Chapter 12 -- Jem dies when a chimney collapses on him. A significant character death that students may find upsetting given the bond between Tom and Jem
- Framing language: "This chapter deals with something very sad. The author included it because it was a real danger for climbing boys in this era. If you need a moment, that is completely fine"
- Watch for: students who withdraw, become visibly upset, or who have experienced loss recently
- Protocol: if a student becomes distressed, acknowledge their feelings privately. Do not require them to continue reading. Follow the school's wellbeing referral process. After reading, normalise the emotional response: "Good readers feel things when they read"

TEACHER NOTES:
Teacher reads aloud for these chapters to control pacing around the sensitive content in Chapter 12. Three pause points cover character analysis and the dual-timeline connection.

WATCH FOR:
- Students who disconnect during the reading -- gently re-engage with a quiet question
- Students who want to discuss Jem's death immediately -- allow brief acknowledgement, then redirect to the pause point

[Literacy: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause here. "It's an omen, thought Tom as he looked up and saw a flash of blue between the black"
- Ask: what does the author mean by "It's an omen"? What is Tom feeling? [Tom sees blue sky through the chimney and takes it as a sign of hope -- their escape plan will work]
- Ask: knowing what happens next, what effect does this create? [Dramatic irony -- the reader feels the tragedy more because Tom was so hopeful]

DO:
- Display the quote and read it aloud
- Ten seconds of think time
- Cold Call 3 students -- push for the connection between hope and what is about to happen

CFU CHECKPOINT:
Technique: Cold Call
Script:
- "What is an omen?" [A sign of something to come]
- "Is Tom seeing this as a good omen or a bad omen?" [Good -- the blue sky represents hope]
- Scan for: students connecting Tom's hope to the coming tragedy
PROCEED:
- Most students connect the omen to the emotional setup. Continue reading.
PIVOT:
- Most likely issue -- students take "omen" literally and miss the emotional setup
- Reteach: "Tom is underground in a dark chimney. He sees one flash of blue sky. To him, that tiny piece of sky means tomorrow he will be free. The author puts this here on PURPOSE, right before something terrible happens. Why?"
- Re-check: "What emotion does the author want us to feel right now?"

TEACHER NOTES:
This pause point is about authorial intent and dramatic irony. The omen comes just before Jem's death -- the author deliberately creates hope to make the loss more impactful.

WATCH FOR:
- Students who define "omen" but cannot explain why the author placed it here -- push: "Why does the author show us Tom's hope RIGHT BEFORE the tragedy?"

[Literacy: Pause Point 1 | VTLM 2.0: Higher-Order Questioning]`;

const NOTES_PAUSE2 = `SAY:
- Pause here. "...it could be any one of many"
- Context: Thomas's great-granddaughter Millie has visited, and Thomas is thinking about which ghost from his past is haunting him
- Ask: why might Thomas be particularly upset that Miss Hildegard put Millie in a cupboard? [The cupboard is a small, dark, enclosed space -- like a chimney. Thomas was a climbing boy. He knows what it is like to be trapped in dark spaces]
- Ask: why does the author say "it could be any one of many"? [Thomas has so many painful memories he cannot pinpoint which one is haunting him]

DO:
- Display the quote
- Think-Pair-Share: 20 seconds think, 30 seconds pair, then share
- Draw the connection between Tom's childhood and Thomas's reaction as an old man

CFU CHECKPOINT:
Technique: Think-Pair-Share
Script:
- "Think for 20 seconds: why would a cupboard upset Thomas so much?"
- "Share with your partner for 30 seconds"
- "Hands down -- I am selecting pairs. [Name], what did you discuss?"
- Scan for: students connecting the enclosed space to Tom's chimney experiences
PROCEED:
- Most pairs connect the cupboard to Tom's past. Continue reading.
PIVOT:
- Most likely issue -- students see the cupboard as generic punishment and miss the chimney connection
- Reteach: "Think about Tom as a boy. He was sent up chimneys -- tiny, dark, suffocating. Now he is old, and someone puts his great-granddaughter in a cupboard. A small, dark space. What would that remind him of?"
- Re-check: "Why would Thomas react more strongly than other adults?"

TEACHER NOTES:
This pause point connects the two timelines. The author uses Millie's experience to trigger Thomas's memories -- childhood trauma persists into old age. This is sophisticated character development.

WATCH FOR:
- Students who only see surface action (Millie in a cupboard) without connecting to Tom's past -- the chimney/cupboard parallel is the key insight
- Students making strong cross-timeline connections -- this is the analytical depth SC2 targets

[Literacy: Pause Point 2 | VTLM 2.0: Teacher-Led Discussion]`;

const NOTES_IDO_APPOSITIVE = `SAY:
- Time for our sentence-level writing focus: appositives
- An appositive is a noun or noun phrase placed right next to another noun to rename it or tell us more about it
- Think of it as a label that sits beside a name
- Watch me. Start with: "Tom climbed into the chimney"
- I want more information about Tom. I add an appositive: "Tom, the youngest of Master Jack's boys, climbed into the chimney"
- Notice "the youngest of Master Jack's boys" sits right next to "Tom" and tells us more. That is the appositive
- Notice the commas. The appositive is set off by commas because it is extra information. If I remove it, the sentence still works

DO:
- Display the model sentence and point to the appositive and the noun it renames
- Think aloud: "I check -- if I remove the appositive, does the sentence still make sense? Yes. So the commas are correct and the appositive is extra detail"
- Show a second quick example: "Jem, Tom's closest friend, whispered his escape plan"

TEACHER NOTES:
The remove-and-check think-aloud is the key self-checking move. Students need this when writing their own appositives.

MISCONCEPTIONS:
- Misconception: an appositive is the same as an adjective
  Why: both add description, so students conflate them
  Impact: students write adjectives when asked for appositives, or use adjective punctuation
  Quick correction: "An adjective describes -- like 'brave Tom'. An appositive RENAMES -- like 'Tom, the youngest boy'. An appositive is a noun phrase, not an adjective"

WATCH FOR:
- Students confused by the commas -- emphasise: "The commas are like brackets. They hold the extra information"
- Students mixing up appositives with relative clauses from earlier lessons -- clarify: appositives have no connecting word like "who" or "which"

[Literacy: I Do -- Modelling | VTLM 2.0: Explicit Teaching]`;

const NOTES_CFU = `SAY:
- Quick check. I will show you a sentence. Find the appositive
- On your whiteboards, write ONLY the appositive. You have 15 seconds
- Three, two, one -- hold up

DO:
- Display the sentence clearly
- Use Show Me Boards -- count down, then scan
- Scan for: "a cruel and violent man"

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Write ONLY the appositive on your whiteboard. 15 seconds"
- Count down, scan
- Look for: "a cruel and violent man" or close paraphrases
PROCEED:
- Most boards show the correct appositive. Reveal the answer and move to We Do / You Do.
PIVOT:
- Most likely issue -- students write the subject noun or the whole sentence
- Reteach: "The appositive sits between the commas. Find the two commas first. What is between them? That is the appositive. It renames the noun that comes right before the first comma"
- Re-check with a new sentence on the board: "Jem, Tom's closest friend, whispered the plan." What is the appositive?

TEACHER NOTES:
Show Me Boards give visible evidence from every student. The "find the commas" strategy gives students a concrete tool for identifying appositives.

WATCH FOR:
- Students who write "Master Jack" -- they have found the noun, not the appositive. Redirect: "That is WHO the appositive is about. What does it TELL US about Master Jack?"
- Students who identify the appositive quickly -- they are ready for the We Do / You Do

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- The appositive is "a cruel and violent man". It sits right next to "Master Jack" and tells us more about who he is
- Notice: remove it, and the sentence still works. The appositive is extra detail between the commas

DO:
- Reveal the answer on the bar
- Highlight the commas as the signal
- Transition to We Do / You Do

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_WEDO = `SAY:
- On your worksheet, you have sentences from our novel. Add an appositive to each noun marked with a dot
- Section A gives you a bank of appositives to choose from. Section B -- you write your own
- Let's do sentence 1 together. "Tom stole the coins from the bedroom." Who can suggest an appositive for Tom? [Take 2-3 suggestions -- e.g., "a desperate boy with nothing to lose", "the smallest of the climbing boys"]
- Good. Write it in with commas: "Tom, [your appositive], stole the coins from the bedroom"
- Now complete the rest. Section A first, then Section B. You have 8 minutes

DO:
- Distribute worksheets
- Model sentence 1 together (We Do), then release for You Do
- Circulate -- check commas and that appositives are noun phrases, not adjectives or clauses

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: provide the bank of descriptive noun phrases already on the worksheet and direct students to the three Section A sentences only. Each blank is clearly marked with commas
- Extra Notes: refer students back to the I Do model on screen
EXTENDING PROMPT:
- Task: write 2 original sentences about events from Chapters 11-14 that each contain an appositive. Then write a short paragraph (3-4 sentences) about one character where at least one sentence uses an appositive

TEACHER NOTES:
Sentence 1 with the class is the We Do. The remaining sentences become You Do. Section A scaffolds with a bank; Section B is the independent application.

WATCH FOR:
- Students writing adjectives instead of noun phrases -- redirect: "'Brave' is an adjective. 'A brave young boy' is a noun phrase"
- Students forgetting commas -- remind: "Commas hold the appositive in place"
- Students finishing Section A quickly and correctly -- they are ready for Section B

[Literacy: We Do / You Do | VTLM 2.0: Guided + Independent Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Let's check our answers. Your appositive does not need to match mine exactly. What matters is a noun phrase that renames the noun, set off by commas
- Sentence 1: "Tom, a desperate boy planning his escape, stole the coins from the bedroom"
- Sentence 2: "Jem, Tom's closest friend and fellow climbing boy, got stuck in the chimney again"
- Sentence 3: "Master Jack, the cruel chimney sweep master, whipped the boys as punishment"

DO:
- Reveal model answers
- Read each full sentence aloud with the appositive
- Invite students to share alternatives that also work
- Celebrate creative, accurate appositives

[Literacy: We Do / You Do Reveal | VTLM 2.0: Guided + Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: identifying what characters reveal through actions and words -- thumbs?
- SC2: explaining how the author develops Tom's character across chapters -- thumbs?
- SC3: adding an appositive to a sentence with commas -- thumbs?
- Turn and Tell: which moment across Chapters 11 to 14 was the biggest turning point for Tom, and why?

DO:
- Thumbs check each SC -- note any thumbs-down for follow-up next lesson
- The turn-and-tell targets the reading analysis -- students should identify a specific moment and justify it
- Preview: next lesson we continue with Chapters 15 to 17 -- Tom faces the consequences of his arrest

TEACHER NOTES:
The closing reconnects the reading arc to today's writing focus. Students thumbs-down on SC3 may benefit from reviewing the worksheet at home or receiving additional practice next lesson.

WATCH FOR:
- Students thumbs-down on SC3 -- check if it is terminology confusion or genuine difficulty with the concept
- Students who want to discuss Jem's death further -- acknowledge the emotion and encourage them to write about it

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two printable resources today
- The ${WORKSHEET_RESOURCE.name} is for the We Do / You Do appositive activity
- The ${ANSWER_KEY_RESOURCE.name} is for teacher reference

DO:
- Print the worksheet before the lesson (one per student)
- Print the answer key (teacher copy only)

TEACHER NOTES:
Section A has a bank to scaffold; Section B is independent. The answer key shows model responses and alternatives that also work.

[Literacy: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Chapters 11-14: Escape and Arrest -- Lesson 6";

  // Slide 1 -- Title
  titleSlide(
    pres,
    "Chapters 11-14",
    "Escape and Arrest",
    "Lesson 6  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  // Slide 2 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to analyse how an author develops character across key moments, and to add detail to sentences using appositives",
    ],
    [
      "I can identify what a character's actions and words reveal about their feelings",
      "I can explain how the author uses events across chapters to develop Tom's character",
      "I can add an appositive to a sentence to give the reader more information about a noun",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // Slide 3 -- Vocab: hoard
  vocabSlide(
    pres,
    "hoard",
    "noun / verb",
    "A secret store of valuable things; or, to collect and hide things away. People hoard things they want to protect or save for later.",
    "The boys had been hoarding stolen items, saving everything they could for their planned escape.",
    NOTES_VOCAB_HOARD,
    FOOTER
  );

  // Slide 4 -- Vocab: occupied
  vocabSlide(
    pres,
    "occupied",
    "adjective",
    "Being used or taken up by someone or something. A room is occupied when someone is in it. A person is occupied when they are busy.",
    "Tom crept into the bedroom only to discover it was occupied by a sleeping, drunk patron.",
    NOTES_VOCAB_OCCUPIED,
    FOOTER
  );

  // Slide 5 -- Vocab: resolved
  vocabSlide(
    pres,
    "resolved",
    "adjective / verb",
    "Firmly decided or determined. When you have resolved to do something, your mind is made up. Stronger than simply 'decided'.",
    "Despite the tragedy, Tom resolved to still try to escape the next day.",
    NOTES_VOCAB_RESOLVED,
    FOOTER
  );

  // Slide 6 -- Reading Launch
  contentSlide(
    pres,
    "Teacher Read Aloud",
    C.PRIMARY,
    "Chapters 11-14",
    [
      "Reading Mode: Teacher Read Aloud",
      "Chapter 11: Jem gets stuck again -- the escape plan becomes urgent",
      "Chapter 12: The boys head to a Public House -- tragedy strikes",
      "Chapter 13: Thomas reflects as Millie visits",
      "Chapter 14: The Bow Street Runners arrive",
      "Focus: how does Tom respond? What does this reveal about his character?",
    ],
    NOTES_READING,
    FOOTER
  );

  // Slide 7 -- Pause Point 1
  quoteSlide(
    pres,
    "Pause Point 1",
    "Chapter 12 -- p. 57",
    "It's an omen, thought Tom as he looked up and saw a flash of blue between the black.",
    "p. 57",
    "What does the author mean by 'It's an omen'? What is Tom feeling right before everything changes?",
    NOTES_PAUSE1,
    FOOTER
  );

  // Slide 8 -- Pause Point 2
  quoteSlide(
    pres,
    "Pause Point 2",
    "Chapter 13 -- p. 64",
    "...it could be any one of many.",
    "p. 64",
    "Given what we know about Tom's childhood, why might Thomas be particularly upset by Miss Hildegard putting Millie in a cupboard?",
    NOTES_PAUSE2,
    FOOTER
  );

  // Slide 9 -- I Do: Appositives
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Appositives",
    "An appositive is a noun or noun phrase placed right next to another noun to rename it or tell us more about it.\n\nIt is set off by commas.\n\nIf you remove it, the sentence still makes sense.\n\nAppositives are NOT:\n- Adjectives (brave, small)\n- Relative clauses (who was brave)",
    "Model:\n\n\"Tom climbed into the chimney.\"\n\nAdd an appositive about Tom:\n\n\"Tom, the youngest of Master Jack's boys, climbed into the chimney.\"\n\nCheck: remove it. \"Tom climbed into the chimney.\" Still works. The appositive is extra detail.",
    NOTES_IDO_APPOSITIVE,
    FOOTER
  );

  // Slides 10-11 -- CFU with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Find the Appositive",
      "Show Me Boards",
      "\"Master Jack, a cruel and violent man, sent the boys up the chimney.\"\n\nOn your whiteboard, write ONLY the appositive.\n\nHint: find the two commas first. What is between them?",
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
        x: 0.7, y: ansY + 0.10, w: 2.8, h: 0.28, rectRadius: 0.08,
        fill: { color: C.WHITE },
      });
      slide.addText("a cruel and violent man", {
        x: 0.7, y: ansY + 0.10, w: 2.8, h: 0.28,
        fontSize: 12, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addText("The appositive renames \"Master Jack\" -- it tells us more about who he is. Remove it and the sentence still works: \"Master Jack sent the boys up the chimney.\"", {
        x: 3.7, y: ansY + 0.08, w: 5.6, h: 0.84,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // Slides 12-13 -- We Do / You Do with reveal
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.SECONDARY);
      addBadge(s, "We Do / You Do", { color: C.SECONDARY, w: 2.0 });
      addTitle(s, "Add an Appositive");

      const sentences = [
        { num: "1", text: "Tom stole the coins from the bedroom.", color: C.PRIMARY },
        { num: "2", text: "Jem got stuck in the chimney again.", color: C.SECONDARY },
        { num: "3", text: "Master Jack whipped the boys as punishment.", color: C.ACCENT },
      ];

      sentences.forEach((sent, i) => {
        const sy = CONTENT_TOP + i * 1.18;
        addCard(s, 0.5, sy, 9, 1.06, { strip: sent.color, fill: C.WHITE });
        s.addShape("roundRect", {
          x: 0.7, y: sy + 0.10, w: 0.50, h: 0.28, rectRadius: 0.08,
          fill: { color: sent.color },
        });
        s.addText(sent.num, {
          x: 0.7, y: sy + 0.10, w: 0.50, h: 0.28,
          fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
        s.addText(sent.text, {
          x: 1.35, y: sy + 0.10, w: 7.8, h: 0.28,
          fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
        s.addText("Add an appositive after the noun. Use commas.", {
          x: 1.35, y: sy + 0.52, w: 7.8, h: 0.40,
          fontSize: 12, fontFace: FONT_B, color: C.MUTED, italic: true, valign: "middle", margin: 0,
        });
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO);
      return s;
    },
    (slide) => {
      const answers = [
        { y: CONTENT_TOP + 0.52, text: "\"Tom, a desperate boy planning his escape, stole the coins from the bedroom.\"" },
        { y: CONTENT_TOP + 1.70, text: "\"Jem, Tom's closest friend and fellow climbing boy, got stuck in the chimney again.\"" },
        { y: CONTENT_TOP + 2.88, text: "\"Master Jack, the cruel chimney sweep master, whipped the boys as punishment.\"" },
      ];
      answers.forEach((ans) => {
        slide.addShape("roundRect", {
          x: 0.7, y: ans.y, w: 8.5, h: 0.40, rectRadius: 0.06,
          fill: { color: C.BG_LIGHT },
        });
        slide.addText(ans.text, {
          x: 0.8, y: ans.y + 0.02, w: 8.2, h: 0.36,
          fontSize: 11, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
        });
      });
      slide.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // Slide 14 -- Closing
  closingSlide(
    pres,
    "Which moment across Chapters 11 to 14 was the biggest turning point for Tom, and why? Tell your partner.",
    [
      "I can identify what a character's actions and words reveal",
      "I can explain how the author develops Tom's character across chapters",
      "I can add an appositive to a sentence using commas",
    ],
    NOTES_CLOSING
  );

  // Slide 15 -- Resources
  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // ---------------------------------------------------------------------------
  // PDF 1 -- Appositives Worksheet
  // ---------------------------------------------------------------------------
  const ws = createPdf({ title: WORKSHEET_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Appositives -- Add Detail to Sentences", {
    color: C.PRIMARY,
    subtitle: "Chapters 11-14: Escape and Arrest",
    lessonInfo: "Lesson 6 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(
    ws,
    "An appositive is a noun or noun phrase placed next to another noun to rename it or give more detail. It is set off by commas. Example: \"Tom, the youngest of the climbing boys, looked up at the chimney.\"",
    wsY,
    { color: C.PRIMARY }
  );

  wsY = addSectionHeading(ws, "Section A: Add an Appositive (use the bank)", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Choose an appositive from the bank and add it to each sentence. Remember to use commas.", wsY);
  wsY += 4;
  wsY = addTipBox(
    ws,
    "Appositive Bank:  a desperate boy with nothing to lose   |   Tom's closest friend and fellow climbing boy   |   the cruel chimney sweep master   |   a sleeping, drunk patron",
    wsY,
    { color: C.SECONDARY }
  );

  wsY = addBodyText(ws, "1. Tom stole the coins from the bedroom.", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "2. Jem got stuck in the chimney again.", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "3. Master Jack whipped the boys as punishment.", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Section B: Add Your Own Appositive", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "Now write your own appositive for each sentence. Use commas. Your appositive should be a noun phrase that renames the noun.", wsY);
  wsY += 4;

  wsY = addBodyText(ws, "4. The Bow Street Runners searched the boys.", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "5. Thomas remembered the climbing boys he had known.", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "6. Millie wanted to take over the farm one day.", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });

  addPdfFooter(ws, "Lesson 6 | Chapters 11-14 | Appositives Worksheet");
  await writePdf(ws, WORKSHEET_PDF_PATH);

  // ---------------------------------------------------------------------------
  // PDF 2 -- Answer Key
  // ---------------------------------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Appositives -- Answer Key", {
    color: C.PRIMARY,
    subtitle: "Teacher Reference",
    lessonInfo: "Lesson 6 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(
    ak,
    "Marking note: appositives that are different nouns phrases but still rename the noun correctly (and use commas) should be marked correct. Look for noun phrases, NOT adjectives or relative clauses.",
    akY,
    { color: C.ALERT }
  );

  akY = addSectionHeading(ak, "Section A -- Model Answers", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "1. Tom, a desperate boy with nothing to lose, stole the coins from the bedroom.", akY, { fontSize: 11 });
  akY = addBodyText(ak, "2. Jem, Tom's closest friend and fellow climbing boy, got stuck in the chimney again.", akY, { fontSize: 11 });
  akY = addBodyText(ak, "3. Master Jack, the cruel chimney sweep master, whipped the boys as punishment.", akY, { fontSize: 11 });
  akY += 8;

  akY = addSectionHeading(ak, "Section B -- Sample Responses", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "4. The Bow Street Runners, the law enforcers of 18th century London, searched the boys.", akY, { fontSize: 11 });
  akY = addBodyText(ak, "5. Thomas, an old man with a long and painful memory, remembered the climbing boys he had known.", akY, { fontSize: 11 });
  akY = addBodyText(ak, "6. Millie, Thomas's great-granddaughter, wanted to take over the farm one day.", akY, { fontSize: 11 });
  akY += 10;

  akY = addSectionHeading(ak, "Common Student Errors", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "- Writing an adjective instead of a noun phrase (\"brave\" instead of \"a brave boy\")", akY);
  akY = addBodyText(ak, "- Forgetting commas around the appositive", akY);
  akY = addBodyText(ak, "- Writing a relative clause (\"who was cruel\") instead of an appositive", akY);
  akY = addBodyText(ak, "- Placing the appositive far from the noun it renames", akY);

  addPdfFooter(ak, "Lesson 6 | Answer Key -- TEACHER COPY");
  await writePdf(ak, ANSWER_KEY_PDF_PATH);

  // ---------------------------------------------------------------------------
  // Write PPTX
  // ---------------------------------------------------------------------------
  const outName = path.join(OUT_DIR, "Tom_Lesson6_Chapters_11-14.pptx");
  fs.mkdirSync(OUT_DIR, { recursive: true });
  await pres.writeFile({ fileName: outName });
  console.log("PPTX written to " + outName);
  console.log("PDF written to " + WORKSHEET_PDF_PATH);
  console.log("PDF written to " + ANSWER_KEY_PDF_PATH);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
