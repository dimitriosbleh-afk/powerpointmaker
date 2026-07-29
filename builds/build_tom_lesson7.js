"use strict";

// Tom Appleby, Convict Boy -- Lesson 7: Chapters 15-17 -- Prison and Survival
// Year 5/6 Literacy | Novel study
// Reading: Character analysis; literary devices (personification, metaphor, imagery)
// Writing (sentence-level): Add relative clauses to sentences (who, whom, which, whose, that, when, where, why)
// Sensitivity: Chapter 16 contains content about death and discussion of hanging

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

const SESSION_NUMBER = 7;
const FOOTER = "Chapters 15-17 | Lesson 7 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson7_Prison_and_Survival";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const WORKSHEET_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Relative Clauses Worksheet",
  "Student worksheet: identify and add relative clauses using content from Chapters 15-17."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Relative Clauses Answer Key",
  "Teacher reference: model relative clauses, punctuation notes and marking guidance."
);
const RESOURCE_ITEMS = [WORKSHEET_RESOURCE, ANSWER_KEY_RESOURCE];
const WORKSHEET_PDF_PATH = path.join(OUT_DIR, WORKSHEET_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- We pick up after Tom's arrest. Three chapters today -- 15, 16 and 17
- Our writing focus is relative clauses -- adding extra information about a noun using words like who, which or that
- Chapter 16 deals with prison life including a discussion of hanging. We will handle this sensitively

DO:
- Display title slide as students settle
- Have novels on desks, bookmarked at Chapter 15

TEACHER NOTES:
Chapters 15-17 cover Newgate Prison, Tom's meeting with Bald Sally, and his sentencing. Two sensitive moments: talk of hanging (Ch 15), and prison conditions (Ch 16). Relative clauses are a continuation of earlier grammar work.

WATCH FOR:
- Students still processing Jem's death from last lesson -- acknowledge briefly, then move on
- Students who need a quick reminder of where we left Tom (arrested at the cellar)

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_LI_SC = `SAY:
- Today's learning intention: we are learning to analyse how an author builds character under pressure, and to add detail to sentences using relative clauses
- Read the success criteria from the slide. SC1 is the floor -- everyone can identify one thing Tom does or feels. SC2 is the core -- how the author builds Tom's survival. SC3 is the writing target
- Relative clauses -- we have met these before. Today we revise, then practise

DO:
- Choral read LI and SCs
- Do not over-explain relative clauses -- the I Do covers it

TEACHER NOTES:
The writing target (SC3) is the exit focus. Any student who finishes early extends into writing two-clause sentences.

WATCH FOR:
- Students unsure what a relative clause is -- reassure: "We'll build this together"
- Students who remember clearly from previous lessons -- invite them to help during the CFU

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_REPULSED = `SAY:
- First word: repulsed. Choral read: repulsed
- Repulsed means to feel strong disgust -- a strong, negative feeling that makes you want to turn away
- In Chapter 16, Tom is repulsed by some of what he sees in prison -- the smells, the dirt, the other prisoners
- Quick check: if someone is repulsed by food, do they want to eat it? [No -- they want to turn away]

DO:
- Display word, choral read, define, example
- Sentence completion on boards: "I was repulsed by the smell of ___"
- Cold Call 2 students: "Give me an example of something someone might be repulsed by"

TEACHER NOTES:
"Repulsed" connects to the sensory detail the author uses in prison scenes. Students often know the feeling but not the word.

WATCH FOR:
- Students using "repulsed" for mild dislike -- clarify: it is a strong word, not just "didn't like"
- Students who confuse it with "repelled" (pushed back) -- both come from the same root but "repulsed" means emotionally disgusted

[Literacy: Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_VOCAB_CONDEMNED = `SAY:
- Second word: condemned. Choral read: condemned
- Condemned means officially declared guilty, or expressing strong disapproval
- When a judge condemns someone, they formally declare that person guilty -- often with a serious punishment
- In these chapters, Tom worries about being condemned to death for his theft
- Word family: condemn (verb), condemned (adjective), condemnation (noun)

DO:
- Display word, choral read, define, example
- Ask: what is the difference between "accused" and "condemned"? [Accused = someone says you did it. Condemned = officially declared guilty]
- Partner talk 20 seconds: a time someone was unfairly condemned (in a story they know)

TEACHER NOTES:
"Condemned" is central to Tom's fear and the plot of Chapter 15-17. Linking it to the judicial process helps students read the court scene.

WATCH FOR:
- Students who only hear the "condemn to death" meaning -- broaden: a building can be condemned (declared unsafe)
- Students confusing "condemn" with "condone" -- opposite meanings. Condone means to accept; condemn means to strongly disapprove

[Literacy: Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_VOCAB_EARNEST = `SAY:
- Third word: earnest. Choral read: earnest
- Earnest means serious and sincere -- you really mean what you say, no joking
- When Tom speaks to Bald Sally earnestly, he is serious, not playing around
- Quick check: is "earnest" closer to "joking" or "serious"? [Serious]

DO:
- Display word, choral read, define, example
- Sentence completion: "She earnestly asked me to ___"
- Ask: can you think of a time you spoke earnestly about something?

TEACHER NOTES:
"Earnest" is a subtle character word -- it tells us how Tom is speaking, which reveals his emotional state. Tie it to the novel when students encounter it.

WATCH FOR:
- Students confusing "earnest" with "ernest" (the name) -- they sound the same
- Students interpreting it as "strongly" -- closer to "sincerely"

[Literacy: Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_READING = `SAY:
- Three chapters today -- 15, 16 and 17. I will read aloud again
- Focus: how does Tom survive? What does the author show us about his character under extreme pressure?
- A note on Chapter 16 -- it talks about hanging, which was a real punishment in this time. If this feels heavy, that is okay. We will not dwell on gruesome detail
- Find Chapter 15 now

DO:
- Give 30 seconds to find the chapter
- Teacher reads aloud to control pacing
- Two pause points today: p.72 and p.81

SENSITIVITY ADVISORY:
- What it is: Chapter 16 discusses hanging as a punishment. Chapter 15 describes harsh prison conditions
- Framing language: "This was how the justice system worked in 18th century England. The author shows us this because it helps us understand how scared Tom was and why escape mattered so much"
- Watch for: students who seem overwhelmed, students who have experienced loss
- Protocol: pause reading if needed, acknowledge the heaviness, do not require continued engagement if a student opts out

TEACHER NOTES:
These chapters are dark but the focus is survival. Frame Tom's adaptation as resourcefulness, not glorification of theft. The author shows a child forced into survival strategies.

WATCH FOR:
- Students processing the severity of Tom's situation -- this is appropriate engagement
- Students who want to discuss capital punishment -- briefly acknowledge, redirect to the character analysis focus

[Literacy: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause here. "Rats survived."
- This line comes right after Tom figures out how to survive in the prison
- Ask: how does this short sentence add to what we know about Tom? [Tom compares himself to rats -- he has lowered himself to whatever it takes to live. It shows both his resilience and how much he has been reduced]
- Ask: why does the author use such a short sentence here? [The short sentence hits hard. It makes us stop and feel the weight. No extra words -- just the brutal truth]

DO:
- Display the quote
- Ten seconds think time
- Cold Call 3 students
- Draw attention to the sentence LENGTH as a craft choice

CFU CHECKPOINT:
Technique: Cold Call
Script:
- "How does comparing himself to a rat change how we see Tom?"
- "Why might the author have used only two words for this sentence?"
- Scan for: students reading the craft choice, not just the plot
PROCEED:
- Most students engage with the sentence as a craft choice. Continue reading.
PIVOT:
- Most likely issue -- students read only the literal meaning (there were rats in the prison)
- Reteach: "The author could have written 'Tom survived like a rat, by being tough and clever.' Instead, the author wrote two words: 'Rats survived.' What does that short sentence DO to us as readers?"
- Re-check: "Why would the author choose brevity here?"

TEACHER NOTES:
This pause point targets authorial craft -- sentence length as a deliberate choice. This is SC2 territory: how the author develops character through writing choices.

WATCH FOR:
- Students who only see literal rats -- redirect to Tom's comparison
- Students connecting this to earlier "he had resolved" lines -- excellent: this is cross-chapter synthesis

[Literacy: Pause Point 1 | VTLM 2.0: Higher-Order Questioning]`;

const NOTES_PAUSE2 = `SAY:
- Pause here. "...and wrapped his courage cloak more firmly round his shoulders."
- This is a metaphor. Tom does not really have a cloak -- the author is describing how Tom mentally prepares himself
- Ask: what is the author comparing courage to? [A cloak -- something you wrap around yourself, that covers and protects you]
- Ask: the author has used this metaphor before. Why do you think they repeat it? [It shows Tom's courage is something he has to put on deliberately. It takes effort. He has to keep doing it again and again]

DO:
- Display the quote
- Think-Pair-Share: 20 seconds think, 30 seconds pair
- Draw attention to this as a recurring metaphor -- ask students to flag it when they see it again

CFU CHECKPOINT:
Technique: Think-Pair-Share
Script:
- "What does the cloak stand for?"
- "Why does the author use the same metaphor again and again?"
- Scan for: students identifying the metaphor AND the repetition
PROCEED:
- Most pairs identify the cloak as courage. Continue reading.
PIVOT:
- Most likely issue -- students interpret "cloak" literally
- Reteach: "A cloak is something you put on. You can take it off. You have to wrap it around yourself. Now replace 'cloak' with 'courage'. What does that tell us about Tom's bravery?"
- Re-check: "Is Tom brave naturally, or does he have to work at it?"

TEACHER NOTES:
This metaphor is one of the novel's recurring motifs. Recognising it as repeated imagery supports SC2 -- students analysing craft across chapters.

WATCH FOR:
- Students who spot the metaphor but not its repetition -- tell them the author uses this image throughout the novel
- Students confusing metaphor and simile -- a simile uses like or as, a metaphor does not

[Literacy: Pause Point 2 | VTLM 2.0: Teacher-Led Discussion]`;

const NOTES_IDO_REL = `SAY:
- Our writing focus today: relative clauses
- A relative clause is a group of words that gives extra information about a noun. It starts with a relative pronoun -- who, whom, which, whose, that -- or a relative adverb -- when, where, why
- Watch me. Start with: "Tom was arrested"
- I want to tell the reader WHICH Tom. I add a relative clause: "Tom, who had stolen the coins, was arrested"
- Notice "who had stolen the coins" comes right after Tom -- the noun it tells us about
- Notice the commas. This is EXTRA information. Remove it, and the sentence still works: "Tom was arrested"
- Try another: "The judge condemned the prisoners that he believed were guilty"
- This time no commas. Why? Because "that he believed were guilty" is ESSENTIAL -- it tells us WHICH prisoners. Without it, we do not know which prisoners were condemned

DO:
- Display the terminology clearly
- Point to the relative pronoun and the noun it modifies in each example
- Think aloud the removing test: "Remove the clause. Does the sentence still tell us WHO? If yes, use commas. If no, do not"

TEACHER NOTES:
The essential vs non-essential distinction is tricky. Use the "remove and check" test consistently. Students can ignore the formal terminology if they follow the test.

MISCONCEPTIONS:
- Misconception: a relative clause is the same as an appositive
  Why: both add extra information about a noun
  Impact: students use the wrong punctuation or confuse the grammar terms
  Quick correction: "An appositive is a NOUN PHRASE. A relative clause starts with a CONNECTING WORD like who, which, or that. If there is no connecting word, it's an appositive"
- Misconception: always use commas with relative clauses
  Why: that is how appositives work, and students overgeneralise
  Impact: incorrect punctuation
  Quick correction: "If the clause is essential (tells us WHICH one), no commas. If the clause is extra (adds info we could live without), commas"

WATCH FOR:
- Students placing the clause too far from the noun -- remind: right after the noun it modifies
- Students starting a sentence with a relative clause -- relative clauses NEVER start a sentence

[Literacy: I Do -- Modelling | VTLM 2.0: Explicit Teaching]`;

const NOTES_CFU = `SAY:
- Quick check. On your whiteboard, find the relative clause in this sentence
- Write the whole relative clause, starting with the relative pronoun
- You have 20 seconds

DO:
- Display the sentence
- Use Show Me Boards -- count down, then scan
- Scan for: "who was exhausted from climbing"

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Write ONLY the relative clause on your board. 20 seconds"
- Count down, scan
- Look for: "who was exhausted from climbing" or close variants
PROCEED:
- Most boards show the correct clause. Reveal and move to We Do.
PIVOT:
- Most likely issue -- students write the whole sentence or miss the relative pronoun
- Reteach: "A relative clause starts with a CONNECTING WORD -- who, which, that, whose, when, where. Find the connecting word first. What follows it, up to the next comma or the end of the clause? That is your relative clause"
- Re-check with a new sentence: "The prison, which was dark and crowded, terrified Tom." What is the relative clause?

TEACHER NOTES:
Show Me Boards give visible evidence from every student. The "find the connecting word first" strategy is the scaffold.

WATCH FOR:
- Students who only write the relative pronoun ("who") -- they've started but not finished the clause
- Students who write the full sentence -- they haven't isolated the clause

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- The relative clause is "who was exhausted from climbing". It starts with "who" and tells us more about Tom
- Notice: remove it and the sentence still works. "Tom slept for twelve hours straight." So this clause is NON-ESSENTIAL -- commas are correct
- Now you practise on your worksheet

DO:
- Reveal the answer
- Highlight the relative pronoun "who" as the starting signal
- Transition to We Do / You Do

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_WEDO = `SAY:
- On your worksheet, you have sentences. Add a relative clause to each, answering either "Which one?" or "What kind?"
- Let's do sentence 1 together. "The judge sentenced Tom." Which judge? Add a relative clause
- Who can suggest one? [Take 2-3 -- e.g., "who wore a long white wig", "that had a stern face"]
- Great. Write it in: "The judge, who wore a long white wig, sentenced Tom"
- Ask yourselves: is my clause essential or non-essential? If essential, no commas. If extra, commas
- Work through the rest. Section A scaffolds. Section B is independent. 10 minutes

DO:
- Distribute worksheets
- Model sentence 1 together (We Do), release for You Do
- Circulate -- check for relative pronouns and correct punctuation

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: use the provided bank of relative clauses on the worksheet. Students select and place each clause with correct commas using the marked positions
- Extra Notes: refer to the I Do model on screen; students can work in pairs
EXTENDING PROMPT:
- Task: write 3 original sentences about Chapters 15-17 that each contain a relative clause. At least one must be essential (no commas), and one must be non-essential (with commas). Label each as essential or non-essential

TEACHER NOTES:
The essential vs non-essential distinction is the challenging part. The answer key shows which is which for each sentence.

WATCH FOR:
- Students omitting relative pronouns -- remind: every relative clause starts with who, whom, which, whose, that, when, where, or why
- Students putting commas on essential clauses -- reteach with the removal test
- Students confusing relative clauses with appositives -- check that the clause starts with a connecting word

[Literacy: We Do / You Do | VTLM 2.0: Guided + Independent Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Let's check. Your clause does not need to match mine exactly -- what matters is that it starts with a relative pronoun, sits beside the noun, and has the right punctuation
- Sentence 1: "The judge, who wore a long white wig, sentenced Tom." Non-essential -- commas
- Sentence 2: "The prisoner who slept beside Tom snored loudly." Essential -- no commas (tells us WHICH prisoner)
- Sentence 3: "Newgate Prison, which was overcrowded and dirty, frightened Tom." Non-essential -- commas

DO:
- Reveal each model answer
- Read each full sentence aloud with the relative clause
- Invite students to share alternatives that also work

[Literacy: We Do / You Do Reveal | VTLM 2.0: Guided + Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: identifying how Tom survives -- thumbs?
- SC2: explaining how the author builds Tom's survival through craft choices -- thumbs?
- SC3: adding a relative clause to a sentence with correct punctuation -- thumbs?
- Turn and Tell: what does the author want us to understand about the justice system through Tom's court hearing?

DO:
- Thumbs check each SC
- The turn-and-tell targets the reading analysis -- push students toward the author's intent, not just plot summary
- Preview: next lesson, Chapters 18-19 -- Tom is put on the ship, and we meet Sailor Sam

TEACHER NOTES:
The closing reconnects the reading arc to today's writing focus. Any student thumbs-down on SC3 can revisit the worksheet at home.

WATCH FOR:
- Students thumbs-down on SC3 -- is it punctuation or the concept of relative clause? Different pivots needed
- Students with strong responses on SC2 -- these are the analytic depth indicators

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two printable resources today
- The ${WORKSHEET_RESOURCE.name} is for the We Do / You Do relative clause activity
- The ${ANSWER_KEY_RESOURCE.name} is for teacher reference, including essential/non-essential notes

DO:
- Print the worksheet before the lesson (one per student)
- Print the answer key (teacher copy only)

TEACHER NOTES:
Section A has a bank to scaffold. Section B is independent. The answer key specifies essential vs non-essential for each sentence.

[Literacy: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Chapters 15-17: Prison and Survival -- Lesson 7";

  // Slide 1 -- Title
  titleSlide(
    pres,
    "Chapters 15-17",
    "Prison and Survival",
    "Lesson 7  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // Slide 2 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to analyse how the author builds Tom's survival under pressure, and to add detail to sentences using relative clauses",
    ],
    [
      "I can identify what Tom does or says to survive in the prison",
      "I can explain how the author uses craft choices to show Tom's character under pressure",
      "I can add a relative clause to a sentence using correct punctuation",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // Slide 3 -- Vocab: repulsed
  vocabSlide(
    pres,
    "repulsed",
    "adjective / verb",
    "To feel strong disgust -- a powerful negative feeling that makes you want to turn away. Stronger than 'disliked'.",
    "Tom was repulsed by the stench of the overcrowded prison cell.",
    NOTES_VOCAB_REPULSED,
    FOOTER
  );

  // Slide 4 -- Vocab: condemned
  vocabSlide(
    pres,
    "condemned",
    "verb / adjective",
    "Officially declared guilty, often with a serious punishment; or, to strongly disapprove of something. Courts condemn; so do people with strong views.",
    "Tom feared he would be condemned to hang for his theft.",
    NOTES_VOCAB_CONDEMNED,
    FOOTER
  );

  // Slide 5 -- Vocab: earnest
  vocabSlide(
    pres,
    "earnest",
    "adjective",
    "Serious and sincere -- you really mean what you say, no joking. Someone earnest is not playing around.",
    "Tom spoke earnestly to Bald Sally, hoping she would help him.",
    NOTES_VOCAB_EARNEST,
    FOOTER
  );

  // Slide 6 -- Reading Launch
  contentSlide(
    pres,
    "Teacher Read Aloud",
    C.PRIMARY,
    "Chapters 15-17",
    [
      "Reading Mode: Teacher Read Aloud",
      "Chapter 15: Tom pleads guilty -- sent to Newgate Prison",
      "Chapter 16: Tom meets Bald Sally, a midwife",
      "Chapter 17: Tom's court hearing -- sentenced to transportation",
      "Focus: how does Tom survive? What craft choices does the author make?",
    ],
    NOTES_READING,
    FOOTER
  );

  // Slide 7 -- Pause Point 1: Rats survived
  quoteSlide(
    pres,
    "Pause Point 1",
    "Chapter 15 -- p. 72",
    "Rats survived.",
    "p. 72",
    "How does comparing Tom to a rat change how we see him? Why has the author used such a short sentence here?",
    NOTES_PAUSE1,
    FOOTER
  );

  // Slide 8 -- Pause Point 2: courage cloak metaphor
  quoteSlide(
    pres,
    "Pause Point 2",
    "Chapter 17 -- p. 81",
    "...and wrapped his courage cloak more firmly round his shoulders.",
    "p. 81",
    "What is the author comparing courage to? Why might the author use this same image more than once?",
    NOTES_PAUSE2,
    FOOTER
  );

  // Slide 9 -- I Do: Relative Clauses
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Relative Clauses",
    "A relative clause gives extra information about a noun.\n\nIt starts with:\n- who / whom / which / whose / that\n- when / where / why\n\nIt sits right after the noun it tells us about.\n\nPunctuation:\n- ESSENTIAL (tells us WHICH one)  =  no commas\n- EXTRA (adds info we could live without)  =  commas",
    "Example 1 (extra):\n\"Tom, who had stolen the coins, was arrested.\"\n\nExample 2 (essential):\n\"The prisoners that he believed were guilty were condemned.\"\n\nTest: remove the clause. Does the sentence still tell us WHICH one? If yes, use commas.",
    NOTES_IDO_REL,
    FOOTER
  );

  // Slides 10-11 -- CFU with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Find the Relative Clause",
      "Show Me Boards",
      "\"Tom, who was exhausted from climbing, slept for twelve hours straight.\"\n\nOn your whiteboard, write ONLY the relative clause.\n\nHint: find the connecting word first (who, which, that, when, where, why).",
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
        x: 0.7, y: ansY + 0.10, w: 3.3, h: 0.28, rectRadius: 0.08,
        fill: { color: C.WHITE },
      });
      slide.addText("who was exhausted from climbing", {
        x: 0.7, y: ansY + 0.10, w: 3.3, h: 0.28,
        fontSize: 11, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addText("Starts with \"who\" and sits right after \"Tom\". Remove it and the sentence still works -- so this is NON-ESSENTIAL. Commas are correct.", {
        x: 4.2, y: ansY + 0.08, w: 5.1, h: 0.84,
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
      addTitle(s, "Add a Relative Clause");

      const sentences = [
        { num: "1", text: "The judge sentenced Tom.", color: C.PRIMARY },
        { num: "2", text: "The prisoner snored loudly.", color: C.SECONDARY },
        { num: "3", text: "Newgate Prison frightened Tom.", color: C.ACCENT },
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
        s.addText("Add a relative clause. Decide: essential (no commas) or extra (commas)?", {
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
        { y: CONTENT_TOP + 0.52, text: "\"The judge, who wore a long white wig, sentenced Tom.\"  (extra -- commas)" },
        { y: CONTENT_TOP + 1.70, text: "\"The prisoner who slept beside Tom snored loudly.\"  (essential -- no commas)" },
        { y: CONTENT_TOP + 2.88, text: "\"Newgate Prison, which was overcrowded and dirty, frightened Tom.\"  (extra -- commas)" },
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
    "What does the author want us to understand about the justice system through Tom's court hearing? Tell your partner.",
    [
      "I can identify what Tom does or says to survive",
      "I can explain how the author uses craft choices",
      "I can add a relative clause with correct punctuation",
    ],
    NOTES_CLOSING
  );

  // Slide 15 -- Resources


  // ---------------------------------------------------------------------------
  // PDF 1 -- Relative Clauses Worksheet
  // ---------------------------------------------------------------------------
  const ws = createPdf({ title: WORKSHEET_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Relative Clauses -- Add Detail to Sentences", {
    color: C.PRIMARY,
    subtitle: "Chapters 15-17: Prison and Survival",
    lessonInfo: "Lesson 7 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(
    ws,
    "A relative clause starts with who, whom, which, whose, that, when, where, or why. It sits right after the noun it tells us about. Essential clause (tells us WHICH one) = no commas. Extra clause (adds information) = commas.",
    wsY,
    { color: C.PRIMARY }
  );

  wsY = addSectionHeading(ws, "Section A: Add a Relative Clause (use the bank)", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Choose a relative clause from the bank. Decide: essential (no commas) or extra (commas).", wsY);
  wsY += 4;
  wsY = addTipBox(
    ws,
    "Clause Bank:   who wore a long white wig   |   that slept beside Tom   |   which was overcrowded and dirty   |   whose baby Sally delivered   |   where Tom was held before trial",
    wsY,
    { color: C.SECONDARY }
  );

  wsY = addBodyText(ws, "1. The judge sentenced Tom.", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "2. The prisoner snored loudly.", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "3. Newgate Prison frightened Tom.", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Section B: Write Your Own", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "Write your own relative clause for each sentence. Decide: essential or extra? Mark your choice.", wsY);
  wsY += 4;

  wsY = addBodyText(ws, "4. Bald Sally helped Tom.   (essential / extra?)", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "5. The court hearing was over quickly.   (essential / extra?)", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "6. Tom wore his courage cloak.   (essential / extra?)", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });

  addPdfFooter(ws, "Lesson 7 | Chapters 15-17 | Relative Clauses Worksheet");
  await writePdf(ws, WORKSHEET_PDF_PATH);

  // ---------------------------------------------------------------------------
  // PDF 2 -- Answer Key
  // ---------------------------------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Relative Clauses -- Answer Key", {
    color: C.PRIMARY,
    subtitle: "Teacher Reference",
    lessonInfo: "Lesson 7 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(
    ak,
    "Marking note: accept any relative clause that starts with who/whom/which/whose/that/when/where/why and sits right after the noun. Punctuation must match the essential/extra status. Use the removal test: can the sentence stand without the clause?",
    akY,
    { color: C.ALERT }
  );

  akY = addSectionHeading(ak, "Section A -- Model Answers", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "1. The judge, who wore a long white wig, sentenced Tom.  (extra -- commas)", akY, { fontSize: 11 });
  akY = addBodyText(ak, "2. The prisoner that slept beside Tom snored loudly.  (essential -- no commas, tells WHICH prisoner)", akY, { fontSize: 11 });
  akY = addBodyText(ak, "3. Newgate Prison, which was overcrowded and dirty, frightened Tom.  (extra -- commas)", akY, { fontSize: 11 });
  akY += 8;

  akY = addSectionHeading(ak, "Section B -- Sample Responses", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "4. Bald Sally, who was a midwife in Newgate, helped Tom.  (extra)", akY, { fontSize: 11 });
  akY = addBodyText(ak, "   OR: Bald Sally, whose baby needed delivering, helped Tom.  (extra)", akY, { fontSize: 11 });
  akY += 4;
  akY = addBodyText(ak, "5. The court hearing, which Tom had dreaded for weeks, was over quickly.  (extra)", akY, { fontSize: 11 });
  akY += 4;
  akY = addBodyText(ak, "6. Tom wore his courage cloak when he stood before the judge.  (essential -- WHEN he wore it)", akY, { fontSize: 11 });
  akY += 10;

  akY = addSectionHeading(ak, "Essential vs Non-Essential -- Quick Guide", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "- Essential: the clause tells us WHICH one. Removing it changes meaning.  No commas.", akY);
  akY = addBodyText(ak, "- Non-essential / Extra: the clause adds info. Removing it keeps the meaning clear.  Use commas.", akY);
  akY += 4;
  akY = addBodyText(ak, "Removal test: read the sentence without the clause. If it still tells the reader WHICH one, the clause was extra (commas). If meaning is lost, the clause was essential (no commas).", akY);
  akY += 10;

  akY = addSectionHeading(ak, "Common Student Errors", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- Omitting the relative pronoun -- clause must start with who/which/that/whose/when/where/why", akY);
  akY = addBodyText(ak, "- Starting a sentence with the relative clause -- relative clauses never begin a sentence", akY);
  akY = addBodyText(ak, "- Using commas on essential clauses -- use the removal test", akY);
  akY = addBodyText(ak, "- Confusing relative clauses with appositives (no connecting word)", akY);

  addPdfFooter(ak, "Lesson 7 | Answer Key -- TEACHER COPY");
  await writePdf(ak, ANSWER_KEY_PDF_PATH);

  // ---------------------------------------------------------------------------
  // Write PPTX
  // ---------------------------------------------------------------------------
  const outName = path.join(OUT_DIR, "Tom_Lesson7_Chapters_15-17.pptx");
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
