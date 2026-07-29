"use strict";

// Tom Unit -- Session 6: Chapters 11-14 -- Escape and Arrest
// Week 2, Session 6, Grade 5/6 Literacy
// Reading: Character analysis across 4 chapters (Jem's fear, Jem's death, Thomas/Millie, Tom's arrest)
// Writing: Appositives (revise concept, identify, add to sentences)
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
  addTextOnShape, withReveal,
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
const FOOTER = "Chapters 11-14 | Session 6 | Week 2 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Session6_Escape_and_Arrest";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const WORKSHEET_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Appositives Worksheet",
  "Student worksheet: identify and add appositives to sentences using novel content."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Answer Key",
  "Teacher reference: model answers for appositive sentences with marking guidance."
);
const RESOURCE_ITEMS = [WORKSHEET_RESOURCE, ANSWER_KEY_RESOURCE];
const WORKSHEET_PDF_PATH = path.join(OUT_DIR, WORKSHEET_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to Session 6. We are continuing with our novel today, picking up at Chapter 11
- We have four chapters to read today -- 11, 12, 13 and 14. There is a lot happening for Tom in these chapters
- Our writing focus today is appositives -- adding extra detail about a noun in a sentence

DO:
- Display title slide as students settle
- Have copies of the novel on desks, bookmarked at Chapter 11

TEACHER NOTES:
Session 6 of 10. Chapters 11-14 cover a dramatic arc: Jem's second chimney incident, the escape plan, Jem's death, and Tom's arrest. Chapter 12 requires a sensitivity advisory. The appositive writing connects to the reading objective of expanding description.

WATCH FOR:
- Students who may need a brief recap of Chapters 9-10 -- Jem's secret stash, the escape plan, the boys stealing items
- Students still processing content from the previous week

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands today: reading Chapters 11-14 for character development, and learning to add appositives to our sentences
- Read the success criteria. Our first one is about identifying what characters reveal through their actions. Our second goes deeper into how the author builds character across chapters. Our third is the writing target -- appositives

DO:
- Choral read the LI, then the SCs
- Do not over-explain appositives yet -- the I Do covers this

TEACHER NOTES:
SC1 is the floor -- every student can point to evidence of what a character does or says and explain what it shows. SC2 is the core -- tracking how the author develops Tom across multiple chapters. SC3 is the writing application -- adding appositives to sentences.

WATCH FOR:
- Students who look blank at "appositives" -- reassure: "We will learn this together step by step"
- Students already familiar with appositives from prior learning -- they can support peers

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_HOARD = `SAY:
- Our first vocabulary word: hoard. Read it with me: hoard [students repeat]
- Hoard can be a noun or a verb. As a noun, it means a secret store of valuable things. As a verb, it means to collect and hide things away
- In our novel, the boys have been hoarding stolen items -- hiding them away for their escape plan
- Quick check: if someone has a hoard of chocolate, have they shared it or hidden it? [Hidden it]

DO:
- Display word, choral read, define, give example
- Cold Call 2 students: "What might someone hoard, and why?"
- Image discrimination: show 3 images -- a pile of coins hidden under a mattress, someone sharing food at a table, a treasure chest. Which shows a hoard? [Coins under mattress or treasure chest]

TEACHER NOTES:
Pre-teaching "hoard" before reading supports comprehension of the boys' escape preparations. The word recurs when Tom steals coins in Chapter 12.

WATCH FOR:
- Students who confuse "hoard" with "horde" (a large group) -- clarify spelling and meaning
- Students who only think of "hoard" as negative -- expand: squirrels hoard nuts, people hoard memories

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_VOCAB_OCCUPIED = `SAY:
- Second word: occupied. Read it with me: occupied [students repeat]
- Occupied means being used or taken up by someone or something. A room is occupied when someone is in it. A person is occupied when they are busy
- In our novel, Tom enters a bedroom and finds it occupied -- there is a sleeping, drunk patron in the bed. This is important because Tom is trying to steal from the room
- Turn and talk: tell your partner a time when you found something was occupied -- a seat on the bus, a bathroom, a computer [30 seconds]

DO:
- Display word, choral read, define, example
- Turn and Talk: 30 seconds, then Cold Call 2 pairs
- Thumbs Up/Down CFU: "If a classroom is empty, is it occupied?" [Thumbs down -- no]

TEACHER NOTES:
"Occupied" is versatile across contexts. Teaching both the physical (a room occupied by someone) and abstract (a person occupied with a task) meanings builds transfer.

WATCH FOR:
- Students who confuse "occupied" with "preoccupied" -- acknowledge the connection but distinguish: occupied = busy/in use; preoccupied = distracted by worry
- Students giving examples only about rooms -- push for the "busy" meaning too

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_VOCAB_RESOLVED = `SAY:
- Third word: resolved. Read it with me: resolved [students repeat]
- Resolved means firmly decided or determined. When you have resolved to do something, you have made up your mind and nothing will change it
- In our novel, after something terrible happens in Chapter 12, Tom resolves to still try to escape. Despite everything, he is determined
- Ask: What is the difference between "decided" and "resolved"? [Resolved is stronger -- it means you are absolutely certain and won't change your mind]

DO:
- Display word, choral read, define, example
- Cold Call: "Can you think of a time you resolved to do something difficult?" [Take 2-3 responses]
- Sentence completion: "After failing the test, she resolved to ___" [study harder, practise every night, ask for help]

TEACHER NOTES:
"Resolved" carries emotional weight in the novel -- Tom's determination despite grief. This word connects directly to character analysis in the pause points.

WATCH FOR:
- Students who think "resolved" only means "solved" (as in a problem) -- clarify: resolved has two meanings. Here it means determined, not solved
- Students who give weak examples -- push for genuine determination, not casual decisions

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_READING = `SAY:
- Chapters 11 through 14 today. A lot happens. I will read aloud and we will pause at key moments
- Focus: watch how Tom responds to the events around him. What do his actions and decisions reveal about his character?
- Before we start -- Chapter 12 contains some difficult content. A character we know dies. If this feels uncomfortable, that is okay. We will talk about it together
- Find Chapter 11 now

DO:
- Give students 30 seconds to find the chapter
- Read aloud -- teacher reads for these chapters given the sensitive content
- Plan for 3 pause points: p.57, p.64, p.67

SENSITIVITY ADVISORY:
- What it is: Chapter 12 -- Jem dies when a chimney collapses on him. This is a significant character death that students may find upsetting, particularly given the bond between Tom and Jem established in earlier chapters
- Framing language: "This chapter deals with something very sad. The author has included this because it was a real danger for climbing boys in this era. If you need a moment, that is completely fine"
- Watch for: Students who withdraw, become visibly upset, or who have experienced loss recently
- Protocol: If a student becomes distressed, acknowledge their feelings privately. Do not require them to continue reading. Follow the school's wellbeing referral process if needed. After reading, normalise the emotional response: "It is okay to feel sad about this. Good readers feel things when they read"

TEACHER NOTES:
Teacher reads aloud for these chapters to control pacing around the sensitive content in Chapter 12. The alternating pause points cover character analysis (p.57, p.67) and the Thomas/Tom connection (p.64). Mention the other pause points from the planning document verbally if time permits.

WATCH FOR:
- Students who disconnect during the reading -- gently re-engage with a quiet question
- Students who want to discuss Jem's death immediately -- allow brief acknowledgement, then redirect to the pause point discussion

[General: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause here. "It's an omen, thought Tom as he looked up and saw a flash of blue between the black"
- Ask: What does the author mean by "It's an omen"? What is Tom thinking? [Tom sees blue sky through the chimney and takes it as a sign -- a good sign that their escape plan will work. He is hopeful]
- Ask: What does this tell us about Tom's state of mind right before everything changes? [He is excited, hopeful, looking for signs that things will get better. This makes what happens next even more devastating]

DO:
- Display the quote, read aloud
- Give 10 seconds of think time
- Cold Call 3 students for interpretation
- Push for the connection between hope and what is about to happen

CFU CHECKPOINT:
Technique: Cold Call

Script:
- "What is an omen?" [A sign of something to come]
- "Is Tom seeing this as a good omen or a bad omen?" [Good -- the blue sky represents hope and freedom]
- "Knowing what happens next in this chapter, what effect does this create?" [Dramatic irony -- the reader feels the tragedy more because Tom was so hopeful]
- Scan for: students connecting Tom's hope to the coming tragedy

PROCEED (>=80%): Most students understand the omen and its emotional effect. Continue reading.
PIVOT (<80%): Most likely issue -- students take "omen" literally and miss the emotional setup. Reteach: "Tom is underground in a dark chimney. He sees one flash of blue sky. To him, that tiny piece of sky means everything -- it means tomorrow he will be free. The author puts this here on PURPOSE, right before something terrible happens. Why?" Re-check: "What emotion does the author want us to feel?"

TEACHER NOTES:
This pause point is about authorial intent and dramatic irony. The omen of blue sky comes just before Jem's death -- the author deliberately creates hope to make the loss more impactful.

WATCH FOR:
- Students who define "omen" but cannot explain why the author placed it here -- push: "Why does the author show us Tom's hope RIGHT BEFORE the tragedy?"
- Students who are already anticipating Jem's death from the chapter summary -- redirect to the text evidence

[General: Pause Point 1 | VTLM 2.0: Higher-Order Questioning]`;

const NOTES_PAUSE2 = `SAY:
- Pause here. "...it could be any one of many"
- Context: Thomas's great-granddaughter Millie has visited, and Thomas is thinking about which ghost from his past is haunting him
- Ask: Given what the author has already told us about Tom and Thomas, why might Thomas be particularly upset by Miss Hildegard putting Millie in a cupboard? [The cupboard is a small, dark, enclosed space -- like a chimney. Thomas was a climbing boy. He knows what it is like to be trapped in dark spaces. Putting a child in a cupboard would bring back those memories]
- Ask: Why does the author say "it could be any one of many"? What does this tell us about Thomas's past? [He has so many painful memories that he cannot pinpoint which one is haunting him. His past was full of suffering]

DO:
- Display the quote
- Use Think-Pair-Share: 20 seconds think, 30 seconds pair, then share
- Draw the connection between Tom's childhood experiences and Thomas's reaction as an old man

CFU CHECKPOINT:
Technique: Think-Pair-Share

Script:
- "Think for 20 seconds: why would a cupboard upset Thomas so much?"
- "Share with your partner for 30 seconds"
- "Hands down -- I am selecting pairs. [Name], what did you discuss?"
- Scan for: students connecting the cupboard to Tom's chimney experiences

PROCEED (>=80%): Most pairs connect the enclosed space to Tom's past. Continue reading.
PIVOT (<80%): Most likely issue -- students see the cupboard as generic punishment and miss the chimney connection. Reteach: "Think about what Tom did as a boy. He was sent up chimneys -- tiny, dark, suffocating spaces. Now imagine he is old, and someone puts his great-granddaughter in a cupboard. A small, dark space. What would that remind him of?" Re-check: "Why would Thomas react more strongly than other adults?"

TEACHER NOTES:
This pause point connects the two timelines. The author uses Millie's experience to trigger Thomas's memories, showing how childhood trauma persists into old age. This is sophisticated character development across the dual narrative.

WATCH FOR:
- Students who only see surface action (Millie in a cupboard) without connecting to Tom's past -- the chimney/cupboard parallel is the key insight
- Students who are making strong connections between the timelines -- excellent: this is the analytical depth SC2 targets

[General: Pause Point 2 | VTLM 2.0: Teacher-Led Discussion]`;

const NOTES_PAUSE3 = `SAY:
- Final pause. "Come on now, boy. Yer nicked"
- Ask: What is going on here? [The Bow Street Runners have arrived. They search the boys and find the money Tom stole. Tom is being arrested]
- Ask: Think about everything that has happened in these four chapters. Tom planned to escape. His best friend died. He resolved to keep going. And now he is arrested. What does this tell us about Tom's situation? [No matter what Tom does, he cannot escape his circumstances. Every plan fails. The author keeps building hope and then crushing it]

DO:
- Display the quote
- Let the moment land -- this is the culmination of four chapters of drama
- Cold Call 3-4 students for their analysis
- Draw the thread: hope (omen) -> tragedy (Jem) -> determination (resolved) -> arrest (nicked)

TEACHER NOTES:
The final pause point ties together the arc across all four chapters. The pattern of hope followed by disaster is a deliberate authorial choice that students should recognise.

WATCH FOR:
- Students who see each event in isolation -- help them see the pattern: the author repeatedly gives Tom hope and then takes it away
- Students who are emotionally affected by the cumulative weight of the chapters -- acknowledge this as a sign of engaged reading

[General: Pause Point 3 | VTLM 2.0: Deep Comprehension]`;

const NOTES_PAUSE3_REVEAL = `SAY:
- The author creates a pattern across these chapters: hope, then loss. Omen of blue sky, then Jem's death. Tom resolves to escape, then he is arrested
- This pattern is deliberate. The author wants us to understand that for children like Tom, the system is rigged against them. No matter how brave or determined they are, the odds are stacked
- This connects to our vocabulary word "resolved" -- Tom's resolve is tested again and again

DO:
- Reveal the summary card
- Draw the connection back to the vocabulary
- Transition to writing: "Now let us look at how we can add detail to our sentences using appositives"

[General: Pause Point 3 Reveal | VTLM 2.0: Deep Comprehension]`;

const NOTES_IDO_APPOSITIVE = `SAY:
- Now we are working on sentence-level writing: appositives
- An appositive is a noun or noun phrase placed right next to another noun to rename it or give more information about it
- Think of it as a label that sits beside a name to tell the reader more
- Watch me. Here is a sentence: "Tom climbed into the chimney"
- I want to add more information about Tom. I am going to add an appositive: "Tom, the youngest of Master Jack's boys, climbed into the chimney"
- Notice: "the youngest of Master Jack's boys" sits right next to "Tom" and tells us more about who Tom is. That is the appositive
- Notice the commas. The appositive is set off by commas because it is extra information. If I removed it, the sentence still works: "Tom climbed into the chimney"

DO:
- Display the terminology and model sentence clearly
- Point to the appositive and the noun it renames
- Think aloud: "I check -- if I remove the appositive, does the sentence still make sense? Yes. So the commas are correct and the appositive is extra detail"
- Show a second quick example: "Jem, Tom's closest friend, whispered his escape plan"

TEACHER NOTES:
The think-aloud about removing the appositive to check the sentence is the key metacognitive move. Students need this self-checking strategy when writing their own.

MISCONCEPTIONS:
- Misconception: An appositive is the same as an adjective
  Why: Both add description, so students conflate them
  Impact: Students write adjectives when asked for appositives, or use adjective punctuation
  Quick correction: "An adjective describes -- like 'brave Tom'. An appositive RENAMES -- like 'Tom, the youngest boy'. The appositive is a noun phrase, not an adjective"

WATCH FOR:
- Students who look confused by the commas -- emphasise: "The commas are like brackets. They hold the extra information"
- Students who mix up appositives with relative clauses (from Session 2) -- clarify: appositives have no connecting word like "who" or "which"

[General: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU = `SAY:
- Quick check. I am going to show you a sentence. You need to find the appositive
- "Master Jack, a cruel and violent man, sent the boys up the chimney"
- On your whiteboards, write the appositive. You have 15 seconds
- Three, two, one -- hold up! [Scan]

DO:
- Use Show Me Boards
- Scan for: "a cruel and violent man"
- Students who write "Master Jack" have identified the noun, not the appositive
- Students who write the whole sentence need redirection

CFU CHECKPOINT:
Technique: Show Me Boards

Script:
- "Write ONLY the appositive on your whiteboard. 15 seconds"
- Count down, scan
- Look for: "a cruel and violent man" or close paraphrases

PROCEED (>=80%): Most boards show the correct appositive. Move to We Do.
PIVOT (<80%): Most likely issue -- students write the subject noun or the whole sentence instead of isolating the appositive. Reteach: "The appositive sits between the commas. Find the two commas first. What is between them? That is the appositive. It renames the noun that comes right before the first comma." Re-check with a new sentence: "Jem, Tom's closest friend, whispered the plan." What is the appositive?

TEACHER NOTES:
Show Me Boards provide visible evidence from every student. The "find the commas" strategy gives students a concrete tool for identifying appositives.

WATCH FOR:
- Students who write "Master Jack" -- they have found the noun, not the appositive. Redirect: "That is WHO the appositive is about. What does it TELL US about Master Jack?"
- Students who correctly identify the appositive quickly -- they are ready for the We Do

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- The appositive is "a cruel and violent man". It sits right next to "Master Jack" and tells us more about who he is
- Notice: if I remove it, the sentence still works -- "Master Jack sent the boys up the chimney." The appositive is the extra detail between the commas
- Now you will practise adding appositives to sentences on your worksheets

DO:
- Reveal the answer
- Highlight the commas as the signal
- Distribute the Session 6 Appositives Worksheet
- Transition to We Do

[General: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_WEDO = `SAY:
- On your worksheet, you have sentences from our novel. Your job is to add an appositive to each sentence
- Section A gives you a bank of appositives to choose from. Section B -- you write your own
- Let's do the first one together. "Tom stole the coins from the bedroom." Who can suggest an appositive for Tom? [Take 2-3 suggestions -- e.g., "a desperate boy with nothing to lose", "the smallest of the climbing boys"]
- Good. Write it in: "Tom, [your appositive], stole the coins from the bedroom." Remember the commas
- Now complete the rest. Section A first, then Section B. You have 8 minutes

DO:
- Distribute worksheets (if not already done)
- Do sentence 1 collaboratively (We Do), then release for You Do on the rest
- Circulate -- check commas and that appositives are noun phrases, not adjectives or clauses

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide a word bank of descriptive noun phrases (e.g., "a young chimney sweep", "Tom's best friend", "the owner of the boys") and sentence frames with blanks clearly marked. Students select from the bank and write the appositive with commas
- Extra Notes: Students can refer back to the I Do model on screen

EXTENDING PROMPT:
- Task: Write 2 original sentences about events from Chapters 11-14 that each contain an appositive. Then write a short paragraph (3-4 sentences) about one character where at least one sentence uses an appositive to add detail

TEACHER NOTES:
The transition from We Do (sentence 1) to You Do (remaining sentences) within the same worksheet is a clean gradual release. Section A (with bank) scaffolds; Section B (without bank) is the independent application.

WATCH FOR:
- Students who write adjectives instead of noun phrases -- redirect: "An appositive is a noun phrase. 'Brave' is an adjective. 'A brave young boy' is a noun phrase"
- Students who forget commas -- remind: "The commas hold the appositive in place. Without them, the sentence is confusing"
- Students finishing Section A quickly and correctly -- they are ready for Section B

[General: We Do / You Do | VTLM 2.0: Guided + Independent Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Let's check our answers. Here are model appositives for each sentence
- Your answers do not need to match mine exactly. What matters is that your appositive is a noun phrase that renames the noun, and it is set off by commas
- For sentence 1: "Tom, a desperate boy planning his escape, stole the coins from the bedroom"
- For sentence 2: "Jem, Tom's closest friend and fellow climbing boy, got stuck in the chimney again"
- For sentence 3: "Master Jack, the cruel chimney sweep master, whipped the boys as punishment"

DO:
- Reveal model answers
- Read each full sentence aloud with the appositive
- Invite students to share alternatives that also work
- Celebrate creative, accurate appositives

[General: We Do / You Do Reveal | VTLM 2.0: Guided + Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: identifying what characters reveal through actions and words -- thumbs? [scan]
- SC2: explaining how the author develops Tom's character across chapters -- thumbs? [scan]
- SC3: adding an appositive to a sentence -- thumbs? [scan]
- Turn and talk: Which moment across Chapters 11 to 14 was the biggest turning point for Tom, and why?

DO:
- Run through each SC with thumbs check
- The turn-and-talk targets the reading analysis -- students should identify a specific moment and justify it
- Preview: "Next session we continue with Chapters 15 to 17 -- Tom faces the consequences of his arrest"

TEACHER NOTES:
The closing reconnects to the reading arc before ending. Students who show "thumbs down" on SC3 may benefit from reviewing the worksheet at home or receiving additional practice next session.

WATCH FOR:
- Students "thumbs down" on SC3 -- check if it is terminology confusion or genuine difficulty with the concept
- Students who want to discuss Jem's death further -- acknowledge the emotion and encourage them to write about it

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two printable resources today
- The ${WORKSHEET_RESOURCE.name} is for the We Do and You Do appositive activity
- The ${ANSWER_KEY_RESOURCE.name} is for teacher reference

DO:
- Print the worksheet before the lesson (one per student)
- Print the answer key (teacher copy only)
- Click any resource card to open the PDF

TEACHER NOTES:
The worksheet progresses from identifying appositives to adding them with a bank, then without a bank. The answer key includes alternative valid responses.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Chapters 11-14: Escape and Arrest -- Session 6";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "Chapters 11-14",
    "Escape and Arrest",
    "Session 6  |  Week 2  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // =========================================================================
  // SLIDE 2 -- Learning Intention & Success Criteria
  // =========================================================================
  liSlide(
    pres,
    [
      "We are learning to analyse how an author develops character across key moments, and to add detail to sentences using appositives",
    ],
    [
      "I can identify what a character's actions and words reveal about their feelings",
      "I can explain how the author uses events across chapters to develop character",
      "I can add an appositive to a sentence to give the reader more information about a noun",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Vocabulary: hoard
  // =========================================================================
  vocabSlide(
    pres,
    "hoard",
    "noun / verb",
    "A secret store of valuable things (noun); to collect and hide things away secretly (verb). People hoard things they want to protect or save for later.",
    "The boys had been hoarding stolen items in a secret place, saving everything they could for their planned escape.",
    NOTES_VOCAB_HOARD,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Vocabulary: occupied
  // =========================================================================
  vocabSlide(
    pres,
    "occupied",
    "adjective",
    "Being used or taken up by someone or something. A room is occupied when someone is in it. A person is occupied when they are busy with something.",
    "Tom crept into the bedroom only to discover it was occupied by a sleeping, drunk patron.",
    NOTES_VOCAB_OCCUPIED,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- Vocabulary: resolved
  // =========================================================================
  vocabSlide(
    pres,
    "resolved",
    "adjective / verb",
    "Firmly decided or determined. When you have resolved to do something, you have made up your mind and nothing will change it. Stronger than simply 'decided'.",
    "Despite the tragedy, Tom resolved to still try to escape the next day.",
    NOTES_VOCAB_RESOLVED,
    FOOTER
  );

  // =========================================================================
  // SLIDE 6 -- Reading Launch
  // =========================================================================
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
      "Focus: How does Tom respond to each event? What does this reveal about his character?",
    ],
    NOTES_READING,
    FOOTER
  );

  // =========================================================================
  // SLIDE 7 -- Pause Point 1: "It's an omen"
  // =========================================================================
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

  // =========================================================================
  // SLIDE 8 -- Pause Point 2: Thomas and Millie
  // =========================================================================
  quoteSlide(
    pres,
    "Pause Point 2",
    "Chapter 13 -- p. 64",
    "...it could be any one of many.",
    "p. 64",
    "Given what the author has told us about Tom's childhood, why might Thomas be particularly upset by Miss Hildegard putting Millie in a cupboard?",
    NOTES_PAUSE2,
    FOOTER
  );

  // =========================================================================
  // SLIDES 9-10 -- Pause Point 3 with Reveal: "Yer nicked"
  // =========================================================================
  withReveal(
    () => quoteSlide(
      pres,
      "Pause Point 3",
      "Chapter 14 -- p. 67",
      "'Come on now, boy. Yer nicked.'",
      "p. 67",
      "What is going on here? Think about everything that has happened across these four chapters. What pattern is the author creating?",
      NOTES_PAUSE3,
      FOOTER
    ),
    (slide) => {
      const ansY = 3.62;
      slide.addShape("roundRect", {
        x: 0.5, y: ansY, w: 9, h: 1.38, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      });
      slide.addText("Hope and Loss", {
        x: 0.75, y: ansY + 0.08, w: 3, h: 0.30,
        fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true, margin: 0,
      });
      slide.addText("The author creates a deliberate pattern: hope followed by disaster. Blue sky omen -> Jem's death. Tom resolves to escape -> arrested. For children like Tom, the system is stacked against them. Every plan crumbles.", {
        x: 0.75, y: ansY + 0.40, w: 8.5, h: 0.88,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, valign: "top", margin: 0,
      });
      slide.addNotes(NOTES_PAUSE3_REVEAL);
    }
  );

  // =========================================================================
  // SLIDE 11 -- I Do: Appositives
  // =========================================================================
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Appositives",
    "An appositive is a noun or noun phrase placed next to another noun to rename it or give more information.\n\nIt is set off by commas.\n\nIf you remove it, the sentence still makes sense.\n\nAppositives are NOT:\n- Adjectives (brave, small)\n- Relative clauses (who was brave)",
    "Model:\n\n\"Tom climbed into the chimney.\"\n\nAdd an appositive about Tom:\n\n\"Tom, the youngest of Master Jack's boys, climbed into the chimney.\"\n\nCheck: Remove it. \"Tom climbed into the chimney.\" Still works. The appositive is extra detail.",
    NOTES_IDO_APPOSITIVE,
    FOOTER
  );

  // =========================================================================
  // SLIDES 12-13 -- CFU: Find the Appositive (withReveal)
  // =========================================================================
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Find the Appositive",
      "Show Me Boards",
      "\"Master Jack, a cruel and violent man, sent the boys up the chimney.\"\n\nOn your whiteboard, write ONLY the appositive.\n\nHint: Find the two commas first. What is between them?",
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
        x: 0.7, y: ansY + 0.10, w: 2.6, h: 0.28, rectRadius: 0.08,
        fill: { color: C.WHITE },
      });
      slide.addText("a cruel and violent man", {
        x: 0.7, y: ansY + 0.10, w: 2.6, h: 0.28,
        fontSize: 12, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addText("The appositive renames \"Master Jack\" -- it tells us more about who he is. Remove it and the sentence still works: \"Master Jack sent the boys up the chimney.\"", {
        x: 3.5, y: ansY + 0.08, w: 5.8, h: 0.84,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // =========================================================================
  // SLIDES 14-15 -- We Do / You Do: Add Appositives (withReveal)
  // =========================================================================
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
        s.addText("Add an appositive after the underlined noun. Use commas.", {
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

  // =========================================================================
  // SLIDE 16 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    "Which moment across Chapters 11 to 14 was the biggest turning point for Tom, and why? Tell your partner.",
    [
      "I can identify what a character's actions and words reveal about their feelings",
      "I can explain how the author uses events across chapters to develop character",
      "I can add an appositive to a sentence to give the reader more information about a noun",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 17 -- Resources
  // =========================================================================


  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: Appositives Worksheet -----------------------------------------
  const ws = createPdf({ title: WORKSHEET_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Appositives -- Add Detail to Sentences", {
    color: C.PRIMARY,
    subtitle: "Chapters 11-14: Escape and Arrest",
    lessonInfo: "Session 6 | Week 2 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "An appositive is a noun or noun phrase placed next to another noun to rename it or give more detail. It is set off by commas. Example: \"Tom, the youngest of the climbing boys, looked up at the chimney.\"", wsY, { color: C.PRIMARY });

  // Section A: Add appositives with a bank
  wsY = addSectionHeading(ws, "Section A: Add an Appositive (use the bank)", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Choose an appositive from the bank and add it to each sentence. Remember to use commas.", wsY);
  wsY += 4;
  wsY = addTipBox(ws, "Appositive Bank: a desperate boy with nothing to lose  |  Tom's closest friend and fellow climbing boy  |  the cruel chimney sweep master  |  a sleeping, drunk patron", wsY, { color: C.SECONDARY });

  wsY = addBodyText(ws, "1. Tom stole the coins from the bedroom.", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "2. Jem got stuck in the chimney again.", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "3. Master Jack whipped the boys as punishment.", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "4. The man in the bedroom did not wake up.", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 10;

  // Section B: Add appositives without a bank
  wsY = addSectionHeading(ws, "Section B: Add an Appositive (your own words)", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "Write your own appositive to add to each sentence. Remember to use commas.", wsY);
  wsY += 4;

  wsY = addBodyText(ws, "5. Tom looked up at the sky through the chimney.", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "6. The Bow Street Runners searched the boys.", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "7. Millie came to visit Thomas on his birthday.", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });

  addPdfFooter(ws, "Session 6 | Appositives Worksheet");

  // --- PDF 2: Answer Key ---------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Appositives -- Answer Key", {
    color: C.ALERT,
    subtitle: "Teacher Reference -- Chapters 11-14",
    lessonInfo: "Session 6 | Week 2 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Accept any completion where the appositive is a noun phrase that renames or adds detail about the noun, and is correctly set off by commas. Model answers and alternatives below.", akY, { color: C.ALERT });

  akY = addSectionHeading(ak, "Section A: With Bank", akY, { color: C.PRIMARY });

  akY = addBodyText(ak, "1. \"Tom, a desperate boy with nothing to lose, stole the coins from the bedroom.\"", akY);
  akY = addBodyText(ak, "2. \"Jem, Tom's closest friend and fellow climbing boy, got stuck in the chimney again.\"", akY);
  akY = addBodyText(ak, "3. \"Master Jack, the cruel chimney sweep master, whipped the boys as punishment.\"", akY);
  akY = addBodyText(ak, "4. \"The man in the bedroom, a sleeping, drunk patron, did not wake up.\"", akY);
  akY += 10;

  akY = addSectionHeading(ak, "Section B: Own Words (accept any valid noun phrase)", akY, { color: C.ACCENT });

  akY = addBodyText(ak, "5. Model: \"Tom, a small boy covered in soot, looked up at the sky through the chimney.\"", akY);
  akY = addBodyText(ak, "   Alternative: \"Tom, the bravest of the climbing boys, looked up at the sky...\"", akY, { italic: true });
  akY += 4;
  akY = addBodyText(ak, "6. Model: \"The Bow Street Runners, officers of the law, searched the boys.\"", akY);
  akY = addBodyText(ak, "   Alternative: \"The Bow Street Runners, two stern-looking men, searched the boys.\"", akY, { italic: true });
  akY += 4;
  akY = addBodyText(ak, "7. Model: \"Millie, Thomas's great-granddaughter, came to visit Thomas on his birthday.\"", akY);
  akY = addBodyText(ak, "   Alternative: \"Millie, a cheerful young girl, came to visit Thomas on his birthday.\"", akY, { italic: true });
  akY += 14;

  akY = addSectionHeading(ak, "Common Errors to Watch For", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- Adjectives instead of noun phrases (e.g., \"Tom, brave, stole...\" -- missing the noun)", akY);
  akY = addBodyText(ak, "- Missing commas around the appositive", akY);
  akY = addBodyText(ak, "- Relative clauses instead of appositives (e.g., \"Tom, who was brave, stole...\" -- this uses 'who' and is a different structure)", akY);
  akY = addBodyText(ak, "- Appositives that do not make sense with the noun (e.g., adding a Tom appositive to a sentence about Master Jack)", akY);

  addPdfFooter(ak, "Session 6 | Answer Key -- TEACHER COPY");

  // --- Write all files ----------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Session6.pptx` }),
    writePdf(ws, WORKSHEET_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Session6.pptx`);
  console.log("Done: " + WORKSHEET_RESOURCE.name);
  console.log("Done: " + ANSWER_KEY_RESOURCE.name);
}

build().catch(console.error);
