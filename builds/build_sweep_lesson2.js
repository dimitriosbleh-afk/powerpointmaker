"use strict";

// Sweep Unit -- Lesson 2: Chapters 4-6 + Relative Clauses (Identify)
// Week 1, Session 2, Grade 5/6 Literacy
// Novel: "Sweep" by Jonathan Auxier

const pptxgen = require("pptxgenjs");
const path = require("path");
const fs = require("fs");

const { createTheme, weekToVariant } = require("../themes/factory");
const T = createTheme("literacy", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  addTextOnShape, addInstructionCard, withReveal,
  titleSlide, liSlide, contentSlide,
  cfuSlide, closingSlide,
  vocabSlide, quoteSlide, modellingSlide,
} = T;

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  addWriteLine,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 2;
const FOOTER = "Sweep | Lesson 2 | Week 1 | Year 5/6 Literacy";
const OUT_DIR = "output/Sweep_Lesson2_Chapters4_6";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const WORKSHEET_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Relative Clauses Worksheet",
  "Student worksheet: identify and punctuate relative clauses in sentences from the novel."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Answer Key",
  "Teacher reference: model answers for relative clause identification."
);
const RESOURCE_ITEMS = [WORKSHEET_RESOURCE, ANSWER_KEY_RESOURCE];
const WORKSHEET_PDF_PATH = path.join(OUT_DIR, WORKSHEET_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to Lesson 2. Today we read Chapters 4 to 6 of Sweep
- Something interesting happens with the structure in these chapters -- the author moves between time periods again
- Our writing focus today is relative clauses. This is a new grammar concept that will help you write more complex, detailed sentences

DO:
- Display title slide as students settle
- Have copies of Sweep on desks, bookmarked at Chapter 4

TEACHER NOTES:
Session 2 of the Sweep unit. Chapters 4-6 introduce Master Jack and the chimney sweep world (Tom's story) and flash forward to old Thomas (Chapter 5). The writing focus is identifying relative clauses -- the first step before constructing them. This is a brand new grammar concept for this term.

WATCH FOR:
- Students eager to find out what happens to Tom after the workhouse -- good engagement
- Students who remember the courage cloak metaphor from Lesson 1 -- it does not reappear in these chapters but will return

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands again: reading Chapters 4-6 and learning about relative clauses
- Read the success criteria. SC1 is about understanding how the author switches between time periods. SC2 is about character development -- what we learn about characters from how they are described. SC3 is our writing focus -- identifying relative clauses in sentences

DO:
- Choral read the LI, then each SC
- Briefly preview: "A relative clause is a group of words that gives extra information about a noun. We will learn to spot them today"

TEACHER NOTES:
SC1 targets the dual-timeline structure (young Tom / old Thomas). SC2 targets character analysis through authorial description. SC3 introduces relative clause identification -- the prerequisite for constructing them in future lessons. SC1 is the floor: noticing the time shift. SC2 extends to explaining purpose. SC3 is the new grammar skill.

WATCH FOR:
- Students unsure what "relative clause" means -- reassure: "This is completely new. We will learn it together step by step"
- Students who recall the time shift from Lesson 1 (Ch1 vs Ch2) -- excellent connection

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_OVERPOWERING = `SAY:
- First vocabulary word: overpowering. Read it with me: overpowering [students repeat]
- Overpowering means so strong that you cannot resist it or ignore it. It overwhelms your senses or your ability to cope
- In the novel, the smell in Master Jack's cellar is overpowering. It is so strong it hits you the moment you walk in -- soot, unwashed bodies, damp
- Ask: What is the difference between a strong smell and an overpowering smell? [A strong smell you notice. An overpowering smell you cannot escape -- it fills your nose and you cannot think about anything else. It takes over]

DO:
- Display word, choral read, define, give example
- Cold Call 2 students: name something that could be overpowering (a smell, a feeling, a sound)
- Sentence completion: "The ___ was so overpowering that ___"

TEACHER NOTES:
"Overpowering" connects to the sensory detail the author uses to describe Tom's new environment with Master Jack. The word reinforces how extreme and hostile the conditions are. It also works in emotional contexts (overpowering grief, overpowering fear) which connects to earlier chapters.

WATCH FOR:
- Students who only think of smell -- extend: "Can a feeling be overpowering? Can noise be overpowering?"
- Students who give strong examples -- affirm and connect to the novel

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_VOCAB_DEDICATION = `SAY:
- Second word: dedication. Read it with me: dedication [students repeat]
- Dedication means giving a lot of time and effort to something because you believe it matters. It shows commitment and loyalty
- In Chapter 5, Thomas reflects on his family with dedication -- he has spent his life building something for them. His dedication to his farm and his family defines who he is
- Ask: What is the difference between doing something and doing something with dedication? [Dedication means you keep going even when it is hard. You care deeply. It is not just effort -- it is effort driven by belief and purpose]

DO:
- Display word, choral read, define, example
- Turn and Talk: "Think of someone you know who shows dedication. What do they do? Why is it dedication and not just effort?"
- Take 2 responses

TEACHER NOTES:
"Dedication" connects to old Thomas's character and the thematic contrast between his settled, purposeful life and young Tom's unstable, powerless existence. The word is also useful in academic writing and information reports -- a Tier 2 word with broad application.

WATCH FOR:
- Students who confuse dedication with stubbornness -- "Dedication is positive. You are committed to something worthwhile. Stubbornness can be about refusing to change even when you should"
- Students who connect dedication to sporting or artistic pursuits -- excellent real-world examples

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_READING = `SAY:
- Chapters 4 to 6 today. Three chapters with very different moods
- Chapter 4: Tom arrives at the workhouse and is handed off. Something unexpected happens -- he is sold to a man called Master Jack
- Chapter 5: we jump forward again to old Thomas. His son Joshua visits. Thomas reflects on his life
- Chapter 6: back to young Tom. He arrives somewhere new and meets other boys
- Reading mode: student read aloud. We have four pause points
- Watch for the author's structure -- she keeps switching between young Tom and old Thomas. Think about WHY

DO:
- Give students 30 seconds to find Chapter 4
- Select first reader
- Plan reader rotations -- Chapter 5 is shorter, good for a less confident reader

TEACHER NOTES:
The dual timeline becomes clearer in these chapters. Ch5 (Thomas) is a deliberate contrast to Ch4 and Ch6 (Tom). The author places a quiet, reflective chapter between two tense, action-driven chapters. Students should begin to notice this structural choice.

WATCH FOR:
- Students who are confused by the time jump in Ch5 -- "The author is doing this on purpose. She wants us to see both versions of this life"
- Students who react to Tom being sold -- this is a confronting concept. Acknowledge it

[General: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause. "'That way,' she ordered. 'And mind your manners.'" Page 15
- What have we learned about Tom's situation from this conversation? [He has arrived at the workhouse. The woman is not welcoming -- she gives orders. Tom is treated like an inconvenience, not a child who needs care. "Mind your manners" suggests he must be obedient and quiet]
- Ask: How does the way this woman speaks to Tom compare to how Mrs Wilson treated him at the end of Chapter 3? [Mrs Wilson gave him breakfast -- some kindness. This woman gives him orders. The contrast shows that kindness is rare in Tom's world]

DO:
- Display the quote
- Cold Call 3 students
- Push for textual evidence in responses

TEACHER NOTES:
The workhouse scene establishes the institutional indifference Tom faces. The woman's dialogue reveals character through speech -- a key authorial technique for SC2. The contrast with Mrs Wilson is worth noting for character analysis.

WATCH FOR:
- Students who notice the power dynamic -- adults command, Tom obeys
- Students who compare this to modern treatment of children -- valid connection

[General: Pause Point 1 | VTLM 2.0: Active Reading]`;

const NOTES_PAUSE2 = `SAY:
- Pause. "Good day to you, milord." Page 17
- What is going on here? [Tom is being sold. A man -- Master Jack -- has come to buy a boy to work as a chimney sweep. The language "milord" is sarcastic or mocking -- Master Jack is no lord]
- Ask: Why does the author include this moment? What does it tell us about this world? [Children were bought and sold like goods. The casual, almost business-like nature of it shows how normalised this cruelty was. The author wants us to feel the injustice]

DO:
- Display the quote
- Think-Pair-Share: 30 seconds think, 30 seconds share
- Emphasise: this is based on real historical practices

CFU CHECKPOINT:
Technique: Think-Pair-Share

Script:
- "Think for 30 seconds: what does this scene reveal about how children like Tom were treated in this time period?"
- "Share with your partner"
- "I am selecting pairs. [Name], what did you discuss?"
- Scan for: students connecting the selling of children to broader themes of power and injustice

PROCEED (>=80%): Most pairs identify the injustice and connect it to the historical context. Continue reading.
PIVOT (<80%): Most likely issue -- students describe WHAT happens but not WHY the author included it. Reteach: "The author chose to show us this scene. She could have just said 'Tom was taken away.' Instead she showed us the transaction -- the buying and selling. Why show it? Because she wants us to FEEL the injustice, not just know about it." Re-check: "What emotion is the author trying to create in you right now?"

TEACHER NOTES:
This moment connects to the cross-curricular HASS content that will appear in Lessons 4-5 (convict transportation, treatment of children in 18th century England). Flag it lightly now without going deep into the historical context.

WATCH FOR:
- Students who are shocked by the concept of buying children -- validate: "It is shocking. That is exactly the reaction the author wants"
- Students who connect to other stories about child labour -- good intertextual thinking

[General: Pause Point 2 | VTLM 2.0: Deep Comprehension]`;

const NOTES_PAUSE3 = `SAY:
- Pause. "Someone was in the rose garden, on the other side of the house." Page 21
- The author has switched to old Thomas again. Why might she be doing this -- changing between times in Tom's life in the different chapters?
- Ask: What have we learned about Tom/Thomas from this flash-forward? [Thomas is old, settled, has a farm, has sons. He is reflective and thoughtful. His life turned out well -- but we do not yet know how he got from chimney sweep boy to elderly farmer in Australia]
- The author is creating a mystery: how does this frightened boy become this old man?

DO:
- Display the quote
- Cold Call 2 students
- Draw attention to the CONTRAST: young Tom in a dirty cellar vs old Thomas in a rose garden

TEACHER NOTES:
Chapter 5 is structurally important. The author places a calm, reflective chapter between two tense chapters to create contrast and build the mystery of Tom's journey. The rose garden vs the cellar is a powerful visual contrast. This connects to SC1 (understanding the dual timeline structure).

WATCH FOR:
- Students who predict a connection between the ghost (Ch1) and Tom's story -- strong inferencing
- Students who notice the rose garden as a symbol of beauty/peace vs Tom's harsh reality

[General: Pause Point 3 | VTLM 2.0: Structural Analysis]`;

const NOTES_PAUSE4 = `SAY:
- Final pause. "Let's you say hello to 'em then, shall we?" Page 26
- What do you think the author wants us to know about Tom's situation now? [Tom has been taken to Master Jack's place. There are other boys already there -- Big Bill, Little Will, and Jem. Tom is about to enter a world of chimney sweeping. The casual, rough way Master Jack speaks shows he does not care about these boys]
- Ask: The personification on this page -- "His feet already screamed with cold." What does this add? [It makes the physical suffering vivid and immediate. Feet cannot literally scream, but the word "screamed" tells us the cold is not just uncomfortable -- it is painful, urgent, desperate]

DO:
- Display the quote
- Finger voting: "How safe do you think Tom is right now? 1 = very safe, 5 = in danger" [Expect 4s and 5s]
- Take 2-3 responses

TEACHER NOTES:
This closes the reading. Tom is now embedded in Master Jack's world with the other boys. The personification of the cold is a strong example of authorial technique for SC2. The introduction of Big Bill, Little Will, and Jem sets up important relationships for future chapters.

WATCH FOR:
- Students who notice the other boys' names (Big Bill, Little Will) -- the size-based names suggest their physical condition and role
- Students who predict Tom will befriend Jem -- good narrative anticipation

[General: Pause Point 4 | VTLM 2.0: Critical Thinking]`;

const NOTES_RC_IDO = `SAY:
- Now our writing focus. Relative clauses -- this is brand new, so we will build it step by step
- A relative clause is a group of words that gives EXTRA INFORMATION about a noun. It starts with a relative pronoun: who, whom, which, whose, that. Or a relative adverb: when, where, why
- Watch: "Tom, who had lost his father, was sent to the workhouse." The relative clause is "who had lost his father." It tells us more about Tom
- The relative clause sits right after the noun it describes. It works like an adjective -- it adds detail
- Now, essential vs non-essential. If the relative clause is essential for meaning, no commas. If it adds extra information that could be removed, use commas
- Essential: "The boy that Master Jack chose was Tom." Remove "that Master Jack chose" and you lose the meaning -- which boy?
- Non-essential: "Tom, who was only ten years old, followed Master Jack to the cart." Remove "who was only ten years old" and the sentence still makes sense

DO:
- Display the relative pronoun/adverb list
- Display both example sentences
- Point to the relative clause in each, underline it
- Circle the relative pronoun
- Show where commas go for non-essential clauses

MISCONCEPTIONS:
- Misconception: Every clause starting with "who" or "which" is a relative clause
  Why: Students overgeneralise the relative pronoun signal. Sometimes "who" starts a question ("Who did this?") not a relative clause
  Impact: Students misidentify question words as relative pronouns in analysis tasks
  Quick correction: "A relative clause gives extra information ABOUT a noun that comes before it. If 'who' starts a question, it is not a relative clause -- it is an interrogative. Check: is there a noun right before it that the clause describes?"

TEACHER NOTES:
This is the first exposure to relative clauses. The I Do focuses on identification and punctuation, not construction -- that comes in later lessons. All examples use Sweep content to maintain textual connection. The essential/non-essential distinction is important for punctuation accuracy.

WATCH FOR:
- Students who confuse relative clauses with other subordinate clauses -- "Look for the relative pronoun (who/which/that/whose/where/when). That is your signal"
- Students who are overwhelmed by the terminology -- reassure: "You already use these in your speaking. Today we are just learning to spot them and name them"

[General: I Do -- Relative Clauses | VTLM 2.0: Explicit Teaching]`;

const NOTES_WEDO = `SAY:
- Let's practise spotting relative clauses together. I will show you sentences from the novel
- "The woman, who showed no warmth towards Tom, pointed down the corridor."
- With your partner: find and underline the relative clause. Circle the relative pronoun. Does it need commas?
- [Take responses] The relative clause is "who showed no warmth towards Tom." The relative pronoun is "who." It IS non-essential -- you could remove it and the sentence still works. So yes, commas
- Next one: "The cart that carried Tom away from the workhouse rattled over the cobblestones."
- Find the relative clause. Does this one need commas?

DO:
- Display sentences one at a time
- Give 30 seconds partner discussion for each
- Take responses, model the underlining and circling
- Click to next slide to reveal answers after students respond

CFU CHECKPOINT:
Technique: Show Me Boards

Script:
- Display: "The boys who slept in the cellar huddled together for warmth."
- "On your whiteboard, underline the relative clause and write E for essential or NE for non-essential. 30 seconds. Show me"
- Scan for: correct identification of "who slept in the cellar" AND correct label (E -- essential, because removing it changes meaning: which boys?)

PROCEED (>=80%): Most boards show correct clause and E label. Move to You Do.
PIVOT (<80%): Most likely issue -- students label it NE because it has "who" and they think all "who" clauses need commas. Reteach: "Read the sentence without the clause: 'The boys huddled together for warmth.' Which boys? We do not know. The clause is ESSENTIAL because it tells us WHICH boys. If removing the clause makes the sentence unclear, it is essential -- no commas." Re-check with a new sentence: "Tom, who missed his father terribly, tried to sleep." E or NE? [NE -- we already know who Tom is]

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Identify the relative pronoun in each sentence only (who, which, that, whose, where, when). Do not worry about essential vs non-essential yet -- just find the signal word
- Extra Notes: Provide a reference card listing all relative pronouns and adverbs

EXTENDING PROMPT:
- Task: Write two sentences about the novel -- one with an essential relative clause and one with a non-essential relative clause. Punctuate correctly and label each

TEACHER NOTES:
The We Do uses novel content different from the I Do examples. The essential/non-essential distinction is the key analytical skill. The CFU specifically tests this distinction because it is the most common error point.

WATCH FOR:
- Students who identify the clause but mislabel essential/non-essential -- the "remove and check" test is the key strategy
- Students who correctly identify and label -- affirm and share reasoning

[General: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_YOUDO = `SAY:
- On the worksheet, you have eight sentences from the novel, each containing a relative clause
- FIRST: Underline the relative clause in each sentence
- NEXT: Circle the relative pronoun or relative adverb
- THEN: Write E for essential or NE for non-essential, and add commas where needed for non-essential clauses

DO:
- Distribute the Session 2 Relative Clauses Worksheet
- Circulate -- check that students are underlining the full clause, not just the pronoun
- After 5 minutes: "You should be on sentence 4 or 5 by now"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Complete sentences 1-4 only. Focus on finding and underlining the relative clause and circling the relative pronoun. The E/NE labelling is optional for these students
- Extra Notes: Encourage students to use the "remove and read" test -- cover the clause with their finger and read the sentence without it

EXTENDING PROMPT:
- Task: After completing all eight sentences, write three original sentences about Chapters 4-6: one using "who," one using "which," and one using "where." Include at least one essential and one non-essential clause with correct punctuation

TEACHER NOTES:
The You Do sentences are different from those used in the I Do and We Do to ensure genuine transfer. The sentences progress in difficulty -- sentences 1-4 have clearer relative clauses, sentences 5-8 require more careful analysis. The mix of essential and non-essential examples ensures students cannot default to one pattern.

WATCH FOR:
- Students who underline too much or too little -- "The relative clause starts at the relative pronoun and ends when the extra information about the noun is complete"
- Students who add commas to every relative clause -- "Only non-essential. If removing the clause changes the meaning, no commas"
- Students who finish quickly and accurately -- strong grammar instinct

[General: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: I can explain why the author switches between time periods -- thumbs? [scan]
- SC2: I can describe how the author develops characters through description and dialogue -- thumbs? [scan]
- SC3: I can identify a relative clause and its relative pronoun in a sentence -- thumbs? [scan]
- Turn and talk: Which character from Chapters 4-6 do you find most interesting so far, and why? Use a relative clause in your answer if you can

DO:
- Run through each SC with thumbs check
- The turn-and-talk integrates reading comprehension with the new grammar skill
- Preview: "Next lesson we read Chapters 7-8. Tom starts working as a chimney sweep"

TEACHER NOTES:
The closing integrates both strands. Students who can describe a character using a relative clause are demonstrating reading comprehension and grammar simultaneously. Students showing thumbs-down on SC3 will need additional practice -- relative clauses are revisited in future lessons.

WATCH FOR:
- Students thumbs-down on SC3 -- relative clauses are new, expect some uncertainty. Reassure: "We will keep practising these"
- Students who spontaneously use relative clauses in their turn-and-talk -- excellent transfer

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${WORKSHEET_RESOURCE.name} has eight sentences from the novel for you to practise with
- The ${ANSWER_KEY_RESOURCE.name} is for teacher reference

DO:
- Print the worksheet before the lesson (one per student)
- Print the answer key (teacher copy only)
- Click any resource card to open the PDF

TEACHER NOTES:
The worksheet uses exclusively Sweep content to maintain textual connection. Sentences progress in difficulty.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Sweep - Lesson 2 - Chapters 4-6 + Relative Clauses";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "Sweep",
    "Chapters 4-6 -- Sold, Reflected, Arrived",
    "Lesson 2  |  Week 1  |  Year 5/6 Literacy",
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
  // SLIDE 2 -- LI/SC
  // =========================================================================
  liSlide(
    pres,
    [
      "We are learning to analyse how an author develops characters and structures a narrative across time, and to identify relative clauses in sentences",
    ],
    [
      "I can explain why the author switches between two time periods in the story",
      "I can describe how the author develops a character through description and dialogue",
      "I can identify a relative clause and its relative pronoun in a sentence",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Vocabulary: overpowering
  // =========================================================================
  vocabSlide(
    pres,
    "overpowering",
    "adjective",
    "So strong that you cannot resist it or ignore it. It overwhelms your senses or your ability to cope -- it takes over completely.",
    "The stench of soot and damp in Master Jack's cellar was overpowering, hitting Tom the moment he descended the stairs.",
    NOTES_VOCAB_OVERPOWERING,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Vocabulary: dedication
  // =========================================================================
  vocabSlide(
    pres,
    "dedication",
    "noun",
    "Giving a lot of time and effort to something because you believe it matters. It shows commitment, loyalty, and purpose that goes beyond simple effort.",
    "Thomas had built his farm with decades of dedication, and every fence post and fruit tree showed it.",
    NOTES_VOCAB_DEDICATION,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- Reading Launch
  // =========================================================================
  contentSlide(
    pres,
    "Read Aloud",
    C.PRIMARY,
    "Chapters 4-6",
    [
      "Reading Mode: Student read aloud",
      "Ch 4: Tom arrives at the workhouse and is sold to Master Jack",
      "Ch 5: Flash-forward -- old Thomas reflects on family and farm",
      "Ch 6: Tom meets the other boys in Master Jack's cellar",
      "Focus: why does the author switch between time periods?",
    ],
    NOTES_READING,
    FOOTER
  );

  // =========================================================================
  // SLIDE 6 -- Pause Point 1 (p.15)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 1", "Chapter 4 -- p. 15",
    "'That way,' she ordered. 'And mind your manners.'",
    "p. 15",
    "What have we learned about Tom's situation from this conversation?",
    NOTES_PAUSE1, FOOTER
  );

  // =========================================================================
  // SLIDE 7 -- Pause Point 2 (p.17)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 2", "Chapter 4 -- p. 17",
    "Good day to you, milord.",
    "p. 17",
    "What is going on here? What does this scene reveal about how children were treated?",
    NOTES_PAUSE2, FOOTER
  );

  // =========================================================================
  // SLIDE 8 -- Pause Point 3 (p.21)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 3", "Chapter 5 -- p. 21",
    "Someone was in the rose garden, on the other side of the house.",
    "p. 21",
    "Why does the author keep switching between time periods? What do we learn from this flash-forward?",
    NOTES_PAUSE3, FOOTER
  );

  // =========================================================================
  // SLIDE 9 -- Pause Point 4 (p.26)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 4", "Chapter 6 -- p. 26",
    "Let's you say hello to 'em then, shall we?",
    "p. 26",
    "What does the author want us to know about Tom's situation now?",
    NOTES_PAUSE4, FOOTER
  );

  // =========================================================================
  // SLIDE 10 -- I Do: Relative Clauses
  // =========================================================================
  modellingSlide(
    pres,
    "I Do",
    "Relative Clauses",
    "A relative clause gives EXTRA\nINFORMATION about a noun.\n\nStarts with a relative pronoun:\nwho, whom, which, whose, that\n\nOr a relative adverb:\nwhen, where, why\n\nPlaced immediately AFTER the\nnoun it describes.",
    "Essential (no commas):\n\"The boy that Master Jack chose\nwas Tom.\"\nRemove it = meaning changes\n\nNon-essential (commas):\n\"Tom, who was only ten years old,\nfollowed Master Jack to the cart.\"\nRemove it = meaning stays clear",
    NOTES_RC_IDO,
    FOOTER
  );

  // =========================================================================
  // SLIDES 11-12 -- We Do: Relative Clauses (withReveal)
  // =========================================================================
  withReveal(
    () => contentSlide(
      pres,
      "We Do",
      C.SUCCESS,
      "Spot the Relative Clause",
      [
        "Find and underline the relative clause. Circle the relative pronoun. Essential or non-essential?",
        "\"The woman, who showed no warmth towards Tom, pointed down the corridor.\"",
        "\"The cart that carried Tom away from the workhouse rattled over the cobblestones.\"",
        "Discuss with your partner, then write E or NE on your whiteboard",
      ],
      NOTES_WEDO,
      FOOTER
    ),
    (slide) => {
      const ansY = 3.90;
      slide.addShape("roundRect", {
        x: 0.5, y: ansY, w: 9, h: 1.05, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      });

      slide.addShape("roundRect", {
        x: 0.7, y: ansY + 0.08, w: 1.5, h: 0.28, rectRadius: 0.08,
        fill: { color: C.WHITE },
      });
      slide.addText("Answers", {
        x: 0.7, y: ansY + 0.08, w: 1.5, h: 0.28,
        fontSize: 10, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      slide.addText("1. \"who showed no warmth towards Tom\" -- NE (commas needed, extra info)\n2. \"that carried Tom away from the workhouse\" -- E (no commas, tells us WHICH cart)", {
        x: 2.3, y: ansY + 0.06, w: 7.0, h: 0.90,
        fontSize: 12.5, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_WEDO);
    }
  );

  // =========================================================================
  // SLIDE 13 -- You Do: Relative Clauses
  // =========================================================================
  contentSlide(
    pres,
    "You Do",
    C.ACCENT,
    "Relative Clauses -- Independent Practice",
    [
      "FIRST: Underline the relative clause in each sentence",
      "NEXT: Circle the relative pronoun or relative adverb",
      "THEN: Write E (essential) or NE (non-essential) and add commas where needed",
    ],
    NOTES_YOUDO,
    FOOTER,
    (slide, layoutGuide) => {
      const tipY = layoutGuide.panelTopPadded + 0.1;
      addCard(slide, layoutGuide.rightX, tipY, layoutGuide.rightW, 2.3, {
        strip: C.SECONDARY, fill: C.BG_CARD,
      });
      slide.addText("Remember", {
        x: layoutGuide.rightX + 0.15, y: tipY + 0.08, w: 3.8, h: 0.28,
        fontSize: 11, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
      });
      slide.addText([
        { text: "Remove the clause and re-read", options: { bullet: true, breakLine: true, fontSize: 12, color: C.CHARCOAL } },
        { text: "If the meaning changes = Essential (no commas)", options: { bullet: true, breakLine: true, fontSize: 12, color: C.CHARCOAL } },
        { text: "If the meaning stays clear = Non-essential (commas)", options: { bullet: true, fontSize: 12, color: C.CHARCOAL } },
      ], {
        x: layoutGuide.rightX + 0.15, y: tipY + 0.44, w: 3.8, h: 1.6,
        fontFace: FONT_B, valign: "top", margin: 0, paraSpaceAfter: 6,
      });
    }
  );

  // =========================================================================
  // SLIDE 14 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    "Which character from Chapters 4-6 do you find most interesting? Use a relative clause in your answer.",
    [
      "I can explain why the author switches between two time periods in the story",
      "I can describe how the author develops a character through description and dialogue",
      "I can identify a relative clause and its relative pronoun in a sentence",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 15 -- Resources
  // =========================================================================


  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // Sentences for the worksheet (all from Sweep content)
  const rcSentences = [
    { num: 1, text: "Tom who had never seen London before stared at the dark streets.", clause: "who had never seen London before", pronoun: "who", type: "NE", corrected: "Tom, who had never seen London before, stared at the dark streets." },
    { num: 2, text: "The cellar that Master Jack used as a bedroom was damp and freezing.", clause: "that Master Jack used as a bedroom", pronoun: "that", type: "E", corrected: "The cellar that Master Jack used as a bedroom was damp and freezing." },
    { num: 3, text: "Jem who was the oldest boy showed Tom where to sleep.", clause: "who was the oldest boy", pronoun: "who", type: "NE", corrected: "Jem, who was the oldest boy, showed Tom where to sleep." },
    { num: 4, text: "The rags which replaced Tom's boots offered no warmth at all.", clause: "which replaced Tom's boots", pronoun: "which", type: "E", corrected: "The rags which replaced Tom's boots offered no warmth at all." },
    { num: 5, text: "Lettice whose voice was harsh demanded Tom's boots and breeches.", clause: "whose voice was harsh", pronoun: "whose", type: "NE", corrected: "Lettice, whose voice was harsh, demanded Tom's boots and breeches." },
    { num: 6, text: "The cart that rattled through the night carried Tom to his new life.", clause: "that rattled through the night", pronoun: "that", type: "E", corrected: "The cart that rattled through the night carried Tom to his new life." },
    { num: 7, text: "Thomas whose farm overlooked the valley thought about his sons.", clause: "whose farm overlooked the valley", pronoun: "whose", type: "NE", corrected: "Thomas, whose farm overlooked the valley, thought about his sons." },
    { num: 8, text: "The place where the boys slept smelled of soot and damp.", clause: "where the boys slept", pronoun: "where", type: "E", corrected: "The place where the boys slept smelled of soot and damp." },
  ];

  // --- PDF 1: Relative Clauses Worksheet -----------------------------------
  const ws = createPdf({ title: WORKSHEET_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Relative Clauses Worksheet", {
    color: C.NAVY,
    subtitle: "Sweep -- Chapters 4-6",
    lessonInfo: "Sweep | Lesson 2 | Week 1 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "A relative clause gives extra information about a noun. It starts with a relative pronoun (who, whom, which, whose, that) or a relative adverb (when, where, why).\n\nEssential (E) = removing it changes the meaning. No commas.\nNon-essential (NE) = removing it keeps the meaning clear. Use commas.", wsY, { color: C.NAVY });

  wsY = addSectionHeading(ws, "Underline the relative clause. Circle the relative pronoun/adverb. Write E or NE.", wsY, { color: C.NAVY });

  for (const s of rcSentences) {
    wsY = addBodyText(ws, `${s.num}. ${s.text}`, wsY, { fontSize: 11 });
    wsY += 2;
    wsY = addWriteLine(ws, "Relative clause:", wsY, {});
    wsY = addWriteLine(ws, "E or NE?", wsY, {});
    wsY += 6;
  }

  addPdfFooter(ws, "Sweep | Lesson 2 | Relative Clauses Worksheet");

  // --- PDF 2: Answer Key ---------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Relative Clauses -- Answer Key", {
    color: C.ALERT,
    subtitle: "Teacher Reference -- Chapters 4-6",
    lessonInfo: "Sweep | Lesson 2 | Week 1 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Students should underline the full relative clause, circle the relative pronoun/adverb, and correctly label E (essential) or NE (non-essential). For NE clauses, check that commas are added correctly.", akY, { color: C.ALERT });

  for (const s of rcSentences) {
    akY = addSectionHeading(ak, `Sentence ${s.num}`, akY, { color: s.type === "NE" ? C.SECONDARY : C.NAVY });
    akY = addBodyText(ak, `Corrected: ${s.corrected}`, akY, { fontSize: 11 });
    akY = addBodyText(ak, `Relative clause: "${s.clause}"`, akY, { fontSize: 10, italic: true });
    akY = addBodyText(ak, `Relative pronoun/adverb: ${s.pronoun} | Type: ${s.type} (${s.type === "E" ? "essential -- no commas" : "non-essential -- commas needed"})`, akY, { fontSize: 10 });
    akY += 6;
  }

  addPdfFooter(ak, "Sweep | Lesson 2 | Answer Key -- TEACHER COPY");

  // --- Write all files ------------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Sweep_Lesson2.pptx` }),
    writePdf(ws, WORKSHEET_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log(`PPTX written to ${OUT_DIR}/Sweep_Lesson2.pptx`);
  console.log(`Done: ${WORKSHEET_RESOURCE.name}`);
  console.log(`Done: ${ANSWER_KEY_RESOURCE.name}`);
}

build().catch(console.error);
