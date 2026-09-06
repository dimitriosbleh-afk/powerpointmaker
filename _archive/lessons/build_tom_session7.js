"use strict";

// Tom Unit -- Session 7: Chapters 15-17 -- Prison and Survival
// Week 2, Session 7, Grade 5/6 Literacy
// Reading: Character analysis, literary devices (personification, metaphor, imagery)
// Writing: Relative clauses (revision + application with novel content)
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
const FOOTER = "Chapters 15-17 | Session 7 | Week 2 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Session7_Prison_and_Survival";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const WORKSHEET_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Relative Clauses Worksheet",
  "Student worksheet: add relative clauses to sentences using novel content."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Answer Key",
  "Teacher reference: model relative clauses with essential/non-essential marking guidance."
);
const RESOURCE_ITEMS = [WORKSHEET_RESOURCE, ANSWER_KEY_RESOURCE];
const WORKSHEET_PDF_PATH = path.join(OUT_DIR, WORKSHEET_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Session 7 today. We are continuing with Chapters 15, 16, and 17
- Last session, Tom was arrested by the Bow Street Runners after Jem's death. Today we find out what happens next
- Our writing focus today is relative clauses -- adding extra information about a noun using words like "who", "which", and "that"

DO:
- Display title slide as students settle
- Have copies of the novel on desks, bookmarked at Chapter 15

TEACHER NOTES:
Session 7 of 10. Chapters 15-17 cover Tom's trial and imprisonment in Newgate, meeting Bald Sally, and his sentencing to transportation. Chapter 16 discusses death and hanging -- requires sensitivity advisory. The relative clauses work was introduced in Session 2; this session revises and applies it with novel content.

WATCH FOR:
- Students who may need a brief recap of Chapter 14 -- Tom was arrested for stealing coins
- Students who are still processing Jem's death from last session

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands today: reading Chapters 15-17 for character development, and adding relative clauses to our writing
- Read the success criteria. SC1 is about character through actions. SC2 is about the author's use of metaphor -- especially one metaphor that keeps coming back. SC3 is the writing target: relative clauses

DO:
- Choral read the LI, then the SCs
- Brief reminder: "Some of you may remember relative clauses from earlier. We will revise together"

TEACHER NOTES:
SC1 targets character resilience through Tom's adaptation to prison life. SC2 targets the recurring "courage cloak" metaphor. SC3 is the sentence-level writing target with relative clauses.

WATCH FOR:
- Students confident with relative clauses from Session 2 -- they can model for peers
- Students who missed Session 2 -- the I Do revises from scratch

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_REPULSED = `SAY:
- Our first vocabulary word: repulsed. Read it with me: repulsed [students repeat]
- Repulsed means feeling intense disgust or revulsion. When you are repulsed by something, you find it so unpleasant that you want to get away from it
- In our novel, Tom encounters things in Newgate Prison that would repulse most people -- the filth, the overcrowding, the smell
- Sentence completion: "She was repulsed by ___." Give me an example [the mouldy food, the smell of the drain, the sight of the rotting fruit]

DO:
- Display word, choral read, define, give example
- Cold Call 2 students for sentence completions
- Quick dramatisation: "Show me your face when you are repulsed by something" [students make disgusted faces]

TEACHER NOTES:
"Repulsed" carries strong emotional weight. The dramatisation engages kinaesthetic learners and makes the word memorable. Connects to prison descriptions in Chapter 15.

WATCH FOR:
- Students who confuse "repulsed" with "repelled" -- they are close synonyms, acknowledge this
- Students who think "repulsed" means "fought back" (as in "repulsed the attack") -- clarify: same word, different meaning in this context

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_VOCAB_CONDEMN = `SAY:
- Second word: condemn. Read it with me: condemn [students repeat]
- Condemn means to express strong disapproval, or to sentence someone to a punishment. A judge condemns a criminal to prison
- In our novel, Tom is condemned -- sentenced to punishment for his crime. The system condemns him
- Ask: What is the difference between "condemn" and "punish"? [Condemn is the judgement or declaration. Punish is carrying it out. A judge condemns; the prison punishes]

DO:
- Display word, choral read, define, example
- Synonym sort: write "condemn" on one side of the board. Ask: which words are close? [sentence, judge, denounce, criticise] Which are different? [praise, forgive, pardon]
- Connect to the novel: Tom is condemned by the justice system

TEACHER NOTES:
"Condemn" operates at both literal (legal sentencing) and figurative (social judgement) levels. Both meanings are relevant to the novel's themes.

WATCH FOR:
- Students who only think of "condemn" as legal -- expand to everyday usage: "People condemned the decision"
- Students who pronounce the silent "n" -- model: "con-DEM, not con-DEM-n"

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_VOCAB_EARNEST = `SAY:
- Third word: earnest. Read it with me: earnest [students repeat]
- Earnest means showing sincere and intense conviction. An earnest person is genuine and serious about what they are saying or doing. Not joking, not pretending
- In our novel, characters speak earnestly at key moments -- when the stakes are high and they mean every word
- Turn and talk: describe a time when someone was earnest with you. What made you know they were being serious and genuine? [30 seconds]

DO:
- Display word, choral read, define, example
- Turn and Talk: 30 seconds, Cold Call 2 pairs
- Thumbs Up/Down: "If someone is being sarcastic, are they being earnest?" [Thumbs down]

TEACHER NOTES:
"Earnest" is a character quality word that supports the analysis of dialogue in the novel. Tom's earnest desire to survive and Bald Sally's earnest advice are key moments.

WATCH FOR:
- Students who think "earnest" just means "serious" -- refine: earnest includes genuine care and conviction, not just seriousness
- Students who confuse it with the name "Ernest" -- acknowledge the connection (Oscar Wilde's play) but clarify the meaning

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_READING = `SAY:
- Chapters 15, 16, and 17 today. Tom has been arrested and these chapters follow what happens next
- I will read aloud again for these chapters. There is difficult content about prison, punishment, and the threat of hanging
- Chapter 16 discusses death and people being hanged. This was a real part of the justice system in this period. If you find this difficult, that is a normal response
- Focus: watch how Tom adapts and survives. What does his behaviour in prison reveal about his character?

DO:
- Give students 30 seconds to find Chapter 15
- Teacher reads aloud to control pacing around sensitive content
- Plan for 2 pause points: p.72 and p.81

SENSITIVITY ADVISORY:
- What it is: Chapter 15 -- Tom learns he may be hanged for theft. Chapter 16 -- death and the realities of Newgate Prison including discussion of hanging as punishment
- Framing language: "The justice system in this period was very different from today. Children could be sentenced to death for stealing. This is part of the history we are learning about"
- Watch for: Students who become anxious about the death penalty content, students who have experienced the justice system in their families
- Protocol: Frame as historical context. If a student becomes distressed, acknowledge privately and offer an alternative: "You can step out for a moment if you need to." Follow the school's wellbeing referral process

TEACHER NOTES:
Teacher reads aloud for control over pacing and tone around the sensitive content. Chapter 16 features Bald Sally, a midwife character who shows Tom kindness -- this is a deliberate hopeful counterpoint to the grim setting. Chapter 17 brings relief when the judge reduces Tom's charge.

WATCH FOR:
- Students who disconnect during the prison descriptions -- gently re-engage
- Students who are relieved by Chapter 17's outcome -- good: they are emotionally invested

[General: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause here. "Rats survived"
- This is about Tom in Newgate Prison. He has been watching how life works in the prison
- Ask: How does what we just read add to our understanding of Tom? [Tom is observing and learning. He watches the rats survive and applies the same logic to himself. He becomes a survivor -- stealing blankets, stealing money from visitors, doing whatever it takes to stay alive]
- Ask: What does this two-word sentence tell us? Why does the author make it so short? [Short sentence = impact. Tom has reduced survival to its simplest form. Rats survived. He will too]

DO:
- Display the quote
- Give 10 seconds of think time
- Cold Call for responses
- Push for the connection between the rats and Tom's survival strategy

CFU CHECKPOINT:
Technique: Cold Call

Script:
- "What are the rats doing that Tom notices?" [Surviving -- finding food, adapting, persisting]
- "So what does Tom decide to do?" [The same -- he becomes a survivor, adapting to prison life]
- "Why only two words? Why not a longer description?" [The author makes it blunt and simple. Survival is not poetic -- it is raw and basic]
- Scan for: students connecting the brevity of the sentence to its meaning

PROCEED (>=80%): Most students see the connection between rats and Tom's survival. Continue reading.
PIVOT (<80%): Most likely issue -- students see the rats literally and miss the metaphorical connection to Tom. Reteach: "The author did not put this sentence here to tell us about rats. Tom is watching the rats and thinking: if THEY can survive in here, so can I. The rats are his model for prison survival." Re-check: "What does Tom start doing after this?" [Stealing blankets and money -- surviving]

TEACHER NOTES:
This two-word sentence is a masterclass in brevity. The author strips away all description to show Tom's mindset: pure survival mode. This connects to SC1 (character through actions).

WATCH FOR:
- Students who focus only on the literal rats -- push to the metaphorical level
- Students who notice the sentence length -- excellent: this is authorial craft

[General: Pause Point 1 | VTLM 2.0: Higher-Order Questioning]`;

const NOTES_PAUSE2 = `SAY:
- Pause here. "...and wrapped his courage cloak more firmly round his shoulders"
- We have seen this metaphor before. Ask: When did we first encounter the courage cloak? [Earlier in the novel -- Tom uses this image to give himself strength]
- Ask: What does the author want us to know by REPEATING this metaphor here? [Tom needs courage again. The courage cloak is not a real thing -- it is Tom's way of finding strength inside himself. The author repeats it to show that Tom draws on the same inner resource again and again, especially when things are at their worst]
- Ask: How do things look for Tom right now? [He has just learned his sentence -- transportation. It is terrifying but also a relief -- he will not be hanged. The courage cloak shows he is gathering his strength for what comes next]

DO:
- Display the quote
- Use Think-Pair-Share: 20 seconds think, 30 seconds pair, share
- Draw attention to the recurring nature of the metaphor

TEACHER NOTES:
The recurring "courage cloak" metaphor is a key authorial device. Its repetition shows Tom's resilience -- he keeps reaching for the same internal resource. This directly targets SC2.

WATCH FOR:
- Students who remember the earlier appearance of the metaphor -- excellent recall and connection
- Students who take the cloak literally -- redirect: "Is it a real cloak? What is Tom actually doing?"

[General: Pause Point 2 | VTLM 2.0: Deep Comprehension]`;

const NOTES_IDO_RELATIVE = `SAY:
- Now we are working on sentence-level writing: relative clauses. Some of you may remember these from earlier
- A relative clause is a type of dependent clause that provides extra information about a noun. It begins with a relative pronoun: who, whom, which, whose, that -- or a relative adverb: when, where, why
- Why use them? They let us add detail without starting a new sentence. They make writing more sophisticated
- Watch me. "Tom entered Newgate Prison." I want to add information about the prison
- I add a relative clause: "Tom entered Newgate Prison, which was overcrowded and filthy"
- "which was overcrowded and filthy" -- this is the relative clause. It starts with "which" and tells us more about the prison
- Punctuation: this clause is non-essential -- it adds extra information but the sentence works without it. So I use commas
- Essential example: "The prisoners who stole food were punished." No commas -- "who stole food" is essential because it tells us WHICH prisoners

DO:
- Display terminology and model sentences
- Point to the relative pronoun as the start signal
- Model the comma test: remove the clause. Does the sentence still make sense AND still identify the right noun?
- Show one essential (no commas) and one non-essential (commas) example

TEACHER NOTES:
This revises the relative clause concept from Session 2 and applies it to novel content. The essential vs non-essential distinction is the key punctuation decision students need to make.

MISCONCEPTIONS:
- Misconception: All relative clauses need commas
  Why: Students overgeneralise from non-essential examples
  Impact: Incorrect punctuation in formal writing
  Quick correction: "Try removing the clause. If you no longer know WHICH noun the sentence is about, the clause is essential and gets NO commas. If the sentence still identifies the noun clearly, the clause is non-essential and gets commas"

WATCH FOR:
- Students who confuse relative clauses with appositives from Session 6 -- clarify: appositives are noun phrases (no connecting word); relative clauses start with who/which/that
- Students who remember this from Session 2 -- they can support peers during We Do

[General: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU = `SAY:
- Quick check. I am going to show you two sentences. You tell me: which one has a relative clause?
- Sentence A: "Tom, a young boy from the workhouse, was sentenced to transportation"
- Sentence B: "Tom, who had been arrested for theft, was sentenced to transportation"
- Finger vote: hold up 1 for Sentence A, hold up 2 for Sentence B
- Three, two, one -- show me! [Scan]

DO:
- Use Finger Voting
- Scan for: 2 fingers (Sentence B -- "who had been arrested for theft" is a relative clause starting with "who")
- Sentence A has an appositive, not a relative clause -- this tests the distinction

CFU CHECKPOINT:
Technique: Finger Voting

Script:
- "Hold up 1 for A, 2 for B. Which has a relative clause?"
- Count down, scan
- Look for: mostly 2 fingers

PROCEED (>=80%): Most show 2. Move to We Do.
PIVOT (<80%): Most likely issue -- students confuse appositives and relative clauses. Reteach: "Look for the signal word. Sentence B starts with 'who' -- that is a relative pronoun, so it is a relative clause. Sentence A -- 'a young boy from the workhouse' -- has no signal word. It is a noun phrase sitting next to the name. That is an appositive, which we learned last session." Re-check: "The boy who survived the prison became a successful man -- relative clause or appositive?"

TEACHER NOTES:
This CFU deliberately tests the appositive vs relative clause distinction since both were taught in this unit. The finger vote is quick and gives whole-class data.

WATCH FOR:
- Students who hold up 1 -- they may think any extra information in commas is a relative clause
- Students who are unsure and wait to copy others -- watch for delayed responses

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Sentence B: "Tom, who had been arrested for theft, was sentenced to transportation"
- "who had been arrested for theft" is the relative clause. It starts with "who" -- the relative pronoun
- Sentence A has an appositive -- "a young boy from the workhouse" -- no signal word, just a noun phrase
- The difference: relative clauses have a connecting word (who, which, that). Appositives do not

DO:
- Reveal the answer
- Highlight the relative pronoun "who" as the key signal
- Transition to We Do: "Now you will add relative clauses to sentences on your worksheets"

[General: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_WEDO = `SAY:
- On your worksheet, you have sentences from our novel. Your job is to add a relative clause to each sentence
- Remember: start with a relative pronoun (who, whom, which, whose, that) or a relative adverb (when, where, why)
- Then decide: is your clause essential or non-essential? That determines whether you need commas
- Let's do the first one together. "Newgate Prison terrified Tom." I want to add detail about the prison
- What relative pronoun works here? [which] "Newgate Prison, which was dark and overcrowded, terrified Tom"
- Is this essential or non-essential? [Non-essential -- we already know which prison. Commas needed]
- Now complete the rest. You have 8 minutes

DO:
- Distribute the Session 7 Relative Clauses Worksheet
- Do sentence 1 collaboratively (We Do), then release for You Do
- Circulate -- check relative pronoun usage and comma decisions

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide sentence frames with the relative pronoun already filled in and a word bank of clause completions. Students select and punctuate. E.g., "Tom, who ___, learned to survive" with options: was only ten years old / had been a climbing boy / came from the workhouse
- Extra Notes: Students can refer back to the I Do examples on screen

EXTENDING PROMPT:
- Task: Write a paragraph (4-5 sentences) about Tom's experience in Chapters 15-17 where at least 2 sentences contain relative clauses (one essential, one non-essential). Label each clause and explain the punctuation choice

TEACHER NOTES:
The worksheet ensures a range of relative pronouns (who, which, whose, that, where, when) and a mix of essential and non-essential clauses. Students who mastered appositives last session should find the transition manageable.

WATCH FOR:
- Students who add appositives instead of relative clauses -- remind: "Start with a relative pronoun"
- Students unsure about essential vs non-essential -- use the removal test: "Take it out. Do you still know which noun?"
- Students finishing quickly and accurately -- direct to the extending paragraph task

[General: We Do / You Do | VTLM 2.0: Guided + Independent Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Let's check. Here are model relative clauses for each sentence
- Your answers do not need to match exactly. Check: does your clause start with a relative pronoun? Is the punctuation correct for essential or non-essential?
- Sentence 1: "Newgate Prison, which was dark and overcrowded, terrified Tom"
- Sentence 2: "Bald Sally, who helped deliver a baby in the prison, gave Tom advice about survival"
- Sentence 3: "The judge who heard Tom's case reduced his charge"
- Notice: sentence 3 has NO commas because the clause is essential -- it tells us WHICH judge

DO:
- Reveal model answers
- Highlight the comma difference between essential and non-essential
- Celebrate accurate student work

[General: We Do / You Do Reveal | VTLM 2.0: Guided + Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: identifying what Tom's actions reveal about his character -- thumbs? [scan]
- SC2: explaining how the author uses the repeated courage cloak metaphor -- thumbs? [scan]
- SC3: adding a relative clause to a sentence with correct punctuation -- thumbs? [scan]
- Turn and talk: Tom survived by watching rats, stealing, and adapting. Does this make him admirable, or has the system forced him to become someone he would not otherwise be?

DO:
- Run through each SC with thumbs check
- The turn-and-talk is deliberately provocative -- there is no single right answer
- Preview: "Next session we continue with Chapters 18-19 -- Tom begins his journey to Botany Bay"

TEACHER NOTES:
The closing question challenges students to evaluate character rather than just describe it. This pushes beyond SC1 and SC2 into genuine critical thinking about the author's portrayal of survival.

WATCH FOR:
- Students who struggle to see both sides -- prompt: "Can something be both? Can Tom be admirable AND shaped by a cruel system?"
- Students "thumbs down" on SC3 -- check if it is the essential/non-essential distinction causing confusion

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two printable resources today
- The ${WORKSHEET_RESOURCE.name} is for the We Do and You Do relative clauses activity
- The ${ANSWER_KEY_RESOURCE.name} is for teacher reference

DO:
- Print the worksheet before the lesson (one per student)
- Print the answer key (teacher copy only)
- Click any resource card to open the PDF

TEACHER NOTES:
The worksheet includes a mix of relative pronouns and both essential and non-essential clause opportunities. The answer key flags which clauses are essential vs non-essential.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Chapters 15-17: Prison and Survival -- Session 7";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "Chapters 15-17",
    "Prison and Survival",
    "Session 7  |  Week 2  |  Year 5/6 Literacy",
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
  // SLIDE 2 -- LI / SC
  // =========================================================================
  liSlide(
    pres,
    [
      "We are learning to analyse how an author reveals character through adversity, and to add detail to sentences using relative clauses",
    ],
    [
      "I can identify what Tom's actions in prison reveal about his character",
      "I can explain how the author uses a repeated metaphor to show Tom's resilience",
      "I can add a relative clause to a sentence using the correct punctuation",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Vocabulary: repulsed
  // =========================================================================
  vocabSlide(
    pres,
    "repulsed",
    "adjective / verb",
    "Feeling intense disgust or revulsion. When you are repulsed by something, you find it so deeply unpleasant that you want to get away from it.",
    "Tom was repulsed by the conditions in Newgate Prison -- the filth, the stench, and the overcrowding.",
    NOTES_VOCAB_REPULSED,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Vocabulary: condemn
  // =========================================================================
  vocabSlide(
    pres,
    "condemn",
    "verb",
    "To express strong disapproval of something, or to sentence someone to a particular punishment. A judge condemns a criminal. Society condemns injustice.",
    "The court condemned Tom for stealing, even though he was only a child driven by desperation.",
    NOTES_VOCAB_CONDEMN,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- Vocabulary: earnest
  // =========================================================================
  vocabSlide(
    pres,
    "earnest",
    "adjective",
    "Showing sincere and intense conviction. An earnest person is genuine and serious about what they are saying or doing. Not joking, not pretending.",
    "Bald Sally spoke to Tom in earnest, giving him real advice about how to survive in the prison.",
    NOTES_VOCAB_EARNEST,
    FOOTER
  );

  // =========================================================================
  // SLIDE 6 -- Reading Launch
  // =========================================================================
  contentSlide(
    pres,
    "Teacher Read Aloud",
    C.PRIMARY,
    "Chapters 15-17",
    [
      "Reading Mode: Teacher Read Aloud",
      "Chapter 15: Tom is sent to Newgate Prison -- he learns to survive",
      "Chapter 16: Tom meets Bald Sally and helps deliver a baby",
      "Chapter 17: Tom's court hearing -- the charge is reduced",
      "Focus: How does Tom adapt to each new situation? What does survival look like?",
    ],
    NOTES_READING,
    FOOTER
  );

  // =========================================================================
  // SLIDE 7 -- Pause Point 1: "Rats survived"
  // =========================================================================
  quoteSlide(
    pres,
    "Pause Point 1",
    "Chapter 15 -- p. 72",
    "Rats survived.",
    "p. 72",
    "How does what we just read add to our understanding of Tom? Why does the author use only two words?",
    NOTES_PAUSE1,
    FOOTER
  );

  // =========================================================================
  // SLIDE 8 -- Pause Point 2: "courage cloak"
  // =========================================================================
  quoteSlide(
    pres,
    "Pause Point 2",
    "Chapter 17 -- p. 81",
    "...and wrapped his courage cloak more firmly round his shoulders.",
    "p. 81",
    "What does the author want us to know by repeating this metaphor? How do things look for Tom now?",
    NOTES_PAUSE2,
    FOOTER
  );

  // =========================================================================
  // SLIDE 9 -- I Do: Relative Clauses
  // =========================================================================
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Relative Clauses",
    "A relative clause adds extra information about a noun.\n\nStarts with a relative pronoun:\nwho, whom, which, whose, that\n\nOr a relative adverb:\nwhen, where, why\n\nEssential clause = no commas\n(tells us WHICH one)\n\nNon-essential clause = commas\n(adds extra detail)",
    "Model:\n\n\"Tom entered Newgate Prison.\"\n\nAdd a relative clause:\n\"Tom entered Newgate Prison, which was overcrowded and filthy.\"\n\nNon-essential: remove it and we still know which prison. Commas needed.\n\nEssential example:\n\"The prisoners who stole food were punished.\"\nNo commas -- tells us WHICH prisoners.",
    NOTES_IDO_RELATIVE,
    FOOTER
  );

  // =========================================================================
  // SLIDES 10-11 -- CFU: Relative Clause or Appositive? (withReveal)
  // =========================================================================
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Relative Clause or Appositive?",
      "Finger Voting",
      "Which sentence has a RELATIVE CLAUSE?\n\nA: \"Tom, a young boy from the workhouse, was sentenced to transportation.\"\n\nB: \"Tom, who had been arrested for theft, was sentenced to transportation.\"\n\nHold up 1 for A, 2 for B.",
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
        x: 0.7, y: ansY + 0.10, w: 1.8, h: 0.28, rectRadius: 0.08,
        fill: { color: C.WHITE },
      });
      slide.addText("Sentence B", {
        x: 0.7, y: ansY + 0.10, w: 1.8, h: 0.28,
        fontSize: 12, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addText("\"who had been arrested for theft\" starts with the relative pronoun \"who\" -- it is a relative clause. Sentence A has an appositive (a noun phrase with no connecting word).", {
        x: 2.7, y: ansY + 0.08, w: 6.6, h: 0.84,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // =========================================================================
  // SLIDES 12-13 -- We Do / You Do: Add Relative Clauses (withReveal)
  // =========================================================================
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.SECONDARY);
      addBadge(s, "We Do / You Do", { color: C.SECONDARY, w: 2.0 });
      addTitle(s, "Add a Relative Clause");

      const sentences = [
        { num: "1", text: "Newgate Prison terrified Tom.", hint: "Add detail about the prison (which...)", color: C.PRIMARY },
        { num: "2", text: "Bald Sally gave Tom advice about survival.", hint: "Add detail about Sally (who...)", color: C.SECONDARY },
        { num: "3", text: "The judge heard Tom's case.", hint: "Add detail about which judge (who/that...)", color: C.ACCENT },
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
        s.addText(sent.hint, {
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
        { y: CONTENT_TOP + 0.52, text: "\"Newgate Prison, which was dark and overcrowded, terrified Tom.\" (non-essential)" },
        { y: CONTENT_TOP + 1.70, text: "\"Bald Sally, who helped deliver a baby in the prison, gave Tom advice.\" (non-essential)" },
        { y: CONTENT_TOP + 2.88, text: "\"The judge who heard Tom's case reduced his charge.\" (essential -- no commas)" },
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
  // SLIDE 14 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    "Tom survived by watching rats, stealing, and adapting. Does this make him admirable, or has the system forced him to become someone he otherwise would not be? Tell your partner.",
    [
      "I can identify what Tom's actions in prison reveal about his character",
      "I can explain how the author uses a repeated metaphor to show Tom's resilience",
      "I can add a relative clause to a sentence using the correct punctuation",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 15 -- Resources
  // =========================================================================


  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: Relative Clauses Worksheet ------------------------------------
  const ws = createPdf({ title: WORKSHEET_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Relative Clauses -- Add Detail to Sentences", {
    color: C.PRIMARY,
    subtitle: "Chapters 15-17: Prison and Survival",
    lessonInfo: "Session 7 | Week 2 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "A relative clause starts with a relative pronoun (who, whom, which, whose, that) or a relative adverb (when, where, why). Essential clauses = no commas. Non-essential clauses = commas.", wsY, { color: C.PRIMARY });

  wsY = addSectionHeading(ws, "Section A: Add a Relative Clause (hint provided)", wsY, { color: C.PRIMARY });

  wsY = addBodyText(ws, "1. Newgate Prison terrified Tom. (Add detail about the prison using \"which\")", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 4;
  wsY = addBodyText(ws, "Essential or non-essential? Circle one:   ESSENTIAL   /   NON-ESSENTIAL", wsY, { fontSize: 10, italic: true });
  wsY += 8;

  wsY = addBodyText(ws, "2. Bald Sally gave Tom advice about survival. (Add detail about Sally using \"who\")", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 4;
  wsY = addBodyText(ws, "Essential or non-essential? Circle one:   ESSENTIAL   /   NON-ESSENTIAL", wsY, { fontSize: 10, italic: true });
  wsY += 8;

  wsY = addBodyText(ws, "3. The judge heard Tom's case. (Add detail about which judge using \"who\" or \"that\")", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 4;
  wsY = addBodyText(ws, "Essential or non-essential? Circle one:   ESSENTIAL   /   NON-ESSENTIAL", wsY, { fontSize: 10, italic: true });
  wsY += 12;

  wsY = addSectionHeading(ws, "Section B: Add a Relative Clause (your choice)", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "Choose the best relative pronoun or adverb. Decide whether commas are needed.", wsY);
  wsY += 4;

  wsY = addBodyText(ws, "4. Tom remembered the cellar. (Use \"where\")", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 8;

  wsY = addBodyText(ws, "5. The stolen money was found by the Bow Street Runners. (Use \"that\" or \"which\")", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 8;

  wsY = addBodyText(ws, "6. Tom's sentence was reduced. (Use \"whose\" or \"which\")", wsY, { fontSize: 12 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 12;

  wsY = addSectionHeading(ws, "Challenge: Write Your Own", wsY, { color: C.ALERT });
  wsY = addBodyText(ws, "Write 2 sentences about Chapters 15-17. One must have an essential relative clause (no commas) and one must have a non-essential relative clause (commas). Label each.", wsY);
  wsY = addLinedArea(ws, wsY, 4, { lineSpacing: 26 });

  addPdfFooter(ws, "Session 7 | Relative Clauses Worksheet");

  // --- PDF 2: Answer Key ---------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Relative Clauses -- Answer Key", {
    color: C.ALERT,
    subtitle: "Teacher Reference -- Chapters 15-17",
    lessonInfo: "Session 7 | Week 2 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Accept any relative clause that starts with the correct pronoun/adverb and fits the sentence logically. Check that essential/non-essential identification matches the punctuation.", akY, { color: C.ALERT });

  akY = addSectionHeading(ak, "Section A", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "1. \"Newgate Prison, which was dark and overcrowded, terrified Tom.\" -- NON-ESSENTIAL (we know which prison)", akY);
  akY = addBodyText(ak, "2. \"Bald Sally, who helped deliver a baby in the prison, gave Tom advice.\" -- NON-ESSENTIAL (Sally is named)", akY);
  akY = addBodyText(ak, "3. \"The judge who heard Tom's case reduced his charge.\" -- ESSENTIAL (tells us WHICH judge -- no commas)", akY);
  akY += 10;

  akY = addSectionHeading(ak, "Section B", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "4. \"Tom remembered the cellar where the boys had been kept by Master Jack.\" -- ESSENTIAL", akY);
  akY = addBodyText(ak, "   Alternative: \"Tom remembered the cellar, where he had slept on the cold floor.\" -- NON-ESSENTIAL", akY, { italic: true });
  akY += 4;
  akY = addBodyText(ak, "5. \"The stolen money that had been hidden in Tom's clothes was found by the Bow Street Runners.\" -- ESSENTIAL", akY);
  akY += 4;
  akY = addBodyText(ak, "6. \"Tom's sentence, which could have been death, was reduced to transportation.\" -- NON-ESSENTIAL", akY);
  akY += 14;

  akY = addSectionHeading(ak, "Common Errors", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- Using commas with essential clauses (\"The prisoners, who stole food, were punished\" -- incorrect if it specifies WHICH prisoners)", akY);
  akY = addBodyText(ak, "- Omitting commas with non-essential clauses", akY);
  akY = addBodyText(ak, "- Using appositives instead of relative clauses (no connecting word)", akY);
  akY = addBodyText(ak, "- Starting a sentence with a relative clause (relative clauses follow the noun they modify)", akY);

  addPdfFooter(ak, "Session 7 | Answer Key -- TEACHER COPY");

  // --- Write all files ----------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Session7.pptx` }),
    writePdf(ws, WORKSHEET_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Session7.pptx`);
  console.log("Done: " + WORKSHEET_RESOURCE.name);
  console.log("Done: " + ANSWER_KEY_RESOURCE.name);
}

build().catch(console.error);
