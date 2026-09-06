"use strict";

// Tom Unit -- Session 8: Chapters 18-19 -- Journey to Plymouth
// Week 2, Session 8, Grade 5/6 Literacy
// Reading: Character analysis, literary devices (metaphor, repetition/rule of threes)
// Writing: Note-taking -- KPAS (key words, phrases, abbreviations, symbols)

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

const SESSION_NUMBER = 8;
const FOOTER = "Chapters 18-19 | Session 8 | Week 2 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Session8_Journey_to_Plymouth";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const WORKSHEET_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "KPAS Note-Taking Worksheet",
  "Student worksheet: practise converting sentences into key words, phrases, abbreviations and symbols."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Answer Key",
  "Teacher reference: model KPAS conversions for each sentence."
);
const RESOURCE_ITEMS = [WORKSHEET_RESOURCE, ANSWER_KEY_RESOURCE];
const WORKSHEET_PDF_PATH = path.join(OUT_DIR, WORKSHEET_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Session 8 today. Two chapters: 18 and 19
- Last session Tom was sentenced to seven years transportation. Today we follow his journey towards the ship
- Our writing focus is a practical skill: note-taking using key words, phrases, abbreviations and symbols

DO:
- Display title slide as students settle
- Have copies of the novel on desks, bookmarked at Chapter 18

TEACHER NOTES:
Session 8 of 10. Chapters 18-19 cover Tom's transport to Plymouth and boarding the Scarborough, plus a reflective Thomas chapter. The note-taking skill (KPAS) is a study skill that supports the information report writing in Sessions 9-10.

WATCH FOR:
- Students who need a recap of Chapter 17 -- Tom was sentenced to transportation, not hanging
- Students who are curious about where Tom is going -- Botany Bay and the First Fleet

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands: reading Chapters 18-19 for character development, and learning to take efficient notes
- Read the success criteria. SC1 is about character through dialogue. SC2 is about the author's use of repetition. SC3 is the note-taking skill -- converting full sentences into key words, phrases, abbreviations and symbols

DO:
- Choral read the LI, then the SCs
- Brief preview: "Note-taking is a skill you will use across all your subjects. Today we learn a system for doing it efficiently"

TEACHER NOTES:
SC1 targets the Sam/Tom relationship -- Sam's dialogue reveals hope and experience. SC2 targets the "Down, down, down" repetition. SC3 is the KPAS writing skill that bridges into the information report sessions.

WATCH FOR:
- Students who think note-taking is just copying -- the KPAS system teaches them to transform, not transcribe
- Students excited about the First Fleet connection -- channel this into Sessions 9-10

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_IMPENETRABLE = `SAY:
- First vocabulary word: impenetrable. Read it with me: impenetrable [students repeat]
- Impenetrable means impossible to pass through or see through. An impenetrable wall cannot be broken. Impenetrable darkness cannot be seen through
- In our novel, Tom descends into the ship and encounters impenetrable darkness below deck -- he literally cannot see anything
- Ask: What is the opposite of impenetrable? [Penetrable, see-through, clear, passable]

DO:
- Display word, choral read, define, give example
- Cold Call: "What else could be described as impenetrable?" [A forest, fog, armour, a fortress, silence]
- Sentence completion: "The fog was so impenetrable that ___" [we could not see our hands, the ship could not navigate, everyone was lost]

TEACHER NOTES:
"Impenetrable" is a strong descriptive word that connects to the sensory imagery in Chapter 18. Teaching both literal (physical barrier) and figurative (impenetrable silence) builds transfer.

WATCH FOR:
- Students who struggle with the word length -- break it down: im-PENE-trable. "Penetrate" means to pass through. "Im" means not. So: cannot be passed through
- Students who only think of physical barriers -- extend to abstract: "impenetrable mystery"

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_VOCAB_QUEASY = `SAY:
- Second word: queasy. Read it with me: queasy [students repeat]
- Queasy means feeling sick or nauseous, especially in your stomach. You might feel queasy on a boat, in a car, or when you see something unpleasant
- In our novel, Tom boards a ship for the first time. The rocking, the smell, the cramped conditions below deck -- it is very likely making him queasy
- Quick dramatisation: show me your face when you feel queasy [students demonstrate]
- Turn and talk: when was the last time you felt queasy? What caused it? [30 seconds]

DO:
- Display word, choral read, define, example
- Dramatisation for engagement
- Turn and Talk: 30 seconds, Cold Call 2 pairs
- Connect to the novel: "Tom has never been on a ship before. He has been underground in chimneys and now he is in the belly of a boat"

TEACHER NOTES:
"Queasy" is a sensory word that connects to the physical experience of the ship voyage. The dramatisation and personal connection make it memorable.

WATCH FOR:
- Students who confuse "queasy" with "uneasy" -- they are related (both uncomfortable) but queasy is specifically about physical nausea
- Students who have motion sickness experiences -- their personal connection to the word is strong

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_READING = `SAY:
- Chapters 18 and 19. Tom leaves prison and begins the journey to Botany Bay
- I will read Chapter 18 aloud. Chapter 19 is shorter -- a Thomas reflection chapter
- Focus: watch for two things. First, the repetition of a word at a key moment. Second, how Sam's character is revealed through what he says to Tom
- Find Chapter 18 now

DO:
- Give students 30 seconds to find the chapter
- Teacher reads aloud
- Plan for 2 pause points: p.90 and p.95

TEACHER NOTES:
Chapter 18 introduces Sailor Sam, a significant character who provides hope and mentorship. Chapter 19 reveals Thomas's life outcomes (grazier, magistrate, large family), giving students a glimpse of where Tom's story ends. The "Down, down, down" repetition in Chapter 18 is a key literary device.

WATCH FOR:
- Students fascinated by the First Fleet details -- acknowledge and connect to Sessions 9-10
- Students who notice the shift in tone when Sam appears -- excellent character awareness

[General: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause here. "It was impossible to see"
- Just before this, the author wrote "Down, down, down..." as Tom descends into the ship
- Ask: Why might the author have repeated "down"? [The repetition creates a sense of descent -- Tom keeps going deeper and deeper. It feels relentless and claustrophobic. Each "down" takes him further from the light, further from freedom]
- Ask: What does this make you feel about the situation right now? [Trapped, suffocating, hopeless. Tom is being swallowed by the ship]

DO:
- Display the quote
- Read the "Down, down, down" slowly and deliberately
- Give think time, then Cold Call

CFU CHECKPOINT:
Technique: Cold Call

Script:
- "Why three 'downs' and not just one?" [Repetition creates intensity -- one "down" is a direction, three is a journey into darkness]
- "What does this remind you of from earlier in the novel?" [The chimneys -- Tom was always going down into dark, tight spaces. The ship is another version of the same experience]
- Scan for: students connecting the descent into the ship to the chimney experiences

PROCEED (>=80%): Most students understand the effect of repetition and connect to the chimney imagery. Continue reading.
PIVOT (<80%): Most likely issue -- students describe the literal meaning without the emotional effect. Reteach: "Read it aloud: 'Down, down, down.' Each word drags you further. The author is not just telling us Tom went below deck -- they are making us FEEL it. This is called the rule of threes -- repeating something three times for effect." Re-check: "What is the emotional effect?"

TEACHER NOTES:
The "Down, down, down" repetition uses the rule of threes for emotional intensity. It also creates a structural parallel with Tom's chimney experiences -- always descending into darkness.

WATCH FOR:
- Students who connect this to the chimney imagery -- excellent: pattern recognition across the novel
- Students who feel the claustrophobia of the description -- acknowledge: "The author wants you to feel that"

[General: Pause Point 1 | VTLM 2.0: Higher-Order Questioning]`;

const NOTES_PAUSE2 = `SAY:
- Pause here. "And Sam's voice was full of hope, for himself as well as Tom"
- Ask: What did we learn from this conversation between Sam and Tom? [Sam tells Tom about Botany Bay, about sailing the seas. He gives Tom something he has not had in a long time -- hope. And the author tells us Sam needs that hope too]
- Ask: How is the author making you feel right now about these characters? [Warm, hopeful, connected. After chapters of darkness and loss, Sam represents a new beginning. The author is deliberately shifting the tone]

DO:
- Display the quote
- Use Think-Pair-Share: 20 seconds think, 30 seconds pair, share
- Draw attention to "for himself as well as Tom" -- Sam is not just a mentor figure; he needs hope too

TEACHER NOTES:
This pause point reveals character through dialogue and relationship. Sam's hope is not one-directional -- both characters need each other. This is sophisticated characterisation that targets SC1.

WATCH FOR:
- Students who only see Sam as helping Tom -- push: "The author says the hope is for Sam too. Why might Sam need hope?"
- Students who notice the tonal shift from the darkness of prison -- excellent awareness of authorial craft

[General: Pause Point 2 | VTLM 2.0: Deep Comprehension]`;

const NOTES_IDO_KPAS = `SAY:
- Now we are learning a practical writing skill: note-taking using KPAS
- KPAS stands for Key words, Phrases, Abbreviations, and Symbols
- When you take notes, you do NOT copy full sentences. You transform them -- pull out the important information and write it quickly
- Watch me. Here is a sentence from the text: "In chains, Tom and the other prisoners are transported by wagon to Plymouth where the first fleet ships are anchored"
- Step 1: Find the key words -- the words that carry the important information. Who? Tom and prisoners. What? Transported. Where? Plymouth. How? In chains, by wagon. Why? First fleet ships
- Step 2: Drop the little words -- "the", "and", "are", "by", "where"
- My notes: "Tom + prisoners, chained, transported wagon -> Plymouth, first fleet ships anchored"
- I used + for "and", -> for "leads to/results in"

DO:
- Display the full sentence and the KPAS conversion side by side
- Model the thinking: "I ask: who, what, when, where, why, how? Those are my key words"
- Show the abbreviations and symbols on the slide
- Think aloud: "I check -- can I reconstruct the meaning from my notes? Yes. Good notes let you rebuild the original idea"

TEACHER NOTES:
The I Do models the complete KPAS process with a novel-content sentence. The key metacognitive move is the reconstruction check -- can the student rebuild meaning from the notes?

MISCONCEPTIONS:
- Misconception: Note-taking means copying key sentences from the text
  Why: Students are used to highlighting or copying rather than transforming
  Impact: Notes become too long to be useful and students do not process the information
  Quick correction: "If your notes look like the original text, you have copied, not transformed. Good notes are SHORTER and use YOUR system of abbreviations and symbols"

WATCH FOR:
- Students who want to copy the whole sentence -- redirect: "Which words can you drop and still keep the meaning?"
- Students who drop too many words and lose meaning -- the reconstruction check catches this

[General: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU = `SAY:
- Quick check. I will show you a sentence. You convert it to KPAS notes on your whiteboard
- "Once aboard the ship, Tom and the prisoners are ordered to strip and wash in half-barrels of seawater while their hair is clipped short to remove lice"
- Use key words, abbreviations and symbols. You have 30 seconds
- Hold up your boards! [Scan]

DO:
- Use Show Me Boards
- Scan for: key information captured (Tom + prisoners, strip + wash, seawater, hair clipped, remove lice)
- Scan for: appropriate use of symbols (+ for and, -> for results in)

CFU CHECKPOINT:
Technique: Show Me Boards

Script:
- "Convert this sentence to KPAS notes. 30 seconds"
- Count down, scan
- Look for: notes that are shorter than the original but capture who/what/where/why

PROCEED (>=80%): Most boards show transformed notes with key info captured. Move to We Do.
PIVOT (<80%): Most likely issue -- students copy most of the sentence instead of transforming it. Reteach: "Start with the questions: Who? Tom and prisoners. What happened? Ordered to strip and wash. In what? Seawater. What else? Hair clipped. Why? Remove lice. Now use symbols: Tom + prisoners, ordered strip + wash, seawater, hair clipped -> remove lice. That is half the length but keeps all the meaning." Re-check with a shorter sentence.

TEACHER NOTES:
Show Me Boards let every student demonstrate their KPAS attempt simultaneously. The model answer should be shared after scanning -- there is no single "correct" version, but all versions should be shorter than the original.

WATCH FOR:
- Students who copy word-for-word -- they need the "drop little words" strategy
- Students who over-abbreviate and lose meaning -- the reconstruction check helps
- Students who use creative but clear abbreviations -- celebrate and share

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Here is one way to convert that sentence. Your version does not need to match exactly
- Model: "Tom + prisoners, ordered strip + wash, half-barrels seawater w/ carbolic soap, hair clipped -> remove lice"
- Check: can you rebuild the original meaning from these notes? [Yes] Then the notes work
- The key: notes should be SHORTER than the original but keep all the important information

DO:
- Reveal model KPAS conversion
- Compare with 2-3 student examples on whiteboards -- celebrate different valid approaches
- Transition to We Do on the worksheet

[General: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_WEDO = `SAY:
- On your worksheet, you have sentences from the novel. Convert each one to KPAS notes
- Remember: key words, phrases, abbreviations and symbols
- Use the abbreviation and symbol guide on the worksheet
- Let's do sentence 1 together. [Read it aloud, take suggestions, model the conversion on the board]
- Now complete sentences 2 through 5 independently. You have 8 minutes

DO:
- Distribute the Session 8 KPAS Note-Taking Worksheet
- Do sentence 1 collaboratively (We Do), then release for You Do
- Circulate -- check that students are transforming, not copying

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide the key words already underlined in the sentences. Students only need to write the underlined words with appropriate symbols. Reduces the cognitive load of identifying key information
- Extra Notes: Students can use the abbreviation guide on the worksheet for support

EXTENDING PROMPT:
- Task: Write a short passage (3-4 sentences) summarising what happens in Chapters 18-19 using full sentences. Then convert your OWN passage into KPAS notes. Compare the two versions -- which is more efficient?

TEACHER NOTES:
The progression from We Do (sentence 1) to You Do (sentences 2-5) mirrors the gradual release. The worksheet includes a reference box of common abbreviations and symbols for student support.

WATCH FOR:
- Students who copy rather than transform -- the key question: "Is your version shorter?"
- Students who enjoy the efficiency of the system -- they will use it in Sessions 9-10
- Students finishing quickly -- direct to the extending passage task

[General: We Do / You Do | VTLM 2.0: Guided + Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: identifying what Sam and Tom's dialogue reveals -- thumbs? [scan]
- SC2: explaining the effect of the "Down, down, down" repetition -- thumbs? [scan]
- SC3: converting sentences into KPAS notes -- thumbs? [scan]
- Turn and talk: Chapter 19 tells us Thomas became a grazier and a magistrate with a huge family. Knowing where Tom ends up, does that change how you feel about what he is going through now?

DO:
- Run through each SC with thumbs check
- The turn-and-talk connects the two timelines -- students reflect on how knowing the ending affects their experience of the journey
- Preview: "Next session we start reading about Tom's voyage AND we begin planning our information report about 18th century England"

TEACHER NOTES:
The closing question is about narrative structure -- knowing the ending changes how we read the middle. This is a sophisticated reader response concept.

WATCH FOR:
- Students who say "it makes it less sad because we know he survives" vs "it makes it more interesting because we see HOW he survived" -- both are valid
- Students "thumbs down" on SC3 -- KPAS will be practised again in Session 9

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two printable resources today
- The ${WORKSHEET_RESOURCE.name} is for the KPAS note-taking activity
- The ${ANSWER_KEY_RESOURCE.name} is for teacher reference

DO:
- Print the worksheet before the lesson (one per student)
- Print the answer key (teacher copy only)

TEACHER NOTES:
The worksheet includes a reference box of common abbreviations and symbols. The answer key provides model KPAS conversions but emphasises that student versions may differ -- the test is whether meaning is preserved.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Chapters 18-19: Journey to Plymouth -- Session 8";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "Chapters 18-19",
    "Journey to Plymouth",
    "Session 8  |  Week 2  |  Year 5/6 Literacy",
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
      "We are learning to analyse how an author uses dialogue and repetition to develop character, and to take efficient notes using key words, phrases, abbreviations and symbols",
    ],
    [
      "I can identify what a character's dialogue reveals about their personality and hopes",
      "I can explain how the author uses repetition to create emotional effect",
      "I can convert sentences into key words, phrases, abbreviations and symbols (KPAS)",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Vocabulary: impenetrable
  // =========================================================================
  vocabSlide(
    pres,
    "impenetrable",
    "adjective",
    "Impossible to pass through or see through. An impenetrable wall cannot be broken. Impenetrable darkness cannot be seen through. Can also describe something impossible to understand.",
    "Tom descended into the impenetrable darkness below deck, unable to see anything at all.",
    NOTES_VOCAB_IMPENETRABLE,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Vocabulary: queasy
  // =========================================================================
  vocabSlide(
    pres,
    "queasy",
    "adjective",
    "Feeling sick or nauseous, especially in your stomach. You might feel queasy on a boat, in a car, or when confronted with something unpleasant.",
    "The rocking of the ship and the stench below deck made Tom feel queasy for the first time in his life.",
    NOTES_VOCAB_QUEASY,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- Reading Launch
  // =========================================================================
  contentSlide(
    pres,
    "Teacher Read Aloud",
    C.PRIMARY,
    "Chapters 18-19",
    [
      "Reading Mode: Teacher Read Aloud",
      "Chapter 18: Tom is transported to Plymouth and boards the Scarborough",
      "Chapter 19: Thomas reflects on his life -- family, career, legacy",
      "Focus: How does the author use repetition at a key moment?",
      "Focus: What does Sam's dialogue reveal about his character?",
    ],
    NOTES_READING,
    FOOTER
  );

  // =========================================================================
  // SLIDE 6 -- Pause Point 1: "Down, down, down"
  // =========================================================================
  quoteSlide(
    pres,
    "Pause Point 1",
    "Chapter 18 -- p. 90",
    "It was impossible to see.",
    "p. 90",
    "Why might the author have repeated 'down'? What does this make you feel about the situation right now?",
    NOTES_PAUSE1,
    FOOTER
  );

  // =========================================================================
  // SLIDE 7 -- Pause Point 2: Sam's voice
  // =========================================================================
  quoteSlide(
    pres,
    "Pause Point 2",
    "Chapter 18 -- p. 95",
    "And Sam's voice was full of hope, for himself as well as Tom.",
    "p. 95",
    "What did we learn from this conversation? How is the author making you feel about these characters?",
    NOTES_PAUSE2,
    FOOTER
  );

  // =========================================================================
  // SLIDE 8 -- I Do: KPAS Note-Taking
  // =========================================================================
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Note-Taking: KPAS",
    "KPAS = Key words, Phrases,\nAbbreviations, Symbols\n\nWhen taking notes, TRANSFORM\n-- do not copy.\n\nAsk: who / what / when /\nwhere / why / how?\n\nCommon abbreviations:\nb/c = because\nw/ = with, w/o = without\n\nCommon symbols:\n= means / definition\n+ and\n-> leads to / results in",
    "Full sentence:\n\"In chains, Tom and the other prisoners are transported by wagon to Plymouth where the first fleet ships are anchored.\"\n\nKPAS notes:\nTom + prisoners, chained,\ntransported wagon -> Plymouth,\nfirst fleet ships anchored\n\nCheck: Can I rebuild the meaning\nfrom my notes? Yes.",
    NOTES_IDO_KPAS,
    FOOTER
  );

  // =========================================================================
  // SLIDES 9-10 -- CFU: Convert a Sentence (withReveal)
  // =========================================================================
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Convert to KPAS Notes",
      "Show Me Boards",
      "\"Once aboard the ship, Tom and the prisoners are ordered to strip and wash in half-barrels full of seawater using fatty cakes of carbolic soap while their hair is clipped short to remove lice.\"\n\nConvert this to KPAS notes on your whiteboard. 30 seconds.",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      const ansY = 3.8;
      slide.addShape("roundRect", {
        x: 0.5, y: ansY, w: 9, h: 1.2, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      });
      slide.addText("Model KPAS Notes", {
        x: 0.75, y: ansY + 0.08, w: 3, h: 0.26,
        fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true, margin: 0,
      });
      slide.addText("Tom + prisoners, ordered strip + wash, half-barrels seawater w/ carbolic soap, hair clipped -> remove lice\n\nYour version may differ -- the test: is it shorter AND does it keep the meaning?", {
        x: 0.75, y: ansY + 0.36, w: 8.5, h: 0.76,
        fontSize: 12, fontFace: FONT_B, color: C.WHITE, valign: "top", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // =========================================================================
  // SLIDE 11 -- We Do / You Do: KPAS Practice
  // =========================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do / You Do", { color: C.SECONDARY, w: 2.0 });
    addTitle(s, "KPAS Practice -- Your Turn");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.4, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Convert each sentence to KPAS notes on your worksheet:", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.28,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0,
    });
    s.addText("First: We do sentence 1 together\nNext: Complete sentences 2-5 independently\nThen: Check -- can you rebuild the meaning from your notes?", {
      x: 0.75, y: CONTENT_TOP + 0.44, w: 8.4, h: 0.80,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const refY = CONTENT_TOP + 1.54;
    addCard(s, 0.5, refY, 9, SAFE_BOTTOM - refY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Quick Reference", {
      x: 0.75, y: refY + 0.08, w: 3, h: 0.28,
      fontSize: 11, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("b/c = because     w/ = with     w/o = without\n= means/definition     + and     -> leads to / results in\nDrop little words: the, a, an, is, are, was, were", {
      x: 0.75, y: refY + 0.38, w: 8.4, h: 0.80,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
  }

  // =========================================================================
  // SLIDE 12 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    "Chapter 19 tells us Thomas became a grazier and a magistrate with a huge family. Knowing where Tom ends up, does that change how you feel about what he is going through now?",
    [
      "I can identify what a character's dialogue reveals about their personality and hopes",
      "I can explain how the author uses repetition to create emotional effect",
      "I can convert sentences into key words, phrases, abbreviations and symbols (KPAS)",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 13 -- Resources
  // =========================================================================


  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: KPAS Note-Taking Worksheet ------------------------------------
  const ws = createPdf({ title: WORKSHEET_RESOURCE.name });
  let wsY = addPdfHeader(ws, "KPAS Note-Taking Practice", {
    color: C.PRIMARY,
    subtitle: "Chapters 18-19: Journey to Plymouth",
    lessonInfo: "Session 8 | Week 2 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "KPAS = Key words, Phrases, Abbreviations, Symbols. Transform sentences into short notes that capture the important information. Drop little words (the, a, is, are, was). Use the guide below.", wsY, { color: C.PRIMARY });

  wsY = addSectionHeading(ws, "Abbreviation and Symbol Guide", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "b/c = because  |  w/ = with  |  w/o = without  |  = means / definition  |  + and  |  -> leads to / results in / cause and effect", wsY, { fontSize: 10 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Convert each sentence to KPAS notes:", wsY, { color: C.PRIMARY });

  wsY = addBodyText(ws, "1. Tom was transported by cart with other prisoners from London to Plymouth, a journey that took three days and three nights along muddy, dangerous roads.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 6;

  wsY = addBodyText(ws, "2. Tom was put aboard a ship called the Scarborough and was assigned a berth below deck with a sailor named Sam.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 6;

  wsY = addBodyText(ws, "3. Sam told Tom about Botany Bay and what it was like to sail the seven seas, filling him with hope about the life that waited for him.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 6;

  wsY = addBodyText(ws, "4. Thomas had five children, twenty-seven grandchildren and forty-three great-grandchildren, and had worked as both a grazier and a magistrate.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 6;

  wsY = addBodyText(ws, "5. The convicts were forced to strip and wash in half-barrels of seawater using carbolic soap, and their hair was clipped short to remove lice before they were allowed below deck.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 24 });
  wsY += 10;

  wsY = addSectionHeading(ws, "Reconstruction Check", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "Choose one set of your notes. Can you explain the full meaning to a partner using only your notes? If yes, your notes work. If no, add the missing key words.", wsY);

  addPdfFooter(ws, "Session 8 | KPAS Note-Taking Worksheet");

  // --- PDF 2: Answer Key ---------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "KPAS Note-Taking -- Answer Key", {
    color: C.ALERT,
    subtitle: "Teacher Reference -- Chapters 18-19",
    lessonInfo: "Session 8 | Week 2 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Student notes will vary. Accept any version that is shorter than the original and preserves the key information. The test: can the student reconstruct the meaning from their notes?", akY, { color: C.ALERT });

  akY = addSectionHeading(ak, "Model KPAS Conversions", akY, { color: C.PRIMARY });

  akY = addBodyText(ak, "1. Tom + prisoners, transported cart London -> Plymouth, 3 days + 3 nights, muddy roads", akY);
  akY += 4;
  akY = addBodyText(ak, "2. Tom aboard Scarborough, berth below deck w/ Sailor Sam", akY);
  akY += 4;
  akY = addBodyText(ak, "3. Sam told Tom re Botany Bay + sailing, filled w/ hope -> new life", akY);
  akY += 4;
  akY = addBodyText(ak, "4. Thomas: 5 children, 27 grandchildren, 43 great-grandchildren, grazier + magistrate", akY);
  akY += 4;
  akY = addBodyText(ak, "5. Convicts ordered strip + wash, half-barrels seawater w/ carbolic soap, hair clipped -> remove lice", akY);
  akY += 14;

  akY = addSectionHeading(ak, "What to Look For", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- Notes that are significantly shorter than the originals (at least half the length)", akY);
  akY = addBodyText(ak, "- Key information preserved: who, what, when, where, why, how", akY);
  akY = addBodyText(ak, "- Use of at least some abbreviations or symbols (not required to use all)", akY);
  akY = addBodyText(ak, "- Little words dropped (the, a, an, is, are, was, were)", akY);
  akY = addBodyText(ak, "- Student can reconstruct meaning from notes (the reconstruction check)", akY);
  akY += 10;
  akY = addBodyText(ak, "Common issues:", akY, { fontSize: 11 });
  akY = addBodyText(ak, "- Copying full sentences instead of transforming", akY);
  akY = addBodyText(ak, "- Over-abbreviating so meaning is lost", akY);
  akY = addBodyText(ak, "- Keeping too many little words (reads like slightly shorter sentences)", akY);

  addPdfFooter(ak, "Session 8 | Answer Key -- TEACHER COPY");

  // --- Write all files ----------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Session8.pptx` }),
    writePdf(ws, WORKSHEET_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Session8.pptx`);
  console.log("Done: " + WORKSHEET_RESOURCE.name);
  console.log("Done: " + ANSWER_KEY_RESOURCE.name);
}

build().catch(console.error);
