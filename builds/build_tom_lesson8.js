"use strict";

// Tom Appleby, Convict Boy -- Lesson 8: Chapters 18-19 -- Aboard the Scarborough
// Year 5/6 Literacy | Novel study
// Reading: Character analysis; literary devices (metaphor, repetition / rule of threes)
// Writing (sentence-level): Note-taking -- convert sentences to keywords, phrases, abbreviations, symbols (KPAS)

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
const FOOTER = "Chapters 18-19 | Lesson 8 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson8_Aboard_the_Scarborough";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const WORKSHEET_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Note Taking Worksheet",
  "Student worksheet: convert sentences to keywords, phrases, abbreviations and symbols."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Note Taking Answer Key",
  "Teacher reference: model notes and marking guidance."
);
const RESOURCE_ITEMS = [WORKSHEET_RESOURCE, ANSWER_KEY_RESOURCE];
const WORKSHEET_PDF_PATH = path.join(OUT_DIR, WORKSHEET_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Two chapters today -- 18 and 19. Tom boards the ship that will take him to Botany Bay
- Our writing focus today is a new skill: note-taking -- how to capture important information without writing every word

DO:
- Display title slide as students settle
- Have novels on desks, bookmarked at Chapter 18

TEACHER NOTES:
Chapter 18 is long but visually rich -- the transportation to the ship and Tom's first experiences on board. Chapter 19 is a short modern-timeline chapter revisiting Thomas. The writing skill shifts from sentence detail to note-taking.

WATCH FOR:
- Students who want to summarise last lesson's court scene -- allow a 15 second recap, then move on
- Students who have never taken formal notes -- today is an introduction, not a mastery target

[Literacy: Title | VTLM 2.0: Establishing Purpose]`;

const NOTES_LI_SC = `SAY:
- Two strands today: reading Chapters 18-19 for character under new pressures, and learning to take notes using keywords, phrases, abbreviations and symbols
- Read the success criteria from the slide
- SC1 is the floor -- everyone can point to something the author shows us about Tom's new situation. SC2 is the core -- analysing craft. SC3 is the writing target -- converting sentences to notes

DO:
- Choral read LI and SCs
- Do not pre-teach KPAS yet -- the I Do covers it

TEACHER NOTES:
SC3 is genuinely new. Frame it as a useful life skill -- researchers, journalists and students use this every day.

WATCH FOR:
- Students who think notes must use full sentences -- gently challenge this now
- Students who use emojis or texting shortcuts -- celebrate the impulse, then refine it

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_IMPENETRABLE = `SAY:
- First word: impenetrable. Choral read: impenetrable
- Impenetrable means impossible to enter or pass through -- nothing can get in
- Break it down: IM means not. PENETRABLE means able to be passed through. So impenetrable = NOT able to be passed through
- In the novel, Tom looks into impenetrable darkness below deck. He cannot see through it at all
- Quick check: is a thick forest impenetrable? [Yes -- hard or impossible to get through]

DO:
- Display word, choral read, define, example
- Morphology moment: show "im-" prefix = "not" (impossible, imperfect, improbable)
- Cold Call: "Name something that could be impenetrable"

TEACHER NOTES:
"Impenetrable" teaches both the word and a useful morphology strategy. The "im-" prefix appears often -- noticing it helps students decode future words.

WATCH FOR:
- Students who pronounce it "impen-TABLE" -- stress the correct syllables: im-PEN-e-tra-ble
- Students confusing it with "inevitable" -- different meaning (sure to happen)

[Literacy: Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_VOCAB_QUEASY = `SAY:
- Second word: queasy. Choral read: queasy
- Queasy means feeling sick in your stomach -- nauseous, wobbly, unwell. Often from motion or something unpleasant
- In the novel, Tom feels queasy as the ship rocks in the harbour. He is not used to the movement
- Quick check: if you are queasy, do you feel well or unwell? [Unwell]

DO:
- Display word, choral read, define, example
- Ask: "What might make YOU feel queasy?" -- take 3 quick answers (car rides, boats, spinning, bad smells)
- Sentence completion: "The smell of ___ made me feel queasy"

TEACHER NOTES:
"Queasy" is a concrete, relatable word. Students often have experience of this feeling but not the vocabulary.

WATCH FOR:
- Students who think queasy means very scared -- close but not quite; it is the stomach feeling
- Students spelling it "queezy" -- note the unusual ending -asy

[Literacy: Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_READING = `SAY:
- Chapters 18 and 19. Tom is taken to the ship and boards it for the first time. Chapter 19 is short -- we return to old Thomas
- Focus: how does the author show us what Tom is feeling as he faces this new ordeal?
- I will read aloud -- pause at one key moment today on page 89
- Find Chapter 18 now

DO:
- Give 30 seconds to find the chapter
- Read aloud at a steady pace
- One pause point today -- p. 89

TEACHER NOTES:
These chapters are dense with sensory detail. Point out the author's use of sound, smell, and movement as Tom boards the ship. The short Chapter 19 gives brief respite and reminds us old Tom survived.

WATCH FOR:
- Students getting lost in the nautical vocabulary -- tiller, berth, pannikins. Do not stop to define every word; teach these as they arise or skip over
- Students connecting "down" in Ch 18 with earlier "down a chimney" imagery -- excellent cross-chapter thinking

[Literacy: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause here. "Down, down, down..."
- Three words. Same word. Three times. This is called REPETITION -- specifically the rule of threes
- Ask: what is Tom experiencing here? [He is being taken below deck. Deeper and deeper into the dark hold of the ship]
- Ask: why has the author repeated "down"? [To make us FEEL the descent. Each repetition takes us further down with Tom. It feels endless. It feels inescapable]
- Connect: where have we seen a similar descent before? [Down the chimney. Tom is always being forced downward -- into dark, enclosed spaces]

DO:
- Display the quote
- Think time: 10 seconds
- Cold Call 3 students
- Draw the craft connection: repetition creates rhythm and emphasis

CFU CHECKPOINT:
Technique: Cold Call
Script:
- "What is the author making us feel with those three words?"
- "Where else has Tom gone 'down' in this novel?"
- Scan for: students connecting the repetition to emotional effect AND earlier chapters
PROCEED:
- Most students engage with repetition as a craft choice. Continue reading.
PIVOT:
- Most likely issue -- students read the repetition as just stating a fact
- Reteach: "The author could have written 'Tom went down deep into the ship.' Instead they wrote 'Down, down, down.' What is DIFFERENT about how that makes us feel?"
- Re-check: "Why three, not two or four?"

TEACHER NOTES:
This pause point teaches the rule of threes and repetition as deliberate craft. The recurring "down" motif ties this scene to earlier chimney scenes -- this is analytic depth (SC2).

WATCH FOR:
- Students who treat repetition as a mistake -- clarify: purposeful repetition is a craft tool
- Students connecting "down the chimney" to "down below deck" -- celebrate this insight

[Literacy: Pause Point 1 | VTLM 2.0: Higher-Order Questioning]`;

const NOTES_IDO_KPAS = `SAY:
- Our writing focus: taking notes. When you read, you do not always need full sentences. You can capture the important ideas using keywords, phrases, abbreviations and symbols -- we call this KPAS
- Why? Because notes are for YOU. They are a memory aid. They should be fast to write and easy to understand when you come back to them
- Watch me. Here is a sentence from our text:
- "In chains, Tom and the other prisoners are transported by wagon to Plymouth where the first fleet ships are anchored."
- First, I find the KEY WORDS. Who? Tom and other prisoners. What doing? Transported. Where? To Plymouth. How? By wagon, in chains
- Now I use abbreviations and symbols
- My notes: "In chains, Tom + other prisoners  transported by wagon  Plymouth (First Fleet ships anchored)"
- Notice: no "the", no "are", no "where". Small words are dropped. The core meaning stays

DO:
- Display the sentence and the notes side by side
- Point out each key word in colour
- Show the arrow, the plus sign, the brackets
- Think aloud: "Does this note still make sense if I come back to it in a week? Yes"

TEACHER NOTES:
The key move is identifying WHO / WHAT DOING / WHEN / WHERE / WHY / HOW, then keeping only those. Everything else is decoration.

WATCH FOR:
- Students who copy the sentence and then shorten -- reteach: find the key words FIRST, then write the note
- Students whose notes are so short they lose meaning -- notes must still make sense later

[Literacy: I Do -- Modelling | VTLM 2.0: Explicit Teaching]`;

const NOTES_IDO_SYMBOLS = `SAY:
- Here are the most useful abbreviations and symbols. You do not need to memorise them all today -- we will build up over time
- Abbreviations:
-   b/c = because
-   w/ = with
-   w/o = without
- Symbols:
-   = means "is" or "means" or "equals"
-   + means "and"
-   an arrow  means "leads to", "results in", "cause and effect"
- When taking notes, drop small words: the, a, and, is, are, that

DO:
- Display the abbreviations and symbols clearly
- Choral read each one: "b slash c equals because..."
- Ask: if you see "T + J w/o food  queasy", can you read it back in a full sentence? [Tom and Jem, without food, felt queasy]

TEACHER NOTES:
The abbreviation and symbol set is small by design -- mastery over addition. Students will pick up more over time. Today's focus is the core set.

WATCH FOR:
- Students inventing their own abbreviations -- encourage this as long as they remember what it means
- Students using emojis -- allow with a caveat: emojis work IF you will remember what they mean

[Literacy: I Do -- Abbreviations | VTLM 2.0: Explicit Teaching]`;

const NOTES_CFU = `SAY:
- Quick check. I will show you a sentence. On your whiteboard, identify ONLY the key words -- the who, what doing, where, when, why, how
- Ignore the small words like "the", "and", "is"
- You have 20 seconds

DO:
- Display the sentence
- Use Show Me Boards -- count down, then scan
- Scan for: Tom, prisoners, ordered, strip, wash, half-barrels, carbolic soap, hair clipped

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- "Write ONLY the key words. Drop the small words. 20 seconds"
- Count down, scan
- Look for: Tom, prisoners, stripped, washed, barrels, carbolic soap, hair clipped -- or close variants
PROCEED:
- Most boards show the key words. Reveal and move to the converted note.
PIVOT:
- Most likely issue -- students copy the whole sentence
- Reteach: "Ask yourself: if I cross out 'the', 'and', 'is', 'are' -- what is LEFT? That is your key words"
- Re-check with a new shorter sentence: "Tom was seasick for three days." [key words: Tom, seasick, three days]

TEACHER NOTES:
Show Me Boards give visible evidence. The "cross out small words" strategy is the scaffold.

WATCH FOR:
- Students copying the whole sentence -- redirect with crossing-out strategy
- Students missing quantities (three, five, hours) -- these are key details

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Great. The key words are: Tom + prisoners, ordered to strip, wash, half-barrels w/ carbolic soap, hair clipped short
- Now I convert with abbreviations and symbols: [Tom + prisoners ordered strip + wash, half-barrels w/ carbolic soap + hair clipped short  remove lice]
- One sentence, 22 words. My note: about 15 words. Saves time, keeps the meaning

DO:
- Reveal the key words
- Then reveal the converted note
- Transition to We Do / You Do

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_WEDO = `SAY:
- On your worksheet, you have three sentences from our text. Your job: convert each sentence to a note using keywords, phrases, abbreviations and symbols
- Section A gives you a partial scaffold -- some words are underlined as key words to help you start
- Section B is independent -- you find the key words yourself, then convert
- Let's do sentence 1 together. Find who, what doing, where, when, why, how
- Then decide which abbreviations or symbols could help
- 10 minutes on this section

DO:
- Distribute worksheets
- Model sentence 1 together (We Do), release for You Do
- Circulate -- check that key ideas survive the conversion

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: complete Section A only, using the underlined key words as a starting point. Replace "because" with b/c and "and" with +. That is enough for today
- Extra Notes: pair with a stronger student for the conversion step
EXTENDING PROMPT:
- Task: take a paragraph from a different chapter and convert it into notes. Then swap with a partner -- can they read your notes back as a paragraph?

TEACHER NOTES:
The extension (read notes back to recreate the paragraph) is a quality check. If notes can be read back meaningfully, they are good notes.

WATCH FOR:
- Students whose notes are too short -- lost key information
- Students whose notes are too long -- still copying the sentence
- Students inventing useful symbols -- celebrate and share with the class

[Literacy: We Do / You Do | VTLM 2.0: Guided + Independent Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Let's check. Your notes do not need to match mine exactly. What matters is that the key idea survives and the note is shorter than the sentence
- Sentence 1: "Sailor Sam, who shared Tom's berth, told him about Botany Bay and the seven seas."
-   Notes: [Sailor Sam (shared Tom's berth)  told Tom  Botany Bay + seven seas]
- Sentence 2: "Tom felt queasy as the ship rocked gently at anchor in Plymouth harbour."
-   Notes: [Tom queasy  ship rocking @ anchor Plymouth]
- Sentence 3: "Thomas had five children, twenty-seven grandchildren and forty-three great-grandchildren."
-   Notes: [Thomas  5 children, 27 grandchildren, 43 great-grandchildren]

DO:
- Reveal each model note
- Invite students to share their version and compare
- Celebrate efficient, still-readable notes

[Literacy: We Do / You Do Reveal | VTLM 2.0: Guided + Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: identifying Tom's feelings about the new ordeal -- thumbs?
- SC2: explaining craft like repetition or sensory detail -- thumbs?
- SC3: converting a sentence to notes with abbreviations and symbols -- thumbs?
- Turn and Tell: what is the most important change for Tom in these two chapters, and why?

DO:
- Thumbs check each SC
- The turn-and-tell asks for a change + justification -- push for evidence
- Preview: next lesson we read Chapters 20 and 21, and we begin planning a body paragraph for our information report

TEACHER NOTES:
The note-taking skill will be used again in Lesson 9 when students take notes from a non-fiction article. Flag this connection.

WATCH FOR:
- Students thumbs-down on SC3 -- check: is it knowing the symbols, or finding key words? Different reteach
- Students using SC3 skills spontaneously in other subjects -- celebrate this

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two printable resources today
- The ${WORKSHEET_RESOURCE.name} is for the We Do / You Do note-taking activity
- The ${ANSWER_KEY_RESOURCE.name} is for teacher reference

DO:
- Print the worksheet before the lesson (one per student)
- Print the answer key (teacher copy only)

TEACHER NOTES:
Section A scaffolds by underlining key words. Section B is independent. Answer key shows model notes; accept variations if the meaning is preserved.

[Literacy: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Chapters 18-19: Aboard the Scarborough -- Lesson 8";

  // Slide 1 -- Title
  titleSlide(
    pres,
    "Chapters 18-19",
    "Aboard the Scarborough",
    "Lesson 8  |  Year 5/6 Literacy",
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
      "We are learning to analyse how an author shows character under new pressures, and to take notes using keywords, phrases, abbreviations and symbols",
    ],
    [
      "I can identify what Tom feels as he faces a new ordeal",
      "I can explain how the author uses craft choices like repetition to create effect",
      "I can convert a sentence into notes using keywords, abbreviations and symbols",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // Slide 3 -- Vocab: impenetrable
  vocabSlide(
    pres,
    "impenetrable",
    "adjective",
    "Impossible to enter or pass through -- nothing can get in. The 'im-' prefix means 'not', so impenetrable = not able to be passed through.",
    "Below deck, Tom stared into the impenetrable darkness of the ship's hold.",
    NOTES_VOCAB_IMPENETRABLE,
    FOOTER
  );

  // Slide 4 -- Vocab: queasy
  vocabSlide(
    pres,
    "queasy",
    "adjective",
    "Feeling sick in your stomach -- nauseous, wobbly, unwell. Often caused by motion or something unpleasant.",
    "Tom felt queasy as the ship rocked gently at anchor in Plymouth harbour.",
    NOTES_VOCAB_QUEASY,
    FOOTER
  );

  // Slide 5 -- Reading Launch
  contentSlide(
    pres,
    "Teacher Read Aloud",
    C.PRIMARY,
    "Chapters 18-19",
    [
      "Reading Mode: Teacher Read Aloud",
      "Chapter 18: Tom is transported to Plymouth and boards the Scarborough",
      "Chapter 19: Thomas reflects on his five children and many descendants",
      "Focus: how does the author show Tom's feelings through sensory detail?",
    ],
    NOTES_READING,
    FOOTER
  );

  // Slide 6 -- Pause Point 1: Down, down, down
  quoteSlide(
    pres,
    "Pause Point 1",
    "Chapter 18 -- p. 89",
    "Down, down, down...",
    "p. 89",
    "Why has the author repeated 'down' three times? What does this make us feel? Where have we seen Tom go 'down' before?",
    NOTES_PAUSE1,
    FOOTER
  );

  // Slide 7 -- I Do: KPAS concept + modelled conversion
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Note-Taking with KPAS",
    "KPAS =\n\n- Keywords\n- Phrases\n- Abbreviations\n- Symbols\n\nFind the key information:\nWHO / WHAT DOING / WHEN / WHERE / WHY / HOW\n\nDrop small words:\nthe, a, and, is, are, that\n\nNotes are for YOU -- fast to write, easy to read back.",
    "Original sentence:\n\"In chains, Tom and the other prisoners are transported by wagon to Plymouth where the first fleet ships are anchored.\"\n\nConverted note:\n\"In chains, Tom + other prisoners -> transported by wagon -> Plymouth (First Fleet ships anchored)\"",
    NOTES_IDO_KPAS,
    FOOTER
  );

  // Slide 8 -- I Do: Abbreviations and Symbols reference
  modellingSlide(
    pres,
    "I Do -- Tools",
    "Abbreviations & Symbols",
    "Abbreviations:\n\n- b/c = because\n- w/ = with\n- w/o = without\n\nUse sparingly. You do not need to use them all.",
    "Symbols:\n\n=  means 'is' or 'equals' or 'means'\n+  means 'and'\narrow  means 'leads to', 'results in', 'cause and effect'\n\nExample:\n[T + J w/o food -> queasy]\nReads as: Tom and Jem, without food, felt queasy",
    NOTES_IDO_SYMBOLS,
    FOOTER
  );

  // Slides 9-10 -- CFU with reveal
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Find the Key Words",
      "Show Me Boards",
      "\"Once aboard the ship, Tom and the prisoners are ordered to strip and wash in half-barrels full of seawater using fatty cakes of carbolic soap while their hair is clipped short to remove lice.\"\n\nOn your whiteboard, write ONLY the key words.",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      const ansY = 3.9;
      slide.addShape("roundRect", {
        x: 0.5, y: ansY, w: 9, h: 1.16, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      });
      slide.addText("Key words:", {
        x: 0.75, y: ansY + 0.08, w: 8.5, h: 0.24,
        fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true, margin: 0,
      });
      slide.addText("Tom + prisoners, ordered strip + wash, half-barrels w/ carbolic soap + hair clipped short -> remove lice", {
        x: 0.75, y: ansY + 0.34, w: 8.5, h: 0.72,
        fontSize: 12, fontFace: FONT_B, color: C.WHITE, valign: "top", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // Slides 11-12 -- We Do / You Do with reveal
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.SECONDARY);
      addBadge(s, "We Do / You Do", { color: C.SECONDARY, w: 2.0 });
      addTitle(s, "Convert Sentences to Notes");

      const sentences = [
        { num: "1", text: "Sailor Sam, who shared Tom's berth, told him about Botany Bay and the seven seas.", color: C.PRIMARY },
        { num: "2", text: "Tom felt queasy as the ship rocked gently at anchor in Plymouth harbour.", color: C.SECONDARY },
        { num: "3", text: "Thomas had five children, 27 grandchildren and 43 great-grandchildren.", color: C.ACCENT },
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
          x: 1.35, y: sy + 0.08, w: 7.8, h: 0.52,
          fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
          fit: "shrink", shrinkText: true,
        });
        s.addText("Convert to notes: keywords + abbreviations + symbols.", {
          x: 1.35, y: sy + 0.62, w: 7.8, h: 0.36,
          fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true, valign: "middle", margin: 0,
        });
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO);
      return s;
    },
    (slide) => {
      const answers = [
        { y: CONTENT_TOP + 0.62, text: "[Sailor Sam (shared Tom's berth) -> told Tom -> Botany Bay + seven seas]" },
        { y: CONTENT_TOP + 1.80, text: "[Tom queasy -> ship rocking @ anchor, Plymouth]" },
        { y: CONTENT_TOP + 2.98, text: "[Thomas -> 5 children, 27 grandchildren, 43 great-grandchildren]" },
      ];
      answers.forEach((ans) => {
        slide.addShape("roundRect", {
          x: 0.7, y: ans.y, w: 8.5, h: 0.32, rectRadius: 0.06,
          fill: { color: C.BG_LIGHT },
        });
        slide.addText(ans.text, {
          x: 0.8, y: ans.y + 0.02, w: 8.2, h: 0.28,
          fontSize: 11, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
        });
      });
      slide.addNotes(NOTES_WEDO_REVEAL);
    }
  );

  // Slide 13 -- Closing
  closingSlide(
    pres,
    "What is the most important change for Tom in these two chapters, and why? Tell your partner.",
    [
      "I can identify what Tom feels as he faces a new ordeal",
      "I can explain how the author uses craft like repetition",
      "I can convert a sentence to notes using KPAS",
    ],
    NOTES_CLOSING
  );

  // Slide 14 -- Resources


  // ---------------------------------------------------------------------------
  // PDF 1 -- Note-Taking Worksheet
  // ---------------------------------------------------------------------------
  const ws = createPdf({ title: WORKSHEET_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Note-Taking -- Sentences to KPAS", {
    color: C.PRIMARY,
    subtitle: "Chapters 18-19: Aboard the Scarborough",
    lessonInfo: "Lesson 8 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(
    ws,
    "Notes keep the key meaning using Keywords, Phrases, Abbreviations and Symbols. Ask: WHO / WHAT DOING / WHEN / WHERE / WHY / HOW? Then drop small words like 'the', 'and', 'is'.",
    wsY,
    { color: C.PRIMARY }
  );

  wsY = addTipBox(
    ws,
    "Abbreviations:  b/c = because   |   w/ = with   |   w/o = without          Symbols:  =  is/means/equals   |   +  and   |   arrow  leads to / causes",
    wsY,
    { color: C.SECONDARY }
  );

  wsY = addSectionHeading(ws, "Section A: Key Words Highlighted", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "The key words are shown in CAPITALS. Convert each sentence to a note using those words, plus any abbreviations or symbols you think will help.", wsY);
  wsY += 4;

  wsY = addBodyText(ws, "1. SAILOR SAM, who SHARED TOM'S BERTH, TOLD him about BOTANY BAY and the SEVEN SEAS.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "2. TOM felt QUEASY as the SHIP ROCKED gently at ANCHOR in PLYMOUTH HARBOUR.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "3. THOMAS had 5 CHILDREN, 27 GRANDCHILDREN and 43 GREAT-GRANDCHILDREN.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Section B: Find the Key Words Yourself", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "No help this time. Find the key words, then convert each sentence to a note.", wsY);
  wsY += 4;

  wsY = addBodyText(ws, "4. Tom was reluctantly given a straw mattress and a thin blanket by the ship's steward.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "5. Because the ship was becalmed for three days, the prisoners became restless and hungry.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });
  wsY += 6;
  wsY = addBodyText(ws, "6. Thomas was a grazier and a magistrate in the colony after his convict days ended.", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });

  addPdfFooter(ws, "Lesson 8 | Chapters 18-19 | Note-Taking Worksheet");
  await writePdf(ws, WORKSHEET_PDF_PATH);

  // ---------------------------------------------------------------------------
  // PDF 2 -- Answer Key
  // ---------------------------------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Note-Taking -- Answer Key", {
    color: C.PRIMARY,
    subtitle: "Teacher Reference",
    lessonInfo: "Lesson 8 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(
    ak,
    "Marking note: accept any note that preserves the key meaning AND is clearly shorter than the original sentence. There are many valid conversions. Focus on whether key information survived.",
    akY,
    { color: C.ALERT }
  );

  akY = addSectionHeading(ak, "Section A -- Model Notes", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "1. [Sailor Sam (shared Tom's berth) -> told Tom -> Botany Bay + seven seas]", akY, { fontSize: 11 });
  akY = addBodyText(ak, "2. [Tom queasy -> ship rocking @ anchor, Plymouth]", akY, { fontSize: 11 });
  akY = addBodyText(ak, "3. [Thomas -> 5 children, 27 grandchildren, 43 great-grandchildren]", akY, { fontSize: 11 });
  akY += 8;

  akY = addSectionHeading(ak, "Section B -- Model Notes", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "4. [Tom -> reluctantly given straw mattress + thin blanket (from ship's steward)]", akY, { fontSize: 11 });
  akY = addBodyText(ak, "5. [Ship becalmed 3 days -> prisoners restless + hungry]", akY, { fontSize: 11 });
  akY = addBodyText(ak, "6. [Thomas = grazier + magistrate after convict days]", akY, { fontSize: 11 });
  akY += 10;

  akY = addSectionHeading(ak, "Common Student Errors", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "- Copying the whole sentence -- reteach the 'cross out small words' strategy", akY);
  akY = addBodyText(ak, "- Dropping key numbers or names -- remind students these are usually the most important key words", akY);
  akY = addBodyText(ak, "- Using symbols without meaning (adding emojis that don't aid recall) -- redirect to the core set", akY);
  akY = addBodyText(ak, "- Notes so short the meaning is lost -- test: read the notes aloud a day later. Still clear?", akY);

  addPdfFooter(ak, "Lesson 8 | Answer Key -- TEACHER COPY");
  await writePdf(ak, ANSWER_KEY_PDF_PATH);

  // ---------------------------------------------------------------------------
  // Write PPTX
  // ---------------------------------------------------------------------------
  const outName = path.join(OUT_DIR, "Tom_Lesson8_Chapters_18-19.pptx");
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
