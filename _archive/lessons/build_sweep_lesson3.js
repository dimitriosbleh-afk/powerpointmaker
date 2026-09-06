"use strict";

// Sweep Unit -- Lesson 3: Chapters 7-8 + Single Paragraph Outline (SPO)
// Week 1, Session 3, Grade 5/6 Literacy
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
  addWriteLine, addTwoColumnOrganiser,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 3;
const FOOTER = "Sweep | Lesson 3 | Week 1 | Year 5/6 Literacy";
const OUT_DIR = "output/Sweep_Lesson3_Chapters7_8";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const WORKSHEET_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "SPO Worksheet",
  "Student worksheet: single paragraph outline with topic sentence, supporting details, and concluding sentence."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Answer Key",
  "Teacher reference: model SPO responses for Chapters 7-8."
);
const RESOURCE_ITEMS = [WORKSHEET_RESOURCE, ANSWER_KEY_RESOURCE];
const WORKSHEET_PDF_PATH = path.join(OUT_DIR, WORKSHEET_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Lesson 3 today. We are reading Chapters 7 and 8 of Sweep
- Big events in these chapters -- Tom starts working as a chimney sweep and faces real danger
- Our writing focus builds on summary sentences. Today we learn the Single Paragraph Outline -- a way to plan a whole paragraph before writing it

DO:
- Display title slide as students settle
- Have copies of Sweep bookmarked at Chapter 7

TEACHER NOTES:
Session 3 of the Sweep unit. Chapters 7-8 are action-heavy and emotionally intense. Tom climbs his first chimney and rescues Jem when he gets stuck. The courage cloak metaphor returns. The SPO builds directly on the summary sentence skill from Lesson 1 -- it extends from one sentence to a planned paragraph.

WATCH FOR:
- Students who remember the courage cloak from Lesson 1 -- it returns in Chapter 7
- Students who are anxious about what happens to Tom -- reassure without spoiling

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands: reading Chapters 7-8 and learning to plan a paragraph using a Single Paragraph Outline
- Read the success criteria. SC1 is about the courage cloak metaphor -- what it reveals about Tom's character. SC2 is about how the author builds tension. SC3 is our writing skill -- planning a paragraph with a topic sentence, supporting details, and concluding sentence

DO:
- Choral read the LI, then each SC
- Briefly preview: "An SPO is a plan for one paragraph. It has three parts: a topic sentence, supporting details, and a concluding sentence"

TEACHER NOTES:
SC1 targets the recurring courage cloak motif. SC2 targets tension-building techniques (pacing, short sentences, sensory detail). SC3 introduces the SPO as a paragraph planning tool. The progression: SC1 is recognition, SC2 is analysis, SC3 is application to writing.

WATCH FOR:
- Students who think SPO means writing a full paragraph immediately -- clarify: "Today we PLAN the paragraph. The optional extension is writing it"
- Students who remember the courage cloak from Lesson 1 -- excellent recall

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_CONTORTED = `SAY:
- First vocabulary word: contorted. Read it with me: contorted [students repeat]
- Contorted means twisted or bent out of its normal shape. When something is contorted, it looks strained, uncomfortable, or distorted
- In Chapter 8, Jem's body is contorted as he gets stuck in the chimney. He is twisted into an unnatural position -- his limbs are bent in ways they should not be
- Ask: If I said someone's face was contorted with pain, what would it look like? [Screwed up, twisted, grimacing -- every muscle pulling in a different direction. The pain is so strong it changes the shape of the face]

DO:
- Display word, choral read, define, give examples
- Cold Call 2 students: describe something contorted (a face, a body, a piece of metal)
- Quick activity: "Show me a contorted face" [students make exaggerated expressions] -- this builds embodied understanding

TEACHER NOTES:
"Contorted" is a powerful descriptive word that connects directly to the chimney scene in Chapter 8. It conveys physical suffering vividly. The embodied activity (making a contorted face) is brief but effective for Year 5/6 -- it makes the word memorable through physical experience.

WATCH FOR:
- Students who think contorted just means bent -- extend: "Contorted implies force, strain, something pushed beyond its natural shape"
- Students who connect it to the physical reality of chimney sweeping -- excellent contextual understanding

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_VOCAB_SYMPATHETICALLY = `SAY:
- Second word: sympathetically. Read it with me: sympathetically [students repeat]
- Sympathetically means in a way that shows you understand and care about someone else's suffering or feelings. You feel with them
- In the story, Jem speaks sympathetically to Tom. He understands what Tom is going through because he has been through it himself. He does not mock or dismiss -- he acknowledges Tom's fear and pain
- Ask: What is the difference between doing something sympathetically and doing something kindly? [Kindness is being nice. Sympathy is deeper -- it means you understand the other person's experience because you have felt something similar or can imagine it]

DO:
- Display word, choral read, define, example
- Thumbs up/down: "The teacher marked the test sympathetically." Does this make sense? [It could -- if the teacher understood the student struggled and was gentle in their feedback]
- Turn and Talk: "When has someone treated you sympathetically? What did they do?"

TEACHER NOTES:
"Sympathetically" connects to the relationship between Jem and Tom. Jem's sympathy is rooted in shared experience -- he was once in Tom's position. This word also connects to character analysis (SC2) and the broader theme of compassion in harsh circumstances.

WATCH FOR:
- Students who confuse sympathy with empathy -- for this lesson, do not labour the distinction. Both involve understanding another's feelings
- Students who notice that Jem is the only character who treats Tom sympathetically -- important character observation

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_READING = `SAY:
- Chapters 7 and 8 today. These are intense chapters
- Chapter 7: Tom goes out with Master Jack and the boys to work in London at night. He watches the other boys clean chimneys, then has to climb one himself. The courage cloak comes back
- Chapter 8: something goes wrong in a chimney. Real danger
- Reading mode: student read aloud. Three pause points
- The author uses a lot of sensory detail in these chapters -- pay attention to what you can see, hear, smell, and feel in the writing

DO:
- Give students 30 seconds to find Chapter 7
- Select first reader
- Remind students: "If any part of this reading is difficult or upsetting, that is okay. The author writes honestly about hard situations"

TEACHER NOTES:
These chapters are the physical and emotional heart of the chimney sweep storyline. The sensory writing (soot, darkness, cold, fear) is intense and deliberate. The courage cloak metaphor returns on p.35, connecting back to Lesson 1. Jem emerges as an important character -- his relationship with Tom develops here.

WATCH FOR:
- Students disturbed by the description of chimney work -- the physical reality is confronting
- Students who recall the courage cloak from Lesson 1 -- strong memory and connection-making

[General: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause. "He tried to imagine the courage cloak again, all red and warm about him." Page 35
- The metaphor of the courage cloak recurs again. What is the author trying to tell us?
- Ask: Why does Tom need the courage cloak at this moment? [He is about to climb a chimney for the first time. He is terrified. The courage cloak is his way of finding strength when everything around him is frightening]
- Notice the colour detail: "all red and warm." Ask: Why red and warm? [Red suggests strength, passion, fire. Warm is the opposite of what Tom feels -- he is cold, scared. The cloak is what he wishes he had. It is imaginary comfort in a real nightmare]

DO:
- Display the quote
- Cold Call 3 students
- Connect back to Lesson 1: "We first saw this metaphor in Chapter 2. It is becoming Tom's signature move -- wrapping courage around himself when things are worst"

CFU CHECKPOINT:
Technique: Cold Call

Script:
- "I am going to ask three of you. No hands -- I choose. [Name 1]: what does the courage cloak represent for Tom?"
- "[Name 2]: why does the author keep bringing this metaphor back?"
- "[Name 3]: what would happen to Tom without the courage cloak -- what would he be like?"
- Scan for: students connecting the metaphor to Tom's resilience and the author's deliberate motif choices

PROCEED (>=80%): Students articulate the metaphor's purpose and its recurring nature. Continue reading.
PIVOT (<80%): Most likely issue -- students describe the cloak literally without connecting it to character development. Reteach: "Forget the cloak for a second. Tom is about to climb inside a chimney in the dark. He is ten years old. He is terrified. Now -- the cloak. It is not real. He imagines it. What does that tell us about Tom as a person? He creates his own courage. The author brings it back to show us this is WHO Tom is -- someone who finds strength even when there is none." Re-check: "In one sentence, what does the courage cloak tell us about Tom's character?"

TEACHER NOTES:
The courage cloak is now established as a recurring motif. Three appearances (Ch2 p.10, Ch7 p.27, Ch7 p.35) form a pattern that students should be tracking. This connects to SC1 and to the broader literary concept of motif as an authorial choice.

WATCH FOR:
- Students who track the three appearances of the courage cloak -- excellent literary awareness
- Students who connect the cloak to bravery vs recklessness -- "Tom is not reckless. He is frightened but acts anyway. That is the difference"

[General: Pause Point 1 | VTLM 2.0: Literary Analysis]`;

const NOTES_PAUSE2 = `SAY:
- Pause. "Finally Tom slept." Page 35
- Two questions. First: how does what we just read add to our understanding of the characters Jem and Master Jack?
- [Jem shows kindness and knowledge -- he explains things to Tom, shares what he knows. Master Jack is cruel -- the scars on Jem's back tell the real story. Master Jack punishes boys who try to escape]
- Second: how are things looking for Tom? [He is trapped. The cellar is freezing. The boys huddle together for warmth. But there is a small light -- Jem. Jem is a friend in a dark place]

DO:
- Display the quote
- Give 20 seconds think time
- Take responses for both questions separately
- Push for character evidence from the text

TEACHER NOTES:
Jem's characterisation deepens here -- he is both victim and mentor. The scars on his back are a powerful symbol of Master Jack's cruelty. The huddle for warmth is a small act of solidarity that contrasts with the institutional cruelty of their situation. Both details support SC2 (character development through description).

WATCH FOR:
- Students who focus on Master Jack's cruelty -- valid, but also prompt: "What about Jem? What kind of person is he?"
- Students who notice the scars as a warning and a character detail simultaneously

[General: Pause Point 2 | VTLM 2.0: Character Analysis]`;

const NOTES_PAUSE3 = `SAY:
- Final pause. "They were free." Page 44
- What is going on here? [Jem got stuck in a chimney. Tom climbed in to save him. Together they got out. "They were free" refers to escaping the chimney -- but there is a double meaning. Are they really free?]
- Ask: The author uses two words. "They were free." Why so short? Why not a longer, more detailed sentence? [After pages of tension, panic, and struggle, the short sentence is a release. The reader can breathe. The simplicity matches the relief. Less is more here]

DO:
- Display the quote
- Let students react -- this is an emotional moment
- Cold Call: "Are they REALLY free? Free from the chimney yes -- but free from Master Jack?"

CFU CHECKPOINT:
Technique: Finger Voting (1-5 scale)

Script:
- "Rate the tension level during the chimney rescue scene. 1 = calm, 5 = extreme tension. Show me: 3, 2, 1"
- Scan for: mostly 4s and 5s
- Follow up: "[Name], what specific detail made the tension highest for you?"
- Then: "Now rate Tom's situation OVERALL. 1 = hopeful, 5 = desperate"
- Scan for: expect 3s and 4s -- the rescue gives hope but they are still with Master Jack

PROCEED (>=80%): Students distinguish between the immediate relief and the ongoing entrapment. Move to writing focus.
PIVOT (<80%): Most likely issue -- students think "free" means truly free. Reteach: "Free from the chimney, yes. But look around. Where are they going back to? Master Jack's cellar. The author gives us one beautiful moment of freedom inside a larger prison. That contrast is deliberate." Re-check: "Thumbs up if you think Tom and Jem are truly free. Thumbs down if you think they are still trapped." [Expect mostly thumbs down]

TEACHER NOTES:
This is the emotional climax of the reading section. The rescue scene builds tension through pacing, sensory detail, and short sentences. The ambiguity of "free" is a sophisticated interpretive question. Students who see the double meaning are demonstrating strong inferential reading.

WATCH FOR:
- Students who see only the literal freedom -- push for the deeper reading
- Students who connect Tom's bravery in the rescue to the courage cloak -- excellent thematic connection
- Students who are emotionally engaged -- the rescue scene is designed to produce relief

[General: Pause Point 3 | VTLM 2.0: Deep Comprehension]`;

const NOTES_SPO_IDO = `SAY:
- Now our writing focus. In Lesson 1 we wrote summary sentences -- one sentence capturing a main idea. Today we level up to a whole paragraph, but we plan it first using a Single Paragraph Outline, or SPO
- An SPO has three parts: a topic sentence, supporting details, and a concluding sentence
- The topic sentence introduces the main idea. It tells the reader what the paragraph is about. It often answers who, what, when, where, why, or how
- Supporting details are evidence from the text that backs up the topic sentence. Usually 2-3 details
- The concluding sentence wraps up the paragraph and reminds the reader of the main idea without repeating the topic sentence word for word
- Watch me model one. My main idea is: "How does the description of London's streets reflect Tom's new life?"

DO:
- Display the SPO framework on the left
- Model the planning process step by step on the right
- Think aloud: "My topic sentence needs to introduce this idea. My supporting details come from the text. My concluding sentence connects back to the main idea"

MISCONCEPTIONS:
- Misconception: A concluding sentence just repeats the topic sentence
  Why: Students have been told concluding sentences "go back to the main idea" and interpret this as repeating
  Impact: Paragraphs feel circular and undeveloped. Students do not learn to synthesise
  Quick correction: "The concluding sentence reminds the reader of the main idea but uses DIFFERENT words and often adds a final thought. It closes the door -- it does not reopen it"

TEACHER NOTES:
The SPO is a paragraph planning tool, not a full paragraph. The I Do models the PLANNING process. The optional extension invites students to write the paragraph from their plan. All examples use Chapters 7-8 content. The who/what/when/where/why/how prompts connect back to the summary sentence technique from Lesson 1.

WATCH FOR:
- Students who want to skip the plan and write directly -- "The plan is the skill today. It makes the writing faster and more focused"
- Students who struggle to distinguish topic sentence from supporting detail -- "The topic sentence is the big idea. Supporting details are the evidence"

[General: I Do -- SPO | VTLM 2.0: Explicit Modelling]`;

const NOTES_WEDO = `SAY:
- Let's plan one together. Our main idea is: "Tom adapts to his new life as a chimney sweep"
- First, the topic sentence. With your partner, brainstorm: how would you introduce this idea in one sentence?
- [Take suggestions] One possibility: "Despite the harsh conditions, Tom begins to adapt to his terrifying new life as a chimney sweep"
- Now supporting details. What evidence from the text supports this idea? [Tom watches the other boys and learns. He follows Jem up the chimney. He endures the cold and the dark. He uses the courage cloak to find strength]
- Finally, a concluding sentence. Remember -- do not repeat the topic sentence. Wrap it up with a final thought
- One possibility: "Tom's ability to face each new horror shows a growing resilience that he will need for what lies ahead"

DO:
- Display the topic prompt
- Give 30 seconds for partner brainstorm on the topic sentence
- Take 3 responses before showing model
- Click to next slide to reveal the full model SPO after discussion

CFU CHECKPOINT:
Technique: Show Me Boards

Script:
- "Write a topic sentence for this idea on your whiteboard: 'Tom adapts to his new life as a chimney sweep.' 45 seconds. Show me"
- Scan for: sentences that introduce the MAIN IDEA (adaptation) with specific context (chimney sweep, harsh conditions). Sentences that just describe events are summaries, not topic sentences
- "A topic sentence tells the reader what the paragraph will be ABOUT. It is not a summary of events -- it is a statement of the main idea"

PROCEED (>=80%): Most boards show a clear topic sentence. Move to You Do.
PIVOT (<80%): Most likely issue -- students write a summary of events rather than a topic sentence. Reteach: "There is a difference. A summary says WHAT happened: 'Tom went up a chimney.' A topic sentence says what the paragraph is ABOUT: 'Tom begins to adapt to his new life.' The topic sentence is bigger -- it contains the whole paragraph's idea in one sentence." Re-check: "Rewrite. Start with 'Despite...' or 'Although...' to help frame it as an idea, not an event."

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Write only the topic sentence using the frame: "In Chapters 7-8, Tom shows ___ when he ___." Focus on identifying the main idea before worrying about supporting details
- Extra Notes: Students can use their notes from the reading discussion to find textual evidence

EXTENDING PROMPT:
- Task: Complete the full SPO, then write the paragraph from the plan. Aim for 4-5 sentences total. Ensure the concluding sentence adds a new thought rather than repeating the topic sentence

TEACHER NOTES:
The We Do uses different content from the I Do (I Do analysed London's streets, We Do analyses Tom's adaptation). The reveal shows a complete model SPO. The topic sentence is the most challenging part -- distinguishing it from a summary sentence is the key learning move.

WATCH FOR:
- Students who write event summaries instead of topic sentences -- the key distinction for this lesson
- Students who find strong supporting details from the text -- affirm and share

[General: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_YOUDO = `SAY:
- On the worksheet, you have a choice of main ideas to plan an SPO for
- Choose ONE: "Master Jack is cruel and uncaring" OR "Tom's resourcefulness saves Jem when he gets stuck in a chimney"
- FIRST: Write your topic sentence -- the big idea for your paragraph
- NEXT: Find 2-3 supporting details from Chapters 7-8. Use evidence from the text
- THEN: Write a concluding sentence that wraps up without repeating the topic sentence

DO:
- Distribute the Session 3 SPO Worksheet
- Circulate -- check that students write the topic sentence BEFORE finding details
- After 5 minutes: "You should have your topic sentence and at least one supporting detail"
- Conference with students who are struggling with the topic sentence

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the provided sentence frame for the topic sentence. Then find ONE supporting detail from the text and write it in your own words
- Extra Notes: Help these students locate a specific page reference for their supporting detail

EXTENDING PROMPT:
- Task: Complete the SPO for both main ideas (not just one). Then choose your stronger SPO and write the full paragraph (4-5 sentences)

TEACHER NOTES:
The You Do offers a choice of main ideas, both different from the I Do and We Do. "Master Jack is cruel" focuses on character analysis. "Tom's resourcefulness saves Jem" focuses on plot analysis. Both require evidence from the text. The optional paragraph extension bridges from planning to writing.

WATCH FOR:
- Students who copy text directly instead of using their own words for supporting details -- "Put the book down and tell me what happened. Now write THAT"
- Students who write a strong concluding sentence that adds insight -- share with the class

[General: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: I can explain what the courage cloak metaphor reveals about Tom -- thumbs? [scan]
- SC2: I can identify how the author builds tension in the chimney scenes -- thumbs? [scan]
- SC3: I can plan a paragraph using a topic sentence, supporting details, and concluding sentence -- thumbs? [scan]
- Turn and talk: What was the most tense moment in Chapters 7-8? What technique did the author use to create that tension?

DO:
- Run through each SC with thumbs check
- The turn-and-talk targets SC2 specifically -- tension-building techniques
- Preview: "Next lesson we read Chapters 9-10 and begin planning an information report about convict settlement"

TEACHER NOTES:
The closing targets both strands. Students who can identify a tense moment AND name a technique are demonstrating analytical reading. Students showing thumbs-down on SC3 may need a refresher on the topic-sentence/supporting-detail distinction next session.

WATCH FOR:
- Students who name specific techniques (short sentences, sensory detail, pacing) -- strong analytical language
- Students thumbs-down on SC3 -- the SPO is new, so some uncertainty is expected

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${WORKSHEET_RESOURCE.name} has the SPO planning template for your chosen main idea
- The ${ANSWER_KEY_RESOURCE.name} is for teacher reference

DO:
- Print the worksheet before the lesson (one per student)
- Print the answer key (teacher copy only)
- Click any resource card to open the PDF

TEACHER NOTES:
The worksheet provides structured space for the SPO components with guided prompts. The sentence frames in the enabling section are on the worksheet itself.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Sweep - Lesson 3 - Chapters 7-8 + SPO";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "Sweep",
    "Chapters 7-8 -- Into the Chimney",
    "Lesson 3  |  Week 1  |  Year 5/6 Literacy",
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
      "We are learning to analyse how an author uses recurring motifs and tension-building techniques, and to plan a paragraph using a Single Paragraph Outline",
    ],
    [
      "I can explain what the recurring courage cloak metaphor reveals about Tom",
      "I can identify how the author builds tension during the chimney scenes",
      "I can plan a paragraph with a topic sentence, supporting details, and concluding sentence",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Vocabulary: contorted
  // =========================================================================
  vocabSlide(
    pres,
    "contorted",
    "adjective",
    "Twisted or bent out of its normal shape. When something is contorted, it looks strained, uncomfortable, or distorted -- pushed beyond its natural position.",
    "Jem's body was contorted in the narrow chimney, his limbs twisted at painful angles as he struggled to move.",
    NOTES_VOCAB_CONTORTED,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Vocabulary: sympathetically
  // =========================================================================
  vocabSlide(
    pres,
    "sympathetically",
    "adverb",
    "In a way that shows you understand and care about someone else's suffering or feelings. You feel with them because you recognise their experience.",
    "Jem spoke sympathetically to Tom, knowing exactly what the boy was feeling because he had been through it all before.",
    NOTES_VOCAB_SYMPATHETICALLY,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- Reading Launch
  // =========================================================================
  contentSlide(
    pres,
    "Read Aloud",
    C.PRIMARY,
    "Chapters 7-8",
    [
      "Reading Mode: Student read aloud",
      "Ch 7: Tom goes to work in London's streets. He climbs his first chimney",
      "Ch 8: Jem gets stuck in a chimney. Tom climbs in to rescue him",
      "Focus: how does the author build tension? Watch for the courage cloak",
    ],
    NOTES_READING,
    FOOTER
  );

  // =========================================================================
  // SLIDE 6 -- Pause Point 1 (p.35)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 1", "Chapter 7 -- p. 35",
    "He tried to imagine the courage cloak again, all red and warm about him.",
    "p. 35",
    "The courage cloak is back. What is the author trying to tell us about Tom?",
    NOTES_PAUSE1, FOOTER
  );

  // =========================================================================
  // SLIDE 7 -- Pause Point 2 (p.35)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 2", "Chapter 7 -- p. 35",
    "Finally Tom slept.",
    "p. 35",
    "How does what we just read add to our understanding of Jem and Master Jack? How are things looking for Tom?",
    NOTES_PAUSE2, FOOTER
  );

  // =========================================================================
  // SLIDE 8 -- Pause Point 3 (p.44)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 3", "Chapter 8 -- p. 44",
    "They were free.",
    "p. 44",
    "What is going on here? Why does the author use such a short sentence?",
    NOTES_PAUSE3, FOOTER
  );

  // =========================================================================
  // SLIDE 9 -- I Do: SPO
  // =========================================================================
  modellingSlide(
    pres,
    "I Do",
    "Single Paragraph Outline (SPO)",
    "Three parts:\n\n1. Topic Sentence\nIntroduces the main idea.\nTells the reader what the\nparagraph is about.\n\n2. Supporting Details\n2-3 pieces of evidence from\nthe text.\n\n3. Concluding Sentence\nWraps up the main idea\nwithout repeating the\ntopic sentence.",
    "Model: How does the description\nof London reflect Tom's life?\n\nTopic sentence:\n\"The horror of London's streets\nmirrors the cruelty of Tom's\nnew existence.\"\n\nSupporting details:\n- Tenements, excrement, stench\n- Dark streets, no light\n- Tom scampers through filth\n\nConcluding sentence:\n\"The city itself seems to\nwork against Tom, as hostile\nas the people who control him.\"",
    NOTES_SPO_IDO,
    FOOTER
  );

  // =========================================================================
  // SLIDES 10-11 -- We Do: SPO (withReveal)
  // =========================================================================
  withReveal(
    () => contentSlide(
      pres,
      "We Do",
      C.SUCCESS,
      "Plan a Paragraph Together",
      [
        "Main idea: Tom adapts to his new life as a chimney sweep",
        "Step 1: Write a topic sentence on your whiteboard",
        "Step 2: Find 2-3 supporting details from the text",
        "Step 3: Draft a concluding sentence",
      ],
      NOTES_WEDO,
      FOOTER
    ),
    (slide) => {
      const ansY = 3.70;
      slide.addShape("roundRect", {
        x: 0.5, y: ansY, w: 9, h: 1.30, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      });

      slide.addShape("roundRect", {
        x: 0.7, y: ansY + 0.08, w: 1.5, h: 0.28, rectRadius: 0.08,
        fill: { color: C.WHITE },
      });
      slide.addText("Model SPO", {
        x: 0.7, y: ansY + 0.08, w: 1.5, h: 0.28,
        fontSize: 10, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      slide.addText("TS: \"Despite the harsh conditions, Tom begins to adapt to his terrifying new life as a chimney sweep.\"\nSD: Watches other boys and learns | Follows Jem up the chimney | Uses courage cloak to find strength\nCS: \"Tom's ability to face each new horror shows a growing resilience that he will need for what lies ahead.\"", {
        x: 2.3, y: ansY + 0.06, w: 7.0, h: 1.16,
        fontSize: 11.5, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_WEDO);
    }
  );

  // =========================================================================
  // SLIDE 12 -- You Do: SPO
  // =========================================================================
  contentSlide(
    pres,
    "You Do",
    C.ACCENT,
    "SPO -- Independent Practice",
    [
      "Choose ONE main idea:",
      "FIRST: Write your topic sentence",
      "NEXT: Find 2-3 supporting details from Chapters 7-8",
      "THEN: Write a concluding sentence that adds a final thought",
    ],
    NOTES_YOUDO,
    FOOTER,
    (slide, layoutGuide) => {
      const tipY = layoutGuide.panelTopPadded + 0.1;
      addCard(slide, layoutGuide.rightX, tipY, layoutGuide.rightW, 2.5, {
        strip: C.PRIMARY, fill: C.BG_CARD,
      });
      slide.addText("Choose One", {
        x: layoutGuide.rightX + 0.15, y: tipY + 0.08, w: 3.8, h: 0.28,
        fontSize: 11, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
      });
      slide.addText([
        { text: "A. Master Jack is cruel and uncaring", options: { bullet: true, breakLine: true, fontSize: 13, color: C.CHARCOAL, bold: true } },
        { text: "OR", options: { breakLine: true, fontSize: 11, color: C.MUTED, align: "center" } },
        { text: "B. Tom's resourcefulness saves Jem when he gets stuck", options: { bullet: true, fontSize: 13, color: C.CHARCOAL, bold: true } },
      ], {
        x: layoutGuide.rightX + 0.15, y: tipY + 0.48, w: 3.8, h: 1.8,
        fontFace: FONT_B, valign: "top", margin: 0, paraSpaceAfter: 8,
      });
    }
  );

  // =========================================================================
  // SLIDE 13 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    "What was the most tense moment in Chapters 7-8? What technique did the author use to create that tension?",
    [
      "I can explain what the recurring courage cloak metaphor reveals about Tom",
      "I can identify how the author builds tension during the chimney scenes",
      "I can plan a paragraph with a topic sentence, supporting details, and concluding sentence",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 14 -- Resources
  // =========================================================================


  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: SPO Worksheet ------------------------------------------------
  const ws = createPdf({ title: WORKSHEET_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Single Paragraph Outline (SPO)", {
    color: C.NAVY,
    subtitle: "Sweep -- Chapters 7-8",
    lessonInfo: "Sweep | Lesson 3 | Week 1 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "An SPO plans ONE paragraph with three parts:\n1. Topic Sentence -- introduces the main idea (what the paragraph is ABOUT)\n2. Supporting Details -- 2-3 pieces of evidence from the text\n3. Concluding Sentence -- wraps up without repeating the topic sentence", wsY, { color: C.NAVY });

  wsY = addSectionHeading(ws, "Choose ONE main idea:", wsY, { color: C.NAVY });
  wsY = addBodyText(ws, "A. Master Jack is cruel and uncaring", wsY, { bold: true, fontSize: 12 });
  wsY = addBodyText(ws, "B. Tom's resourcefulness saves Jem when he gets stuck in a chimney", wsY, { bold: true, fontSize: 12 });
  wsY += 4;
  wsY = addBodyText(ws, "I chose: ____", wsY, { fontSize: 11 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Topic Sentence", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Sentence frame (optional): \"In Chapters 7-8, [character] shows ___ when ___\"", wsY, { fontSize: 10, italic: true });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 28 });
  wsY += 6;

  wsY = addSectionHeading(ws, "Supporting Details (find 2-3 from the text)", wsY, { color: C.SECONDARY });
  wsY = addWriteLine(ws, "Detail 1:", wsY, {});
  wsY += 4;
  wsY = addWriteLine(ws, "Page:", wsY, {});
  wsY += 6;
  wsY = addWriteLine(ws, "Detail 2:", wsY, {});
  wsY += 4;
  wsY = addWriteLine(ws, "Page:", wsY, {});
  wsY += 6;
  wsY = addWriteLine(ws, "Detail 3:", wsY, {});
  wsY += 4;
  wsY = addWriteLine(ws, "Page:", wsY, {});
  wsY += 8;

  wsY = addSectionHeading(ws, "Concluding Sentence", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Remember: wrap up the main idea with a DIFFERENT wording. Add a final thought.", wsY, { fontSize: 10, italic: true });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 28 });
  wsY += 8;

  wsY = addSectionHeading(ws, "Extension: Write the full paragraph from your SPO", wsY, { color: C.ACCENT });
  wsY = addLinedArea(ws, wsY, 6, { lineSpacing: 26 });

  addPdfFooter(ws, "Sweep | Lesson 3 | SPO Worksheet");

  // --- PDF 2: Answer Key ---------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "SPO -- Answer Key", {
    color: C.ALERT,
    subtitle: "Teacher Reference -- Chapters 7-8",
    lessonInfo: "Sweep | Lesson 3 | Week 1 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Accept any SPO where: (1) the topic sentence states a main idea (not just an event), (2) supporting details come from the text with specific evidence, and (3) the concluding sentence wraps up without repeating the topic sentence.", akY, { color: C.ALERT });

  akY = addSectionHeading(ak, "Option A: Master Jack is cruel and uncaring", akY, { color: C.NAVY });
  akY = addBodyText(ak, "Model Topic Sentence: \"Master Jack's treatment of the boys reveals a man driven entirely by greed and indifference to their suffering.\"", akY);
  akY = addBodyText(ak, "Supporting Details:", akY, { bold: true, fontSize: 10 });
  akY = addBodyText(ak, "- He orders Jem to show his scars as a warning to Tom (p.33) -- using fear and physical punishment to control", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- The boys sleep in a freezing cellar with no blankets (p.35) -- their basic needs are ignored", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- He sends young boys up dangerous chimneys for profit (p.27-28) -- their lives are less important than his income", akY, { fontSize: 10 });
  akY = addBodyText(ak, "Model Concluding Sentence: \"Every detail of Master Jack's world shows that the boys exist only to serve his business, never as children deserving of care.\"", akY);
  akY += 8;

  akY = addSectionHeading(ak, "Option B: Tom's resourcefulness saves Jem", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "Model Topic Sentence: \"When Jem becomes trapped in a chimney, Tom's quick thinking and bravery transform him from a frightened newcomer into a rescuer.\"", akY);
  akY = addBodyText(ak, "Supporting Details:", akY, { bold: true, fontSize: 10 });
  akY = addBodyText(ak, "- Tom hears Jem's cries and immediately climbs into the chimney after him (p.40) -- he acts without hesitation", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Tom uses what he has learned from watching the other boys to navigate the chimney (p.41) -- he applies knowledge under pressure", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Together they dislodge Jem and escape (p.43-44) -- Tom's courage saves a life", akY, { fontSize: 10 });
  akY = addBodyText(ak, "Model Concluding Sentence: \"Tom's rescue of Jem proves that courage is not the absence of fear but the decision to act despite it -- the courage cloak made real.\"", akY);

  addPdfFooter(ak, "Sweep | Lesson 3 | Answer Key -- TEACHER COPY");

  // --- Write all files ------------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Sweep_Lesson3.pptx` }),
    writePdf(ws, WORKSHEET_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log(`PPTX written to ${OUT_DIR}/Sweep_Lesson3.pptx`);
  console.log(`Done: ${WORKSHEET_RESOURCE.name}`);
  console.log(`Done: ${ANSWER_KEY_RESOURCE.name}`);
}

build().catch(console.error);
