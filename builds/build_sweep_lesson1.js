"use strict";

// Sweep Unit -- Lesson 1: Chapters 1-3 + Summary Sentences
// Week 1, Session 1, Grade 5/6 Literacy
// Novel: "Sweep" by Jonathan Auxier
// Note: Chapter 2 contains content about death (Pa dies).

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

const SESSION_NUMBER = 1;
const FOOTER = "Sweep | Lesson 1 | Week 1 | Year 5/6 Literacy";
const OUT_DIR = "output/Sweep_Lesson1_Chapters1_3";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const WORKSHEET_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Summary Sentences Worksheet",
  "Student worksheet: write summary sentences using who, what, where, when, why, how."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Answer Key",
  "Teacher reference: model summary sentences for Chapters 1-3."
);
const RESOURCE_ITEMS = [WORKSHEET_RESOURCE, ANSWER_KEY_RESOURCE];
const WORKSHEET_PDF_PATH = path.join(OUT_DIR, WORKSHEET_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome back to Term 2. We are starting a new novel study today -- Sweep by Jonathan Auxier
- This is a story set in two time periods: one follows a young boy named Tom in historical England, the other follows a 90-year-old man named Thomas in Australia
- We will also be learning a new writing skill this term -- summary sentences. This is a brand new style of writing for us

DO:
- Display title slide as students settle
- Distribute copies of the novel
- If students have not encountered the novel before, briefly show the cover and let them examine it

TEACHER NOTES:
First session of Term 2, first lesson of the Sweep unit. Students are meeting a new novel and a new writing style (summary sentences). Assume mixed readiness -- do not presume familiarity with the text, the author, or the summary sentence form.

WATCH FOR:
- Students who are excited about a new novel -- channel that into careful reading
- Students who may need reassurance about Term 2 expectations -- keep the tone warm and welcoming

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands today: reading and responding to the first three chapters, and learning to write summary sentences
- Read the success criteria. SC1 is about identifying what the author does to shape how we feel. SC2 is about explaining what those choices do to the reader. SC3 is about writing summary sentences using question prompts
- Some of you may have written summaries before. Today we are learning a specific technique for doing it well

DO:
- Choral read the LI, then each SC
- Briefly clarify: a summary sentence captures the main idea in one clear sentence

TEACHER NOTES:
SC1 targets recognition of authorial techniques (imagery, metaphor, simile). SC2 targets analysis of effect on the reader. SC3 targets the writing skill -- summary sentences using who/what/where/when/why/how. SC1 is the floor: noticing what the author does. SC2 extends to explaining purpose. SC3 introduces the new writing form.

WATCH FOR:
- Students unsure about literary devices -- reassure: "We will learn to spot these together as we read"
- Students who think summary means retelling everything -- "A summary sentence captures the MAIN idea, not every detail"

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_GRASPED = `SAY:
- Our first vocabulary word: grasped. Read it with me: grasped [students repeat]
- Grasped means to take hold of something firmly, to grip it tightly. But it also has a second meaning: to understand something
- In the novel, Tom grasped the bars of the stocks where his father was held. He physically gripped them
- Ask: If I said "She grasped the concept quickly," am I talking about physical holding or understanding? [Understanding -- grasped can mean to get hold of an idea in your mind]

DO:
- Display word, choral read, define, give both meanings
- Cold Call 2 students: use "grasped" in a sentence -- one physical, one understanding
- Thumbs up/down: "He grasped the rope tightly." Is this physical or understanding? [Physical -- thumbs up for physical]

TEACHER NOTES:
"Grasped" is a Tier 2 word with dual meaning (physical grip and cognitive understanding). Pre-teaching it before reading helps students access both uses in the novel and in their own writing. The dual-meaning nature makes it a strong vocabulary choice for Year 5/6.

WATCH FOR:
- Students who only know the physical meaning -- the understanding meaning is equally important and more useful in academic writing
- Students who confuse grasped with grabbed -- "Grasped suggests a firm, purposeful hold. Grabbed is quicker and less controlled"

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_VOCAB_LINGERED = `SAY:
- Second word: lingered. Read it with me: lingered [students repeat]
- Lingered means to stay somewhere longer than necessary, to be slow to leave. There is a feeling of not wanting to go, of holding on to a moment
- In the story, the smell of the fire lingered in the air. It did not disappear quickly -- it stayed and hung around
- Ask: What is the difference between staying and lingering? [Lingering suggests you are staying by choice, slowly, almost reluctantly. Staying is neutral. Lingering has emotion attached to it]

DO:
- Display word, choral read, define, example
- Turn and Talk: "Tell your partner about a time something lingered -- a smell, a feeling, a memory"
- Take 2 responses

TEACHER NOTES:
"Lingered" connects to the novel's atmospheric writing. The author uses sensory language extensively, and "lingered" captures the way experiences stay with characters. It is also useful for students' own descriptive writing.

WATCH FOR:
- Students who think lingered only applies to people -- it applies to smells, feelings, memories, sounds, and thoughts too
- Students who give strong sensory examples -- affirm: "That is exactly the kind of moment an author would describe as lingering"

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_READING = `SAY:
- We are reading Chapters 1 to 3 today. Three chapters, three very different scenes
- Chapter 1: we meet a 90-year-old man named Thomas who sees something strange on his lawn
- Chapter 2: we jump back in time to meet a young boy named Tom and his father. This chapter deals with some serious content -- Pa is unwell and passes away. If any of this feels difficult, that is okay
- Chapter 3: Tom faces a big change in his life
- Reading mode: I will read Chapter 1, then student readers for Chapters 2 and 3
- We have four pause points across the three chapters

DO:
- Give students 30 seconds to find the first page
- Begin reading Chapter 1 aloud (teacher read)
- Plan reader rotations for Chapters 2 and 3

SENSITIVITY ADVISORY:
- What it is: Chapter 2 describes Pa being placed in the stocks for blasphemy. He becomes ill and dies. Tom witnesses this
- Framing language: "This chapter deals with something serious. A character we care about passes away. If this is hard to read about, that is completely understandable"
- Watch for: Students who become withdrawn, tearful, or agitated. Students who have experienced loss may find this content triggering
- Protocol: If a student becomes distressed, acknowledge their feelings quietly and privately. Do not require them to continue reading aloud. Follow the school's wellbeing referral process if needed

TEACHER NOTES:
The dual timeline structure (old Thomas in Australia, young Tom in England) may initially confuse some students. Chapter 1 is set in Australia. Chapter 2 jumps to historical England. Flag this shift explicitly when it happens. The death content in Chapter 2 is handled with gravity by the author -- follow that tone.

WATCH FOR:
- Students confused by the time shift between Chapter 1 and Chapter 2 -- "The author is telling two stories that are connected. We will discover how as we read"
- Students affected by the death content -- see sensitivity advisory above

[General: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause here. "If there was a ghost down there, he'd find it." Page 2
- What is going on here? [A 90-year-old man named Thomas sees something on his lawn at night. He decides to go and investigate. He is brave -- or stubborn -- enough to go looking for a ghost]
- Notice the author's imagery on this page: "shining autumn gold" and the similes -- "Gumtrees looked like old, bent men" and "like girls with long white limbs, reaching up into the sky"
- Ask: Why does the author describe the trees as looking like people? [It creates an eerie, alive feeling. The landscape feels watched. It sets up the idea that something supernatural might be happening]

DO:
- Display the quote
- Give 15 seconds think time
- Cold Call 3 students for initial responses
- Draw attention to the imagery on the page

TEACHER NOTES:
This opening establishes the Australian setting and the supernatural element. The similes personify the landscape, creating atmosphere. This is a good early opportunity to notice authorial technique without heavy analysis.

WATCH FOR:
- Students who are already hooked by the ghost element -- excellent engagement
- Students who notice the Australian setting (gumtrees) -- connect: "This story moves between Australia and England. We will see why"

[General: Pause Point 1 | VTLM 2.0: Active Reading]`;

const NOTES_PAUSE2 = `SAY:
- Pause. "The sightless eye stared at nothing." Page 6
- How is the author making you feel right now towards Tom and Pa? [Sympathy, sadness, maybe anger at the unfairness. Pa was punished for printing a pamphlet. Tom is trying to care for his father. The "sightless eye" suggests Pa is very unwell or dying]
- The author chose the word "sightless" rather than "blind" or "closed." Ask: What is the effect of "sightless"? [It sounds more final, more permanent. "Blind" suggests the eye cannot see. "Sightless" suggests there is nothing behind it anymore]

DO:
- Display the quote, read aloud with gravity
- Give 20 seconds think time
- Think-Pair-Share: 30 seconds think, 30 seconds pair, then share

CFU CHECKPOINT:
Technique: Think-Pair-Share

Script:
- "Think for 30 seconds: how is the author making you feel towards Tom and Pa right now? What specific words or phrases create that feeling?"
- "Share with your partner for 30 seconds"
- "I am selecting pairs. [Name], what did your pair discuss?"
- Scan for: students identifying emotional response AND connecting it to specific word choices

PROCEED (>=80%): Most pairs name an emotion and connect it to a word choice. Continue reading.
PIVOT (<80%): Most likely issue -- students describe their feeling but cannot point to specific words causing it. Reteach: "Let me show you. The author wrote 'sightless eye stared at nothing.' Three word choices: sightless, stared, nothing. Each one adds emptiness. That is the author CHOOSING words to make you feel a certain way." Re-check: "Find one more word on this page that makes you feel something. What feeling does it create?"

SENSITIVITY ADVISORY:
- What it is: Pa is dying in the stocks. Tom is witnessing his father's decline
- Framing language: "This is a heavy moment in the story. The author writes about it honestly because it matters to Tom's character"
- Watch for: Visible distress, withdrawal, or students who stop reading along
- Protocol: Acknowledge quietly. Offer a brief break if needed. Do not single out affected students

TEACHER NOTES:
This is the emotional core of Chapter 2. Pa's death is the inciting incident that drives Tom's story forward. The author's word choices here are deliberate and worth noticing -- this connects directly to SC1 (identifying authorial choices) and SC2 (explaining their effect).

WATCH FOR:
- Students who are affected by the content -- handle with care (see sensitivity advisory)
- Students who notice the contrast between Ch1 (old Thomas, relatively calm) and Ch2 (young Tom, devastating) -- this is the dual timeline at work
- Students who use emotional language in their responses -- affirm: "That is exactly what the author intended"

[General: Pause Point 2 | VTLM 2.0: Deep Comprehension]`;

const NOTES_PAUSE3 = `SAY:
- Pause. "A cloak made out of courage would be thick and warm and keep him safe..." Page 10
- The author is using a metaphor to describe courage. Why might she have chosen this?
- Ask: What does comparing courage to a cloak tell us? [A cloak is something you wrap around yourself for protection. Courage is not just a feeling -- it is something Tom actively puts on, like armour. It keeps him warm and safe in a cold, dangerous world]
- This metaphor will come back. Remember it

DO:
- Display the quote
- Cold Call 2 students
- Emphasise: this is a METAPHOR -- courage IS a cloak, not LIKE a cloak
- Note to students: "Keep this image in your mind. The author brings it back later"

TEACHER NOTES:
The courage cloak is a recurring motif in the novel. Flagging it now prepares students to track its reappearance. The distinction between metaphor (courage IS a cloak) and simile (courage is LIKE a cloak) is worth noting briefly but not belaboring in this first session.

WATCH FOR:
- Students who confuse metaphor and simile -- brief clarification is fine but do not derail into a full literary devices lesson
- Students who connect the cloak to Tom's situation -- he needs courage because he has lost everything

[General: Pause Point 3 | VTLM 2.0: Literary Analysis]`;

const NOTES_PAUSE4 = `SAY:
- Final pause. "Now be quiet." Page 13
- How are things looking for Tom now? [Not good. He has been orphaned. The magistrate is sending him to the workhouse. Mr Tupper is not kind to him. Tom has no say in what happens to him. Adults are making decisions about his life and telling him to be quiet]
- Ask: What does the command "Now be quiet" reveal about how adults treat Tom? [They do not see him as a person with feelings or rights. He is a problem to be dealt with, not a child to be cared for. The short, blunt command shows no warmth]

DO:
- Display the quote
- Finger voting: "How would you rate Tom's situation right now? 1 = hopeful, 5 = desperate" [Expect mostly 4s and 5s]
- Take 2-3 responses explaining their rating

CFU CHECKPOINT:
Technique: Finger Voting (1-5 scale)

Script:
- "Rate Tom's situation right now. 1 finger means hopeful, 5 fingers means desperate. Show me on my count: 3, 2, 1, show"
- Scan for: mostly 4s and 5s. Students who show 1 or 2 may not have understood the gravity of Tom's situation
- "I see mostly 4s and 5s. [Name], why did you choose [number]?"
- Follow up: "What specific detail from the chapter makes you feel that way?"

PROCEED (>=80%): Most students rate 4-5 and can justify with textual evidence. Move to writing focus.
PIVOT (<80%): Most likely issue -- students rate Tom's situation as hopeful because he got breakfast from Mrs Wilson. Reteach: "Mrs Wilson gave him breakfast, yes. But where is he going? The workhouse. Who chose this for him? The magistrate. Did anyone ask Tom what he wanted? No. One kind moment does not change the overall direction." Re-check: "Knowing he is heading to the workhouse, what number would you give now?"

TEACHER NOTES:
This closes the reading section. Tom's powerlessness is a key theme -- adults control his fate. The transition from reading to writing happens after this slide. Students should carry their understanding of the chapters into the summary sentence work.

WATCH FOR:
- Students who focus only on Mrs Wilson's kindness -- redirect to the bigger picture
- Students who are already predicting what happens next -- excellent engagement, but hold predictions for now

[General: Pause Point 4 | VTLM 2.0: Critical Thinking]`;

const NOTES_IDO = `SAY:
- Now our writing focus. Summary sentences -- this is a new skill for us this term
- A summary sentence captures the main idea of something you have read in one clear, concise sentence
- The trick is using question prompts to help you: who, what doing, when, where, why, how. You do not need all of them -- just the ones that matter for your summary
- Watch me model one. I am going to summarise Chapter 1
- First, I brainstorm: Who? Thomas, a 90-year-old man. What doing? Sees something on his lawn at night. When? At night. Where? On his farm in Australia. Why? He notices something strange. How? He decides to go investigate
- Now I choose what matters most and write: "One night, Thomas, a 90-year-old farmer, spots a mysterious figure on his lawn and decides to investigate despite his age"
- Notice: I started with the "when" detail, included who, what, and where. I did not need every question word

DO:
- Display the brainstorm framework (who/what/where/when/why/how)
- Model the brainstorm aloud, writing key words on the slide
- Model writing the summary sentence, thinking aloud about word choices
- Point out: "I started with 'when' to set the scene. You can start with any question word"

MISCONCEPTIONS:
- Misconception: A summary means retelling every event in order
  Why: Students confuse summary with retelling. Retelling is sequential; summarising is selective
  Impact: Students write long, unfocused paragraphs instead of concise sentences
  Quick correction: "A summary picks out the MAIN idea. If you are writing more than two lines, you are retelling, not summarising. Ask yourself: what is the ONE big thing that happened?"

TEACHER NOTES:
This is the first exposure to summary sentences as a formal writing technique. The brainstorm-then-write process is the key scaffold. Modelling with Chapter 1 keeps it connected to the reading. The optional "start with when" prompt gives students a concrete entry point without being prescriptive.

WATCH FOR:
- Students who want to include every detail -- "Pick the main idea, not every event"
- Students who write a sentence that is too vague (e.g. "Thomas did something") -- "Add specific detail. What did he see? Where was he?"

[General: I Do -- Summary Sentences | VTLM 2.0: Explicit Modelling]`;

const NOTES_WEDO = `SAY:
- Let's try one together. I want you to summarise this: "Why was Tom's father in the stocks?"
- First, brainstorm with your partner. Use the question prompts: who, what, when, where, why
- Who? [Pa / Tom's father] What happened? [He was put in the stocks] Why? [He printed a pamphlet about the King -- blasphemy] Where? [In the town]
- Now write a summary sentence on your whiteboard. You have 60 seconds
- One possibility: "Pa was placed in the stocks for a week after printing a pamphlet that criticised the King, which was considered blasphemy"
- Compare yours to mine. Did you capture the main idea?

DO:
- Display the summary prompt
- Give 30 seconds for partner brainstorm
- Give 60 seconds for whiteboard writing
- Take 3 responses before revealing model answer
- Click to next slide to reveal the model answer after students have responded

CFU CHECKPOINT:
Technique: Show Me Boards

Script:
- "Write your summary sentence on your whiteboard. 60 seconds. Hold up on my signal: 3, 2, 1, show"
- Scan for: sentences that include WHO (Pa), WHAT (put in stocks), and WHY (blasphemy/pamphlet). These three elements are essential
- "I can see most of you have included who and what. Check: did you include WHY he was in the stocks? That is the key detail"

PROCEED (>=80%): Most boards show a sentence with who, what, and why. Move to You Do.
PIVOT (<80%): Most likely issue -- students write a retelling rather than a single sentence, or miss the WHY. Reteach: "A summary sentence is ONE sentence. If you have written three sentences, pick the most important one. And the WHY matters -- it is not just that Pa was in the stocks, it is WHY: blasphemy. That word tells us about the time period and the injustice." Re-check: "Rewrite in one sentence. Include who, what, and why. 30 seconds."

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the sentence frame: "___ was placed in the stocks because ___." Fill in the blanks using information from the chapter
- Extra Notes: Students who struggle can refer back to their chapter notes or the text

EXTENDING PROMPT:
- Task: Write two summary sentences -- one for Chapter 2 and one for Chapter 3. Each must start with a different question word (e.g. one starts with "when," the other with "after")

TEACHER NOTES:
The We Do uses different content from the I Do (I Do summarised Ch1, We Do summarises a specific event from Ch2). The reveal shows one model answer but emphasises that multiple correct versions exist. The key check is whether students captured the essential information.

WATCH FOR:
- Students who write more than one sentence -- "Trim it down. One sentence, main idea only"
- Students who miss "blasphemy" as the reason -- "Why was he punished? What word describes what he did?"
- Students whose sentence captures who-what-why clearly -- affirm and share

[General: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_YOUDO = `SAY:
- On the worksheet, you have three summary tasks from the chapters we read today
- First: brainstorm who, what, where, when, why, how for each prompt. Write key words only -- not full sentences yet
- Next: write a summary sentence for each prompt using your brainstorm notes
- Then: check each sentence -- does it capture the main idea in one clear sentence? Could someone who has not read the chapter understand the key point?

DO:
- Distribute the Session 1 Summary Sentences Worksheet
- Circulate -- check that students brainstorm BEFORE writing the sentence
- After 5 minutes: "You should be writing your second summary sentence by now"
- Conference with students who are struggling -- help them identify the main idea first

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Complete only the first summary prompt using the sentence frame provided on the worksheet. Focus on identifying who, what, and why before writing
- Extra Notes: Oral rehearsal helps -- have the student say their sentence aloud before writing

EXTENDING PROMPT:
- Task: After completing all three summaries, write a fourth summary sentence that captures the OVERALL main idea across all three chapters in a single sentence. This requires synthesising across chapters rather than summarising a single event

TEACHER NOTES:
The You Do prompts are different from the I Do and We Do content to ensure genuine transfer. Students summarise: what happens to Tom after his father dies, and what important decision Tom makes. These require selecting the main idea from complex events. The optional "start with when" prompt is on the worksheet as a scaffold, not a requirement.

WATCH FOR:
- Students who retell instead of summarise -- redirect: "One sentence. Main idea"
- Students who write strong, concise summaries -- share with the class as models
- Students who finish early and write the synthesis sentence -- this demonstrates deeper understanding

[General: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: I can identify the authorial choices the author uses to shape how we feel -- thumbs? [scan]
- SC2: I can explain the effect of at least one authorial choice on the reader -- thumbs? [scan]
- SC3: I can write a summary sentence using who, what, where, when, why, how question prompts -- thumbs? [scan]
- Turn and talk: What was the most powerful moment across Chapters 1-3? Can you summarise it in one sentence?

DO:
- Run through each SC with thumbs check
- The turn-and-talk combines reading comprehension with the new writing skill
- Preview: "Next lesson we read Chapters 4-6 and continue developing our writing skills"

TEACHER NOTES:
The closing integrates both strands. Students who can summarise a powerful moment in one sentence are demonstrating both reading comprehension and the new writing technique. Students showing thumbs-down on SC3 may need additional modelling in the next session.

WATCH FOR:
- Students thumbs-down on SC3 -- summary sentences are new, so some uncertainty is expected. Reassure: "We will keep practising this"
- Students who name the courage cloak metaphor as the most powerful moment -- excellent literary sensitivity

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- One worksheet and one answer key today
- The ${WORKSHEET_RESOURCE.name} has the three summary prompts plus brainstorming space
- The ${ANSWER_KEY_RESOURCE.name} is for teacher reference

DO:
- Print the worksheet before the lesson (one per student)
- Print the answer key (teacher copy only)
- Click any resource card to open the PDF

TEACHER NOTES:
The worksheet scaffolds the brainstorm-then-write process with structured space for question-word notes before the summary sentence line.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Sweep - Lesson 1 - Chapters 1-3 + Summary Sentences";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "Sweep",
    "Chapters 1-3 -- A Ghost, a Father, a Workhouse",
    "Lesson 1  |  Week 1  |  Year 5/6 Literacy",
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
      "We are learning to identify and analyse how an author uses language to shape the reader's response, and to write summary sentences that capture the main idea",
    ],
    [
      "I can identify the authorial choices the author uses to shape how we feel",
      "I can explain the effect of at least one authorial choice on the reader",
      "I can write a summary sentence using who, what, where, when, why, how question prompts",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Vocabulary: grasped
  // =========================================================================
  vocabSlide(
    pres,
    "grasped",
    "verb",
    "To take hold of something firmly, to grip it tightly. Also means to understand something -- to get hold of an idea in your mind.",
    "Tom grasped the bars of the stocks, his knuckles white, refusing to let go of his father.",
    NOTES_VOCAB_GRASPED,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Vocabulary: lingered
  // =========================================================================
  vocabSlide(
    pres,
    "lingered",
    "verb",
    "To stay somewhere longer than necessary, to be slow to leave. There is a feeling of not wanting to go, of holding on to a moment or a place.",
    "The smell of smoke lingered in the cold night air long after the fire had died.",
    NOTES_VOCAB_LINGERED,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- Reading Launch
  // =========================================================================
  contentSlide(
    pres,
    "Read Aloud",
    C.PRIMARY,
    "Chapters 1-3",
    [
      "Reading Mode: Teacher read (Ch 1), Student read aloud (Ch 2-3)",
      "Ch 1: A 90-year-old Thomas sees a ghost on his lawn",
      "Ch 2: Tom tries to care for Pa in the stocks. Pa passes away",
      "Ch 3: Tom is orphaned and sent to the workhouse",
      "Focus: what choices does the author make to shape how we feel?",
    ],
    NOTES_READING,
    FOOTER
  );

  // =========================================================================
  // SLIDE 6 -- Pause Point 1 (p.2)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 1", "Chapter 1 -- p. 2",
    "If there was a ghost down there, he'd find it.",
    "p. 2",
    "What is going on here? What do the similes on this page tell us about the setting?",
    NOTES_PAUSE1, FOOTER
  );

  // =========================================================================
  // SLIDE 7 -- Pause Point 2 (p.6)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 2", "Chapter 2 -- p. 6",
    "The sightless eye stared at nothing.",
    "p. 6",
    "How is the author making you feel right now towards Tom and Pa?",
    NOTES_PAUSE2, FOOTER
  );

  // =========================================================================
  // SLIDE 8 -- Pause Point 3 (p.10)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 3", "Chapter 2 -- p. 10",
    "A cloak made out of courage would be thick and warm and keep him safe...",
    "p. 10",
    "The author uses a metaphor to describe courage. Why might she have chosen a cloak?",
    NOTES_PAUSE3, FOOTER
  );

  // =========================================================================
  // SLIDE 9 -- Pause Point 4 (p.13)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 4", "Chapter 3 -- p. 13",
    "Now be quiet.",
    "p. 13",
    "How are things looking for Tom now? What does this command reveal about how adults treat him?",
    NOTES_PAUSE4, FOOTER
  );

  // =========================================================================
  // SLIDE 10 -- I Do: Summary Sentences
  // =========================================================================
  modellingSlide(
    pres,
    "I Do",
    "Summary Sentences",
    "Question prompts:\n\nWho?\nWhat doing?\nWhen?\nWhere?\nWhy?\nHow?\n\nUse the prompts that matter\nfor YOUR summary.\nYou do not need all of them.",
    "Model: Chapter 1\n\nWho? Thomas, 90 years old\nWhat? Sees something on his lawn\nWhen? At night\nWhere? His farm in Australia\n\nSummary sentence:\n\"One night, Thomas, a 90-year-old farmer, spots a mysterious figure on his lawn and decides to investigate despite his age.\"",
    NOTES_IDO,
    FOOTER
  );

  // =========================================================================
  // SLIDES 11-12 -- We Do: Summary Sentences (withReveal)
  // =========================================================================
  withReveal(
    () => contentSlide(
      pres,
      "We Do",
      C.SUCCESS,
      "Write a Summary Sentence",
      [
        "Summarise: Why was Tom's father in the stocks?",
        "Step 1: Brainstorm with your partner -- who, what, why",
        "Step 2: Write a summary sentence on your whiteboard",
        "Step 3: Hold up on my signal",
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
        x: 0.7, y: ansY + 0.08, w: 1.9, h: 0.28, rectRadius: 0.08,
        fill: { color: C.WHITE },
      });
      slide.addText("Model Answer", {
        x: 0.7, y: ansY + 0.08, w: 1.9, h: 0.28,
        fontSize: 10, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      slide.addText("\"Pa was placed in the stocks for a week after printing a pamphlet that criticised the King, which was considered blasphemy.\"", {
        x: 2.7, y: ansY + 0.06, w: 6.6, h: 0.90,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_WEDO);
    }
  );

  // =========================================================================
  // SLIDE 13 -- You Do: Summary Sentences
  // =========================================================================
  contentSlide(
    pres,
    "You Do",
    C.ACCENT,
    "Summary Sentences -- Independent Practice",
    [
      "FIRST: Brainstorm who, what, where, when, why, how for each prompt",
      "NEXT: Write a summary sentence for each prompt on the worksheet",
      "THEN: Check -- does each sentence capture the main idea in one clear sentence?",
    ],
    NOTES_YOUDO,
    FOOTER,
    (slide, layoutGuide) => {
      const promptY = layoutGuide.panelTopPadded + 0.1;
      addCard(slide, layoutGuide.rightX, promptY, layoutGuide.rightW, 2.6, {
        strip: C.PRIMARY, fill: C.BG_CARD,
      });
      slide.addText("Your Prompts", {
        x: layoutGuide.rightX + 0.15, y: promptY + 0.08, w: 3.8, h: 0.28,
        fontSize: 11, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
      });
      slide.addText([
        { text: "1. What happens to Tom after his father dies?", options: { bullet: true, breakLine: true, fontSize: 12.5, color: C.CHARCOAL } },
        { text: "2. What important decision does Tom make the night his father dies?", options: { bullet: true, breakLine: true, fontSize: 12.5, color: C.CHARCOAL } },
        { text: "Optional: Summarise all three chapters in one sentence", options: { bullet: true, fontSize: 11.5, color: C.MUTED, italic: true } },
      ], {
        x: layoutGuide.rightX + 0.15, y: promptY + 0.42, w: 3.8, h: 2.0,
        fontFace: FONT_B, valign: "top", margin: 0, paraSpaceAfter: 6,
      });
    }
  );

  // =========================================================================
  // SLIDE 14 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    "What was the most powerful moment across Chapters 1-3? Can you summarise it in one sentence?",
    [
      "I can identify the authorial choices the author uses to shape how we feel",
      "I can explain the effect of at least one authorial choice on the reader",
      "I can write a summary sentence using who, what, where, when, why, how question prompts",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 15 -- Resources
  // =========================================================================


  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: Summary Sentences Worksheet -----------------------------------
  const ws = createPdf({ title: WORKSHEET_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Summary Sentences Worksheet", {
    color: C.NAVY,
    subtitle: "Sweep -- Chapters 1-3",
    lessonInfo: "Sweep | Lesson 1 | Week 1 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "A summary sentence captures the MAIN IDEA in one clear sentence.\nUse question prompts to help you brainstorm: Who? What doing? When? Where? Why? How?\nYou do not need all of them -- choose the prompts that matter for your summary.\nOptional: Try starting your summary sentence with the 'when' detail.", wsY, { color: C.NAVY });

  // Prompt 1
  wsY = addSectionHeading(ws, "Prompt 1: What happens to Tom after his father dies?", wsY, { color: C.NAVY });
  wsY = addBodyText(ws, "Brainstorm:", wsY, { fontSize: 10, bold: true });
  wsY += 2;
  wsY = addWriteLine(ws, "Who?", wsY, {});
  wsY = addWriteLine(ws, "What?", wsY, {});
  wsY = addWriteLine(ws, "Where?", wsY, {});
  wsY = addWriteLine(ws, "Why?", wsY, {});
  wsY += 4;
  wsY = addBodyText(ws, "Summary sentence:", wsY, { fontSize: 10, bold: true });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 28 });
  wsY += 8;

  // Prompt 2
  wsY = addSectionHeading(ws, "Prompt 2: What important decision does Tom make the night his father dies?", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Brainstorm:", wsY, { fontSize: 10, bold: true });
  wsY += 2;
  wsY = addWriteLine(ws, "Who?", wsY, {});
  wsY = addWriteLine(ws, "What?", wsY, {});
  wsY = addWriteLine(ws, "When?", wsY, {});
  wsY = addWriteLine(ws, "Why?", wsY, {});
  wsY += 4;
  wsY = addBodyText(ws, "Summary sentence:", wsY, { fontSize: 10, bold: true });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 28 });
  wsY += 8;

  // Extension prompt
  wsY = addSectionHeading(ws, "Extension: Summarise all three chapters in ONE sentence", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "This is harder -- you need to find the ONE main idea that connects Chapters 1, 2, and 3.", wsY, { fontSize: 10, italic: true });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 28 });

  addPdfFooter(ws, "Sweep | Lesson 1 | Summary Sentences Worksheet");

  // --- PDF 2: Answer Key ---------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Summary Sentences -- Answer Key", {
    color: C.ALERT,
    subtitle: "Teacher Reference -- Chapters 1-3",
    lessonInfo: "Sweep | Lesson 1 | Week 1 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Accept any answer that: (1) captures the main idea, (2) is a single sentence, and (3) uses specific detail from the text. Model answers below show one possible response.", akY, { color: C.ALERT });

  akY = addSectionHeading(ak, "Prompt 1: What happens to Tom after his father dies?", akY, { color: C.NAVY });
  akY = addBodyText(ak, "Model: After Pa dies in the stocks, Tom is taken by Mr Tupper to the magistrate, who declares that Tom must be sent to the workhouse because he is now an orphan with no one to care for him.", akY);
  akY = addBodyText(ak, "Key details to look for: Pa's death, magistrate's decision, workhouse, Tom being orphaned.", akY, { italic: true, fontSize: 10 });
  akY += 8;

  akY = addSectionHeading(ak, "Prompt 2: What important decision does Tom make?", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "Model: The night his father dies, Tom decides to wrap courage around himself like a cloak and face whatever comes next, even though he is alone and frightened.", akY);
  akY = addBodyText(ak, "Key details to look for: the courage cloak metaphor, Tom's resolve, the night of Pa's death.", akY, { italic: true, fontSize: 10 });
  akY += 8;

  akY = addSectionHeading(ak, "Extension: All three chapters in one sentence", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "Model: Sweep opens with an elderly Thomas seeing a ghost on his Australian farm, then takes us back to young Tom's life in England where he loses his father and is sent to the workhouse, revealing the beginning of a journey shaped by loss and courage.", akY);
  akY = addBodyText(ak, "This is a challenging task. Accept any sentence that connects the two timelines or identifies the overarching theme of the opening chapters.", akY, { italic: true, fontSize: 10 });

  addPdfFooter(ak, "Sweep | Lesson 1 | Answer Key -- TEACHER COPY");

  // --- Write all files ------------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Sweep_Lesson1.pptx` }),
    writePdf(ws, WORKSHEET_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log(`PPTX written to ${OUT_DIR}/Sweep_Lesson1.pptx`);
  console.log(`Done: ${WORKSHEET_RESOURCE.name}`);
  console.log(`Done: ${ANSWER_KEY_RESOURCE.name}`);
}

build().catch(console.error);
