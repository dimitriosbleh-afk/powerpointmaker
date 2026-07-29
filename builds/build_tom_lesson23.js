"use strict";

// Tom Unit -- Lesson 23: Chapters 48-51 -- Emma Arrives, Marines Leave + KPAS Note-Taking (continued)
// Week 6, Lesson 23, Grade 5/6 Literacy
// Reading: Ch 48 (Emma arrives), Ch 49 (Emma plans to find a husband), Ch 50 (Emma marries),
//          Ch 51 (marines sent home; Sergeant resigns to stay -- says Tom is family)
// Sentence-level: KPAS note-taking continued from Lesson 22
// Sensitivity: Ch 48 contains brief adult content -- handled in framing

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const { createTheme, weekToVariant } = require("../themes/factory");
// Unit cohesion: all Tom lessons use the same variant.
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

const SESSION_NUMBER = 23;
const FOOTER = "Chapters 48-51 | Lesson 23 | Week 6 | Year 5/6 Literacy";
const OUT_DIR = "output/Tom_Lesson23_Emma_Marines_Leave_KPAS";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const KPAS_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "KPAS Note-Taking Practice",
  "Student worksheet: convert sentences from Chapters 48-51 into key words, phrases, abbreviations and symbols."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "KPAS Note-Taking Practice Answer Key",
  "Teacher reference: model KPAS conversions for the practice worksheet."
);
const RESOURCE_ITEMS = [KPAS_RESOURCE, ANSWER_KEY_RESOURCE];
const KPAS_PDF_PATH = path.join(OUT_DIR, KPAS_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Lesson 23. Four chapters today: 48, 49, 50 and 51
- A new person arrives at Sergeant's Cove: a woman called Emma
- Then big news: the marines are being sent home -- and what the Sergeant decides will shape the rest of the book
- After reading we practise KPAS note-taking with sentences from these chapters

DO:
- Display title slide as students settle
- Have copies of the novel ready
- Have the KPAS practice worksheet ready (do not distribute yet)

TEACHER NOTES:
This is a content-rich lesson covering four chapters. Chapter 48 contains brief adult content (a convict woman is assigned to the Sergeant). Frame this before reading: the Sergeant has requested help around the house and Emma has been sent. The novel treats this matter-of-factly and so do we. The Sergeant's decision in Chapter 51 to resign and stay is the emotional turning point of the unit.

WATCH FOR:
- Students who find Chapter 48 confusing or uncomfortable -- pre-frame, then read briskly without lingering
- Students reacting to Tom running away in Ch 51 -- pause and discuss

[Literacy: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_RESOURCES = `SAY:
- Two resources today
- The ${KPAS_RESOURCE.name} is your practice sheet
- The ${ANSWER_KEY_RESOURCE.name} is for teacher use

DO:
- Print the practice sheet (one per student)
- Print one answer key for teacher use
- Distribute the practice after the I Do / We Do

TEACHER NOTES:
KPAS is now a familiar tool. Today's sentences come from Chapters 48-51 and connect directly to what students have just read. The skill will be used again next lesson when students take notes from non-fiction reading for the conclusion of their information report.

[Literacy: Resources | VTLM 2.0: Student Resources]`;

const NOTES_HOOK = `SAY:
- Before we read, a quick connection
- Last lesson we left the colony with a second fleet that brought women but no extra food
- One of those women is about to arrive at Sergeant's Cove. Her name is Emma
- Quick prediction: how do you think Tom and Rob will feel about a new person joining the household? Turn and tell

DO:
- 45 seconds turn and tell
- Cold call 2-3 students
- Brief frame: in Chapter 48 Emma arrives. The household will change. So will Tom

TEACHER NOTES:
This launch activates the prior knowledge from Lesson 22 (second fleet, women's ship) and sets up Chapter 48. Keep it tight -- 90 seconds total. The prediction sharpens attention for the read aloud.

WATCH FOR:
- Students who predict excitement or curiosity -- valid
- Students who predict resistance or worry -- also valid; celebrate the textual basis (the boys have built a family)

[Literacy: Hook | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands today: analysing the author's choices in Chapters 48-51, and continuing KPAS note-taking
- Read the success criteria

DO:
- Choral read the LI, then the SCs
- Brief: "Last lesson we LEARNED KPAS. Today we PRACTISE it with new sentences"

TEACHER NOTES:
SC1 targets analysis of character and authorial choice in Chapters 48-51. SC2 targets the underlying KPAS skill from Lesson 22. SC3 is application to fresh sentences from today's chapters.

WATCH FOR:
- Students who think today is just more reading -- redirect: "We also build our note-taking, ready for our writing later in the week"

[Literacy: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB1 = `SAY:
- Two explicit vocabulary words from this section
- First: frank. Frank means honest and direct -- saying what you really think, even when it is uncomfortable
- Emma is frank with Tom in Chapter 49 when she asks him about the Sergeant

DO:
- Display the word and meaning
- Read the meaning and example aloud
- Quick oral routine: "Use frank in a sentence with a partner. 20 seconds"

TEACHER NOTES:
"Frank" is one of the explicit vocab words for this lesson. Emma's frankness with Tom is a key character beat -- she is not playing games, she is laying out her plan. Knowing this word helps students notice Emma's directness.

WATCH FOR:
- Students who confuse 'frank' with the name 'Frank' -- clarify: same word, used as an adjective here
- Students who can use it in a sentence -- celebrate

[Literacy: Vocabulary | VTLM 2.0: Build Knowledge / Vocabulary]`;

const NOTES_VOCAB2 = `SAY:
- Second vocabulary word: famished
- Famished means extremely hungry -- much stronger than just 'hungry'
- After a long day with no food you might say "I am famished"

DO:
- Display the word and meaning
- Read the meaning and example aloud
- Quick oral routine: "Tell your partner about a time you felt famished"

TEACHER NOTES:
"Famished" appears in this section of the novel. It connects to the colony's food shortage theme from earlier chapters. Strong students will notice the food motif building.

WATCH FOR:
- Students who say "starving" -- accept and connect: "Yes, famished is close to starving"
- Students who use the word naturally -- celebrate

[Literacy: Vocabulary | VTLM 2.0: Build Knowledge / Vocabulary]`;

const NOTES_READING = `SAY:
- Four chapters today. I will read aloud
- Five pause points
- The chapters cover: Emma arriving (Ch 48), Emma's plan to find a husband (Ch 49), Emma marrying (Ch 50), and big news about the marines (Ch 51)

DO:
- 30 seconds for students to find Chapter 48
- Pre-frame Ch 48: "The Sergeant has asked for a convict woman to help around the house. Emma is sent. The novel mentions this briefly -- we read on"
- Read at a steady pace
- Take pause point 1 at p.256, pause point 2 at p.258, pause point 3 at p.260, pause point 4 at p.263, pause point 5 at p.267
- Read all four chapters

TEACHER NOTES:
The reading takes 18-22 minutes. Chapter 51 is the emotional centre -- Tom runs away briefly then resolves to stay, and the Sergeant chooses family over returning to England. Plan to slow slightly at the end so the emotional weight lands.

WATCH FOR:
- Students who get distracted in Ch 48 -- gently redirect
- Students moved by the Sergeant's decision at the end of Ch 51 -- name it: "Yes. The author has just told us they are family now"

[Literacy: Read Aloud | VTLM 2.0: Build Knowledge / Reading]`;

const NOTES_PAUSE1 = `SAY:
- Stop. We just heard Emma tell Tom about herself, and the narrator confirms: "...and Emma was London bred too."
- Turn to your partner: what have we learned about Emma from this conversation? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: she is from London like Tom; she was a convict; she has had a hard life; she is willing to chat openly; she sizes things up quickly; she is practical

TEACHER NOTES:
This pause point invites character inference from dialogue. Strong students will name the connection back to Tom -- they are both from London. Some students will simply summarise what Emma said. Push for inference: "What does her WAY of talking tell us, not just WHAT she said?"

WATCH FOR:
- Students who only summarise the facts -- prompt: "What kind of person speaks like that?"
- Students who connect Emma to Tom (both London) -- excellent

[Literacy: Pause Point | VTLM 2.0: Comprehension / Character Analysis]`;

const NOTES_PAUSE2 = `SAY:
- Stop. The narrator tells us about the quiet of the house: "...except for the patter of the bush rats in the shingles."
- Turn to your partner: what do you think the author wants us to know? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: the silence is heavy; the Sergeant did not go to Emma; Tom and Rob can hear because the house is small; the author is showing us, not telling us; something awkward is unspoken; the household has changed but not as expected

TEACHER NOTES:
A 'showing not telling' moment. The author does not say "the Sergeant did not visit Emma's room". The author tells us what the boys HEAR -- nothing but rats. Strong students will name the inference.

WATCH FOR:
- Students who describe the literal sound -- prompt: "Why does the author mention what they HEAR rather than what happened?"
- Students who name the inference (he did not go to her) -- celebrate the literary thinking

[Literacy: Pause Point | VTLM 2.0: Comprehension / Author Craft]`;

const NOTES_PAUSE3 = `SAY:
- Stop. Emma has just said: "I'll let ye know what I find."
- Turn to your partner: what have we learned about Emma from this conversation? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: she has a plan; she is practical, not romantic; she will not wait around; she takes charge of her own life; she is honest about what she wants; she likes Tom enough to keep him updated

TEACHER NOTES:
This adds to the picture of Emma. She is not bitter that the Sergeant has not been with her. She moves on -- and she does it efficiently. Strong students will name her agency.

WATCH FOR:
- Students who say Emma is unkind -- prompt: "Is this unkind, or just direct? Look at her words again"
- Students who name her agency or practicality -- celebrate

[Literacy: Pause Point | VTLM 2.0: Comprehension / Character Analysis]`;

const NOTES_PAUSE4 = `SAY:
- Stop. The narrator tells us about the Sergeant: "He was happier than he had ever been in his life."
- Turn to your partner: how is the author making you feel right now about these characters? 60 seconds

DO:
- 60 seconds Turn & Talk
- Cold Call 2-3 students
- Listen for: relieved for the Sergeant; warm; the unusual family is working; happiness has been earned after a long hard journey; satisfied; hopeful

TEACHER NOTES:
This is a quietly important emotional beat. The Sergeant came to the colony stiffly. He has lost a son. Now he is happier than ever. Strong students will note that this happiness comes from the household with the boys, not from what the colony imposed on him.

WATCH FOR:
- Students who describe only the Sergeant's actions -- prompt: "What feeling does the author leave you with?"
- Students who connect this to earlier sadness in the Sergeant -- excellent

[Literacy: Pause Point | VTLM 2.0: Comprehension / Author Craft]`;

const NOTES_PAUSE5 = `SAY:
- Stop. We have just heard Tom reflect: "No matter how bad they are, I can survive them."
- Two questions for your partner. 90 seconds
- Given what the author has already told us about Tom, what do you think he will do?
- What is the significance of using the courage cloak in this situation?

DO:
- 90 seconds Turn & Talk
- Cold Call 2-3 students for the first question, then 2-3 for the second
- Listen for the first: Tom will probably come back; he is used to bad news; he has survived before; he will not run forever
- Listen for the second: the courage cloak is from his mother; he is drawing on his mother again; he is reaching for what his mother gave him; this is what gives him strength; the author is showing the cloak as a real source of courage

TEACHER NOTES:
This pause point is the heart of Chapter 51. Tom is grieving the loss of the marines (and the Sergeant going home). He runs but reaches for the courage cloak -- the gift his mother gave him at the start of the book. Connect back: "Where do we first meet the courage cloak? What did Tom's mother say?" The cloak is a motif the author uses across the whole novel.

WATCH FOR:
- Students who only answer one question -- nudge to the other
- Students who connect the cloak back to Tom's mother and the start of the book -- excellent literary thinking; celebrate openly

[Literacy: Pause Point | VTLM 2.0: Comprehension / Motif and Character]`;

const NOTES_REVISE = `SAY:
- Quick revision of KPAS -- Key words, Phrases, Abbreviations, Symbols
- This is from last lesson. Quick recall
- Abbreviations: b/c (because), w/ (with), w/o (without)
- Symbols: = (means), + (and), -> (cause / result), up arrow (more), down arrow (less), / (new idea)
- Drop small words like "the", "and", "a" when taking notes

DO:
- Display the KPAS card
- Choral read each abbreviation and symbol
- Quick oral check: "What does -> mean? What does the / mean?"

TEACHER NOTES:
Keep this brisk -- it is recall, not new teaching. If students struggle, do not slow the lesson down here. The I Do and We Do will refresh the procedure.

WATCH FOR:
- Students who recall confidently -- celebrate retention
- Students who hesitate -- the I Do will rebuild the procedure

[Literacy: Revise | VTLM 2.0: Retention and Recall]`;

const NOTES_IDO = `SAY:
- Watch me convert one sentence from Chapter 51 into KPAS notes
- Sentence: "The marines were being sent back to England because the new soldiers from the second fleet had taken over their job."
- Step 1: find the key words. WHO? marines. WHAT DOING? being sent back to England. WHY? new soldiers took over
- Step 2: drop the small words -- the, were, being, back, because, the, from, the, had
- Step 3: use abbreviations and symbols
- My notes: "marines sent back to England b/c new soldiers took over"
- I can shorten more: "marines -> England b/c new soldiers took over"

DO:
- Display the original sentence and the KPAS notes side by side
- Talk through each step
- Highlight which words were dropped and which were replaced
- Note the second version uses the arrow to mean 'are going to' -- show this as a flexible use

TEACHER NOTES:
Show the thinking, not just the answer. The arrow here is creative -- it can mean 'moving toward'. Acknowledge that there are several valid versions; the goal is shorter notes that still capture the meaning.

WATCH FOR:
- Students who want one perfect answer -- redirect: "Many versions work. The TEST is: shorter, still meaningful"
- Students with their own creative use of arrows -- celebrate

[Literacy: I Do | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_WEDO = `SAY:
- Together. Mini-whiteboards out
- Sentence from Chapter 50: "Emma found a husband within a fortnight and was married soon after."
- Take 60 seconds. Convert this sentence into KPAS notes
- Boards up when called

DO:
- 60 seconds silent work on whiteboards
- Boards up
- Pick 2-3 strong examples to read aloud
- Reveal sample notes: "Emma found husband w/in fortnight + married soon after"
- Quick discussion: "What did you drop? What did you keep?"

TEACHER NOTES:
This is guided practice. Most students should produce workable notes. Common alternatives: "Emma -> husband 2 wks / married", "Emma found husband fortnight + married", "Emma married w/in fortnight". All are acceptable if they capture the key idea concisely. Notice w/in for 'within' may be new -- celebrate if students invent it.

WATCH FOR:
- Students who copy the whole sentence -- redirect: "Drop the small words"
- Students who lose the time information (fortnight) -- redirect: "Time is important here. Keep it"

[Literacy: We Do | VTLM 2.0: Guided Practice]`;

const NOTES_CFU_BUILD = `SAY:
- Quick check. Two sets of notes for the same sentence. Which set captures the key idea BEST?
- Sentence: "The Sergeant said that he would resign from the marines so that the boys could stay together as a family in the colony."
- A: "Sergeant resign marines"
- B: "Sergeant resign marines -> boys + Sergeant stay together as family in colony"
- Hold up A or B. Three, two, one -- show

DO:
- Display both
- Show Me Boards
- Scan: most students should choose B
- Cold Call 1-2 students: "Why B?"

TEACHER NOTES:
A is too short -- it loses the WHY and the consequence. B uses the arrow to capture the cause (resigns SO THAT they can stay together) and keeps the important detail that they will stay as a family. Notes need to be shorter than the original AND still meaningful.

WATCH FOR:
- Students who pick A because it is shortest -- redirect: "Shorter is not always better. Does A tell you the WHOLE idea?"
- Students who pick B and explain the arrow -- they have the principle

[Literacy: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Best notes: B
- B is shorter than the original but keeps the key idea: Sergeant resigns, and the reason is so they can stay as family
- A loses the reason completely
- Aim for B-style notes in your own work

DO:
- Display the reveal banner
- Read B aloud
- Pivot if many picked A: "Compare A and B. Can you guess the full meaning from A alone? Now from B?"

CFU CHECKPOINT:
Technique: Show Me Boards (A or B)
Script:
- "Hold up A or B. Which set of notes captures the key idea best?"
- Scan for: most students choose B (>=80%)
PROCEED (>=80%): Most chose B. Release to the worksheet.
PIVOT (<80%): Most likely issue -- students think shortest is best. Reteach: "Notes are shorter AND still carry the meaning. Drop the small words, keep the WHY."

TEACHER NOTES:
After reveal, release students to the worksheet. The principle (shorter AND still meaningful) is the takeaway.

WATCH FOR:
- Students who self-correct toward B-style notes -- celebrate

[Literacy: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Your turn. The worksheet has six sentences from Chapters 48-51
- For each one: convert it into KPAS notes
- Use the abbreviation and symbol bank at the top
- Drop the small words. Keep the WHY where there is one
- 12 minutes. I will circulate

DO:
- Distribute the worksheet
- Circulate -- prioritise students who looked unsure during the CFU
- Quick conferences: "Read me your notes. Shorter than the original? Still meaningful?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Use the enabling version of the worksheet which provides each sentence with the key words already underlined -- students just convert the underlined words into KPAS form
- Extra Notes: This focuses cognitive load on the conversion step rather than on identifying key words

EXTENDING PROMPT:
- Task: After completing the worksheet, write your OWN sentence about Chapter 51 and convert it into KPAS notes. Then swap with a partner -- can your partner reconstruct your sentence from your notes?
- Extra Notes: This pushes students to test the principle (shorter AND still meaningful) against a real audience

TEACHER NOTES:
The 12-minute block is independent practice. The worksheet uses sentences directly from today's chapters so the link between reading and note-taking stays close. Circulate actively.

WATCH FOR:
- Students who keep small words -- prompt: "Can you drop 'the' and 'a'?"
- Students who lose the cause or consequence -- prompt: "Use -> or b/c to capture the WHY"
- Students whose notes are so short the meaning is gone -- prompt: "Read it back. Does it still mean the same thing?"

[Literacy: You Do | VTLM 2.0: Independent Application]`;

const NOTES_CLOSING = `SAY:
- Quick self-check. Show me on your fingers
- SC1: I described what we learned about Emma, Tom and the Sergeant in Chapters 48-51 -- 1 to 5
- SC2: I can use key words, phrases, abbreviations and symbols to take shorter notes -- 1 to 5
- SC3: I converted sentences from Chapters 48-51 into KPAS notes -- 1 to 5
- Then turn to your partner: tell them what the courage cloak means to Tom now, at this point in the story

DO:
- Run each SC
- Listen in on partner shares
- Wrap up: "Next session we read Chapter 52 -- the LAST chapter of the book. Then we use our note-taking for our writing"

TEACHER NOTES:
The reflection prompt connects back to the most important pause point. The cloak has carried meaning across the whole novel; collect what students say so you can build on it in Lesson 24. Collect worksheets to scan for common patterns.

WATCH FOR:
- Students who name the cloak as a link back to Tom's mother -- excellent recall
- Students unsure -- check in privately

[Literacy: Closing | VTLM 2.0: Review and Reflect]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "Tom Appleby Ch 48-51 + KPAS Note-Taking -- Lesson 23";

  // SLIDE 1 -- Title
  titleSlide(
    pres,
    "Chapters 48-51",
    "Emma Arrives, Marines Leave + KPAS Note-Taking",
    "Lesson 23  |  Week 6  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  // SLIDE 2 -- Resources (immediately after title)
  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    T,
    FOOTER,
    NOTES_RESOURCES
  );

  // SLIDE 3 -- Hook / Launch (prior knowledge -> new learning)
  contentSlide(
    pres,
    "Launch",
    C.PRIMARY,
    "Who is About to Arrive at the Cove?",
    [
      "Last lesson: the second fleet brought women but no extra food",
      "One of those women is about to arrive at Sergeant's Cove",
      "Her name is Emma",
      "Prediction (turn and tell): how will Tom and Rob feel about a new person joining the household?",
    ],
    NOTES_HOOK,
    FOOTER
  );

  // SLIDE 4 -- LI / SC
  liSlide(
    pres,
    [
      "We are learning to analyse the author's choices in Chapters 48 to 51 and to keep practising our KPAS note-taking with sentences from these chapters",
    ],
    [
      "I can describe what we learned about Emma, Tom and the Sergeant in Chapters 48 to 51",
      "I can use key words, phrases, abbreviations and symbols to take shorter notes",
      "I can convert sentences from Chapters 48 to 51 into KPAS notes",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // SLIDE 5 -- Vocabulary: frank
  vocabSlide(
    pres,
    "frank",
    "adjective",
    "Honest and direct -- saying what you really think, even when it might be uncomfortable.",
    "Emma was frank with Tom when she asked him about the Sergeant.",
    NOTES_VOCAB1,
    FOOTER
  );

  // SLIDE 6 -- Vocabulary: famished
  vocabSlide(
    pres,
    "famished",
    "adjective",
    "Extremely hungry -- much stronger than just 'hungry'.",
    "After two days with little to eat, the convicts were famished.",
    NOTES_VOCAB2,
    FOOTER
  );

  // SLIDE 7 -- Reading anchor
  contentSlide(
    pres,
    "Read",
    C.PRIMARY,
    "Chapters 48, 49, 50 & 51",
    [
      "Ch 48: Emma arrives at the cove. Tom chats with her. A cow and calf are brought home",
      "Ch 49: Emma asks Tom about the Sergeant. She plans to find a husband",
      "Ch 50: Emma finds a husband and marries within two weeks",
      "Ch 51: The marines are being sent home. Tom runs -- then returns. The Sergeant resigns to stay",
      "Five pause points along the way",
    ],
    NOTES_READING,
    FOOTER
  );

  // SLIDE 8 -- Pause point 1 (p.256)
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 48  |  Emma talks with Tom",
    "...and Emma was London bred too.",
    "p.256",
    "What have we learned about Emma from this conversation?",
    NOTES_PAUSE1,
    FOOTER
  );

  // SLIDE 9 -- Pause point 2 (p.258)
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 48  |  The quiet of the house",
    "...except for the patter of the bush rats in the shingles.",
    "p.258",
    "What do you think the author wants us to know?",
    NOTES_PAUSE2,
    FOOTER
  );

  // SLIDE 10 -- Pause point 3 (p.260)
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 49  |  Emma's plan",
    "I'll let ye know what I find.",
    "p.260",
    "What have we learned about Emma from this conversation?",
    NOTES_PAUSE3,
    FOOTER
  );

  // SLIDE 11 -- Pause point 4 (p.263)
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 50  |  The Sergeant",
    "He was happier than he had ever been in his life.",
    "p.263",
    "How is the author making you feel right now about these characters?",
    NOTES_PAUSE4,
    FOOTER
  );

  // SLIDE 12 -- Pause point 5 (p.267) -- two-question pause
  quoteSlide(
    pres,
    "Pause Point",
    "Chapter 51  |  Tom and the courage cloak",
    "No matter how bad they are, I can survive them.",
    "p.267",
    "What do you think Tom will do? What is the significance of using the courage cloak here?",
    NOTES_PAUSE5,
    FOOTER
  );

  // SLIDE 13 -- Revise: KPAS
  contentSlide(
    pres,
    "Revise",
    C.SECONDARY,
    "KPAS -- Key Words, Phrases, Abbreviations, Symbols",
    [
      "Note-taking = writing the IMPORTANT bits in shorter form",
      "Abbreviations:  b/c (because)  |  w/ (with)  |  w/o (without)",
      "Symbols:  =  (means)  |  +  (and)  |  ->  (cause / result)",
      "More symbols:  up arrow (more)  |  down arrow (less)  |  /  (new idea)",
      "Shorter AND still meaningful -- both at once",
    ],
    NOTES_REVISE,
    FOOTER
  );

  // SLIDE 14 -- I Do: model KPAS conversion
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Convert a Sentence into KPAS Notes",
    "Sentence from Ch 51:\n\n\"The marines were being\nsent back to England\nbecause the new soldiers\nfrom the second fleet had\ntaken over their job.\"\n\nStep 1: find the key words\n  -> marines\n  -> sent back to England\n  -> new soldiers took over\n\nStep 2: drop small words\nthe, were, being, the, from, the, had",
    "Step 3: use abbreviations\n  -> b/c  (because)\n  -> -> (moving toward / so)\n\nMy notes:\n\n\"marines sent back to England\nb/c new soldiers took over\"\n\nEven shorter:\n\n\"marines -> England b/c\nnew soldiers took over\"\n\nNotice: same meaning, fewer words",
    NOTES_IDO,
    FOOTER
  );

  // SLIDE 15 -- We Do
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do", { color: C.SECONDARY });
    addTitle(s, "Together -- Whiteboards", { color: C.SECONDARY });

    addCard(s, 0.5, CONTENT_TOP, 9, 1.45, { strip: C.SECONDARY, fill: C.WHITE });
    s.addText("Convert this sentence into KPAS notes (Ch 50):", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.32,
      fontSize: 14, fontFace: FONT_B, color: C.SECONDARY, bold: true, margin: 0,
    });
    s.addText("\"Emma found a husband within a fortnight and was married soon after.\"", {
      x: 0.75, y: CONTENT_TOP + 0.46, w: 8.4, h: 0.90,
      fontSize: 17, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
      fit: "shrink", shrinkText: true, valign: "middle",
    });

    const tipY = CONTENT_TOP + 1.65;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Your job:", {
      x: 0.75, y: tipY + 0.10, w: 4, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- Find the KEY words\n- Drop the small words (a, within, was, after)\n- Keep the time information (fortnight)\n- 60 seconds. Boards up when called", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: 0.95,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_WEDO);
  }

  // SLIDE 16 + 17 -- CFU: which notes capture the key idea best? (reveal pair)
  function buildCfuBase() {
    const slide = pres.addSlide();
    addTopBar(slide, C.ALERT);
    addBadge(slide, "CFU", { color: C.ALERT });
    addTitle(slide, "Which Notes Capture the Key Idea Best?", { color: C.ALERT });

    const stampW = 1.3;
    slide.addShape("roundRect", {
      x: 9.5 - stampW, y: 0.20, w: stampW, h: 0.32, rectRadius: 0.08,
      fill: { color: C.WHITE }, line: { color: C.ALERT, width: 1.5 },
    });
    slide.addText("CHECK", {
      x: 9.5 - stampW, y: 0.20, w: stampW, h: 0.32,
      fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    slide.addShape("roundRect", {
      x: 0.5, y: CONTENT_TOP, w: 4.5, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ALERT },
    });
    slide.addText("Show Me Boards: A or B", {
      x: 0.5, y: CONTENT_TOP, w: 4.5, h: 0.40,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    slide.addText("Sentence:  \"The Sergeant said that he would resign from the marines so that the boys could stay together as a family in the colony.\"", {
      x: 0.5, y: CONTENT_TOP + 0.55, w: 9, h: 0.65,
      fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, bold: true, margin: 0, valign: "middle",
      fit: "shrink", shrinkText: true,
    });

    const cardY = CONTENT_TOP + 1.30;
    const cardH = 0.95;
    addCard(slide, 0.5, cardY, 9, cardH, { strip: C.MUTED, fill: C.WHITE });
    slide.addText("A", {
      x: 0.7, y: cardY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
    });
    slide.addText("\"Sergeant resign marines\"", {
      x: 1.2, y: cardY + 0.18, w: 8.0, h: cardH - 0.30,
      fontSize: 16, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    const cardBY = cardY + cardH + 0.18;
    const cardBH = SAFE_BOTTOM - cardBY - 0.55;
    addCard(slide, 0.5, cardBY, 9, cardBH, { strip: C.SECONDARY, fill: C.BG_CARD });
    slide.addText("B", {
      x: 0.7, y: cardBY + 0.10, w: 0.5, h: 0.4,
      fontSize: 22, fontFace: FONT_H, color: C.SECONDARY, bold: true, margin: 0,
    });
    slide.addText("\"Sergeant resign marines -> boys + Sergeant stay together as family in colony\"", {
      x: 1.2, y: cardBY + 0.14, w: 8.0, h: cardBH - 0.24,
      fontSize: 15, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(slide, FOOTER);
    slide.addNotes(NOTES_CFU_BUILD);
    return slide;
  }

  withReveal(
    buildCfuBase,
    (slide) => {
      const revealY = 4.68;
      addCard(slide, 0.5, revealY, 9, 0.42, { fill: C.SUCCESS });
      slide.addText("Best notes: B  --  shorter AND keeps the WHY", {
        x: 0.5, y: revealY, w: 9, h: 0.42,
        fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // SLIDE 18 -- You Do: worksheet
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "You Do", { color: C.PRIMARY, w: 1.5 });
    addTitle(s, "Convert to KPAS Notes");

    addCard(s, 0.5, CONTENT_TOP, 9, 1.85, { strip: C.PRIMARY, fill: C.WHITE });
    s.addText("Six sentences from Chapters 48-51", {
      x: 0.75, y: CONTENT_TOP + 0.10, w: 8.4, h: 0.30,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("First:    Read each sentence aloud quietly\nNext:     Underline the KEY words (who, what doing, where, when, why)\nThen:     Drop the small words (the, and, a, of)\nFinally:  Use abbreviations and symbols where helpful", {
      x: 0.75, y: CONTENT_TOP + 0.46, w: 8.4, h: 1.30,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    const tipY = CONTENT_TOP + 2.00;
    addCard(s, 0.5, tipY, 9, SAFE_BOTTOM - tipY, { strip: C.ACCENT, fill: C.BG_CARD });
    s.addText("Time: 12 minutes", {
      x: 0.75, y: tipY + 0.10, w: 4, h: 0.30,
      fontSize: 13, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("- The abbreviation and symbol bank is at the top of your worksheet\n- Read your notes back aloud -- do they still capture the meaning?\n- Aim for SHORTER but still meaningful", {
      x: 0.75, y: tipY + 0.42, w: 8.4, h: 0.85,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_YOUDO);
  }

  // SLIDE 19 -- Closing
  closingSlide(
    pres,
    {
      reflectionPrompt: "Tell your partner what the courage cloak means to Tom now, at this point in the story.",
      scItems: [
        "I can describe what we learned about Emma, Tom and the Sergeant in Chapters 48 to 51",
        "I can use key words, phrases, abbreviations and symbols to take shorter notes",
        "I can convert sentences from Chapters 48 to 51 into KPAS notes",
      ],
      selfAssessment: {
        prompt: "Show on your fingers: 1 (need help) to 5 (got it)",
        options: ["1-2", "3", "4-5"],
      },
    },
    NOTES_CLOSING
  );

  // ---- PDF: KPAS Practice -------------------------------------------------
  const ws = createPdf({ title: KPAS_RESOURCE.name });
  let wsY = addPdfHeader(ws, "KPAS Note-Taking Practice", {
    color: C.PRIMARY,
    subtitle: "Convert each sentence from Chapters 48-51 into notes",
    lessonInfo: "Lesson 23 | Week 6 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "For each sentence: find the KEY words, drop the small words (the, and, a, of), and use abbreviations and symbols where helpful. Aim for SHORTER notes that still carry the meaning.", wsY, { color: C.PRIMARY });

  wsY = addSectionHeading(ws, "Abbreviation & Symbol Bank", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Abbreviations:   b/c = because   |   w/ = with   |   w/o = without", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "Symbols:   =  means / definition   |   +  and   |   ->  cause / result", wsY, { fontSize: 11 });
  wsY = addBodyText(ws, "More:   up arrow  more / increase   |   down arrow  less / decrease   |   /  full stop / new idea", wsY, { fontSize: 11 });
  wsY += 8;

  wsY = addSectionHeading(ws, "1.  Sentence from Chapter 48", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "\"Emma arrived at Sergeant's Cove with a small bundle of her belongings and a tired look on her face.\"", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "My KPAS notes:", wsY, { fontSize: 10 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;

  wsY = addSectionHeading(ws, "2.  Sentence from Chapter 48", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "\"The Sergeant brought home a cow and her calf so that the household would have fresh milk.\"", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "My KPAS notes:", wsY, { fontSize: 10 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;

  wsY = addSectionHeading(ws, "3.  Sentence from Chapter 49", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "\"Emma told Tom that she would look for a husband when she went to collect the rations.\"", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "My KPAS notes:", wsY, { fontSize: 10 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;

  wsY = addSectionHeading(ws, "4.  Sentence from Chapter 50", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "\"Emma found a husband within a fortnight and was married soon after.\"", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "My KPAS notes:", wsY, { fontSize: 10 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;

  wsY = addSectionHeading(ws, "5.  Sentence from Chapter 51", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "\"The marines were being sent back to England because the new soldiers from the second fleet had taken over their job.\"", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "My KPAS notes:", wsY, { fontSize: 10 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 4;

  wsY = addSectionHeading(ws, "6.  Sentence from Chapter 51", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "\"The Sergeant decided to resign from the marines so that the boys and the Sergeant could stay together as a family in the colony.\"", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "My KPAS notes:", wsY, { fontSize: 10 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY += 6;

  wsY = addSectionHeading(ws, "Challenge", wsY, { color: C.ALERT });
  wsY = addBodyText(ws, "Write your OWN sentence about Chapter 51 (one or two lines), then convert it into KPAS notes. Swap with a partner -- can your partner reconstruct your sentence from your notes?", wsY, { fontSize: 11, italic: true });
  wsY = addBodyText(ws, "My sentence:", wsY, { fontSize: 10 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });
  wsY = addBodyText(ws, "My KPAS notes:", wsY, { fontSize: 10 });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 22 });

  addPdfFooter(ws, "Lesson 23 | KPAS Note-Taking Practice");

  // ---- PDF: Answer Key ---------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "KPAS Note-Taking Practice -- Answer Key", {
    color: C.SECONDARY,
    subtitle: "Teacher reference -- many valid versions for each",
    lessonInfo: "Lesson 23 | Week 6 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Many versions of KPAS notes are valid. The principle is: shorter than the original AND still meaningful. Accept any version that captures the key idea using key words and (where helpful) abbreviations and symbols.", akY, { color: C.SECONDARY });

  akY = addSectionHeading(ak, "1. Emma arrives (Ch 48)", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sample notes:  \"Emma arrived cove w/ small bundle + tired face\"", akY, { fontSize: 11 });
  akY = addBodyText(ak, "Alt:  \"Emma -> cove / bundle + tired\"  or  \"Emma arrived cove tired\"", akY, { fontSize: 10, italic: true });
  akY += 4;

  akY = addSectionHeading(ak, "2. Cow and calf (Ch 48)", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sample notes:  \"Sergeant brought home cow + calf -> fresh milk\"", akY, { fontSize: 11 });
  akY = addBodyText(ak, "Alt:  \"cow + calf home / for milk\"  or  \"Sergeant got cow + calf -> milk\"", akY, { fontSize: 10, italic: true });
  akY += 4;

  akY = addSectionHeading(ak, "3. Emma's plan (Ch 49)", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sample notes:  \"Emma told Tom -> will find husband at rations\"", akY, { fontSize: 11 });
  akY = addBodyText(ak, "Alt:  \"Emma -> looking for husband when collecting rations\"  or  \"Emma: husband hunt at rations\"", akY, { fontSize: 10, italic: true });
  akY += 4;

  akY = addSectionHeading(ak, "4. Emma marries (Ch 50)", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sample notes:  \"Emma found husband w/in fortnight + married soon after\"", akY, { fontSize: 11 });
  akY = addBodyText(ak, "Alt:  \"Emma -> husband fortnight / married\"  or  \"Emma found husband in 2 wks -> married\"", akY, { fontSize: 10, italic: true });
  akY += 4;

  akY = addSectionHeading(ak, "5. Marines sent back (Ch 51)", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sample notes:  \"marines sent back to England b/c new soldiers took over\"", akY, { fontSize: 11 });
  akY = addBodyText(ak, "Alt:  \"marines -> England b/c new soldiers (2nd fleet) took job\"  or  \"new soldiers -> marines leave\"", akY, { fontSize: 10, italic: true });
  akY += 4;

  akY = addSectionHeading(ak, "6. Sergeant resigns (Ch 51)", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Sample notes:  \"Sergeant resign marines -> boys + Sergeant stay as family in colony\"", akY, { fontSize: 11 });
  akY = addBodyText(ak, "Alt:  \"Sergeant -> resign / family stays together\"  or  \"Sergeant resigns b/c boys = family\"", akY, { fontSize: 10, italic: true });
  akY += 6;

  akY = addSectionHeading(ak, "Common Errors to Watch For", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- Notes that are still full sentences. Reteach: drop the small words.", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Notes that lose the WHY (no b/c or arrow). Reteach: capture cause as well as event.", akY, { fontSize: 10 });
  akY = addBodyText(ak, "- Notes that are so short the meaning is lost. Reteach: 'read your notes back -- can you reconstruct the sentence?'", akY, { fontSize: 10 });

  addPdfFooter(ak, "Lesson 23 | KPAS Note-Taking Practice Answer Key -- TEACHER USE");

  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/Tom_Lesson23.pptx` }),
    writePdf(ws, KPAS_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log("PPTX written to " + `${OUT_DIR}/Tom_Lesson23.pptx`);
  console.log("Done: " + KPAS_RESOURCE.name);
  console.log("Done: " + ANSWER_KEY_RESOURCE.name);
}

build().catch(console.error);
