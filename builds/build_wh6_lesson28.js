"use strict";

// War Horse Unit -- Lessons 28-29: Chapter 21 + Topic Sentences + 2nd Body Paragraph
// Week 6, Session 6, Grade 5/6 Literacy
// Final chapter reading + topic sentence practice + Australian WW1 reading + 2nd body paragraph

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  addTextOnShape, withReveal,
  titleSlide, liSlide, contentSlide,
  cfuSlide, taskSlide, closingSlide,
  vocabSlide, quoteSlide, modellingSlide,
} = require("../themes/wh6_helpers");

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 6;
const FOOTER = "War Horse | Lessons 28-29 | Week 6 | Year 5/6 Literacy";
const OUT_DIR = "output/WH6_Lesson28_Chapter21_2nd_Para";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const SPO_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "SPO Template",
  "Single paragraph outline template for the second body paragraph."
);
const VOCAB_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Tier 2-3 Vocabulary Reference",
  "Reference sheet with tier 2 and tier 3 vocabulary for persuasive writing about WW1."
);
const RESOURCE_ITEMS = [SPO_RESOURCE, VOCAB_RESOURCE];
const SPO_PDF_PATH = path.join(OUT_DIR, SPO_RESOURCE.fileName);
const VOCAB_PDF_PATH = path.join(OUT_DIR, VOCAB_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Today we finish War Horse with Chapter 21 -- the final chapter
- We also read about Australia's involvement in World War 1 and write our second body paragraph
- Big session today -- reading and writing working together

DO:
- Display title slide as students settle
- Have copies of War Horse bookmarked at Chapter 21 (p. 175)
- Have the Britannica Kids article ready on devices or printed

TEACHER NOTES:
Session 6 combines Lessons 28 and 29. Chapter 21 is the resolution -- Joey returns to Albert. The Australian WW1 reading provides new evidence for the persuasive text. The writing focus is on topic sentences (from L28) and the second body paragraph with tier 2/3 vocabulary (from L29).

WATCH FOR:
- Students who are eager to find out how the story ends -- channel into careful reading
- Students who need a recap of the auction from last session

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Three things today: finish War Horse, practise writing strong topic sentences, and write our second body paragraph
- Read the success criteria. SC1 is about topic sentences. SC2 is about planning with the SPO. SC3 is about using precise, formal vocabulary

DO:
- Choral read the LI, then the SCs
- Connect to prior learning: "We used the SPO in our last writing session. Today we use it again for a new argument"

TEACHER NOTES:
SC1 (topic sentences) connects to the L28 sentence-level writing focus. SC2 (SPO planning) is a repeat from Session 4 -- students should be more confident now. SC3 (tier 2/3 vocabulary) is the new persuasive language tool from L29.

WATCH FOR:
- Students who are confident with the SPO from Session 4 -- they should move through planning quickly
- Students who are unsure about "tier 2 and tier 3 vocabulary" -- explain briefly: "precise, subject-specific words that make your writing sound more knowledgeable"

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_STERN = `SAY:
- Our vocabulary word: stern. Read it with me: stern [students repeat]
- Stern means very serious, strict, and showing disapproval. A stern person does not smile easily -- their expression tells you this is not a moment for messing around
- In Chapter 21, Emilie's grandfather speaks sternly to Albert at first. He is serious and protective -- this horse belonged to his granddaughter
- Ask: What is the difference between stern and angry? [Stern is controlled and serious. Angry is heated and emotional. A stern teacher might say "That is not acceptable" calmly. An angry teacher might raise their voice]

DO:
- Display word, choral read, define, give example
- Connect to the chapter: the grandfather's stern demeanour softens as he understands Albert's story
- Cold Call 2 students: describe a time someone was stern with you

TEACHER NOTES:
"Stern" appears on p.179 ("sternly"). It captures the grandfather's protective authority. The word also connects to persuasive writing -- a persuasive text sometimes needs a stern, authoritative tone.

WATCH FOR:
- Students who think stern means mean -- "Stern is not cruel. It is serious with purpose. A stern person cares deeply about something"
- Students who connect stern to authority figures in their lives -- good real-world application

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_READING = `SAY:
- Chapter 21 -- the final chapter of War Horse. This is where everything is resolved
- After the auction, Albert must face the reality that Joey was sold to Emilie's grandfather
- But there is a conversation ahead that changes everything
- Reading mode: student read aloud. Pages 175 to 182. Two pause points
- Find page 175 now

DO:
- Give students 30 seconds to find p. 175
- Select first reader
- This is the final chapter -- let the emotional moments breathe

TEACHER NOTES:
Chapter 21 resolves the central narrative. The grandfather's story connects back to Emilie, creating a satisfying full-circle moment. Albert's promise to keep Emilie's memory alive is the emotional climax. The chapter ends with Joey returning to England -- the promise from Chapter 1 fulfilled.

WATCH FOR:
- Students who are emotional about Emilie's death -- handle sensitively. The revelation comes mid-chapter
- Students who remember Emilie from earlier chapters -- excellent recall

[General: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause here. "It's true, sir, goes to their 'eads, sir. Must do, mustn't it?" Page 177
- What is going on? What does the author want us to know? [The men are talking about the grandfather -- how he goes on about his granddaughter and the horse. Sergeant Thunder is explaining to Albert that the grandfather has a deep emotional connection to Joey through Emilie]
- Ask: Why is this moment important for Albert? [He is about to learn that someone else loved Joey just as much as he did. This is confronting -- he thought Joey was HIS horse, but Emilie's grandfather feels the same way]
- Notice the simile on p.176: "Sounds mad as a hatter." This is informal soldier speech. Why does Morpurgo use it? [It makes the soldiers feel real. They speak casually, even about emotional things. It contrasts with the gravity of the situation]

DO:
- Display the quote
- Use Cold Call for responses
- Push for perspective: Albert is not the only person who loves Joey

CFU CHECKPOINT:
Technique: Cold Call

Script:
- "What does Albert learn from this conversation that challenges his view of Joey?"
- Cold call 3 students
- Scan for: Albert realises Joey's story is bigger than just his own -- Joey mattered to other people too

PROCEED (>=80%): Most students grasp the multiple perspectives. Continue reading.
PIVOT (<80%): Most likely issue -- students only see Albert's perspective. Reteach: "Put yourself in the grandfather's shoes. His granddaughter Emilie was sick. Joey was her comfort. Then Emilie died. The grandfather promised to find Joey. Would you let that horse go easily?" Re-check: "Whose claim to Joey is stronger -- Albert's or the grandfather's? Why?"

TEACHER NOTES:
This pause point targets SC1 indirectly -- the grandfather's perspective is revealed through the soldiers' dialogue. The competing claims to Joey mirror real emotional complexity. There is no clear "right" answer about who Joey belongs to -- that is the point.

WATCH FOR:
- Students who insist Albert's claim is the only valid one -- push: "The grandfather's love for Joey came through his granddaughter. Is that love less real?"
- Students who are surprised by Emilie's death -- this is revealed in the grandfather's conversation and may be confronting

SENSITIVITY ADVISORY:
- What it is: Emilie's death is revealed during the grandfather's conversation. This is the death of a child character students have grown attached to
- Framing language: "This part of the story includes something sad. Emilie became ill and passed away. If this is difficult for you, that is okay"
- Watch for: Students who become quiet or withdrawn -- this may trigger personal connections
- Protocol: Acknowledge feelings privately, do not require discussion participation, follow school wellbeing processes if needed

[General: Pause Point 1 | VTLM 2.0: Teacher-Led Discussion]`;

const NOTES_PAUSE2 = `SAY:
- "I will sell my Emilie's horse to you." Page 180
- How do things look now? [The grandfather agrees to sell Joey to Albert. But there is a condition -- Albert must promise to care for Joey and keep Emilie's memory alive]
- Ask: Why does the grandfather change his mind? What convinces him? [He sees that Albert's love for Joey is genuine. When he hears Albert's full story -- how he raised Joey from a foal, how he searched for him through the war -- he recognises a kindred spirit. Both men love this horse deeply]
- The resolution is about understanding and empathy. Neither man is wrong. They find a way forward together

DO:
- Display the quote
- Let this moment land -- it is the emotional resolution of the entire novel
- Cold Call: "What does Albert promise? Why does it matter?"

TEACHER NOTES:
This is the narrative climax of the novel. The grandfather's decision is driven by recognising shared love -- not by money or rights. Albert's promise to remember Emilie connects the novel's themes of memory, loyalty, and the bonds between humans and animals. This connects directly to the persuasive writing topic about the significance of animals in war.

WATCH FOR:
- Students who notice the grandfather's generosity -- he could have kept Joey or sold for more money
- Students who connect this to the broader theme: Joey's journey touched many lives, not just Albert's

[General: Pause Point 2 | VTLM 2.0: Deep Comprehension]`;

const NOTES_READING2 = `SAY:
- Now we switch to our second reading. This is research for your persuasive writing
- We are reading from the Britannica Kids article: Australia in World War 1
- Sections: Introduction, Attitudes Toward War, Australians Who Served, Women in the War, Impact of the War, Changing Status of Women
- Watch the BTN video included in the Introduction
- Take notes as you read. You are looking for evidence for your second body paragraph

DO:
- Direct students to the Britannica Kids article on devices or distribute printed copies
- Play the BTN video from the Introduction section
- Select readers for the article sections
- Monitor note-taking

TEACHER NOTES:
This reading shifts from the novel to historical source material. Students are building their evidence base for the second body paragraph. The Australian focus broadens their perspective beyond the British/European experience. Incidental vocabulary: sacrifice, casualty, disruption, conscription, declining, enlistment, shortages, hardships, patriotism.

SOURCES:
Australia in World War 1 - Britannica Kids

WATCH FOR:
- Students who are still emotionally processing the War Horse ending -- give them a moment before transitioning
- Students who take better notes this time around -- they have practised the skill

[General: Reading | VTLM 2.0: Structured Reading Practice]`;

const NOTES_TS_IDO = `SAY:
- Before we plan our paragraph, let's sharpen our topic sentences
- A strong topic sentence does two things: states the argument clearly AND makes the reader want to keep reading
- Four strategies for interesting topic sentences: use different sentence types (question, exclamation, command), use vivid words, start with an adverb or prepositional phrase, or make a bold claim
- Watch: here are some supporting details about the grandfather. Read them from the slide
- From these details, I can generate a topic sentence: "Driven by a promise he could never break, Emilie's grandfather searched tirelessly across France for the horse that had meant everything to his granddaughter."
- That starts with a participial phrase, uses vivid words ("tirelessly," "meant everything"), and makes the reader curious

DO:
- Display the SD set and the model TS
- Think aloud: "I read the SDs first, then I ask: what is the MAIN IDEA that connects all of them? That becomes my topic sentence"
- Show the connection between each SD and how the TS captures the essence

TEACHER NOTES:
This I Do combines the topic sentence review from L28 with a brief modelling moment. The SD sets are from the lesson plan. The modelling is deliberately brisk since students practised topic sentences in Session 3 (Lesson 23).

WATCH FOR:
- Students who write a TS that only covers one SD instead of the connecting idea -- "Your TS should be the umbrella that covers ALL the supporting details"
- Students who remember the strategies from Session 3 -- build on that: "You already know sentence types and vivid words. Now add starting with an interesting phrase"

[General: I Do -- Topic Sentences | VTLM 2.0: Explicit Teaching]`;

const NOTES_TIER23 = `SAY:
- One more persuasive tool before you write. Tier 2 and tier 3 vocabulary
- Tier 2 words are precise, formal words that work across subjects: significant, demonstrate, essential, consequence, contribution, impact
- Tier 3 words are subject-specific: conscription, armistice, casualties, commemoration, allied forces
- Using these words makes your writing sound more knowledgeable and credible. Compare:
- Basic: "Lots of people died in the war and it was really sad"
- With tier 2/3: "The devastating casualties of the war had a significant impact on communities across Australia"
- The second version sounds like someone who KNOWS the topic

DO:
- Display the tier 2/3 examples
- Read both versions aloud -- the contrast should be obvious
- Distribute the Session 6 Tier 2-3 Vocabulary Reference sheet

TEACHER NOTES:
Tier 2/3 vocabulary is the final persuasive language tool in the unit (after modality in Session 4 and nominalisations). The reference sheet provides a word bank students can draw from. This is about USING the words, not memorising definitions.

WATCH FOR:
- Students who try to use every tier 2/3 word at once -- "Pick 2-3 that fit naturally. Forcing in too many makes your writing sound unnatural"
- Students who are unsure which words fit their argument -- direct them to the reference sheet

[General: I Do -- Vocabulary for Persuasion | VTLM 2.0: Explicit Teaching]`;

const NOTES_CFU = `SAY:
- Quick check. Which topic sentence is stronger for a body paragraph about sacrifice?
- A: "People gave up things during the war."
- B: "The immense sacrifices made by Australian families during World War 1 demonstrate the true cost of conflict."
- Hold up fingers: 1 for A, 2 for B

DO:
- Display both options
- Allow 10 seconds to read and decide
- Signal: fingers up

CFU CHECKPOINT:
Technique: Finger Voting

Script:
- "1 finger for A, 2 fingers for B. Which is the stronger topic sentence?"
- Scan for: 2 fingers (B is stronger -- tier 2/3 vocabulary, specific, authoritative)
- Follow up: "What makes B stronger? Name a specific word." [immense, sacrifices, demonstrate, conflict -- all tier 2/3]

PROCEED (>=80%): Most show 2 fingers and can identify the stronger vocabulary. Move to We Do.
PIVOT (<80%): Most likely issue -- students think A is "simpler and clearer." Reteach: "Both are clear. But which one sounds like it was written by someone who really knows the topic? 'Gave up things' is vague -- what things? 'Immense sacrifices' is precise. 'The war' is generic. 'World War 1' is specific. Persuasive writing needs precision." Re-check: "Rewrite 'Animals were useful' using one tier 2 or tier 3 word."

TEACHER NOTES:
This CFU checks whether students can distinguish between basic and tier 2/3 vocabulary in context. The follow-up question requires them to identify the specific words that make B stronger.

WATCH FOR:
- Students who choose B but cannot explain WHY -- they are sensing the difference without analysing it. Push: "Point to the exact word that makes it stronger"
- Students who think longer = better -- "It is not about length. It is about precision"

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_WEDO = `SAY:
- Now plan your second body paragraph. You have your notes from the Australian WW1 reading
- Pick a NEW argument -- different from your first body paragraph
- Ideas from the reading: the sacrifice of Australian families, the impact on communities, the role of women, the changing home front, the cost of war on young people
- With your partner: discuss your argument. What evidence from the reading supports it?
- Then fill in your SPO template. Make your topic sentence strong -- use what we practised today

DO:
- Distribute the Session 6 SPO Template
- Give 2 minutes for partner discussion
- Circulate as students plan -- check topic sentences for clarity, modality, and tier 2/3 vocabulary
- After 5 minutes: "You should have your TS and at least one SD"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide the SPO template with sentence starters for each section. TS starter: "One significant impact of World War 1 was..." SD starters: "For example, ..." / "This meant that ..." / "As a result, ..."
- Extra Notes: Students can draw evidence directly from their reading notes

EXTENDING PROMPT:
- Task: After completing the SPO, write the full paragraph. Then revise one sentence to include a nominalisation (from Session 4) AND a tier 2/3 word in the same sentence

TEACHER NOTES:
The We Do is guided planning for the second body paragraph. Students have done this once before (Session 4) so should be more confident with the SPO structure. The new element is incorporating tier 2/3 vocabulary.

WATCH FOR:
- Students who repeat their first body paragraph argument -- "You need a DIFFERENT argument. What else did you learn from the reading?"
- Students whose topic sentences are flat despite the I Do practice -- prompt: "Can you make that more interesting? Try starting with a strong claim or a vivid phrase"

[General: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_YOUDO = `SAY:
- Write your second body paragraph from your SPO
- First: check your SPO is complete
- Next: write your paragraph. Try to include at least one tier 2 or tier 3 vocabulary word from the reference sheet
- Then: reread and edit. Check: is your topic sentence strong? Does each sentence support your argument? Have you used precise vocabulary?

DO:
- Students write independently
- Circulate -- look for tier 2/3 vocabulary in use
- After 7 minutes: "Start rereading and editing"
- Collect or photograph finished paragraphs

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Write the paragraph following the SPO structure. Focus on getting the TS-SD-CS structure right. Do not worry about tier 2/3 vocabulary -- the structure is the priority
- Extra Notes: These students can refer to the teacher's modelled paragraph from Session 4 as a structural guide

EXTENDING PROMPT:
- Task: After writing and editing, write a brief reflection: "How is my second body paragraph different from my first? What persuasive techniques did I use?" Identify at least one instance each of modality, nominalisation, and tier 2/3 vocabulary across both paragraphs

TEACHER NOTES:
This is the second time students write from an SPO. The quality should be noticeably higher than Session 4. The editing step is integrated -- students reread and revise before submitting.

WATCH FOR:
- Students who rush the editing step -- "Reading your work aloud quietly helps you catch errors your eyes miss"
- Students whose second paragraph sounds very similar to their first -- push for a genuinely different argument and evidence

[General: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: writing strong topic sentences -- thumbs? [scan]
- SC2: planning with the SPO -- thumbs? [scan]
- SC3: using tier 2 and tier 3 vocabulary -- thumbs? [scan]
- Turn and talk: We finished War Horse today. How does Joey's story connect to what you are writing in your persuasive text? What did Joey's journey teach you about the role of animals in war?

DO:
- Run through each SC with thumbs check
- The turn-and-talk bridges the novel study and the persuasive writing -- Joey's story IS evidence for their persuasive text
- Preview: "Next session we read about Indigenous Australians in WW1 and write our concluding paragraph"

TEACHER NOTES:
The closing explicitly connects the novel study to the persuasive writing. Joey's journey is not just a story -- it is evidence. Students who can articulate this connection are demonstrating deep understanding of both the narrative and the persuasive purpose.

WATCH FOR:
- Students who are emotional about finishing War Horse -- acknowledge: "We have been on a journey with Joey. It is okay to feel that"
- Students who make the connection between Joey's experiences and their persuasive arguments -- excellent synthesis

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two resources for today
- The ${SPO_RESOURCE.name} is for planning your second body paragraph
- The ${VOCAB_RESOURCE.name} has tier 2 and tier 3 words you can use in your persuasive writing

DO:
- Print the SPO template before the lesson (one per student)
- Print the vocabulary reference (one per student or per pair)
- Click any resource card to open the PDF

TEACHER NOTES:
Students should keep both completed SPO templates (Session 4 and Session 6) with their persuasive text drafts. They will need all paragraphs for the final editing session.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "War Horse - Lessons 28-29 - Chapter 21 + 2nd Body Paragraph";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "War Horse",
    "Chapter 21 + Second Body Paragraph",
    "Lessons 28-29  |  Week 6  |  Year 5/6 Literacy",
    NOTES_TITLE
  );

  const theme = { C, FONT_H, FONT_B, addTopBar, addTitle, addFooter, addCard };
  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    theme,
    FOOTER,
    NOTES_RESOURCES
  );

  // =========================================================================
  // SLIDE 2 -- LI/SC
  // =========================================================================
  liSlide(
    pres,
    [
      "We are learning to write engaging topic sentences, and to plan and write a second body paragraph using precise vocabulary",
    ],
    [
      "I can write an interesting topic sentence that makes the reader want to keep reading",
      "I can plan a body paragraph using the single paragraph outline template",
      "I can use tier 2 and tier 3 vocabulary to add credibility to my persuasive writing",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Vocabulary: stern
  // =========================================================================
  vocabSlide(
    pres,
    "stern",
    "adjective",
    "Very serious, strict, and showing disapproval. A stern person does not smile easily -- their expression tells you this is not a moment for joking.",
    "The grandfather spoke sternly at first, his voice heavy with the weight of a promise he had made to his dying granddaughter.",
    NOTES_VOCAB_STERN,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Reading Launch: Chapter 21
  // =========================================================================
  contentSlide(
    pres,
    "Read Aloud",
    C.PRIMARY,
    "Chapter 21 -- Pages 175-182",
    [
      "Reading Mode: Student Read Aloud",
      "The final chapter of War Horse",
      "Albert and the grandfather have a conversation that changes everything",
      "Focus: how does understanding another person's perspective lead to resolution?",
    ],
    NOTES_READING,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- Pause Point 1 (Ch21 p.177)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 1", "Chapter 21 -- p. 177",
    "It's true, sir, goes to their 'eads, sir. Must do, mustn't it?",
    "p. 177",
    "What's going on? What does Albert learn that challenges his view of Joey?",
    NOTES_PAUSE1, FOOTER
  );

  // =========================================================================
  // SLIDE 6 -- Pause Point 2 (Ch21 p.180)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 2", "Chapter 21 -- p. 180",
    "I will sell my Emilie's horse to you.",
    "p. 180",
    "How do things look now? What convinces the grandfather to let Joey go?",
    NOTES_PAUSE2, FOOTER
  );

  // =========================================================================
  // SLIDE 7 -- Australian WW1 Reading
  // =========================================================================
  contentSlide(
    pres,
    "Read Aloud",
    C.SECONDARY,
    "Australia in World War 1",
    [
      "Britannica Kids: Australia in World War 1",
      "Sections: Introduction, Attitudes Toward War, Australians Who Served",
      "Women in the War, Impact of the War, Changing Status of Women",
      "Watch the BTN video in the Introduction",
      "Take notes: what evidence could you use for your second body paragraph?",
    ],
    NOTES_READING2,
    FOOTER
  );

  // =========================================================================
  // SLIDE 8 -- Topic Sentences + Tier 2/3 Vocab (I Do)
  // =========================================================================
  modellingSlide(
    pres,
    "I Do",
    "Strong Topic Sentences + Precise Vocabulary",
    "Strategies for interesting topic sentences:\n- Change the sentence type\n- Use vivid words\n- Start with an interesting phrase\n- Make a bold claim\n\nSupporting details:\n- The grandfather was the only surviving member of his family\n- He wanted to keep his promise to Emilie\n- He went to horse auctions all over France",
    "Model TS: \"Driven by a promise he could never break, Emilie's grandfather searched tirelessly across France for the horse that meant everything to his granddaughter.\"\n\nTier 2/3 vocabulary adds authority:\nBasic: \"Lots of people died\"\nPrecise: \"The devastating casualties had a significant impact on communities\"",
    NOTES_TS_IDO,
    FOOTER
  );

  // =========================================================================
  // SLIDES 9-10 -- CFU (withReveal)
  // =========================================================================
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Which Topic Sentence Is Stronger?",
      "Finger Voting",
      "A) \"People gave up things during the war.\"\n\nB) \"The immense sacrifices made by Australian families during World War 1 demonstrate the true cost of conflict.\"\n\nHold up: 1 finger for A, 2 fingers for B.",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      const ansY = 4.10;
      slide.addShape("roundRect", {
        x: 0.5, y: ansY, w: 9, h: 0.85, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      });

      slide.addShape("roundRect", {
        x: 0.7, y: ansY + 0.10, w: 1.4, h: 0.28, rectRadius: 0.08,
        fill: { color: C.WHITE },
      });
      slide.addText("Answer", {
        x: 0.7, y: ansY + 0.10, w: 1.4, h: 0.28,
        fontSize: 10, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      slide.addText("B is stronger. Key words: immense, sacrifices, demonstrate, conflict -- all tier 2/3 vocabulary that adds authority and precision.", {
        x: 2.3, y: ansY + 0.08, w: 7.0, h: 0.68,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU);
    }
  );

  // =========================================================================
  // SLIDE 11 -- We Do: Plan 2nd SPO
  // =========================================================================
  contentSlide(
    pres,
    "We Do",
    C.SUCCESS,
    "Plan Your Second Body Paragraph",
    [
      "Pick a NEW argument -- different from your first body paragraph",
      "Ideas: sacrifice of families, impact on communities, role of women, cost on young people, the home front",
      "Discuss with your partner: what is your argument? What evidence supports it?",
      "Fill in your SPO: strong TS with tier 2/3 vocabulary -> Supporting Details -> CS",
    ],
    NOTES_WEDO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 12 -- You Do
  // =========================================================================
  taskSlide(
    pres,
    "You Do",
    "Write Your Second Body Paragraph",
    [
      { label: "FIRST", instruction: "Check your SPO is complete: topic sentence, supporting details, concluding sentence" },
      { label: "NEXT", instruction: "Write your paragraph. Include at least one tier 2 or tier 3 word from the reference sheet" },
      { label: "THEN", instruction: "Reread and edit: Is your TS strong? Does each sentence support your argument? Is your vocabulary precise?" },
    ],
    NOTES_YOUDO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 13 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    "We finished War Horse today. How does Joey's story connect to what you are writing? What did his journey teach you about the role of animals in war?",
    [
      "I can write an interesting topic sentence that makes the reader want to keep reading",
      "I can plan a body paragraph using the single paragraph outline template",
      "I can use tier 2 and tier 3 vocabulary to add credibility to my persuasive writing",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 14 -- Resources
  // =========================================================================

  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: SPO Template --------------------------------------------------
  const spo = createPdf({ title: SPO_RESOURCE.name });
  let spoY = addPdfHeader(spo, "Single Paragraph Outline (SPO)", {
    color: C.PRIMARY,
    subtitle: "Second Body Paragraph Planner",
    lessonInfo: "War Horse | Lessons 28-29 | Week 6 | Year 5/6 Literacy",
    showNameDate: true,
  });

  spoY = addTipBox(spo, "This is your second body paragraph. Your argument must be DIFFERENT from your first paragraph.\n\nUse tier 2 and tier 3 vocabulary to make your writing sound knowledgeable and precise.", spoY, { color: C.PRIMARY });

  spoY = addSectionHeading(spo, "My Argument (in one sentence):", spoY, { color: C.PRIMARY });
  spoY = addLinedArea(spo, spoY, 1, { lineSpacing: 28 });
  spoY += 6;

  spoY = addSectionHeading(spo, "Topic Sentence (TS):", spoY, { color: C.PRIMARY });
  spoY = addBodyText(spo, "Make it strong: use an interesting opening, high modality, and tier 2/3 vocabulary.", spoY, { fontSize: 9, italic: true });
  spoY = addLinedArea(spo, spoY, 2, { lineSpacing: 26 });
  spoY += 6;

  spoY = addSectionHeading(spo, "Supporting Detail 1 (SD1):", spoY, { color: C.SECONDARY });
  spoY = addBodyText(spo, "Evidence from the reading that supports your argument.", spoY, { fontSize: 9, italic: true });
  spoY = addLinedArea(spo, spoY, 2, { lineSpacing: 26 });
  spoY += 6;

  spoY = addSectionHeading(spo, "Supporting Detail 2 (SD2):", spoY, { color: C.SECONDARY });
  spoY = addLinedArea(spo, spoY, 2, { lineSpacing: 26 });
  spoY += 6;

  spoY = addSectionHeading(spo, "Supporting Detail 3 (SD3) -- Optional:", spoY, { color: C.SECONDARY });
  spoY = addLinedArea(spo, spoY, 2, { lineSpacing: 26 });
  spoY += 6;

  spoY = addSectionHeading(spo, "Concluding Sentence (CS):", spoY, { color: C.ACCENT });
  spoY = addBodyText(spo, "Wrap up your argument and link back to your position.", spoY, { fontSize: 9, italic: true });
  spoY = addLinedArea(spo, spoY, 2, { lineSpacing: 26 });
  spoY += 10;

  spoY = addSectionHeading(spo, "My Full Paragraph:", spoY, { color: C.PRIMARY, fontSize: 14 });
  spoY = addLinedArea(spo, spoY, 10, { lineSpacing: 26 });

  addPdfFooter(spo, "War Horse | Lessons 28-29 | SPO Template");

  // --- PDF 2: Tier 2-3 Vocabulary Reference ---------------------------------
  const voc = createPdf({ title: VOCAB_RESOURCE.name });
  let vocY = addPdfHeader(voc, "Tier 2 and Tier 3 Vocabulary", {
    color: C.ACCENT,
    subtitle: "Persuasive Writing Word Bank",
    lessonInfo: "War Horse | Lessons 28-29 | Week 6 | Year 5/6 Literacy",
    showNameDate: false,
  });

  vocY = addTipBox(voc, "Tier 2 words are precise, formal words that work across many subjects. Tier 3 words are specific to a subject area. Using both makes your persuasive writing sound knowledgeable and authoritative.\n\nPick 2-3 words that fit naturally into your writing. Do not force in too many.", vocY, { color: C.PRIMARY });

  vocY = addSectionHeading(voc, "Tier 2: General Academic Vocabulary", vocY, { color: C.PRIMARY });
  vocY = addBodyText(voc, "significant - important and worth noting", vocY);
  vocY = addBodyText(voc, "demonstrate - to show or prove clearly", vocY);
  vocY = addBodyText(voc, "essential - absolutely necessary", vocY);
  vocY = addBodyText(voc, "consequence - a result or effect of an action", vocY);
  vocY = addBodyText(voc, "contribution - something given to help achieve a purpose", vocY);
  vocY = addBodyText(voc, "impact - a strong effect or influence", vocY);
  vocY = addBodyText(voc, "devastating - causing great damage or distress", vocY);
  vocY = addBodyText(voc, "immense - extremely large or great", vocY);
  vocY = addBodyText(voc, "evident - clearly seen or understood", vocY);
  vocY = addBodyText(voc, "fundamental - forming a necessary base or core", vocY);
  vocY += 8;

  vocY = addSectionHeading(voc, "Tier 3: World War 1 Vocabulary", vocY, { color: C.SECONDARY });
  vocY = addBodyText(voc, "conscription - compulsory enlistment for military service", vocY);
  vocY = addBodyText(voc, "armistice - an agreement to stop fighting", vocY);
  vocY = addBodyText(voc, "casualties - people killed or injured in war", vocY);
  vocY = addBodyText(voc, "commemoration - an act of honouring and remembering", vocY);
  vocY = addBodyText(voc, "allied forces - countries that fought together against a common enemy", vocY);
  vocY = addBodyText(voc, "sacrifice - giving up something valuable for a greater cause", vocY);
  vocY = addBodyText(voc, "patriotism - love and devotion to one's country", vocY);
  vocY = addBodyText(voc, "enlistment - the act of joining the armed forces", vocY);
  vocY += 8;

  vocY = addSectionHeading(voc, "Example Sentences", vocY, { color: C.ACCENT });
  vocY = addBodyText(voc, "Basic: \"Many people helped during the war and it changed things.\"", vocY);
  vocY = addBodyText(voc, "With tier 2/3: \"The contribution of Australian soldiers was significant, and the devastating consequences of the conflict fundamentally changed the nation.\"", vocY, { italic: true });
  vocY += 4;
  vocY = addBodyText(voc, "Basic: \"Women worked hard when the men were away fighting.\"", vocY);
  vocY = addBodyText(voc, "With tier 2/3: \"The essential contribution of women during the war demonstrated their immense capacity for leadership and resilience.\"", vocY, { italic: true });

  addPdfFooter(voc, "War Horse | Lessons 28-29 | Tier 2-3 Vocabulary Reference");

  // --- Write all files ------------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/WH6_Lesson28.pptx` }),
    writePdf(spo, SPO_PDF_PATH),
    writePdf(voc, VOCAB_PDF_PATH),
  ]);

  console.log(`PPTX written to ${OUT_DIR}/WH6_Lesson28.pptx`);
  console.log(`Done: ${SPO_RESOURCE.name}`);
  console.log(`Done: ${VOCAB_RESOURCE.name}`);
}

build().catch(console.error);
