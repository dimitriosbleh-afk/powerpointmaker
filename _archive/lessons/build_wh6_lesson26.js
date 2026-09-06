"use strict";

// War Horse Unit -- Lessons 26-27: Chapters 19-20 + Conjunctions
// Week 6, Session 5, Grade 5/6 Literacy
// War Horse Chapters 19-20 reading + because/but/so + subordinating conjunctions

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
  addProblem,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 5;
const FOOTER = "War Horse | Lessons 26-27 | Week 6 | Year 5/6 Literacy";
const OUT_DIR = "output/WH6_Lesson26_Chapters19_20";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const WORKSHEET_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Conjunctions Worksheet",
  "Student worksheet: complete sentences using coordinating and subordinating conjunctions."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Answer Key",
  "Teacher reference: model answers for conjunction sentence completion."
);
const RESOURCE_ITEMS = [WORKSHEET_RESOURCE, ANSWER_KEY_RESOURCE];
const WORKSHEET_PDF_PATH = path.join(OUT_DIR, WORKSHEET_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to Week 6. Today we are reading the final chapters of War Horse -- Chapters 19 and 20
- Big events ahead: the war ends, but there is a twist about the horses
- Our writing focus is conjunctions -- because, but, so, and subordinating conjunctions

DO:
- Display title slide as students settle
- Have copies of War Horse on desks, bookmarked at Chapter 19 (p. 158)

TEACHER NOTES:
Session 5 of the unit, first session of Week 6. Combines Lessons 26-27. Chapters 19-20 cover the end of the war, the devastating news about the horses, and the dramatic auction. The sentence-level writing covers both coordinating (because/but/so) and subordinating conjunctions.

WATCH FOR:
- Students who are eager to find out what happens -- channel that into careful reading
- Students who need a recap of Chapter 18 (Joey's illness and recovery)

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands: reading the next chapters and working on our sentence writing using conjunctions
- Read the success criteria. SC1 is about character analysis -- what do their actions reveal? SC2 is about using because, but, and so. SC3 is about subordinating conjunctions for complex sentences

DO:
- Choral read the LI, then the SCs
- If needed, briefly clarify: a conjunction is a joining word that connects ideas

TEACHER NOTES:
SC1 targets the reading comprehension strand through character perspective analysis. SC2 and SC3 are the sentence-level writing targets -- coordinating then subordinating conjunctions. The progression is deliberate: simpler joining words first, then more complex sentence structures.

WATCH FOR:
- Students who are unsure what "subordinating conjunction" means -- reassure: "We will learn this together. You already know some of these words"
- Students who are more focused on the story than the writing -- use the story content for the conjunction examples

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_SOLEMN = `SAY:
- Our first vocabulary word: solemn. Read it with me: solemn [students repeat]
- Solemn means very serious and without any humour. A solemn expression shows gravity and weight -- this is not a moment for laughing
- In Chapter 19, the atmosphere is solemn. The war is ending but the news about the horses changes everything
- Ask: What is the difference between being sad and being solemn? [Solemn is a deeper, quieter seriousness. You can be sad and cry. Solemn is more still -- a heavy, respectful seriousness]

DO:
- Display word, choral read, define, give example
- Connect to the chapter: the men receive solemn news about the horses
- Cold Call 2 students: when have you experienced a solemn moment?

TEACHER NOTES:
"Solemn" appears on p.163 ("shook each other solemnly by the hand"). Pre-teaching it prepares students for the emotional weight of the chapter ending. The word also connects to the theme of dignity and respect throughout the novel.

WATCH FOR:
- Students who think solemn just means sad -- refine: "Solemn is about seriousness and gravity, not just sadness. A funeral can be solemn. A judge can look solemn"
- Students who connect to Remembrance/ANZAC Day ceremonies -- excellent real-world connection

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_VOCAB_GRIM = `SAY:
- Second word: grim. Read it with me: grim [students repeat]
- Grim means very serious, gloomy, or worrying. Something grim makes you feel uneasy or concerned. A grim situation looks bad and there seems to be no easy way out
- In Chapter 20, there are grim moments during the auction. The arrival of the butcher from Cambrai is described as grim -- a man who buys horses not to care for them
- Ask: If someone described the weather as grim, what would it look like? [Dark, grey, cold, threatening -- not just cloudy but heavy and unpleasant]

DO:
- Display word, choral read, define, example
- Connect to the auction scene: the grim reality that horses may be sold to the wrong person
- Thumbs up/down: "A birthday party is grim." [Thumbs down -- parties are not grim]

TEACHER NOTES:
"Grim" is a versatile Tier 2 word that appears in many contexts. In Chapter 20, the word captures the tension of the auction -- the stakes are high and the outcome is uncertain. It also connects to the broader grim realities of war.

WATCH FOR:
- Students who confuse grim with angry -- "Grim is not about anger. It is about a heavy, serious, worrying mood"
- Students who think grim only describes people -- it also describes situations, weather, news, and places

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_READING = `SAY:
- Chapters 19 and 20 today. Major events ahead
- Chapter 19: the war carries on, then something devastating happens, and then the war ends -- but with a cruel twist for the horses
- Chapter 20: a dramatic auction that will decide Joey's fate
- Reading mode: student read aloud. We have three pause points across both chapters
- Pages 158 to 174. Find page 158 now

DO:
- Give students 30 seconds to find p. 158
- Select first reader -- choose a confident reader
- Plan reader rotations. Chapter 20 (p.167) is a natural break point for new readers

TEACHER NOTES:
These two chapters form one emotional arc: the war ends, but the horses face a new threat. David's death (Ch19) and the auction (Ch20) are the emotional peaks. The men's solidarity in pooling money is a key character moment.

WATCH FOR:
- Students upset by David's death (p.161) -- it is sudden and shocking. Allow a moment
- Students who are anxious about Joey's fate at the auction -- this is excellent engagement

[General: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause here. "Major Martin said nothing, but turned on his heel and walked away." Page 166
- How are things looking for Joey and Albert now? [The war has ended but all horses must stay in France. They will be sold at auction. Albert cannot take Joey home]
- Ask: Why does Major Martin walk away without speaking? What does this tell us about him? [He is upset too. He knows this is cruel but cannot change it. His silence shows he disagrees with the order but must follow it. Walking away is his way of showing he cannot bear to discuss it]
- Notice: Morpurgo uses repetition in Chapter 19 (p.159): "back to the war, back to the front line, back to the whine and roar of the shells." Why? [The repetition shows the relentless, endless nature of war. Each "back to" adds another burden]

DO:
- Display the quote, read aloud
- Give 10 seconds think time
- Use Think-Pair-Share: 30 seconds think, 30 seconds pair, share

CFU CHECKPOINT:
Technique: Think-Pair-Share

Script:
- "Think for 30 seconds: what does Major Martin's silence and walking away reveal about his character?"
- "Share with your partner for 30 seconds"
- "I am selecting pairs. [Name], what did your pair discuss?"
- Scan for: students identifying Martin's internal conflict -- following orders he disagrees with

PROCEED (>=80%): Most pairs identify the emotional tension. Continue to Chapter 20.
PIVOT (<80%): Most likely issue -- students describe WHAT Martin does but not WHY. Reteach: "Major Martin is a military officer. He has to follow orders. But walking away without a word -- that is not just following orders. What feeling makes someone turn away and say nothing?" [Grief, frustration, helplessness] Re-check: "What one word describes how Martin feels about this order?"

TEACHER NOTES:
This pause covers the Chapter 19 climax -- the devastating news that horses cannot return to England. Martin's wordless departure is a powerful character moment. The repetition on p.159 foreshadows the relentless final stretch of the war.

WATCH FOR:
- Students who think Martin is cold or uncaring because he walks away -- redirect: "Sometimes silence says more than words"
- Students who connect the auction announcement to real historical practice -- the British army did sell horses in France after WW1

[General: Pause Point 1 | VTLM 2.0: Teacher-Led Discussion]`;

const NOTES_PAUSE2 = `SAY:
- Chapter 20 now. "It's the last thing we'll be doing for 'em, least we can do for 'em seems to me." Page 169
- What is going on? What does the author want us to know? [Sergeant Thunder is organising a collection -- all the men are pooling their money to try to buy Joey at the auction. Even Major Martin contributed. They are doing everything they can]
- Ask: Why is it important that Morpurgo shows ALL the men contributing, not just Albert? [It shows the bond between soldiers and horses is universal, not just one man's story. The horses earned the respect and love of everyone. This strengthens the argument that animals were valued in the war]

DO:
- Display the quote
- Cold Call 3 students for responses
- Push for the author's PURPOSE -- why did Morpurgo write this scene this way?

TEACHER NOTES:
Sergeant Thunder's leadership and Major Martin's contribution show the universal bond between soldiers and horses. This connects to the persuasive writing topic -- the value of animals in war is not just one person's opinion but shared by many.

WATCH FOR:
- Students who focus only on Albert -- broaden: "This is not just Albert's story anymore. Every man in this hospital feels the same"
- Students who notice Martin's contribution -- he officially follows orders but personally supports the men

[General: Pause Point 2 | VTLM 2.0: Higher-Order Questioning]`;

const NOTES_PAUSE3 = `SAY:
- Final pause. "...and there was some delay before he brought his hammer down on the table and I was sold." Page 174
- What just happened? [Joey has been sold at auction. An old man -- Emilie's grandfather -- outbid the butcher. The auctioneer hesitated, creating tension, but the grandfather won]
- Ask: Why does Morpurgo include the butcher character? [To create tension and raise the stakes. Without a villain at the auction, there is no drama. The butcher represents the worst possible outcome for Joey -- being sold for meat. This makes the grandfather's bid feel like a rescue]
- The characterisation on p.172 is powerful: "a thin, wiry little man with weasel eyes who wore on his face a smile so full of consummate greed." Morpurgo wants us to FEEL the threat

DO:
- Display the quote
- Let students react -- this is a high-tension moment
- Cold Call: "Did you expect this outcome? Why or why not?"

TEACHER NOTES:
The auction is the narrative climax of the novel's final act. The grandfather's appearance connects back to the Emilie chapters, creating a satisfying story loop. The characterisation of the butcher (p.172) is a masterclass in villain description -- worth noting for students' own descriptive writing.

WATCH FOR:
- Students who remember Emilie's grandfather from earlier chapters -- excellent recall and connection-making
- Students who are relieved -- the emotional release after tension is a sign of deep engagement
- Students who notice the "delay" and wonder why the auctioneer hesitated

[General: Pause Point 3 | VTLM 2.0: Deep Comprehension]`;

const NOTES_CONJ_IDO = `SAY:
- Now sentence-level writing. Today we are working with conjunctions -- joining words that connect ideas
- Three coordinating conjunctions first: because, but, so
- "Because" tells us WHY something is true. "But" shows a change in direction. "So" shows what happens as a result
- Watch: "The horses cannot return to England..." I can finish this three different ways:
- "...because they have been deemed not healthy enough to take space on the transport home." -- WHY
- "...but many of the men are opposed to this order." -- CHANGE
- "...so they will be sold at auction in France." -- RESULT
- Now subordinating conjunctions. These start a DEPENDENT clause -- a part that cannot stand alone
- Unless, after, even though, although, while, when, before, if
- "Unless the men pool their money together, they have no chance of buying Joey."
- The "unless" part cannot stand by itself. It DEPENDS on the main clause

DO:
- Display the conjunction types and examples
- Read each example aloud, emphasising the conjunction
- Point out: coordinating = joins equal ideas. Subordinating = one idea depends on the other

MISCONCEPTIONS:
- Misconception: "Because" always comes in the middle of a sentence
  Why: Students see "because" in mid-sentence position most often and overgeneralise
  Impact: Students avoid starting sentences with "Because..." which limits their sentence variety
  Quick correction: "You CAN start a sentence with 'because' -- just make sure you finish the thought. 'Because the war ended, the horses were no longer needed.' That is a complete sentence"

TEACHER NOTES:
This slide combines the conjunction work from Lessons 26 and 27. The coordinating conjunctions (because/but/so) are revision; the subordinating conjunctions are the new learning. All examples use War Horse content to maintain textual connection. The sentence stems from the lesson plan are used directly.

WATCH FOR:
- Students who confuse "but" and "so" -- "but changes direction, so shows a result. 'I was tired BUT I kept going' vs 'I was tired SO I went to bed'"
- Students who write fragments after subordinating conjunctions -- they need both the dependent AND independent clause

[General: I Do -- Conjunctions | VTLM 2.0: Explicit Teaching]`;

const NOTES_WEDO = `SAY:
- Let's practise together. I will give you a dependent clause starting with a subordinating conjunction. You add the independent clause to complete the sentence
- "After Sergeant Thunder is outbid at the auction, ..."
- With your partner, discuss: how could we finish this sentence? [Take suggestions]
- One possibility: "After Sergeant Thunder is outbid at the auction, him and all the other men feel defeated and helpless."
- Now your turn with a partner. Complete this one: "Even though Albert's father was Joey's original owner, ..."

DO:
- Display the dependent clauses one at a time
- Give 30 seconds partner discussion time for each
- Take 2-3 responses for each, discuss quality
- Model if needed: show how the independent clause completes the meaning

CFU CHECKPOINT:
Technique: Finger Voting

Script:
- Show three options for completing "Even though Albert's father was Joey's original owner, ...":
  A) "Even though Albert's father was Joey's original owner. He felt sad."
  B) "Even though Albert's father was Joey's original owner, Emilie's grandfather believes Joey should belong to him."
  C) "Even though Albert's father was Joey's original owner, because he sold him."
- "Hold up 1, 2, or 3 fingers for A, B, or C."
- Scan for: 2 fingers (B is correct -- complete sentence with independent clause)

PROCEED (>=80%): Most show 2 fingers. Move to You Do.
PIVOT (<80%): Most likely issue -- students choose A (two separate sentences, not a complex sentence) or C (double conjunction). Reteach: "A has a full stop after 'owner' -- that cuts the sentence in half. The dependent clause needs to be JOINED to the independent clause. C adds 'because' which creates a double conjunction -- pick one. B flows naturally: 'Even though X, Y.'" Re-check: "Complete this: 'Although Joey was exhausted, ...' Is your answer a complete sentence with one conjunction?"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Students complete sentences using only because, but, and so (coordinating conjunctions). Provide a sentence stem: "The horses cannot return to England..." and three sentence starters: "...because ___", "...but ___", "...so ___"
- Extra Notes: Focus on the meaning of each conjunction rather than subordinating structures

EXTENDING PROMPT:
- Task: Write three complex sentences about the auction scene using three DIFFERENT subordinating conjunctions (e.g., unless, although, while). Then identify the dependent and independent clause in each

TEACHER NOTES:
The We Do uses the exact dependent clauses from the Lesson 27 plan. The finger voting CFU tests whether students understand what makes a complete complex sentence versus a fragment or double conjunction.

WATCH FOR:
- Students who write two separate sentences instead of one complex sentence -- "Join them. The subordinating conjunction is the bridge"
- Students who add extra conjunctions -- "One conjunction per join. If you have 'even though' you do not need 'because' as well"

[General: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_YOUDO = `SAY:
- On the worksheet, you have six sentences to complete
- First: complete three sentences using because, but, and so. The sentence stem is about the War Horse characters
- Next: complete three sentences by adding an independent clause to the dependent clause
- Then: write one original complex sentence about Chapters 19 or 20 using a subordinating conjunction of your choice

DO:
- Distribute the Session 5 Conjunctions Worksheet
- Circulate -- check for complete sentences, not fragments
- After 7 minutes: "You should be working on your original sentence by now"

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Complete only the first three sentences (because/but/so). Use the reading content to help generate ideas
- Extra Notes: These students may benefit from oral rehearsal before writing

EXTENDING PROMPT:
- Task: After completing all exercises, write a short paragraph (3-4 sentences) about the auction that uses at least two different subordinating conjunctions and one coordinating conjunction. Underline each conjunction

TEACHER NOTES:
The You Do uses different sentence content from the We Do to ensure genuine transfer. The original sentence at the end requires students to generate both the dependent and independent clause, demonstrating full understanding.

WATCH FOR:
- Students who write fragments -- the dependent clause without the independent clause
- Students who use the same conjunction for every sentence -- encourage variety
- Students who write sophisticated complex sentences -- affirm and share with the class

[General: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: what characters' actions reveal about their perspective -- thumbs? [scan]
- SC2: using because, but, and so -- thumbs? [scan]
- SC3: using subordinating conjunctions for complex sentences -- thumbs? [scan]
- Turn and talk: What was the most dramatic moment across Chapters 19 and 20? Can you describe it using a complex sentence with a subordinating conjunction?

DO:
- Run through each SC with thumbs check
- The turn-and-talk combines reading comprehension with sentence writing -- a natural integration
- Preview: "Next session we read the final chapter of War Horse and continue our persuasive writing"

TEACHER NOTES:
The closing integrates both strands -- reading comprehension and sentence writing. Students who can describe a dramatic moment using a complex sentence are demonstrating both SCs simultaneously.

WATCH FOR:
- Students thumbs-down on SC3 -- subordinating conjunctions may need revisiting in the next sentence-level writing focus
- Students who describe the auction as the most dramatic moment -- this is the intended climax

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two printable resources today
- The ${WORKSHEET_RESOURCE.name} has sentence completion exercises for both types of conjunctions
- The ${ANSWER_KEY_RESOURCE.name} is for teacher reference

DO:
- Print the worksheet before the lesson (one per student)
- Print the answer key (teacher copy only)
- Click any resource card to open the PDF

TEACHER NOTES:
The worksheet exercises use War Horse content exclusively, maintaining the textual connection throughout the session.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "War Horse - Lessons 26-27 - Chapters 19-20 + Conjunctions";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "War Horse",
    "Chapters 19-20 -- The Auction",
    "Lessons 26-27  |  Week 6  |  Year 5/6 Literacy",
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
      "We are learning to analyse how authorial choices shape perspective, and to use conjunctions to build compound and complex sentences",
    ],
    [
      "I can explain what a character's actions and dialogue reveal about their perspective",
      "I can use because, but, and so to extend sentences with clear meaning",
      "I can use subordinating conjunctions to write complex sentences that vary my writing",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Vocabulary: solemn
  // =========================================================================
  vocabSlide(
    pres,
    "solemn",
    "adjective",
    "Very serious and without any humour. A solemn expression shows gravity and weight -- this is not a moment for laughing or joking.",
    "The men shook each other solemnly by the hand, their faces heavy with the knowledge that their horses would not be coming home.",
    NOTES_VOCAB_SOLEMN,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Vocabulary: grim
  // =========================================================================
  vocabSlide(
    pres,
    "grim",
    "adjective",
    "Very serious, gloomy, or worrying. Something grim makes you feel uneasy and concerned -- there seems to be no easy way out.",
    "A grim silence fell over the yard as the men realised what the auction could mean for their beloved horses.",
    NOTES_VOCAB_GRIM,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- Reading Launch
  // =========================================================================
  contentSlide(
    pres,
    "Read Aloud",
    C.PRIMARY,
    "Chapters 19-20 -- Pages 158-174",
    [
      "Reading Mode: Student Read Aloud",
      "Ch 19: The war continues, a sudden loss, and devastating news about the horses",
      "Ch 20: The men's plan, the dramatic auction, and an unexpected rescuer",
      "Focus: what do characters' actions reveal about their perspective and values?",
    ],
    NOTES_READING,
    FOOTER
  );

  // =========================================================================
  // SLIDE 6 -- Pause Point 1 (Ch19 p.166)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 1", "Chapter 19 -- p. 166",
    "Major Martin said nothing, but turned on his heel and walked away.",
    "p. 166",
    "How are things looking for Joey and Albert now? Why does Martin walk away in silence?",
    NOTES_PAUSE1, FOOTER
  );

  // =========================================================================
  // SLIDE 7 -- Pause Point 2 (Ch20 p.169)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 2", "Chapter 20 -- p. 169",
    "It's the last thing we'll be doing for 'em, least we can do for 'em seems to me.",
    "p. 169",
    "What's going on? What does the author want us to know about the men's bond with the horses?",
    NOTES_PAUSE2, FOOTER
  );

  // =========================================================================
  // SLIDE 8 -- Pause Point 3 (Ch20 p.174)
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 3", "Chapter 20 -- p. 174",
    "...and there was some delay before he brought his hammer down on the table and I was sold.",
    "p. 174",
    "What just happened? Why did Morpurgo include the butcher character?",
    NOTES_PAUSE3, FOOTER
  );

  // =========================================================================
  // SLIDE 9 -- Conjunctions I Do
  // =========================================================================
  modellingSlide(
    pres,
    "I Do",
    "Conjunctions -- Joining Ideas",
    "Coordinating: because, but, so\n\n\"The horses cannot return to England...\"\n\n...because they have been deemed not healthy enough for transport.\n(WHY)\n\n...but many of the men are opposed.\n(CHANGE)\n\n...so they will be sold at auction.\n(RESULT)",
    "Subordinating: unless, after, even though, although, while, when\n\nThese start a DEPENDENT clause:\n\n\"Unless the men pool their money, they have no chance of buying Joey.\"\n\n\"After Sergeant Thunder is outbid, the men feel defeated.\"\n\n\"Even though Albert's father was Joey's original owner, the grandfather claims him.\"",
    NOTES_CONJ_IDO,
    FOOTER
  );

  // =========================================================================
  // SLIDES 10-11 -- We Do: Conjunction Practice (withReveal)
  // =========================================================================
  withReveal(
    () => contentSlide(
      pres,
      "We Do",
      C.SUCCESS,
      "Complete the Sentences",
      [
        "Add an independent clause to complete each sentence:",
        "\"After Sergeant Thunder is outbid at the auction, ...\"",
        "\"Even though Albert's father was Joey's original owner, ...\"",
        "Discuss with your partner, then write your answer",
      ],
      NOTES_WEDO,
      FOOTER
    ),
    (slide) => {
      const ansY = 3.90;
      slide.addShape("roundRect", {
        x: 0.5, y: ansY, w: 9, h: 1.10, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      });

      slide.addShape("roundRect", {
        x: 0.7, y: ansY + 0.08, w: 1.7, h: 0.28, rectRadius: 0.08,
        fill: { color: C.WHITE },
      });
      slide.addText("Model Answers", {
        x: 0.7, y: ansY + 0.08, w: 1.7, h: 0.28,
        fontSize: 10, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      slide.addText("\"After Sergeant Thunder is outbid, him and the men feel defeated and helpless.\"\n\"Even though Albert's father was Joey's original owner, Emilie's grandfather believes Joey should belong to him.\"", {
        x: 2.6, y: ansY + 0.06, w: 6.7, h: 0.96,
        fontSize: 12.5, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_WEDO);
    }
  );

  // =========================================================================
  // SLIDE 12 -- You Do
  // =========================================================================
  taskSlide(
    pres,
    "You Do",
    "Conjunction Practice",
    [
      { label: "FIRST", instruction: "Complete three sentences using because, but, and so with the War Horse sentence stem on the worksheet" },
      { label: "NEXT", instruction: "Complete three sentences by adding an independent clause to each dependent clause" },
      { label: "THEN", instruction: "Write one original complex sentence about Chapters 19-20 using a subordinating conjunction" },
    ],
    NOTES_YOUDO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 13 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    "What was the most dramatic moment across Chapters 19 and 20? Describe it using a complex sentence with a subordinating conjunction.",
    [
      "I can explain what a character's actions and dialogue reveal about their perspective",
      "I can use because, but, and so to extend sentences with clear meaning",
      "I can use subordinating conjunctions to write complex sentences that vary my writing",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 14 -- Resources
  // =========================================================================

  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: Conjunctions Worksheet ----------------------------------------
  const ws = createPdf({ title: WORKSHEET_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Conjunctions Worksheet", {
    color: C.PRIMARY,
    subtitle: "Chapters 19-20: The Auction",
    lessonInfo: "War Horse | Lessons 26-27 | Week 6 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "Coordinating conjunctions join equal ideas:\n- because = tells WHY\n- but = shows a CHANGE in direction\n- so = shows a RESULT\n\nSubordinating conjunctions start a dependent clause (a part that cannot stand alone):\n- unless, after, even though, although, while, when, before, if", wsY, { color: C.PRIMARY });

  // Part A: because/but/so
  wsY = addSectionHeading(ws, "Part A: Complete using because, but, or so", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Sentence stem: \"The horses cannot return to England...\"", wsY, { fontSize: 12, bold: true });
  wsY += 4;

  wsY = addBodyText(ws, "1. The horses cannot return to England because", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 26 });
  wsY += 4;

  wsY = addBodyText(ws, "2. The horses cannot return to England but", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 26 });
  wsY += 4;

  wsY = addBodyText(ws, "3. The horses cannot return to England so", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 26 });
  wsY += 10;

  // Part B: subordinating conjunctions
  wsY = addSectionHeading(ws, "Part B: Add an independent clause to complete the sentence", wsY, { color: C.SECONDARY });

  wsY = addBodyText(ws, "4. Unless the men all pool their money together,", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 26 });
  wsY += 4;

  wsY = addBodyText(ws, "5. After Sergeant Thunder is outbid at the auction,", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 26 });
  wsY += 4;

  wsY = addBodyText(ws, "6. Even though Albert's father was Joey's original owner,", wsY, { fontSize: 11 });
  wsY = addLinedArea(ws, wsY, 1, { lineSpacing: 26 });
  wsY += 10;

  // Part C: original sentence
  wsY = addSectionHeading(ws, "Part C: Write your own", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "Write one original complex sentence about Chapters 19 or 20. Use a subordinating conjunction and underline it.", wsY, { fontSize: 10, italic: true });
  wsY = addLinedArea(ws, wsY, 2, { lineSpacing: 26 });

  addPdfFooter(ws, "War Horse | Lessons 26-27 | Conjunctions Worksheet");

  // --- PDF 2: Answer Key ---------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Conjunctions Worksheet -- Answer Key", {
    color: C.ALERT,
    subtitle: "Teacher Reference -- Chapters 19-20",
    lessonInfo: "War Horse | Lessons 26-27 | Week 6 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Accept any answer that: (1) uses the correct conjunction type, (2) completes a grammatically correct sentence, and (3) makes sense in the context of the story. Model answers below show one possible response.", akY, { color: C.ALERT });

  akY = addSectionHeading(ak, "Part A: because / but / so", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "1. ...because they have been deemed not healthy enough to take up space on the transport home.", akY);
  akY = addBodyText(ak, "2. ...but many of the men are opposed to this order from Major Martin.", akY);
  akY = addBodyText(ak, "3. ...so they will be sold off in an auction in France.", akY);
  akY += 8;

  akY = addSectionHeading(ak, "Part B: Subordinating conjunctions", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "4. Unless the men all pool their money together, they would have no chance to be able to take Joey back to England.", akY);
  akY = addBodyText(ak, "5. After Sergeant Thunder is outbid at the auction, him and all the other men from the veterinary hospital feel defeated and helpless.", akY);
  akY = addBodyText(ak, "6. Even though Albert's father was Joey's original owner, Emilie's grandfather believes Joey should belong to him and his Emilie.", akY);
  akY += 8;

  akY = addSectionHeading(ak, "Part C: Student original -- marking guide", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "Check for:", akY, { bold: true });
  akY = addBodyText(ak, "- Uses a subordinating conjunction (unless, after, even though, although, while, when, before, if)", akY);
  akY = addBodyText(ak, "- Contains both a dependent clause AND an independent clause", akY);
  akY = addBodyText(ak, "- Makes sense in the context of Chapters 19-20", akY);
  akY = addBodyText(ak, "- Is a single, complete sentence (not two sentences or a fragment)", akY);
  akY += 6;
  akY = addBodyText(ak, "Example: \"Although the men raised enough money, they were still outbid by the butcher from Cambrai.\"", akY, { italic: true });

  addPdfFooter(ak, "War Horse | Lessons 26-27 | Answer Key -- TEACHER COPY");

  // --- Write all files ------------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/WH6_Lesson26.pptx` }),
    writePdf(ws, WORKSHEET_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  console.log(`PPTX written to ${OUT_DIR}/WH6_Lesson26.pptx`);
  console.log(`Done: ${WORKSHEET_RESOURCE.name}`);
  console.log(`Done: ${ANSWER_KEY_RESOURCE.name}`);
}

build().catch(console.error);
