"use strict";

// War Horse Unit -- Lesson 21: Chapter 16 -- The Coin Toss
// Week 5, Session 1, Grade 5/6 Literacy
// Chapter 16 reading + subordinating conjunctions sentence-level writing

const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const {
  C, FONT_H, FONT_B,
  SAFE_BOTTOM, CONTENT_TOP,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  withReveal,
  titleSlide, liSlide, contentSlide,
  cfuSlide, taskSlide, closingSlide,
  vocabSlide, quoteSlide, modellingSlide,
} = require("../themes/wh5_helpers");

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 1;
const FOOTER = "War Horse | Lesson 21 of 25 | Week 5 | Year 5/6 Literacy";
const OUT_DIR = "output/WH5_Lesson21_The_Coin_Toss";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const WORKSHEET_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Subordinating Conjunctions Worksheet",
  "Student worksheet: add an independent clause to complete three complex sentences using subordinating conjunctions."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Answer Key",
  "Teacher reference: model completions for subordinating conjunction sentences with alternative valid answers."
);
const RESOURCE_ITEMS = [WORKSHEET_RESOURCE, ANSWER_KEY_RESOURCE];
const WORKSHEET_PDF_PATH = path.join(OUT_DIR, WORKSHEET_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to Week 5. We are back with War Horse today for Chapter 16
- Last week we started our persuasive writing about horses in WWI. Today we return to the novel
- Chapter 16 picks up right where Chapter 15 left off -- Joey is stuck in no man's land

DO:
- Display title slide as students settle
- Have copies of War Horse on desks, bookmarked at Chapter 16 (p. 127)

TEACHER NOTES:
Lesson 21 of 25, first session of Week 5. Chapter 16 features a remarkable moment of humanity between enemy soldiers. The subordinating conjunctions work connects naturally to the novel's complex sentences.

WATCH FOR:
- Students who may need a brief recap of Chapter 15 -- Joey ran through the battlefield, ended up in no man's land
- Students still processing the emotional content from the previous chapters

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands today: reading Chapter 16 and building complex sentences with subordinating conjunctions
- Read the success criteria. Ask: What do you think a subordinating conjunction is? Some of you may remember this from earlier in the year [Take 1-2 guesses]

DO:
- Choral read the LI, then the SCs
- Do not over-explain subordinating conjunctions yet -- the I Do covers this

TEACHER NOTES:
SC1 targets character analysis through dialogue and action. SC2 addresses literary devices (metaphor and personification). SC3 is the sentence-level writing target. The progression moves from comprehension to analysis to application.

WATCH FOR:
- Students who look blank at "subordinating conjunctions" -- reassure: "We will learn this together step by step"

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_WARY = `SAY:
- Our first vocabulary word: wary. Read it with me: wary [students repeat]
- Wary means cautious and watchful, especially because you think something might be dangerous or go wrong
- In Chapter 16, the soldiers are wary of each other. They are enemies who have climbed into no man's land -- of course they are watchful and cautious
- Quick check: if you were wary of a dog, would you run towards it or keep your distance? [Keep your distance]

DO:
- Display word, choral read, define, give example
- Connect to the chapter: the soldiers' wariness is a key part of the tension
- Cold Call 2 students: use "wary" in a sentence about something from their own life

TEACHER NOTES:
Pre-teaching "wary" before reading ensures students can process the soldiers' cautious interaction without stopping to decode vocabulary. The word carries significant weight in this chapter.

WATCH FOR:
- Students who confuse "wary" with "weary" (tired) -- clarify the spelling and meaning difference
- Students who think "wary" only means scared -- refine: "Wary is careful and watchful, not necessarily afraid"

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_VOCAB_COARSE = `SAY:
- Second word: coarse. Read it with me: coarse [students repeat]
- Coarse means rough in texture or quality. It can describe physical things like coarse sand or coarse fabric, or it can describe language that is rough or vulgar
- In Chapter 16, Morpurgo uses "coarse" to describe a character's voice or appearance -- rough, weathered
- Ask: What is the OPPOSITE of coarse? [Smooth, soft, fine, refined]

DO:
- Display word, choral read, define, example
- Quick synonym/antonym check to deepen understanding
- Ask 1-2 students to describe something coarse from their experience

TEACHER NOTES:
"Coarse" is a versatile Tier 2 word. Teaching both the physical and figurative meanings expands students' descriptive vocabulary for their own writing.

WATCH FOR:
- Students who confuse "coarse" with "course" (a path or class) -- clarify the spelling and that these are different words
- Students who only think of "coarse" as rude language -- expand to texture

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_READING_INTRO = `SAY:
- Chapter 16. Joey is stuck in no man's land, trapped in barbed wire. As day breaks, both sides of soldiers can see him
- Today's reading is Student Read Aloud. I will select readers. We have three pause points
- Focus: watch how the two soldiers interact. They are enemies in a war, but they find common ground over a horse
- Pages 127 to 134 -- find your page now

DO:
- Give students 30 seconds to find p. 127
- Select first reader -- choose a confident reader for the opening
- Plan reader rotations: change every half-page or at natural breaks

TEACHER NOTES:
Chapter 16 is a pivotal chapter. The coin toss scene between the Welsh and German soldiers shows humanity persisting in the midst of war. Students should notice that the soldiers treat each other with respect despite being enemies.

SENSITIVITY ADVISORY:
- What it is: The chapter takes place in no man's land during WWI. Joey is injured and trapped. The setting is a battlefield
- Framing language: "This chapter shows something remarkable -- enemies choosing kindness"
- Watch for: Students still affected by Topthorn's death or the battlefield violence
- Protocol: Frame the chapter positively -- this is a moment of hope and humanity

WATCH FOR:
- Readers who rush the dialogue -- coach: "Give the soldiers different voices. The Welshman and the German sound different"
- Students who miss the humour in the coin toss -- point it out

[General: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause here. "He too held up a white handkerchief in one hand and began also to work his way through the wire towards me"
- What is happening? [Two soldiers from opposite sides have both climbed into no man's land, both waving white flags, both coming to help Joey]
- Ask: Why is this moment significant? Why is it remarkable that BOTH sides come out? [They are enemies -- they could shoot each other. But they both choose to help the horse instead]

DO:
- Display the quote, read aloud
- Give 10 seconds of think time before taking responses
- Use Think-Pair-Share: 30 seconds think, 30 seconds pair, then share

CFU CHECKPOINT:
Technique: Think-Pair-Share

Script:
- "Think for 30 seconds: what makes this moment remarkable?"
- "Share with your partner for 30 seconds"
- "Hands down -- I am selecting pairs. [Name], what did your pair discuss?"
- Scan for: students identifying the tension between enemies and compassion

PROCEED (>=80%): Most pairs recognise the significance of enemies cooperating. Continue reading.
PIVOT (<80%): Most likely issue -- students focus on the action (helping the horse) without seeing the deeper significance (enemies choosing compassion). Reteach: "These men will go back to shooting at each other in a few hours. But right now, they both chose to risk their lives for a horse. What does that tell us about them?" Re-check: "Is this about the horse, or is it about something bigger?"

TEACHER NOTES:
The first pause point establishes the central tension of the chapter: enemies finding common ground. This connects to the novel's broader theme about the absurdity of war.

WATCH FOR:
- Students who only see surface action -- push: "Why white handkerchiefs? What do they symbolise?"
- Students who are surprised enemies would cooperate -- great teaching moment about shared humanity

[General: Pause Point 1 | VTLM 2.0: Teacher-Led Discussion]`;

const NOTES_PAUSE2 = `SAY:
- Stop here. "Still, not your fault I don't suppose. Nor mine, neither come to that"
- What is going on? Who is talking? [The Welsh soldier is talking to Joey about the war]
- Ask: What does this quote tell us about how the soldier views the war? [He does not blame himself or Joey -- the war is not their fault. It is something that happened TO them, not something they chose]

DO:
- Display the quote
- Use Cold Call for responses
- Push for deeper interpretation -- this is a character analysis moment

TEACHER NOTES:
This quote reveals the Welsh soldier's perspective on war: it is something inflicted upon ordinary people (and animals), not something they chose. This connects directly to SC1 (character through dialogue).

WATCH FOR:
- Students who miss the significance of "nor mine, neither" -- the soldier includes himself alongside the horse as victims of the war
- Students who can identify what the character says but not what it REVEALS about his thinking

[General: Pause Point 2 | VTLM 2.0: Higher-Order Questioning]`;

const NOTES_PAUSE3 = `SAY:
- Final pause. "...the line of khaki soldiers who began now to laugh and cheer with delight as I limped towards them through the gap in the wire"
- What have we learned from this whole scene? What does the author want us to know? [Even in war, people can choose compassion. The soldiers on both sides are cheering -- they are united in joy over a horse being saved]
- Ask: Why does Morpurgo have the soldiers LAUGH and CHEER? [It humanises them. For a moment, they are not enemies -- they are just people who are happy to see a horse rescued]

DO:
- Display the quote, read with warmth in your voice
- Let the moment land -- this is a hopeful scene
- Cold Call 2-3 students for their interpretation

TEACHER NOTES:
The final pause point brings the chapter's theme full circle. The laughter and cheering from both sides shows war's absurdity -- these men can share joy one moment and return to fighting the next.

WATCH FOR:
- Students who see this as simply "happy ending" -- push: "Is it really a happy ending? What happens next? They go back to fighting"
- Students making connections to Lesson 19's persuasive topic -- excellent: this scene could support either side of the argument

[General: Pause Point 3 | VTLM 2.0: Deep Comprehension]`;

const NOTES_PAUSE3_REVEAL = `SAY:
- Morpurgo wants us to see that war divides people who, in ordinary life, would have no reason to be enemies
- The coin toss, the laughter, the cooperation -- all of it shows shared humanity beneath the uniforms
- This connects to our persuasive writing topic. Whether you are arguing for or against using horses, this chapter gives you evidence about how the war affected both the animals AND the people

DO:
- Reveal the answer card
- Draw the connection to the persuasive writing unit
- Transition to literary terms: "Now let's look at the specific language Morpurgo uses in this chapter"

[General: Pause Point 3 Reveal | VTLM 2.0: Deep Comprehension]`;

const NOTES_LITERARY = `SAY:
- Chapter 16 uses metaphor and personification. Let's look at both
- Metaphor: "his face was lined and creased with years" -- the Welsh soldier's face IS lined and creased. This is not literally about lines drawn on his face -- it means the years of hardship have marked him
- Metaphor: "his face red and still smooth with youth" -- the German soldier, by contrast, is young. His face IS smooth. Notice how Morpurgo uses physical descriptions as metaphors for experience and innocence
- Personification: "a smile lighting his worn face" -- can a smile actually produce light? No. But Morpurgo makes the smile an active force that brings warmth and brightness to a face worn by war

DO:
- Display all three examples
- Read each aloud with emphasis
- Ask: "What is the contrast between the two soldiers? How does Morpurgo use their faces to show it?" [Age and experience vs youth and innocence]

TEACHER NOTES:
These metaphors and the personification work together to characterise the two soldiers. The contrast between the weathered Welshman and the young German adds depth to the scene.

MISCONCEPTIONS:
- Misconception: "His face was lined with years" is literal -- he had lines on his face
  Why: Students take figurative language literally when the metaphor is subtle
  Impact: Reduces ability to interpret character descriptions in other texts
  Quick correction: "Years cannot literally line a face. What Morpurgo means is that a lifetime of experience has left its mark on him -- wrinkles, stress, hardship"

WATCH FOR:
- Students who can name the device but not explain the EFFECT -- push: "So what? What does the personification DO for the reader?"
- Students who see the contrast between the soldiers -- great: this supports SC2

[General: I Do -- Literary Terms | VTLM 2.0: Explicit Teaching]`;

const NOTES_CONJ_IDO = `SAY:
- Now we are working on sentence-level writing: subordinating conjunctions
- A subordinating conjunction joins a dependent clause to an independent clause to make a complex sentence
- Dependent clause: cannot stand alone as a sentence. It depends on something else
- Independent clause: CAN stand alone as a complete sentence
- Examples of subordinating conjunctions: since, before, even though, although, when, because, if, after, while, until
- Watch me. Here is a dependent clause: "Since Joey is trapped within the wires and can't walk freely, ..."
- I need to add an independent clause to complete the sentence: "...the soldiers need to intervene to help get him out of no man's land"
- The independent clause completes the thought. It can stand alone: "The soldiers need to intervene to help get him out of no man's land" -- that is a complete sentence

DO:
- Display the terminology and the model sentence clearly
- Point to the dependent clause and independent clause as separate parts
- Think aloud: "I check -- can my second part stand alone as a sentence? Yes, it can. So it is an independent clause"
- Show the subordinating conjunction highlighted in the dependent clause

TEACHER NOTES:
The think-aloud about checking whether the second part can stand alone is the key metacognitive move. Students need this self-checking strategy when they write their own sentences.

WATCH FOR:
- Students who look confused by the terminology -- reassure: "The names sound complex, but the idea is simple: one part can't stand alone, the other can"
- Students who mix up dependent and independent -- consistent gesture: point DOWN for dependent (needs support), point UP for independent (stands on its own)

[General: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU = `SAY:
- Quick check. I am going to read you a clause. You tell me: is it a DEPENDENT clause or an INDEPENDENT clause?
- "Before the Welsh soldier realises the German can speak English"
- Can that stand alone as a complete sentence? Think about it
- Thumbs up if it is INDEPENDENT (can stand alone), thumbs down if it is DEPENDENT (needs more)
- Three, two, one -- show me! [Scan]

DO:
- Use Thumbs Up/Down
- Scan for: thumbs down (DEPENDENT -- it starts with "before" and leaves you waiting for what happens next)

CFU CHECKPOINT:
Technique: Thumbs Up/Down

Script:
- "Thumbs up for independent, thumbs down for dependent"
- Read the clause, count down, scan
- Look for: mostly thumbs down (dependent)

PROCEED (>=80%): Most show thumbs down. Move to We Do/You Do.
PIVOT (<80%): Most likely issue -- students think any clause with a subject and verb is independent. Reteach: "Read it aloud. Does it feel complete? 'Before the Welsh soldier realises the German can speak English...' You are left waiting. Before what? That is how you know it is dependent -- it depends on more information." Re-check: "The soldiers decided to flip a coin" -- independent or dependent? [Independent -- it is a complete thought]

TEACHER NOTES:
Thumbs Up/Down is quick and decisive. The chosen clause starts with "before" which is a clear subordinating conjunction signal.

WATCH FOR:
- Students who guess without thinking -- push: "Read it in your head first. Does it feel finished?"
- Students who are confident but wrong -- the re-check with a clear independent clause helps calibrate

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- Thumbs down -- it is DEPENDENT. "Before the Welsh soldier realises the German can speak English" -- you are left hanging. Before that happened... what? It needs an independent clause to complete it
- The word "before" is the subordinating conjunction. It creates the dependency
- If I said "The Welsh soldier realises the German can speak English" -- that IS independent. Adding "before" at the start makes it dependent

DO:
- Reveal the answer
- Highlight the subordinating conjunction "before" as the signal word
- Transition to We Do/You Do

[General: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_WEDO = `SAY:
- Your turn. On the worksheet, you have three dependent clauses from Chapter 16. Each one starts with a subordinating conjunction
- You need to add an independent clause to complete each sentence
- Remember: your independent clause must be a COMPLETE thought that could stand on its own
- Let's do the first one together. "Since Joey is trapped within the wires and can't walk freely, ..." What could we add? [Take 2-3 suggestions, discuss which works best]
- Now complete the other two independently. You have 5 minutes

DO:
- Distribute the Session 1 Subordinating Conjunctions Worksheet
- Do the first sentence collaboratively (We Do), then release for independent work on sentences 2 and 3
- Circulate -- check that completions are genuine independent clauses, not fragments

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide word bank cards with key vocabulary from the chapter (intervene, cooperate, rescue, realise, communicate) and model sentence frames: "...the soldiers decide to ___" / "...he thinks it was ___"
- Extra Notes: Students can refer to the teacher's model from the I Do slide

EXTENDING PROMPT:
- Task: Write your OWN complex sentence about Chapter 16 using a subordinating conjunction not on the worksheet (e.g., although, while, until, after). Then label the dependent clause and the independent clause

TEACHER NOTES:
The transition from We Do (sentence 1) to You Do (sentences 2-3) within the same activity is a clean gradual release. Circulating during the You Do is the key assessment opportunity.

WATCH FOR:
- Students who write a second dependent clause instead of an independent one -- check: "Can your second part stand alone as a sentence?"
- Students who write very short completions ("...they helped") -- push for detail and connection to the chapter

[General: We Do / You Do | VTLM 2.0: Guided + Independent Practice]`;

const NOTES_WEDO_REVEAL = `SAY:
- Let's check our answers. Here are model completions for each sentence
- For sentence 1: "Since Joey is trapped within the wires and can't walk freely, the soldiers need to intervene to help get him out of no man's land"
- For sentence 2: "Before the Welsh soldier realises the German can speak English, he thinks it was useless to come out to try and help Joey"
- For sentence 3: "Even though the two soldiers will return to fighting one another in a few hours, they were able to be kind to each other to help Joey"
- Your answers don't need to match mine exactly. What matters is that your independent clause is a complete thought that makes sense with the dependent clause

DO:
- Reveal model completions
- Read each full sentence aloud
- Invite students to share alternatives that also work
- Celebrate creative, accurate completions

[General: We Do / You Do Reveal | VTLM 2.0: Guided + Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: key moments showing character through dialogue and action -- thumbs? [scan]
- SC2: metaphor and personification revealing character -- thumbs? [scan]
- SC3: adding an independent clause to a dependent clause -- thumbs? [scan]
- Turn and talk: What moment in Chapter 16 surprised you the most, and why?

DO:
- Run through each SC with thumbs check
- The turn-and-talk is open-ended -- let students share what struck them most
- Preview: "Next lesson we continue with Chapter 17 -- Joey arrives at the veterinary hospital"

TEACHER NOTES:
The closing reconnects to the reading experience before ending on the writing skill. Students "down" on SC3 may benefit from reviewing the worksheet at home.

WATCH FOR:
- Students "thumbs down" on SC3 -- check if it is terminology confusion or genuine difficulty with the concept
- Students who want to discuss the persuasive writing connection -- acknowledge and encourage

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two printable resources today
- The ${WORKSHEET_RESOURCE.name} is for the We Do and You Do activity
- The ${ANSWER_KEY_RESOURCE.name} is for teacher reference with multiple valid completions

DO:
- Print the worksheet before the lesson (one per student)
- Print the answer key (teacher copy only)
- Click any resource card to open the PDF

TEACHER NOTES:
The worksheet separates the three sentences with clear space for writing. The answer key includes alternative valid completions to support flexible marking.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "War Horse - Lesson 21 - Chapter 16: The Coin Toss";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "War Horse",
    "Chapter 16 -- The Coin Toss",
    "Lesson 21  |  Week 5  |  Year 5/6 Literacy",
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
  // SLIDE 2 -- Learning Intention & Success Criteria
  // =========================================================================
  liSlide(
    pres,
    [
      "We are learning to analyse how an author reveals character through dialogue and action, and to build complex sentences using subordinating conjunctions",
    ],
    [
      "I can identify key moments where the author shows character through what they say and do",
      "I can explain how the author uses metaphor and personification to reveal character",
      "I can add an independent clause to a dependent clause to create a complex sentence",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Vocabulary: wary
  // =========================================================================
  vocabSlide(
    pres,
    "wary",
    "adjective",
    "Cautious and watchful, especially because you think something might be dangerous or could go wrong. A wary person keeps their guard up and watches carefully.",
    "The soldiers were wary of each other as they carefully approached the injured horse in no man's land.",
    NOTES_VOCAB_WARY,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Vocabulary: coarse
  // =========================================================================
  vocabSlide(
    pres,
    "coarse",
    "adjective",
    "Rough in texture, quality, or manner. Can describe physical things (coarse sand, coarse fabric) or language and behaviour that is rough or unrefined.",
    "The old soldier's coarse hands told the story of a lifetime of hard work before the war began.",
    NOTES_VOCAB_COARSE,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- Reading Introduction
  // =========================================================================
  contentSlide(
    pres,
    "Read Aloud",
    C.PRIMARY,
    "Chapter 16 -- Pages 127-134",
    [
      "Reading Mode: Student Read Aloud",
      "Joey is stuck in no man's land, trapped in barbed wire",
      "Both sides can see him as day breaks",
      "Two soldiers -- one Welsh, one German -- climb out to help",
      "Focus: How do the soldiers interact? What does their behaviour reveal?",
    ],
    NOTES_READING_INTRO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 6 -- Pause Point 1
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 1", "Chapter 16 -- p. 128",
    "He too held up a white handkerchief in one hand and began also to work his way through the wire towards me.",
    "p. 128",
    "What's happening? Why is it remarkable that BOTH sides have sent someone out?",
    NOTES_PAUSE1, FOOTER
  );

  // =========================================================================
  // SLIDE 7 -- Pause Point 2
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 2", "Chapter 16 -- p. 131",
    "Still, not your fault I don't suppose. Nor mine, neither come to that.",
    "p. 131",
    "What's going on? What does this tell us about how the soldier views the war?",
    NOTES_PAUSE2, FOOTER
  );

  // =========================================================================
  // SLIDES 8-9 -- Pause Point 3 (withReveal)
  // =========================================================================
  withReveal(
    () => quoteSlide(
      pres, "Pause Point 3", "Chapter 16 -- p. 134",
      "...the line of khaki soldiers who began now to laugh and cheer with delight as I limped towards them through the gap in the wire.",
      "p. 134",
      "What have we learned from this conversation? What does the author want us to know about war and humanity?",
      NOTES_PAUSE3, FOOTER
    ),
    (slide) => {
      const ansY = 3.62;
      slide.addShape("roundRect", {
        x: 0.5, y: ansY, w: 9, h: 1.38, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      });
      slide.addText("Shared Humanity", {
        x: 0.75, y: ansY + 0.08, w: 3, h: 0.30,
        fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true, margin: 0,
      });
      slide.addText("War divides people who would otherwise have no reason to be enemies. The coin toss, the laughter, the cooperation -- all show shared humanity beneath the uniforms. Morpurgo reminds us that soldiers on both sides are just people, united by compassion for a horse.", {
        x: 0.75, y: ansY + 0.40, w: 8.5, h: 0.88,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, valign: "top", margin: 0,
      });
      slide.addNotes(NOTES_PAUSE3_REVEAL);
    }
  );

  // =========================================================================
  // SLIDE 10 -- Literary Terms: Metaphor + Personification
  // =========================================================================
  {
    const s = pres.addSlide();
    addTopBar(s, C.ACCENT);
    addBadge(s, "Literary Devices", { color: C.ACCENT, w: 2.0 });
    addTitle(s, "Metaphor & Personification in Chapter 16");

    // Metaphor 1
    const m1Y = CONTENT_TOP;
    addCard(s, 0.5, m1Y, 9, 1.05, { strip: C.PRIMARY, fill: C.WHITE });
    s.addShape("roundRect", {
      x: 0.7, y: m1Y + 0.08, w: 1.4, h: 0.26, rectRadius: 0.08,
      fill: { color: C.PRIMARY },
    });
    s.addText("Metaphor", {
      x: 0.7, y: m1Y + 0.08, w: 1.4, h: 0.26,
      fontSize: 10, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("\"...his face was lined and creased with years.\" (p. 129)", {
      x: 0.75, y: m1Y + 0.42, w: 8.4, h: 0.26,
      fontSize: 14, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    });
    s.addText("The Welsh soldier -- experienced, weathered, marked by life", {
      x: 0.75, y: m1Y + 0.72, w: 8.4, h: 0.26,
      fontSize: 11, fontFace: FONT_B, color: C.MUTED, margin: 0,
    });

    // Metaphor 2
    const m2Y = m1Y + 1.17;
    addCard(s, 0.5, m2Y, 9, 1.05, { strip: C.SECONDARY, fill: C.WHITE });
    s.addShape("roundRect", {
      x: 0.7, y: m2Y + 0.08, w: 1.4, h: 0.26, rectRadius: 0.08,
      fill: { color: C.SECONDARY },
    });
    s.addText("Metaphor", {
      x: 0.7, y: m2Y + 0.08, w: 1.4, h: 0.26,
      fontSize: 10, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("\"...his face red and still smooth with youth...\" (p. 130)", {
      x: 0.75, y: m2Y + 0.42, w: 8.4, h: 0.26,
      fontSize: 14, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    });
    s.addText("The German soldier -- young, inexperienced, still untouched by war", {
      x: 0.75, y: m2Y + 0.72, w: 8.4, h: 0.26,
      fontSize: 11, fontFace: FONT_B, color: C.MUTED, margin: 0,
    });

    // Personification
    const pY = m2Y + 1.17;
    addCard(s, 0.5, pY, 9, 1.05, { strip: C.ACCENT, fill: C.WHITE });
    s.addShape("roundRect", {
      x: 0.7, y: pY + 0.08, w: 1.8, h: 0.26, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    });
    s.addText("Personification", {
      x: 0.7, y: pY + 0.08, w: 1.8, h: 0.26,
      fontSize: 10, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });
    s.addText("\"...a smile lighting his worn face.\" (p. 133)", {
      x: 0.75, y: pY + 0.42, w: 8.4, h: 0.26,
      fontSize: 14, fontFace: FONT_H, color: C.CHARCOAL, italic: true, margin: 0,
    });
    s.addText("The smile is given the power to produce light -- warmth breaking through hardship", {
      x: 0.75, y: pY + 0.72, w: 8.4, h: 0.26,
      fontSize: 11, fontFace: FONT_B, color: C.MUTED, margin: 0,
    });

    // Contrast note at bottom
    const noteY = pY + 1.12;
    if (noteY + 0.24 <= SAFE_BOTTOM) {
      s.addText("Notice the contrast: the old Welshman marked by years vs the young German still smooth with youth.", {
        x: 0.75, y: noteY, w: 8.5, h: 0.24,
        fontSize: 11, fontFace: FONT_B, color: C.ACCENT, bold: true, italic: true, margin: 0,
      });
    }

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LITERARY);
  }

  // =========================================================================
  // SLIDE 11 -- Subordinating Conjunctions: I Do
  // =========================================================================
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Subordinating Conjunctions",
    "A subordinating conjunction joins a DEPENDENT clause to an INDEPENDENT clause:\n\nDEPENDENT clause: cannot stand alone\nINDEPENDENT clause: CAN stand alone\n\nSubordinating conjunctions:\nsince, before, even though, although, when, because, if, after, while, until",
    "Model:\n\nDEPENDENT: \"Since Joey is trapped within the wires and can't walk freely, ...\"\n\n+ INDEPENDENT: \"...the soldiers need to intervene to help get him out of no man's land.\"\n\nCheck: Can the independent clause stand alone? Yes -- it is a complete sentence.",
    NOTES_CONJ_IDO,
    FOOTER
  );

  // =========================================================================
  // SLIDES 12-13 -- CFU: Dependent or Independent? (withReveal)
  // =========================================================================
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Dependent or Independent?",
      "Thumbs Up / Down",
      "\"Before the Welsh soldier realises the German can speak English\"\n\nCan this clause stand alone as a complete sentence?\n\nThumbs UP = INDEPENDENT (yes, it can stand alone)\nThumbs DOWN = DEPENDENT (no, it needs more)",
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
      slide.addText("DEPENDENT", {
        x: 0.7, y: ansY + 0.10, w: 1.8, h: 0.28,
        fontSize: 12, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addText("It starts with \"before\" (a subordinating conjunction) and leaves you waiting. Before what happened? It DEPENDS on an independent clause to complete the thought.", {
        x: 2.7, y: ansY + 0.08, w: 6.6, h: 0.84,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // =========================================================================
  // SLIDES 14-15 -- We Do / You Do: Complete Sentences (withReveal)
  // =========================================================================
  withReveal(
    () => {
      const s = pres.addSlide();
      addTopBar(s, C.SECONDARY);
      addBadge(s, "We Do / You Do", { color: C.SECONDARY, w: 2.0 });
      addTitle(s, "Complete the Sentence");

      const stems = [
        { conj: "Since", text: "Since Joey is trapped within the wires and can't walk freely, ...", color: C.PRIMARY },
        { conj: "Before", text: "Before the Welsh soldier realises the German can speak English, ...", color: C.SECONDARY },
        { conj: "Even though", text: "Even though the two soldiers will return to fighting one another in a few hours, ...", color: C.ACCENT },
      ];

      stems.forEach((stem, i) => {
        const sy = CONTENT_TOP + i * 1.18;
        addCard(s, 0.5, sy, 9, 1.06, { strip: stem.color, fill: C.WHITE });
        s.addShape("roundRect", {
          x: 0.7, y: sy + 0.10, w: 1.8, h: 0.28, rectRadius: 0.08,
          fill: { color: stem.color },
        });
        s.addText(String(i + 1) + ". " + stem.conj, {
          x: 0.7, y: sy + 0.10, w: 1.8, h: 0.28,
          fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
        s.addText(stem.text, {
          x: 0.75, y: sy + 0.48, w: 8.4, h: 0.48,
          fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });

      addFooter(s, FOOTER);
      s.addNotes(NOTES_WEDO);
      return s;
    },
    (slide) => {
      const answers = [
        { y: CONTENT_TOP + 0.48, text: "...the soldiers need to intervene to help get him out of no man's land." },
        { y: CONTENT_TOP + 1.66, text: "...he thinks it was useless to come out to try and help Joey." },
        { y: CONTENT_TOP + 2.84, text: "...they were able to be kind to each other to help Joey." },
      ];
      answers.forEach((ans) => {
        slide.addShape("roundRect", {
          x: 0.7, y: ans.y, w: 8.5, h: 0.48, rectRadius: 0.06,
          fill: { color: C.BG_LIGHT },
        });
        slide.addText(ans.text, {
          x: 0.8, y: ans.y + 0.02, w: 8.2, h: 0.44,
          fontSize: 12, fontFace: FONT_H, color: C.CHARCOAL, italic: true, valign: "middle", margin: 0,
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
    "What moment in Chapter 16 surprised you the most, and why? Tell your partner.",
    [
      "I can identify key moments where the author shows character through dialogue and action",
      "I can explain how the author uses metaphor and personification to reveal character",
      "I can add an independent clause to a dependent clause to create a complex sentence",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 17 -- Resources
  // =========================================================================

  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: Subordinating Conjunctions Worksheet -------------------------
  const ws = createPdf({ title: WORKSHEET_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Subordinating Conjunctions -- Sentence Building", {
    color: C.PRIMARY,
    subtitle: "Chapter 16: The Coin Toss",
    lessonInfo: "War Horse | Lesson 21 | Week 5 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "Each sentence below starts with a dependent clause. Your job is to add an independent clause to complete the sentence. Remember: your independent clause must be able to stand alone as a complete sentence!", wsY, { color: C.PRIMARY });

  // Sentence 1
  wsY = addSectionHeading(ws, "1. Complete the sentence:", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Since Joey is trapped within the wires and can't walk freely,", wsY, { fontSize: 12, italic: true });
  wsY = addLinedArea(ws, wsY, 3, { lineSpacing: 26 });
  wsY += 10;

  // Sentence 2
  wsY = addSectionHeading(ws, "2. Complete the sentence:", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Before the Welsh soldier realises the German can speak English,", wsY, { fontSize: 12, italic: true });
  wsY = addLinedArea(ws, wsY, 3, { lineSpacing: 26 });
  wsY += 10;

  // Sentence 3
  wsY = addSectionHeading(ws, "3. Complete the sentence:", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "Even though the two soldiers will return to fighting one another in a few hours,", wsY, { fontSize: 12, italic: true });
  wsY = addLinedArea(ws, wsY, 3, { lineSpacing: 26 });
  wsY += 14;

  // Challenge
  wsY = addSectionHeading(ws, "CHALLENGE: Write Your Own!", wsY, { color: C.ALERT, fontSize: 14 });
  wsY = addTipBox(ws, "Write your own complex sentence about Chapter 16 using a subordinating conjunction from this list: although, while, until, after, when, if. Label the dependent clause (DC) and the independent clause (IC).", wsY, { color: C.ALERT });
  wsY = addLinedArea(ws, wsY, 3, { lineSpacing: 26 });

  addPdfFooter(ws, "War Horse | Lesson 21 | Subordinating Conjunctions Worksheet");

  // --- PDF 2: Answer Key ---------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Subordinating Conjunctions -- Answer Key", {
    color: C.ALERT,
    subtitle: "Teacher Reference -- Chapter 16: The Coin Toss",
    lessonInfo: "War Horse | Lesson 21 | Week 5 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Accept any completion where the independent clause is a genuine complete sentence that makes sense with the dependent clause. Model answers and alternatives are provided below.", akY, { color: C.ALERT });

  akY = addSectionHeading(ak, "1. Since Joey is trapped within the wires and can't walk freely,...", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Model: \"...the soldiers need to intervene to help get him out of no man's land.\"", akY);
  akY = addBodyText(ak, "Alternative: \"...both sides of soldiers try to encourage him to come towards them.\"", akY, { italic: true });
  akY = addBodyText(ak, "Alternative: \"...two brave soldiers climb out with wire cutters to free him.\"", akY, { italic: true });
  akY += 10;

  akY = addSectionHeading(ak, "2. Before the Welsh soldier realises the German can speak English,...", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "Model: \"...he thinks it was useless to come out to try and help Joey.\"", akY);
  akY = addBodyText(ak, "Alternative: \"...there is an awkward silence as the two men stare at each other.\"", akY, { italic: true });
  akY = addBodyText(ak, "Alternative: \"...he assumes they will not be able to communicate at all.\"", akY, { italic: true });
  akY += 10;

  akY = addSectionHeading(ak, "3. Even though the two soldiers will return to fighting one another in a few hours,...", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "Model: \"...they were able to be kind to each other to help Joey.\"", akY);
  akY = addBodyText(ak, "Alternative: \"...they share a moment of genuine humanity and respect.\"", akY, { italic: true });
  akY = addBodyText(ak, "Alternative: \"...they cooperate to free the horse and even have a laugh together.\"", akY, { italic: true });
  akY += 14;

  akY = addSectionHeading(ak, "Common Errors to Watch For", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- Completions that are fragments, not complete sentences (e.g., \"...helping him\" instead of \"...the soldiers help him\")", akY);
  akY = addBodyText(ak, "- Completions that start with another subordinating conjunction, creating a run-on dependent clause", akY);
  akY = addBodyText(ak, "- Completions that don't connect logically to the dependent clause", akY);
  akY = addBodyText(ak, "- Very short completions with no detail (e.g., \"...they helped\") -- push for specificity", akY);

  addPdfFooter(ak, "War Horse | Lesson 21 | Answer Key -- TEACHER COPY");

  // --- Write all files ----------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/WH5_Lesson21.pptx` }),
    writePdf(ws, WORKSHEET_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  const pptxPath = `${OUT_DIR}/WH5_Lesson21.pptx`;
  console.log("PPTX written to " + pptxPath);
  console.log(`Done: ${WORKSHEET_RESOURCE.name}.pdf`);
  console.log(`Done: ${ANSWER_KEY_RESOURCE.name}.pdf`);
}

build().catch(console.error);
