"use strict";

// War Horse Unit -- Lesson 22: Chapter 17 -- The Reunion
// Week 5, Session 2, Grade 5/6 Literacy
// Chapter 17 reading + because/but/so conjunctions sentence-level writing

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

const SESSION_NUMBER = 2;
const FOOTER = "War Horse | Lesson 22 of 25 | Week 5 | Year 5/6 Literacy";
const OUT_DIR = "output/WH5_Lesson22_The_Reunion";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const WORKSHEET_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Because But So Worksheet",
  "Student worksheet: complete three sentences using because, but, and so with a War Horse sentence stem."
);
const ANSWER_KEY_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Answer Key",
  "Teacher reference: model completions for because/but/so sentences with alternative valid answers."
);
const RESOURCE_ITEMS = [WORKSHEET_RESOURCE, ANSWER_KEY_RESOURCE];
const WORKSHEET_PDF_PATH = path.join(OUT_DIR, WORKSHEET_RESOURCE.fileName);
const ANSWER_KEY_PDF_PATH = path.join(OUT_DIR, ANSWER_KEY_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to Lesson 22. We are continuing with War Horse Chapter 17 today
- Last lesson we read Chapter 16 -- the coin toss between the Welsh and German soldiers. Today we find out what happens to Joey next
- Chapter 17 is a reunion chapter. Joey ends up at a veterinary hospital

DO:
- Display title slide as students settle
- Have copies of War Horse on desks, bookmarked at Chapter 17 (p. 135)

TEACHER NOTES:
Lesson 22 of 25, second session of Week 5. Chapter 17 is the emotional climax many students have been waiting for -- Albert finds Joey. The because/but/so work is simpler than yesterday's subordinating conjunctions, giving students a lighter sentence-level task alongside the emotional reading.

WATCH FOR:
- Students who need a brief recap of Chapter 16 -- the coin toss, Joey rescued from no man's land, taken to the British side
- Students who may be excited or emotional about the possibility of Albert and Joey reuniting

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Two strands again today: reading Chapter 17 and sentence-level writing with because, but, and so
- Read the success criteria. Some of you may have used because, but, and so before. Today we are being precise about what each one does

DO:
- Choral read the LI, then the SCs
- Do not over-explain the conjunctions yet -- the I Do covers this

TEACHER NOTES:
SC1 targets comprehension through dialogue analysis. SC2 addresses the repetition literary device at the chapter's climax. SC3 is the sentence-level writing target -- simpler than yesterday's subordinating conjunctions but still precise.

WATCH FOR:
- Students who assume because/but/so is "easy" -- the precision of choosing the RIGHT conjunction for the RIGHT purpose is the challenge
- Students who are eager to find out if Albert finds Joey -- channel that energy into careful reading

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_VOCAB_IMMACULATE = `SAY:
- Our first vocabulary word: immaculate. Read it with me: immaculate [students repeat]
- Immaculate means perfectly clean, neat, or tidy. Without a single mark or flaw
- In Chapter 17, Joey is covered in mud and blood from the battlefield. He is the OPPOSITE of immaculate right now. But by the end of the chapter, the soldiers clean him up
- Quick check: if your bedroom was immaculate, would there be clothes on the floor? [No -- immaculate means perfectly tidy]

DO:
- Display word, choral read, define, give example
- Connect to the chapter: the contrast between Joey's filthy state and what immaculate means creates dramatic tension
- Cold Call 2 students: describe something immaculate you have seen

TEACHER NOTES:
Pre-teaching "immaculate" before reading helps students appreciate the transformation Joey undergoes in this chapter and the care the soldiers take in cleaning him.

WATCH FOR:
- Students who confuse "immaculate" with just "clean" -- refine: "Immaculate is more than clean. It is PERFECTLY clean, without a single mark"
- Students who know it from "immaculate conception" -- acknowledge but steer to the general meaning

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_VOCAB_INQUISITIVE = `SAY:
- Second word: inquisitive. Read it with me: inquisitive [students repeat]
- Inquisitive means eager to learn or know things. An inquisitive person asks lots of questions and wants to find out more
- In Chapter 17, a character is inquisitive about Joey -- asking questions, looking closely, wanting to know more about this horse
- Ask: What is the OPPOSITE of inquisitive? [Uninterested, indifferent, apathetic]

DO:
- Display word, choral read, define, example
- Quick synonym check: curious, nosy, questioning
- Ask 1-2 students: when have YOU been inquisitive about something?

TEACHER NOTES:
"Inquisitive" is a strong Tier 2 word that students can use in their own writing. It connects to the character of David in Chapter 17, who becomes increasingly curious about whether this horse might be the Joey that Albert has described.

WATCH FOR:
- Students who only associate "inquisitive" with being nosy in a negative sense -- broaden: "Inquisitive is usually positive. It means you want to understand and learn"
- Students who confuse it with "inquiry" -- connect: same root, both about asking questions

[General: I Do -- Vocabulary | VTLM 2.0: Building Vocabulary]`;

const NOTES_READING_INTRO = `SAY:
- Chapter 17. Joey has been rescued from no man's land and taken to a veterinary hospital
- He is covered in mud and blood. He is assigned to a worker to be cleaned up so the vet can check him
- Today's reading is Student Read Aloud again. I will select readers. We have three pause points
- Focus: pay attention to what the characters SAY. The dialogue in this chapter reveals something very important
- Pages 135 to 145 -- find your page now

DO:
- Give students 30 seconds to find p. 135
- Select first reader -- choose a confident reader for the opening
- Plan reader rotations at natural breaks in dialogue

SENSITIVITY ADVISORY:
- What it is: Joey is injured and in a veterinary hospital during WWI. The chapter describes him being cleaned of mud and blood
- Framing language: "This chapter is about care and recognition -- people taking time to look after an animal properly"
- Watch for: Students who may be anxious about whether Joey will be okay -- reassure through tone
- Protocol: This is ultimately a hopeful chapter. Frame it as a reunion story

TEACHER NOTES:
Chapter 17 is the emotional payoff of the novel. Albert and Joey are reunited but the recognition is gradual -- Albert does not immediately know it is Joey because of the mud and blood. The dramatic irony (readers know before Albert does) creates real tension.

WATCH FOR:
- Readers who rush through the dialogue -- coach: "Slow down in the conversation parts. Let us hear what each character is saying"
- Students who figure out it is Albert before he is named -- great inference work

[General: Reading Launch | VTLM 2.0: Structured Reading Practice]`;

const NOTES_PAUSE1 = `SAY:
- Pause here. "...proper jumpy he was till I got to know him and he got to know me." Page 139
- What is going on? Who is speaking? [Albert is talking to David about a horse -- the horse he used to know. He is describing Joey without realising he is cleaning Joey right now]
- Ask: What does Albert mean by "proper jumpy"? What does this tell us about the relationship between Albert and Joey? [Joey was nervous at first, but they built trust over time. It shows the bond was not instant -- it was earned]

DO:
- Display the quote, read aloud
- Give 10 seconds of think time before taking responses
- Use Think-Pair-Share: 30 seconds think, 30 seconds pair, then share

CFU CHECKPOINT:
Technique: Think-Pair-Share

Script:
- "Think for 30 seconds: what is Albert revealing about his relationship with Joey?"
- "Share with your partner for 30 seconds"
- "Hands down -- I am selecting pairs. [Name], what did your pair discuss?"
- Scan for: students recognising dramatic irony -- Albert is describing Joey while unknowingly cleaning him

PROCEED (>=80%): Most pairs recognise Albert is talking about his horse. Continue reading.
PIVOT (<80%): Most likely issue -- students do not realise WHO is speaking or that he is talking about the SAME horse he is cleaning. Reteach: "Think about what Albert is describing. A jumpy horse that he got to know. Now look at the horse he is currently cleaning. What might Morpurgo be setting up?" Re-check: "Who do you think the horse is?"

TEACHER NOTES:
This is the first hint of dramatic irony in the chapter. Albert is reminiscing about Joey while literally cleaning Joey. The reader may figure it out before Albert does.

WATCH FOR:
- Students who recognise the dramatic irony immediately -- affirm and ask them not to spoil it for others
- Students who focus on "proper jumpy" as funny language -- acknowledge the colloquialism but push for the deeper meaning

[General: Pause Point 1 | VTLM 2.0: Teacher-Led Discussion]`;

const NOTES_PAUSE2 = `SAY:
- Stop here. "Old Sergeant Thunder won't mind, not if I've done all he told me, and I have." Page 141
- What have we learned from this conversation? [Albert cares deeply about doing a good job with the horses. He has been following orders carefully. He respects his sergeant but also has confidence in his own work]
- Ask: What kind of person is Albert? What does this dialogue reveal about his character? [He is conscientious, caring, thorough, and has a strong sense of responsibility towards the animals]

DO:
- Display the quote
- Use Cold Call for responses
- Push for character analysis -- this is a direct SC1 opportunity

TEACHER NOTES:
This quote shows Albert's character: dutiful, responsible, confident in his work. The casual way he talks about "Old Sergeant Thunder" also shows he has earned respect within the unit. This is character revealed through dialogue (SC1).

WATCH FOR:
- Students who focus only on surface meaning -- push: "What does it tell us about Albert as a PERSON, not just what he is doing?"
- Students who notice Albert's affection for the horses in general -- excellent observation: he treats all horses well, not just Joey

[General: Pause Point 2 | VTLM 2.0: Higher-Order Questioning]`;

const NOTES_PAUSE3 = `SAY:
- Final pause. "Not often," Albert said. "Not often, and not this time." Page 145
- What just happened? [Albert has realised the horse he has been cleaning IS Joey. David pointed out that the horse matches Albert's description exactly. Albert looks and confirms it]
- Ask: Read those words again. "Not often... Not often, and not this time." Why does Morpurgo repeat "not often"? [It builds emotion. The repetition makes the moment land harder. Albert is saying: I am not often wrong about horses -- and I am not wrong now. This IS Joey]

DO:
- Display the quote, read with emotional weight
- Let the moment land -- this is the climax of the whole chapter
- Cold Call 2-3 students for their interpretation of the repetition

TEACHER NOTES:
This is the emotional climax. The repetition of "not often" creates a rhythmic, almost breathless quality. Albert's certainty builds through the repetition. This connects directly to the literary device slide.

WATCH FOR:
- Students who are emotionally affected -- give space for that. This reunion has been building for the entire novel
- Students who can identify WHAT the repetition does (builds emotion, creates certainty) vs just noting it exists

[General: Pause Point 3 | VTLM 2.0: Deep Comprehension]`;

const NOTES_PAUSE3_REVEAL = `SAY:
- Morpurgo uses repetition here for emotional impact
- The full exchange: "Not often wrong, am I?" "Not often," Albert said. "Not often, and not this time."
- The phrase "not often" appears three times in quick succession. Each repetition adds weight and certainty
- This is the moment everything in the novel has been building towards. After years apart, through a whole war, Albert and Joey have found each other again
- And notice -- it is the ACT OF CARING (cleaning the horse properly) that leads to the reunion. Albert's thoroughness reveals Joey beneath the mud

DO:
- Reveal the analysis card
- Read the full exchange aloud with appropriate emotion
- Draw the connection: the literary device (repetition) serves the theme (reunion through care)
- Transition: "Now let's move to our sentence-level writing"

[General: Pause Point 3 Reveal | VTLM 2.0: Deep Comprehension]`;

const NOTES_CONJ_IDO = `SAY:
- Now sentence-level writing. Today we are revising three important conjunctions: because, but, and so
- These are coordinating and subordinating conjunctions that show DIFFERENT relationships between ideas
- BECAUSE tells us WHY something is true -- it gives a reason
- BUT tells us about a CHANGE IN DIRECTION -- something unexpected or contrasting
- SO tells us WHAT HAPPENS AS A RESULT -- it shows a consequence
- Watch me. Our sentence stem is: "Albert doesn't think the horse could be Joey..."
- With BECAUSE: "Albert doesn't think the horse could be Joey because he couldn't recognise him at all underneath all of the mud and blood." -- BECAUSE gives the reason
- With BUT: "Albert doesn't think the horse could be Joey, but he does his best to take care of the horse anyway." -- BUT shows the contrast
- With SO: "Albert doesn't think the horse could be Joey, so he misses the initial signs that Joey gives him to let him know that it is him." -- SO shows the consequence

DO:
- Display the three conjunctions with their purposes clearly
- Model each one with the sentence stem
- Think aloud: "I choose the conjunction based on what relationship I want to show. Am I explaining why? Use because. Am I showing a contrast? Use but. Am I showing a result? Use so"
- Point to the conjunction in each model sentence

TEACHER NOTES:
The think-aloud about choosing the RIGHT conjunction for the RIGHT relationship is the key metacognitive move. Students often default to "because" for everything -- push them to see that but and so create different meanings.

WATCH FOR:
- Students who look like they already know this -- the precision of choosing correctly is the real challenge, even if the words are familiar
- Students who confuse "but" and "so" -- "but" changes direction, "so" continues in the same direction but shows consequence

[General: I Do -- Modelling | VTLM 2.0: Explicit Teaching / Modelling]`;

const NOTES_CFU = `SAY:
- Quick check. I will read a sentence. You show me which conjunction fits best
- "Albert cleaned the horse carefully ___ he wanted to do a thorough job for the vet"
- Think: does the missing word explain WHY, show a CONTRAST, or show a RESULT?
- Hold up one finger for BECAUSE, two for BUT, three for SO
- Three, two, one -- show me! [Scan]

DO:
- Use finger voting
- Scan for: one finger (BECAUSE -- it explains WHY Albert cleaned carefully)

CFU CHECKPOINT:
Technique: Finger Voting (1 = because, 2 = but, 3 = so)

Script:
- Read the sentence with a clear pause at the gap
- Count down, scan
- Look for: mostly one finger (because)

PROCEED (>=80%): Most show one finger. Move to You Do.
PIVOT (<80%): Most likely issue -- students choose "so" because they see a cause-and-effect relationship. Reteach: "Read it again. The second part -- 'he wanted to do a thorough job' -- is that a RESULT of cleaning, or the REASON for cleaning? He cleaned carefully BECAUSE he wanted to do a good job. The reason comes after 'because'. If it said 'so the vet would be pleased', THEN it would be 'so' -- because 'the vet would be pleased' is the RESULT." Re-check with: "The horse was very dirty ___ the soldiers spent extra time cleaning him." [so -- result]

TEACHER NOTES:
Finger voting is quick and decisive. The because/so confusion is the most common error. The re-check sentence uses "so" to calibrate.

WATCH FOR:
- Students who guess without processing -- push: "Read the second part. Is it a reason, a contrast, or a result?"
- Students who choose "but" -- this is less common but indicates they do not see the explanatory relationship

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_CFU_REVEAL = `SAY:
- One finger -- BECAUSE. "Albert cleaned the horse carefully BECAUSE he wanted to do a thorough job for the vet"
- The second part explains WHY he cleaned carefully. It is a reason, not a result or a contrast
- Remember: BECAUSE = why. BUT = change of direction. SO = what happens next as a result

DO:
- Reveal the answer
- Highlight BECAUSE as giving the reason
- Quick contrast: "If I changed it to 'Albert cleaned the horse carefully, SO the vet was impressed' -- now 'so' works because 'the vet was impressed' is a RESULT"
- Transition to You Do

[General: CFU Reveal | VTLM 2.0: Formative Assessment]`;

const NOTES_YOUDO = `SAY:
- Your turn. On the worksheet, you have one sentence stem: "Albert doesn't think the horse could be Joey..."
- You need to complete this stem THREE times -- once with because, once with but, once with so
- Each completion must make sense with the conjunction. Because gives a reason. But shows a contrast. So shows a result
- You have 5 minutes. Work independently

DO:
- Distribute the Session 2 Because But So Worksheet
- Circulate -- check that students are matching the conjunction to the right type of relationship
- Common error: using "because" phrasing after "so" or vice versa

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide a word bank with key ideas from the chapter (mud, blood, recognise, signs, clean, care) and show the I Do model sentences on the board for reference
- Extra Notes: Students can use the teacher's models as a starting template and modify them

EXTENDING PROMPT:
- Task: Write a FOURTH sentence about Chapter 17 using a conjunction of your choice. Then write a sentence explaining WHY you chose that conjunction -- what relationship does it show?

TEACHER NOTES:
The single stem completed three ways forces students to think about the CONJUNCTION rather than the content. They must vary their thinking while keeping the same starting point.

WATCH FOR:
- Students who write the same type of completion for all three -- check: "Read your 'but' sentence. Does it actually show a contrast?"
- Students who write very short completions -- push for detail connected to the chapter
- Students who confuse "so" and "because" -- the most common mix-up

[General: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: key moments showing character through dialogue -- thumbs? [scan]
- SC2: how the author uses repetition for emotional impact -- thumbs? [scan]
- SC3: completing sentences with because, but, and so -- thumbs? [scan]
- Turn and talk: What is the most important moment in Chapter 17, and why?

DO:
- Run through each SC with thumbs check
- The turn-and-talk invites reflection on the emotional climax
- Preview: "Next lesson we continue with the final chapters of War Horse"

TEACHER NOTES:
The closing reconnects to the reading experience. Most students will identify the reunion as the most important moment, but some may highlight Albert's description of Joey or David's realisation -- all are valid.

WATCH FOR:
- Students "thumbs down" on SC3 -- check if it is a genuine confusion between the conjunctions
- Students who want to discuss what happens next in the novel -- encourage curiosity

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two printable resources today
- The ${WORKSHEET_RESOURCE.name} is for the You Do activity -- one stem, three conjunctions
- The ${ANSWER_KEY_RESOURCE.name} is for teacher reference with model answers

DO:
- Print the worksheet before the lesson (one per student)
- Print the answer key (teacher copy only)
- Click any resource card to open the PDF

TEACHER NOTES:
The worksheet uses the same sentence stem three times with space for each conjunction. The answer key includes alternative valid completions.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "War Horse - Lesson 22 - Chapter 17: The Reunion";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "War Horse",
    "Chapter 17 -- The Reunion",
    "Lesson 22  |  Week 5  |  Year 5/6 Literacy",
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
      "We are learning to analyse how an author uses dialogue to reveal character, and to build sentences using because, but, and so",
    ],
    [
      "I can identify key moments where the author uses dialogue to show what characters think and feel",
      "I can explain how the author uses repetition to create emotional impact",
      "I can complete sentences using because, but, and so to show different relationships between ideas",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Vocabulary: immaculate
  // =========================================================================
  vocabSlide(
    pres,
    "immaculate",
    "adjective",
    "Perfectly clean, neat, or tidy. Without a single mark or flaw. An immaculate room has everything in its place and not a speck of dust.",
    "After hours of careful cleaning, the soldiers had transformed the filthy horse into an almost immaculate state.",
    NOTES_VOCAB_IMMACULATE,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Vocabulary: inquisitive
  // =========================================================================
  vocabSlide(
    pres,
    "inquisitive",
    "adjective",
    "Eager to learn or know things. An inquisitive person asks lots of questions and wants to find out more about the world around them.",
    "David became increasingly inquisitive as the features he uncovered matched Albert's description of his horse.",
    NOTES_VOCAB_INQUISITIVE,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- Reading Introduction
  // =========================================================================
  contentSlide(
    pres,
    "Read Aloud",
    C.PRIMARY,
    "Chapter 17 -- Pages 135-145",
    [
      "Reading Mode: Student Read Aloud",
      "Joey is taken to a veterinary hospital, covered in mud and blood",
      "A worker is assigned to clean him up so the vet can assess him",
      "Focus: listen to the dialogue carefully -- what do the characters reveal about themselves and about Joey?",
    ],
    NOTES_READING_INTRO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 6 -- Pause Point 1
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 1", "Chapter 17 -- p. 139",
    "...proper jumpy he was till I got to know him and he got to know me.",
    "p. 139",
    "What's going on? Who is speaking, and what are they describing?",
    NOTES_PAUSE1, FOOTER
  );

  // =========================================================================
  // SLIDE 7 -- Pause Point 2
  // =========================================================================
  quoteSlide(
    pres, "Pause Point 2", "Chapter 17 -- p. 141",
    "Old Sergeant Thunder won't mind, not if I've done all he told me, and I have.",
    "p. 141",
    "What have we learned from this conversation? What kind of person is Albert?",
    NOTES_PAUSE2, FOOTER
  );

  // =========================================================================
  // SLIDES 8-9 -- Pause Point 3 (withReveal) + Repetition Literary Device
  // =========================================================================
  withReveal(
    () => quoteSlide(
      pres, "Pause Point 3", "Chapter 17 -- p. 145",
      "\"Not often,\" Albert said. \"Not often, and not this time.\"",
      "p. 145",
      "What just happened? Why does Morpurgo repeat \"not often\"?",
      NOTES_PAUSE3, FOOTER
    ),
    (slide) => {
      const ansY = 3.50;
      slide.addShape("roundRect", {
        x: 0.5, y: ansY, w: 9, h: 1.50, rectRadius: 0.10,
        fill: { color: C.SUCCESS },
      });

      // Label pill
      slide.addShape("roundRect", {
        x: 0.7, y: ansY + 0.10, w: 1.4, h: 0.28, rectRadius: 0.08,
        fill: { color: C.WHITE },
      });
      slide.addText("Repetition", {
        x: 0.7, y: ansY + 0.10, w: 1.4, h: 0.28,
        fontSize: 11, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Full exchange
      slide.addText("\"Not often wrong, am I?\" ... \"Not often,\" Albert said. \"Not often, and not this time.\"", {
        x: 2.3, y: ansY + 0.06, w: 7.0, h: 0.36,
        fontSize: 12, fontFace: FONT_H, color: C.WHITE, italic: true, margin: 0,
      });

      // Analysis
      slide.addText("The phrase \"not often\" appears three times, building certainty and emotional weight. Each repetition deepens Albert's conviction -- this IS Joey. Morpurgo uses repetition to make the reunion moment land with full force. The act of caring (cleaning the horse) leads to recognition.", {
        x: 0.75, y: ansY + 0.48, w: 8.5, h: 0.92,
        fontSize: 12.5, fontFace: FONT_B, color: C.WHITE, valign: "top", margin: 0,
      });
      slide.addNotes(NOTES_PAUSE3_REVEAL);
    }
  );

  // =========================================================================
  // SLIDE 10 -- Because / But / So: I Do
  // =========================================================================
  modellingSlide(
    pres,
    "I Do -- Watch Me",
    "Because, But, and So",
    "Three conjunctions, three different jobs:\n\nBECAUSE - tells us WHY (gives a reason)\nBUT - shows a CHANGE IN DIRECTION (contrast)\nSO - shows WHAT HAPPENS AS A RESULT (consequence)\n\nSentence stem:\n\"Albert doesn't think the horse could be Joey...\"",
    "BECAUSE: \"...because he couldn't recognise him at all underneath all of the mud and blood.\"\n(Reason: why Albert thinks this)\n\nBUT: \"...but he does his best to take care of the horse anyway.\"\n(Contrast: he still helps despite not recognising Joey)\n\nSO: \"...so he misses the initial signs that Joey gives him.\"\n(Result: what happens because of Albert's belief)",
    NOTES_CONJ_IDO,
    FOOTER
  );

  // =========================================================================
  // SLIDES 11-12 -- CFU: Which Conjunction? (withReveal)
  // =========================================================================
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Which Conjunction?",
      "Finger Voting",
      "\"Albert cleaned the horse carefully ___ he wanted to do a thorough job for the vet.\"\n\nWhich conjunction fits best?\n\n1 finger = BECAUSE (reason)\n2 fingers = BUT (contrast)\n3 fingers = SO (result)",
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
      slide.addText("BECAUSE", {
        x: 0.7, y: ansY + 0.10, w: 1.8, h: 0.28,
        fontSize: 12, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      slide.addText("The second part explains WHY Albert cleaned carefully -- \"he wanted to do a thorough job\" is the REASON, not a result or a contrast. BECAUSE introduces reasons.", {
        x: 2.7, y: ansY + 0.08, w: 6.6, h: 0.84,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU_REVEAL);
    }
  );

  // =========================================================================
  // SLIDE 13 -- You Do: Complete with Because / But / So
  // =========================================================================
  taskSlide(
    pres,
    "You Do",
    "Complete the Sentence Three Ways",
    [
      { label: "BECAUSE", instruction: "Albert doesn't think the horse could be Joey because..." },
      { label: "BUT", instruction: "Albert doesn't think the horse could be Joey, but..." },
      { label: "SO", instruction: "Albert doesn't think the horse could be Joey, so..." },
    ],
    NOTES_YOUDO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 14 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    "What is the most important moment in Chapter 17, and why? Tell your partner.",
    [
      "I can identify key moments where the author uses dialogue to show what characters think and feel",
      "I can explain how the author uses repetition to create emotional impact",
      "I can complete sentences using because, but, and so to show different relationships between ideas",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 15 -- Resources
  // =========================================================================

  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: Because But So Worksheet -------------------------------------
  const ws = createPdf({ title: WORKSHEET_RESOURCE.name });
  let wsY = addPdfHeader(ws, "Because, But, and So -- Sentence Building", {
    color: C.PRIMARY,
    subtitle: "Chapter 17: The Reunion",
    lessonInfo: "War Horse | Lesson 22 | Week 5 | Year 5/6 Literacy",
    showNameDate: true,
  });

  wsY = addTipBox(ws, "Complete the sentence stem THREE times using a different conjunction each time. Each conjunction shows a different relationship:\n- BECAUSE = tells us WHY (gives a reason)\n- BUT = shows a CHANGE IN DIRECTION (contrast)\n- SO = shows WHAT HAPPENS AS A RESULT (consequence)", wsY, { color: C.PRIMARY });

  // Sentence stem display
  wsY = addSectionHeading(ws, "Sentence Stem:", wsY, { color: C.CHARCOAL });
  wsY = addBodyText(ws, "\"Albert doesn't think the horse could be Joey...\"", wsY, { fontSize: 13, bold: true });
  wsY += 8;

  // BECAUSE
  wsY = addSectionHeading(ws, "1. Complete with BECAUSE:", wsY, { color: C.PRIMARY });
  wsY = addBodyText(ws, "Albert doesn't think the horse could be Joey because", wsY, { fontSize: 11, italic: true });
  wsY = addLinedArea(ws, wsY, 3, { lineSpacing: 26 });
  wsY += 10;

  // BUT
  wsY = addSectionHeading(ws, "2. Complete with BUT:", wsY, { color: C.SECONDARY });
  wsY = addBodyText(ws, "Albert doesn't think the horse could be Joey, but", wsY, { fontSize: 11, italic: true });
  wsY = addLinedArea(ws, wsY, 3, { lineSpacing: 26 });
  wsY += 10;

  // SO
  wsY = addSectionHeading(ws, "3. Complete with SO:", wsY, { color: C.ACCENT });
  wsY = addBodyText(ws, "Albert doesn't think the horse could be Joey, so", wsY, { fontSize: 11, italic: true });
  wsY = addLinedArea(ws, wsY, 3, { lineSpacing: 26 });
  wsY += 14;

  // Challenge
  wsY = addSectionHeading(ws, "CHALLENGE:", wsY, { color: C.ALERT, fontSize: 14 });
  wsY = addTipBox(ws, "Write your own sentence about Chapter 17 using a conjunction of your choice (because, but, or so). Then explain WHY you chose that conjunction -- what relationship does it show?", wsY, { color: C.ALERT });
  wsY = addLinedArea(ws, wsY, 3, { lineSpacing: 26 });

  addPdfFooter(ws, "War Horse | Lesson 22 | Because But So Worksheet");

  // --- PDF 2: Answer Key ---------------------------------------------------
  const ak = createPdf({ title: ANSWER_KEY_RESOURCE.name });
  let akY = addPdfHeader(ak, "Because, But, and So -- Answer Key", {
    color: C.ALERT,
    subtitle: "Teacher Reference -- Chapter 17: The Reunion",
    lessonInfo: "War Horse | Lesson 22 | Week 5 | Year 5/6 Literacy",
    showNameDate: false,
  });

  akY = addTipBox(ak, "Accept any completion that correctly uses the conjunction to show the right type of relationship (reason, contrast, or result). Model answers and alternatives below.", akY, { color: C.ALERT });

  akY = addSectionHeading(ak, "1. BECAUSE (reason -- tells us WHY)", akY, { color: C.PRIMARY });
  akY = addBodyText(ak, "Model: \"...because he couldn't recognise him at all underneath all of the mud and blood.\"", akY);
  akY = addBodyText(ak, "Alternative: \"...because the horse looks completely different from how he remembers Joey.\"", akY, { italic: true });
  akY = addBodyText(ak, "Alternative: \"...because it has been so long since he last saw him and the horse is in such bad condition.\"", akY, { italic: true });
  akY += 10;

  akY = addSectionHeading(ak, "2. BUT (contrast -- change in direction)", akY, { color: C.SECONDARY });
  akY = addBodyText(ak, "Model: \"...but he does his best to take care of the horse anyway.\"", akY);
  akY = addBodyText(ak, "Alternative: \"...but something about the horse feels strangely familiar to him.\"", akY, { italic: true });
  akY = addBodyText(ak, "Alternative: \"...but he still talks gently to the horse as he cleans him.\"", akY, { italic: true });
  akY += 10;

  akY = addSectionHeading(ak, "3. SO (result -- what happens as a consequence)", akY, { color: C.ACCENT });
  akY = addBodyText(ak, "Model: \"...so he misses the initial signs that Joey gives him to let him know that it is him.\"", akY);
  akY = addBodyText(ak, "Alternative: \"...so he carries on cleaning without paying special attention to the horse.\"", akY, { italic: true });
  akY = addBodyText(ak, "Alternative: \"...so David has to point out that the features match Albert's description.\"", akY, { italic: true });
  akY += 14;

  akY = addSectionHeading(ak, "Common Errors to Watch For", akY, { color: C.ALERT });
  akY = addBodyText(ak, "- Confusing BECAUSE and SO: \"because\" gives a reason (looking backward at cause), \"so\" shows a result (looking forward at consequence)", akY);
  akY = addBodyText(ak, "- Using \"but\" without a genuine contrast -- the completion must change direction from the stem", akY);
  akY = addBodyText(ak, "- Fragments instead of complete clauses after the conjunction", akY);
  akY = addBodyText(ak, "- Very short or vague completions -- push for detail connected to Chapter 17", akY);

  addPdfFooter(ak, "War Horse | Lesson 22 | Answer Key -- TEACHER COPY");

  // --- Write all files ----------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/WH5_Lesson22.pptx` }),
    writePdf(ws, WORKSHEET_PDF_PATH),
    writePdf(ak, ANSWER_KEY_PDF_PATH),
  ]);

  const pptxPath = `${OUT_DIR}/WH5_Lesson22.pptx`;
  console.log("PPTX written to " + pptxPath);
  console.log(`Done: ${WORKSHEET_RESOURCE.name}.pdf`);
  console.log(`Done: ${ANSWER_KEY_RESOURCE.name}.pdf`);
}

build().catch(console.error);
