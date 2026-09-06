"use strict";

// War Horse Unit -- Lessons 24-25: Plan and Write a Body Paragraph
// Week 5, Session 4, Grade 5/6 Literacy
// WW1 supplementary reading + SPO + modality + nominalisations + paragraph writing

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
} = require("../themes/wh5_helpers");

const {
  createPdf, writePdf, addPdfHeader, addSectionHeading,
  addBodyText, addTipBox, addPdfFooter, addLinedArea,
  addResourceSlide, getSessionResourceFolder, makeSessionResource,
  addWriteLine, addStepInstructions,
} = require("../themes/pdf_helpers");

const SESSION_NUMBER = 4;
const FOOTER = "War Horse | Lessons 24-25 | Week 5 | Year 5/6 Literacy";
const OUT_DIR = "output/WH5_Lesson24_Body_Paragraph";
const RES_DIR = path.join(OUT_DIR, getSessionResourceFolder(SESSION_NUMBER));

const SPO_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "SPO Template",
  "Single paragraph outline template for students to plan their body paragraph."
);
const LANG_RESOURCE = makeSessionResource(
  SESSION_NUMBER,
  "Persuasive Language Reference",
  "Reference sheet with high modality language and nominalisation examples for persuasive writing."
);
const RESOURCE_ITEMS = [SPO_RESOURCE, LANG_RESOURCE];
const SPO_PDF_PATH = path.join(OUT_DIR, SPO_RESOURCE.fileName);
const LANG_PDF_PATH = path.join(OUT_DIR, LANG_RESOURCE.fileName);
fs.mkdirSync(RES_DIR, { recursive: true });

// ---------------------------------------------------------------------------
// Teacher Notes
// ---------------------------------------------------------------------------

const NOTES_TITLE = `SAY:
- Welcome to today's lesson. We are combining our reading and writing work
- First we will read more about World War 1 from the National Geographic Kids article
- Then we will plan and write a body paragraph for our persuasive text

DO:
- Display title slide as students settle
- Have the NatGeo Kids article ready on devices or printed copies
- Have copies of the SPO template ready to distribute later

TEACHER NOTES:
This session combines Lessons 24 and 25 of the unit. Students read WW1 source material then plan and write a body paragraph using the SPO template. The writing focus is on modality and nominalisations as persuasive tools.

WATCH FOR:
- Students who need a recap of the persuasive text macrostructure from previous sessions
- Students who have not yet finished their introductory paragraph from the prior session

[General: Title | VTLM 2.0: Establishing Purpose and Relevance]`;

const NOTES_LI_SC = `SAY:
- Read the learning intention from the slide
- Today we are planning AND writing a body paragraph for our persuasive text
- Read the success criteria. SC1 is about knowing what a body paragraph does. SC2 is about using the SPO template to plan it. SC3 is about using persuasive language -- modality and nominalisations

DO:
- Choral read the LI, then the SCs
- Do not explain modality or nominalisations yet -- the I Do covers both

TEACHER NOTES:
SC1 is the foundation -- every student should be able to name the purpose and features of a body paragraph. SC2 is the core target -- planning with the SPO. SC3 is the depth -- incorporating persuasive language techniques.

WATCH FOR:
- Students who look uncertain about "modality" or "nominalisations" -- reassure: "We will learn these together today"
- Students who are confident from the introductory paragraph work -- they are ready

[General: Learning Intention | VTLM 2.0: Clear Learning Intention]`;

const NOTES_READING = `SAY:
- Before we write, we need to build our knowledge. We are reading from the National Geographic Kids article about World War 1
- Today we are reading these sections: Why did it start, How long did it last, Life in the trenches, A new type of war, Animals during World War 1, What was life like in Britain, Women during World War 1, Children during World War 1, and Armistice Day
- As you read, take notes. You are looking for evidence and information you could use in your persuasive text
- Reading mode: student read aloud. I will select readers

DO:
- Distribute printed articles or direct students to the URL on devices
- Select readers and rotate every 1-2 sections
- Allow students to jot notes as they read -- this is active reading
- Monitor note-taking quality -- are students capturing key facts, not copying whole sentences?

TEACHER NOTES:
The reading covers two sets of article sections (from Lessons 24 and 25). Students are building their evidence base for persuasive writing. The note-taking is formative -- it feeds directly into their SPO planning later. Incidental vocabulary across both readings: outbreak, civilisation, Europe, allies, central powers, trenches, Western Front, no man's land, artillery, home front, zeppelin airships, raids, shelling, battlecruisers, conscientious objection, conscription, armistice day, commemoration.

WATCH FOR:
- Students who copy whole sentences instead of taking notes -- coach: "Write the key idea in your own words, not the author's exact sentence"
- Students who are not taking notes at all -- prompt: "What is one thing you have learned so far? Write that down"

SOURCES:
World War 1 Facts for Kids - National Geographic Kids Australia

[General: Reading | VTLM 2.0: Structured Reading Practice]`;

const NOTES_MACROSTRUCTURE = `SAY:
- Quick revision. A persuasive text has three main parts. What are they? [Introduction, body paragraphs, conclusion]
- We have already written our introduction. Today we are focusing on the body paragraphs
- A body paragraph has ONE job: present one key argument that supports your position
- The structure is simple: Topic Sentence, Supporting Details, Concluding Sentence. We call this the SPO -- single paragraph outline

DO:
- Display the SPO structure on the slide
- Point to each element as you explain it
- Emphasise: ONE argument per paragraph, not multiple

CFU CHECKPOINT:
Technique: Choral Response

Script:
- "What does the topic sentence do?" [States the main argument of the paragraph]
- "What do the supporting details do?" [Provide evidence, examples, and explanation]
- "What does the concluding sentence do?" [Wraps up the argument and links back to the position]
- Scan for: confident choral responses on all three

PROCEED (>=80%): Most students respond accurately. Move to modality.
PIVOT (<80%): Most likely issue -- students confuse body paragraph features with introduction features. Reteach: "The introduction tells the reader your WHOLE position. A body paragraph zooms in on ONE argument. Think of it like a camera -- the introduction is the wide shot, the body paragraph is the close-up on one specific point." Re-check: "If my persuasive text argues that animals were important in WW1, what might ONE body paragraph focus on? [One specific role -- like transport, or communication, or companionship]"

TEACHER NOTES:
This is revision from L24. The SPO template is a planning tool that students will use throughout the unit. Keep the revision brisk -- students encountered persuasive macrostructure in the previous sessions.

WATCH FOR:
- Students who think they need to include ALL their arguments in one paragraph -- reinforce: one argument per paragraph
- Students who are unclear on the difference between a topic sentence and a thesis statement

[General: I Do -- Revision | VTLM 2.0: Activating Prior Knowledge]`;

const NOTES_MODALITY = `SAY:
- Now, persuasive language. The first tool is MODALITY -- how strongly you say something
- High modality means you are very certain and definite: must, always, certainly, undoubtedly, every, without question
- Low modality is softer: might, sometimes, perhaps, could, possibly
- In a persuasive text, we mostly use HIGH modality to sound confident and convincing
- Watch: "Animals helped in the war." That is neutral. Now with high modality: "Animals were undoubtedly essential to the war effort." -- much more persuasive
- But sometimes we CONTROL modality for precision. "Most soldiers relied on horses" is more credible than "Every single soldier relied on horses" -- because the second one is not actually true

DO:
- Display the modality examples on the slide
- Read both versions aloud so students hear the difference
- Emphasise: high modality for confidence, but not so high that you lose credibility

TEACHER NOTES:
Modality is a key persuasive tool from the curriculum (VC2E6LA08). The distinction between HIGH modality for persuasion and CONTROLLED modality for precision is important -- students who only use "always" and "must" can sound aggressive rather than persuasive. The sweet spot is confident but credible.

WATCH FOR:
- Students who think higher modality is always better -- push: "If I say 'Every horse that ever lived was used in the war,' is that convincing or just wrong?"
- Students who default to low modality because they are uncertain writers -- encourage: "In a persuasive text, your job is to sound certain"

[General: I Do -- Modality | VTLM 2.0: Explicit Teaching]`;

const NOTES_SPO_MODEL = `SAY:
- Watch me plan a body paragraph using the SPO template. My argument is: horses were useful for transport during the war
- Topic sentence: I need a strong opening that states my argument clearly. "Horses were undoubtedly one of the most essential forms of transport during World War 1."
- Notice the high modality: "undoubtedly" and "most essential"
- Supporting details: I need evidence. SD1: Horses carried soldiers, weapons, and supplies to the front lines. SD2: Without horses, armies could not move artillery or ammunition across muddy, damaged terrain where vehicles could not go. SD3: Horses also pulled ambulance carts to rescue wounded soldiers
- Concluding sentence: wraps up and connects back to my position. "The critical role horses played in wartime transport clearly demonstrates their importance to the war effort."

DO:
- Display the SPO template being filled in step by step
- Think aloud at each decision point -- why this evidence, why this order
- Point out the modality words as you use them

TEACHER NOTES:
This is the worked example from L24. The teacher models the PLANNING stage, not the writing yet. Students see how to take notes from reading and organise them into the SPO structure. The transport argument connects directly to the "Animals during WW1" section they read earlier.

WATCH FOR:
- Students who want to start writing their paragraph immediately -- redirect: "Plan first, write second. The SPO is your blueprint"
- Students who notice the modality language in the model -- affirm and ask them to identify which words are high modality

[General: I Do -- Worked Example | VTLM 2.0: Modelling]`;

const NOTES_WRITING_MODEL = `SAY:
- Now watch me turn my SPO into a written paragraph. I follow my plan -- the SPO tells me what to write
- Read from slide: "Horses were undoubtedly one of the most essential forms of transport during World War 1. They carried soldiers, weapons, and vital supplies to the front lines, often through terrain that no vehicle could cross. Without horses, armies would have been unable to move artillery or ammunition across the muddy, shell-damaged battlefields of the Western Front. Horses also pulled ambulance carts, rescuing wounded soldiers who would otherwise have been left behind. The critical role horses played in wartime transport clearly demonstrates their importance to the war effort."
- Notice: my SPO had three supporting details, and each one became a sentence. The concluding sentence links back to my position
- I also added a nominalisation: "importance." That used to be the adjective "important" -- I turned it into a noun. Nominalisations make your writing sound more formal and authoritative

DO:
- Display the complete paragraph alongside the SPO
- Draw lines connecting SPO elements to paragraph sentences
- Highlight the nominalisation: "importance" from "important"
- Briefly explain: nominalisations turn verbs or adjectives into nouns. They make writing sound more formal

MISCONCEPTIONS:
- Misconception: Every supporting detail needs to be exactly one sentence
  Why: Students see three SDs become three sentences and assume a 1:1 rule
  Impact: Students write unnaturally short or long sentences to match their SD count
  Quick correction: "A supporting detail can become one, two, or even three sentences. The SPO gives you the IDEAS -- you decide how many sentences each idea needs"

TEACHER NOTES:
This slide models the WRITING stage. The paragraph demonstrates the SPO-to-paragraph workflow. The nominalisation introduction is brief here -- the concept is new and will be practised more in subsequent sessions. Keep it light: name it, show one example, move on.

WATCH FOR:
- Students who think the paragraph must match the SPO word-for-word -- "The SPO is a plan, not a script. You expand and connect the ideas when you write"
- Students who are intrigued by nominalisations -- they will get more practice with these

[General: I Do -- Writing Model | VTLM 2.0: Modelling]`;

const NOTES_CFU = `SAY:
- Quick check. I am going to show you two sentences. Tell me which one uses higher modality and would be more persuasive
- Sentence A: "Horses sometimes helped with transport during the war."
- Sentence B: "Horses were essential to the transportation of troops and supplies during the war."
- On your whiteboards: write A or B, and write the word that makes it high modality

DO:
- Display both sentences
- Allow 30 seconds for whiteboard responses
- Signal: boards up

CFU CHECKPOINT:
Technique: Show Me Boards

Script:
- "Write: A or B, plus the high modality word. Boards up in 30 seconds."
- Scan for: B + "essential" (or "were essential")
- Also accept if students note the nominalisation "transportation" in B -- that is a bonus observation

PROCEED (>=80%): Most show B with "essential." Move to We Do.
PIVOT (<80%): Most likely issue -- students choose A because it sounds "nicer" or more polite. Reteach: "In everyday conversation, softening language is polite. But in persuasive writing, your job is to CONVINCE. 'Sometimes helped' sounds uncertain. 'Were essential' sounds like a fact. Which one would convince a reader?" Re-check: "Which sounds more persuasive -- 'Animals could be useful' or 'Animals were vital'? Show me on your boards."

TEACHER NOTES:
This CFU checks modality identification before students apply it in their own writing. Sentence B also contains a nominalisation ("transportation") which observant students may notice.

WATCH FOR:
- Students who write A because "sometimes" is more accurate -- affirm the thinking but redirect: "In persuasive writing, we choose the language that makes our argument strongest, not the language that hedges"
- Students who identify "essential" AND "transportation" -- excellent, they are already seeing nominalisations

[General: CFU | VTLM 2.0: Formative Assessment]`;

const NOTES_WEDO = `SAY:
- Your turn to plan. You are going to complete your own SPO for the first body paragraph of your persuasive text
- Think about the reading we did today. What is ONE strong argument you can make about the role of animals in World War 1?
- It could be about transport, communication, companionship, morale, medical support -- but pick just ONE
- With your partner, discuss: what is your argument, and what evidence from the reading supports it?
- After you discuss, start filling in your SPO template. Topic sentence first -- make it strong, with high modality

DO:
- Distribute the Session 4 SPO Template
- Give 2 minutes for partner discussion
- Circulate as students begin their SPO -- check topic sentences for clarity and modality
- After 5 minutes: "You should have your topic sentence and at least one supporting detail"

CFU CHECKPOINT:
Technique: Cold Call

Script:
- After 3 minutes of planning, cold call 2-3 students: "Read me your topic sentence."
- Listen for: a clear argument (not a list of points), high modality language
- If a topic sentence is too vague: "That is a good start. Can you make the argument sharper? What EXACTLY are you arguing?"

PROCEED (>=80%): Most topic sentences state a clear argument with modality. Students continue planning.
PIVOT (<80%): Most likely issue -- topic sentences are too vague or list multiple arguments. Reteach: "Your topic sentence has ONE job: state ONE argument. Watch: 'Horses were important and dogs helped and pigeons carried messages' -- that is three arguments. Pick ONE. 'Horses were essential for transport.' That is one clear argument." Re-check: cold call 2 more students for revised topic sentences.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Provide the SPO template with the topic sentence pre-filled ("Horses were an important part of the war because they helped with transport.") and one supporting detail started. Students complete the remaining SDs and concluding sentence
- Extra Notes: Students can use their reading notes directly

EXTENDING PROMPT:
- Task: After completing the SPO, write a second topic sentence for a DIFFERENT argument. Below it, list how the modality would change if writing for a different audience (e.g., younger students vs. a government official)

TEACHER NOTES:
The We Do is guided planning. Students discuss with partners first then work on their SPO individually with teacher support. The SPO template is the critical scaffold -- it structures thinking before writing.

WATCH FOR:
- Students who try to fit multiple arguments into one paragraph -- redirect: "Save that for your next body paragraph"
- Students whose supporting details are opinions rather than evidence -- "Can you point to something from the article that supports this?"

[General: We Do | VTLM 2.0: Scaffold Practice]`;

const NOTES_YOUDO = `SAY:
- Now write your body paragraph from your SPO
- First: check your SPO is complete -- topic sentence, at least two supporting details, concluding sentence
- Next: write your paragraph. Turn each part of your SPO into full sentences. Try to include at least one high modality word
- Then: reread your paragraph. Check: does each sentence connect to your argument? Have you used high modality language? Fix anything that needs improving

DO:
- Students write independently
- Circulate -- check that paragraphs follow the SPO structure
- After 8 minutes: "Start rereading and editing now"
- Collect or photograph finished paragraphs for formative assessment

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: Students use their pre-filled SPO to write a shorter paragraph (topic sentence + 2 supporting sentences + concluding sentence). Focus on following the SPO structure rather than adding persuasive language
- Extra Notes: The enabling scaffold PDF has sentence starters for each SPO section

EXTENDING PROMPT:
- Task: After writing the paragraph, add a sentence that includes a nominalisation. Then write a brief note explaining how the nominalisation makes the writing sound more formal

TEACHER NOTES:
This is independent writing time. The editing procedures from L25 are embedded: students reread and self-edit before finishing. The key quality check is whether the paragraph follows the SPO structure and uses modality.

WATCH FOR:
- Students who write without referring to their SPO -- "Your SPO is your plan. Check it as you write"
- Students who forget a concluding sentence -- this is the most commonly omitted element
- Students who use modality naturally -- affirm it: "I can see you have used 'certainly' -- great persuasive language"

[General: You Do | VTLM 2.0: Independent Practice]`;

const NOTES_CLOSING = `SAY:
- Success criteria check. SC1: the purpose and features of a body paragraph -- thumbs? [scan]
- SC2: planning using the SPO template -- thumbs? [scan]
- SC3: using high modality language to strengthen arguments -- thumbs? [scan]
- Turn and talk: What was the strongest piece of evidence you found in the reading today? How did you use it in your paragraph?

DO:
- Run through each SC with thumbs check
- The turn-and-talk connects reading to writing -- the evidence loop
- Collect SPOs and paragraphs
- Preview: "Next session we continue with our War Horse reading and more writing"

TEACHER NOTES:
The closing connects reading (evidence gathering) to writing (evidence use). Students who are thumbs-down on SC3 may need the persuasive language reference sheet for homework or next session.

WATCH FOR:
- Students thumbs-down on SC2 -- they may need additional guided practice with the SPO in the next writing session
- Students who can name modality but did not use it in their writing -- the gap between knowledge and application

[General: Closing | VTLM 2.0: Review and Reflect]`;

const NOTES_RESOURCES = `SAY:
- Two printable resources for today
- The ${SPO_RESOURCE.name} is your planning template -- use it to organise your argument before writing
- The ${LANG_RESOURCE.name} has examples of high modality language and nominalisations you can use

DO:
- Print the SPO template before the lesson (one per student)
- Print the language reference (one per student or one per pair)
- Click any resource card to open the PDF

TEACHER NOTES:
The SPO template will be reused in Session 6 for the second body paragraph. Students should keep their completed SPO with their persuasive text drafts.

[General: Resources | VTLM 2.0: Student Resources]`;

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------

async function build() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "Claude";
  pres.title = "War Horse - Lessons 24-25 - Plan and Write a Body Paragraph";

  // =========================================================================
  // SLIDE 1 -- Title
  // =========================================================================
  titleSlide(
    pres,
    "War Horse",
    "Plan and Write a Body Paragraph",
    "Lessons 24-25  |  Week 5  |  Year 5/6 Literacy",
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
      "We are learning to plan and write a persuasive body paragraph using a single paragraph outline and persuasive language techniques",
    ],
    [
      "I can identify the purpose and features of a body paragraph in a persuasive text",
      "I can plan a body paragraph using the single paragraph outline template",
      "I can use high modality language and nominalisations to strengthen my arguments",
    ],
    NOTES_LI_SC,
    FOOTER
  );

  // =========================================================================
  // SLIDE 3 -- Reading Launch
  // =========================================================================
  contentSlide(
    pres,
    "Read Aloud",
    C.PRIMARY,
    "World War 1 -- National Geographic Kids",
    [
      "Reading Mode: Student Read Aloud -- take notes as you read",
      "Sections: Why did it start? | How long did it last? | Life in the trenches",
      "A new type of war | Animals during WW1 | Life in Britain during WW1",
      "Women during WW1 | Children during WW1 | Armistice Day",
      "Focus: What evidence could you use in your persuasive text?",
    ],
    NOTES_READING,
    FOOTER
  );

  // =========================================================================
  // SLIDE 4 -- Body Paragraph Features + SPO
  // =========================================================================
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "Body Paragraph Structure -- The SPO",
    [
      "Purpose: present ONE key argument that supports your position",
      "Topic Sentence (TS) -- states your argument clearly",
      "Supporting Details (SD) -- evidence, examples, explanation",
      "Concluding Sentence (CS) -- wraps up and links back to your position",
    ],
    NOTES_MACROSTRUCTURE,
    FOOTER
  );

  // =========================================================================
  // SLIDE 5 -- High Modality Language
  // =========================================================================
  modellingSlide(
    pres,
    "I Do",
    "High Modality Language",
    "Modality = how strongly you say something\n\nHIGH modality (persuasive): must, always, certainly, undoubtedly, essential, every, vital, clearly\n\nLOW modality (softer): might, sometimes, perhaps, could, possibly",
    "Neutral: \"Animals helped in the war.\"\n\nHigh modality: \"Animals were undoubtedly essential to the war effort.\"\n\nControlled: \"Most soldiers relied on horses\" is more credible than \"Every single soldier relied on horses\"",
    NOTES_MODALITY,
    FOOTER
  );

  // =========================================================================
  // SLIDE 6 -- Teacher Models SPO
  // =========================================================================
  modellingSlide(
    pres,
    "I Do",
    "Modelling the SPO -- Horses as Transport",
    "TS: Horses were undoubtedly one of the most essential forms of transport during World War 1.\n\nSD1: Carried soldiers, weapons, and supplies to the front lines\nSD2: Armies could not move artillery across muddy terrain without horses\nSD3: Horses pulled ambulance carts to rescue wounded soldiers",
    "CS: The critical role horses played in wartime transport clearly demonstrates their importance to the war effort.\n\nNotice: high modality words -- \"undoubtedly\", \"most essential\", \"critical\", \"clearly\"",
    NOTES_SPO_MODEL,
    FOOTER
  );

  // =========================================================================
  // SLIDE 7 -- Teacher Models Writing from SPO
  // =========================================================================
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "From SPO to Paragraph",
    [
      "Horses were undoubtedly one of the most essential forms of transport during World War 1. They carried soldiers, weapons, and vital supplies to the front lines, often through terrain that no vehicle could cross.",
      "Without horses, armies would have been unable to move artillery or ammunition across the muddy, shell-damaged battlefields of the Western Front. Horses also pulled ambulance carts, rescuing wounded soldiers who would otherwise have been left behind.",
      "The critical role horses played in wartime transport clearly demonstrates their importance to the war effort.",
    ],
    NOTES_WRITING_MODEL,
    FOOTER
  );

  // =========================================================================
  // SLIDES 8-9 -- CFU (withReveal)
  // =========================================================================
  withReveal(
    () => cfuSlide(
      pres,
      "CFU",
      "Which Has Higher Modality?",
      "Show Me Boards",
      "A) \"Horses sometimes helped with transport during the war.\"\n\nB) \"Horses were essential to the transportation of troops and supplies during the war.\"\n\nOn your whiteboard: write A or B, and write the high modality word.",
      NOTES_CFU,
      FOOTER
    ),
    (slide) => {
      const ansY = 4.05;
      slide.addShape("roundRect", {
        x: 0.5, y: ansY, w: 9, h: 0.90, rectRadius: 0.10,
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

      slide.addText("B -- \"essential\" is high modality. Also notice \"transportation\" -- a nominalisation (noun form of \"transport\").", {
        x: 2.3, y: ansY + 0.08, w: 7.0, h: 0.72,
        fontSize: 13, fontFace: FONT_B, color: C.WHITE, valign: "middle", margin: 0,
      });
      slide.addNotes(NOTES_CFU);
    }
  );

  // =========================================================================
  // SLIDE 10 -- We Do: Plan Own SPO
  // =========================================================================
  contentSlide(
    pres,
    "We Do",
    C.SUCCESS,
    "Plan Your Body Paragraph",
    [
      "Choose ONE argument about animals in World War 1",
      "Ideas: transport, communication, companionship, morale, medical support",
      "Discuss with your partner: what is your argument? What evidence supports it?",
      "Fill in your SPO template: Topic Sentence (with high modality) -> Supporting Details -> Concluding Sentence",
    ],
    NOTES_WEDO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 11 -- You Do: Write Paragraph
  // =========================================================================
  taskSlide(
    pres,
    "You Do",
    "Write Your Body Paragraph",
    [
      { label: "FIRST", instruction: "Check your SPO is complete: topic sentence, supporting details, concluding sentence" },
      { label: "NEXT", instruction: "Write your paragraph from your SPO. Include at least one high modality word to strengthen your argument" },
      { label: "THEN", instruction: "Reread your paragraph and edit: Does each sentence connect to your argument? Is your language persuasive?" },
    ],
    NOTES_YOUDO,
    FOOTER
  );

  // =========================================================================
  // SLIDE 12 -- Closing
  // =========================================================================
  closingSlide(
    pres,
    "What was the strongest piece of evidence you found in the reading today? How did you use it in your paragraph?",
    [
      "I can identify the purpose and features of a body paragraph in a persuasive text",
      "I can plan a body paragraph using the single paragraph outline template",
      "I can use high modality language and nominalisations to strengthen my arguments",
    ],
    NOTES_CLOSING
  );

  // =========================================================================
  // SLIDE 13 -- Resources
  // =========================================================================

  // =========================================================================
  // Generate companion PDFs
  // =========================================================================

  // --- PDF 1: SPO Template --------------------------------------------------
  const spo = createPdf({ title: SPO_RESOURCE.name });
  let spoY = addPdfHeader(spo, "Single Paragraph Outline (SPO)", {
    color: C.PRIMARY,
    subtitle: "Body Paragraph Planner",
    lessonInfo: "War Horse | Lessons 24-25 | Week 5 | Year 5/6 Literacy",
    showNameDate: true,
  });

  spoY = addTipBox(spo, "A body paragraph has ONE job: present one key argument that supports your position.\n\nUse this template to plan your paragraph BEFORE you write it.", spoY, { color: C.PRIMARY });

  spoY = addSectionHeading(spo, "My Argument (in one sentence):", spoY, { color: C.PRIMARY });
  spoY = addLinedArea(spo, spoY, 1, { lineSpacing: 28 });
  spoY += 6;

  spoY = addSectionHeading(spo, "Topic Sentence (TS):", spoY, { color: C.PRIMARY });
  spoY = addBodyText(spo, "State your argument clearly. Use high modality language (must, essential, undoubtedly, clearly).", spoY, { fontSize: 9, italic: true });
  spoY = addLinedArea(spo, spoY, 2, { lineSpacing: 26 });
  spoY += 6;

  spoY = addSectionHeading(spo, "Supporting Detail 1 (SD1):", spoY, { color: C.SECONDARY });
  spoY = addBodyText(spo, "Evidence, example, or explanation that supports your argument.", spoY, { fontSize: 9, italic: true });
  spoY = addLinedArea(spo, spoY, 2, { lineSpacing: 26 });
  spoY += 6;

  spoY = addSectionHeading(spo, "Supporting Detail 2 (SD2):", spoY, { color: C.SECONDARY });
  spoY = addLinedArea(spo, spoY, 2, { lineSpacing: 26 });
  spoY += 6;

  spoY = addSectionHeading(spo, "Supporting Detail 3 (SD3) -- Optional:", spoY, { color: C.SECONDARY });
  spoY = addLinedArea(spo, spoY, 2, { lineSpacing: 26 });
  spoY += 6;

  spoY = addSectionHeading(spo, "Concluding Sentence (CS):", spoY, { color: C.ACCENT });
  spoY = addBodyText(spo, "Wrap up the argument and link back to your position.", spoY, { fontSize: 9, italic: true });
  spoY = addLinedArea(spo, spoY, 2, { lineSpacing: 26 });
  spoY += 10;

  spoY = addSectionHeading(spo, "My Full Paragraph:", spoY, { color: C.PRIMARY, fontSize: 14 });
  spoY = addBodyText(spo, "Write your body paragraph here, following your SPO plan.", spoY, { fontSize: 9, italic: true });
  spoY = addLinedArea(spo, spoY, 10, { lineSpacing: 26 });

  addPdfFooter(spo, "War Horse | Lessons 24-25 | SPO Template");

  // --- PDF 2: Persuasive Language Reference ---------------------------------
  const lang = createPdf({ title: LANG_RESOURCE.name });
  let langY = addPdfHeader(lang, "Persuasive Language Reference", {
    color: C.ACCENT,
    subtitle: "Modality and Nominalisations",
    lessonInfo: "War Horse | Lessons 24-25 | Week 5 | Year 5/6 Literacy",
    showNameDate: false,
  });

  langY = addSectionHeading(lang, "High Modality Language", langY, { color: C.PRIMARY });
  langY = addBodyText(lang, "High modality words make your writing sound confident and persuasive. Use them in your topic sentences and concluding sentences.", langY, { fontSize: 10 });
  langY += 4;
  langY = addBodyText(lang, "must  |  always  |  certainly  |  undoubtedly  |  essential  |  vital  |  clearly  |  without question  |  every  |  significant  |  critical  |  fundamental", langY, { fontSize: 11, bold: true });
  langY += 8;
  langY = addBodyText(lang, "Example: \"Animals were undoubtedly essential to the war effort.\"", langY, { fontSize: 10, italic: true });
  langY += 6;
  langY = addTipBox(lang, "Control your modality for precision. \"Most soldiers relied on horses\" is more credible than \"Every single soldier relied on horses.\" Choose the strongest language that is STILL truthful.", langY, { color: C.ALERT });

  langY = addSectionHeading(lang, "Nominalisations", langY, { color: C.SECONDARY });
  langY = addBodyText(lang, "A nominalisation turns a verb or adjective into a noun. This makes your writing sound more formal and authoritative.", langY, { fontSize: 10 });
  langY += 4;
  langY = addBodyText(lang, "important -> importance  |  transport -> transportation  |  contribute -> contribution", langY, { fontSize: 11, bold: true });
  langY = addBodyText(lang, "survive -> survival  |  destroy -> destruction  |  significant -> significance", langY, { fontSize: 11, bold: true });
  langY += 6;
  langY = addBodyText(lang, "Before: \"Horses were important because they transported supplies.\"", langY, { fontSize: 10 });
  langY = addBodyText(lang, "After: \"The importance of horses in the transportation of supplies cannot be overstated.\"", langY, { fontSize: 10, italic: true });
  langY += 8;

  langY = addSectionHeading(lang, "Combining Both Tools", langY, { color: C.ACCENT });
  langY = addBodyText(lang, "The strongest persuasive writing uses BOTH modality and nominalisations together:", langY, { fontSize: 10 });
  langY += 4;
  langY = addBodyText(lang, "\"The contribution of animals to the war effort was undoubtedly significant, and their sacrifice must always be remembered.\"", langY, { fontSize: 11, bold: true });
  langY += 4;
  langY = addBodyText(lang, "High modality: undoubtedly, must, always", langY, { fontSize: 9, italic: true });
  langY = addBodyText(lang, "Nominalisations: contribution (from contribute), sacrifice (from sacrifice/verb form)", langY, { fontSize: 9, italic: true });

  addPdfFooter(lang, "War Horse | Lessons 24-25 | Persuasive Language Reference");

  // --- Write all files ------------------------------------------------------
  await Promise.all([
    pres.writeFile({ fileName: `${OUT_DIR}/WH5_Lesson24.pptx` }),
    writePdf(spo, SPO_PDF_PATH),
    writePdf(lang, LANG_PDF_PATH),
  ]);

  console.log(`PPTX written to ${OUT_DIR}/WH5_Lesson24.pptx`);
  console.log(`Done: ${SPO_RESOURCE.name}`);
  console.log(`Done: ${LANG_RESOURCE.name}`);
}

build().catch(console.error);
