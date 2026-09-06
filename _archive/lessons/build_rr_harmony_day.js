"use strict";

// Respectful Relationships + Harmony Day — Grade 5/6 Wellbeing
// Single 60-minute session: Harmony Day context + Emotional Intensity activities
// Groups of 2-4, fun and achievable

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");
const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf,
  addPdfHeader, addSectionHeading, addBodyText, addTipBox,
  addPdfFooter, addLinedArea, addWriteLine,
  addResourceSlide, makeSessionResource, formatSessionResourceFileName,
  getSessionResourceFolder,
  PAGE, hex,
} = require("../themes/pdf_helpers");

// ── Theme ──
const T = createTheme("wellbeing", "grade56", weekToVariant(12));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, cfuSlide, closingSlide,
  scenarioSlide, reflectionSlide, pairShareSlide,
  withReveal,
  addTopBar, addBadge, addTitle, addCard, addFooter, addTextOnShape,
  addInstructionCard,
  CONTENT_TOP, SAFE_BOTTOM, SLIDE_W,
  runSlideDiagnostics,
} = T;

// ── Output paths ──
const UNIT = "RR_Harmony_Day_Emotional_Intensity";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "Harmony Day - Emotional Intensity.pptx";
const FOOTER = "Harmony Day | Emotional Intensity | Grade 5/6 Wellbeing";
const SESSION = 1;
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

// ── Resources ──
const INTENSITY_CARDS_RESOURCE = makeSessionResource(
  SESSION,
  "Emotions Intensity Cards",
  "Printable emotion word pairs for the statues game and group activity."
);
const ROLLER_COASTER_RESOURCE = makeSessionResource(
  SESSION,
  "Roller Coaster Worksheet",
  "Casey's Day roller coaster diagram with emotional intensity word list."
);
const RESOURCE_ITEMS = [INTENSITY_CARDS_RESOURCE, ROLLER_COASTER_RESOURCE];

fs.mkdirSync(RES_DIR, { recursive: true });

// ═══════════════════════════════════════════════════════════════
// Teacher Notes
// ═══════════════════════════════════════════════════════════════

const NOTES_TITLE = [
  "SAY:",
  "- Happy Harmony Day, everyone. Today is March 20 -- Harmony Day in Australia",
  "- Our session today blends Harmony Day with our Respectful Relationships work on understanding emotions",
  "- We are going to explore how emotions can range from mild to intense, and how understanding this helps us support each other -- which is what Harmony Day is all about",
  "",
  "DO:",
  "- Display slide as students settle",
  "- Have the room cleared for movement space (push desks to edges if possible)",
  "",
  "TEACHER NOTES:",
  "Single 60-minute session combining Harmony Day (March 20) with RR emotional intensity content. The session moves through three connected activities: emotion statues game, Casey's Day roller coaster, and a group creative task. All group work is in pairs or small groups of 3-4.",
  "",
  "WATCH FOR:",
  "- Students who may feel uncomfortable with the movement activity -- offer the mirror-pairs option",
  "- Students who need a moment to transition from regular class to wellbeing mode",
  "",
  "[General: Title | VTLM 2.0: Establishing Purpose]",
].join("\n");

const NOTES_LI = [
  "SAY:",
  "- Read the learning intention from the slide",
  "- Read through the three success criteria together",
  "- SC1 is about recognising that emotions come in different strengths. SC2 is about using precise words to describe those strengths. SC3 is about identifying what triggers stronger emotions",
  "",
  "DO:",
  "- Choral read the LI, then each SC",
  "- Brief check: Thumbs up if you have heard the word 'intensity' before",
  "",
  "TEACHER NOTES:",
  "The LI connects Harmony Day's theme of understanding others with the RR emotional intensity content. SC1 is the foundation -- every student should grasp the concept. SC2 builds vocabulary. SC3 requires application and empathy.",
  "",
  "WATCH FOR:",
  "- Students unsure about 'intensity' -- explain simply: how strong or weak an emotion feels",
  "- Students who associate emotions only with negative feelings -- we will address both positive and negative",
  "",
  "[General: LI/SC | VTLM 2.0: Clear Learning Intention]",
].join("\n");

const NOTES_HARMONY = [
  "SAY:",
  "- Harmony Day is celebrated on March 20 every year across Australia",
  "- The message is: Everyone Belongs. It is about respect, inclusiveness, and a sense of belonging for everyone",
  "- Today we are connecting this to our emotions work. When we understand how people feel -- and how strongly they feel -- we become better at including and supporting everyone",
  "- Ask: Why might understanding emotional intensity help us be more inclusive? [When we recognise someone is feeling something strongly, we can respond with more care and empathy]",
  "",
  "DO:",
  "- Point to the key messages on slide",
  "- Allow 30 seconds for a Turn and Talk: What does 'Everyone Belongs' mean to you?",
  "",
  "TEACHER NOTES:",
  "This slide bridges Harmony Day into the RR content. Keep it brief (2-3 minutes). The connection is genuine: understanding emotional intensity builds empathy, which is a core Harmony Day value.",
  "",
  "WATCH FOR:",
  "- Students who share personal experiences of not belonging -- acknowledge warmly but do not dwell; the lesson will give structured ways to explore this",
  "",
  "[General: Context Setting | VTLM 2.0: Prior Knowledge Activation]",
].join("\n");

const NOTES_INTENSITY_INTRO = [
  "SAY:",
  "- Emotions are not just on or off. They come in different strengths -- like a volume dial",
  "- Being a bit annoyed is different from being furious. Feeling pleased is different from being ecstatic",
  "- We call this emotional intensity -- how strongly we feel an emotion",
  "- Having precise words for these different levels helps us understand ourselves and communicate with others",
  "",
  "DO:",
  "- Point to the intensity scale visual on the slide",
  "- Model one example: 'I might feel a bit nervous before a test. But if I find out I have a surprise exam right now -- that is anxious or even panicked. Same family of emotion, different intensity'",
  "",
  "TEACHER NOTES:",
  "This is the I Do -- teacher-led explanation of the core concept. Keep it concrete with the volume dial analogy. The slide shows the visual scale; teacher narration brings it to life.",
  "",
  "WATCH FOR:",
  "- Students nodding along -- good sign of connection to experience",
  "- Students who seem to think there are only 'happy' and 'sad' -- the intensity concept expands this",
  "",
  "[General: I Do | VTLM 2.0: Explicit Explanation]",
].join("\n");

const NOTES_STATUES = [
  "SAY:",
  "- We are going to play the Emotion Statues game. Here is how it works",
  "- I will call out an emotion word -- like 'annoyed'. You freeze into a statue showing that emotion with your face and body",
  "- Then I will wind up the intensity. I will call out a stronger word -- like 'furious'. You change your statue to show that stronger feeling",
  "- Watch me first: [demonstrate annoyed -> furious with face and posture]",
  "",
  "DO:",
  "- Clear space in the room. Students stand with hands by sides",
  "- Option A: students spread out in a circle. Option B: pairs face each other in mirror position",
  "- Run 3-4 rounds using the Emotions Intensity Cards",
  "- After each round, arrange half the class to observe the other half",
  "",
  "TEACHER NOTES:",
  "This is the We Do -- guided, active practice. The game builds empathy and emotional vocabulary simultaneously. Mirror pairs work well for students who are self-conscious. Keep pacing brisk -- about 60-90 seconds per round.",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Students work in pairs, showing just facial expressions (not full body) for each emotion. Limit to 2 emotion pairs they feel confident with.",
  "- Extra Notes: Some students find full-body expression confronting. Facial expression alone still builds the vocabulary connection.",
  "EXTENDING PROMPT:",
  "- Task: After the main rounds, extending students create their own emotion intensity pair and demonstrate it. They explain what events might trigger each level.",
  "",
  "WATCH FOR:",
  "- Students who are very still or reluctant -- invite them to do small expressions, even just with their hands",
  "- Students who are being silly rather than expressive -- redirect: 'Show me what furious really looks like in your face'",
  "- Readiness signal: students can visibly distinguish between the mild and intense version of the same emotion",
  "",
  "[General: We Do | VTLM 2.0: Scaffold Practice]",
].join("\n");

const NOTES_STATUES_DEBRIEF = [
  "SAY:",
  "- What did all these emotions have in common? [They were all negative or challenging emotions -- the kinds we feel when life gets tough]",
  "- Could you see the difference between the mild version and the intense version? What changed in people's bodies?",
  "- Ask: Why is it useful to have different words for different intensities? [It helps us communicate exactly how we feel, and it helps others understand and respond better]",
  "",
  "DO:",
  "- Students return to seats or sit in a circle",
  "- Cold Call 3-4 students for observations about what they noticed",
  "",
  "CFU CHECKPOINT:",
  "Technique: Cold Call with follow-up probing",
  "Script:",
  "- Say: I am going to pick a few people. Tell me: what was the difference between the first emotion and the wound-up version? Not just the word -- what changed in your body or face?",
  "- Scan for: students who can articulate physical and emotional differences, not just name the words",
  "PROCEED: If 80%+ can describe the difference between mild and intense, move to Casey's Day.",
  "PIVOT: If students struggle to distinguish intensity levels, replay one round slowly with the class and narrate: 'See how annoyed is tense shoulders but furious is clenched fists, tight jaw, whole body rigid.' Then re-check with a new pair.",
  "",
  "TEACHER NOTES:",
  "Debrief is essential -- it is where the learning is made explicit. The game is engaging but without debrief it stays as 'fun activity'. This CFU checks whether students have genuinely grasped the intensity concept.",
  "",
  "WATCH FOR:",
  "- Students who describe emotions as just 'bad' or 'good' without gradation -- prompt: 'Was it the same kind of bad, or did it get worse?'",
  "- Students who connect to real experiences -- affirm this and note for sensitivity",
  "",
  "[General: CFU | VTLM 2.0: Monitor Progress]",
].join("\n");

const NOTES_CASEY_INTRO = [
  "SAY:",
  "- Now we are going to look at how emotions change throughout a single day -- like a roller coaster with highs and lows",
  "- I am going to read you a story about Casey's morning. Listen for how Casey's emotions change with each event",
  "- As I read, think about: is each emotion mild, medium, or intense?",
  "",
  "DO:",
  "- Read Casey's Day story aloud with expression",
  "- Pause briefly after each event to let students register the emotional shift",
  "- After reading, distribute the roller coaster handout",
  "",
  "TEACHER NOTES:",
  "Casey's Day is the bridge between the statues game (physical expression of intensity) and analytical labelling. The story is deliberately relatable -- school morning, sibling issues, friend dynamics. Reading aloud models emotional interpretation.",
  "",
  "WATCH FOR:",
  "- Students who are already labelling emotions as you read -- great sign of engagement",
  "- Students who connect Casey's experiences to their own mornings",
  "",
  "[General: I Do | VTLM 2.0: Explicit Modelling]",
].join("\n");

const NOTES_CASEY_WE_DO = [
  "SAY:",
  "- Let's map Casey's roller coaster together. Look at the diagram on your handout",
  "- First event: Casey steps on mushy biscuits. How does Casey feel? [Disgusted, grossed out]. Where would that sit on the roller coaster -- mild, medium, or intense? [Probably medium -- it is unpleasant but not devastating]",
  "- Second event: Dad made a favourite breakfast. [Happy, pleased, grateful]. Is that a high or a low? [Definitely a high -- a positive moment]",
  "- I want you to work with your group to label the rest of Casey's events on the roller coaster. Use the Emotional Intensity Word List on your handout to find precise words",
  "",
  "DO:",
  "- Model the first two events on the board or slide, mapping them onto the roller coaster",
  "- Students work in groups of 2-4 to complete the remaining events",
  "- Circulate and check labelling. Are groups using intensity words or just basic emotion words?",
  "- Allow 8-10 minutes for group work",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Students label events as simply 'positive' or 'negative' and rate intensity as 1 (mild), 2 (medium), or 3 (intense) using numbers rather than vocabulary words.",
  "- Extra Notes: This gives a concrete numerical scaffold before introducing vocabulary.",
  "EXTENDING PROMPT:",
  "- Task: After completing Casey's roller coaster, students identify which single event would have the most lasting emotional impact and write 2-3 sentences explaining why, using at least two intensity words.",
  "",
  "TEACHER NOTES:",
  "This is the We Do. Teacher models the first two events, then groups complete the rest. The shift from whole-class to group work is the release. The Emotional Intensity Word List handout is the scaffold -- it gives vocabulary options at different intensity levels.",
  "",
  "WATCH FOR:",
  "- Groups using only basic words (happy, sad, angry) -- redirect to the word list: 'Can you find a more precise word that shows HOW happy?'",
  "- Groups rushing through without discussing -- prompt: 'Talk about it first. Do you all agree on the intensity level?'",
  "- Readiness signal: groups using at least 2 different intensity levels for the same emotion family",
  "",
  "[General: We Do | VTLM 2.0: Scaffold Practice]",
].join("\n");

const NOTES_CFU_CASEY = [
  "SAY:",
  "- Quick check. I am going to read one of Casey's events. On your whiteboard, write the emotion AND a word that shows the intensity",
  "- Event: Casey was the last one picked for basketball, and heard Lou say 'Don't pick him. He'll make us lose'",
  "- Write your answer now. You have 30 seconds",
  "",
  "DO:",
  "- Students write on whiteboards",
  "- On signal, hold up boards",
  "- Scan for: emotion family (hurt, rejected, humiliated, embarrassed) AND intensity level (this is a high-intensity negative)",
  "",
  "CFU CHECKPOINT:",
  "Technique: Show Me Boards",
  "Script:",
  "- Say: Write the emotion Casey felt AND a word from the intensity list. 30 seconds. Hold up on my signal.",
  "- Scan for: students writing intensity-level words (humiliated, crushed, devastated) rather than just 'sad'. Accept any emotion family that fits (hurt, rejected, embarrassed) as long as the intensity is at the stronger end.",
  "PROCEED: If 80%+ use an intensity word (not just 'sad' or 'angry'), proceed to the creative task.",
  "PIVOT: If most students write only basic emotion words, reteach: 'Sad is the family. But HOW sad? A little sad is disappointed. Very sad is devastated. Being picked last and hearing someone say you will make them lose -- that is not just sad. That is closer to humiliated or crushed.' Re-check with a second event.",
  "",
  "TEACHER NOTES:",
  "This CFU targets SC2 -- using precise vocabulary for emotional intensity. The chosen event is deliberately high-impact to test whether students can identify strong intensity, not just name the emotion family.",
  "",
  "WATCH FOR:",
  "- Students who write 'sad' -- they understand the emotion but not the intensity layer yet",
  "- Students who write 'angry' -- valid, but probe: 'Angry at who? And how angry?'",
  "",
  "[General: CFU | VTLM 2.0: Monitor Progress]",
].join("\n");

const NOTES_YOU_DO = [
  "SAY:",
  "- Now your group is going to create your own roller coaster day story",
  "- You will choose one of these topics: A day in the life of a new Prep student, A day in the life of a new student joining your class, or The evening of a new babysitter",
  "- Your story needs at least six emotions, including at least one high, one low, and one in-between",
  "- Use the Emotional Intensity Word List to choose precise words",
  "",
  "DO:",
  "- Display the three topic options on the slide",
  "- Groups of 2-4 choose their topic",
  "- Step 1: List the highs, lows, and in-betweens for your character's day",
  "- Step 2: Draw the roller coaster and map the events and emotions",
  "- Step 3: Each person writes up the story individually",
  "- Allow 15 minutes. Circulate and support",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Group maps only 4 events (2 highs, 2 lows) and uses the intensity word list as a menu to select from. They draw the roller coaster but do not need to write a full narrative -- dot points are fine.",
  "EXTENDING PROMPT:",
  "- Task: Group includes at least one moment of mixed emotions (e.g., excited AND nervous about meeting a new class) and explains how mixed emotions can also vary in intensity.",
  "",
  "TEACHER NOTES:",
  "This is the You Do -- independent group application. The topics are deliberately empathy-driven: all three scenarios involve imagining how SOMEONE ELSE feels, linking back to Harmony Day. Different content from We Do (Casey's given story vs their own created story), increased complexity (they must invent events and match emotions).",
  "",
  "WATCH FOR:",
  "- Groups who pick a topic but struggle to start -- prompt: 'What is the first thing that happens in your character's day? How do they feel about it?'",
  "- Groups creating only dramatic/extreme events -- encourage some in-between moments too",
  "- Readiness signal: groups mapping emotions at different intensity levels on their roller coaster with precise vocabulary",
  "",
  "[General: You Do | VTLM 2.0: Supported Application]",
].join("\n");

const NOTES_SHARE = [
  "SAY:",
  "- Let's hear some of your roller coaster stories",
  "- As each group shares, listen for the emotion words they chose. Did they pick words that show intensity, or just basic emotion words?",
  "- After each group, we will give one piece of feedback: What emotion word did they use that really showed the intensity well?",
  "",
  "DO:",
  "- Invite 2-3 groups to share (based on time available)",
  "- After each share, ask the class: 'What was the best intensity word they used?'",
  "- Connect back to Harmony Day: 'When we can imagine how intensely someone else might feel, we are practising empathy -- and that is how everyone belongs'",
  "",
  "TEACHER NOTES:",
  "Gallery share serves dual purpose: celebrating group work and reinforcing vocabulary through peer modelling. Keep it brisk -- 2 minutes per group maximum. The Harmony Day connection at the end brings the session full circle.",
  "",
  "WATCH FOR:",
  "- Groups who are reluctant to share -- offer to share just their roller coaster diagram",
  "- Rich vocabulary use -- celebrate it publicly",
  "",
  "[General: Share | VTLM 2.0: Feedback and Reflection]",
].join("\n");

const NOTES_CLOSING = [
  "SAY:",
  "- Let's look back at our success criteria from the start of the session",
  "- SC1: I can explain that emotions range from mild to intense. Give me a thumbs up, sideways, or down",
  "- SC2: I can use precise words to describe different levels of emotional intensity. Thumbs?",
  "- SC3: I can identify events that might trigger stronger emotions in myself or others. Thumbs?",
  "- Today is Harmony Day. Understanding how people feel -- and how strongly they feel -- is one of the most important ways we show respect and make sure everyone belongs",
  "",
  "DO:",
  "- Display SC on slide. Read each one. Students show thumbs for each",
  "- Note students showing thumbs-down on any SC for follow-up",
  "- Final Turn and Talk prompt on slide",
  "",
  "TEACHER NOTES:",
  "The closing ties the emotional intensity learning back to Harmony Day's core message. Students self-assess against the SC and reflect on how understanding emotional intensity connects to belonging and respect.",
  "",
  "WATCH FOR:",
  "- Students consistently at thumbs-sideways on SC2 -- they may need more vocabulary exposure in future RR sessions",
  "- Students who make genuine connections between emotional intensity and Harmony Day -- affirm this thinking",
  "",
  "[General: Closing | VTLM 2.0: Review and Reflection]",
].join("\n");

const NOTES_RESOURCES = [
  "SAY:",
  "- Here are the resources for today's session. The Emotions Intensity Cards are for the statues game. The Roller Coaster Worksheet has Casey's story diagram and the word list",
  "",
  "DO:",
  "- Print Emotions Intensity Cards -- one set per class (cut into individual cards)",
  "- Print Roller Coaster Worksheet -- one per student or one per group",
  "",
  "TEACHER NOTES:",
  "Both resources support the session activities. Print the intensity cards before the lesson and cut them. The roller coaster worksheet includes Casey's Day text extract, the roller coaster diagram, and the emotional intensity word list.",
  "",
  "WATCH FOR:",
  "- Ensure all groups have access to the word list during the You Do activity",
  "",
  "[General: Resources | VTLM 2.0: Preparation]",
].join("\n");

// ═══════════════════════════════════════════════════════════════
// Build
// ═══════════════════════════════════════════════════════════════

async function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";

  // ================================================================
  // SLIDE 1: Title
  // ================================================================
  titleSlide(
    pres,
    "Harmony Day\nEmotional Intensity",
    "Respectful Relationships",
    "Grade 5/6 Wellbeing | March 20",
    NOTES_TITLE
  );

  addResourceSlide(
    pres,
    RESOURCE_ITEMS,
    { C, FONT_H, FONT_B },
    FOOTER,
    NOTES_RESOURCES
  );

  // ================================================================
  // SLIDE 2: LI / SC
  // ================================================================
  liSlide(
    pres,
    ["We are learning to understand emotional intensity so we can better recognise and respect how others feel"],
    [
      "I can explain that emotions range from mild to intense",
      "I can use precise words to describe different levels of emotional intensity",
      "I can identify events that might trigger stronger emotions in myself or others",
    ],
    NOTES_LI,
    FOOTER
  );

  // ================================================================
  // SLIDE 3: Harmony Day Context
  // ================================================================
  contentSlide(
    pres,
    "Harmony Day",
    C.PRIMARY,
    "Everyone Belongs",
    [
      "March 20 - Harmony Day Australia",
      "Respect  |  Inclusiveness  |  Belonging",
      "Today's connection:",
      "When we understand how strongly people feel, we can respond with empathy and care",
    ],
    NOTES_HARMONY,
    FOOTER
  );

  // ================================================================
  // SLIDE 4: Emotional Intensity - I Do
  // ================================================================
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "What Is Emotional Intensity?",
    [
      "Emotions come in different strengths",
      "Mild: annoyed ... nervous ... pleased",
      "Medium: angry ... anxious ... happy",
      "Intense: furious ... panicked ... ecstatic",
    ],
    NOTES_INTENSITY_INTRO,
    FOOTER,
    (slide, layoutGuide) => {
      // Visual: intensity scale on the right
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW;
      const topY = layoutGuide.panelTopPadded;

      // Scale background
      addCard(slide, rx, topY, rw, 3.2, { fill: C.BG_CARD, strip: C.SECONDARY });

      // Scale title
      slide.addText("Intensity Scale", {
        x: rx + 0.15, y: topY + 0.1, w: rw - 0.3, h: 0.4,
        fontSize: 14, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
      });

      // Low band
      addTextOnShape(slide, "LOW", {
        x: rx + 0.3, y: topY + 0.6, w: rw - 0.6, h: 0.55, rectRadius: 0.06,
        fill: { color: C.SUCCESS },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

      // Medium band
      addTextOnShape(slide, "MEDIUM", {
        x: rx + 0.3, y: topY + 1.25, w: rw - 0.6, h: 0.55, rectRadius: 0.06,
        fill: { color: C.ACCENT },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

      // High band
      addTextOnShape(slide, "HIGH", {
        x: rx + 0.3, y: topY + 1.9, w: rw - 0.6, h: 0.55, rectRadius: 0.06,
        fill: { color: C.ALERT },
      }, { fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true });

      // Arrow indicator
      slide.addText("Volume dial -->", {
        x: rx + 0.15, y: topY + 2.6, w: rw - 0.3, h: 0.4,
        fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true, align: "center", margin: 0,
      });
    }
  );

  // ================================================================
  // SLIDE 5: Emotion Statues Game - We Do
  // ================================================================
  contentSlide(
    pres,
    "We Do",
    C.SECONDARY,
    "Emotion Statues Game",
    [
      "How to play:",
      "1.  Stand with hands by your sides",
      "2.  Teacher calls an emotion word -- freeze into a statue",
      "3.  Teacher winds up the intensity -- change your statue",
      "4.  Half the class watches, then swap",
      "Mirror option: face a partner and match poses",
    ],
    NOTES_STATUES,
    FOOTER,
    (slide, layoutGuide) => {
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW;
      const topY = layoutGuide.panelTopPadded;

      // Example pairs card
      addCard(slide, rx, topY, rw, 3.0, { fill: C.BG_CARD, strip: C.PRIMARY });
      slide.addText("Example Pairs", {
        x: rx + 0.15, y: topY + 0.1, w: rw - 0.3, h: 0.35,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
      });

      const pairs = [
        ["annoyed", "furious"],
        ["nervous", "terrified"],
        ["disappointed", "devastated"],
        ["worried", "panicked"],
      ];

      pairs.forEach((pair, i) => {
        const py = topY + 0.55 + i * 0.55;
        slide.addText([
          { text: pair[0], options: { fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL } },
          { text: "  -->  ", options: { fontSize: 12, fontFace: FONT_B, color: C.MUTED } },
          { text: pair[1], options: { fontSize: 13, fontFace: FONT_B, color: C.ALERT, bold: true } },
        ], {
          x: rx + 0.25, y: py, w: rw - 0.5, h: 0.45,
          valign: "middle", margin: 0,
        });
      });
    }
  );

  // ================================================================
  // SLIDE 6: Statues Debrief - CFU
  // ================================================================
  cfuSlide(
    pres,
    "Check",
    "What Did You Notice?",
    "Cold Call",
    "What changed in your body or face when the emotion got more intense?",
    NOTES_STATUES_DEBRIEF,
    FOOTER
  );

  // ================================================================
  // SLIDE 7: Casey's Day - I Do (Read Aloud)
  // ================================================================
  contentSlide(
    pres,
    "I Do",
    C.PRIMARY,
    "A Day in the Life of Casey",
    [
      "Listen for the emotional highs and lows",
      "As I read, think about:",
      "What emotion does Casey feel at each moment?",
      "Is it mild, medium, or intense?",
      "What event triggered the change?",
    ],
    NOTES_CASEY_INTRO,
    FOOTER,
    (slide, layoutGuide) => {
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW;
      const topY = layoutGuide.panelTopPadded;

      // Roller coaster visual hint
      addCard(slide, rx, topY, rw, 2.8, { fill: C.BG_CARD, strip: C.ACCENT });
      slide.addText("Casey's Roller Coaster", {
        x: rx + 0.15, y: topY + 0.1, w: rw - 0.3, h: 0.35,
        fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true, margin: 0,
      });

      // Simple roller coaster path representation
      const events = [
        { label: "Mushy biscuit", level: "low" },
        { label: "Favourite brekky", level: "high" },
        { label: "Sister late", level: "low" },
        { label: "Friend saved seat", level: "high" },
        { label: "Maths test", level: "low" },
        { label: "Picked last", level: "vlow" },
        { label: "Rosie kind", level: "high" },
      ];

      events.forEach((ev, i) => {
        const ey = topY + 0.55 + i * 0.3;
        const dotColor = ev.level === "high" ? C.SUCCESS : ev.level === "vlow" ? C.ALERT : C.ACCENT;
        slide.addShape("roundRect", {
          x: rx + 0.25, y: ey, w: 0.2, h: 0.2, rectRadius: 0.1,
          fill: { color: dotColor },
        });
        slide.addText(ev.label, {
          x: rx + 0.55, y: ey - 0.02, w: rw - 0.8, h: 0.24,
          fontSize: 10, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });
    }
  );

  // ================================================================
  // SLIDE 8: Casey's Day - We Do (Group Labelling)
  // ================================================================
  withReveal(
    () => contentSlide(
      pres,
      "We Do",
      C.SECONDARY,
      "Map Casey's Roller Coaster",
      [
        "With your group (2-4):",
        "1.  Look at the roller coaster diagram on your handout",
        "2.  For each event, name the emotion Casey felt",
        "3.  Use the Emotional Intensity Word List to find a precise word",
        "4.  Write the word on the roller coaster at the right height",
        "We will do the first two together, then you continue",
      ],
      NOTES_CASEY_WE_DO,
      FOOTER
    ),
    (slide) => {
      addTextOnShape(slide, "Example: Mushy biscuit = disgusted (medium)  |  Favourite breakfast = grateful (mild-medium)", {
        x: 0.5, y: 4.3, w: 9, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SECONDARY },
      }, { fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // ================================================================
  // SLIDE 9: CFU - Casey's Emotions
  // ================================================================
  withReveal(
    () => cfuSlide(
      pres,
      "Check",
      "Show Me Boards",
      "Whiteboards",
      "Casey was picked last for basketball and heard: 'Don't pick him. He'll make us lose.'\n\nWrite the EMOTION and an INTENSITY word.",
      NOTES_CFU_CASEY,
      FOOTER
    ),
    (slide) => {
      addTextOnShape(slide, "Strong answers: humiliated, crushed, devastated, deeply hurt, rejected", {
        x: 0.5, y: 4.3, w: 9, h: 0.55, rectRadius: 0.08,
        fill: { color: C.SUCCESS },
      }, { fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true });
    }
  );

  // ================================================================
  // SLIDE 10: You Do - Create Your Own Roller Coaster
  // ================================================================
  contentSlide(
    pres,
    "You Do",
    C.SUCCESS,
    "Create Your Roller Coaster Day",
    [
      "In your group, choose a topic:",
      "A.  A day in the life of a new Prep student",
      "B.  A day in the life of a new student joining YOUR class",
      "C.  The evening of a new babysitter",
      "Your story needs:",
      "At least 6 emotions (highs, lows, and in-betweens)",
      "Precise intensity words from the word list",
    ],
    NOTES_YOU_DO,
    FOOTER,
    (slide, layoutGuide) => {
      const rx = layoutGuide.rightX;
      const rw = layoutGuide.rightW;
      const topY = layoutGuide.panelTopPadded;

      // Steps card
      addCard(slide, rx, topY, rw, 2.8, { fill: C.BG_CARD, strip: C.SUCCESS });
      slide.addText("Steps", {
        x: rx + 0.15, y: topY + 0.1, w: rw - 0.3, h: 0.35,
        fontSize: 13, fontFace: FONT_H, color: C.SUCCESS, bold: true, margin: 0,
      });

      const steps = [
        "First: Choose your topic and list the events",
        "Next: Draw your roller coaster -- map highs, lows, in-betweens",
        "Then: Label each point with an emotion + intensity word",
        "Finally: Write up the story (each person writes their own)",
      ];
      steps.forEach((step, i) => {
        slide.addText(step, {
          x: rx + 0.2, y: topY + 0.55 + i * 0.5, w: rw - 0.4, h: 0.45,
          fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle", margin: 0,
        });
      });
    }
  );

  // ================================================================
  // SLIDE 11: Gallery Share
  // ================================================================
  pairShareSlide(
    pres,
    "Gallery Share",
    [
      "Share your roller coaster with the class",
      "What emotion word did this group use that really showed intensity?",
      "How does understanding emotional intensity help us make sure everyone belongs?",
    ],
    NOTES_SHARE,
    FOOTER
  );

  // ================================================================
  // SLIDE 12: Closing
  // ================================================================
  closingSlide(
    pres,
    "Turn and Talk: What is one thing you will do differently now that you understand emotional intensity?",
    [
      "I can explain that emotions range from mild to intense",
      "I can use precise words to describe different levels of emotional intensity",
      "I can identify events that might trigger stronger emotions in myself or others",
    ],
    NOTES_CLOSING
  );

  // ================================================================
  // SLIDE 13: Resources
  // ================================================================


  // ── Write PPTX ──
  const pptxPath = path.join(LESSON_FOLDER, PPTX_NAME);
  await pres.writeFile({ fileName: pptxPath });
  console.log("PPTX written to", pptxPath);

  // ═══════════════════════════════════════════════════════════════
  // Companion PDFs
  // ═══════════════════════════════════════════════════════════════

  // ── PDF 1: Emotions Intensity Cards ──
  await generateIntensityCards();

  // ── PDF 2: Roller Coaster Worksheet ──
  await generateRollerCoasterWorksheet();

  console.log("All resources written to:", RES_DIR);
}

// ═══════════════════════════════════════════════════════════════
// PDF Generation
// ═══════════════════════════════════════════════════════════════

async function generateIntensityCards() {
  const doc = createPdf({ title: "Session 1 Emotions Intensity Cards" });

  let y = addPdfHeader(doc, "Emotions Intensity Cards", {
    subtitle: "Print and cut - one set per class",
    color: hex(C.PRIMARY),
    lessonInfo: "Harmony Day | Emotional Intensity | Grade 5/6",
  });

  y = addTipBox(doc, "Cut along the dotted lines. Each card shows an emotion at two intensity levels. Use for the Statues Game and group discussions.", y, { color: hex(C.SECONDARY) });

  const pairs = [
    { mild: "Annoyed", intense: "Furious", family: "Anger" },
    { mild: "Nervous", intense: "Terrified", family: "Fear" },
    { mild: "Disappointed", intense: "Devastated", family: "Sadness" },
    { mild: "Worried", intense: "Panicked", family: "Anxiety" },
    { mild: "Embarrassed", intense: "Humiliated", family: "Shame" },
    { mild: "Irritated", intense: "Enraged", family: "Anger" },
    { mild: "Uneasy", intense: "Petrified", family: "Fear" },
    { mild: "Glum", intense: "Heartbroken", family: "Sadness" },
    { mild: "Unsure", intense: "Overwhelmed", family: "Confusion" },
    { mild: "Jealous", intense: "Consumed with Envy", family: "Jealousy" },
  ];

  y = addSectionHeading(doc, "Emotion Intensity Pairs (Cut Out Cards)", y, { color: hex(C.PRIMARY) });

  const margin = doc.page.margins.left || 50;
  const cardW = (doc.page.width - margin * 2 - 20) / 2;
  const cardH = 70;

  pairs.forEach((pair, i) => {
    if (y + cardH + 10 > doc.page.height - 80) {
      doc.addPage();
      y = 50;
    }

    const col = i % 2;
    const xPos = margin + col * (cardW + 20);

    if (col === 0 && i > 0) {
      // same row
    }

    // Card border
    doc.rect(xPos, y, cardW, cardH).dash(3, { space: 3 }).stroke(hex(C.MUTED));

    // Family label
    doc.fontSize(9).font("Sans-Italic").fillColor(hex(C.MUTED))
      .text(pair.family, xPos + 5, y + 5, { width: cardW - 10 });

    // Mild word
    doc.fontSize(14).font("Sans").fillColor(hex(C.CHARCOAL))
      .text(pair.mild, xPos + 5, y + 20, { width: cardW - 10 });

    // Arrow
    doc.fontSize(12).font("Sans").fillColor(hex(C.MUTED))
      .text("-->", xPos + 5, y + 38, { width: cardW - 10, align: "center" });

    // Intense word
    doc.fontSize(14).font("Sans-Bold").fillColor(hex(C.ALERT))
      .text(pair.intense, xPos + 5, y + 50, { width: cardW - 10 });

    if (col === 1 || i === pairs.length - 1) {
      y += cardH + 10;
    }
  });

  addPdfFooter(doc, "Harmony Day | Emotional Intensity | Grade 5/6");

  const filePath = path.join(LESSON_FOLDER, INTENSITY_CARDS_RESOURCE.fileName);
  await writePdf(doc, filePath);
  console.log("PDF written:", filePath);
}

async function generateRollerCoasterWorksheet() {
  const doc = createPdf({ title: "Session 1 Roller Coaster Worksheet" });

  let y = addPdfHeader(doc, "Casey's Roller Coaster Day", {
    subtitle: "Emotional Intensity Worksheet",
    color: hex(C.PRIMARY),
    lessonInfo: "Harmony Day | Emotional Intensity | Grade 5/6",
  });

  // Casey's story extract
  y = addSectionHeading(doc, "Casey's Morning", y, { color: hex(C.PRIMARY) });

  y = addBodyText(doc, "The day started like any other. I woke up and climbed out of bed. Eww, my foot landed right on something slimy and mushy -- my baby sister's half chewed biscuits. 'Dadddddd!' 'Just get ready for school,' was all he said.", y);

  y = addBodyText(doc, "Walking out into the kitchen I noticed he had made my favourite breakfast: egg and bacon muffin. 'Thanks Dad,' I said. He smiled at me and ruffled my hair.", y);

  y = addBodyText(doc, "I ate and went quickly to get dressed for school to have time to play with my friends before the bell. But I had to wait ages until my big sister was finally ready to leave for school. It was three minutes to nine! We were going to be late. My teacher gets so annoyed when we are late.", y);

  y = addBodyText(doc, "When I arrived in class, after sneaking back from the office sign-in, I saw that my best friend had saved me a seat. Then I discovered we were doing a Maths test on decimals -- and last week I got 7 out of 10 wrong!", y);

  y = addBodyText(doc, "When the recess bell rang, Jordan suggested, 'Let's play basketball'. As they picked teams I stood and waited and waited and waited. I was the last one chosen. I could hear Lou saying, 'Don't pick him. He'll make us lose.' I turned around and walked off. As I trudged over to the taps, Rosie called, 'Why don't we just sit and talk? I don't even like basketball.'", y);

  // Roller coaster diagram area
  y = addSectionHeading(doc, "Casey's Roller Coaster", y, { color: hex(C.SECONDARY) });

  y = addTipBox(doc, "Map Casey's emotions on the roller coaster below. Label each event with an emotion word from the Intensity Word List. Place highs at the top and lows at the bottom.", y, { color: hex(C.SECONDARY) });

  // Draw a simple roller coaster grid
  const margin = doc.page.margins.left || 50;
  const gridW = doc.page.width - margin * 2;
  const gridH = 120;

  if (y + gridH + 20 > doc.page.height - 80) {
    doc.addPage();
    y = 50;
  }

  // Y-axis labels
  doc.fontSize(8).font("Sans").fillColor(hex(C.CHARCOAL));
  doc.text("HIGH", margin - 5, y, { width: 35, align: "right" });
  doc.text("MED", margin - 5, y + gridH / 2 - 5, { width: 35, align: "right" });
  doc.text("LOW", margin - 5, y + gridH - 10, { width: 35, align: "right" });

  // Grid box
  const gridX = margin + 40;
  const gridActualW = gridW - 40;
  doc.rect(gridX, y, gridActualW, gridH).stroke(hex(C.MUTED));

  // Horizontal midline
  doc.moveTo(gridX, y + gridH / 2).lineTo(gridX + gridActualW, y + gridH / 2)
    .dash(2, { space: 3 }).stroke(hex(C.MUTED));

  // Event markers along the x-axis
  const events = ["Mushy\nbiscuit", "Favourite\nbrekky", "Sister\nlate", "Friend\nsaved seat", "Maths\ntest", "Picked\nlast", "Rosie's\nkindness"];
  const evSpacing = gridActualW / (events.length + 1);
  doc.undash();
  events.forEach((ev, i) => {
    const ex = gridX + evSpacing * (i + 1);
    doc.fontSize(7).font("Sans").fillColor(hex(C.CHARCOAL))
      .text(ev, ex - 25, y + gridH + 3, { width: 50, align: "center" });
    // Dotted vertical guide
    doc.moveTo(ex, y).lineTo(ex, y + gridH).dash(1, { space: 3 }).stroke(hex(C.MUTED));
    doc.undash();
  });

  y += gridH + 35;

  // Emotional Intensity Word List
  if (y + 180 > doc.page.height - 80) {
    doc.addPage();
    y = 50;
  }

  y = addSectionHeading(doc, "Emotional Intensity Word List", y, { color: hex(C.ACCENT) });

  const wordTable = [
    { family: "Sadness", mild: "Disappointed, Let down, Blue", medium: "Sad, Upset, Unhappy", intense: "Devastated, Heartbroken, Grief-stricken" },
    { family: "Anger", mild: "Annoyed, Irritated, Bothered", medium: "Angry, Cross, Frustrated", intense: "Furious, Enraged, Livid" },
    { family: "Fear", mild: "Nervous, Uneasy, Unsettled", medium: "Scared, Worried, Anxious", intense: "Terrified, Petrified, Panicked" },
    { family: "Happiness", mild: "Pleased, Content, Glad", medium: "Happy, Cheerful, Delighted", intense: "Ecstatic, Overjoyed, Elated" },
    { family: "Shame", mild: "Awkward, Self-conscious", medium: "Embarrassed, Ashamed", intense: "Humiliated, Mortified, Disgraced" },
    { family: "Surprise", mild: "Curious, Puzzled", medium: "Surprised, Shocked", intense: "Stunned, Astonished, Flabbergasted" },
  ];

  // Table header
  const colWidths = [gridW * 0.15, gridW * 0.28, gridW * 0.28, gridW * 0.29];
  const headerY = y;
  doc.fontSize(9).font("Sans-Bold").fillColor(hex(C.WHITE));

  // Header background
  doc.rect(margin, headerY, gridW, 18).fill(hex(C.PRIMARY));
  doc.fillColor(hex(C.WHITE));
  doc.text("Family", margin + 3, headerY + 4, { width: colWidths[0] - 6 });
  doc.text("Mild", margin + colWidths[0] + 3, headerY + 4, { width: colWidths[1] - 6 });
  doc.text("Medium", margin + colWidths[0] + colWidths[1] + 3, headerY + 4, { width: colWidths[2] - 6 });
  doc.text("Intense", margin + colWidths[0] + colWidths[1] + colWidths[2] + 3, headerY + 4, { width: colWidths[3] - 6 });

  y = headerY + 18;

  wordTable.forEach((row, i) => {
    if (y + 22 > doc.page.height - 80) {
      doc.addPage();
      y = 50;
    }
    const rowH = 22;
    const bgColor = i % 2 === 0 ? "F5F5F5" : "FFFFFF";
    doc.rect(margin, y, gridW, rowH).fill(bgColor);
    doc.rect(margin, y, gridW, rowH).stroke(hex(C.MUTED));

    doc.fontSize(8).font("Sans-Bold").fillColor(hex(C.PRIMARY))
      .text(row.family, margin + 3, y + 5, { width: colWidths[0] - 6 });
    doc.fontSize(8).font("Sans").fillColor(hex(C.CHARCOAL))
      .text(row.mild, margin + colWidths[0] + 3, y + 5, { width: colWidths[1] - 6 });
    doc.text(row.medium, margin + colWidths[0] + colWidths[1] + 3, y + 5, { width: colWidths[2] - 6 });
    doc.font("Sans-Bold").fillColor(hex(C.ALERT))
      .text(row.intense, margin + colWidths[0] + colWidths[1] + colWidths[2] + 3, y + 5, { width: colWidths[3] - 6 });

    y += rowH;
  });

  // Your Roller Coaster section
  doc.addPage();
  y = 50;

  y = addSectionHeading(doc, "Your Roller Coaster Day", y, { color: hex(C.SUCCESS) });

  y = addTipBox(doc, "Choose a topic: A) A day in the life of a new Prep student, B) A day in the life of a new student joining your class, or C) The evening of a new babysitter.", y, { color: hex(C.SUCCESS) });

  y = addWriteLine(doc, "Topic chosen", y);
  y = addWriteLine(doc, "Group members", y);

  y += 5;
  y = addSectionHeading(doc, "Plan Your Events", y, { color: hex(C.SECONDARY) });

  // Events planning table
  for (let i = 1; i <= 6; i++) {
    if (y + 30 > doc.page.height - 80) {
      doc.addPage();
      y = 50;
    }
    y = addWriteLine(doc, "Event " + i, y);
    y = addWriteLine(doc, "Emotion + Intensity", y);
    y += 5;
  }

  // Blank roller coaster
  if (y + gridH + 30 > doc.page.height - 80) {
    doc.addPage();
    y = 50;
  }

  y = addSectionHeading(doc, "Draw Your Roller Coaster", y, { color: hex(C.PRIMARY) });

  // Y-axis
  doc.fontSize(8).font("Sans").fillColor(hex(C.CHARCOAL));
  doc.text("HIGH", margin - 5, y, { width: 35, align: "right" });
  doc.text("MED", margin - 5, y + gridH / 2 - 5, { width: 35, align: "right" });
  doc.text("LOW", margin - 5, y + gridH - 10, { width: 35, align: "right" });

  doc.rect(gridX, y, gridActualW, gridH).stroke(hex(C.MUTED));
  doc.moveTo(gridX, y + gridH / 2).lineTo(gridX + gridActualW, y + gridH / 2)
    .dash(2, { space: 3 }).stroke(hex(C.MUTED));
  doc.undash();

  y += gridH + 15;

  // Writing space
  if (y + 20 > doc.page.height - 80) {
    doc.addPage();
    y = 50;
  }
  y = addSectionHeading(doc, "Write Your Story", y, { color: hex(C.PRIMARY) });
  y = addLinedArea(doc, y, 12);

  addPdfFooter(doc, "Harmony Day | Emotional Intensity | Grade 5/6");

  const filePath = path.join(LESSON_FOLDER, ROLLER_COASTER_RESOURCE.fileName);
  await writePdf(doc, filePath);
  console.log("PDF written:", filePath);
}

// ── Run ──
build().catch((err) => {
  console.error("Build failed:", err);
  process.exit(1);
});
