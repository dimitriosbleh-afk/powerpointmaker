"use strict";

// Inquiry - Shaping Australia, Lesson 4 (Term 2, Week 4)
// Year 5/6 | Why people migrated to Australia after 1788.
// Continues from Lesson 3 (Captain Cook and European Settlement). Students
// study who came in four waves (Gold Rush Chinese, Post-WWII Europeans,
// Vietnamese refugees, modern skilled migration), learn the push factor /
// pull factor framework, complete migration profile cards, then write a
// push/pull explanation paragraph for one cultural group.
//
// Sensitive content. Acknowledgement of Country is read carefully because
// every migrant came to land that already belonged to First Nations
// peoples. The White Australia Policy and discrimination are named, not
// glossed over.

const path = require("path");
const fs = require("fs");
const PptxGenJS = require("pptxgenjs");
const { createTheme, weekToVariant } = require("../themes/factory");
const {
  createPdf, writePdf, addPdfHeader, addBodyText, addTipBox,
  addPdfFooter, addResourceSlide, makeSessionResource,
  getSessionResourceFolder, PAGE, hex,
} = require("../themes/pdf_helpers");

// -- Theme --
// Unit cohesion: Lessons 1, 2 and 3 used variant 0 (Explorer - olive).
// All lessons in the unit MUST share a variant. weekToVariant(1) -> 0.
const T = createTheme("inquiry", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, closingSlide,
  exitTicketSlide, withReveal,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  addInstructionCard,
  CONTENT_TOP, SAFE_BOTTOM,
} = T;

// -- Output paths --
const UNIT = "Shaping_Australia";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "Lesson 4 - Why People Came to Australia.pptx";
const FOOTER = "Inquiry | Year 5/6 | Shaping Australia | Lesson 4";
const SESSION = 4;
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

// -- Resources --
const PROFILE_CARDS = makeSessionResource(
  SESSION,
  "Migration Profile Cards",
  "Four migration profile cards: Chinese (1850s Gold Rush), Italian (post-WWII), Vietnamese (1970s refugees), Indian (modern skilled migration). Each card has space to record when, why and one fact."
);
const PUSH_PULL_SCAFFOLD = makeSessionResource(
  SESSION,
  "Push and Pull Explanation Scaffold",
  "Structured A4 scaffold to explain why ONE cultural group migrated to Australia, using push factors and pull factors with sentence starters."
);
const RESOURCE_ITEMS = [PROFILE_CARDS, PUSH_PULL_SCAFFOLD];

fs.mkdirSync(RES_DIR, { recursive: true });

// ===============================================================
// Teacher Notes
// ===============================================================

const NOTES_TITLE = [
  "SAY:",
  "- Welcome back to Shaping Australia",
  "- Last lesson we learnt about Captain Cook, the First Fleet and the impact of British settlement",
  "- Today we ask: after 1788, who else came to this land, and why did they come?",
  "- We will learn about four big migration stories that helped shape modern Australia",
  "",
  "DO:",
  "- Display this slide as students enter",
  "- Have the class KWL chart from Lessons 1, 2 and 3 visible",
  "- Have the Migration Profile Cards and Push and Pull Scaffold printed - one of each per student",
  "",
  "TEACHER NOTES:",
  "Lesson 4 builds on the settlement history from Lesson 3 and broadens it to the many migrant groups who followed. Some sensitive content - the White Australia Policy and refugee experiences are named, not avoided.",
  "",
  "WATCH FOR:",
  "- Students from migrant families - welcome stories they want to share, but never put a student on the spot",
  "",
  "[Inquiry: Title | VTLM 2.0: Establishing Purpose]",
].join("\n");

const NOTES_RESOURCES = [
  "SAY:",
  "- Two printed resources today: a set of migration profile cards and a push and pull explanation scaffold",
  "- We also bring back the class KWL chart from Lessons 1, 2 and 3",
  "",
  "DO:",
  "- Print the Migration Profile Cards - one per student (A4)",
  "- Print the Push and Pull Explanation Scaffold - one per student (A4)",
  "- Have the Lesson 1 to 3 class KWL chart visible",
  "- Mini-whiteboards ready for the CFU check",
  "- Pencils, erasers, and the student inquiry books for additional notes",
  "",
  "TEACHER NOTES:",
  "Both resources are designed for a single 60-minute lesson. The profile cards are the recording tool for SC1 and SC2; the scaffold structures the explanation paragraph for SC3. Both are referred to on the You Do slides.",
  "",
  "WATCH FOR:",
  "- Print single-sided so students have full writing room for the profile cards",
  "",
  "[Inquiry: Resources | VTLM 2.0: Preparation]",
].join("\n");

const NOTES_AOC = [
  "SAY:",
  "- Before we begin, we pause and acknowledge",
  "- We acknowledge the Traditional Owners of the land on which we are learning today",
  "- We pay our respects to Elders past and present, and to all Aboriginal and Torres Strait Islander peoples",
  "- We thank them for caring for this land for tens of thousands of years",
  "- Today we talk about the many people who came to this land. They all came to land that already belonged to First Nations peoples",
  "",
  "DO:",
  "- Read the Acknowledgement aloud, slowly. The class stands quietly",
  "- Replace the bracketed nation name with the local Country if your school's standard wording uses it",
  "- Allow a few seconds of silence after the Acknowledgement before moving on",
  "",
  "TEACHER NOTES:",
  "The Acknowledgement matters today especially because we are looking at later migration to Australia. Every migrant group, including British settlers, arrived on First Nations Country.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: A formal recognition of First Nations Australians as the original and continuing custodians of this land.",
  "- Framing language: 'We acknowledge. We pay our respects. We thank.'",
  "- Watch for: Students who treat it as routine - reset the tone with: 'This matters today especially. We do this properly'.",
  "- Protocol: If your school has a Welcome to Country recorded by a local Elder, you may use that instead.",
  "",
  "WATCH FOR:",
  "- Quiet, respectful student behaviour - praise it briefly afterwards",
  "",
  "[Inquiry: Acknowledgement | VTLM 2.0: Cultural Protocol]",
].join("\n");

const NOTES_LAUNCH = [
  "SAY:",
  "- Last lesson we learnt that the British settled here from 1788 with the First Fleet",
  "- Today we ask: after 1788, who else came to live in Australia? And why did they come?",
  "- Hold this question in your head: are the reasons people came the same, or are they different?",
  "",
  "DO:",
  "- Project the class KWL chart beside this slide",
  "- Read the slide aloud",
  "- Quick partner share: 'Name one country your family or someone you know came from. Why did they come?'",
  "- Cold call 2-3 students to share. Add new countries to the class chart",
  "",
  "TEACHER NOTES:",
  "The launch reactivates Lesson 3 (the First Fleet started British settlement) and points it at today's focus. Keep it brisk. Some students will share family stories - welcome them, but do not require it.",
  "",
  "WATCH FOR:",
  "- Students who say 'I have always lived here' - acknowledge: 'That is fine. Many Australians have. Today we learn about families who came later'",
  "- Students who confuse 'where I was born' with 'where my family came from' - either answer is welcome",
  "",
  "[Inquiry: Launch | VTLM 2.0: Activating Prior Knowledge]",
].join("\n");

const NOTES_VOCAB = [
  "SAY:",
  "- Four inquiry words for today",
  "- Read each word and meaning with me",
  "- 'Push factor' and 'pull factor' are the two big ones - we will use them all lesson",
  "",
  "DO:",
  "- Read each word aloud. Have students repeat 'push factor' and 'pull factor' after you",
  "- Point to the small icons as you go",
  "",
  "CFU CHECKPOINT:",
  "Technique: Finger Voting",
  "Script:",
  "- Say: Hold up 1 finger if a war is a PUSH factor. Hold up 2 if it is a PULL factor",
  "- Scan for: most students showing 1 (war pushes people away from home)",
  "PROCEED: If 80% show 1, move on.",
  "PIVOT: If many show 2, say: 'A push factor pushes you AWAY from your home. A pull factor pulls you TOWARDS a new place. War pushes people away'. Re-check with: 'Is finding a good job a push or a pull?' (answer 2 - pulls people toward a country).",
  "",
  "TEACHER NOTES:",
  "These four words are the inquiry vocabulary for the rest of the lesson. Push factor and pull factor are the conceptual heart of the lesson - everything else hangs on them.",
  "",
  "WATCH FOR:",
  "- Students mixing up push and pull - re-anchor with a hand action: hands pushing forward for 'push', hands pulling in for 'pull'",
  "",
  "[Inquiry: Vocabulary | VTLM 2.0: Establishing Knowledge]",
].join("\n");

const NOTES_LI = [
  "SAY:",
  "- Read the Learning Intention aloud",
  "- Read each Success Criterion. We will come back to these at the end to check how we went",
  "",
  "DO:",
  "- Point to each I can statement as you read it",
  "- Leave the slide visible for about 20 seconds",
  "",
  "TEACHER NOTES:",
  "Internal tier mapping (do not share with students). SC1 - everyone can list cultural groups. SC2 - most students explain reasons. SC3 - students sort reasons into push factors and pull factors.",
  "",
  "WATCH FOR:",
  "- Students who think 'multicultural' just means 'many people' - clarify: 'It means many cultures living together. Different languages, foods, holidays, beliefs - all in one country'",
  "",
  "[Inquiry: LI/SC | VTLM 2.0: Learning Intentions]",
].join("\n");

const NOTES_IDO_WAVES = [
  "SAY:",
  "- After 1788, many more people came to live in Australia. We will look at four big stories",
  "- One - the 1850s Gold Rush. Gold was found in Victoria and New South Wales. People came from China, the United States, Britain, Germany and other countries to look for gold",
  "- Two - after World War Two, from 1945 to about 1970. Europe was destroyed by the war. Australia needed workers. About 1.8 million migrants came in this period - from Italy, Greece, the Netherlands, Germany, Malta, Poland, Hungary and many more places",
  "- Three - from the late 1970s. The Vietnam War ended and many Vietnamese families fled their country. Around 90,000 came to Australia, many by boat, as refugees",
  "- Four - the last 30 years. Skilled migration. Many people have come from India, China, the Philippines, the United Kingdom, South Africa and many other countries to work and study",
  "- One important thing - until 1973, the Australian government had a policy called the White Australia Policy. It made it very hard for non-white people to come here. That policy was wrong and was ended",
  "",
  "DO:",
  "- Read the slide content slowly",
  "- Point to each wave on the timeline as you talk about it",
  "- Stop after wave four and ask: 'Why do you think Australia is called multicultural today?'",
  "- Take 2-3 student responses",
  "",
  "TEACHER NOTES:",
  "These four waves are real history and well documented. Numbers used (1.8 million post-war migrants, 90,000 Vietnamese) are widely accepted estimates from the Department of Home Affairs and the Australian Bureau of Statistics. The White Australia Policy is named because it is a key fact about who could and could not come.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: The history of migration to Australia includes a long period of legal racial discrimination (the White Australia Policy, 1901 to 1973).",
  "- Framing language: 'For 72 years Australian law made it harder for non-white people to come here. That policy was wrong. It ended in 1973'.",
  "- Watch for: Students who say 'why do we let people in now?' - redirect: 'Today we ask why people came. We do not judge whether they should have'.",
  "- Protocol: Keep the focus on history and facts. Modern migration debates belong in a different lesson if your school chooses to teach them.",
  "",
  "WATCH FOR:",
  "- Students who ask about more recent groups (Sudanese, Syrian, Afghan) - welcome the question: 'Yes - many groups have come since 2000 as refugees. We use Vietnam as our example because it is the biggest single wave of refugees in Australian history'",
  "- Students who say 'no one came after the British' - re-anchor with the Gold Rush timeline point",
  "",
  "SOURCES:",
  "- Migration figures: Department of Home Affairs, Australian Bureau of Statistics, National Museum of Australia",
  "- White Australia Policy: National Museum of Australia, Australian Parliament House education resources",
  "",
  "[Inquiry: I Do (Waves) | VTLM 2.0: Direct Instruction]",
].join("\n");

const NOTES_IDO_PUSH_PULL = [
  "SAY:",
  "- Why do people leave one country and move to another? Historians use two big ideas",
  "- One - push factors. These are reasons that push people AWAY from their home country",
  "- Push factors include war, persecution, poverty, lack of work, famine, and disasters",
  "- Two - pull factors. These are reasons that pull people TOWARDS a new country",
  "- Pull factors include jobs, gold, safety, family already in the new country, education, and freedom",
  "- Almost every migration story uses BOTH. People are pushed by something AND pulled towards something",
  "",
  "DO:",
  "- Read each card aloud, pausing on each side",
  "- Use hand actions: hands pushing forward for 'push', hands pulling in for 'pull'",
  "- Use the Italian example at the bottom to show BOTH working together",
  "- Ask: 'For Italian families after WWII - what was the push? What was the pull?'",
  "- Take 1-2 student responses",
  "",
  "TEACHER NOTES:",
  "The push/pull framework is the conceptual heart of the lesson. Students who hold this idea can explain almost any migration story. The Italian post-WWII example is the warmest worked example because both push (poverty and war damage) and pull (jobs and family chain migration) were strong.",
  "",
  "WATCH FOR:",
  "- Students who think push and pull are opposites - clarify: 'They work together, not against each other. Most migrants have both'",
  "- Students who say 'people just move because they want to' - prompt: 'What makes them want to? That is the push or the pull'",
  "",
  "[Inquiry: I Do (Push and Pull) | VTLM 2.0: Conceptual Knowledge]",
].join("\n");

const CFU_Q_TEXT = "A family flees a war in their home country and comes to Australia where their cousin already lives. Name ONE push factor and ONE pull factor in their story.";

const NOTES_CFU_Q = [
  "SAY:",
  "- Mini-whiteboards out",
  "- Write two short answers: one push factor and one pull factor from the story",
  "- One minute",
  "",
  "DO:",
  "- Set a 1-minute timer",
  "- Scan whiteboards as students hold them up",
  "",
  "CFU CHECKPOINT:",
  "Technique: Show Me Boards",
  "Script:",
  "- Say: Boards up on three. One, two, three",
  "- Scan for: a clear PUSH (war, fleeing, danger) AND a clear PULL (cousin in Australia, family, safety)",
  "PROCEED: If 80% show one of each, reveal and move to the profile cards.",
  "PIVOT: If many show two pushes or two pulls, say: 'You named good reasons, but check: a push is what they leave BEHIND. A pull is what brings them HERE. Sort yours into push and pull'. Re-check.",
  "",
  "TEACHER NOTES:",
  "This hinge checks the threshold concept. Without sorting reasons into push and pull, students cannot complete the explanation task that comes later.",
  "",
  "WATCH FOR:",
  "- Students who write 'war' as the pull - prompt: 'Did the war bring them HERE, or did they run AWAY from it?'",
  "- Students who write only one of the two - prompt: 'You have one. Now find the other half'",
  "",
  "[Inquiry: CFU | VTLM 2.0: Checking for Understanding]",
].join("\n");

const NOTES_CFU_A = [
  "SAY:",
  "- Push factor: a war in their home country - that pushed them away",
  "- Pull factor: a cousin already living in Australia - that pulled them towards us",
  "- Almost every migration story has both halves. Hold this idea as we move on",
  "",
  "DO:",
  "- Click to reveal the answer",
  "- Acknowledge close answers",
  "- Move on promptly to the profile cards",
  "",
  "TEACHER NOTES:",
  "The reveal anchors the rest of the lesson. The profile cards and the explanation paragraph both rely on students sorting reasons into push and pull.",
  "",
  "WATCH FOR:",
  "- Students still updating their boards after the reveal - that is fine, that is learning",
  "",
  "[Inquiry: CFU Reveal | VTLM 2.0: Feedback]",
].join("\n");

const NOTES_WE_DO_PROFILE = [
  "SAY:",
  "- We are going to record what we know about four migrant groups - Chinese, Italian, Vietnamese and Indian",
  "- A profile card has three boxes: WHEN they came, WHY they came, and ONE fact",
  "- Watch me - I will model the Italian card together",
  "",
  "DO:",
  "- Project the Italian profile card on the slide",
  "- Think aloud: 'WHEN - after World War Two, mostly between 1945 and 1970. I will write 'After WWII, 1945 to 1970'",
  "- Think aloud: 'WHY - Italy was destroyed by the war. Lots of poverty. Australia needed workers. I'll write 'War damage and poverty (push). Jobs in Australia (pull)'",
  "- Think aloud: 'One fact - many Italian men worked on the Snowy Mountains Scheme building dams. I'll write that'",
  "- Pause: 'Notice the push and pull labels in the WHY box. We will do that for every card'",
  "",
  "TEACHER NOTES:",
  "This We Do is short - one full card modelled. The aim is to show the THINKING for the WHY box, especially how to label push factors and pull factors. Students will fill the other three cards in the You Do.",
  "",
  "WATCH FOR:",
  "- Students who write reasons without labelling push or pull - prompt them to add the label as they go",
  "- Students who try to fill all four cards during the model - 'Watch me first, then you do yours'",
  "",
  "[Inquiry: We Do (Profile) | VTLM 2.0: Modelling]",
].join("\n");

const NOTES_YOU_DO_PROFILE = [
  "SAY:",
  "- Take your Migration Profile Cards sheet",
  "- Four cards: Chinese, Italian (already modelled - you can keep that one), Vietnamese, Indian",
  "- For each card, fill in WHEN, WHY (with push and pull labels), and ONE fact",
  "- Use the information bank on the back of the sheet if you need help",
  "- You have 12 minutes - work with your shoulder partner",
  "",
  "DO:",
  "- Distribute the Migration Profile Cards - one per student",
  "- Set a 12-minute timer",
  "- Circulate. Prompt quiet pairs with: 'What pushed Chinese families to leave China in the 1850s? What pulled them here?'",
  "- Encourage students to label every reason as push or pull - that is the key extension",
  "- Photograph filled cards before they leave the room",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Complete only two cards instead of all four (Italian is already modelled). Use the information bank to circle reasons before writing them",
  "- Extra Notes: A pair can sit close to the teacher group for one-on-one prompting if needed.",
  "EXTENDING PROMPT:",
  "- Task: After the four cards, write one sentence comparing two groups - 'Both ___ and ___ had ___ as a push factor, but their pull factors were different because ___'.",
  "",
  "TEACHER NOTES:",
  "The profile cards are the recording tool for SC1 (list cultures) and SC2 (explain why). Students who complete all four cards with push and pull labels are demonstrating SC3 as well. Look for accurate use of the information bank - students do not need to invent facts.",
  "",
  "WATCH FOR:",
  "- Pairs missing the push and pull labels - prompt: 'Add the label in brackets after each reason'",
  "- Pairs copying word-for-word from the bank - prompt: 'Choose the reason that matters most, then explain it in your own words'",
  "",
  "[Inquiry: You Do (Profile) | VTLM 2.0: Application]",
].join("\n");

const NOTES_WE_DO_EXPLAIN = [
  "SAY:",
  "- Now we use what we know to write one good paragraph about ONE group",
  "- We will model a Vietnamese explanation - watch how I use my profile card",
  "- The scaffold has three parts: introduce the group, explain the push, explain the pull, finish with a fact",
  "- Watch me model the introduction and the push paragraph",
  "",
  "DO:",
  "- Project the explanation scaffold on the slide",
  "- Model only the introduction and the push paragraph. Read aloud as you write",
  "- 'Many Vietnamese families came to Australia in the late 1970s and 1980s'",
  "- 'They were pushed out of their home country by the Vietnam War. After the war ended, many people were in danger. Some were sent to prison camps. They had to leave to be safe'",
  "- Think aloud: 'I used my profile card to find the push. I used the word 'pushed' so I show I am using the framework'",
  "- Stop modelling there - students take it from this point",
  "",
  "TEACHER NOTES:",
  "The model deliberately stops after the push paragraph so students have a clear next step. The scaffold provides sentence starters; the conversation about WHY they migrated is where the learning happens. Use Vietnamese as the model because it has the clearest push factor in the lesson.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: The Vietnamese refugee story includes war trauma, separation from family, dangerous boat journeys and many deaths at sea.",
  "- Framing language: 'This is a serious story. We tell it with respect. We use words like 'fled', 'refugee' and 'sought safety' - not 'illegal' or 'invader'.",
  "- Watch for: Students who use modern political language - redirect to historical accuracy: 'These people were refugees by international law. They were welcomed by Australia at the time'.",
  "- Protocol: If any student has a family connection to refugee experience, follow their lead. Never put a student on the spot.",
  "",
  "WATCH FOR:",
  "- Students who copy the model word-for-word - prompt: 'Choose your own group, not Vietnamese'",
  "- Students who write only one factor - prompt: 'You need a push AND a pull. Find the other one'",
  "",
  "[Inquiry: We Do (Explain) | VTLM 2.0: Modelling]",
].join("\n");

const NOTES_YOU_DO_EXPLAIN = [
  "SAY:",
  "- Take your Push and Pull Explanation Scaffold",
  "- Choose ONE group from your profile cards - Chinese, Italian, Vietnamese or Indian",
  "- Use the four paragraph starters - they help you stay on track",
  "- P1: Introduce the group and when they came",
  "- P2: Explain the PUSH factors - what pushed them away from home",
  "- P3: Explain the PULL factors - what pulled them to Australia",
  "- P4: One fact about the group's contribution to Australia today",
  "- You have 15 minutes. Quality over length",
  "",
  "DO:",
  "- Distribute the Push and Pull Explanation Scaffold - one per student",
  "- Set a 15-minute timer",
  "- Circulate. Prompt quiet students with: 'Look at your profile card. What did you write in the WHY box?'",
  "- Look for: clear use of 'push factor' and 'pull factor' as terms",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Fill in only the four sentence starters at the top of each paragraph. One sentence per paragraph is enough",
  "- Extra Notes: Sit close to the teacher group for one-on-one prompting if needed.",
  "EXTENDING PROMPT:",
  "- Task: After the four paragraphs, add a fifth paragraph comparing your group with one other group - 'Both ___ and ___ had ___ as a push factor, but ___'",
  "",
  "TEACHER NOTES:",
  "This task assesses SC2 (explain why) and SC3 (use push and pull factors). Look for accurate use of content from the profile cards - that is the evidence of learning. Sophistication of language is a bonus, not the target.",
  "",
  "WATCH FOR:",
  "- Students stuck on the introduction - prompt: 'Just write the first sentence and keep going'",
  "- Students writing very short answers - prompt: 'Add one more fact from your profile card'",
  "",
  "[Inquiry: You Do (Explain) | VTLM 2.0: Application]",
].join("\n");

const NOTES_EXIT = [
  "SAY:",
  "- Sticky note - one sentence before you leave",
  "- Finish this sentence: 'One cultural group that came to Australia was ___. They came because of the push factor ___ and the pull factor ___.'",
  "- Stick it on the class chart on your way out",
  "",
  "DO:",
  "- Hand each student a sticky note (or small slip of paper)",
  "- Allow 3 minutes",
  "- Collect by sticking on the class chart - L column",
  "",
  "TEACHER NOTES:",
  "This exit ticket assesses SC2 directly - the reasons different cultures migrated to Australia - and uses the push and pull vocabulary from SC3. Look for accurate cultures, accurate pushes and accurate pulls.",
  "",
  "WATCH FOR:",
  "- Students who write only the group - prompt: 'Add the push and the pull'",
  "- Students who name a group but mix up the push and pull - acknowledge the group, then prompt: 'Sort: which was the push, which was the pull?'",
  "",
  "[Inquiry: Exit Ticket | VTLM 2.0: Evidence of Learning]",
].join("\n");

const NOTES_CLOSING = [
  "SAY:",
  "- Let's come back to our success criteria. Thumbs up, sideways, or down for each",
  "- I can list different cultures that migrated to Australia",
  "- I can explain why different cultures migrated to Australia",
  "- I can sort reasons into push factors and pull factors",
  "- Reflection: which migration story stood out to you, and why?",
  "",
  "DO:",
  "- Read each I can statement and pause for thumbs",
  "- Acknowledge that today's lesson covered many cultures and many reasons",
  "- Tell students: next lesson we will look at what life was like for migrants when they arrived",
  "",
  "TEACHER NOTES:",
  "Close the lesson warmly. Many students will be thinking about their own family stories. A respectful close lets them carry the learning out of the room.",
  "",
  "WATCH FOR:",
  "- Students who show thumbs down on SC3 - flag for a small-group push/pull review next lesson",
  "- Students who linger with family stories - give them time",
  "",
  "[Inquiry: Closing | VTLM 2.0: Review and Reflect]",
].join("\n");

// ===============================================================
// Build function
// ===============================================================

async function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.title = "Shaping Australia - Lesson 4 - Why People Came to Australia";
  pres.author = "Year 5/6 Inquiry";

  // -- Slide 1: Title --
  titleSlide(
    pres,
    "Why People Came to Australia",
    "Different cultures, different reasons",
    "Year 5/6 Inquiry  |  Term 2  |  Lesson 4",
    NOTES_TITLE
  );

  // -- Slide 2: Teacher Resources --
  addResourceSlide(pres, {
    resources: RESOURCE_ITEMS,
    studentTools: [
      "Mini-whiteboards (one per student)",
      "Sticky notes for exit ticket",
      "Inquiry book or pencil for additional notes",
    ],
    boardSetup: [
      "Class KWL chart from Lessons 1 to 3 visible",
      "Migration Profile Cards and Push and Pull Scaffold printed - one of each per student",
    ],
  }, T, FOOTER, NOTES_RESOURCES);

  // -- Slide 3: Acknowledgement of Country --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Acknowledgement", { color: C.PRIMARY, w: 2.2 });
    addTitle(s, "Acknowledgement of Country");

    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;
    addCard(s, 0.5, cardY, 9, cardH, { strip: C.PRIMARY, fill: C.WHITE });

    s.addText("We acknowledge the Traditional Owners of the land on which we are learning today.", {
      x: 0.85, y: cardY + 0.35, w: 8.3, h: 0.75,
      fontSize: 20, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
    s.addText("We pay our respects to Elders past and present, and to all Aboriginal and Torres Strait Islander peoples.", {
      x: 0.85, y: cardY + 1.20, w: 8.3, h: 0.75,
      fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
    s.addText("We thank them for caring for this land for tens of thousands of years.", {
      x: 0.85, y: cardY + 2.05, w: 8.3, h: 0.55,
      fontSize: 16, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
    s.addText("Every migrant who came after 1788 came to land that already belonged to First Nations peoples.", {
      x: 0.85, y: cardY + 2.75, w: 8.3, h: 0.75,
      fontSize: 14, fontFace: FONT_B, color: C.PRIMARY, bold: true,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_AOC);
  }

  // -- Slide 4: Launch --
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "Launch", { color: C.SECONDARY, w: 1.4 });
    addTitle(s, "From Lesson 3 to today");

    // Left: prior knowledge card
    addInstructionCard(s, [
      { role: "header", text: "We already know" },
      { role: "body", text: "Captain Cook claimed the east coast for Britain in 1770" },
      { role: "body", text: "The First Fleet arrived at Sydney Cove in 1788 - the start of British settlement" },
      { role: "body", text: "British settlement had a huge and harmful impact on First Nations peoples" },
      { role: "body", text: "The British called the land terra nullius - this was not true" },
    ], {
      x: 0.5, y: CONTENT_TOP, w: 4.5, h: SAFE_BOTTOM - CONTENT_TOP,
      strip: C.SECONDARY, fill: C.WHITE,
      headerColor: C.SECONDARY,
    });

    // Right: today's focus card
    const rX = 5.2;
    const rW = 4.3;
    const rH = SAFE_BOTTOM - CONTENT_TOP;
    addCard(s, rX, CONTENT_TOP, rW, rH, { strip: C.PRIMARY, fill: C.BG_LIGHT });

    s.addText("Today we ask", {
      x: rX + 0.2, y: CONTENT_TOP + 0.15, w: rW - 0.4, h: 0.36,
      fontSize: 15, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("After 1788, who else came to live in Australia - and why did they come?", {
      x: rX + 0.2, y: CONTENT_TOP + 0.60, w: rW - 0.4, h: 1.25,
      fontSize: 19, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "left", valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Question prompt at the bottom
    s.addShape("roundRect", {
      x: rX + 0.2, y: CONTENT_TOP + rH - 1.45, w: rW - 0.4, h: 1.20, rectRadius: 0.08,
      fill: { color: C.ACCENT },
    });
    s.addText("Turn and tell:", {
      x: rX + 0.35, y: CONTENT_TOP + rH - 1.30, w: rW - 0.7, h: 0.32,
      fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true, margin: 0,
    });
    s.addText("Name one country your family or someone you know came from. Why did they come?", {
      x: rX + 0.35, y: CONTENT_TOP + rH - 0.95, w: rW - 0.7, h: 0.70,
      fontSize: 13, fontFace: FONT_B, color: C.WHITE,
      valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_LAUNCH);
  }

  // -- Slide 5: Vocabulary --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Vocabulary", { color: C.PRIMARY, w: 1.7 });
    addTitle(s, "Four Inquiry Words");

    const vocab = [
      { word: "migrate", meaning: "to move from one country to another to live - millions of people have migrated to Australia since 1788", color: C.PRIMARY },
      { word: "push factor", meaning: "a reason that pushes people AWAY from their home country - war, poverty, persecution, famine", color: C.ALERT },
      { word: "pull factor", meaning: "a reason that pulls people TOWARDS a new country - jobs, safety, family, education", color: C.SUCCESS },
      { word: "multicultural", meaning: "having many different cultures living together - languages, foods, beliefs and holidays all in one country", color: C.ACCENT },
    ];

    const cardW = 4.45;
    const cardH = 1.8;
    const gapX = 0.1;
    const gapY = 0.1;

    vocab.forEach((v, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 0.5 + col * (cardW + gapX);
      const y = CONTENT_TOP + row * (cardH + gapY);

      addCard(s, x, y, cardW, cardH, { strip: v.color, fill: C.WHITE });
      s.addText(v.word, {
        x: x + 0.2, y: y + 0.15, w: cardW - 0.4, h: 0.50,
        fontSize: 22, fontFace: FONT_H, color: v.color, bold: true, margin: 0,
        fit: "shrink",
      });
      s.addText(v.meaning, {
        x: x + 0.2, y: y + 0.70, w: cardW - 0.4, h: cardH - 0.85,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, margin: 0,
        valign: "top",
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_VOCAB);
  }

  // -- Slide 6: LI / SC --
  liSlide(
    pres,
    ["We are learning about the reasons why people immigrated to Australia, so we can understand our multicultural history"],
    [
      "I can list different cultures that migrated to Australia",
      "I can explain why different cultures migrated to Australia",
      "I can sort reasons into push factors and pull factors",
    ],
    NOTES_LI,
    FOOTER
  );

  // -- Slide 7: I Do - Waves of migration --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "I Do  -  Waves", { color: C.PRIMARY, w: 1.85 });
    addTitle(s, "Four big waves of migration");

    // Timeline strip across the slide
    const tlY = CONTENT_TOP + 0.30;
    const tlX = 0.7;
    const tlW = 8.6;
    const tlH = 0.08;

    s.addShape("rect", {
      x: tlX, y: tlY, w: tlW, h: tlH,
      fill: { color: C.PRIMARY },
    });

    const points = [
      { x: 0.08, label: "1850s", note: "Gold Rush - Chinese miners", color: C.ACCENT },
      { x: 0.36, label: "1945 - 1970", note: "Post-WWII - Italian, Greek, Dutch ...", color: C.SECONDARY },
      { x: 0.64, label: "Late 1970s", note: "Vietnamese refugees", color: C.ALERT },
      { x: 0.92, label: "1990s onwards", note: "Skilled migration - India, China ...", color: C.SUCCESS },
    ];

    const labelW = 1.85;
    const noteW = 2.10;

    points.forEach((p) => {
      const px = tlX + p.x * tlW;
      const py = tlY + tlH / 2;

      // Dot
      s.addShape("roundRect", {
        x: px - 0.14, y: py - 0.14, w: 0.28, h: 0.28, rectRadius: 0.14,
        fill: { color: p.color },
        line: { color: C.WHITE, width: 1.5 },
      });

      // Year label
      s.addText(p.label, {
        x: px - labelW / 2, y: py + 0.20, w: labelW, h: 0.30,
        fontSize: 12, fontFace: FONT_B, color: p.color, bold: true,
        align: "center", margin: 0,
      });

      // Event note
      s.addText(p.note, {
        x: px - noteW / 2, y: py + 0.52, w: noteW, h: 0.55,
        fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "top", margin: 0,
      });
    });

    // Bottom: key idea card with White Australia Policy note
    const keyY = CONTENT_TOP + 2.25;
    const keyH = SAFE_BOTTOM - keyY - 0.05;
    addCard(s, 0.5, keyY, 9, keyH, { strip: C.ALERT, fill: C.BG_LIGHT });

    s.addText("Important to know", {
      x: 0.75, y: keyY + 0.10, w: 5, h: 0.28,
      fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
    });
    s.addText("From 1901 to 1973, the White Australia Policy made it very hard for non-white people to come here. That policy was wrong - it was ended in 1973. After that, migration from Asia, Africa and the Middle East grew.", {
      x: 0.75, y: keyY + 0.42, w: 8.5, h: keyH - 0.50,
      fontSize: 14, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_WAVES);
  }

  // -- Slide 8: I Do - Push factors and Pull factors --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "I Do  -  Why?", { color: C.PRIMARY, w: 1.65 });
    addTitle(s, "Push factors and pull factors");

    // Two big cards side by side: Push (red/alert) and Pull (green/success)
    const cardY = CONTENT_TOP;
    const upperH = 2.70;
    const cardW = (9 - 0.20) / 2;

    // PUSH card
    {
      const x = 0.5;
      addCard(s, x, cardY, cardW, upperH, { strip: C.ALERT, fill: C.WHITE });

      // Arrow icon
      s.addShape("roundRect", {
        x: x + (cardW - 0.55) / 2, y: cardY + 0.15, w: 0.55, h: 0.55, rectRadius: 0.28,
        fill: { color: C.ALERT },
      });
      s.addText("->", {
        x: x + (cardW - 0.55) / 2, y: cardY + 0.15, w: 0.55, h: 0.55,
        fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      s.addText("PUSH FACTORS", {
        x: x + 0.2, y: cardY + 0.80, w: cardW - 0.4, h: 0.32,
        fontSize: 16, fontFace: FONT_H, color: C.ALERT, bold: true,
        align: "center", margin: 0,
      });
      s.addText("Reasons that push people AWAY from their home country", {
        x: x + 0.2, y: cardY + 1.15, w: cardW - 0.4, h: 0.50,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", valign: "top", margin: 0,
      });
      s.addText([
        { text: "War or conflict", options: { bullet: true, breakLine: true } },
        { text: "Persecution or unfair treatment", options: { bullet: true, breakLine: true } },
        { text: "Poverty or famine", options: { bullet: true } },
      ], {
        x: x + 0.30, y: cardY + 1.65, w: cardW - 0.4, h: upperH - 1.75,
        fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
        paraSpaceAfter: 3,
      });
    }

    // PULL card
    {
      const x = 0.5 + cardW + 0.20;
      addCard(s, x, cardY, cardW, upperH, { strip: C.SUCCESS, fill: C.WHITE });

      // Arrow icon
      s.addShape("roundRect", {
        x: x + (cardW - 0.55) / 2, y: cardY + 0.15, w: 0.55, h: 0.55, rectRadius: 0.28,
        fill: { color: C.SUCCESS },
      });
      s.addText("<-", {
        x: x + (cardW - 0.55) / 2, y: cardY + 0.15, w: 0.55, h: 0.55,
        fontSize: 22, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      s.addText("PULL FACTORS", {
        x: x + 0.2, y: cardY + 0.80, w: cardW - 0.4, h: 0.32,
        fontSize: 16, fontFace: FONT_H, color: C.SUCCESS, bold: true,
        align: "center", margin: 0,
      });
      s.addText("Reasons that pull people TOWARDS a new country", {
        x: x + 0.2, y: cardY + 1.15, w: cardW - 0.4, h: 0.50,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", valign: "top", margin: 0,
      });
      s.addText([
        { text: "Jobs and a better life", options: { bullet: true, breakLine: true } },
        { text: "Safety, freedom and family already here", options: { bullet: true, breakLine: true } },
        { text: "Gold, education and opportunity", options: { bullet: true } },
      ], {
        x: x + 0.30, y: cardY + 1.65, w: cardW - 0.4, h: upperH - 1.75,
        fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL, valign: "top", margin: 0,
        paraSpaceAfter: 3,
      });
    }

    // Bottom: worked example card
    const exY = cardY + upperH + 0.15;
    const exH = SAFE_BOTTOM - exY - 0.05;
    addCard(s, 0.5, exY, 9, exH, { strip: C.PRIMARY, fill: C.BG_LIGHT });

    s.addText("Worked example - Italian families after WWII", {
      x: 0.75, y: exY + 0.10, w: 8.5, h: 0.30,
      fontSize: 12, fontFace: FONT_B, color: C.PRIMARY, bold: true, margin: 0,
    });
    s.addText("PUSH: Italy was destroyed by the war. Many families had no work and not enough food.   PULL: Australia needed workers for big projects like the Snowy Mountains Scheme. Italians who came earlier wrote home to their families.", {
      x: 0.75, y: exY + 0.45, w: 8.5, h: exH - 0.55,
      fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_PUSH_PULL);
  }

  // -- Slide 9 / 10: CFU hinge with reveal --
  function buildCfuBase() {
    const s = pres.addSlide();
    addTopBar(s, C.ALERT);
    addBadge(s, "CFU", { color: C.ALERT });
    addTitle(s, "Sort the push from the pull", { color: C.ALERT });

    // Technique pill
    s.addShape("roundRect", {
      x: 0.5, y: CONTENT_TOP, w: 2.8, h: 0.40, rectRadius: 0.08,
      fill: { color: C.ALERT },
    });
    s.addText("Show Me Boards", {
      x: 0.5, y: CONTENT_TOP, w: 2.8, h: 0.40,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
    });

    const aH = 0.65;
    const aY = SAFE_BOTTOM - aH - 0.05;
    const qY = CONTENT_TOP + 0.56;
    const qH = aY - 0.20 - qY;

    addCard(s, 0.5, qY, 9, qH, { strip: C.ALERT, fill: C.WHITE });
    s.addText(CFU_Q_TEXT, {
      x: 0.75, y: qY + 0.30, w: 8.5, h: qH - 0.60,
      fontSize: 20, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle",
      align: "center", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    // Answer bar placeholder
    s.addShape("roundRect", {
      x: 0.5, y: aY, w: 9, h: aH, rectRadius: 0.1,
      fill: { color: C.BG_LIGHT },
      line: { color: C.MUTED, width: 0.6, dashType: "dash" },
    });
    s.addText("Answer revealed next click", {
      x: 0.5, y: aY, w: 9, h: aH,
      fontSize: 11, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_CFU_Q);
    return s;
  }

  withReveal(
    buildCfuBase,
    (s) => {
      const aH = 0.65;
      const aY = SAFE_BOTTOM - aH - 0.05;
      s.addShape("roundRect", {
        x: 0.5, y: aY, w: 9, h: aH, rectRadius: 0.1,
        fill: { color: C.SUCCESS },
      });
      s.addText("PUSH: war in their home country   |   PULL: cousin already in Australia (family)", {
        x: 0.5, y: aY, w: 9, h: aH,
        fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });
      s.addNotes(NOTES_CFU_A);
    }
  );

  // -- Slide 11: We Do - Model a profile card --
  contentSlide(
    pres,
    "We Do  -  Model",
    C.SECONDARY,
    "Filling in a migration profile card",
    [
      "Each card has three boxes: WHEN, WHY, ONE FACT",
      "In the WHY box, label every reason as (push) or (pull)",
      "Watch the teacher model the Italian card together",
      "Then you fill in three more groups",
    ],
    NOTES_WE_DO_PROFILE,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const panelH = SAFE_BOTTOM - topY - 0.05;

      addCard(s, rX, topY, rW, panelH, { strip: C.SECONDARY, fill: C.BG_LIGHT });

      s.addText("Modelled card - Italian", {
        x: rX + 0.2, y: topY + 0.10, w: rW - 0.4, h: 0.28,
        fontSize: 12, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      // Mock profile card
      const cX = rX + 0.30;
      const cW = rW - 0.60;
      const cY = topY + 0.45;
      const cH = panelH - 0.55;

      s.addShape("rect", {
        x: cX, y: cY, w: cW, h: cH,
        fill: { color: C.WHITE },
        line: { color: C.MUTED, width: 0.8 },
      });

      // Group title
      s.addShape("rect", {
        x: cX, y: cY, w: cW, h: 0.30,
        fill: { color: C.SECONDARY },
      });
      s.addText("Italian families", {
        x: cX, y: cY, w: cW, h: 0.30,
        fontSize: 11, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Three rows: WHEN, WHY, FACT
      const rowGap = 0.04;
      const rowTotal = cH - 0.30 - rowGap * 2;
      const whenH = rowTotal * 0.18;
      const whyH = rowTotal * 0.52;
      const factH = rowTotal * 0.30;
      let ry = cY + 0.30;

      // WHEN
      s.addShape("rect", {
        x: cX, y: ry, w: cW, h: whenH,
        fill: { color: C.BG_LIGHT },
        line: { color: C.MUTED, width: 0.4 },
      });
      s.addText("WHEN: After WWII, 1945 - 1970", {
        x: cX + 0.10, y: ry, w: cW - 0.20, h: whenH,
        fontSize: 9, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
        valign: "middle", margin: 0,
      });
      ry += whenH + rowGap;

      // WHY
      s.addShape("rect", {
        x: cX, y: ry, w: cW, h: whyH,
        fill: { color: C.WHITE },
        line: { color: C.MUTED, width: 0.4 },
      });
      s.addText([
        { text: "WHY:", options: { bold: true, breakLine: true } },
        { text: "War damage and poverty (push)", options: { breakLine: true } },
        { text: "Jobs in Australia, family already here (pull)", options: {} },
      ], {
        x: cX + 0.10, y: ry + 0.05, w: cW - 0.20, h: whyH - 0.10,
        fontSize: 9, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "top", margin: 0, paraSpaceAfter: 2,
      });
      ry += whyH + rowGap;

      // FACT
      s.addShape("rect", {
        x: cX, y: ry, w: cW, h: factH,
        fill: { color: C.BG_LIGHT },
        line: { color: C.MUTED, width: 0.4 },
      });
      s.addText("ONE FACT: Many Italian men worked on the Snowy Mountains Scheme", {
        x: cX + 0.10, y: ry + 0.05, w: cW - 0.20, h: factH - 0.10,
        fontSize: 9, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        valign: "top", margin: 0,
      });
    }
  );

  // -- Slide 12: You Do - Complete profile cards --
  contentSlide(
    pres,
    "You Do  -  Profile cards",
    C.ACCENT,
    "Complete your migration profile cards",
    [
      "One Migration Profile Cards sheet per student",
      "Italian is already modelled - complete Chinese, Vietnamese and Indian",
      "Label every reason in the WHY box as (push) or (pull)",
      "Use the information bank on the back if you need help",
      "12 minutes - work with your shoulder partner",
    ],
    NOTES_YOU_DO_PROFILE,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const panelH = SAFE_BOTTOM - topY - 0.05;

      addCard(s, rX, topY, rW, panelH, { strip: C.ACCENT, fill: C.WHITE });

      s.addText("Your cards", {
        x: rX + 0.2, y: topY + 0.10, w: rW - 0.4, h: 0.28,
        fontSize: 11, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });

      // Mock 2x2 grid of mini cards
      const grdX = rX + 0.25;
      const grdY = topY + 0.45;
      const grdW = rW - 0.50;
      const grdH = panelH - 0.55;
      const miniGap = 0.08;
      const miniW = (grdW - miniGap) / 2;
      const miniH = (grdH - miniGap) / 2;

      const groups = [
        { name: "Chinese", color: C.ACCENT },
        { name: "Italian", color: C.SECONDARY, modelled: true },
        { name: "Vietnamese", color: C.ALERT },
        { name: "Indian", color: C.SUCCESS },
      ];

      groups.forEach((g, i) => {
        const col = i % 2;
        const row = Math.floor(i / 2);
        const mx = grdX + col * (miniW + miniGap);
        const my = grdY + row * (miniH + miniGap);

        // Card outline
        s.addShape("rect", {
          x: mx, y: my, w: miniW, h: miniH,
          fill: { color: g.modelled ? C.BG_LIGHT : C.WHITE },
          line: { color: C.MUTED, width: 0.6 },
        });

        // Title strip
        s.addShape("rect", {
          x: mx, y: my, w: miniW, h: 0.22,
          fill: { color: g.color },
        });
        s.addText(g.name, {
          x: mx, y: my, w: miniW, h: 0.22,
          fontSize: 9, fontFace: FONT_B, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });

        // Box labels
        const labY = my + 0.27;
        const labH = miniH - 0.30;
        s.addText([
          { text: g.modelled ? "WHEN: filled in" : "WHEN: __________", options: { breakLine: true } },
          { text: g.modelled ? "WHY: filled in" : "WHY: __________", options: { breakLine: true } },
          { text: g.modelled ? "FACT: filled in" : "FACT: __________", options: {} },
        ], {
          x: mx + 0.08, y: labY, w: miniW - 0.16, h: labH,
          fontSize: 7, fontFace: FONT_B, color: g.modelled ? C.MUTED : C.CHARCOAL, italic: g.modelled,
          valign: "top", margin: 0, paraSpaceAfter: 2,
        });
      });
    }
  );

  // -- Slide 13: We Do - Model the explanation paragraph --
  contentSlide(
    pres,
    "We Do  -  Model",
    C.SECONDARY,
    "Writing a push and pull explanation",
    [
      "Use the scaffold - four paragraphs",
      "Use your profile card for the facts",
      "Use the words 'push factor' and 'pull factor'",
      "Watch the teacher model the introduction and the push paragraph",
    ],
    NOTES_WE_DO_EXPLAIN,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const panelH = SAFE_BOTTOM - topY - 0.05;

      addCard(s, rX, topY, rW, panelH, { strip: C.SECONDARY, fill: C.BG_LIGHT });

      s.addText("Modelled opening (Vietnamese)", {
        x: rX + 0.2, y: topY + 0.10, w: rW - 0.4, h: 0.28,
        fontSize: 11, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      // Mock paragraph preview
      const lX = rX + 0.30;
      const lW = rW - 0.60;
      const lY = topY + 0.45;
      const lH = panelH - 0.55;

      s.addShape("rect", {
        x: lX, y: lY, w: lW, h: lH,
        fill: { color: C.WHITE },
        line: { color: C.MUTED, width: 0.8 },
      });

      s.addText("P1 - Introduction", {
        x: lX + 0.10, y: lY + 0.10, w: lW - 0.20, h: 0.22,
        fontSize: 9, fontFace: FONT_B, color: C.PRIMARY, bold: true,
        margin: 0,
      });
      s.addText("Many Vietnamese families came to Australia in the late 1970s and 1980s.", {
        x: lX + 0.10, y: lY + 0.33, w: lW - 0.20, h: 0.42,
        fontSize: 9, fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, valign: "top", margin: 0,
      });

      s.addText("P2 - Push factor", {
        x: lX + 0.10, y: lY + 0.85, w: lW - 0.20, h: 0.22,
        fontSize: 9, fontFace: FONT_B, color: C.ALERT, bold: true,
        margin: 0,
      });
      s.addText("They were pushed out of their home country by the Vietnam War. After the war ended, many people were in danger.", {
        x: lX + 0.10, y: lY + 1.08, w: lW - 0.20, h: lH - 1.18,
        fontSize: 9, fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, valign: "top", margin: 0,
      });
    }
  );

  // -- Slide 14: You Do - Write explanation --
  contentSlide(
    pres,
    "You Do  -  Write",
    C.ACCENT,
    "Write your push and pull explanation",
    [
      "Choose ONE group from your profile cards",
      "P1: introduce the group and when they came",
      "P2: explain the PUSH factors",
      "P3: explain the PULL factors",
      "P4: one fact about their contribution today",
    ],
    NOTES_YOU_DO_EXPLAIN,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const panelH = SAFE_BOTTOM - topY - 0.05;

      addCard(s, rX, topY, rW, panelH, { strip: C.ACCENT, fill: C.WHITE });

      s.addText("Your scaffold", {
        x: rX + 0.2, y: topY + 0.10, w: rW - 0.4, h: 0.28,
        fontSize: 11, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });

      // Mock scaffold: 4 paragraph blocks
      const blockX = rX + 0.30;
      const blockW = rW - 0.60;
      const blockTopY = topY + 0.45;
      const totalBlockH = panelH - 0.55;
      const gap = 0.06;
      const blockH = (totalBlockH - gap * 3) / 4;

      const starters = [
        { n: "P1", color: C.PRIMARY, text: "Many ___ families came to Australia in ..." },
        { n: "P2", color: C.ALERT, text: "PUSH: They were pushed away from home because ..." },
        { n: "P3", color: C.SUCCESS, text: "PULL: They were pulled to Australia because ..." },
        { n: "P4", color: C.SECONDARY, text: "Today, this group contributes ..." },
      ];

      starters.forEach((b, i) => {
        const by = blockTopY + i * (blockH + gap);
        s.addShape("rect", {
          x: blockX, y: by, w: blockW, h: blockH,
          fill: { color: C.BG_LIGHT },
          line: { color: C.MUTED, width: 0.5 },
        });
        // Tab marker
        s.addShape("rect", {
          x: blockX, y: by, w: 0.30, h: blockH,
          fill: { color: b.color },
        });
        s.addText(b.n, {
          x: blockX, y: by, w: 0.30, h: blockH,
          fontSize: 9, fontFace: FONT_B, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
        s.addText(b.text, {
          x: blockX + 0.40, y: by, w: blockW - 0.50, h: blockH,
          fontSize: 9, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
          align: "left", valign: "middle", margin: 0,
        });
      });
    }
  );

  // -- Slide 15: Exit Ticket --
  exitTicketSlide(
    pres,
    [
      "Sticky note - finish this sentence and stick it on the class chart on your way out:\n\n'One cultural group that came to Australia was ___. They came because of the push factor ___ and the pull factor ___.'",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2, title: "One thing before you leave" }
  );

  // -- Slide 16: Closing --
  closingSlide(
    pres,
    {
      reflectionPrompt: "Which migration story stood out to you today, and why?",
      scItems: [
        "I can list different cultures that migrated to Australia",
        "I can explain why different cultures migrated to Australia",
        "I can sort reasons into push factors and pull factors",
      ],
      selfAssessment: {
        prompt: "Thumbs up, sideways, or down for each one",
        options: ["Got it", "Getting there", "Need more"],
      },
    },
    NOTES_CLOSING
  );

  // -- Write PPTX --
  fs.mkdirSync(LESSON_FOLDER, { recursive: true });
  await pres.writeFile({ fileName: path.join(LESSON_FOLDER, PPTX_NAME) });
  console.log("PPTX written to", path.join(LESSON_FOLDER, PPTX_NAME));

  // ===============================================================
  // PDFs
  // ===============================================================
  await buildProfileCardsPdf();
  await buildPushPullScaffoldPdf();
}

// ───────────────────────────────────────────────────────────────
//  PDF 1: Migration profile cards
// ───────────────────────────────────────────────────────────────

async function buildProfileCardsPdf() {
  const doc = createPdf({ title: "Migration Profile Cards" });
  let y = addPdfHeader(doc, "Migration Profile Cards", {
    subtitle: "Four migrant groups. Italian is modelled. Complete Chinese, Vietnamese and Indian.",
    color: C.PRIMARY,
    lessonInfo: "Year 5/6 Inquiry  |  Shaping Australia  |  Lesson 4",
  });

  // Brief instruction
  y = addBodyText(doc,
    "For each card, fill in WHEN they came, WHY they came (label every reason as 'push' or 'pull'), and ONE fact about them.",
    y, { fontSize: 10.5, italic: true, color: "4B5563" }
  );

  y += 6;

  // Cards - one per row, 4 rows
  const tableX = PAGE.MARGIN;
  const tableW = PAGE.CONTENT_W;
  const cardH = 100;
  const cardGap = 8;
  const labelW = 70;

  const cards = [
    {
      group: "Chinese",
      color: C.ACCENT,
      modelled: false,
    },
    {
      group: "Italian  -  modelled",
      color: C.SECONDARY,
      modelled: true,
      modelWhen: "After WWII, 1945 - 1970",
      modelWhy: "War damage and poverty (push). Jobs in Australia, family already here (pull).",
      modelFact: "Many Italian men worked on the Snowy Mountains Scheme.",
    },
    {
      group: "Vietnamese",
      color: C.ALERT,
      modelled: false,
    },
    {
      group: "Indian",
      color: C.SUCCESS,
      modelled: false,
    },
  ];

  cards.forEach((c) => {
    // Title bar
    doc.save();
    doc.rect(tableX, y, tableW, 18).fill(hex(c.color));
    doc.fontSize(11).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(c.group, tableX + 8, y + 5, { width: tableW - 16 });
    doc.restore();

    // Card body
    doc.save();
    doc.rect(tableX, y + 18, tableW, cardH).lineWidth(0.8).strokeColor("#9CA3AF").stroke();
    doc.restore();

    // Internal dividers
    const rowH = cardH / 3;
    for (let i = 1; i < 3; i++) {
      doc.save();
      doc.moveTo(tableX, y + 18 + rowH * i).lineTo(tableX + tableW, y + 18 + rowH * i)
        .lineWidth(0.4).strokeColor("#D1D5DB").stroke();
      doc.restore();
    }

    // Label column background
    doc.save();
    doc.rect(tableX, y + 18, labelW, cardH).fill(hex("F4F1E6"));
    doc.restore();
    doc.save();
    doc.moveTo(tableX + labelW, y + 18).lineTo(tableX + labelW, y + 18 + cardH)
      .lineWidth(0.4).strokeColor("#D1D5DB").stroke();
    doc.restore();

    // Labels
    const labels = ["WHEN", "WHY", "ONE FACT"];
    labels.forEach((label, i) => {
      doc.fontSize(9).font("Sans-Bold").fillColor(hex(C.CHARCOAL || "2D3142"));
      doc.text(label, tableX + 6, y + 18 + rowH * i + rowH / 2 - 6, { width: labelW - 12 });
    });

    // Modelled content (only for Italian)
    if (c.modelled) {
      doc.fontSize(9.5).font("Sans-Italic").fillColor(hex(C.CHARCOAL || "2D3142"));
      doc.text(c.modelWhen, tableX + labelW + 8, y + 18 + 6, { width: tableW - labelW - 16 });
      doc.text(c.modelWhy, tableX + labelW + 8, y + 18 + rowH + 6, { width: tableW - labelW - 16 });
      doc.text(c.modelFact, tableX + labelW + 8, y + 18 + rowH * 2 + 6, { width: tableW - labelW - 16 });
    } else {
      // Light writing guides
      const writingX = tableX + labelW + 8;
      const writingW = tableW - labelW - 16;
      for (let i = 0; i < 3; i++) {
        const rowY = y + 18 + rowH * i;
        const innerH = rowH;
        const linesPerRow = innerH > 30 ? 2 : 1;
        const lineSpacing = (innerH - 10) / (linesPerRow + 1);
        for (let li = 1; li <= linesPerRow; li++) {
          const ly = rowY + 6 + li * lineSpacing;
          doc.save();
          doc.moveTo(writingX, ly).lineTo(writingX + writingW, ly)
            .lineWidth(0.4).strokeColor("#D1D5DB").stroke();
          doc.restore();
        }
      }
    }

    y = y + 18 + cardH + cardGap;
  });

  // Add information bank on a new page
  doc.addPage();
  y = PAGE.MARGIN;

  doc.fontSize(14).font("Sans-Bold").fillColor(hex(C.PRIMARY));
  doc.text("Information bank  -  use these facts to help fill your cards", PAGE.MARGIN, y);
  y += 22;

  const bankItems = [
    {
      group: "Chinese",
      color: C.ACCENT,
      facts: [
        "Came in large numbers during the 1850s Gold Rush in Victoria.",
        "Push: poverty in China; war and unrest at home.",
        "Pull: gold; chance of a better life.",
        "Many Chinese miners walked from South Australia to the Victorian goldfields.",
        "Faced unfair laws and discrimination; many later returned to China.",
      ],
    },
    {
      group: "Vietnamese",
      color: C.ALERT,
      facts: [
        "Came in the late 1970s and 1980s after the Vietnam War ended in 1975.",
        "Push: war, danger, prison camps, fear of persecution.",
        "Pull: safety; family members who came earlier.",
        "Around 90,000 came to Australia. Some came by boat. Australia accepted them as refugees.",
        "Today, Vietnamese is one of the most spoken non-English languages in Australia.",
      ],
    },
    {
      group: "Indian",
      color: C.SUCCESS,
      facts: [
        "Many came from the 1990s onwards through skilled migration and student visas.",
        "Push: limited work opportunities at home; high cost of education.",
        "Pull: good jobs in technology, engineering, healthcare and finance; Australian universities; family reunion.",
        "India is now one of the biggest sources of new migrants to Australia.",
        "Festivals like Diwali are now celebrated in many Australian cities.",
      ],
    },
  ];

  bankItems.forEach((bi) => {
    // Title bar
    doc.save();
    doc.rect(PAGE.MARGIN, y, PAGE.CONTENT_W, 16).fill(hex(bi.color));
    doc.fontSize(10).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(bi.group, PAGE.MARGIN + 8, y + 4, { width: PAGE.CONTENT_W - 16 });
    doc.restore();
    y += 18;

    bi.facts.forEach((f) => {
      doc.fontSize(9.5).font("Sans").fillColor(hex("2D3142"));
      doc.text("- " + f, PAGE.MARGIN + 12, y, { width: PAGE.CONTENT_W - 24 });
      y = doc.y + 3;
    });
    y += 6;
  });

  // Extension prompt
  y += 4;
  y = addTipBox(doc,
    "Extension: under your last card, write one sentence starting with 'Both ___ and ___ had ___ as a push factor, but their pull factors were different because ___.'",
    y,
    { color: C.ACCENT }
  );

  addPdfFooter(doc, "Year 5/6 Inquiry  -  Shaping Australia  -  Lesson 4  -  Migration Profile Cards");
  const outPath = path.join(RES_DIR, "Session 4 Migration Profile Cards.pdf");
  await writePdf(doc, outPath);
  console.log("PDF written to", outPath);
}

// ───────────────────────────────────────────────────────────────
//  PDF 2: Push and pull explanation scaffold
// ───────────────────────────────────────────────────────────────

async function buildPushPullScaffoldPdf() {
  const doc = createPdf({ title: "Push and Pull Explanation Scaffold" });
  let y = addPdfHeader(doc, "Push and Pull Explanation", {
    subtitle: "Choose ONE group from your profile cards. Use 'push factor' and 'pull factor' in your writing.",
    color: C.PRIMARY,
    lessonInfo: "Year 5/6 Inquiry  |  Shaping Australia  |  Lesson 4",
  });

  // Group selector box
  y = addTipBox(doc,
    "Group I am writing about: ____________________________________   (Chinese / Italian / Vietnamese / Indian / Other)",
    y,
    { color: C.SECONDARY }
  );

  y += 6;

  // Four scaffold blocks
  const blocks = [
    { label: "P1", color: C.PRIMARY, prompt: "Introduce the group and when they came.\nTry: 'Many ___ families came to Australia in ___.'" },
    { label: "P2", color: C.ALERT, prompt: "Explain the PUSH factors - what pushed them away from home.\nTry: 'They were pushed out of their home country by ___. This pushed them to leave because ___.'" },
    { label: "P3", color: C.SUCCESS, prompt: "Explain the PULL factors - what pulled them to Australia.\nTry: 'They were pulled to Australia by ___. This pulled them here because ___.'" },
    { label: "P4", color: C.SECONDARY, prompt: "One fact about the group today.\nTry: 'Today, ___ Australians ___.'" },
  ];

  const blockX = PAGE.MARGIN;
  const blockW = PAGE.CONTENT_W;
  const labelW = 30;
  const blockH = 110;
  const blockGap = 10;

  blocks.forEach((b) => {
    // Label tab
    doc.save();
    doc.rect(blockX, y, labelW, blockH).fill(hex(b.color));
    doc.fontSize(13).font("Sans-Bold").fillColor("#FFFFFF");
    doc.text(b.label, blockX, y + blockH / 2 - 8, { width: labelW, align: "center" });
    doc.restore();

    // Body area
    doc.save();
    doc.rect(blockX + labelW, y, blockW - labelW, blockH).lineWidth(0.8).strokeColor("#9CA3AF").stroke();
    doc.restore();

    // Prompt (italic, top)
    doc.fontSize(9).font("Sans-Italic").fillColor(hex("6B7280"));
    doc.text(b.prompt, blockX + labelW + 10, y + 6, { width: blockW - labelW - 20 });

    // Writing lines
    const linesStartY = y + 44;
    const lineCount = 4;
    const lineGap = (blockH - 44 - 6) / lineCount;
    for (let li = 0; li < lineCount; li++) {
      const ly = linesStartY + li * lineGap;
      doc.save();
      doc.moveTo(blockX + labelW + 10, ly).lineTo(blockX + blockW - 10, ly)
        .lineWidth(0.6).strokeColor("#000000").stroke();
      doc.restore();
    }

    y = y + blockH + blockGap;
  });

  addPdfFooter(doc, "Year 5/6 Inquiry  -  Shaping Australia  -  Lesson 4  -  Push and Pull Explanation");
  const outPath = path.join(RES_DIR, "Session 4 Push and Pull Explanation Scaffold.pdf");
  await writePdf(doc, outPath);
  console.log("PDF written to", outPath);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
