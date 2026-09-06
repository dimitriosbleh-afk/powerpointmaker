"use strict";

// Inquiry - Shaping Australia, Lesson 2 (Term 2, Week 2)
// Year 5/6 | Aboriginal Culture: Two Cultures, One Country.
// Reads the Aboriginal Heritage history page like a picture story book,
// explores Aboriginal Dreaming, and the AIATSIS map of Indigenous
// Australia. Class then compares Indigenous and non-Indigenous cultures,
// reflects on the "what if" of peaceful settlement, and adds to the KWL
// chart from Lesson 1.
//
// User-supplied links (preserved exactly):
//   Aboriginal Heritage history page
//     https://www.aboriginalheritage.org/history/history/
//   Aboriginal Dreaming and the Dreamtime
//     https://www.aboriginal-art-australia.com/aboriginal-art-library/understanding-aboriginal-dreaming-and-the-dreamtime/
//   AIATSIS Map of Indigenous Australia
//     https://aiatsis.gov.au/explore/map-indigenous-australia
//   ABC: The Stolen Generation
//     https://www.youtube.com/watch?v=CE7hCCO9jv8
//   Ask Us Anything
//     https://www.youtube.com/watch?v=SHVbVBLlhCM
//   The Stolen Generation: Mona
//     https://www.youtube.com/watch?v=KE6GYUc47Qg

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
// Term 2 Week 2 -> variant 0 (Explorer - olive) - matches Lesson 1 (unit cohesion)
const T = createTheme("inquiry", "grade56", weekToVariant(1));
const {
  C, FONT_H, FONT_B,
  titleSlide, liSlide, contentSlide, closingSlide,
  exitTicketSlide, pairShareSlide, withReveal,
  addTopBar, addBadge, addTitle, addCard, addFooter,
  addInstructionCard,
  CONTENT_TOP, SAFE_BOTTOM,
} = T;

// -- Output paths --
const UNIT = "Shaping_Australia";
const LESSON_FOLDER = path.join(__dirname, "..", "output", UNIT);
const PPTX_NAME = "Lesson 2 - Aboriginal Culture.pptx";
const FOOTER = "Inquiry | Year 5/6 | Shaping Australia | Lesson 2";
const SESSION = 2;
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

// -- URLs (user-supplied, preserve exactly) --
const URL_HERITAGE = "https://www.aboriginalheritage.org/history/history/";
const URL_DREAMING = "https://www.aboriginal-art-australia.com/aboriginal-art-library/understanding-aboriginal-dreaming-and-the-dreamtime/";
const URL_AIATSIS  = "https://aiatsis.gov.au/explore/map-indigenous-australia";
const URL_BTN_STOLEN = "https://www.youtube.com/watch?v=CE7hCCO9jv8";
const URL_ASK_US     = "https://www.youtube.com/watch?v=SHVbVBLlhCM";
const URL_MONA       = "https://www.youtube.com/watch?v=KE6GYUc47Qg";

// -- Resources --
const COMPARE_CHART = makeSessionResource(
  SESSION,
  "Cultures Compared Chart",
  "Two-column comparison chart with similarities band, discussion prompts, and a What if? reflection space."
);
const RESOURCE_ITEMS = [COMPARE_CHART];

fs.mkdirSync(RES_DIR, { recursive: true });

// ===============================================================
// Teacher Notes
// ===============================================================

const NOTES_TITLE = [
  "SAY:",
  "- Welcome back to our Shaping Australia unit. Last lesson we collected what we thought we knew, watched two clips and started a class wonderings list",
  "- Today we focus on the people who have lived here for tens of thousands of years - First Nations Australians",
  "- We are going to read together, look at a map of Indigenous Australia, learn about Dreaming, and add what we learn to the chart we started last lesson",
  "",
  "DO:",
  "- Display this slide as students enter",
  "- Have the class KWL chart from Lesson 1 visible (or photographed and reprojected) so it is ready to add to",
  "- Have the comparison chart printed - one per pair",
  "",
  "TEACHER NOTES:",
  "This lesson handles sensitive content with care. Read the Sensitivity Advisory sections in each slide before teaching. The framing is respectful inquiry - we listen, we learn, we ask questions, we do not put any student in the position of speaking on behalf of a community.",
  "",
  "WATCH FOR:",
  "- First Nations students, or students with personal/family connection - quietly check in before the lesson begins, offer the option to step out at any point",
  "",
  "[Inquiry: Title | VTLM 2.0: Establishing Purpose]",
].join("\n");

const NOTES_RESOURCES = [
  "SAY:",
  "- A few things to set up for today: one printed chart, the KWL from last lesson, and three websites we will visit together",
  "",
  "DO:",
  "- Print the Cultures Compared Chart - one per pair (A4)",
  "- Have the Lesson 1 KWL chart visible - either the original whiteboard chart or a photo reprojected",
  "- Open the three website tabs on the teacher computer in advance: Aboriginal Heritage history page, Aboriginal Dreaming page, AIATSIS map",
  "- Have mini-whiteboards available for the CFU check",
  "- Sticky notes ready for the exit ticket",
  "- Optional: queue the three videos listed under Videos & media in case you choose to use one this lesson or in a follow-up",
  "",
  "TEACHER NOTES:",
  "We are visiting three external websites this lesson. Pre-load each tab so the class transitions are smooth. The reading site reads like a picture-story-book - stop and discuss often.",
  "",
  "WATCH FOR:",
  "- Internet access on the teacher computer - test all three links before the lesson",
  "",
  "[Inquiry: Resources | VTLM 2.0: Preparation]",
].join("\n");

const NOTES_AOC = [
  "SAY:",
  "- Before we begin today, we pause and acknowledge",
  "- We acknowledge the Traditional Owners of the land on which we are learning today",
  "- We pay our respects to Elders past and present, and to all Aboriginal and Torres Strait Islander peoples",
  "- We thank them for caring for this land for tens of thousands of years",
  "",
  "DO:",
  "- Read the Acknowledgement aloud, slowly. The class stands quietly. No talking, no fidgeting",
  "- Replace the bracketed nation name with the local Country if your school's standard wording uses it",
  "- Allow a few seconds of silence after the Acknowledgement before moving on",
  "",
  "TEACHER NOTES:",
  "An Acknowledgement of Country is the appropriate way to open any lesson on First Nations content. Keep the tone respectful, not performative. The school may have a standard wording - use that if so.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: A formal recognition of First Nations Australians as the original and continuing custodians of this land.",
  "- Framing language: 'We acknowledge. We pay our respects. We thank.'",
  "- Watch for: Students who treat it as routine - reset the tone with: 'This matters. We do this properly'.",
  "- Protocol: If your school has a Welcome to Country recorded by a local Elder, you may use that instead of an Acknowledgement.",
  "",
  "WATCH FOR:",
  "- Quiet, respectful student behaviour - praise it briefly afterwards",
  "",
  "[Inquiry: Acknowledgement | VTLM 2.0: Cultural Protocol]",
].join("\n");

const NOTES_LAUNCH = [
  "SAY:",
  "- Last lesson we started a chart - what we thought we knew, what we learnt from the two clips, and our wonderings",
  "- Some of those wonderings were about the people who lived here long before British settlement",
  "- Today we go looking for answers to those wonderings - and we will collect new ones along the way",
  "- One question to hold in our heads today: how is First Nations culture different from non-Indigenous culture, and what can we learn from both?",
  "",
  "DO:",
  "- Project the Lesson 1 KWL chart beside this slide if possible",
  "- Read the slide aloud",
  "- Quick partner share: 'From last lesson, what is ONE thing about First Nations Australians that you want to find out today?'",
  "- Cold call 2-3 students to share. Star their wondering on the class chart if it appears there",
  "",
  "TEACHER NOTES:",
  "The launch reactivates the wonderings from Lesson 1 and points them at today's focus. Keep it brisk - the launch is the bridge, not the lesson.",
  "",
  "WATCH FOR:",
  "- Students whose wondering is broad ('everything') - prompt: 'Pick one thing to look for today'",
  "- Students who say 'nothing' - reassure: 'That is okay. Listen carefully today and we will collect together'",
  "",
  "[Inquiry: Launch | VTLM 2.0: Activating Prior Knowledge]",
].join("\n");

const NOTES_VOCAB = [
  "SAY:",
  "- Four inquiry words for today",
  "- Read each word and meaning with me",
  "- 'Country' is the big one. For First Nations Australians, Country means more than land - it means place, family, ancestors, language, and stories",
  "",
  "DO:",
  "- Read each word aloud. Have students repeat 'Country' and 'culture' after you",
  "- Point to the small icons as you go",
  "",
  "CFU CHECKPOINT:",
  "Technique: Finger Voting",
  "Script:",
  "- Say: Hold up 1 finger if Country only means dirt and grass. Hold up 2 if Country means land plus family, ancestors, language and stories",
  "- Scan for: most students showing 2",
  "PROCEED: If 80% show 2, move on to the LI / SC.",
  "PIVOT: If many show 1, say: 'Listen again. Country is a much bigger word for First Nations people. Land, water, sky, family, language, stories. All of it together'.",
  "",
  "TEACHER NOTES:",
  "These four words are the inquiry vocabulary. We will return to 'Country' on the AIATSIS map slide and 'Dreaming' on the Aboriginal Art Australia slide.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: Cultural concepts (Country, Dreaming) hold deep meaning for First Nations communities.",
  "- Framing language: 'We use these words carefully and respectfully'.",
  "- Watch for: Students who joke about 'dreaming' as in 'daydreaming' - redirect: 'In this lesson, Dreaming has a special meaning. We will look at that on the next slide'.",
  "",
  "WATCH FOR:",
  "- Students reading 'Indigenous' as a word they do not understand - clarify: 'It means originally from this place'",
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
  "Internal tier mapping (do not share with students). SC1 - everyone can name one feature. SC2 - most students will compare. SC3 - students explain or reflect on what we can learn from both.",
  "",
  "WATCH FOR:",
  "- Students who think 'compare' means 'pick the better one' - clarify: 'Compare means describe how they are different and how they are the same. Not better or worse'",
  "",
  "[Inquiry: LI/SC | VTLM 2.0: Learning Intentions]",
].join("\n");

const NOTES_IDO_HERITAGE = [
  "SAY:",
  "- We are going to visit a website together - the Aboriginal Heritage history page",
  "- We will read it like a picture story book. I read aloud, you listen, we stop and talk along the way",
  "- Your job: notice ONE thing that surprises you, or one thing you have a question about. Hold it in your head until I pause",
  "",
  "DO:",
  "- Open the link from this slide on the teacher computer (project it large)",
  "- Read the page aloud, scrolling slowly",
  "- Stop after each clear section. Quick prompts: 'What stood out?' or 'Did anything surprise you?'",
  "- Capture 1-2 student noticings in the L column of the class KWL chart as you go",
  "- Plan for about 12-15 minutes of guided reading. Skip-skim later sections if time is tight",
  "",
  "SOURCES:",
  "- Aboriginal Heritage history page. Used as supplied: " + URL_HERITAGE,
  "",
  "TEACHER NOTES:",
  "Read the page yourself the night before so you know where the natural pause points are. The page describes deep history including aspects of dispossession - read with the tone of a teacher sharing important history, not as drama.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: The page covers tens of thousands of years of First Nations history including the impacts of British colonisation.",
  "- Framing language: 'This is real Australian history. Some parts are hard. Historians study it so we understand what happened and why'.",
  "- Watch for: First Nations students or students with personal connection - check in quietly during or after.",
  "- Protocol: If a student needs space, quietly offer them a desk task or time outside the room. Do not single out any student to comment.",
  "",
  "WATCH FOR:",
  "- Students wanting to interrupt with comments - hold them with: 'Save it - we'll share when we pause'",
  "- Students who go quiet - that is fine. Quiet thinking is a valid response to heavy content",
  "",
  "[Inquiry: I Do (Read) | VTLM 2.0: Shared Reading]",
].join("\n");

const NOTES_DREAMING = [
  "SAY:",
  "- Now we look at Dreaming - one of the big ideas in Aboriginal culture",
  "- Dreaming, or the Dreamtime, is the way many Aboriginal peoples explain how the world was made, how the land took its shape, and how people, animals and plants are connected",
  "- Dreaming is not a story from long ago that is over. It is a continuing way of understanding the world",
  "- We will read short sections of this site together",
  "",
  "DO:",
  "- Open the Aboriginal Art Australia Dreaming page from the link on this slide",
  "- Read the introduction section aloud. Pause after each paragraph",
  "- Ask the class: 'What does Dreaming help to explain?'",
  "- Listen for: how the land was formed, how people are connected to land, how animals got their features, how to live well together",
  "- Write 1-2 student responses in the L column of the class KWL chart",
  "- Do not retell or invent specific Dreaming stories - if students ask, point them to the page rather than inventing details",
  "",
  "SOURCES:",
  "- Aboriginal Dreaming and the Dreamtime, Aboriginal Art Australia. Used as supplied: " + URL_DREAMING,
  "",
  "TEACHER NOTES:",
  "Treat Dreaming with the respect it deserves. Specific Dreaming stories belong to specific peoples. The teacher's job is to introduce the concept, not to retell stories that belong elsewhere.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: Dreaming is sacred to many Aboriginal peoples and is held with deep cultural significance.",
  "- Framing language: 'These ideas are sacred. We treat them with respect, the way we would treat any community's beliefs'.",
  "- Watch for: Students treating Dreaming like a fairy tale or daydreaming - gently redirect: 'In this lesson, Dreaming is a serious word with deep meaning'.",
  "- Protocol: If a story is not on the page, do not improvise one. Refer the question back to the page or save it for the W column.",
  "",
  "WATCH FOR:",
  "- Students conflating Dreaming with mythology - clarify: 'It is not myth, it is a continuing way of understanding the world'",
  "- Students with deep questions - capture them for the W column",
  "",
  "[Inquiry: We Do (Explore) | VTLM 2.0: Conceptual Knowledge]",
].join("\n");

const NOTES_AIATSIS = [
  "SAY:",
  "- This is the AIATSIS Map of Indigenous Australia",
  "- Look how many language groups there are across Australia. Hundreds",
  "- Each colour and each name on this map is a different First Nations group, with their own language, their own Country, their own stories",
  "- Find our state. Find an area near our school. What groups can you see?",
  "",
  "DO:",
  "- Open the AIATSIS map from the link on this slide. Project it large",
  "- Zoom in to your school's region. Read 2-3 nearby group names aloud",
  "- Quick partner share: 'How many different groups can you count in our state?'",
  "- Take 3-4 contributions",
  "- Add 'Hundreds of First Nations groups across Australia' to the L column of the class KWL chart",
  "",
  "SOURCES:",
  "- AIATSIS Map of Indigenous Australia. Used as supplied: " + URL_AIATSIS,
  "",
  "TEACHER NOTES:",
  "The map is the visual moment of the lesson. Many students think of First Nations Australians as a single group - the map shows the actual diversity. Hundreds of languages, hundreds of cultures.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: The map shows the boundaries of language groups across Australia.",
  "- Framing language: 'Each group on this map is a different culture, with its own language and Country'.",
  "- Watch for: Students who try to claim or correct group names - redirect: 'This map was made with First Nations communities. We trust their work'.",
  "- Protocol: Do not put First Nations students in the position of speaking for any group on the map. If a student has a personal connection and chooses to share, listen.",
  "",
  "WATCH FOR:",
  "- Students reading group names aloud - praise the careful pronunciation, accept the attempt",
  "- Students who say 'I didn't know there were that many' - capture: that is a key learning",
  "",
  "[Inquiry: We Do (Map) | VTLM 2.0: Visual Evidence]",
].join("\n");

const CFU_Q_TEXT = "For First Nations Australians, Country is more than land. Country also includes ____.";

const NOTES_CFU_Q = [
  "SAY:",
  "- Mini-whiteboards out",
  "- Finish this sentence: 'Country also includes ___'",
  "- You can write more than one word. Use the vocabulary slide if you need to",
  "- Thirty seconds",
  "",
  "DO:",
  "- Set a 30-second timer",
  "- Scan whiteboards as students hold them up",
  "",
  "CFU CHECKPOINT:",
  "Technique: Show Me Boards",
  "Script:",
  "- Say: Boards up on three. One, two, three",
  "- Scan for: family, ancestors, language, stories, identity, water, sky, all of these",
  "PROCEED: If 80% show one of the target ideas, reveal and move on.",
  "PIVOT: If many write only 'land' or 'rocks', say: 'Country is bigger than that. Listen: Country includes the land AND the family, ancestors, language and stories that belong to that land'. Re-check with the same prompt.",
  "",
  "TEACHER NOTES:",
  "This hinge checks the key concept of Country. Wrong answers tell you students are still on the everyday English meaning - one re-explanation is usually enough.",
  "",
  "WATCH FOR:",
  "- Students writing 'animals' or 'plants' - acknowledge: those belong to Country too",
  "- Students stuck for an answer - prompt with: 'family, ancestors, language ...'",
  "",
  "[Inquiry: CFU | VTLM 2.0: Checking for Understanding]",
].join("\n");

const NOTES_CFU_A = [
  "SAY:",
  "- Country includes family, ancestors, language and stories - as well as the land, water and sky",
  "- It is all of these together, not separate parts",
  "- We will keep this idea with us as we compare cultures next",
  "",
  "DO:",
  "- Click to reveal the answer",
  "- Acknowledge close answers ('animals', 'sky', 'stories')",
  "- Move on promptly to the comparison",
  "",
  "TEACHER NOTES:",
  "Validate partial answers so students stay engaged for the comparison. The expanded meaning of Country is the foundation for the next slide's comparison work.",
  "",
  "WATCH FOR:",
  "- Students still rewriting after the reveal - that is fine, that is learning",
  "",
  "[Inquiry: CFU Reveal | VTLM 2.0: Feedback]",
].join("\n");

const NOTES_COMPARE = [
  "SAY:",
  "- Pairs - take a Cultures Compared Chart between two",
  "- Column 1: Indigenous culture - what we have learnt today and last lesson",
  "- Column 2: Non-Indigenous culture - the way the British arrived, settled and built modern Australia",
  "- The middle band: things both cultures share - family, stories, caring for community, art, song",
  "- Then the discussion questions at the bottom: how do the cultures differ, what are the benefits and limitations of each way of living, how has settlement helped and hurt our country",
  "",
  "DO:",
  "- Distribute the Cultures Compared Chart - one per pair",
  "- Set a 12-minute timer",
  "- Circulate. Prompt quiet pairs with one of the discussion questions",
  "- Encourage pairs to add at least ONE thing to the middle 'shared' band - this is the key extension",
  "- Photograph filled charts before they leave the room",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Use the sentence frames on the chart - 'In Indigenous culture ___', 'In non-Indigenous culture ___', 'Both cultures ___'",
  "- Extra Notes: It is fine for an enabling pair to write fewer items if they are well thought through.",
  "EXTENDING PROMPT:",
  "- Task: After filling the chart, write ONE sentence under the chart starting with: 'Something I think non-Indigenous Australians could learn from First Nations culture is ___'.",
  "",
  "TEACHER NOTES:",
  "The chart is the recording tool. The conversation is the lesson. The middle band - what is shared - is essential. Both cultures value family, story, song, art, and caring for place.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: A direct comparison of two cultures.",
  "- Framing language: 'Different is not better or worse. Different is different. Today we look at both with respect'.",
  "- Watch for: Students who frame one culture as 'better' or 'modern' - redirect: 'We are looking at differences, not ranking them'.",
  "- Protocol: Do not put First Nations students in the position of speaking for First Nations culture. The chart is filled from what we have learnt this lesson, not from any one student's family experience.",
  "",
  "WATCH FOR:",
  "- Pairs leaving the middle band empty - prompt: 'What do BOTH cultures value?'",
  "- Pairs writing stereotypes - redirect to what we read in the website together",
  "",
  "[Inquiry: You Do (Compare) | VTLM 2.0: Application]",
].join("\n");

const NOTES_WHATIF = [
  "SAY:",
  "- A big What If question to think about",
  "- What would Australia be like today if Captain Cook had settled peacefully and negotiated to live alongside First Nations peoples?",
  "- This is a serious question, not a fairy tale",
  "- Talk in your pairs - share ONE thing you think might be different now",
  "",
  "DO:",
  "- Read the question slowly",
  "- Set a 4-minute pair-share timer",
  "- Circulate, listen, do not jump in",
  "- Take 4-6 contributions on the class chart - add them to a new 'What if?' line",
  "- Acknowledge the heaviness of the question. It is okay to think there might still be hard parts",
  "",
  "TEACHER NOTES:",
  "This is counterfactual thinking - imagining a different past to understand the present. Students will offer a range of answers. Treat all serious answers respectfully. Do not allow joke answers.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: A reflection on the impact of British colonisation on First Nations Australians.",
  "- Framing language: 'There is no perfect answer. We are thinking carefully about a hard question'.",
  "- Watch for: Students who become upset at the unfairness - acknowledge: 'It is hard. Many people in Australia are still working on what fair looks like'.",
  "- Protocol: If a student wants to step out for a moment, allow it without comment. Check in privately afterwards.",
  "",
  "WATCH FOR:",
  "- Students writing simplistic answers - prompt: 'What about language? What about land? What about Country?'",
  "- Students saying 'we wouldn't be here' - acknowledge thoughtfully: 'That is one of the things people have written about. It is a real thought to have'",
  "",
  "[Inquiry: Reflect | VTLM 2.0: Counterfactual Thinking]",
].join("\n");

const NOTES_OPTIONAL_VIDEOS = [
  "SAY:",
  "- These three videos are optional resources",
  "- They cover the Stolen Generations - Aboriginal children who were forcibly taken from their families by Australian governments through the 1900s",
  "- This is heavy content. We will only watch a clip together if there is time and the class is ready, or use them in a follow-up lesson",
  "",
  "DO:",
  "- If using a clip now: choose ONE only - the BTN clip is the most school-friendly first viewing",
  "- Brief the class before pressing play - see the framing language below",
  "- Watch in full, then pause for partner share",
  "- Otherwise: name the topic, list the resources, and tell students they will be revisited later in the unit",
  "",
  "SOURCES:",
  "- BTN: The Stolen Generation. Used as supplied: " + URL_BTN_STOLEN,
  "- Ask Us Anything (Stolen Generations). Used as supplied: " + URL_ASK_US,
  "- The Stolen Generation: Mona. Used as supplied: " + URL_MONA,
  "",
  "TEACHER NOTES:",
  "These resources are optional for this lesson. The Stolen Generations content needs strong framing and time for response. Many schools handle it in a dedicated lesson with a wellbeing staff member briefed in advance.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: The forcible removal of Aboriginal children from their families - a defined period in Australian government policy that caused lasting harm.",
  "- Framing language: 'Australian governments did this. It was wrong. Today we acknowledge it and we listen to people who lived through it'.",
  "- Watch for: Students who become distressed - the protocol below is essential.",
  "- Protocol: Tell students at the start that they may step out at any point. Check in with First Nations students before and after. Have the school's wellbeing staff aware that the lesson is happening.",
  "",
  "WATCH FOR:",
  "- Students laughing or making jokes - this is usually nervousness, not malice. Pause the video, restate the framing, then continue",
  "- Students who go silent - quiet is a respectful response. Check in privately afterwards",
  "",
  "[Inquiry: Optional Extension | VTLM 2.0: Sensitive Content]",
].join("\n");

const NOTES_KWL_UPDATE = [
  "SAY:",
  "- Last task before our exit ticket",
  "- Open your group's KWL chart from last lesson",
  "- Add at least THREE things to the L column from today - what we learnt about First Nations Australians",
  "- Add at least TWO new wonderings to the W column - questions you now want to find out",
  "- Use full sentences for wonderings - start with what, why, how, who, when",
  "",
  "DO:",
  "- Distribute the Lesson 1 KWL charts back to each group of three (or have students retrieve them from their folders)",
  "- Set a 6-minute timer",
  "- Circulate. Prompt quiet groups: 'What was the most surprising thing today?'",
  "- Photograph each updated chart before students put them away",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Use sentence starters: 'Today we learnt that ___', 'I wonder why ___', 'I wonder how ___'",
  "- Extra Notes: Two well-formed sentences are enough.",
  "EXTENDING PROMPT:",
  "- Task: Write ONE wondering that links two events - for example: 'Why did British settlement change the lives of so many First Nations peoples?'",
  "",
  "TEACHER NOTES:",
  "Updating the KWL is the bridge from this lesson to the next. The chart is the unit's living evidence document - it grows each session.",
  "",
  "WATCH FOR:",
  "- Groups who repeat what they wrote last week - prompt: 'What is NEW from today?'",
  "- Groups who finish early - direct them to the extending task",
  "",
  "[Inquiry: You Do (Update Chart) | VTLM 2.0: Consolidation]",
].join("\n");

const NOTES_EXIT = [
  "SAY:",
  "- One thing on a sticky note before you leave",
  "- Finish this sentence: 'One thing I now understand about First Nations culture that I did not understand before is ___'",
  "- Stick it onto the class chart on your way out",
  "",
  "DO:",
  "- Hand each student a sticky note (or small slip of paper)",
  "- Allow 3 minutes",
  "- Collect by sticking on the L column of the class KWL chart",
  "",
  "TEACHER NOTES:",
  "This exit ticket assesses SC2 - describing or comparing what students have learnt. The sticky note format is fast and the responses stay visible across the term.",
  "",
  "WATCH FOR:",
  "- Students who write 'nothing' - prompt: 'Pick the smallest thing. The map. Country. Dreaming. Any one of them'",
  "- Students who write a single word - that is okay if the word is meaningful (e.g. 'Country')",
  "",
  "[Inquiry: Exit Ticket | VTLM 2.0: Evidence of Learning]",
].join("\n");

const NOTES_CLOSING = [
  "SAY:",
  "- Let's come back to our success criteria. Thumbs up, sideways, or down for each one",
  "- I can name some features of Indigenous and non-Indigenous Australian cultures",
  "- I can describe ways the cultures differ and ways they are alike",
  "- I can explain something we can learn from both ways of living",
  "- Reflection: which moment from today will stay with you the most, and why?",
  "",
  "DO:",
  "- Read each I can statement and pause for thumbs",
  "- Acknowledge the seriousness of today's content. These are real people, real history, real culture",
  "- Tell students: next lesson we will pick up the wonderings we starred today",
  "",
  "TEACHER NOTES:",
  "Close the lesson with care. Students will be processing heavy content. A respectful close lets them carry the learning out of the room. If a student lingers with a question, give them time.",
  "",
  "WATCH FOR:",
  "- Students who show thumbs down on SC3 - flag for a small-group prompt next lesson",
  "- Students who seem reflective - this is engagement, not distress; affirm their thinking",
  "",
  "[Inquiry: Closing | VTLM 2.0: Review and Reflect]",
].join("\n");

// ===============================================================
// Build function
// ===============================================================

async function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.title = "Shaping Australia - Lesson 2 - Aboriginal Culture";
  pres.author = "Year 5/6 Inquiry";

  // -- Slide 1: Title --
  titleSlide(
    pres,
    "Aboriginal Culture",
    "Two cultures, one country",
    "Year 5/6 Inquiry  |  Term 2  |  Lesson 2",
    NOTES_TITLE
  );

  // -- Slide 2: Teacher Resources (per §44) --
  addResourceSlide(pres, {
    resources: RESOURCE_ITEMS,
    studentTools: [
      "Mini-whiteboards (one per student)",
      "Sticky notes for exit ticket",
      "Group KWL Charts from Lesson 1 (returned to each group)",
    ],
    boardSetup: [
      "Class KWL chart from Lesson 1 visible (re-projected photo or original)",
      "Three website tabs pre-loaded on teacher computer",
    ],
    videos: [
      "Optional: BTN - The Stolen Generation",
      "Optional: Ask Us Anything (Stolen Generations)",
      "Optional: The Stolen Generation - Mona",
    ],
    urls: [
      URL_HERITAGE,
      URL_DREAMING,
      URL_AIATSIS,
      URL_BTN_STOLEN,
      URL_ASK_US,
      URL_MONA,
    ],
  }, T, FOOTER, NOTES_RESOURCES);

  // -- Slide 3: Acknowledgement of Country --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Acknowledgement", { color: C.PRIMARY, w: 2.2 });
    addTitle(s, "Acknowledgement of Country");

    // Single full-width respect card
    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP;
    addCard(s, 0.5, cardY, 9, cardH, { strip: C.PRIMARY, fill: C.BG_LIGHT });

    s.addText("We acknowledge the Traditional Owners", {
      x: 0.75, y: cardY + 0.30, w: 8.5, h: 0.55,
      fontSize: 24, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", margin: 0,
    });

    s.addText(
      "of the land on which we are learning today.",
      {
        x: 0.75, y: cardY + 0.90, w: 8.5, h: 0.42,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", margin: 0,
      }
    );

    s.addText(
      "We pay our respects to Elders past and present,",
      {
        x: 0.75, y: cardY + 1.45, w: 8.5, h: 0.42,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", margin: 0,
      }
    );

    s.addText(
      "and to all Aboriginal and Torres Strait Islander peoples.",
      {
        x: 0.75, y: cardY + 1.92, w: 8.5, h: 0.42,
        fontSize: 18, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", margin: 0,
      }
    );

    // Concentric gathering motif - small rings drawn with roundRect (LibreOffice
    // does not render small "oval" shapes reliably; roundRect with rectRadius
    // = half side length is a guaranteed circle across LibreOffice/PowerPoint).
    const symCx = 5;
    const symCy = cardY + 2.85;
    const rings = [
      { d: 0.80, color: C.SECONDARY, fill: null },
      { d: 0.50, color: C.ACCENT, fill: null },
      { d: 0.20, color: null, fill: C.PRIMARY },
    ];
    rings.forEach((r) => {
      const opts = {
        x: symCx - r.d / 2, y: symCy - r.d / 2, w: r.d, h: r.d,
        rectRadius: r.d / 2,
      };
      if (r.fill) {
        opts.fill = { color: r.fill };
      } else {
        opts.fill = { color: C.BG_LIGHT };
        opts.line = { color: r.color, width: 2.2 };
      }
      s.addShape("roundRect", opts);
    });

    s.addText(
      "We thank them for caring for this land for tens of thousands of years.",
      {
        x: 0.75, y: cardY + 3.40, w: 8.5, h: 0.42,
        fontSize: 14, fontFace: FONT_B, color: C.MUTED, italic: true,
        align: "center", margin: 0,
      }
    );

    addFooter(s, FOOTER);
    s.addNotes(NOTES_AOC);
  }

  // -- Slide 4: Launch (Recap from Lesson 1) --
  contentSlide(
    pres,
    "Launch",
    C.SECONDARY,
    "Picking up our wonderings",
    [
      "Last lesson we started a class KWL chart",
      "Some wonderings were about First Nations Australians",
      "Today we go looking for answers",
      "Hold this question: how is First Nations culture different?",
    ],
    NOTES_LAUNCH,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.35;

      // Mini KWL chart preview - emphasising the W column we are extending
      addCard(s, rX, topY, rW, cardH, { strip: C.SECONDARY, fill: C.WHITE });

      s.addText("From Lesson 1", {
        x: rX + 0.15, y: topY + 0.10, w: rW - 0.3, h: 0.30,
        fontSize: 11, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      const tableX = rX + 0.15;
      const tableY = topY + 0.50;
      const tableW = rW - 0.3;
      const tableH = cardH - 0.65;
      const colW = tableW / 3;
      const hdrH = 0.45;

      const headers = [
        { label: "K", color: C.SECONDARY },
        { label: "L", color: C.PRIMARY },
        { label: "W", color: C.ACCENT },
      ];
      headers.forEach((h, i) => {
        s.addShape("rect", {
          x: tableX + i * colW, y: tableY, w: colW, h: hdrH,
          fill: { color: h.color },
        });
        s.addText(h.label, {
          x: tableX + i * colW, y: tableY, w: colW, h: hdrH,
          fontSize: 18, fontFace: FONT_H, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
      });

      const rowY = tableY + hdrH;
      const rowH = tableH - hdrH;
      headers.forEach((_, i) => {
        s.addShape("rect", {
          x: tableX + i * colW, y: rowY, w: colW, h: rowH,
          fill: { color: i === 2 ? "FFF3D6" : "FAF5E8" },
          line: { color: C.MUTED, width: 0.6 },
        });
      });

      // Star marker on W column to signal "today extends this"
      s.addText("*  add today", {
        x: tableX + 2 * colW, y: rowY + 0.10, w: colW, h: 0.30,
        fontSize: 9, fontFace: FONT_B, color: C.ACCENT, italic: true,
        align: "center", margin: 0,
      });
    }
  );

  // -- Slide 5: Vocabulary --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Vocabulary", { color: C.PRIMARY, w: 1.7 });
    addTitle(s, "Four Inquiry Words");

    const vocab = [
      { word: "Indigenous", meaning: "originally from this place - First Nations Australians have lived here for tens of thousands of years", color: C.PRIMARY },
      { word: "non-Indigenous", meaning: "not originally from this place - those who came to Australia from 1788 onwards", color: C.SECONDARY },
      { word: "culture", meaning: "the way a group of people live - language, beliefs, stories, art, and how they care for each other", color: C.ACCENT },
      { word: "Country", meaning: "for First Nations Australians, a word for land plus family, ancestors, language and stories", color: C.SUCCESS },
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
    ["We are learning about Aboriginal culture and how it is different from non-Indigenous culture"],
    [
      "I can name some features of Indigenous and non-Indigenous Australian cultures",
      "I can describe ways the cultures differ and ways they are alike",
      "I can explain something we can learn from both ways of living",
    ],
    NOTES_LI,
    FOOTER
  );

  // -- Slide 7: I Do - Aboriginal Heritage website --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "I Do  -  Read", { color: C.PRIMARY, w: 1.85 });
    addTitle(s, "Indigenous Australians  -  read together");

    // Left: instruction card
    addInstructionCard(s, [
      { role: "header", text: "How we read" },
      { role: "body", text: "Like a picture story book" },
      { role: "body", text: "Teacher reads aloud; class listens" },
      { role: "body", text: "Stop and discuss along the way" },
      { role: "body", text: "Watch for: ONE thing that surprises you" },
    ], {
      x: 0.5, y: CONTENT_TOP, w: 3.5, h: SAFE_BOTTOM - CONTENT_TOP,
      strip: C.PRIMARY, fill: C.WHITE,
      headerColor: C.PRIMARY,
    });

    // Right: website link card with browser-style mockup
    const rX = 4.2;
    const rW = 5.3;
    const rH = SAFE_BOTTOM - CONTENT_TOP;

    addCard(s, rX, CONTENT_TOP, rW, rH, { strip: C.SECONDARY, fill: C.BG_DARK });

    // Browser bar
    const bbY = CONTENT_TOP + 0.25;
    const bbH = 0.40;
    s.addShape("roundRect", {
      x: rX + 0.20, y: bbY, w: rW - 0.40, h: bbH, rectRadius: 0.06,
      fill: { color: C.WHITE },
    });
    // 3 browser dots
    [C.ALERT, C.ACCENT, C.SUCCESS].forEach((dotC, di) => {
      s.addShape("roundRect", {
        x: rX + 0.30 + di * 0.18, y: bbY + 0.10, w: 0.16, h: 0.16, rectRadius: 0.08,
        fill: { color: dotC },
      });
    });
    s.addText("aboriginalheritage.org", {
      x: rX + 0.95, y: bbY, w: rW - 1.30, h: bbH,
      fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL,
      align: "left", valign: "middle", margin: 0,
    });

    // Body mockup - title + paragraph lines
    const bodyY = bbY + bbH + 0.15;
    s.addShape("rect", {
      x: rX + 0.20, y: bodyY, w: rW - 0.40, h: rH - 1.35,
      fill: { color: C.WHITE },
      line: { color: C.MUTED, width: 0.5 },
    });

    s.addText("Indigenous Australians", {
      x: rX + 0.40, y: bodyY + 0.15, w: rW - 0.80, h: 0.40,
      fontSize: 17, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      margin: 0,
    });
    s.addText("History  -  Culture  -  Country", {
      x: rX + 0.40, y: bodyY + 0.55, w: rW - 0.80, h: 0.26,
      fontSize: 11, fontFace: FONT_B, color: C.SECONDARY, italic: true,
      margin: 0,
    });

    // Mockup paragraph lines
    const lineGap = 0.22;
    const lineX = rX + 0.40;
    const lineW = rW - 0.80;
    for (let li = 0; li < 6; li++) {
      const lineLen = li === 5 ? lineW * 0.55 : (li % 2 === 0 ? lineW : lineW * 0.85);
      s.addShape("rect", {
        x: lineX, y: bodyY + 0.95 + li * lineGap, w: lineLen, h: 0.06,
        fill: { color: C.MUTED, transparency: 60 },
        line: { color: C.MUTED, width: 0 },
      });
    }

    // Link strip (clickable)
    const linkY = CONTENT_TOP + rH - 0.62;
    s.addShape("roundRect", {
      x: rX + 0.20, y: linkY, w: rW - 0.40, h: 0.42, rectRadius: 0.06,
      fill: { color: C.ACCENT },
    });
    s.addText("Open Indigenous Australians page", {
      x: rX + 0.20, y: linkY, w: rW - 0.40, h: 0.42,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
      hyperlink: { url: URL_HERITAGE, tooltip: "Open the Aboriginal Heritage history page" },
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_HERITAGE);
  }

  // -- Slide 8: We Do 1 - Aboriginal Dreaming --
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do  -  Explore", { color: C.SECONDARY, w: 2.05 });
    addTitle(s, "Aboriginal Dreaming");

    // Left: concept card explaining Dreaming briefly
    addInstructionCard(s, [
      { role: "header", text: "What is Dreaming?" },
      { role: "body", text: "How the world was made" },
      { role: "body", text: "How land took its shape" },
      { role: "body", text: "How people, animals and plants are connected" },
      { role: "body", text: "Continuing - not just long ago" },
    ], {
      x: 0.5, y: CONTENT_TOP, w: 3.7, h: SAFE_BOTTOM - CONTENT_TOP,
      strip: C.SECONDARY, fill: C.WHITE,
      headerColor: C.SECONDARY,
    });

    // Right: link card
    const rX = 4.4;
    const rW = 5.1;
    const rH = SAFE_BOTTOM - CONTENT_TOP;

    addCard(s, rX, CONTENT_TOP, rW, rH, { strip: C.ACCENT, fill: C.BG_DARK });

    // Decorative concentric symbol (gathering motif - generic, not sacred).
    // Drawn with roundRect because LibreOffice does not render small "oval"
    // shapes reliably; roundRect with rectRadius = half side renders as a
    // circle in both LibreOffice and PowerPoint/Google Slides.
    const cx = rX + rW / 2;
    const cy = CONTENT_TOP + 1.10;
    [
      { d: 1.70, color: C.ACCENT },
      { d: 1.20, color: C.SECONDARY },
      { d: 0.70, color: C.SUCCESS },
    ].forEach((ring) => {
      s.addShape("roundRect", {
        x: cx - ring.d / 2, y: cy - ring.d / 2, w: ring.d, h: ring.d,
        rectRadius: ring.d / 2,
        fill: { color: C.BG_DARK },
        line: { color: ring.color, width: 3.0 },
      });
    });
    s.addShape("roundRect", {
      x: cx - 0.10, y: cy - 0.10, w: 0.20, h: 0.20, rectRadius: 0.10,
      fill: { color: C.WHITE },
    });

    s.addText("a continuing way", {
      x: rX + 0.30, y: CONTENT_TOP + 2.20, w: rW - 0.60, h: 0.36,
      fontSize: 16, fontFace: FONT_H, color: C.WHITE, bold: true,
      align: "center", margin: 0,
    });
    s.addText("of understanding the world", {
      x: rX + 0.30, y: CONTENT_TOP + 2.55, w: rW - 0.60, h: 0.32,
      fontSize: 13, fontFace: FONT_B, color: C.SUBTITLE, italic: true,
      align: "center", margin: 0,
    });

    // Link strip
    const linkY = CONTENT_TOP + rH - 0.62;
    s.addShape("roundRect", {
      x: rX + 0.20, y: linkY, w: rW - 0.40, h: 0.42, rectRadius: 0.06,
      fill: { color: C.ACCENT },
    });
    s.addText("Open Dreaming page", {
      x: rX + 0.20, y: linkY, w: rW - 0.40, h: 0.42,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
      hyperlink: { url: URL_DREAMING, tooltip: "Open Aboriginal Dreaming and the Dreamtime" },
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_DREAMING);
  }

  // -- Slide 9: We Do 2 - AIATSIS Map --
  {
    const s = pres.addSlide();
    addTopBar(s, C.SECONDARY);
    addBadge(s, "We Do  -  Map", { color: C.SECONDARY, w: 1.85 });
    addTitle(s, "Map of Indigenous Australia");

    // Left: instruction card
    addInstructionCard(s, [
      { role: "header", text: "Look for" },
      { role: "body", text: "How many language groups across Australia" },
      { role: "body", text: "Find your state on the map" },
      { role: "body", text: "Find groups near our school" },
      { role: "body", text: "Each colour = a different culture" },
    ], {
      x: 0.5, y: CONTENT_TOP, w: 3.5, h: SAFE_BOTTOM - CONTENT_TOP,
      strip: C.SECONDARY, fill: C.WHITE,
      headerColor: C.SECONDARY,
    });

    // Right: stylised map mockup (NOT a real map - a visual placeholder)
    const rX = 4.2;
    const rW = 5.3;
    const rH = SAFE_BOTTOM - CONTENT_TOP;

    addCard(s, rX, CONTENT_TOP, rW, rH, { strip: C.PRIMARY, fill: C.BG_LIGHT });

    s.addText("AIATSIS Map of Indigenous Australia", {
      x: rX + 0.20, y: CONTENT_TOP + 0.15, w: rW - 0.40, h: 0.32,
      fontSize: 13, fontFace: FONT_H, color: C.PRIMARY, bold: true,
      align: "center", margin: 0,
    });
    s.addText("hundreds of First Nations groups, languages, and Countries", {
      x: rX + 0.20, y: CONTENT_TOP + 0.46, w: rW - 0.40, h: 0.26,
      fontSize: 10, fontFace: FONT_B, color: C.MUTED, italic: true,
      align: "center", margin: 0,
    });

    // Stylised continent shape with coloured patches representing language groups
    const mapX = rX + 0.55;
    const mapY = CONTENT_TOP + 0.85;
    const mapW = rW - 1.10;
    const mapH = rH - 1.85;

    // Continent backdrop
    s.addShape("roundRect", {
      x: mapX, y: mapY, w: mapW, h: mapH, rectRadius: 0.20,
      fill: { color: C.WHITE },
      line: { color: C.PRIMARY, width: 1.5 },
    });

    // Coloured group patches - generic visual, not actual boundaries
    const patches = [
      { rx: 0.05, ry: 0.08, rw: 0.30, rh: 0.32, color: C.PRIMARY },
      { rx: 0.36, ry: 0.05, rw: 0.30, rh: 0.30, color: C.SECONDARY },
      { rx: 0.67, ry: 0.10, rw: 0.28, rh: 0.34, color: C.ACCENT },
      { rx: 0.04, ry: 0.42, rw: 0.32, rh: 0.30, color: C.ACCENT },
      { rx: 0.38, ry: 0.40, rw: 0.30, rh: 0.32, color: C.SUCCESS },
      { rx: 0.70, ry: 0.46, rw: 0.25, rh: 0.30, color: C.PRIMARY },
      { rx: 0.20, ry: 0.74, rw: 0.32, rh: 0.22, color: C.SECONDARY },
      { rx: 0.55, ry: 0.78, rw: 0.30, rh: 0.18, color: C.SUCCESS },
    ];
    patches.forEach((p) => {
      s.addShape("roundRect", {
        x: mapX + p.rx * mapW, y: mapY + p.ry * mapH,
        w: p.rw * mapW, h: p.rh * mapH, rectRadius: 0.08,
        fill: { color: p.color, transparency: 25 },
        line: { color: C.WHITE, width: 1.2 },
      });
    });

    // Link strip
    const linkY = CONTENT_TOP + rH - 0.62;
    s.addShape("roundRect", {
      x: rX + 0.20, y: linkY, w: rW - 0.40, h: 0.42, rectRadius: 0.06,
      fill: { color: C.ACCENT },
    });
    s.addText("Open AIATSIS map", {
      x: rX + 0.20, y: linkY, w: rW - 0.40, h: 0.42,
      fontSize: 12, fontFace: FONT_B, color: C.WHITE, bold: true,
      align: "center", valign: "middle", margin: 0,
      hyperlink: { url: URL_AIATSIS, tooltip: "Open AIATSIS Map of Indigenous Australia" },
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_AIATSIS);
  }

  // -- Slide 10 / 10a: CFU hinge with reveal --
  function buildCfuBase() {
    const s = pres.addSlide();
    addTopBar(s, C.ALERT);
    addBadge(s, "CFU", { color: C.ALERT });
    addTitle(s, "What does Country mean?", { color: C.ALERT });

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

    // Reserve answer-bar zone at the bottom
    const aH = 0.65;
    const aY = SAFE_BOTTOM - aH - 0.05;
    const qY = CONTENT_TOP + 0.56;
    const qH = aY - 0.20 - qY;

    addCard(s, 0.5, qY, 9, qH, { strip: C.ALERT, fill: C.WHITE });
    s.addText(CFU_Q_TEXT, {
      x: 0.75, y: qY + 0.30, w: 8.5, h: qH - 0.60,
      fontSize: 22, fontFace: FONT_B, color: C.CHARCOAL, valign: "middle",
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
      s.addText("family  -  ancestors  -  language  -  stories  -  land, water and sky", {
        x: 0.5, y: aY, w: 9, h: aH,
        fontSize: 16, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addNotes(NOTES_CFU_A);
    }
  );

  // -- Slide 11: You Do - Cultures Compared (pair work with chart) --
  contentSlide(
    pres,
    "You Do  -  Compare",
    C.ACCENT,
    "Cultures Compared  -  with your partner",
    [
      "One Cultures Compared Chart per pair",
      "Column 1: Indigenous culture",
      "Column 2: non-Indigenous culture",
      "Middle band: things BOTH cultures share",
      "Then answer the discussion questions",
    ],
    NOTES_COMPARE,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.35;

      // Comparison chart preview
      addCard(s, rX, topY, rW, cardH, { strip: C.ACCENT, fill: C.WHITE });

      s.addText("Cultures Compared", {
        x: rX + 0.15, y: topY + 0.10, w: rW - 0.3, h: 0.30,
        fontSize: 12, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });

      // Two-column structure with shared band
      const tableX = rX + 0.15;
      const tableY = topY + 0.50;
      const tableW = rW - 0.3;
      const colW = tableW / 2 - 0.05;
      const hdrH = 0.38;

      // Headers
      s.addShape("rect", {
        x: tableX, y: tableY, w: colW, h: hdrH,
        fill: { color: C.PRIMARY },
      });
      s.addText("Indigenous", {
        x: tableX, y: tableY, w: colW, h: hdrH,
        fontSize: 11, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addShape("rect", {
        x: tableX + colW + 0.10, y: tableY, w: colW, h: hdrH,
        fill: { color: C.SECONDARY },
      });
      s.addText("non-Indigenous", {
        x: tableX + colW + 0.10, y: tableY, w: colW, h: hdrH,
        fontSize: 11, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Two body columns with lines
      const bodyY = tableY + hdrH;
      const bodyH = (cardH - 0.65) - hdrH - 0.55;
      [0, 1].forEach((i) => {
        const cx = tableX + i * (colW + 0.10);
        s.addShape("rect", {
          x: cx, y: bodyY, w: colW, h: bodyH,
          fill: { color: "FAF5E8" },
          line: { color: C.MUTED, width: 0.5 },
        });
        const lineCount = 4;
        const lineGap = (bodyH - 0.20) / lineCount;
        for (let li = 0; li < lineCount; li++) {
          const ly = bodyY + 0.20 + li * lineGap;
          s.addShape("line", {
            x: cx + 0.10, y: ly,
            w: colW - 0.20, h: 0,
            line: { color: C.MUTED, width: 0.5 },
          });
        }
      });

      // Shared band at bottom
      const sharedY = bodyY + bodyH + 0.05;
      const sharedH = 0.45;
      s.addShape("rect", {
        x: tableX, y: sharedY, w: tableW, h: sharedH,
        fill: { color: C.SUCCESS },
      });
      s.addText("Both cultures share ...", {
        x: tableX, y: sharedY, w: tableW, h: sharedH,
        fontSize: 11, fontFace: FONT_H, color: C.WHITE, bold: true, italic: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // -- Slide 12: What if? Reflection --
  pairShareSlide(
    pres,
    "A big What If...",
    [
      "What would Australia be like today if Captain Cook had settled peacefully and negotiated to live alongside First Nations peoples?",
      "What would be different about language, land, and Country?",
      "What would still be hard, even in a peaceful version of history?",
    ],
    NOTES_WHATIF,
    FOOTER
  );

  // -- Slide 13: Optional video resources --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "Optional", { color: C.MUTED, w: 1.5 });
    addTitle(s, "Optional: Stolen Generations resources");

    // Sensitivity card across the top
    const sensY = CONTENT_TOP;
    const sensH = 0.85;
    addCard(s, 0.5, sensY, 9, sensH, { strip: C.ALERT, fill: C.BG_LIGHT });
    s.addText("Heavy content. Use only with strong framing and time to respond. The Stolen Generations is sensitive and important Australian history.", {
      x: 0.75, y: sensY + 0.10, w: 8.5, h: sensH - 0.20,
      fontSize: 13, fontFace: FONT_B, color: C.ALERT, italic: true,
      valign: "middle", margin: 0,
    });

    // Three video cards in a row
    const cardY = sensY + sensH + 0.20;
    const cardH = SAFE_BOTTOM - cardY - 0.10;
    const cardW = (9 - 0.30) / 3;
    const videos = [
      { title: "BTN: The Stolen Generation", source: "ABC Education", url: URL_BTN_STOLEN, color: C.PRIMARY },
      { title: "Ask Us Anything", source: "Stolen Generations", url: URL_ASK_US, color: C.SECONDARY },
      { title: "The Stolen Generation: Mona", source: "First-person account", url: URL_MONA, color: C.ACCENT },
    ];

    videos.forEach((v, i) => {
      const cx = 0.5 + i * (cardW + 0.15);
      addCard(s, cx, cardY, cardW, cardH, { strip: v.color, fill: C.BG_DARK });

      // Play triangle
      const playSize = 0.55;
      s.addShape("triangle", {
        x: cx + cardW / 2 - playSize / 2,
        y: cardY + 0.30,
        w: playSize, h: playSize,
        fill: { color: C.WHITE },
        rotate: 90,
      });

      s.addText(v.title, {
        x: cx + 0.15, y: cardY + 1.05, w: cardW - 0.30, h: 0.55,
        fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "top", margin: 0,
        fit: "shrink", shrinkText: true,
      });
      s.addText(v.source, {
        x: cx + 0.15, y: cardY + 1.65, w: cardW - 0.30, h: 0.30,
        fontSize: 10, fontFace: FONT_B, color: C.SUBTITLE, italic: true,
        align: "center", margin: 0,
      });

      // Link strip at bottom of card
      const linkY = cardY + cardH - 0.45;
      s.addShape("roundRect", {
        x: cx + 0.15, y: linkY, w: cardW - 0.30, h: 0.34, rectRadius: 0.05,
        fill: { color: v.color },
      });
      s.addText("Open clip", {
        x: cx + 0.15, y: linkY, w: cardW - 0.30, h: 0.34,
        fontSize: 10, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
        hyperlink: { url: v.url, tooltip: v.title },
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_OPTIONAL_VIDEOS);
  }

  // -- Slide 14: You Do - Update KWL chart from Lesson 1 --
  contentSlide(
    pres,
    "You Do  -  Update",
    C.ACCENT,
    "Add to the chart from last lesson",
    [
      "Get your group's KWL chart from last lesson",
      "Add at least THREE things to the L column",
      "Add at least TWO new wonderings to the W column",
      "Use full sentences - start with what, why, how, who, when",
    ],
    NOTES_KWL_UPDATE,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const cardH = 3.35;

      // Mini chart preview - emphasis on L and W columns
      addCard(s, rX, topY, rW, cardH, { strip: C.ACCENT, fill: C.WHITE });

      s.addText("Group KWL Chart  -  Update", {
        x: rX + 0.15, y: topY + 0.10, w: rW - 0.3, h: 0.30,
        fontSize: 11, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });

      const tableX = rX + 0.15;
      const tableY = topY + 0.50;
      const tableW = rW - 0.3;
      const tableH = cardH - 0.65;
      const colW = tableW / 3;
      const hdrH = 0.45;

      const headers = [
        { label: "K", sub: "we know", color: C.SECONDARY, target: "(no change)" },
        { label: "L", sub: "we learnt", color: C.PRIMARY, target: "+ 3 today" },
        { label: "W", sub: "we wonder", color: C.ACCENT, target: "+ 2 today" },
      ];
      headers.forEach((h, i) => {
        s.addShape("rect", {
          x: tableX + i * colW, y: tableY, w: colW, h: hdrH,
          fill: { color: h.color },
        });
        s.addText(h.label, {
          x: tableX + i * colW, y: tableY, w: colW, h: 0.26,
          fontSize: 13, fontFace: FONT_H, color: C.WHITE, bold: true,
          align: "center", valign: "middle", margin: 0,
        });
        s.addText(h.sub + "  -  " + h.target, {
          x: tableX + i * colW, y: tableY + 0.24, w: colW, h: 0.20,
          fontSize: 7.5, fontFace: FONT_B, color: C.WHITE,
          align: "center", valign: "middle", margin: 0,
        });
      });

      const rowY = tableY + hdrH;
      const rowH = tableH - hdrH;
      headers.forEach((_, i) => {
        // Highlight L and W columns where students are adding
        const isAdd = i === 1 || i === 2;
        s.addShape("rect", {
          x: tableX + i * colW, y: rowY, w: colW, h: rowH,
          fill: { color: isAdd ? "FFF3D6" : "FAF5E8" },
          line: { color: C.MUTED, width: 0.6 },
        });
        const lineCount = 4;
        const lineGap = (rowH - 0.20) / lineCount;
        for (let li = 0; li < lineCount; li++) {
          const ly = rowY + 0.30 + li * lineGap;
          s.addShape("line", {
            x: tableX + i * colW + 0.10, y: ly,
            w: colW - 0.20, h: 0,
            line: { color: C.MUTED, width: 0.5 },
          });
        }
      });
    }
  );

  // -- Slide 15: Exit Ticket --
  exitTicketSlide(
    pres,
    [
      "Sticky note - finish this sentence and stick it on the class chart on your way out:\n\n'One thing I now understand about First Nations culture that I did not understand before is ...'",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2, title: "One thing before you leave" }
  );

  // -- Slide 16: Closing --
  closingSlide(
    pres,
    {
      reflectionPrompt: "Which moment from today will stay with you the most, and why?",
      scItems: [
        "I can name some features of Indigenous and non-Indigenous Australian cultures",
        "I can describe ways the cultures differ and ways they are alike",
        "I can explain something we can learn from both ways of living",
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
  // PDF: Cultures Compared Chart
  // ===============================================================
  await buildComparePdf();
}

async function buildComparePdf() {
  const doc = createPdf({ title: "Cultures Compared Chart" });
  let y = addPdfHeader(doc, "Cultures Compared", {
    subtitle: "Indigenous and non-Indigenous Australia. One per pair.",
    color: C.PRIMARY,
    lessonInfo: "Year 5/6 Inquiry  |  Shaping Australia  |  Lesson 2",
  });

  // Brief instruction
  y = addBodyText(doc,
    "List features of each culture in the two columns. Add at least one thing both cultures share in the band below. Then answer the discussion questions with your partner.",
    y, { fontSize: 10.5, italic: true, color: "4B5563" }
  );

  // Two-column comparison table
  const tableX = PAGE.MARGIN;
  const tableW = PAGE.CONTENT_W;
  const colGap = 12;
  const colW = (tableW - colGap) / 2;
  const hdrH = 28;
  const colH = 200;

  // Indigenous header
  doc.save();
  doc.rect(tableX, y, colW, hdrH).fill(hex(C.PRIMARY));
  doc.fontSize(13).font("Sans-Bold").fillColor("#FFFFFF");
  doc.text("Indigenous culture", tableX + 8, y + 7, { width: colW - 16, align: "left" });
  doc.restore();

  // Non-Indigenous header
  const col2X = tableX + colW + colGap;
  doc.save();
  doc.rect(col2X, y, colW, hdrH).fill(hex(C.SECONDARY));
  doc.fontSize(13).font("Sans-Bold").fillColor("#FFFFFF");
  doc.text("non-Indigenous culture", col2X + 8, y + 7, { width: colW - 16, align: "left" });
  doc.restore();

  // Body cells with writing lines
  const bodyY = y + hdrH;
  doc.save();
  doc.rect(tableX, bodyY, colW, colH).lineWidth(0.8).strokeColor("#9CA3AF").stroke();
  doc.rect(col2X, bodyY, colW, colH).lineWidth(0.8).strokeColor("#9CA3AF").stroke();
  doc.restore();

  // Sentence frame line at top of each column
  doc.fontSize(9).font("Sans-Italic").fillColor(hex("6B7280"));
  doc.text("In Indigenous culture ...", tableX + 8, bodyY + 6, { width: colW - 16 });
  doc.text("In non-Indigenous culture ...", col2X + 8, bodyY + 6, { width: colW - 16 });

  // Writing lines in each column
  const lineCount = 6;
  const lineGap = (colH - 30) / lineCount;
  for (let li = 0; li < lineCount; li++) {
    const ly = bodyY + 30 + li * lineGap;
    doc.save();
    doc.moveTo(tableX + 10, ly).lineTo(tableX + colW - 10, ly)
      .lineWidth(0.6).strokeColor("#000000").stroke();
    doc.moveTo(col2X + 10, ly).lineTo(col2X + colW - 10, ly)
      .lineWidth(0.6).strokeColor("#000000").stroke();
    doc.restore();
  }

  y = bodyY + colH + 12;

  // Shared band
  const sharedH = 70;
  doc.save();
  doc.rect(tableX, y, tableW, hdrH).fill(hex(C.SUCCESS));
  doc.fontSize(13).font("Sans-Bold").fillColor("#FFFFFF");
  doc.text("Both cultures share ...", tableX + 8, y + 7, { width: tableW - 16, align: "left" });
  doc.restore();

  doc.save();
  doc.rect(tableX, y + hdrH, tableW, sharedH).lineWidth(0.8).strokeColor("#9CA3AF").stroke();
  // 2 writing lines
  for (let li = 0; li < 2; li++) {
    const ly = y + hdrH + 24 + li * 24;
    doc.moveTo(tableX + 10, ly).lineTo(tableX + tableW - 10, ly)
      .lineWidth(0.6).strokeColor("#000000").stroke();
  }
  doc.restore();

  y = y + hdrH + sharedH + 16;

  // Discussion questions section
  doc.fontSize(12).font("Sans-Bold").fillColor(hex(C.PRIMARY));
  doc.text("Discuss with your partner", tableX, y, { width: tableW });
  y += 18;

  const questions = [
    "How are Indigenous and non-Indigenous cultures different?",
    "What are the benefits and limitations of each way of living?",
    "How has white settlement helped Australia? How has it hurt?",
  ];
  doc.fontSize(10.5).font("Sans").fillColor(hex("2D3142"));
  questions.forEach((q, i) => {
    doc.text((i + 1) + ".  " + q, tableX, y, { width: tableW });
    y = doc.y + 4;
  });

  y += 6;
  // What if? reflection box
  y = addTipBox(doc,
    "What if? - imagine Australia today if Captain Cook had settled peacefully and negotiated to live alongside First Nations peoples. Write one sentence on the back of this sheet about what would be different.",
    y,
    { color: C.ACCENT }
  );

  addPdfFooter(doc, "Year 5/6 Inquiry  -  Shaping Australia  -  Lesson 2  -  Cultures Compared");
  // COMPARE_CHART.fileName already contains the session folder prefix (it
  // is also used as the relative hyperlink target in the PPTX). The actual
  // write path needs the basename only, joined to RES_DIR.
  const outPath = path.join(RES_DIR, "Session 2 Cultures Compared Chart.pdf");
  await writePdf(doc, outPath);
  console.log("PDF written to", outPath);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
