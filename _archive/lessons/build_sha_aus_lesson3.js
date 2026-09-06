"use strict";

// Inquiry - Shaping Australia, Lesson 3 (Term 2, Week 3)
// Year 5/6 | Captain Cook's voyage and European settlement.
// Continues from Lesson 2 (Aboriginal Culture). Students study why the
// British came, the impact of settlement on First Nations Australians,
// complete a Venn diagram comparing Indigenous and British ways of life,
// then write a letter from an Indigenous person to a British settler.
//
// Sensitive content. Sensitivity Advisory blocks live in every relevant
// slide's teacher notes. First Nations students are never placed in the
// position of speaking on behalf of a community; the letter task uses
// what students have learned in Lessons 1 and 2.

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
// Unit cohesion: Lesson 1 and Lesson 2 used variant 0 (Explorer - olive).
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
const PPTX_NAME = "Lesson 3 - Captain Cook and European Settlement.pptx";
const FOOTER = "Inquiry | Year 5/6 | Shaping Australia | Lesson 3";
const SESSION = 3;
const RES_DIR = path.join(LESSON_FOLDER, getSessionResourceFolder(SESSION));

// -- Resources --
const VENN_CHART = makeSessionResource(
  SESSION,
  "Cultures Venn Diagram",
  "Two-circle Venn diagram comparing First Nations and British ways of life around 1770-1800."
);
const LETTER_SCAFFOLD = makeSessionResource(
  SESSION,
  "Letter Writing Scaffold",
  "Structured A4 scaffold for a letter from an Indigenous person to a British settler. Sentence starters and a planning box."
);
const RESOURCE_ITEMS = [VENN_CHART, LETTER_SCAFFOLD];

fs.mkdirSync(RES_DIR, { recursive: true });

// ===============================================================
// Teacher Notes
// ===============================================================

const NOTES_TITLE = [
  "SAY:",
  "- Welcome back to Shaping Australia",
  "- Last lesson we learnt about First Nations Australians - culture, Country, language groups and Dreaming",
  "- Today we travel back to 1770 and look at what happened when British ships first arrived",
  "- We will also work in pairs to compare two very different ways of life",
  "",
  "DO:",
  "- Display this slide as students enter",
  "- Have the class KWL chart from Lessons 1 and 2 visible",
  "- Have the Venn Diagram and Letter Scaffold printed - one of each per student",
  "",
  "TEACHER NOTES:",
  "Lesson 3 builds on the cultural understanding from Lesson 2 and adds historical context. The day handles sensitive content - read the Sensitivity Advisory sections before teaching.",
  "",
  "WATCH FOR:",
  "- First Nations students - quietly check in before the lesson; offer the option to step out at any point",
  "",
  "[Inquiry: Title | VTLM 2.0: Establishing Purpose]",
].join("\n");

const NOTES_RESOURCES = [
  "SAY:",
  "- Two printed resources today: a Venn diagram and a letter-writing scaffold",
  "- We also bring back the class KWL chart from Lessons 1 and 2",
  "",
  "DO:",
  "- Print the Cultures Venn Diagram - one per student (A4)",
  "- Print the Letter Writing Scaffold - one per student (A4)",
  "- Have the Lesson 1 and Lesson 2 class KWL chart visible",
  "- Mini-whiteboards ready for the CFU check",
  "- Pencils, erasers, and the student inquiry books for additional notes",
  "",
  "TEACHER NOTES:",
  "Both resources are designed for a single 60-minute lesson. The Venn diagram is the comparison record; the scaffold structures the letter so students have a clear frame to write in. Both are referred to on the You Do slides.",
  "",
  "WATCH FOR:",
  "- Resources printed double-sided save paper but make the Venn diagram smaller - print single-sided for clearer student writing space",
  "",
  "[Inquiry: Resources | VTLM 2.0: Preparation]",
].join("\n");

const NOTES_AOC = [
  "SAY:",
  "- Before we begin, we pause and acknowledge",
  "- We acknowledge the Traditional Owners of the land on which we are learning today",
  "- We pay our respects to Elders past and present, and to all Aboriginal and Torres Strait Islander peoples",
  "- We thank them for caring for this land for tens of thousands of years",
  "",
  "DO:",
  "- Read the Acknowledgement aloud, slowly. The class stands quietly",
  "- Replace the bracketed nation name with the local Country if your school's standard wording uses it",
  "- Allow a few seconds of silence after the Acknowledgement before moving on",
  "",
  "TEACHER NOTES:",
  "An Acknowledgement of Country is especially important when the lesson covers colonisation. Keep the tone respectful, not performative.",
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
  "- Last lesson we learnt that hundreds of First Nations groups have lived on this land for tens of thousands of years - each with their own language, Country and culture",
  "- Today we ask: what happened when the British arrived in 1770, and then settled here from 1788?",
  "- Hold this question in your head: how did two very different ways of life meet?",
  "",
  "DO:",
  "- Project the class KWL chart beside this slide",
  "- Read the slide aloud",
  "- Quick partner share: 'From Lesson 2, what is ONE thing you remember about how First Nations Australians lived before 1770?'",
  "- Cold call 2-3 students to share. Star their answer on the class chart if it appears there",
  "",
  "TEACHER NOTES:",
  "The launch reactivates Lesson 2 (Country, language groups, Dreaming) and points it at today's focus. Keep it brisk. The launch is the bridge, not the lesson.",
  "",
  "WATCH FOR:",
  "- Students who say 'I don't remember' - prompt: 'What did the AIATSIS map show us?'",
  "- Students confusing 1770 (Cook) with 1788 (First Fleet) - that is okay, we will clarify in the I Do",
  "",
  "[Inquiry: Launch | VTLM 2.0: Activating Prior Knowledge]",
].join("\n");

const NOTES_VOCAB = [
  "SAY:",
  "- Four inquiry words for today",
  "- Read each word and meaning with me",
  "- 'Settle' and 'colony' are the two big ones - we will use them all lesson",
  "",
  "DO:",
  "- Read each word aloud. Have students repeat 'colony' and 'dispossession' after you",
  "- Point to the small icons as you go",
  "",
  "CFU CHECKPOINT:",
  "Technique: Finger Voting",
  "Script:",
  "- Say: Hold up 1 finger if 'colony' means a new town. Hold up 2 if 'colony' means a settlement controlled by a faraway country",
  "- Scan for: most students showing 2",
  "PROCEED: If 80% show 2, move on.",
  "PIVOT: If many show 1, say: 'A colony is more than a new town. It is a place where one country sends its people, and the colony still belongs to and is run by that faraway country - in our case, Britain'. Re-check.",
  "",
  "TEACHER NOTES:",
  "These four words are the historical vocabulary for the rest of the lesson. We will use 'dispossession' carefully on the impact slide.",
  "",
  "WATCH FOR:",
  "- Students who hear 'dispossession' as too long a word - say it as 'dis-pos-ess-ion' and use the simpler meaning: 'losing your land'",
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
  "Internal tier mapping (do not share with students). SC1 - everyone can give one reason. SC2 - most students will explain impact. SC3 - students compare two ways of life.",
  "",
  "WATCH FOR:",
  "- Students who think 'compare' means 'pick the better one' - clarify: 'Compare means describe how they are different and how they are the same. Not better or worse'",
  "",
  "[Inquiry: LI/SC | VTLM 2.0: Learning Intentions]",
].join("\n");

const NOTES_IDO_COOK = [
  "SAY:",
  "- In April 1770, a British ship called the Endeavour sailed up the east coast of this land",
  "- The captain was James Cook. He was a navigator and a mapmaker",
  "- Cook had two jobs from the British government - to watch the planet Venus cross the sun from Tahiti, and a secret job to look for the great southern land",
  "- On 22 August 1770, Cook stood on an island near the tip of Queensland and claimed the whole eastern coast of this land for Britain",
  "- He called it New South Wales. He did this without asking the First Nations peoples who had lived here for tens of thousands of years",
  "- 18 years later, in January 1788, the First Fleet arrived at Sydney Cove with about 1,500 people - convicts, soldiers, sailors, officers and their families. This was the start of British settlement",
  "",
  "DO:",
  "- Read the slide content slowly",
  "- Point to each year on the timeline as you talk about it",
  "- Stop after the 1788 point and ask: 'Why does it matter that Cook claimed the land without asking?'",
  "- Take 2-3 student responses",
  "",
  "TEACHER NOTES:",
  "Historical accuracy matters here. Cook visited - he did not settle. The First Fleet in 1788 began the colony. Many students conflate the two; this slide separates them.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: The arrival of the British is the beginning of colonisation - a deeply painful history for First Nations peoples.",
  "- Framing language: 'Britain claimed land that already belonged to First Nations peoples. This was the start of a long and painful history'.",
  "- Watch for: Students who frame Cook as a hero or villain - redirect: 'Cook was doing his job. The harder question is about the laws and choices that came next'.",
  "- Protocol: Keep the focus on what happened, not on judging individuals. Today we look at impact, not heroes.",
  "",
  "WATCH FOR:",
  "- Students who ask 'why didn't he ask?' - this is a great question to hold for the impact slide",
  "- Students who confuse 1770 and 1788 - re-anchor: '1770 is Cook visiting. 1788 is the First Fleet settling'",
  "",
  "SOURCES:",
  "- Dates and events: British Library, National Library of Australia, AIATSIS",
  "",
  "[Inquiry: I Do (Voyage) | VTLM 2.0: Direct Instruction]",
].join("\n");

const NOTES_IDO_WHY = [
  "SAY:",
  "- Why did the British want to settle here? Three big reasons",
  "- One - Britain's prisons were full. After the American colonies became independent in 1783, Britain could no longer send convicts to America. They needed somewhere else",
  "- Two - Britain wanted resources and strategic position. A colony in the Pacific gave Britain ships' supplies, a base for trade, and a foothold against other European powers",
  "- Three - the British thought the land was empty. They used a Latin phrase, terra nullius, which means 'land belonging to no one'. This was not true - First Nations peoples had lived here for tens of thousands of years",
  "- These reasons help us answer success criterion one - why the Europeans settled",
  "",
  "DO:",
  "- Read each reason aloud, pausing on each card",
  "- Point to the small icon for each reason",
  "- After reason three, pause and ask: 'Was the land really empty?'",
  "- Take 1-2 student responses",
  "",
  "TEACHER NOTES:",
  "These three reasons are widely accepted by historians. The terra nullius reason is the most important for the rest of the lesson - it links directly to dispossession on the next slide.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: 'Terra nullius' was a legal fiction used to justify taking land that already belonged to First Nations peoples.",
  "- Framing language: 'The British called the land empty. It was not. Hundreds of First Nations groups had lived here for tens of thousands of years'.",
  "- Watch for: Students who think 'terra nullius' was a fair reason - clarify: 'It was a wrong idea, and the High Court later said so in 1992 (the Mabo decision)'.",
  "- Protocol: Mention Mabo briefly if students are ready, but do not derail the lesson - it is the focus of a later unit.",
  "",
  "WATCH FOR:",
  "- Students saying 'so the British just took it?' - acknowledge: 'Yes. That is one of the hard truths of this history'",
  "- Students who want the timeline of Mabo - park it: 'Great question. We come back to it later in the unit'",
  "",
  "[Inquiry: I Do (Why) | VTLM 2.0: Conceptual Knowledge]",
].join("\n");

const NOTES_IDO_IMPACT = [
  "SAY:",
  "- When the British settled here, the impact on First Nations Australians was huge - and most of it was harm",
  "- Four big impacts to know today",
  "- One - dispossession. First Nations peoples lost their land, their Country, and with it their food, water and sacred places",
  "- Two - disease. The British brought illnesses like smallpox that First Nations peoples had no immunity to. Around half the people in the Sydney area died from smallpox within a year of the First Fleet arriving",
  "- Three - violence. Many First Nations peoples were killed in conflicts over land - historians call this period the Frontier Wars",
  "- Four - loss of language and culture. Children were later removed from their families - we will learn more about this in a future lesson called the Stolen Generations",
  "- These impacts answer success criterion two",
  "",
  "DO:",
  "- Read each impact aloud, pausing on each card",
  "- After the four impacts, pause for a moment of silence to acknowledge the weight of this content",
  "- Quick pair share: 'Which of these four impacts feels biggest to you, and why?'",
  "- Take 2-3 student responses",
  "",
  "TEACHER NOTES:",
  "This is the heaviest slide of the lesson. Read it carefully, calmly, and without rushing. The four impacts are well documented by historians. The numbers used (around half of the Sydney area dying from smallpox) are accepted historical estimates.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: The harm caused by British settlement to First Nations Australians, including loss of life, land, language and family.",
  "- Framing language: 'This is real Australian history. It is hard. We learn it so we understand'.",
  "- Watch for: Students who go quiet, students who become upset, students who try to make jokes. Each needs a different response.",
  "- Protocol: Tell students before the slide: 'This is the hardest part of today. If you need a moment, signal me'. Offer a desk task or time outside the room if a student needs it. Have a wellbeing team member aware.",
  "",
  "WATCH FOR:",
  "- First Nations students - check in privately during or after the lesson",
  "- Students who say 'that's not fair' - acknowledge: 'You are right. It was not fair'",
  "- Students who say 'why don't we learn this earlier' - acknowledge: 'A lot of Australians are asking the same question'",
  "",
  "SOURCES:",
  "- Frontier Wars and smallpox impact: AIATSIS, Australian War Memorial, ABC Education resources",
  "",
  "[Inquiry: I Do (Impact) | VTLM 2.0: Direct Instruction]",
].join("\n");

const CFU_Q_TEXT = "The British used the Latin phrase terra nullius to justify settling here. What does this phrase mean - and why was it wrong?";

const NOTES_CFU_Q = [
  "SAY:",
  "- Mini-whiteboards out",
  "- Write two short answers: what terra nullius means, and one reason it was wrong",
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
  "- Scan for: 'land belonging to no one' AND 'First Nations peoples had lived here for tens of thousands of years' (or similar)",
  "PROCEED: If 80% show both ideas, reveal and move on to the Venn diagram.",
  "PIVOT: If many show only one half, say: 'You have part of it. Listen: terra nullius is the meaning AND it was wrong because the land was not empty. Both halves matter'. Re-check with the same prompt.",
  "",
  "TEACHER NOTES:",
  "This hinge checks the threshold concept. Without understanding terra nullius and why it was wrong, the comparison and the letter task lose their meaning.",
  "",
  "WATCH FOR:",
  "- Students who only write the translation - prompt for the second half",
  "- Students who only write 'because First Nations people lived here' - prompt for the translation",
  "",
  "[Inquiry: CFU | VTLM 2.0: Checking for Understanding]",
].join("\n");

const NOTES_CFU_A = [
  "SAY:",
  "- Terra nullius means 'land belonging to no one'",
  "- It was wrong because First Nations peoples had lived here for tens of thousands of years - they had Country, language groups, songlines, laws and connections to the land",
  "- This idea is the key to today's lesson - keep it with you as we move to the comparison",
  "",
  "DO:",
  "- Click to reveal the answer",
  "- Acknowledge close answers",
  "- Move on promptly to the Venn diagram",
  "",
  "TEACHER NOTES:",
  "The reveal anchors the rest of the lesson. The Venn diagram and the letter both rely on students holding this idea.",
  "",
  "WATCH FOR:",
  "- Students still updating their boards after the reveal - that is fine, that is learning",
  "",
  "[Inquiry: CFU Reveal | VTLM 2.0: Feedback]",
].join("\n");

const NOTES_WE_DO_VENN = [
  "SAY:",
  "- We are going to compare two ways of life - First Nations and British - around 1770 to 1800",
  "- A Venn diagram has two circles that overlap. The left circle is what is unique to First Nations life. The right circle is what is unique to British life. The middle is what BOTH groups did",
  "- Watch me - I will model one entry in each section together",
  "",
  "DO:",
  "- Project the Venn diagram from the You Do slide (or the printed version)",
  "- Think aloud: 'Where do songlines belong? Those are unique to First Nations - I'll write that on the left'",
  "- Think aloud: 'Where do steam ships belong? Those are British technology - I'll write that on the right'",
  "- Think aloud: 'Where do families and storytelling belong? Both groups had families and told stories - I'll write that in the middle'",
  "- Pause: 'Notice the middle band. Both cultures shared things. That is important - we look for differences AND similarities'",
  "",
  "TEACHER NOTES:",
  "This We Do is short - one entry in each region. The aim is to model the THINKING, not fill the diagram. Students will fill their own diagrams in the You Do.",
  "",
  "WATCH FOR:",
  "- Students who put everything in the outer circles - prompt for the middle band",
  "- Students who only fill the left circle - prompt: 'What did British people do too?'",
  "",
  "[Inquiry: We Do (Model) | VTLM 2.0: Modelling]",
].join("\n");

const NOTES_YOU_DO_VENN = [
  "SAY:",
  "- Take your Venn Diagram sheet",
  "- Left circle: things unique to First Nations life around 1770 to 1800 - what we learnt in Lessons 1 and 2",
  "- Right circle: things unique to British life around 1770 to 1800",
  "- Middle: things BOTH groups did",
  "- Aim for at least three entries in each region. Use the sentence bank at the bottom of the sheet if you need help",
  "- You have 10 minutes - work with your shoulder partner",
  "",
  "DO:",
  "- Distribute the Cultures Venn Diagram - one per student",
  "- Set a 10-minute timer",
  "- Circulate. Prompt quiet pairs with: 'What did First Nations peoples have that British people did not?'",
  "- Encourage at least one entry in the middle band - that is the key extension",
  "- Photograph filled charts before they leave the room",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Use the sentence bank on the sheet - circle the ones that fit each region, then write them in",
  "- Extra Notes: A pair can complete the diagram with only sentence-bank items if needed.",
  "EXTENDING PROMPT:",
  "- Task: After filling the diagram, write one sentence under it starting with: 'The biggest difference between the two ways of life was ___ because ___'.",
  "",
  "TEACHER NOTES:",
  "The Venn diagram is the recording tool for SC3 (compare). The conversation between pair partners is where the learning happens. The middle band is the most important - it teaches that both cultures shared core human values, even when their daily life looked very different.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: A direct comparison of two cultures during the first decades of British settlement.",
  "- Framing language: 'Different is not better or worse. We are looking at differences and similarities, not ranking them'.",
  "- Watch for: Stereotypes - 'they hunted, we farmed' framing without nuance. Redirect to specifics from Lessons 1 and 2: 'songlines, Country, language groups, Dreaming'.",
  "- Protocol: Do not place First Nations students in the position of speaking for First Nations culture. The diagram is filled from what we have learnt in the unit, not from any one student's family experience.",
  "",
  "WATCH FOR:",
  "- Pairs leaving the middle empty - prompt: 'What do BOTH cultures value?'",
  "- Pairs writing stereotypes - redirect to specifics from Lessons 1 and 2",
  "",
  "[Inquiry: You Do (Compare) | VTLM 2.0: Application]",
].join("\n");

const NOTES_WE_DO_LETTER = [
  "SAY:",
  "- Now we use what we know to write a letter",
  "- The letter is from an Indigenous person in 1788 - the year the First Fleet arrived",
  "- The letter is to a British settler - explaining why they should not settle here",
  "- This is NOT speaking on behalf of any real person or any real community. It is a writing task where we use what we have learnt to imagine a respectful conversation that, in real history, did not happen",
  "- Watch me model the opening",
  "",
  "DO:",
  "- Project the letter scaffold from the You Do slide (or the printed version)",
  "- Model only the OPENING line. Use the frame: 'Dear settler, we have lived on this land for tens of thousands of years'",
  "- Think aloud: 'I'm starting with a fact we know is true. I'm not putting words in anyone's mouth - I'm using what we learnt'",
  "- Optional model paragraph 2 starter: 'This land is our Country. Country means family, ancestors, language and stories'",
  "- Stop modelling there - students take it from this point",
  "",
  "TEACHER NOTES:",
  "The framing on this task is essential. We are using historical learning to write a persuasive letter. We are not pretending to speak for First Nations peoples, and we are not asking First Nations students to do so. The scaffold limits the form so students stay on the content.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: A persuasive writing task in the imagined voice of an Indigenous person speaking to a British settler.",
  "- Framing language: 'We use what we know. We do not pretend to be a real person. We write respectfully, using what Lessons 1 and 2 taught us'.",
  "- Watch for: Students writing in stereotypical or disrespectful voice; students using offensive 'olden days' English. Redirect: 'Use what we learnt. Country. Songlines. Hunting and gathering. Dreaming. Tens of thousands of years. Family. Stories'.",
  "- Protocol: Offer First Nations students the alternative task of writing in their own voice as a modern student rather than in an imagined 1788 voice. Either is acceptable - both meet SC2.",
  "",
  "WATCH FOR:",
  "- Students who copy the opening without adapting - prompt: 'Use a new fact from Lesson 2 in your next sentence'",
  "- Students who start with the British settler - re-anchor: 'The letter is FROM the Indigenous person'",
  "",
  "[Inquiry: We Do (Model) | VTLM 2.0: Modelling]",
].join("\n");

const NOTES_YOU_DO_LETTER = [
  "SAY:",
  "- Take your Letter Writing Scaffold",
  "- Use the four paragraph starters - they help you stay on track",
  "- Paragraph 1 - who you are and how long your people have lived here",
  "- Paragraph 2 - what Country means and what you would lose",
  "- Paragraph 3 - one impact you fear (loss of land, family, language, or life)",
  "- Paragraph 4 - what you ask the settler to do instead",
  "- You have 15 minutes. Quality over length",
  "",
  "DO:",
  "- Distribute the Letter Writing Scaffold - one per student",
  "- Set a 15-minute timer",
  "- Circulate. Prompt quiet students with: 'What is one thing from Lesson 2 you could add here?'",
  "- Look for: respect, accuracy, use of vocabulary words (Country, Dreaming, dispossession), persuasive structure",
  "",
  "ENABLING & EXTENDING:",
  "ENABLING PROMPT:",
  "- Task: Fill in only the four sentence starters at the top of each paragraph. One sentence per paragraph is enough",
  "- Extra Notes: Sit close to the teacher group for one-on-one prompting if needed.",
  "EXTENDING PROMPT:",
  "- Task: After the four paragraphs, add a fifth paragraph that imagines a peaceful alternative - 'If we shared this land respectfully ___'",
  "",
  "TEACHER NOTES:",
  "This task assesses SC2 - the impact of European settlement on First Nations peoples. Look for accurate use of content from Lessons 1 and 2 - that is the evidence of learning. Sophistication of language is a bonus, not the target.",
  "",
  "SENSITIVITY ADVISORY:",
  "- What it is: Persuasive writing in the imagined voice of an Indigenous person around 1788.",
  "- Framing language: 'You are using what we learnt to write a respectful, persuasive letter. You are not speaking for any real person'.",
  "- Watch for: Disrespectful tone, stereotypes, offensive 'olden days' English. Redirect to the content of Lessons 1 and 2.",
  "- Protocol: Offer the alternative task (modern voice writing about why the impact still matters today) for any student - especially First Nations students - who prefers it.",
  "",
  "WATCH FOR:",
  "- Students stuck on paragraph 1 - prompt: 'Just write the first sentence and keep going'",
  "- Students writing very short answers - prompt: 'Add one fact from Lesson 1 or 2'",
  "",
  "[Inquiry: You Do (Letter) | VTLM 2.0: Application]",
].join("\n");

const NOTES_EXIT = [
  "SAY:",
  "- Sticky note - one sentence before you leave",
  "- Finish this sentence: 'The biggest impact of British settlement on First Nations peoples was ___ because ___'",
  "- Stick it on the class chart on your way out",
  "",
  "DO:",
  "- Hand each student a sticky note (or small slip of paper)",
  "- Allow 3 minutes",
  "- Collect by sticking on the class chart - L column",
  "",
  "TEACHER NOTES:",
  "This exit ticket assesses SC2 directly - the impact of European settlement on the Indigenous population. Look for accurate impacts (dispossession, disease, violence, loss of language) and a 'because' that links to specific content from the lesson.",
  "",
  "WATCH FOR:",
  "- Students who write only 'land' - prompt: 'Add a because'",
  "- Students who write 'they died' - acknowledge, then prompt for specificity: 'From what? Disease, conflict, both?'",
  "",
  "[Inquiry: Exit Ticket | VTLM 2.0: Evidence of Learning]",
].join("\n");

const NOTES_CLOSING = [
  "SAY:",
  "- Let's come back to our success criteria. Thumbs up, sideways, or down for each",
  "- I can explain why the Europeans settled in Australia",
  "- I can explain the impact that the Europeans had on the Indigenous population",
  "- I can make comparisons between the way that the Indigenous and Europeans lived",
  "- Reflection: which part of today's lesson will stay with you the most, and why?",
  "",
  "DO:",
  "- Read each I can statement and pause for thumbs",
  "- Acknowledge the seriousness of today's content",
  "- Tell students: next lesson we will pick up the wonderings from today's letters",
  "",
  "TEACHER NOTES:",
  "Close the lesson with care. Students will be processing heavy content. A respectful close lets them carry the learning out of the room.",
  "",
  "WATCH FOR:",
  "- Students who show thumbs down on SC3 - flag for a small-group prompt next lesson",
  "- Students who linger with questions - give them time",
  "",
  "[Inquiry: Closing | VTLM 2.0: Review and Reflect]",
].join("\n");

// ===============================================================
// Build function
// ===============================================================

async function build() {
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_16x9";
  pres.title = "Shaping Australia - Lesson 3 - Captain Cook and European Settlement";
  pres.author = "Year 5/6 Inquiry";

  // -- Slide 1: Title --
  titleSlide(
    pres,
    "Captain Cook and Settlement",
    "Two ways of life meet",
    "Year 5/6 Inquiry  |  Term 2  |  Lesson 3",
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
      "Class KWL chart from Lessons 1 and 2 visible",
      "Venn Diagram and Letter Scaffold printed - one of each per student",
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
      x: 0.85, y: cardY + 0.45, w: 8.3, h: 0.85,
      fontSize: 22, fontFace: FONT_H, color: C.CHARCOAL, bold: true,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
    s.addText("We pay our respects to Elders past and present, and to all Aboriginal and Torres Strait Islander peoples.", {
      x: 0.85, y: cardY + 1.45, w: 8.3, h: 0.85,
      fontSize: 17, fontFace: FONT_B, color: C.CHARCOAL,
      align: "center", valign: "middle", margin: 0,
      fit: "shrink", shrinkText: true,
    });
    s.addText("We thank them for caring for this land for tens of thousands of years.", {
      x: 0.85, y: cardY + 2.45, w: 8.3, h: 0.75,
      fontSize: 17, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
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
    addTitle(s, "From Lesson 2 to today");

    // Left: prior knowledge card
    addInstructionCard(s, [
      { role: "header", text: "We already know" },
      { role: "body", text: "First Nations Australians have lived here for tens of thousands of years" },
      { role: "body", text: "Hundreds of language groups, each with their own Country" },
      { role: "body", text: "Country means land plus family, ancestors, language and stories" },
      { role: "body", text: "Dreaming explains how the world was made and is a continuing way of understanding" },
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
    s.addText("What happened when the British arrived in 1770 and settled here from 1788?", {
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
    s.addText("From Lesson 2, what is ONE thing you remember about how First Nations Australians lived before 1770?", {
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
      { word: "voyage", meaning: "a long journey by ship - Cook's voyage took him from England to the Pacific and east coast of Australia", color: C.PRIMARY },
      { word: "settle / settlement", meaning: "to make a new home in a place - the British settled here from 1788 with the First Fleet", color: C.SECONDARY },
      { word: "colony", meaning: "a settlement controlled by a faraway country - New South Wales was a British colony", color: C.ACCENT },
      { word: "dispossession", meaning: "losing your land and what belongs to you - First Nations peoples were dispossessed of their Country", color: C.SUCCESS },
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
    ["We are learning about Captain Cook's voyage to Australia and what happened when the Europeans settled here"],
    [
      "I can explain why the Europeans settled in Australia",
      "I can explain the impact that the Europeans had on the Indigenous population",
      "I can make comparisons between the way that the Indigenous and Europeans lived",
    ],
    NOTES_LI,
    FOOTER
  );

  // -- Slide 7: I Do - Cook's voyage timeline --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "I Do  -  Voyage", { color: C.PRIMARY, w: 1.85 });
    addTitle(s, "Captain Cook to First Fleet");

    // Timeline strip across the slide
    const tlY = CONTENT_TOP + 0.30;
    const tlX = 0.7;
    const tlW = 8.6;
    const tlH = 0.08;

    s.addShape("rect", {
      x: tlX, y: tlY, w: tlW, h: tlH,
      fill: { color: C.PRIMARY },
    });

    // x positions chosen so the leftmost and rightmost labels stay on-slide.
    const points = [
      { x: 0.08, label: "1768", note: "Endeavour leaves England", color: C.SECONDARY },
      { x: 0.34, label: "April 1770", note: "Cook reaches the east coast", color: C.PRIMARY },
      { x: 0.60, label: "22 Aug 1770", note: "Cook claims east coast for Britain", color: C.ALERT },
      { x: 0.92, label: "Jan 1788", note: "First Fleet arrives at Sydney Cove", color: C.ACCENT },
    ];

    const labelW = 1.60;
    const noteW = 1.95;

    points.forEach((p) => {
      const px = tlX + p.x * tlW;
      const py = tlY + tlH / 2;

      // Dot (roundRect for LibreOffice compatibility)
      s.addShape("roundRect", {
        x: px - 0.14, y: py - 0.14, w: 0.28, h: 0.28, rectRadius: 0.14,
        fill: { color: p.color },
        line: { color: C.WHITE, width: 1.5 },
      });

      // Year label below
      s.addText(p.label, {
        x: px - labelW / 2, y: py + 0.20, w: labelW, h: 0.30,
        fontSize: 13, fontFace: FONT_B, color: p.color, bold: true,
        align: "center", margin: 0,
      });

      // Event note below year
      s.addText(p.note, {
        x: px - noteW / 2, y: py + 0.52, w: noteW, h: 0.55,
        fontSize: 11, fontFace: FONT_B, color: C.CHARCOAL,
        align: "center", valign: "top", margin: 0,
      });
    });

    // Bottom: key idea card
    const keyY = CONTENT_TOP + 2.25;
    const keyH = SAFE_BOTTOM - keyY - 0.05;
    addCard(s, 0.5, keyY, 9, keyH, { strip: C.ALERT, fill: C.BG_LIGHT });

    s.addText("Key idea", {
      x: 0.75, y: keyY + 0.10, w: 5, h: 0.28,
      fontSize: 11, fontFace: FONT_B, color: C.ALERT, bold: true, margin: 0,
    });
    s.addText("Cook visited in 1770 and claimed the eastern coast for Britain - he did not settle. The First Fleet in 1788 began the British colony.", {
      x: 0.75, y: keyY + 0.42, w: 8.5, h: keyH - 0.50,
      fontSize: 17, fontFace: FONT_B, color: C.CHARCOAL,
      valign: "top", margin: 0,
      fit: "shrink", shrinkText: true,
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_COOK);
  }

  // -- Slide 8: I Do - Why did the British settle here? --
  {
    const s = pres.addSlide();
    addTopBar(s, C.PRIMARY);
    addBadge(s, "I Do  -  Why?", { color: C.PRIMARY, w: 1.65 });
    addTitle(s, "Why did the British settle here?");

    const reasons = [
      {
        n: "1",
        title: "Full prisons",
        body: "After 1783, Britain could no longer send convicts to America. Their prisons were overflowing.",
        color: C.PRIMARY,
      },
      {
        n: "2",
        title: "Resources & strategy",
        body: "A Pacific colony gave Britain ships' supplies, trade, and a base against rival European powers.",
        color: C.SECONDARY,
      },
      {
        n: "3",
        title: "Terra nullius",
        body: "The British called the land 'belonging to no one'. This was not true - First Nations peoples had lived here for tens of thousands of years.",
        color: C.ALERT,
      },
    ];

    const cardY = CONTENT_TOP;
    const cardH = SAFE_BOTTOM - CONTENT_TOP - 0.05;
    const cardW = (9 - 0.30) / 3;

    reasons.forEach((r, i) => {
      const x = 0.5 + i * (cardW + 0.15);
      addCard(s, x, cardY, cardW, cardH, { strip: r.color, fill: C.WHITE });

      // Big number badge at top
      s.addShape("roundRect", {
        x: x + (cardW - 0.55) / 2, y: cardY + 0.18, w: 0.55, h: 0.55, rectRadius: 0.28,
        fill: { color: r.color },
      });
      s.addText(r.n, {
        x: x + (cardW - 0.55) / 2, y: cardY + 0.18, w: 0.55, h: 0.55,
        fontSize: 26, fontFace: FONT_H, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Title
      s.addText(r.title, {
        x: x + 0.2, y: cardY + 0.85, w: cardW - 0.4, h: 0.45,
        fontSize: 17, fontFace: FONT_H, color: r.color, bold: true,
        align: "center", valign: "middle", margin: 0,
        fit: "shrink",
      });

      // Body
      s.addText(r.body, {
        x: x + 0.2, y: cardY + 1.40, w: cardW - 0.4, h: cardH - 1.55,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
        align: "left", valign: "top", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_WHY);
  }

  // -- Slide 9: I Do - Impact on First Nations Australians --
  {
    const s = pres.addSlide();
    addTopBar(s, C.ALERT);
    addBadge(s, "I Do  -  Impact", { color: C.ALERT, w: 1.85 });
    addTitle(s, "Impact on First Nations Australians", { color: C.ALERT });

    // Sensitivity strip across the top
    const sensY = CONTENT_TOP;
    const sensH = 0.50;
    s.addShape("roundRect", {
      x: 0.5, y: sensY, w: 9, h: sensH, rectRadius: 0.06,
      fill: { color: C.ALERT },
    });
    s.addText("This is real history. Some parts are hard. Listen carefully.", {
      x: 0.5, y: sensY, w: 9, h: sensH,
      fontSize: 13, fontFace: FONT_B, color: C.WHITE, bold: true, italic: true,
      align: "center", valign: "middle", margin: 0,
    });

    const impacts = [
      { title: "Dispossession", body: "Loss of land, Country, sacred places, food and water sources.", color: C.PRIMARY },
      { title: "Disease", body: "Smallpox and other illnesses brought by the British killed around half the people around Sydney in the first year.", color: C.SECONDARY },
      { title: "Violence", body: "Many First Nations peoples were killed in conflicts over land - historians call this the Frontier Wars.", color: C.ALERT },
      { title: "Loss of culture", body: "Language, songs, ceremonies and family connections were broken. Children were later removed.", color: C.ACCENT },
    ];

    const gridY = sensY + sensH + 0.15;
    const gridH = SAFE_BOTTOM - gridY - 0.05;
    const cardW = (9 - 0.15) / 2;
    const cardH = (gridH - 0.15) / 2;

    impacts.forEach((imp, i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 0.5 + col * (cardW + 0.15);
      const y = gridY + row * (cardH + 0.15);

      addCard(s, x, y, cardW, cardH, { strip: imp.color, fill: C.WHITE });
      s.addText(imp.title, {
        x: x + 0.2, y: y + 0.12, w: cardW - 0.4, h: 0.40,
        fontSize: 18, fontFace: FONT_H, color: imp.color, bold: true, margin: 0,
        fit: "shrink",
      });
      s.addText(imp.body, {
        x: x + 0.2, y: y + 0.58, w: cardW - 0.4, h: cardH - 0.70,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL,
        valign: "top", margin: 0,
      });
    });

    addFooter(s, FOOTER);
    s.addNotes(NOTES_IDO_IMPACT);
  }

  // -- Slide 10 / 10a: CFU hinge with reveal --
  function buildCfuBase() {
    const s = pres.addSlide();
    addTopBar(s, C.ALERT);
    addBadge(s, "CFU", { color: C.ALERT });
    addTitle(s, "Why was 'terra nullius' wrong?", { color: C.ALERT });

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
      s.addText("'land belonging to no one'  -  wrong because First Nations peoples had lived here for tens of thousands of years", {
        x: 0.5, y: aY, w: 9, h: aH,
        fontSize: 14, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
        fit: "shrink", shrinkText: true,
      });
      s.addNotes(NOTES_CFU_A);
    }
  );

  // -- Slide 11: We Do - Model the Venn diagram --
  contentSlide(
    pres,
    "We Do  -  Model",
    C.SECONDARY,
    "Comparing two ways of life",
    [
      "Left circle: unique to First Nations life",
      "Right circle: unique to British life",
      "Middle: things BOTH groups did",
      "Watch the teacher model one entry in each region",
    ],
    NOTES_WE_DO_VENN,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const panelH = SAFE_BOTTOM - topY - 0.05;

      addCard(s, rX, topY, rW, panelH, { strip: C.SECONDARY, fill: C.BG_LIGHT });

      s.addText("Venn diagram", {
        x: rX + 0.2, y: topY + 0.12, w: rW - 0.4, h: 0.30,
        fontSize: 12, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      // Two overlapping circles (using roundRect with rectRadius = side/2)
      const cD = 2.20;
      const cy = topY + panelH * 0.50;
      const overlap = 0.55;
      const leftCx = rX + rW / 2 - (cD / 2 - overlap / 2);
      const rightCx = rX + rW / 2 + (cD / 2 - overlap / 2);

      s.addShape("roundRect", {
        x: leftCx - cD / 2, y: cy - cD / 2, w: cD, h: cD, rectRadius: cD / 2,
        fill: { color: C.PRIMARY, transparency: 65 },
        line: { color: C.PRIMARY, width: 1.5 },
      });
      s.addShape("roundRect", {
        x: rightCx - cD / 2, y: cy - cD / 2, w: cD, h: cD, rectRadius: cD / 2,
        fill: { color: C.SECONDARY, transparency: 65 },
        line: { color: C.SECONDARY, width: 1.5 },
      });

      // Labels above the circles (narrowed to avoid overlap - centres are ~1.65" apart)
      s.addText("First Nations", {
        x: leftCx - 0.70, y: cy - cD / 2 - 0.35, w: 1.40, h: 0.30,
        fontSize: 11, fontFace: FONT_B, color: C.PRIMARY, bold: true,
        align: "center", margin: 0,
      });
      s.addText("British", {
        x: rightCx - 0.70, y: cy - cD / 2 - 0.35, w: 1.40, h: 0.30,
        fontSize: 11, fontFace: FONT_B, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      // Sample entries inside each region
      s.addText("songlines", {
        x: leftCx - cD / 2 + 0.10, y: cy - 0.16, w: cD / 2 - 0.15, h: 0.30,
        fontSize: 10, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText("steam ships", {
        x: rightCx - 0.10, y: cy - 0.16, w: cD / 2 - 0.15, h: 0.30,
        fontSize: 10, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText("BOTH", {
        x: (leftCx + rightCx) / 2 - 0.50, y: cy - 0.30, w: 1.0, h: 0.25,
        fontSize: 9, fontFace: FONT_B, color: C.SUCCESS, bold: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText("families", {
        x: (leftCx + rightCx) / 2 - 0.50, y: cy - 0.04, w: 1.0, h: 0.25,
        fontSize: 9, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", valign: "middle", margin: 0,
      });
      s.addText("stories", {
        x: (leftCx + rightCx) / 2 - 0.50, y: cy + 0.18, w: 1.0, h: 0.25,
        fontSize: 9, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "center", valign: "middle", margin: 0,
      });
    }
  );

  // -- Slide 12: You Do - Complete the Venn diagram --
  contentSlide(
    pres,
    "You Do  -  Compare",
    C.ACCENT,
    "Complete your Venn diagram",
    [
      "One Venn Diagram sheet per student",
      "At least 3 entries in each region",
      "Don't forget the middle - what BOTH groups did",
      "Use the sentence bank at the bottom if you need help",
      "10 minutes - work with your shoulder partner",
    ],
    NOTES_YOU_DO_VENN,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const panelH = SAFE_BOTTOM - topY - 0.05;

      addCard(s, rX, topY, rW, panelH, { strip: C.ACCENT, fill: C.WHITE });

      s.addText("Your sheet", {
        x: rX + 0.2, y: topY + 0.10, w: rW - 0.4, h: 0.28,
        fontSize: 11, fontFace: FONT_H, color: C.ACCENT, bold: true,
        align: "center", margin: 0,
      });

      // Mock A4 sheet preview
      const sheetX = rX + 0.30;
      const sheetW = rW - 0.60;
      const sheetY = topY + 0.45;
      const sheetH = panelH - 0.55;
      s.addShape("rect", {
        x: sheetX, y: sheetY, w: sheetW, h: sheetH,
        fill: { color: C.WHITE },
        line: { color: C.MUTED, width: 0.8 },
      });

      // Title bar
      s.addShape("rect", {
        x: sheetX, y: sheetY, w: sheetW, h: 0.24,
        fill: { color: C.ACCENT },
      });
      s.addText("Indigenous and British", {
        x: sheetX, y: sheetY, w: sheetW, h: 0.24,
        fontSize: 9, fontFace: FONT_B, color: C.WHITE, bold: true,
        align: "center", valign: "middle", margin: 0,
      });

      // Two empty circles
      const cD = 1.45;
      const cy = sheetY + 0.24 + (sheetH - 0.24) * 0.42;
      const overlap = 0.45;
      const leftCx = sheetX + sheetW / 2 - (cD / 2 - overlap / 2);
      const rightCx = sheetX + sheetW / 2 + (cD / 2 - overlap / 2);

      s.addShape("roundRect", {
        x: leftCx - cD / 2, y: cy - cD / 2, w: cD, h: cD, rectRadius: cD / 2,
        fill: { color: C.PRIMARY, transparency: 80 },
        line: { color: C.PRIMARY, width: 1.0 },
      });
      s.addShape("roundRect", {
        x: rightCx - cD / 2, y: cy - cD / 2, w: cD, h: cD, rectRadius: cD / 2,
        fill: { color: C.SECONDARY, transparency: 80 },
        line: { color: C.SECONDARY, width: 1.0 },
      });

      // Bottom band: sentence bank label
      const bankY = sheetY + sheetH - 0.45;
      s.addShape("rect", {
        x: sheetX, y: bankY, w: sheetW, h: 0.40,
        fill: { color: C.BG_LIGHT },
        line: { color: C.MUTED, width: 0.5 },
      });
      s.addText("Sentence bank: Country, songlines, ships, farming, hunting & gathering, family, stories, prisons ...", {
        x: sheetX + 0.1, y: bankY, w: sheetW - 0.2, h: 0.40,
        fontSize: 8, fontFace: FONT_B, color: C.CHARCOAL, italic: true,
        align: "left", valign: "middle", margin: 0,
      });
    }
  );

  // -- Slide 13: We Do - Model the letter --
  contentSlide(
    pres,
    "We Do  -  Model",
    C.SECONDARY,
    "A letter from an Indigenous person",
    [
      "Imagine: it is 1788. The First Fleet has just arrived",
      "An Indigenous person writes to a British settler",
      "We use what we have learnt - we do not pretend to be a real person",
      "Watch the teacher model the opening sentence",
    ],
    NOTES_WE_DO_LETTER,
    FOOTER,
    (s, lg) => {
      const rX = lg.rightX;
      const rW = lg.rightW;
      const topY = lg.panelTopPadded;
      const panelH = SAFE_BOTTOM - topY - 0.05;

      addCard(s, rX, topY, rW, panelH, { strip: C.SECONDARY, fill: C.BG_LIGHT });

      s.addText("Modelled opening", {
        x: rX + 0.2, y: topY + 0.10, w: rW - 0.4, h: 0.28,
        fontSize: 11, fontFace: FONT_H, color: C.SECONDARY, bold: true,
        align: "center", margin: 0,
      });

      // Mock letter
      const lX = rX + 0.30;
      const lW = rW - 0.60;
      const lY = topY + 0.45;
      const lH = panelH - 0.55;

      s.addShape("rect", {
        x: lX, y: lY, w: lW, h: lH,
        fill: { color: C.WHITE },
        line: { color: C.MUTED, width: 0.8 },
      });

      s.addText("Dear settler,", {
        x: lX + 0.2, y: lY + 0.20, w: lW - 0.4, h: 0.32,
        fontSize: 13, fontFace: FONT_B, color: C.CHARCOAL, bold: true,
        margin: 0,
      });

      s.addText("We have lived on this land for tens of thousands of years. This land is our Country. Country means family, ancestors, language and stories ...", {
        x: lX + 0.2, y: lY + 0.60, w: lW - 0.4, h: lH - 0.80,
        fontSize: 12, fontFace: FONT_B, color: C.CHARCOAL,
        italic: true, valign: "top", margin: 0,
      });
    }
  );

  // -- Slide 14: You Do - Write the letter --
  contentSlide(
    pres,
    "You Do  -  Write",
    C.ACCENT,
    "Write your letter",
    [
      "Use the Letter Writing Scaffold - four paragraphs",
      "P1: who you are, how long your people have lived here",
      "P2: what Country means, what you would lose",
      "P3: one impact you fear",
      "P4: what you ask the settler to do instead",
    ],
    NOTES_YOU_DO_LETTER,
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
        { n: "P1", color: C.PRIMARY, text: "I am ... I have lived ..." },
        { n: "P2", color: C.SECONDARY, text: "Our Country means ... I would lose ..." },
        { n: "P3", color: C.ALERT, text: "I fear that ..." },
        { n: "P4", color: C.SUCCESS, text: "Please ... instead of ..." },
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
      "Sticky note - finish this sentence and stick it on the class chart on your way out:\n\n'The biggest impact of British settlement on First Nations peoples was ___ because ___.'",
    ],
    NOTES_EXIT,
    FOOTER,
    { assessesSc: 2, title: "One thing before you leave" }
  );

  // -- Slide 16: Closing --
  closingSlide(
    pres,
    {
      reflectionPrompt: "Which part of today's lesson will stay with you the most, and why?",
      scItems: [
        "I can explain why the Europeans settled in Australia",
        "I can explain the impact that the Europeans had on the Indigenous population",
        "I can make comparisons between the way that the Indigenous and Europeans lived",
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
  await buildVennPdf();
  await buildLetterPdf();
}

// ───────────────────────────────────────────────────────────────
//  PDF 1: Venn diagram
// ───────────────────────────────────────────────────────────────

async function buildVennPdf() {
  const doc = createPdf({ title: "Cultures Venn Diagram" });
  let y = addPdfHeader(doc, "Cultures Venn Diagram  -  Indigenous and British", {
    subtitle: "Compare two ways of life around 1770 to 1800. One per student.",
    color: C.PRIMARY,
    lessonInfo: "Year 5/6 Inquiry  |  Shaping Australia  |  Lesson 3",
  });

  // Brief instruction
  y = addBodyText(doc,
    "Left circle: unique to First Nations life. Right circle: unique to British life. Middle: things BOTH groups did. Aim for at least 3 entries in each region.",
    y, { fontSize: 10.5, italic: true, color: "4B5563" }
  );

  y += 4;

  // Two overlapping circles
  const tableX = PAGE.MARGIN;
  const tableW = PAGE.CONTENT_W;
  const diagH = 270;
  const cR = 130;
  const overlap = 80;
  const centreX = tableX + tableW / 2;
  const centreY = y + diagH / 2;
  const leftCx = centreX - (cR - overlap / 2);
  const rightCx = centreX + (cR - overlap / 2);

  // Soft fills
  doc.save();
  doc.circle(leftCx, centreY, cR).fillOpacity(0.18).fillColor(hex(C.PRIMARY)).fill();
  doc.restore();
  doc.save();
  doc.circle(rightCx, centreY, cR).fillOpacity(0.18).fillColor(hex(C.SECONDARY)).fill();
  doc.restore();

  // Outlines
  doc.save();
  doc.lineWidth(1.4).strokeColor(hex(C.PRIMARY)).circle(leftCx, centreY, cR).stroke();
  doc.lineWidth(1.4).strokeColor(hex(C.SECONDARY)).circle(rightCx, centreY, cR).stroke();
  doc.restore();

  // Labels above
  doc.fontSize(13).font("Sans-Bold").fillColor(hex(C.PRIMARY));
  doc.text("First Nations", leftCx - cR, centreY - cR - 18, { width: cR * 2, align: "center" });
  doc.fontSize(13).font("Sans-Bold").fillColor(hex(C.SECONDARY));
  doc.text("British", rightCx - cR, centreY - cR - 18, { width: cR * 2, align: "center" });

  // BOTH label in middle
  doc.fontSize(10).font("Sans-Bold").fillColor(hex(C.SUCCESS));
  doc.text("BOTH", centreX - 30, centreY - cR + 10, { width: 60, align: "center" });

  y = centreY + cR + 24;

  // Sentence bank
  doc.save();
  doc.rect(tableX, y, tableW, 60).fill(hex("F4F1E6")).stroke();
  doc.restore();
  doc.fontSize(11).font("Sans-Bold").fillColor(hex(C.PRIMARY));
  doc.text("Sentence bank  -  use these to help fill your diagram:", tableX + 10, y + 8, { width: tableW - 20 });
  doc.fontSize(10).font("Sans").fillColor(hex("2D3142"));
  doc.text(
    "Country  |  songlines  |  Dreaming  |  hunting and gathering  |  oral storytelling  |  language groups  |  " +
    "First Fleet  |  steam ships  |  farming  |  iron tools  |  written letters  |  prisons  |  " +
    "families  |  stories  |  music and song  |  art  |  caring for children  |  trade",
    tableX + 10, y + 24, { width: tableW - 20 }
  );

  y = y + 60 + 16;

  // Extension prompt at the bottom
  y = addTipBox(doc,
    "Extension: under the diagram, write one sentence starting with 'The biggest difference between the two ways of life was ___ because ___.'",
    y,
    { color: C.ACCENT }
  );

  addPdfFooter(doc, "Year 5/6 Inquiry  -  Shaping Australia  -  Lesson 3  -  Cultures Venn Diagram");
  const outPath = path.join(RES_DIR, "Session 3 Cultures Venn Diagram.pdf");
  await writePdf(doc, outPath);
  console.log("PDF written to", outPath);
}

// ───────────────────────────────────────────────────────────────
//  PDF 2: Letter scaffold
// ───────────────────────────────────────────────────────────────

async function buildLetterPdf() {
  const doc = createPdf({ title: "Letter Writing Scaffold" });
  let y = addPdfHeader(doc, "Letter from an Indigenous person to a British settler", {
    subtitle: "Use what you have learnt in Lessons 1, 2 and 3. Write respectfully.",
    color: C.PRIMARY,
    lessonInfo: "Year 5/6 Inquiry  |  Shaping Australia  |  Lesson 3",
  });

  // addPdfHeader already provides Name/Date - no duplicate row needed.

  // Framing note
  y = addTipBox(doc,
    "It is 1788. The First Fleet has just arrived. You are an Indigenous person writing to a British settler. Use what we have learnt about Country, family, language and Dreaming. You are NOT speaking for any real person.",
    y,
    { color: C.SECONDARY }
  );

  y += 6;

  // Salutation
  doc.fontSize(11).font("Sans-Bold").fillColor(hex(C.CHARCOAL || "2D3142"));
  doc.text("Dear settler,", PAGE.MARGIN, y);
  y += 18;

  // Four scaffold blocks
  const blocks = [
    { label: "P1", color: C.PRIMARY, prompt: "Start by saying who you are and how long your people have lived here.\nTry: 'I am ___. My people have lived here for ___.'" },
    { label: "P2", color: C.SECONDARY, prompt: "Tell the settler what Country means to you, and what you would lose if you are forced off it.\nTry: 'Country means ___. If we lose Country we lose ___.'" },
    { label: "P3", color: C.ALERT, prompt: "Explain one impact you fear from British settlement.\nTry: 'I fear that ___ because ___.'" },
    { label: "P4", color: C.SUCCESS, prompt: "Ask the settler to do something different. Be respectful.\nTry: 'Please ___ instead of ___.'" },
  ];

  const blockX = PAGE.MARGIN;
  const blockW = PAGE.CONTENT_W;
  const labelW = 30;
  const blockH = 95;
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
    const linesStartY = y + 38;
    const lineCount = 3;
    const lineGap = (blockH - 38 - 6) / lineCount;
    for (let li = 0; li < lineCount; li++) {
      const ly = linesStartY + li * lineGap;
      doc.save();
      doc.moveTo(blockX + labelW + 10, ly).lineTo(blockX + blockW - 10, ly)
        .lineWidth(0.6).strokeColor("#000000").stroke();
      doc.restore();
    }

    y = y + blockH + blockGap;
  });

  // Sign-off
  doc.fontSize(11).font("Sans-Bold").fillColor(hex("2D3142"));
  doc.text("From,", PAGE.MARGIN, y + 4);
  doc.moveTo(PAGE.MARGIN + 40, y + 16).lineTo(PAGE.MARGIN + 240, y + 16)
    .lineWidth(0.7).strokeColor("#000000").stroke();

  addPdfFooter(doc, "Year 5/6 Inquiry  -  Shaping Australia  -  Lesson 3  -  Letter Scaffold");
  const outPath = path.join(RES_DIR, "Session 3 Letter Writing Scaffold.pdf");
  await writePdf(doc, outPath);
  console.log("PDF written to", outPath);
}

build().catch((err) => {
  console.error(err);
  process.exit(1);
});
