"""Deep-pass teacher notes for Lesson 4 (Plan an introduction for an information report) per teachernotes.md v2.0."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from polish_notes_lib import polish_deck

SRC = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated/4. literature_presentation Plan an introduction for an information report 4En06V Storm Boy - Information report - with notes.pptx")
OUT = SRC.with_name(SRC.stem + " v2.pptx")

NOTES = {
    1: """TEACHER NOTES:
Title slide. This is the first non-fiction lesson in the unit. Begin once Writer's Notebooks are open at Lesson 4.""",

    2: """TEACHER NOTES:
Read aloud or summarise before today's lesson. The information texts cover Ngarrindjeri culture, Kurangk and the KNY Agreement.

SENSITIVITY ADVISORY:
- What it is: today's texts cover Ngarrindjeri culture, language and the KNY Agreement.
- Framing language: 'The Ngarrindjeri people are the Traditional Custodians of the Coorong.'
- Watch for: students unsure how to talk about Aboriginal and Torres Strait Islander content.
- Protocol: model respectful language, pause if a student raises a concern, follow up privately.""",

    3: """TEACHER NOTES:
Teacher orientation only, not for students. Read once before delivering the unit.""",

    4: """TEACHER NOTES:
Teacher reference for the I Do, We Do, You Do badges and the support and extension icons. Not student-facing.""",

    5: """TEACHER NOTES:
Teacher reference for the response routines used through the deck. Not student-facing.""",

    6: """TEACHER NOTES:
Teacher reference for the sentence-element colour coding used in the modelling slides. Not student-facing.""",

    7: """SAY:
- "Read the learning intention with me."
- "Today is a non-fiction lesson. We are starting an information report on the Coorong."
- "Ask: where might you find an information report? Expected: encyclopedias, websites, textbooks, magazines."
- "If 'information report' feels new, that is okay. We will build it together."

DO:
- Choral read the LI, then track each success criterion with your finger.
- Take 2-3 student responses on where reports appear.

TEACHER NOTES:
SC2 is the structural knowledge for the rest of the unit. SC3 is today's writing focus.

WATCH FOR:
- Students confusing reports with stories. Clarify: 'An information report gives facts, not a story.'""",

    8: """SAY:
- "Boards out, novel out, Writer's Notebook ready."
- "Texta in your hand. Lid checked."

DO:
- Scan the room for missing items.
- Today's writing happens in the Writer's Notebook.

TEACHER NOTES:
Material check.""",

    9: """TEACHER NOTES:
Section divider. The next slides build background knowledge using information texts about the Coorong.""",

    10: """SAY:
- "We are going to take notes from information texts about the Coorong."
- "We use shorthand to write quickly."
- "Slash means new idea. Plus or ampersand means and. Equals means same. Arrow means leads to. Up arrow means increase. Down arrow means decrease. b/c means because. w/ means with. w/o means without."

DO:
- Glue page 3 of the booklet into the Writer's Notebook.
- Display the key.
- Demonstrate two symbols on the board.
- Ask: 'Coorong = wetland. What does that mean?' Expected: the Coorong is the same as a wetland.

TEACHER NOTES:
Note-taking shorthand supports speed and content focus. Students may already use some symbols.

WATCH FOR:
- Students who do not use shorthand. Encourage but do not insist. The goal is fast notes, not perfect symbols.""",

    11: """SAY:
- "First text: 'Where is the Coorong?'"
- "It tells us about the location."
- "The link is to the original on the SA government site. We have summarised it on the next slides."

DO:
- Briefly point to the link.
- Move to the summary content.

TEACHER NOTES:
Bridge slide. Source attribution preserved for teacher reference.""",

    12: """SAY:
- "Second text: Coorong National Park Teaching Resource."
- "It covers the Park, Ramsar Convention, cultural significance, the Ngarrindjeri Vision for Country, and middens."
- "Pages 2, 4, 5 and 8 are summarised on the next slides."

DO:
- Briefly point to the link.
- Continue to the summary.

TEACHER NOTES:
Bridge slide. Source attribution preserved for teacher reference.""",

    13: """SAY:
- "Read this passage with me."
- "Ask: what facts can we pull out? Expected: Coorong NP is a special place in SA's southeast, 200 km SE of Adelaide, long and narrow, from Goolwa Barrage to nearly Kingston."
- "On your whiteboard: write one fact in shorthand. 1 minute. Show me!"

DO:
- Read the passage aloud.
- Model one shorthand note on the board.
- Time 1 minute.
- Scan boards.
- Share two or three good examples.

TEACHER NOTES:
First note-taking practice. Focus on what it is and where it is. This content feeds body paragraph 1.

WATCH FOR:
- Students writing whole sentences. Redirect: 'Just the key idea. Use symbols.'""",

    14: """SAY:
- "New passage. Read with me."
- "Ask: what facts? Expected: home to water birds and migratory species. Holds historical and archaeological sites. Cared for by many groups working together."
- "Add a fact to your notes. 1 minute. Show me!"

DO:
- Read the passage aloud.
- Time 1 minute.
- Scan boards.

WATCH FOR:
- Students unsure about 'migratory'. Quick gloss: 'Animals that move from place to place.'""",

    15: """SAY:
- "New passage. Read with me."
- "Ask: what facts? Expected: 115 species protected, most important waterbird wetland in the Murray-Darling Basin, internationally recognised, designated Ramsar Wetland."
- "Add to your notes. 1 minute. Show me!"

DO:
- Read the passage aloud.
- Highlight 'Ramsar Wetland'. Key term for body paragraph 3.
- Time 1 minute.
- Scan boards.

WATCH FOR:
- Students recording every number. Redirect: 'Pick the most important fact, not all the numbers.'""",

    16: """SAY:
- "New passage. Read with me."
- "Ask: what facts? Expected: wetlands important for migratory shorebirds. Some birds travel from Alaska and Siberia. Ramsar Convention is an international agreement to protect wetlands."
- "Add to your notes. 1 minute. Show me!"

DO:
- Read aloud.
- Define Ramsar Convention briefly.
- Time 1 minute.

WATCH FOR:
- Students confusing Ramsar with another name. Repeat the spelling: 'R-A-M-S-A-R.'""",

    17: """SAY:
- "New passage. Read with me, slowly and respectfully."
- "Ask: what facts? Expected: Coorong is important to Ngarrindjeri people. Ngarrindjeri name = Kurangk = 'long narrow neck'. Sites include middens. A midden is a place where Aboriginal communities camped."
- "Add to your notes. 1 minute. Show me!"

DO:
- Read aloud, slowly and respectfully.
- Pronounce Ngarrindjeri carefully: Nga-rrin-jeri.
- Highlight the Kurangk meaning.
- Time 1 minute.

TEACHER NOTES:
Cultural significance content for body paragraph 2. Key terms: Ngarrindjeri, Kurangk, midden.

WATCH FOR:
- Students unsure how to write Ngarrindjeri. Display the word and let them copy if they need.
- Sensitive moments. Take time to answer respectfully.""",

    18: """SAY:
- "New passage. Read with me."
- "Ask: what facts? Expected: 2009, Ngarrindjeri negotiated the KNY Agreement with the state government. It recognised Ngarrindjeri ownership of lands and waters."
- "Add to your notes. 1 minute. Show me!"

DO:
- Read aloud.
- Brief explanation: 'KNY means Listen to what Ngarrindjeri people have to say.'
- Time 1 minute.

WATCH FOR:
- Students stalled on the long phrase. Encourage: 'Write KNY in your notes for now.'""",

    19: """SAY:
- "Ask: which state is the Coorong in? 1 NSW, 2 Vic, 3 SA, 4 WA."
- "Show me with fingers."
- "Expected: 3."

DO:
- Read each option.
- Cue finger response.
- Scan, confirm: 'South Australia. From Goolwa to Kingston.'

CFU CHECKPOINT:
Technique: Show Fingers 1-4
Script:
- Ask: 'Show me 1, 2, 3 or 4.' Expected: 3.
- Scan for: 3 across the room.
PROCEED:
- >=80% pick 3. Move on.
PIVOT:
- Most likely: students confuse SA with NSW or Vic.
- Reteach: 'The Coorong is 200 km southeast of Adelaide. Adelaide is in South Australia.'
- Re-check: 'What state is Adelaide in?'

WATCH FOR:
- Students copying neighbours. Cover their hand briefly and re-cue.""",

    20: """SAY:
- "Ask: what was the purpose of Coorong National Park? 1 fresh water, 2 protect plants and animals, 3 community activities, 4 tourists."
- "Show me with fingers."
- "Expected: 2."

DO:
- Read each option.
- Cue finger response.
- Scan, confirm: 'Mainly to protect habitats. Tourists are welcome but not the main purpose.'

CFU CHECKPOINT:
Technique: Show Fingers 1-4
Script:
- Ask: 'Show me 1, 2, 3 or 4.' Expected: 2.
- Scan for: 2 across the room.
PROCEED:
- >=80% pick 2. Move to writing.
PIVOT:
- Most likely: students pick 4 (tourists).
- Reteach: 'A National Park is mainly to protect.'
- Re-check: 'Why was Coorong NP set up - to protect or for tourists?'

WATCH FOR:
- Students who pick 1 or 3. Quick anchor: 'Look for the word that means keeping safe. That is protect.'""",

    21: """TEACHER NOTES:
Section divider. The next slides explicitly teach the structure of an information report and how to plan a title.""",

    22: """SAY:
- "An information report aims to inform. It gives facts."
- "It is not a story. It is not an opinion."
- "Read the slide with me."

DO:
- Choral read the definition.
- Brief contrast: 'A story has characters and a plot. A report has facts and topics.'

WATCH FOR:
- Students who think reports are boring. Reframe: 'Not stories, but full of interesting facts.'""",

    23: """SAY:
- "Information reports have certain features."
- "Present tense: the Coorong IS, not WAS."
- "Third person: it, they. Not I or you."
- "Subheadings to organise sections."
- "Adjectives and adverbials to build description."
- "Technical, precise vocabulary."
- "Action verbs and relating verbs."
- "May include images or diagrams."

DO:
- Read each feature once.
- Quick example: 'Wrong: the Coorong was a wetland. Right: the Coorong is a wetland.'
- Quick noun group example: 'a lagoon' becomes 'a narrow, shallow lagoon'.

TEACHER NOTES:
Reference for the rest of the unit. Today's focus is present tense and precise vocabulary.

WATCH FOR:
- Students unsure about third person. Anchor: 'He, she, it, they. Not I or you.'""",

    24: """SAY:
- "An information report has four parts."
- "TITLE tells the reader what the text is about."
- "INTRODUCTION orients the reader to the topic."
- "BODY PARAGRAPHS give details. They may have subheadings."
- "CONCLUSION reorients and summarises."

DO:
- Choral read each part.
- Quick analogy: 'Like a sandwich. Title is the label. Introduction is the top bread. Body is the filling. Conclusion is the bottom bread.'

TEACHER NOTES:
Anchor for SC2. The four parts are needed for the rest of the unit.

WATCH FOR:
- Students confusing introduction and body. Anchor: 'The introduction tells the reader what the whole report is about. The body goes deep on each part.'""",

    25: """SAY:
- "Ask: what is the purpose of an information text? 1 personal experiences, 2 factual information."
- "Show me with fingers."
- "Expected: 2."

DO:
- Read both options.
- Cue finger response.
- Scan, confirm: 'Reports give facts, not opinions.'

CFU CHECKPOINT:
Technique: Show Fingers 1-2
Script:
- Ask: 'Show me 1 or 2.' Expected: 2.
- Scan for: 2 across the room.
PROCEED:
- >=80% pick 2. Move on.
PIVOT:
- Most likely: students pick 1 because they think 'interesting' must mean opinions.
- Reteach: 'Interesting facts can still be facts. Option 1 is personal experiences.'
- Re-check: 'Is a pelican report giving facts or opinions?'

WATCH FOR:
- Students who pick 1 from doubt. Quick anchor: 'A report tells you what is true, not what someone feels.'""",

    26: """SAY:
- "We are starting our information report on the Coorong."
- "Today: title and planning the introduction."
- "Tomorrow: write the introduction."
- "First we read a model text. The model shows what a great Coorong report looks like."

DO:
- Display the slide.
- Cue the class: 'Listen for the four parts: title, introduction, body, conclusion.'

TEACHER NOTES:
Bridge slide. The model on the next pages is the target.""",

    27: """SAY:
- "Read the model introduction with me."
- "Three sentences. Each does a different job."
- "We will look at each one closely on the next slides."

DO:
- Read aloud.
- Underline each sentence on the slide if you can.
- Mention GST: 'General, Specific, Topic. We will teach this properly soon.'

WATCH FOR:
- Students trying to memorise word for word. Reassure: 'This is a model. Yours will be your own.'""",

    28: """SAY:
- "Body paragraph 1: what it is and where it is located."
- "Read aloud with me."
- "Notice the topic sentence at the start."
- "Notice the summary sentence at the end."

DO:
- Read aloud.
- Point to the topic sentence and the summary sentence.
- Quick observation: 'Tell the reader the topic, give details, sum up.'

TEACHER NOTES:
Body paragraph model. Pattern: topic sentence, then details, then summary.

WATCH FOR:
- Students stuck on technical vocabulary like 'Younghusband Peninsula'. Read place names aloud and move on.""",

    29: """SAY:
- "Body paragraph 2: cultural significance."
- "Read aloud with me, respectfully."
- "Notice the subheading."
- "It tells us how the Coorong is important to the Ngarrindjeri people."
- "It mentions the KNY Agreement (2009) and the Ngarrindjeri name Kurangk."

DO:
- Read aloud, respectfully.
- Pronounce Kurangk and Ngarrindjeri clearly.
- Highlight the subheading.

TEACHER NOTES:
Cultural body paragraph model. KNY Agreement is the key 2009 fact.

WATCH FOR:
- Sensitive moments. Pause and check in if a student raises a question.""",

    30: """SAY:
- "Body paragraph 3: flora and fauna."
- "Flora means plants. Fauna means animals."
- "Read aloud with me."
- "It tells us about the Coorong ecosystem - vulnerable species, migratory birds from Alaska and Siberia, the Bonney Upwelling."

DO:
- Read aloud.
- Define flora and fauna explicitly.
- Brief mention: 'Bonney Upwelling - water from Antarctica brings nutrients up.'

TEACHER NOTES:
Flora and fauna are key technical vocabulary for the unit.

WATCH FOR:
- Students unsure of flora and fauna. Anchor: 'Flora plants. Fauna animals.'""",

    31: """SAY:
- "Conclusion of the model text."
- "Read aloud with me."
- "Notice it does three things: reorients, summarises, and leaves a strong final thought."

DO:
- Read aloud.
- Point to each of the three functions in the text.
- Quick comparison: 'A conclusion echoes the introduction but does not just repeat it.'

TEACHER NOTES:
Conclusion model for later in the unit (Lessons 24 and 25).

WATCH FOR:
- Students who think conclusions just repeat the introduction. Highlight the difference: 'New final thought, not the same words.'""",

    32: """SAY:
- "The TITLE tells the reader what the report is about."
- "Titles can be brief, even one word: Pelicans. South Australia."
- "For broad topics, be more specific: Pelicans and their habitat. The geography of South Australia."
- "A specific title helps the reader know exactly what they will learn."

DO:
- Read each example aloud.
- Quick model: 'Pelicans is brief. Pelicans and their habitat tells us we will learn about where pelicans live.'

TEACHER NOTES:
Anchor for SC3.

WATCH FOR:
- Students who think titles must be long. Reframe: 'A title can be short. It just needs to be clear.'""",

    33: """SAY:
- "The model title is 'The Coorong'."
- "Short but clear. The reader knows the topic."

DO:
- Read the title aloud.
- Quick connection: 'Two words tell us the focus of the whole report.'

TEACHER NOTES:
Reference for student brainstorming on the next slide.""",

    34: """SAY:
- "Look at the three pictures: pelicans, the Murray-Darling River, storms at sea."
- "What could the title be for each?"
- "On your whiteboard: write a title for one or more. 2 minutes. Show me!"

DO:
- Display the pictures.
- Set 2 minutes.
- Scan boards.
- Cold call 2-3 students to share.

WATCH FOR:
- Students writing story-style titles like 'The Pelican's Adventure'. Clarify: 'That sounds like a story. A report title tells us the topic.'""",

    35: """SAY:
- "Now write your own title."
- "Topic: the Coorong."
- "On your whiteboard: brainstorm 2 to 3 possible titles. Make them interesting AND informative. 3 minutes. Show me!"

DO:
- Set 3 minutes.
- Scan boards.
- Cold call 4-5 students to share.
- Reflect: 'Which titles tell the reader what the report is about?'

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Read me one of your titles.' Expected: titles that name the topic clearly.
- Scan for: not story-style.
PROCEED:
- >=80% produce at least one clear title. Move to the Writer's Notebook.
PIVOT:
- Most likely: students write story-style titles like 'My Trip to the Coorong'.
- Reteach: 'A report is not a story. The title tells the topic. Try: The Coorong, or A Guide to the Coorong.'
- Re-check: 'Now write one title that is clear and informative.'

WATCH FOR:
- Students copying the model. Affirm and push: 'Now try a different one.'""",

    36: """SAY:
- "Open your Writer's Notebook to Lesson 4."
- "Question 2: write your title options."
- "Then copy the three subheadings."
- "Body Paragraph 1: What it is and where it is located."
- "Body Paragraph 2: Cultural significance."
- "Body Paragraph 3: Flora and Fauna."
- "Stop there for now."

DO:
- Distribute or open the Writer's Notebook.
- Set 5 minutes.
- Circulate. Check titles are topic-focused.
- Confirm subheadings are written exactly.

WATCH FOR:
- Students skipping the subheadings. Prompt: 'Copy all three.'""",

    37: """SAY:
- "The introduction follows the GST structure."
- "G is GENERAL statement. It introduces the topic and gives context."
- "S is SPECIFIC statement. It addresses the topic using key terms."
- "T is TOPIC statement. It tells the reader what the rest of the report will discuss."

DO:
- Read each definition.
- Quick analogy: 'GST is like a funnel. General is wide. Specific is narrower. Topic points to the body paragraphs.'

TEACHER NOTES:
Anchor for the writing focus. Tomorrow uses GST to write the introduction.

WATCH FOR:
- Students confusing the three parts. Reassure: 'We will look at each on the next slides.'""",

    38: """SAY:
- "GENERAL statement: 'The Coorong is a narrow, shallow lagoon that stretches across the Limestone coast in South Australia.'"
- "Read with me."
- "Topic: the Coorong. Context: it is on the Limestone Coast in SA."
- "A general statement introduces and gives context."

DO:
- Choral read.
- Underline or point to topic and context separately.
- Quick recap: 'Topic is what we are talking about. Context is where or what kind.'

WATCH FOR:
- Students who think a general statement is one word. Clarify: 'It is a full sentence.'""",

    39: """SAY:
- "SPECIFIC statement: 'It is an expansive area of wetlands and waterways that is both culturally and ecologically important.'"
- "A specific statement uses key terms."
- "Ask: what are the key terms? Turn and tell your partner. 30 seconds."
- "Expected: expansive, wetlands, waterways, culturally important, ecologically important."

DO:
- Read aloud.
- Time the partner talk.
- Cold call to share key terms.
- Highlight on the board.
- Note the typo on the slide ('important.hi') - read it as 'important.'

TEACHER NOTES:
The slide has a small typo. Mention briefly so students do not copy it.

WATCH FOR:
- Students picking simple words like 'is'. Refine: 'Key terms are the special words about the topic.'""",

    40: """SAY:
- "TOPIC statement: 'This report will discuss exactly what the Coorong is, where it is located, its cultural significance, and the flora and fauna that rely on it.'"
- "It tells the reader what each body paragraph will cover."
- "Signal phrase: 'This report will discuss'."
- "The four discussion points match the body paragraphs."

DO:
- Read aloud.
- Underline each of the four discussion points.
- Map each point to a body paragraph.

WATCH FOR:
- Students who think the topic statement is the conclusion. Clarify: 'This sets up the body. The conclusion comes at the end.'""",

    41: """SAY:
- "Class discussion. Use what you have learned."
- "What is the Coorong?"
- "Where is it located?"
- "What is its primary purpose?"
- "Who is it significant to?"
- "How can we describe it without using opinion words?"

DO:
- Read each question one at a time.
- Cold call 2-3 responses per question.
- Record key answers on the board.
- Press for factual language.

TEACHER NOTES:
Class brainstorm to generate content for the introduction. Notes feed the Writer's Notebook task.

WATCH FOR:
- Students using opinion words like amazing or beautiful. Redirect: 'Use facts. The Coorong is shallow, narrow, expansive.'""",

    42: """SAY:
- "Quick check. Write what each letter of GST stands for."
- "G stands for..."
- "S stands for..."
- "T stands for..."
- "1 minute. Show me!"
- "Expected: General, Specific, Topic."

DO:
- Time 1 minute.
- Scan boards.
- Confirm answers aloud.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Hold up your board.' Expected: General, Specific, Topic.
- Scan for: all three correct.
PROCEED:
- >=80% answer all three correctly. Move to the You Do.
PIVOT:
- Most likely: students mix Specific and Topic.
- Reteach with the funnel: 'General is wide. Specific is narrower. Topic points to the body paragraphs.'
- Re-check: 'Which one tells the reader what the body paragraphs will cover?'

WATCH FOR:
- Students writing the meanings instead of the words. Re-cue: 'The single word for each letter.'""",

    43: """SAY:
- "Open your Writer's Notebook again."
- "Question 3: write notes on each prompting question."
- "Use information from the texts and the discussion."
- "These notes become tomorrow's introduction."

DO:
- Open the Writer's Notebook.
- Set 8 to 10 minutes.
- Circulate. Check notes, not full sentences.
- Hand the scaffold to identified students.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: scaffold with questions pre-filled and key facts to circle.
- Extra Notes: sit with these students for the first question.
EXTENDING PROMPT:
- Task: after the five questions, draft a brief sentence for each part of the GST structure.

TEACHER NOTES:
This is the You Do. Notes here become the introduction in Lesson 5.

WATCH FOR:
- Students writing long sentences. Redirect: 'Notes only. Use shorthand.'
- Readiness signal: notes for at least 4 of 5 questions.""",

    44: """TEACHER NOTES:
Credits and attribution slide. Not student-facing. End the lesson on the closing reflection slide rather than this one.""",

    45: """SAY:
- "Read each I can statement with me."
- "Show me on your thumbs: up, sideways or down for each one."
- "Pick the I can statement you feel most confident about today."

DO:
- Choral read each SC, scan thumbs each time.
- Note which SC has the most thumbs sideways or down to plan tomorrow.

TEACHER NOTES:
Use this data to decide tomorrow's launch and any small-group reteach.

WATCH FOR:
- Students avoiding the rating. Prompt them to commit to one.
- Patterns where SC2 or SC3 is mostly thumbs down. Flag for reteach.""",
}


def main():
    stats = polish_deck(SRC, OUT, overrides=NOTES)
    print(f"L04 written: {stats}")


if __name__ == "__main__":
    main()
