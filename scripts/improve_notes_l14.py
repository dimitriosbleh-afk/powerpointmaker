"""Deep-pass teacher notes for Lesson 14 (Plan a body paragraph - cultural significance) per teachernotes.md v2.0."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from polish_notes_lib import polish_deck

SRC = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated/14. literature_presentation Plan a body paragraph for an information report 4En06V Storm Boy - Information report - with notes.pptx")
OUT = SRC.with_name(SRC.stem + " v2.pptx")

NOTES = {
    1: """TEACHER NOTES:
Title slide. Today's focus: plan body paragraph 2 of the Coorong information report using a Single Paragraph Outline. The paragraph theme is the cultural significance of the Coorong.""",

    2: """TEACHER NOTES:
Read aloud or summarise before today's lesson.

SENSITIVITY ADVISORY:
- What it is: Aboriginal and Torres Strait Islander histories and cultures, including reference to deceased persons, traditional ownership and culturally significant sites.
- Framing language: 'We are learning from Ngarrindjeri Country today. We approach this with respect.'
- Watch for: students who may have personal connections to the content.
- Protocol: privately check in if a student looks uncertain. Remind students they can step out quietly if needed.""",

    7: """SAY:
- "Read the learning intention with me."
- "Today we plan a body paragraph - not write it, just plan it."
- "Read the success criteria. Three I can statements."
- "We will check them at the end."

DO:
- Choral read the LI, then track each SC with your finger.
- Cold call one student to repeat SC3 in their own words.

TEACHER NOTES:
SC1: name the parts of an information report body paragraph (foundation). SC2: take notes from the article. SC3: write a complete SPO with TS, 3 supporting details and CS (core / exit ticket).

WATCH FOR:
- Students who say 'we did SPO in Lesson 12'. Great. Today's twist is using a non-fiction article as the source.""",

    10: """SOURCES:
Coorong National Park Teacher Resource Pack, page 8.

SAY:
- "Today we are reading a non-fiction text about the Coorong National Park."
- "Specifically focusing on cultural significance and the Ngarrindjeri Vision for Country."
- "Listen out for: Kurangk, Ngarrindjeri, KNY, middens."

DO:
- Display the slide.
- Choral read the four key terms.

TEACHER NOTES:
Introduction to the non-fiction source. Pre-teach the four key terms before reading the longer summary slides.

WATCH FOR:
- Students mispronouncing 'Kurangk' or 'Ngarrindjeri'. Model the pronunciation slowly twice.""",

    11: """SOURCES:
Coorong National Park Teacher Resource Pack.

SAY:
- "Read with me. Cultural significance of the Coorong."
- "The Coorong holds immense cultural significance for the Ngarrindjeri people."
- "The area is known as Kurangk, meaning long narrow neck, in Ngarrindjeri language."
- "The Coorong has culturally significant sites such as middens."
- "A midden is a site where Indigenous communities camped and left the remains of their meals."

DO:
- Choral read the slide in chunks.
- Pause after 'middens' to check understanding: 'What is a midden?'
- Cold call two students for the meaning.

TEACHER NOTES:
Reading slide for cultural significance. Long text. Chunk it for younger readers.

WATCH FOR:
- Students unsure what a midden is. Press: 'It is a campsite that has remains of meals - shells, bones, sometimes tools.'""",

    12: """SOURCES:
Coorong National Park Teacher Resource Pack.

SAY:
- "Read with me. In 2009, the Ngarrindjeri Nation negotiated the Kungun Ngarrindjeri Yunnan Agreement, or KNY, with the state government."
- "Kungun Ngarrindjeri Yunnan means 'Listen to what Ngarrindjeri people have to say'."
- "This agreement recognised Ngarrindjeri ownership of their lands and waters."
- "It set up a process to support their rights and responsibilities for Country."

DO:
- Choral read the slide in chunks.
- Pause after each sentence for understanding.
- Cold call: 'What did the KNY Agreement do?'

TEACHER NOTES:
Reading slide for the KNY Agreement. Stress the meaning of KNY: 'Listen to what Ngarrindjeri people have to say.'

WATCH FOR:
- Students who think KNY is a person's name. Press: 'It is a treaty - an agreement between two groups.'""",

    13: """SAY:
- "Hinge question. What is the Ngarrindjeri people's vision of the Coorong? Select all that apply."
- "Options: 1) respect our country and creation, 2) clean sparkling waters, healthy land, people and living things, 3) protecting our ancestors, 4) caring, sharing, respecting and knowing the lands, water and all living things."
- "Show fingers - one option at a time, thumbs up if it applies."

DO:
- Run options one at a time.
- Wait time on each, then on cue all show.

CFU CHECKPOINT:
Technique: Thumbs Up or Down per option
Script:
- Ask: 'Thumbs up if option 1 is part of the Ngarrindjeri vision, thumbs down if not.' Repeat for each option.
- Scan for: students choosing options that match the article.
PROCEED:
- >=80% select the on-text options. Move to the writing section.
PIVOT:
- Most likely: students show thumbs down on all. Misconception: not making a connection between Caring for Country and the vision.
- Reteach: 'Refer back to slides 11 and 12. What does the vision say about country, land, water and living things?'
- Re-check: 'Show me again.'

TEACHER NOTES:
Multiple-select CFU on the Ngarrindjeri vision. Typically 1, 2 and 4 are direct vision statements; 3 is implied.""",

    15: """SAY:
- "The purpose of an information report is to inform the reader."
- "Information reports classify or describe factual information about a topic."
- "They are not stories. They are not opinions. They are facts."

DO:
- Read the definition aloud.
- Cold call two students: 'Where might you read an information report?'

TEACHER NOTES:
I Do for the purpose of an information report. Establishes the genre.

WATCH FOR:
- Students who confuse information report with persuasive text. Press: 'Facts, not opinions.'""",

    16: """SAY:
- "Our topic is The Coorong."
- "The paragraph focus is cultural significance."
- "That means: how is the Coorong important to people, especially the Ngarrindjeri?"

DO:
- Read the slide aloud.
- Refer to the mentor text link if you have it open.

TEACHER NOTES:
I Do for today's paragraph topic. Cultural significance is body paragraph 2 of the full Coorong report.""",

    17: """SAY:
- "Together. What is the structure of an information report?"
- "Title, Introduction, Body paragraphs (1, 2, 3), Conclusion."
- "We are planning body paragraph 2 today using an SPO."

DO:
- Read the structure aloud.
- Hold up two fingers - we are working on body paragraph 2.
- Cold call: 'What is in body paragraph 1? What is in body paragraph 3?'

TEACHER NOTES:
We Do for the structure of an information report. The slide says 'factual recount' but the genre is information report - same structure applies.

WATCH FOR:
- Students unsure where the body paragraphs sit. Cue: 'Between introduction and conclusion.'""",

    18: """SAY:
- "Together. Features of an information report."
- "Read each feature with me: present tense, third person pronouns, subheadings, simple and expanded noun groups, adjectives and adverbials, technical vocabulary, action and relating verbs, images or labelled diagrams."
- "Today we focus on the features in red: present tense, technical vocabulary, adjectives and adverbials."

DO:
- Choral read each feature.
- Highlight the three focus features.

TEACHER NOTES:
We Do for the features of an information report. Today's focus is the three highlighted features.""",

    19: """SAY:
- "A Single Paragraph Outline puts your writing in an order that makes it easy for readers to understand."
- "It also helps you create a clear and cohesive paragraph."
- "It is like a paragraph plan."

DO:
- Choral read the slide.
- Hold up the SPO graphic if students used one in Lesson 12.

TEACHER NOTES:
I Do for the SPO concept. Should be familiar from Lesson 12.

WATCH FOR:
- Students confident with SPO. Good. Press: 'Today the source is a non-fiction article, not a novel.'""",

    20: """SAY:
- "Together. The parts of a paragraph."
- "Topic sentence: introduces the main idea."
- "Supporting details: provide details that support the topic sentence."
- "Concluding sentence: summarises or concludes the paragraph."

DO:
- Choral read each part.
- Hold up three fingers - three parts.

TEACHER NOTES:
We Do for SPO parts. Same as Lesson 12.""",

    21: """SAY:
- "Together. Identify the parts of this paragraph."
- "I will read it aloud first."
- "Listen for: where is the topic sentence? where are the supporting details? where is the concluding sentence?"
- "Read along with me."

DO:
- Read the paragraph aloud.
- Highlight the topic sentence (first), supporting details (middle), and concluding sentence (last).
- Click to reveal labels.
- Cold call to justify each label.

TEACHER NOTES:
We Do for identifying SPO parts in a real model paragraph. The model is the completed Coorong cultural significance paragraph that students will plan today.

WATCH FOR:
- Students who put the wrong label. Reteach: 'TS at the start, CS at the end, supporting details in the middle.'""",

    22: """SAY:
- "Hinge question. Which features should we include when planning your SPO?"
- "Options: write in present tense, use subheadings to categorise, include opinions, include technical language."
- "Thumbs up if include, thumbs down if not."

DO:
- Run each option one at a time.
- Wait time, then on cue all show.
- Reveal: present tense (yes), subheadings (yes), opinions (no), technical language (yes).

CFU CHECKPOINT:
Technique: Thumbs Up or Down per option
Script:
- Ask: 'Thumbs up if include, thumbs down if not.' Run each option in turn.
- Scan for: thumbs down on opinions, thumbs up on the other three.
PROCEED:
- >=80% correctly classify all four. Move to the topic sentence section.
PIVOT:
- Most likely: students who include opinions. Misconception: confusing information reports with persuasive text.
- Reteach: 'Information reports are facts only. Opinions go in persuasive writing.'
- Re-check: 'Thumbs again on opinions.'

TEACHER NOTES:
Key hinge for genre features.""",

    23: """SAY:
- "Topic sentence purpose. Read with me."
- "Grabs the reader's attention."
- "Introduces the main idea of the paragraph."
- "Tells the reader what to expect."
- "Helps to keep your writing clear and organised."

DO:
- Click to reveal each bullet.
- Choral read each one.

TEACHER NOTES:
I Do for topic sentence purpose. Connects to writing the cultural significance topic sentence.

WATCH FOR:
- Students who think a topic sentence is a question. Press: 'It is a statement that introduces the main idea.'""",

    24: """SAY:
- "Together. Look at this topic sentence."
- "'The Coorong, named Kurangk by the Ngarrindjeri people, is of enormous cultural significance.'"
- "It tells us the Ngarrindjeri word for The Coorong and that it is culturally significant to them."
- "Brainstorm with your partner. How could we paraphrase this in another way?"

DO:
- Partner talk for 60 seconds.
- Cold call two students for paraphrases.

TEACHER NOTES:
We Do for paraphrasing the topic sentence. Sets up the next slide where alternative TS options are revealed.""",

    25: """SAY:
- "Here are two alternative topic sentences."
- "Option A: 'The Coorong is culturally significant to the Ngarrindjeri people who refer to the Coorong as Kurangk meaning long, narrow neck.'"
- "Option B: 'The Ngarrindjeri Indigenous community have a timeless connection to the Coorong, known as the Kurangk in their language.'"
- "Notice each one introduces the main idea but uses different word order."

DO:
- Choral read each option.
- Cold call: 'Which one would you use? Why?'

TEACHER NOTES:
We Do for alternative topic sentences. Provide these to students who need a generated topic sentence for their booklet task.

WATCH FOR:
- Students who copy directly. Press: 'We want your own version. Mix words from these and your own.'""",

    26: """SAY:
- "Supporting details. Read with me."
- "Includes reasons."
- "Gives examples."
- "Provides facts."
- "Helps to keep your writing clear and organised."

DO:
- Click to reveal each bullet.
- Choral read each one.

TEACHER NOTES:
I Do for supporting details purpose.

WATCH FOR:
- Students who include opinions. Press: 'Only facts in supporting details.'""",

    27: """SAY:
- "Watch me find supporting details from the article."
- "Topic sentence: 'The Coorong, named Kurangk by the Ngarrindjeri people, is of enormous cultural significance.'"
- "I read the article and look for facts that support cultural significance."
- "I notice: 'In 2009, the Ngarrindjeri Nation negotiated the KNY Agreement with the state government.'"
- "That is fact 1. I write it in note form: '2009 - KNY agreement = signed -> state government & Ngarrindjeri people.'"

DO:
- Display the SPO frame.
- Walk through reading the article and writing fact 1 in note form.
- Choral read the note.

TEACHER NOTES:
I Do for finding the first supporting detail. Note form, not full sentences.

WATCH FOR:
- Students who write in full sentences. Press: 'SPO is notes only.'""",

    28: """SAY:
- "Together this time."
- "Read this section: 'The agreement recognised Ngarrindjeri ownership of their lands and waters and set up a process to support their rights and responsibilities for Country.'"
- "Ask: what is the main fact here? Expected: traditional ownership recognised + processes for rights."
- "Notes: 'Agreement - traditional owners of land -> established processes for Aboriginal rights of the Coorong.'"

DO:
- Read the article section aloud.
- Whiteboards: students write their own note version first.
- Reveal the model.

TEACHER NOTES:
We Do for the second supporting detail. Stress short note form.

WATCH FOR:
- Students who copy chunks of the original sentence. Cue: 'Shorten and use shorthand.'""",

    29: """SAY:
- "Last supporting detail."
- "Read this section: 'The Ngarrindjeri's cultural heritage involvement is integral to the park's current management. The KNY Agreement has helped the South Australian Government support the Ngarrindjeri in Caring for Country.'"
- "Notes: 'Parks managed with Ngarrindjeri people -> preserve culture & heritage.'"
- "Add the last supporting detail to your SPO."

DO:
- Whiteboards: students write the third detail.
- Reveal the model.
- Choral read the SPO so far (TS + 3 supporting details).

TEACHER NOTES:
We Do for the third supporting detail.""",

    30: """SAY:
- "Here is the SPO with all three supporting details."
- "TS: 'The Coorong, named Kurangk by the Ngarrindjeri people, is of enormous cultural significance.'"
- "1) 2009 - KNY agreement = signed -> state government & Ngarrindjeri people."
- "2) Agreement - traditional owners of the land -> established processes for Aboriginal rights of the Coorong."
- "3) Parks managed with Ngarrindjeri people -> preserve culture & heritage."
- "Choral read the whole SPO so far."

DO:
- Display the SPO.
- Choral read each line.
- Pause for any student questions.

TEACHER NOTES:
We Do summary of the SPO with all three supporting details.""",

    31: """SAY:
- "Hinge question. What is the purpose of the supporting details?"
- "Options: 1) let the reader know what comes next, 2) tell the reader your conclusions, 3) provide additional information that expands on the TS."
- "Show fingers 1, 2 or 3."

DO:
- Wait time. On cue all show together.
- Cold call to justify.

CFU CHECKPOINT:
Technique: Show Fingers 1-3
Script:
- Ask: 'Show me what supporting details do.' Expected: 3.
- Scan for: 3 on most boards.
PROCEED:
- >=80% show 3. Move on to the next We Do.
PIVOT:
- Most likely: students show 1 or 2. Misconception: thinking supporting details are a preview or a conclusion.
- Reteach: 'Supporting details add facts that support the topic sentence's main idea.'
- Re-check: 'Show me again.'

TEACHER NOTES:
Key hinge. Answer: 3.""",

    32: """SAY:
- "We Do. On your whiteboard, write three supporting details that tell us about the KNY agreement, the traditional owners, and cultural significance."
- "Use these prompts to help: which Indigenous group is the Coorong significant to? what does Kurangk mean? what was the KNY agreement?"
- "You can use your notes from yesterday's lesson."

DO:
- Whiteboards. Set timer for 4 minutes.
- Class discussion first if students need it.
- Print the supplementary article and distribute if needed.
- Circulate, check first detail of 3 to 5 students.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: provide the SPO from slide 30 partially completed. Student fills in 1 detail only.
EXTENDING PROMPT:
- Task: find a fourth supporting detail not on the model SPO.

TEACHER NOTES:
We Do for student-written supporting details. Builds independence before the You Do.

WATCH FOR:
- Students writing in full sentences. Cue: 'Notes only.'
- Students rewriting the topic sentence. Cue: 'Details support the TS, they don't repeat it.'""",

    33: """SAY:
- "Concluding sentence. Read with me."
- "A concluding sentence summarises the main ideas of the paragraph and links with the topic sentence."
- "Watch this example: 'By prioritising the conservation and custodianship of the Coorong, both its cultural and ecological importance can be preserved.'"
- "Notice it echoes 'cultural significance' from the topic sentence without repeating it."

DO:
- Display the SPO with all parts including CS.
- Choral read each line.
- Cold call: 'How does the CS link back to the TS?'

TEACHER NOTES:
I Do for the concluding sentence. Connects the SPO format - TS to CS - and shows the example.

WATCH FOR:
- Students who repeat the TS. Press: 'Echo the idea, change the words.'""",

    34: """SAY:
- "Together. Discuss the positive impacts of Ngarrindjeri people as custodians of the land."
- "Read the example CS: 'By prioritising the conservation and custodianship of the Coorong, both its cultural and ecological importance can be preserved.'"
- "Turn and tell your partner: what does custodianship mean? What positive impacts come from it?"

DO:
- Partner talk for 60 seconds.
- Cold call two students to share.
- Listen for ideas about caring for country, preserving culture, protecting nature.

TEACHER NOTES:
We Do for understanding the CS through partner talk. Builds vocabulary (custodianship) before students write their own.

WATCH FOR:
- Students unsure what 'custodianship' means. Quick definition: 'Looking after something on behalf of others - like a caretaker.'""",

    35: """SAY:
- "Brainstorm key words or phrases for your concluding sentence."
- "Word bank: custodians, culture, traditional owners, diversity, conservation, important, country, legacy."
- "Pick 2 to 3 words to use in your CS."

DO:
- Whiteboards. Students write their chosen words.
- Cold call to share choices.

TEACHER NOTES:
We Do for word selection before drafting the CS. Reduces the cognitive load before independent writing.""",

    36: """SAY:
- "Hinge question. What is the purpose of a concluding sentence?"
- "Options: 1) summarises the main ideas, 2) captures the reader's emotions, 3) introduces new events, 4) provides a link to the topic sentence."
- "Thumbs up if it applies, thumbs down if not."

DO:
- Run each option one at a time.
- Wait time, then on cue all show.
- Reveal: 1 yes, 2 no, 3 no, 4 yes.

CFU CHECKPOINT:
Technique: Thumbs Up or Down per option
Script:
- Ask: 'Thumbs up or down for each option.' Run options one at a time.
- Scan for: thumbs up on 1 and 4, thumbs down on 2 and 3.
PROCEED:
- >=80% correctly classify all four. Move to the booklet task.
PIVOT:
- Most likely: students with thumbs up on 2 (emotions) or 3 (new events). Misconception: confusing CS with persuasive techniques or a new paragraph.
- Reteach: 'CS does two things only - summarise the ideas, link to the TS. No emotions, no new events.'
- Re-check: 'Show me again on options 2 and 3.'

TEACHER NOTES:
Key hinge. Answers: 1 (yes), 2 (no), 3 (no), 4 (yes).""",

    37: """SAY:
- "Open your booklet to Lesson 14: Plan a body paragraph for an information report."
- "Plan body paragraph 2 of the Coorong report - cultural significance."
- "First: write a topic sentence. Next: write 3 supporting details in note form. Then: write a concluding sentence."

DO:
- Direct to the booklet.
- Set timer for 12 to 15 minutes.
- Circulate. Check first response of 3 to 5 students.
- Use the alternative TS from slide 25 as scaffold for students who need it.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: provide the topic sentence from slide 25 and the supporting details from slide 30. Students only write the CS.
EXTENDING PROMPT:
- Task: plan body paragraph 1 (location) or body paragraph 3 (flora and fauna) as well.

TEACHER NOTES:
You Do task. Students plan body paragraph 2 using the SPO format.

WATCH FOR:
- Students whose details are not on-topic. Press: 'Stay focused on cultural significance.'
- Students who write a CS that introduces new ideas. Press: 'Just sum up.'
- Fast finishers. Prompt: 'Try the extending task.'""",

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
    print(f"L14 written: {stats}")


if __name__ == "__main__":
    main()
