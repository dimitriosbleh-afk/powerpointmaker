"""Deep-pass teacher notes for Lesson 15 (Write a body paragraph - cultural significance) per teachernotes.md v2.0."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from polish_notes_lib import polish_deck

SRC = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated/15. literature_presentation Write a body paragraph for an information report 4En06V Storm Boy - Information report - with notes.pptx")
OUT = SRC.with_name(SRC.stem + " v2.pptx")

NOTES = {
    1: """TEACHER NOTES:
Title slide. Today's focus: turn the SPO from Lesson 14 into a full body paragraph using information report language features.""",

    2: """TEACHER NOTES:
Cultural sensitivity slide. Read silently before teaching.

SENSITIVITY ADVISORY:
- What it is: Aboriginal and Torres Strait Islander histories and cultures.
- Framing language: 'We are learning from Ngarrindjeri Country today. We approach this with respect.'
- Watch for: students who may have personal connections to the content.
- Protocol: privately check in if a student looks uncertain.""",

    7: """SAY:
- "Read the learning intention with me."
- "Today is writing day. From plan to paragraph."
- "Read the success criteria. Three I can statements."

DO:
- Choral read the LI, then each SC.
- Brief check: 'Hold up your SPO from Lesson 14.' Scan.

TEACHER NOTES:
SC1: write a body paragraph with TS, supporting details and CS (foundation). SC2: use information report language features (core / exit ticket). SC3: proofread and edit.

WATCH FOR:
- Students without their SPO from Lesson 14. Provide the model from slide 30 of Lesson 14 as a backup.""",

    10: """SOURCES:
https://coorongcountry.com.au/indigenous-ngarrindjer-coorong-culture/

SAY:
- "Today we revisit information about Ngarrindjeri culture."
- "We will read about Murrundi (the river) and Kurrangk (Coorong, narrow neck)."
- "Listen out for: Ngarrindjeri, Murrundi, Kurrangk."

DO:
- Display the slide.
- Choral read the three key terms.

TEACHER NOTES:
Introduction to the Ngarrindjeri Culture source. Pre-teach pronunciation of the three key terms.

WATCH FOR:
- Students mispronouncing the Ngarrindjeri words. Model slowly twice.""",

    11: """SOURCES:
https://coorongcountry.com.au/indigenous-ngarrindjer-coorong-culture/

SAY:
- "Read with me."
- "The lower Murray River, Lower Lakes and Coorong region was densely populated."
- "Ngarrindjeri have a timeless connection with the Murray River, lakes and Coorong."
- "They lived in communities, calling this environment their 'supermarket' - everything they needed was here."
- "The river supported waterfowl, mussels, cockles, fish, turtles, kangaroos, possums, native fruits and vegetables."
- "They were known as big, strong and healthy due to their varied diet and knowledge of the land."

DO:
- Choral read the slide in chunks.
- Pause for understanding after 'supermarket'.
- Cold call: 'Why is this environment called a supermarket?'

TEACHER NOTES:
Reading slide. Establishes the deep, sustained Ngarrindjeri connection to Country.""",

    12: """SOURCES:
https://coorongcountry.com.au/indigenous-ngarrindjer-coorong-culture/

SAY:
- "Read with me."
- "Kurrangk - Coorong - means narrow neck."
- "It is a major focus of Ngarrindjeri culture and economy."
- "People still know their special places, visit and care for the country, and tell the stories that keep it alive."
- "After European settlement, Ngarrindjeri continued hunting, gathering and seasonal work."
- "Elders taught children about stars, land, food, plants, medicinal plants, and family history."

DO:
- Choral read the slide in chunks.
- Pause for understanding.
- Cold call: 'How is Ngarrindjeri culture continued today?'

TEACHER NOTES:
Reading slide. Emphasises living, ongoing culture. Stress that this knowledge is continued today by present generations.""",

    13: """SOURCES:
https://coorongcountry.com.au/indigenous-ngarrindjer-coorong-culture/

SAY:
- "Read with me. Visit the Meningie Pelican Path."
- "Reminders of the original inhabitants are found in coastal parks."
- "Aboriginal people lived in the area for thousands of years."
- "Middens - heaps of discarded shellfish remains - are found in these parks."
- "All Aboriginal sites are important culturally and scientifically. Relics must not be disturbed or removed."

DO:
- Choral read the slide in chunks.
- Pause for understanding.
- Cold call: 'What is a midden? Why are they important?'

TEACHER NOTES:
Reading slide. Stress the message: 'Do not disturb or remove relics.' This is both legal and ethical.

WATCH FOR:
- Students who say 'why don't they take some home'. Press: 'Relics belong to Country and to the Ngarrindjeri.'""",

    14: """SAY:
- "Hinge question. Why were the Ngarrindjeri people referred to as big, strong and healthy?"
- "Options: 1) due to healthy populations of waterfowl, mussels, cockles..., 2) because of their varied diet and knowledge of the land, 3) because of their timeless connection with the rivers and lakes, 4) because they traded with groups upstream."
- "Show fingers 1 to 4."

DO:
- Wait time. On cue all show together.
- Cold call to justify.

CFU CHECKPOINT:
Technique: Show Fingers 1-4
Script:
- Ask: 'Show me why Ngarrindjeri people were big, strong and healthy.' Expected: 2.
- Scan for: 2 on most boards.
PROCEED:
- >=80% show 2. Move on.
PIVOT:
- Most likely: students choose a related fact rather than the cause.
- Reteach: 'The text says they were big, strong and healthy DUE TO their varied diet and knowledge of the land. That is the cause.'
- Re-check: 'Show me again.'

TEACHER NOTES:
Text-dependent comprehension CFU. Answer: 2.""",

    15: """SAY:
- "Hinge question. What provided an abundance of food for the Indigenous people?"
- "Options: 1) the sea and coastal lakes, 2) the local land which they hunted on, 3) the people they traded with."
- "Show fingers 1 to 3."

DO:
- Wait time. On cue all show together.
- Cold call to justify.

CFU CHECKPOINT:
Technique: Show Fingers 1-3
Script:
- Ask: 'Show me what provided abundance of food.' Expected: 1.
- Scan for: 1 on most boards.
PROCEED:
- >=80% show 1. Move on.
PIVOT:
- Most likely: students confuse 'abundance' with 'all food sources'.
- Reteach: 'The slide on Pelican Path tells us the sea and coastal lakes provided abundance of food.'
- Re-check: 'Show me again.'

TEACHER NOTES:
Text-dependent CFU. Answer: 1.""",

    16: """SOURCES:
Coorong National Park Teacher Resource Pack, page 8.

SAY:
- "Now we revisit the Coorong National Park source."
- "Specifically focusing on cultural significance and the KNY Agreement."

DO:
- Display the slide.

TEACHER NOTES:
Introduction to the second source. Same source as Lesson 14 - this is a recap before drafting.""",

    17: """SOURCES:
Where is the Coorong, Coorong National Park Teacher Resource Pack.

SAY:
- "Read with me. Cultural significance of the Coorong (recap from Lesson 14)."
- "The Coorong holds immense cultural significance for the Ngarrindjeri people."
- "The area is known as Kurangk, meaning long narrow neck, in Ngarrindjeri language."
- "The Coorong has culturally significant sites such as middens."

DO:
- Choral read the slide.
- Brief check: 'Who remembers what a midden is?'

TEACHER NOTES:
Recap reading slide.""",

    18: """SOURCES:
Where is the Coorong, Coorong National Park Teacher Resource Pack.

SAY:
- "Read with me. KNY Agreement (recap from Lesson 14)."
- "In 2009, the Ngarrindjeri Nation negotiated the Kungun Ngarrindjeri Yunnan Agreement with the state government."
- "The agreement recognised Ngarrindjeri ownership of their lands and waters."

DO:
- Choral read.
- Brief check: 'What does KNY stand for?'

TEACHER NOTES:
Recap reading slide.""",

    19: """SAY:
- "Hinge question. What does Kurangk mean?"
- "Options: 1) long lagoon, 2) Ngarrindjeri cultural heritage, 3) middens, 4) narrow long neck."
- "Show fingers 1 to 4."

DO:
- Wait time. On cue all show together.
- Cold call to justify.

CFU CHECKPOINT:
Technique: Show Fingers 1-4
Script:
- Ask: 'Show me what Kurangk means.' Expected: 4.
- Scan for: 4 on most boards.
PROCEED:
- >=80% show 4. Move to writing.
PIVOT:
- Most likely: students confuse the word with related Coorong terms.
- Reteach: 'Kurangk means long narrow neck. That is the literal translation.'
- Re-check: 'Show me again.'

TEACHER NOTES:
Text-dependent CFU. Answer: 4.""",

    21: """SAY:
- "The purpose of an information report is to inform the reader."
- "Information reports classify or describe factual information about a topic."

DO:
- Read the definition aloud.
- Cold call: 'Where might you read an information report?'

TEACHER NOTES:
I Do recap of the genre purpose.""",

    22: """SAY:
- "Together. Order the structure of a factual recount."
- "Title (1), Introduction (2), Body paragraphs (3), Conclusion (4)."
- "Show fingers 1, 2, 3 or 4 for the correct order."

DO:
- Pause for student response on order.
- Reveal.

TEACHER NOTES:
We Do for ordering the structure.""",

    23: """SAY:
- "Hinge question. A factual recount..."
- "Options: 1) informs the reader, 2) aims to persuade, 3) is written in present tense, 4) is based on factual information."
- "Thumbs up if true, thumbs down if not."

DO:
- Run each option one at a time.
- Wait time on each, then on cue all show.
- Reveal: 1 yes, 2 no, 3 yes, 4 yes.

CFU CHECKPOINT:
Technique: Thumbs Up or Down per option
Script:
- Ask: 'Thumbs up if true, thumbs down if not.' Run each option in turn.
- Scan for: thumbs down on 2 (persuade), thumbs up on 1, 3 and 4.
PROCEED:
- >=80% correctly classify all four. Move to the writing section.
PIVOT:
- Most likely: students put thumbs up on persuade. Misconception: confusing information reports with persuasive writing.
- Reteach: 'Information reports inform. Persuasive writing persuades. Different jobs.'
- Re-check: 'Show me again on persuade.'

TEACHER NOTES:
Key hinge for genre features.""",

    24: """SAY:
- "Our topic is The Coorong."
- "Today's paragraph focus is cultural significance - same as the SPO you planned in Lesson 14."

DO:
- Read the slide aloud.
- Hold up the mentor text if printed.

TEACHER NOTES:
I Do for today's paragraph topic.""",

    25: """SAY:
- "Together. Features of an information report (revisited)."
- "Read each feature with me."
- "Today we focus on the features in red: present tense, technical vocabulary, adjectives and adverbials."

DO:
- Choral read each feature.
- Highlight the three focus features.

TEACHER NOTES:
We Do for features of an information report. Quick recap.""",

    26: """SAY:
- "Together. Identify the technical or precise vocabulary in this extract."
- "Read with me. Look for words that are specific to this topic - they would not appear in everyday talk."
- "I am looking for: Kungun Ngarrindjeri Yunnan, KNY agreement, traditional ownership, Aboriginal rights, custodianship, conservation."

DO:
- Read the extract aloud.
- Highlight technical vocabulary as you go.
- Cold call students for additional terms.

TEACHER NOTES:
We Do for spotting technical vocabulary in body paragraph 2 of the model.

WATCH FOR:
- Students highlighting common words. Press: 'Technical vocabulary is topic-specific. Would you say it at a barbecue?'""",

    27: """SAY:
- "Together. Identify the technical or precise vocabulary in this extract about flora and fauna."
- "Read with me. Look for topic-specific words."
- "I am looking for: ecosystem, vulnerable, endangered, migratory, Bonney Upwelling, marine, nutrients, Antarctica."

DO:
- Read the extract aloud.
- Highlight technical vocabulary as you go.
- Cold call students for additional terms.

TEACHER NOTES:
We Do for spotting technical vocabulary in body paragraph 3.""",

    28: """SAY:
- "Together. Identify the present tense in these sentences."
- "Sentence A: 'This agreement recognises traditional ownership of the lands.' The verb 'recognises' is present tense."
- "Sentence B: 'By prioritising the conservation and custodianship of the Coorong, both its cultural and ecological importance can be preserved.' The verb 'prioritising' is present tense."
- "Notice the -s and -ing endings."

DO:
- Highlight 'recognises' and 'prioritising' on the slide.
- Cold call: 'How can you tell it is present tense?'

TEACHER NOTES:
We Do for identifying present tense verbs.

WATCH FOR:
- Students who say -ed words are present tense. Press: '-ed is past tense. -s and -ing are present.'""",

    29: """SAY:
- "Today we are writing body paragraph 2."
- "Our topic is The Coorong."
- "Our paragraph focus is cultural significance."
- "Use your SPO from Lesson 14 to draft your paragraph."

DO:
- Read the slide aloud.
- Brief check: 'Hold up your SPO.' Scan.
- Provide a printed model SPO for any student missing one.

TEACHER NOTES:
I Do bridging slide. Sets up the writing task.""",

    30: """SAY:
- "Together. Identify the SPO and the paragraph."
- "Look at the two boxes."
- "Box 1 is in note form with arrows and shorthand - that is the SPO."
- "Box 2 is in full sentences - that is the paragraph."
- "Discuss with your partner: which box is the SPO?"

DO:
- Partner talk for 30 seconds.
- Reveal: Box 1 is the SPO, Box 2 is the paragraph.
- Compare them - notice how each SPO note becomes one full sentence.

TEACHER NOTES:
We Do for SPO vs paragraph distinction. Sets up the conversion process.

WATCH FOR:
- Students who think the SPO and paragraph are the same. Press: 'SPO is the plan, paragraph is the writing.'""",

    31: """SAY:
- "Together. SPO vs body paragraph."
- "SPO: brief ideas, notes and symbols, facts not details."
- "Body paragraph: expanded ideas, full sentences, add details."
- "The SPO is the skeleton. The paragraph adds the muscles and skin."

DO:
- Choral read each comparison.
- Hold up two fingers - we have two formats.

TEACHER NOTES:
We Do for the SPO-to-paragraph transformation. The skeleton metaphor is helpful.

WATCH FOR:
- Students unsure about how to expand. Cue: 'Each SPO note becomes one full sentence.'""",

    32: """SAY:
- "Hinge question. How do you change your SPO into a body paragraph?"
- "Option A: expand ideas / record in full sentences / add detail."
- "Option B: expand ideas / record ideas in note form / list events without detail."
- "Show fingers A or B."

DO:
- Wait time. On cue all show together.
- Cold call to justify.

CFU CHECKPOINT:
Technique: Show Fingers A or B
Script:
- Ask: 'Show me how you change your SPO into a paragraph.' Expected: A.
- Scan for: A on most boards.
PROCEED:
- >=80% show A. Move on.
PIVOT:
- Most likely: students show B. Misconception: thinking the paragraph stays in note form.
- Reteach: 'The paragraph is in full sentences with details. Notes were the plan.'
- Re-check: 'Show me again.'

TEACHER NOTES:
Key hinge. Answer: A.""",

    33: """SAY:
- "Writing the second body paragraph - what do we include?"
- "Adjectives - words that describe nouns."
- "Technical language - topic-specific words."
- "Present tense - -s and -ing endings."
- "These are the three features for our paragraph."

DO:
- Read each feature aloud.
- Hold up three fingers - three features.

TEACHER NOTES:
I Do for the three writing focuses.""",

    34: """SAY:
- "Together. Turn supporting detail 1 into a full sentence."
- "SPO note: '2009 - KNY agreement = signed -> state government & Ngarrindjeri people.'"
- "Watch me build it: 'In 2009, the Kungun Ngarrindjeri Yunnan (KNY) agreement was signed between the Ngarrindjeri people and the state government.'"
- "Notice the technical language: Kungun Ngarrindjeri Yunnan, Ngarrindjeri, state."

DO:
- Display the SPO note and the sentence side by side.
- Highlight technical vocabulary in colour.
- Choral read the new sentence.

TEACHER NOTES:
I Do for the first detail conversion. Walks through how an SPO note becomes a full sentence with technical vocabulary.

WATCH FOR:
- Students who write the sentence in past tense. Note: this sentence does use 'was signed' for the historical event - that is correct. Information reports use present tense for general facts but past tense for specific historical events.""",

    35: """SAY:
- "Together. Turn supporting detail 2 into a full sentence."
- "SPO note: 'agreement - traditional owners of the land -> established processes for Aboriginal rights of The Coorong.'"
- "Watch me build it: 'This agreement recognises traditional ownership of the lands and establishes official processes for Aboriginal rights and responsibilities in regards to the Coorong.'"
- "Notice: technical language (Aboriginal, traditional ownership), present tense (recognises, establishes)."

DO:
- Whiteboards: students try their own version first.
- Compare to the model.
- Highlight technical vocab and present tense in colour.

TEACHER NOTES:
We Do for the second detail conversion.

WATCH FOR:
- Students who struggle to expand the note. Cue: 'Each note word can become a phrase.'""",

    36: """SAY:
- "Together. Turn supporting detail 3 into a full sentence."
- "SPO note: 'parks managed with Ngarrindjeri people -> preserve culture & heritage.'"
- "Watch me build it: 'The parks are currently managed with the Ngarrindjeri people to honour and preserve this area's rich cultural heritage.'"
- "Notice: adjectives (rich), technical language (Ngarrindjeri people, cultural heritage), present tense (are, managed)."

DO:
- Cold call a non-volunteer for a sentence based on the SPO note.
- Reveal the model.
- Highlight the three features.

TEACHER NOTES:
We Do for the third detail conversion.

WATCH FOR:
- Students who omit adjectives. Cue: 'Add a word that describes - rich, cultural, important.'""",

    37: """SAY:
- "Now look at the full SPO and the full paragraph side by side."
- "The SPO had the topic sentence, three notes, and a concluding sentence."
- "The paragraph has the same topic sentence, three full sentences, and the same concluding sentence."
- "Choral read the completed paragraph with me."

DO:
- Display the side-by-side comparison.
- Choral read the full paragraph.
- Cold call: 'Which sentence in the paragraph matches detail 2 of the SPO?'

TEACHER NOTES:
We Do summary. Shows the full SPO-to-paragraph transformation. Use this as the model before students write their own.

WATCH FOR:
- Students who skip the topic sentence or the concluding sentence. Press: 'Paragraph needs all three parts.'""",

    38: """SAY:
- "Hinge question. I know my body paragraph is complete when I have..."
- "Options: 1) a topic sentence, 3 supporting details and a concluding statement, 2) written in past tense, 3) introduced a new topic for each supporting detail, 4) included technical vocabulary, adjectives and written in present tense."
- "Thumbs up if true, thumbs down if not."

DO:
- Run each option one at a time.
- Wait time on each, then on cue all show.
- Reveal: 1 yes, 2 no, 3 no, 4 yes.

CFU CHECKPOINT:
Technique: Thumbs Up or Down per option
Script:
- Ask: 'Thumbs up if this is part of a complete body paragraph.' Run each option in turn.
- Scan for: thumbs up on 1 and 4, thumbs down on 2 and 3.
PROCEED:
- >=80% correctly classify all four. Move to the booklet task.
PIVOT:
- Most likely: students with thumbs up on 2 (past tense) or 3 (new topic each time). Misconception: confusing information reports with narrative or recount.
- Reteach: 'Information reports use present tense. Each supporting detail expands the same main idea, not a new topic.'
- Re-check: 'Show me again on options 2 and 3.'

TEACHER NOTES:
Key hinge for completion criteria. Answers: 1 (yes), 2 (no), 3 (no), 4 (yes).""",

    39: """SAY:
- "Open your booklet to Lesson 15: Write a body paragraph for an information report."
- "Use your SPO from Lesson 14 to write your full body paragraph."
- "First: copy your topic sentence. Next: expand each note into a full sentence. Then: write your concluding sentence."
- "Remember the three features: present tense, technical vocab, adjectives."

DO:
- Direct to the booklet.
- Set timer for 15 to 20 minutes.
- Circulate. Check first response of 3 to 5 students.
- Provide the model paragraph from slide 37 as a scaffold for students who need it.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: provide sentence frames for each detail. Student fills the gaps.
EXTENDING PROMPT:
- Task: after writing body paragraph 2, plan and write body paragraph 1 (location) using a fresh SPO.

TEACHER NOTES:
You Do task. Students draft the full body paragraph from their SPO.

WATCH FOR:
- Students writing in past tense. Cue: 'Present tense for facts (use -s and -ing endings).'
- Students missing technical vocabulary. Cue: 'Include Ngarrindjeri, Kurangk, KNY.'
- Students who copy the SPO directly without expanding. Cue: 'Each note becomes one full sentence.'""",

    40: """SAY:
- "After writing your sentences, re-read and reflect on your writing."
- "Use the checklist in your booklet to edit and complete your paragraph."
- "Checklist: TS clearly expresses main idea, supporting details expand the TS, adjectives, precise vocab, present tense, CS summarises, spelling and grammar checked."

DO:
- Direct to the checklist in the booklet.
- Set timer for 5 to 7 minutes.
- Model proofreading one sentence on the board.
- Circulate. Check students mark off each criterion.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: pair students up. Each reads their partner's paragraph and ticks the checklist for them.
EXTENDING PROMPT:
- Task: re-write one sentence with stronger adjectives or more precise vocabulary.

TEACHER NOTES:
You Do part 2 - proofread and edit. SC3 target.

WATCH FOR:
- Students who tick everything without checking. Press: 'Read each criterion and verify.'
- Students who only check spelling. Press: 'Also check tense, vocabulary, structure.'""",

    46: """SAY:
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
    print(f"L15 written: {stats}")


if __name__ == "__main__":
    main()
