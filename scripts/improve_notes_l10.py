"""Deep-pass teacher notes for Lesson 10 (Write a body paragraph for an information report) per teachernotes.md v2.0."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from polish_notes_lib import polish_deck

SRC = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated/10. literature_presentation Write a body paragraph for an information report 4En06V Storm Boy - Information report - with notes.pptx")
OUT = SRC.with_name(SRC.stem + " v2.pptx")

NOTES = {
    1: """TEACHER NOTES:
Title slide. Today's writing focus is writing and editing the first body paragraph for the Coorong information report using the SPO planned in lesson 9.""",

    7: """SAY:
- "Read the learning intention with me."
- "Today we write our first body paragraph using the SPO we planned yesterday."
- "We also edit our writing using a checklist - good writers always check their work."
- "Ask: why do you think editing matters? Expected: catches typos, makes the writing clearer."
- "If editing feels new, that is okay. We will use a checklist together."

DO:
- Choral read the LI, then track each success criterion with your finger.
- Cold call two students for the editing question.

TEACHER NOTES:
SC1 and SC3 are the writing targets the You Do tasks assess.

WATCH FOR:
- Students who cannot articulate why editing matters. Prompt: 'Have you ever read your own work and noticed a typo?'""",

    8: """SAY:
- "Quick check."
- "You need pencil, booklet, novel, mini whiteboard and texta."
- "You also need your SPO from yesterday."
- "Thumbs up when you have all six."

DO:
- Scan the room. Wait for all thumbs.

TEACHER NOTES:
Don't start until materials are out. Students need their SPO from yesterday too.

WATCH FOR:
- Students with no booklet or SPO. They need both.""",

    10: """SOURCES:
Video: Explore the Coorong - https://www.youtube.com/watch?v=Rqc3LYwgq8M

SAY:
- "Watch this short video about the Coorong."
- "Add to your notes from yesterday - especially WHAT the Coorong looks like and WHERE it is."

DO:
- Play the video.
- Give one minute for students to finalise notes.
- Cold call two students for one new observation each.

TEACHER NOTES:
Same video as lesson 9. If students saw it yesterday, replay key sections only.

WATCH FOR:
- Students who do not take notes. Cue: 'Even one new keyword counts.'""",

    11: """SAY:
- "We are taking notes on two non-fiction texts: Where is the Coorong? and Coorong National Park."
- "Use the shorthand from lesson 8 - keywords, phrases, abbreviations, symbols."
- "Focus on what the Coorong is and where it is."

DO:
- Direct students to their note-taking page.
- Set fifteen minutes.
- Circulate. Check the first few notes each student writes.

TEACHER NOTES:
This note-taking step gives students fresh material for expanding the SPO into full sentences.

WATCH FOR:
- Students who copy whole sentences. Redirect to keywords and shorthand.""",

    12: """SOURCES:
Where is the Coorong? - https://www.environment.sa.gov.au/topics/coorong/visitor-experience/where-is-the-coorong

SAY:
- "Read the non-fiction text about where the Coorong is."
- "Take notes using keywords and shorthand."
- "Focus: where is the Coorong, what does it look like."

DO:
- Direct students to the PDF or pre-printed copies.
- Set ten minutes.
- Circulate.

TEACHER NOTES:
First non-fiction text for note-taking. Pre-read so you can model on one paragraph.

WATCH FOR:
- Students who skim and miss key facts. Prompt to slow down and re-read.""",

    13: """SOURCES:
Coorong National Park Teaching Resource - https://cdn.environment.sa.gov.au/parks/docs/coorong-national-park/coorong_np_teacher_resource_2020.pdf

SAY:
- "Now we read the second non-fiction text - Coorong National Park."
- "Focus on pages 5 and 8."
- "Add to your notes."

DO:
- Direct students to the PDF.
- Set ten minutes.
- Circulate.

TEACHER NOTES:
Students add to the notes from the previous lesson.

WATCH FOR:
- Students who do not connect the two texts. Prompt: 'Combine your notes - same topic.'""",

    14: """SAY:
- "Read this text together."
- "Listen for: where the Coorong is and what it features."
- "Add to your notes."

DO:
- Choral read.
- Set five minutes.
- Cold call two students for one fact each.

TEACHER NOTES:
Facts to land: 200 km southeast of Adelaide, long and narrow, channels and lagoons, ocean beach and sand dunes.

WATCH FOR:
- Students who write only one fact. Encourage at least three.""",

    15: """SAY:
- "Read this text together."
- "Listen for: who lives in the Coorong and who looks after it."
- "Add to your notes."

DO:
- Choral read.
- Set five minutes.
- Cold call two students.

TEACHER NOTES:
Facts to land: water birds and migratory species, fishing and tourism, cared for by Department for Environment and Water, Ngarrindjeri people, conservation groups.

WATCH FOR:
- Students who skip the list of carers. Useful for body paragraph 2 later.""",

    16: """SAY:
- "Read this text together."
- "Topic: an internationally important wetland."
- "Listen for facts and figures, especially numbers."

DO:
- Choral read.
- Set five minutes.
- Ask: 'How many species need protection?' Expected: 115.

TEACHER NOTES:
Facts to land: 115 species need protection, 79 birds, 2 amphibians, 15 mammals, 4 reptiles, 15 plants. Most important waterbird wetland in Murray-Darling. Ramsar Wetland.

WATCH FOR:
- Students missing the numbers. Prompt: 'Numbers are almost always keywords.'""",

    17: """SAY:
- "Read this text together."
- "Topic: wetland (continued)."
- "Listen for: where the migratory birds come from."

DO:
- Choral read.
- Set five minutes.
- Ask: 'Where do some birds travel from?' Expected: Alaska, Siberia.

TEACHER NOTES:
Facts to land: shorebirds travel from Alaska and Siberia. Ramsar Convention is an international wetland agreement. SA has six other Ramsar sites.

WATCH FOR:
- Students confused by 'migratory'. Quick gloss: 'Birds that travel long distances.'""",

    18: """SAY:
- "Read this text together, slowly and respectfully."
- "Topic: cultural significance of the Coorong."
- "Listen for: the Ngarrindjeri name and what middens are."

DO:
- Choral read.
- Set five minutes.
- Ask: 'What does Kurangk mean?' Expected: long narrow neck.

SENSITIVITY ADVISORY:
- What it is: discussion of Ngarrindjeri cultural sites and middens.
- Framing language: 'Middens are special places that show how Ngarrindjeri people have lived on this Country for thousands of years.'
- Watch for: any student speaking carelessly about cultural sites.
- Protocol: redirect to respectful language. Pause if a student is upset.

TEACHER NOTES:
This text introduces the cultural focus. Useful for body paragraph 2.

WATCH FOR:
- Students mispronouncing Kurangk or Ngarrindjeri. Model the pronunciation slowly.""",

    19: """SAY:
- "Read this text together, respectfully."
- "Topic: cultural significance (continued)."
- "Listen for: the KNY Agreement and what it means."

DO:
- Choral read.
- Set five minutes.
- Ask: 'When was the KNY Agreement signed?' Expected: 2009.

TEACHER NOTES:
Key fact: KNY = Kungun Ngarrindjeri Yunnan. Means 'Listen to what Ngarrindjeri people have to say.'

WATCH FOR:
- Students confused by the language. Keep concrete: 'An agreement to listen and work together.'""",

    20: """SAY:
- "Check for understanding."
- "What is the Ngarrindjeri people's vision of the Coorong? Select all that apply."
- "Whiteboards: option numbers."

DO:
- Wait. Show me. Scan.
- Reveal: 1, 2, 4 (respect Country and creation; clean sparkling waters; caring, sharing, respecting). Option 3 is not in the source text.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Show me the option numbers.' Expected: 1, 2 and 4.
- Scan for: those numbers on most boards.
PROCEED:
- >=80% correct. Move into the writing focus.
PIVOT:
- Most likely: students only pick one option.
- Reteach: 'The prompt asks for ALL that apply. Re-read each option.'
- Re-check: 'Show me again.'

WATCH FOR:
- Students who only pick one option. They may have skim-read the prompt.""",

    22: """SAY:
- "Quick reminder of what an information report does."
- "It informs the reader. It classifies or describes factual information about a specific topic."
- "We start by reading the mentor text together."

DO:
- Point to the definition.
- Set up: 'You are about to revisit the model text from yesterday.'

TEACHER NOTES:
Review from yesterday. Keep brief - the mentor text is on the next slides.

WATCH FOR:
- Students who cannot recall what an information report is. Quick gloss: 'Tells facts.'""",

    23: """SAY:
- "Read the introduction with me."
- "Notice the General Statement, Specific Topic and Topic Outline (GST formula)."

DO:
- Choral read.
- Cold call: 'What is the General Statement?'
- Cold call: 'What does the Topic Outline list?'

TEACHER NOTES:
Review the GST formula. The introduction sets up the report.

WATCH FOR:
- Students who cannot identify GST parts. Re-read with the labels.""",

    24: """SAY:
- "Read body paragraph one with me."
- "Topic: what the Coorong is and where it is located."
- "This is the paragraph YOU will write today using your SPO."
- "Notice the topic sentence, three supporting details, and concluding sentence."

DO:
- Choral read.
- Highlight the TS, SD, and CS as you go.
- Cold call: 'How many supporting details?' Expected: three.

TEACHER NOTES:
This is the model. Students base their paragraph on it. They do not copy it.

WATCH FOR:
- Students who try to copy the model. Anchor: 'Use it as a guide for STRUCTURE, not for words.'""",

    25: """SAY:
- "Look at the supporting detail sentences."
- "Find the adjectives that describe the water in the Coorong."
- "Whiteboards."

DO:
- Wait. Show me.
- Reveal: freshwater, brackish, saltwater, salty, fresh.
- Discuss: 'Information reports use precise descriptive words.'

TEACHER NOTES:
Links the model to the language features. Adjectives are one of the four focus features.

WATCH FOR:
- Students who write nouns instead of adjectives. Anchor: 'Adjectives describe.'""",

    26: """SAY:
- "Read body paragraph two with me."
- "Topic: cultural significance."
- "Notice the same TS-SD-CS structure."

DO:
- Choral read.
- Cold call: 'What is the topic sentence?'
- Cold call: 'What is the concluding sentence?'

TEACHER NOTES:
Reinforces the structure pattern.

WATCH FOR:
- Students who confuse TS and CS. Anchor: 'TS comes FIRST. CS comes LAST.'""",

    27: """SAY:
- "Read body paragraph three with me."
- "Topic: flora and fauna."
- "Notice the same TS-SD-CS structure."

DO:
- Choral read.
- Cold call: 'What is the topic sentence?'

TEACHER NOTES:
Third review of the structure. Students should now be confident.

WATCH FOR:
- Students who can label but cannot apply. More guided practice next.""",

    28: """SAY:
- "Find the technical or precise vocabulary in this extract."
- "Whiteboards: write the technical words."

DO:
- Wait. Show me.
- Reveal: ecosystem, vulnerable, endangered, wetland, migratory.
- Discuss: 'Technical vocabulary is precise. Use it because the topic needs it.'

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Show me the technical words you found.' Expected: ecosystem, migratory, endangered, wetland.
- Scan for: at least three on most boards.
PROCEED:
- >=80% find at least three. Move into the writing block.
PIVOT:
- Most likely: students pick everyday words like 'home' or 'area'.
- Reteach: 'Technical words are specific to the topic. You would not use ecosystem in everyday speech.'
- Re-check: 'Show me one or two technical words.'

WATCH FOR:
- Students who pick zero technical words. Give them one example to start.""",

    29: """SAY:
- "Read the conclusion with me."
- "The conclusion summarises the report and ends with a memorable closing line."

DO:
- Choral read.
- Cold call: 'Does the conclusion add new information?' Expected: no - it sums up.

TEACHER NOTES:
Close the mentor text review.

WATCH FOR:
- Students who think the conclusion adds new ideas. Clarify: 'It sums up only.'""",

    30: """SAY:
- "Today we write body paragraph one."
- "We use our notes from yesterday and today."
- "The topic is the Coorong. The paragraph focus is what it is and where it is located."
- "The plan from yesterday becomes the paragraph today."

DO:
- Tell students to get their SPO from yesterday ready.
- Point to the topic and paragraph focus.

TEACHER NOTES:
Sets up the writing block. Make sure every student has their SPO from lesson 9.

WATCH FOR:
- Students with no SPO from yesterday. Pair them with a buddy or quickly help them rebuild.""",

    31: """SAY:
- "A body paragraph follows the SPO structure."
- "Sentence 1: topic sentence introduces the location of the Coorong."
- "Sentences 2 to 4: three details that tell us more."
- "Sentence 5: concluding sentence that links back to the topic sentence."
- "Five sentences total."

DO:
- Point to each sentence position.
- Cold call: 'How many supporting detail sentences?' Expected: three.

TEACHER NOTES:
Writing structure. Each SPO becomes five sentences.

WATCH FOR:
- Students who try to write more than five sentences. Fine if controlled, but keep the focus tight.""",

    32: """SAY:
- "Match the topic sentence to a suitable concluding sentence."
- "Whiteboards: write the matching pair number."

DO:
- Wait. Show me.
- Reveal correct pairs:
- TS 'The Coorong is a large expanse...' matches CS 'The Coorong is part of the Murray-Darling basin...'
- TS 'The Coorong, named Kurangk...' matches the cultural-significance CS.

TEACHER NOTES:
Tests whether students see the link between TS and CS - same topic, different words.

WATCH FOR:
- Students who match by length not meaning. Prompt: 'Do they talk about the same big idea?'""",

    33: """SAY:
- "Which statement about topic sentences is true?"
- "Whiteboards: option number."

DO:
- Wait. Show me.
- Reveal: 'Topic sentences express a paragraph's main idea.'

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Show me the option number.' Expected: the option about expressing the main idea.
- Scan for: that option on most boards.
PROCEED:
- >=80% correct. Move into the literacy book task.
PIVOT:
- Most likely: students pick the lengthy option.
- Reteach: 'A topic sentence is one sentence that says the main idea.'
- Re-check: 'Show me again.'

WATCH FOR:
- Students who pick the dot-point option. They confused topic sentence with notes.""",

    34: """SAY:
- "Open your Literacy Book and your SPO from yesterday."
- "Read your topic sentence and concluding sentence to a partner."
- "Ask: do my sentences show a clear link?"
- "If yes, copy your topic sentence into your Literacy Book."
- "If no, edit first, then copy."

DO:
- One minute partner check.
- Set five minutes for editing and copying.
- Circulate. Check the topic sentence each student copies.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: link-check question: 'Do both sentences talk about WHAT the Coorong is and WHERE it is?'
- Extra Notes: students who say no edit before copying.
EXTENDING PROMPT:
- Task: rewrite the topic sentence two ways - one starting with WHAT, one with WHERE.
- Extra Notes: praise variation in structure.

TEACHER NOTES:
First writing You Do for lesson 10. Make sure students have the SPO from lesson 9.

WATCH FOR:
- Students whose TS and CS do not link. Edit before copying.
- Students with no SPO from yesterday. Give them five minutes to plan first.""",

    35: """SAY:
- "Supporting details give additional information about the paragraph focus."
- "The sentences after the topic sentence add information about the Coorong's location and what it looks like."
- "Each detail should be a different fact - not three versions of the same fact."

DO:
- Point to where supporting details sit.
- Cold call: 'How many supporting detail sentences?' Expected: three.

TEACHER NOTES:
Review the role of supporting details before students expand notes into sentences.

WATCH FOR:
- Students who think supporting details all need to start the same way. Prompt for variety.""",

    36: """SAY:
- "Watch this first."
- "We expand SPO notes into full sentences."
- "Each sentence should include adjectives, technical language, and present tense."
- "These three things make our writing sound like a real information report."

DO:
- Write a quick example on the board: 'long, narrow lagoon' (adjectives + technical), 'is located' (present tense).
- Cold call: 'Why present tense?' Expected: the facts are still true now.

TEACHER NOTES:
I Do for the language features. Keep under three minutes.

WATCH FOR:
- Students who slip into past tense. That is the most common mistake.""",

    37: """SAY:
- "Watch this first."
- "Note: more than 100 km long and separated from the ocean by a sandy shore = Younghusband Peninsula."
- "Full sentence: 'It is more than 100 km long and is separated from the ocean by a length of sandy shore named the Younghusband Peninsula.'"
- "Technical: Younghusband Peninsula. Adjective: sandy. Present tense: is, is separated."

DO:
- Build the sentence on the board piece by piece.
- Highlight each language feature.
- Cold call: 'What is the technical word?' Expected: Younghusband Peninsula.

TEACHER NOTES:
I Do model of expanding a note into a full sentence with all three language features.

WATCH FOR:
- Students who keep the note format. Anchor: 'Full sentences with capital letters and full stops.'""",

    38: """SAY:
- "We Do."
- "Note: lagoon, containing freshwater, brackish water and salt water."
- "Turn to your partner. Compose a full sentence with present tense, technical language and adjectives."

DO:
- Two minutes partner talk.
- Cold call non-volunteers.
- Reveal: 'It is a lagoon that contains a mix of freshwater, brackish water, and saltwater across its various stretches of water.'
- Highlight: lagoon (technical), freshwater/brackish/saltwater (adjectives), is/contains (present tense).

TEACHER NOTES:
Use this slide to analyse the text together as a class.

WATCH FOR:
- Students who write past tense ('contained'). Quick reteach.""",

    39: """SAY:
- "Round two."
- "Note: mixture of water = sits between salty water of ocean and fresh water of river."
- "Compose a full sentence."
- "Whiteboards."

DO:
- Set three minutes. Show me.
- Ask a non-volunteer to share their sentence.
- Reveal: 'This is because it sits between the salty water of the ocean and the freshwater rivers of the basin.'
- Identify: salty (adjective), freshwater (adjective + technical), sits (present tense).

TEACHER NOTES:
The slide shows 'Past tense' as a NON-EXAMPLE. Confirm to students that present tense is what they should use.

WATCH FOR:
- Students confused by the 'Past tense' label on the slide. Clarify: 'That is what we AVOID.'""",

    40: """SAY:
- "This is the completed body paragraph."
- "Read it with me."
- "Notice how the topic sentence, three supporting details and concluding sentence flow together."

DO:
- Choral read.
- Then a non-volunteer reads alone.
- Cold call: 'What language features can you spot?' Expected: present tense, adjectives, technical words.

TEACHER NOTES:
Chorally read and review the finished paragraph. Model for the You Do.

WATCH FOR:
- Students who do not see the connection to their own SPO. Hold the SPO and the model side by side.""",

    41: """SAY:
- "Check for understanding."
- "I know my body paragraph is complete when I have..."
- "Whiteboards: write all option numbers that apply."

DO:
- Wait. Show me.
- Reveal: 1, 2, 4. Option 3 (past tense) is FALSE.
- Option 1: TS, three SD, CS. Option 2: new topic per SD. Option 4: technical, adjectives, present tense.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Show me the option numbers - all that apply.' Expected: 1, 2 and 4.
- Scan for: those numbers on most boards.
PROCEED:
- >=80% correct. Move into the You Do.
PIVOT:
- Most likely: students include option 3 (past tense).
- Reteach: 'Information reports are PRESENT tense. Past tense is for stories.'
- Re-check: 'Show me again.'

WATCH FOR:
- Students who pick option 3. Quick reteach on tense before they write.""",

    42: """SAY:
- "Open your Literacy Book."
- "First: expand your SD1 note into a full sentence in your Literacy Book."
- "Next: expand SD2 and SD3."
- "Then: copy your concluding sentence from your SPO."
- "Use present tense, adjectives, technical language."
- "Twenty minutes."

DO:
- Direct students to their SPO and Literacy Book.
- Set twenty minutes.
- Circulate. Read the first SD sentence each student writes.
- Pull two strong examples to share.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: sentence starter for each SD: 'It is...' 'It also...' 'In addition...'
- Extra Notes: students who need support work alongside a buddy.
EXTENDING PROMPT:
- Task: write SD sentences using different starters - one starting with the subject, one with a phrase.
- Extra Notes: praise sentence variety.

TEACHER NOTES:
Main You Do for SC1 and SC3. Track who writes a complete five-sentence paragraph.

WATCH FOR:
- Students who write past tense. Cue: 'Check your verbs.'
- Students whose sentences do not link to the topic sentence. Prompt: 'Does this fact connect to the main idea?'""",

    43: """SAY:
- "Now we edit."
- "After writing, re-read your paragraph."
- "Use the checklist in your booklet."
- "Look for: present tense, adjectives, technical words, capitals, full stops, spelling."
- "Five to ten minutes."

DO:
- Set ten minutes.
- Circulate. Sit next to two students and walk through the checklist.
- Praise students making real edits, not just rereading.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: colour-code the checklist - circle present tense verbs in blue, underline technical words in green.
- Extra Notes: editing as a visual hunt makes it concrete.
EXTENDING PROMPT:
- Task: swap with a partner and edit each other's paragraph.
- Extra Notes: model giving kind, specific feedback.

TEACHER NOTES:
Final You Do for SC3. Editing is a separate skill from writing - protect the time.

WATCH FOR:
- Students who re-read without editing. Prompt: 'Make at least one change you can see.'""",

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
    print(f"L10 written: {stats}")


if __name__ == "__main__":
    main()
