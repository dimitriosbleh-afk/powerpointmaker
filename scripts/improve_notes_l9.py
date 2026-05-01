"""Deep-pass teacher notes for Lesson 9 (Plan a body paragraph for an information report) per teachernotes.md v2.0."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from polish_notes_lib import polish_deck

SRC = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated/9. literature_presentation Plan a body paragraph for an information report 4En06V Storm Boy - Information report - with notes.pptx")
OUT = SRC.with_name(SRC.stem + " v2.pptx")

NOTES = {
    1: """TEACHER NOTES:
Title slide. Today's writing focus is planning a body paragraph for an information report on the Coorong using a Single Paragraph Outline (SPO).""",

    7: """SAY:
- "Read the learning intention with me."
- "Today we plan our first body paragraph for an information report on the Coorong."
- "An information report tells the reader about a topic using facts."
- "We use a Single Paragraph Outline - SPO - to plan each body paragraph before we write."
- "Ask: what do you remember about an SPO? Expected: a plan with a topic sentence, supporting details, concluding sentence."

DO:
- Choral read the LI, then track each success criterion with your finger.
- Cold call two students for the SPO question.

TEACHER NOTES:
SC3 is the planning target the You Do tasks assess.

WATCH FOR:
- Students who cannot recall what an information report is. Quick gloss: 'A non-fiction text that tells facts.'""",

    10: """SOURCES:
Video: Explore the Coorong - https://www.youtube.com/watch?v=Rqc3LYwgq8M

SAY:
- "Watch this short video about the Coorong."
- "This is the place Storm Boy lives."
- "Take notes - what does the Coorong look like, where is it, what lives there?"
- "Use the keywords and shorthand from yesterday."

DO:
- Play the video.
- Give one minute for students to finalise notes.
- Cold call two students for one observation each.

TEACHER NOTES:
Visual anchor for the Coorong. Pre-watch and pre-mark moments to highlight.

WATCH FOR:
- Students who do not take notes during the video. Cue: 'At least one keyword every 30 seconds.'""",

    11: """SOURCES:
Coorong National Park Teaching Resource - https://cdn.environment.sa.gov.au/parks/docs/coorong-national-park/coorong_np_teacher_resource_2020.pdf

SAY:
- "Now we read a non-fiction text about the Coorong."
- "Focus on pages 2 and 4."
- "Add to your notes about WHAT the Coorong is and WHERE it is."
- "Use keywords and shorthand."

DO:
- Direct students to the PDF or pre-printed copies.
- Read pages 2 and 4 aloud or allow paired reading.
- Set ten minutes for note-taking.

TEACHER NOTES:
First non-fiction text reading. Pre-read so you can model note-taking on one paragraph.

WATCH FOR:
- Students copying full sentences. Prompt back to keywords and symbols.""",

    12: """SAY:
- "Read this text together."
- "Listen for: where the Coorong is and what it looks like."
- "Add to your notes."

DO:
- Choral read.
- Set five minutes for note-taking.
- Cold call two students for one fact each.

TEACHER NOTES:
Facts to land: 200 km southeast of Adelaide, long and narrow, Goolwa Barrage to nearly Kingston, channels and lagoons, ocean beach and sand dunes.

WATCH FOR:
- Students who write only one fact. Encourage at least three.""",

    13: """SAY:
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

    14: """SAY:
- "Read this text together."
- "Topic: an internationally important wetland."
- "Listen for facts and figures, especially numbers."

DO:
- Choral read.
- Set five minutes.
- Ask: 'How many species need protection?' Expected: 115.

TEACHER NOTES:
Facts to land: 115 species need protection (79 birds, 2 amphibians, 15 mammals, 4 reptiles, 15 plants). Most important waterbird wetland in Murray-Darling. Ramsar Wetland.

WATCH FOR:
- Students missing the numbers. Prompt: 'Numbers are almost always keywords.'""",

    15: """SAY:
- "Read this text together."
- "Topic: an internationally important wetland (continued)."
- "Listen for: where the migratory birds come from."

DO:
- Choral read.
- Set five minutes.
- Ask: 'Where do some birds travel from?' Expected: Alaska, Siberia.

TEACHER NOTES:
Facts to land: shorebirds travel from Alaska and Siberia. Ramsar Convention is an international wetland agreement. SA has six other Ramsar sites.

WATCH FOR:
- Students confused by 'migratory'. Quick gloss: 'Birds that travel long distances.'""",

    16: """SAY:
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
- Protocol: redirect to respectful language. Pause if a student is upset, follow up at recess if needed.

TEACHER NOTES:
This text introduces the cultural focus. Useful for body paragraph 2.

WATCH FOR:
- Students mispronouncing Kurangk or Ngarrindjeri. Model the pronunciation slowly.""",

    17: """SAY:
- "Read this text together, respectfully."
- "Topic: cultural significance (continued)."
- "Listen for: the KNY Agreement and what it means."

DO:
- Choral read.
- Set five minutes.
- Ask: 'When was the KNY Agreement signed?' Expected: 2009.

TEACHER NOTES:
Key fact: KNY = Kungun Ngarrindjeri Yunnan. Means 'Listen to what Ngarrindjeri people have to say.' Recognises Ngarrindjeri ownership of lands and waters.

WATCH FOR:
- Students confused by the language. Keep concrete: 'An agreement between two governments to listen and work together.'""",

    18: """SAY:
- "Check for understanding."
- "What does the Coorong National Park feature?"
- "Whiteboards: write all option numbers that apply."

DO:
- Wait. Show me. Scan.
- Reveal: 1, 2, 4 (channels, lagoons, ocean beach and sand dunes). Option 3 is wrong - it is the Southern Ocean, not Indian.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Show me the correct option numbers.' Expected: 1, 2 and 4.
- Scan for: those numbers on most boards.
PROCEED:
- >=80% correct. Move into the writing focus.
PIVOT:
- Most likely: students include option 3 (Indian Ocean).
- Reteach: 'The Coorong is in South Australia, on the Southern Ocean side.'
- Re-check: 'Show me the correct numbers.'

WATCH FOR:
- Students who only pick one option. The prompt asks for ALL.""",

    20: """SAY:
- "Watch this first."
- "The purpose of an information report is to inform the reader."
- "It does this by classifying or describing factual information about a specific topic."
- "Today our topic is the Coorong."
- "This is the trap: information reports give facts, not opinions."

DO:
- Point to the definition as you read it.
- Cold call: 'What is the difference between an information report and a story?' Expected: facts vs imagined events.

TEACHER NOTES:
Genre I Do anchor. Keep under three minutes.

WATCH FOR:
- Students who confuse information report with persuasive text. Clarify: 'Facts, not opinions.'""",

    21: """SAY:
- "An information report has this structure."
- "Title. Introduction. Three body paragraphs - each with one main idea. Conclusion."
- "Today we are planning body paragraph one using an SPO."

DO:
- Point to each section on the structure diagram.
- Cold call: 'How many body paragraphs?' Expected: three.
- Cold call: 'What goes in each body paragraph?' Expected: one main idea.

TEACHER NOTES:
Structure overview. Keep under three minutes.

WATCH FOR:
- Students who think a body paragraph can have multiple ideas. Clarify: 'One idea per paragraph.'""",

    22: """SAY:
- "Our topic is the Coorong."
- "The focus for body paragraph one is what the Coorong is and where it is located."
- "Before we plan, we read the mentor text - the model that shows us what good looks like."

DO:
- Point to the topic and the paragraph focus.
- Set up: 'You are about to read a sample report from start to finish.'

TEACHER NOTES:
Sets up the mentor text reading.

WATCH FOR:
- Students unsure what a mentor text is. Quick gloss: 'A model we copy the structure from, not the words.'""",

    23: """SAY:
- "Read the introduction with me."
- "Listen for the General Statement, Specific Topic and Topic Outline (GST formula)."

DO:
- Choral read.
- Cold call: 'What is the general statement?' Expected: the Coorong is a narrow shallow lagoon...
- Cold call: 'What does the report say it will discuss?' Expected: what it is, where, cultural significance, flora and fauna.

TEACHER NOTES:
GST = General statement, Specific topic, Topic outline. The introduction tells readers what is coming.

WATCH FOR:
- Students who cannot separate the three parts. Re-read each sentence with the GST label.""",

    24: """SAY:
- "Read body paragraph one with me."
- "Topic: what the Coorong is and where it is located."
- "Listen for the topic sentence (TS), supporting details (SD) and concluding sentence (CS)."

DO:
- Choral read.
- Cold call: 'What is the topic sentence?' Expected: An integral part of the Murray-Darling basin...
- Cold call: 'What is the concluding sentence?' Expected: The Coorong is a large expanse of different types of water...

TEACHER NOTES:
This is the body paragraph students plan today. Use it as the model.

WATCH FOR:
- Students who cannot find the supporting details. They are the three sentences in the middle.""",

    25: """SAY:
- "Read body paragraph two with me."
- "Topic: cultural significance."
- "Listen for the TS, SD and CS."

DO:
- Choral read.
- Cold call: 'What is the topic sentence?' Expected: The Coorong, named Kurangk by the Ngarrindjeri people...
- Cold call: 'What is the concluding sentence?' Expected: By prioritising the conservation and custodianship...

TEACHER NOTES:
Second model paragraph. Same TS-SD-CS structure as paragraph one.

WATCH FOR:
- Students who cannot find the supporting details. Prompt: 'What facts back up the topic sentence?'""",

    26: """SAY:
- "Read body paragraph three with me."
- "Topic: flora and fauna."
- "Listen for the TS, SD and CS."

DO:
- Choral read.
- Cold call: 'What is the topic sentence?' Expected: The Coorong ecosystem is unique and complex...
- Cold call: 'What is the concluding sentence?' Expected: The Coorong is a rich environment teeming with various species...

TEACHER NOTES:
Third model paragraph. Reinforces the TS-SD-CS structure.

WATCH FOR:
- Students still confused about TS vs CS. Anchor: 'TS introduces. CS sums up.'""",

    27: """SAY:
- "Read the conclusion with me."
- "The conclusion summarises what the report covered and ends with a memorable closing thought."

DO:
- Choral read.
- Cold call: 'What does the conclusion remind us of?' Expected: the report's main points.
- Cold call: 'What is the closing thought?' Expected: the Coorong is an iconic Australian landscape.

TEACHER NOTES:
Close the mentor text. Students have now seen the full structure.

WATCH FOR:
- Students who think the conclusion adds new information. Clarify: 'It summarises and closes.'""",

    28: """SAY:
- "Information reports use specific language features."
- "Read with me: present tense, third person pronouns, subheadings, simple and expanded noun groups, adjectives and adverbials, technical vocabulary, action verbs."
- "Today we focus on subheadings, present tense, technical vocabulary and adjectives."

DO:
- Highlight the four focus features in red on the slide.
- Cold call: 'Why might information reports use third person?' Expected: sounds factual, not personal.

TEACHER NOTES:
Focus on the four red features. Other features come in later lessons.

WATCH FOR:
- Students who try to write in first person. Anchor: 'Third person is the rule.'""",

    29: """SAY:
- "Subheadings categorise and organise our paragraphs."
- "Each new paragraph (each new SPO) gets a new subheading."
- "The subheading tells the reader the key idea coming next."

DO:
- Point to the role of subheadings.
- Cold call: 'What is the subheading for our body paragraph one?' Expected: What it is and where it is located.

TEACHER NOTES:
Keep brief. Students should write a subheading for each SPO.

WATCH FOR:
- Students who do not see the value. Explain: 'They help the reader find what they are looking for.'""",

    30: """SAY:
- "Look at body paragraph one."
- "Write the subheading on your whiteboard."

DO:
- Wait. Show me.
- Reveal: 'What is it and where is it located.'

WATCH FOR:
- Students who write the topic sentence by mistake. The subheading is shorter and sits ABOVE.""",

    31: """SAY:
- "Suggest a subheading for body paragraph two."
- "Write a subheading on your whiteboard - five words or fewer."

DO:
- Wait. Show me.
- Share two student suggestions and the slide examples: 'The cultural significance of the Coorong' or 'Custodianship of the Coorong.'

TEACHER NOTES:
Writing practice with subheadings. Accept any subheading capturing cultural significance in five words or fewer.

WATCH FOR:
- Students who write a full sentence. Cue: 'Subheadings are short.'""",

    32: """SAY:
- "What are the key purposes of a subheading?"
- "Whiteboards: write all option numbers that apply."

DO:
- Wait. Show me.
- Reveal: 1, 3, 4 (organise, indicate change of focus, give the reader the key idea). Option 2 (entertain) is wrong.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Show me the correct option numbers.' Expected: 1, 3 and 4.
- Scan for: those numbers on most boards.
PROCEED:
- >=80% correct. Move into present tense.
PIVOT:
- Most likely: students include option 2.
- Reteach: 'Information reports give facts. Entertaining is for stories.'
- Re-check: 'Show me again.'

WATCH FOR:
- Students who only pick one. The prompt allows multiple correct answers.""",

    33: """SAY:
- "Information reports are usually written in present tense."
- "Present tense is the now tense: is, are, has, have, provides, includes."
- "Read body paragraph one with me. Notice the present tense words."

DO:
- Read the paragraph with the class.
- Highlight the present tense verbs: is located, is more than, is separated, is a lagoon, contains, sits.
- Ask: 'Why might we use present tense for facts?' Expected: because facts are still true now.

TEACHER NOTES:
I Do for present tense. Keep recognition-level, not full grammar lecture.

WATCH FOR:
- Students who try to spot past tense. Confirm: 'Those would be wrong here.'""",

    34: """SAY:
- "We Do."
- "Read each sentence. Identify the word(s) that show present tense."
- "Sentence 1: 'The Coorong is home to various water birds and migratory species, providing them with a safe shelter.'"
- "Sentence 2: 'The Coorong provides a habitat for a wide variety of plants and animals.'"
- "Whiteboards."

DO:
- Wait. Show me.
- Reveal: is, providing (sentence 1). Provides (sentence 2).

WATCH FOR:
- Students who pick a noun by mistake. Anchor: 'Present tense is shown by the verb.'""",

    35: """SAY:
- "Open your Writer's Notebook to today's planning page."
- "Q1: copy the subheading - 'What is it and where is it located.'"
- "Q2 we will work through together on the next slides."

DO:
- Direct students to the planning page.
- Two minutes for Q1.
- Circulate.

TEACHER NOTES:
First You Do task. Subheading only. Q2 builds the SPO across the next slides.

WATCH FOR:
- Students who skip the subheading and start the topic sentence. Redirect.""",

    36: """SAY:
- "Now we plan our SPO using the notes from the non-fiction texts."
- "The SPO is your paragraph plan. It organises ideas before you write."
- "Use class discussion and your notes. You don't need to remember everything."

DO:
- Point to where students should put their notes.
- Set up: 'You will add to this plan in pieces - one part at a time.'

TEACHER NOTES:
Sets up the SPO planning. Students need their notes from earlier in the lesson.

WATCH FOR:
- Students who do not have notes. Pair them with a buddy or summarise key facts on the board.""",

    37: """SAY:
- "An SPO is a Single Paragraph Outline."
- "It puts your writing in an order that is easy for readers to understand."
- "It is like a plan for your paragraph."
- "It has a topic sentence, supporting details, and a concluding sentence."

DO:
- Point to the definition.
- Read along with the class (choral response).
- Cold call: 'Why do we plan before we write?' Expected: so the writing flows and stays organised.

TEACHER NOTES:
The SPO is the planning template. Creating one makes writing each paragraph easier.

WATCH FOR:
- Students who think planning is optional. Anchor: 'The plan IS the writing tool.'""",

    38: """SAY:
- "Can you name the different parts of an SPO?"
- "The three parts: topic sentence, supporting details, concluding sentence."
- "Turn to your partner. Talk about what each part does."

DO:
- One minute partner talk.
- Cold call non-volunteers.
- Confirm: TS = main idea. SD = three details. CS = sums up.

TEACHER NOTES:
Facilitate partner talk so students articulate what each part does.

WATCH FOR:
- Students who can name parts but not what they do. Prompt: 'What is the topic sentence FOR?'""",

    39: """SAY:
- "Look at body paragraph one again."
- "Identify the topic sentence (TS), supporting details (SD), and concluding sentence (CS)."
- "TS expresses the main idea. SD provides three supporting details. CS sums up."

DO:
- Highlight each section as you read.
- Cold call: 'Which sentence is the TS?' Expected: An integral part of the Murray-Darling basin...
- Cold call: 'Which sentence is the CS?' Expected: The Coorong is a large expanse...

TEACHER NOTES:
Reinforces the TS-SD-CS formula in the model text.

WATCH FOR:
- Students who pick the wrong TS. Anchor: 'The TS is the FIRST sentence of the paragraph.'""",

    40: """SAY:
- "True or false?"
- "The SPO of an information report should include three supporting details that align with the subheading."
- "Whiteboards: T or F."

DO:
- Wait. Show me. Scan.
- Reveal: T.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'T or F. Show me.' Expected: T.
- Scan for: T on most boards.
PROCEED:
- >=80% show T. Move into the topic sentence block.
PIVOT:
- Most likely: students show F because they think one detail is enough.
- Reteach: 'Three details give the reader enough evidence.'
- Re-check: 'Show me again.'

WATCH FOR:
- Students who cannot choose. They may not have understood the question.""",

    41: """SAY:
- "The topic sentence is the first sentence of a paragraph."
- "It introduces the main idea."
- "It keeps your writing clear."
- "It grabs the reader's attention."
- "It tells the reader what to expect."

DO:
- Point to each purpose as you read.
- Cold call: 'Why does the topic sentence go FIRST?' Expected: so the reader knows what is coming.

TEACHER NOTES:
Review the purpose and features of a topic sentence.

WATCH FOR:
- Students who think the topic sentence is the title. Clarify: 'It is a full sentence INSIDE the paragraph.'""",

    42: """SAY:
- "Look at the topic sentence from body paragraph one."
- "'An integral part of the Murray-Darling basin, the Coorong is located at the end of the Murray River.'"
- "This tells us TWO things: WHAT the Coorong is and WHERE it is."
- "It answers both parts of our subheading."

DO:
- Highlight the WHAT and the WHERE pieces on the slide.
- Cold call: 'Where does this say the Coorong is?' Expected: at the end of the Murray River.

TEACHER NOTES:
Link the topic sentence back to the subheading. Both must answer the same focus.

WATCH FOR:
- Students who think this is one fact. Anchor: 'It is two: what AND where.'""",

    43: """SAY:
- "Open your Writer's Notebook."
- "Q2: write your topic sentence for body paragraph one."
- "Remember to answer WHAT and WHERE."
- "Five minutes."

DO:
- Set five minutes.
- Circulate. Check the topic sentence each student writes.
- Pull two strong examples to share later.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: sentence frame: 'The Coorong is ___ and is located ___.'
- Extra Notes: students fill in the blanks using their notes.
EXTENDING PROMPT:
- Task: write two versions, one starting with WHAT, one with WHERE.
- Extra Notes: students choose the better one.

TEACHER NOTES:
First writing You Do. Tracks SC2 and SC3.

WATCH FOR:
- Students whose topic sentence misses WHAT or WHERE. Redirect.
- Students who copy the model verbatim. Encourage own words.""",

    44: """SAY:
- "Supporting details give us additional information."
- "They sit between the topic sentence and the concluding sentence."
- "On the SPO they are NOTES - short keywords. We expand them into full sentences when we write."

DO:
- Point to where supporting details sit.
- Cold call: 'How many supporting details?' Expected: three.

TEACHER NOTES:
Supporting details are NOTES on the SPO. They become full sentences in lesson 10.

WATCH FOR:
- Students who try to write full sentences. Redirect: 'Notes only on the SPO.'""",

    45: """SAY:
- "Use your notes to think of three supporting details."
- "Prompts: which river is the Coorong on? How long? How far from Adelaide? Where does it start and finish? What types of water?"
- "Write three on your whiteboard."

DO:
- Set five minutes.
- Circulate. Check that students have varied details, not three versions of the same fact.
- Share two strong examples.

TEACHER NOTES:
Students may use notes from the previous lesson. Facilitate class discussion before writing if needed.

WATCH FOR:
- Students who repeat the same fact in different words. Prompt: 'Three DIFFERENT facts.'""",

    46: """SAY:
- "Open your Writer's Notebook."
- "Q2 supporting details: write three pieces of supporting detail."
- "These are NOTES, not full sentences."
- "Five minutes."

DO:
- Set five minutes.
- Circulate. Check that students wrote NOTES.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: three prompts (one per detail): 'How long? How far from Adelaide? What types of water?'
- Extra Notes: answer each in five words or fewer.
EXTENDING PROMPT:
- Task: choose three details that build a clear picture - the third detail adds something new, not a repeat.
- Extra Notes: praise variety in detail choice.

TEACHER NOTES:
Main You Do for SC3. Track which students write usable notes.

WATCH FOR:
- Students who write three sentences instead of three notes. Redirect.""",

    47: """SAY:
- "The concluding sentence summarises the main ideas of the paragraph."
- "It links back to the topic sentence - same big idea, different wording."
- "It does NOT add new information."

DO:
- Point to the role of the concluding sentence.
- Cold call: 'Should the CS introduce a new fact?' Expected: no - it sums up.

TEACHER NOTES:
CS is often where students go wrong. They add new info. Emphasise: 'Sums up only.'

WATCH FOR:
- Students who think CS introduces the next paragraph. Anchor: 'That is a transition, not a CS.'""",

    48: """SAY:
- "What is the purpose of the concluding sentence?"
- "Whiteboards: option number."

DO:
- Wait. Show me.
- Reveal: 'To remind the reader of the main idea of the paragraph.'

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Show me the option number.' Expected: the option about reminding the reader of the main idea.
- Scan for: that option on most boards.
PROCEED:
- >=80% correct. Move into the You Do.
PIVOT:
- Most likely: students pick the option about telling the reader what comes next.
- Reteach: 'The CS looks BACK at the paragraph it ends, not forward.'
- Re-check: 'Show me again.'

WATCH FOR:
- Students who pick the option about new conclusions. Clarify: 'No new ideas.'""",

    49: """SAY:
- "Open your Writer's Notebook."
- "Q2 CS: write your concluding sentence."
- "Summarise the main ideas. Don't add new information."
- "Five minutes."

DO:
- Set five minutes.
- Circulate. Check that students aren't adding new facts.
- Pull two strong examples.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: sentence frame: 'Overall, the Coorong is ___ and is located ___.'
- Extra Notes: students use their topic sentence as a guide.
EXTENDING PROMPT:
- Task: write a CS that uses different wording but says the same big idea.
- Extra Notes: praise paraphrasing.

TEACHER NOTES:
Final You Do task. Completes the SPO. Students will write the full body paragraph in lesson 10.

WATCH FOR:
- Students who copy the topic sentence verbatim. Prompt to paraphrase.
- Students who add new facts. Cue: 'No new ideas in the CS.'""",

    51: """SAY:
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
    print(f"L09 written: {stats}")


if __name__ == "__main__":
    main()
