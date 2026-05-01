"""Deep-pass teacher notes for Lesson 12 (SPO to summarise a text) per teachernotes.md v2.0."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from polish_notes_lib import polish_deck

SRC = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated/12. literature_presentation Single paragraph outline (SPO) to summarise a text 4En06V Storm Boy - Information report - with notes.pptx")
OUT = SRC.with_name(SRC.stem + " v2.pptx")

NOTES = {
    1: """TEACHER NOTES:
Title slide. Today's focus: read pages 49 to 54 and learn to use a Single Paragraph Outline (SPO) to summarise the main ideas. Begin once materials are out and the novel is open at page 49.""",

    6: """SAY:
- "These are our learning objectives across the unit."
- "Today we focus on the last one: summarising the main ideas using a Single Paragraph Outline."
- "Read the objectives aloud with me."

DO:
- Choral read each objective.
- Highlight the SPO objective verbally.

WATCH FOR:
- Students new to SPO. Reassure: 'We will build it step by step.'""",

    7: """SAY:
- "Read the learning intention with me."
- "Today we read pages 49 to 54 of Storm Boy and learn to use a Single Paragraph Outline (SPO) to summarise main ideas."
- "Ask: what does an SPO have? Expected: topic sentence, supporting details, concluding sentence."
- "If SPO feels new, that is okay. We will build it together."

DO:
- Choral read the LI, then track each success criterion with your finger.
- Cold call one student to repeat SC3 in their own words.

TEACHER NOTES:
SC2 and SC3 are the writing targets the You Do tasks assess.

WATCH FOR:
- Students who cannot recall what an SPO is. Reassure: 'We will model it together.'""",

    8: """SAY:
- "Get your mini whiteboard, texta, booklet, novel and pencil ready."
- "Pencil cases on the floor."

DO:
- Scan the room. Wait until every student has the four items.
- Distribute spares for any student missing materials.

TEACHER NOTES:
Materials check.""",

    10: """SAY:
- "Today we are reading pages 49 to 54 of Storm Boy."
- "We will pause at the spots I have marked to check our thinking."
- "Find page 49."

DO:
- Give 30 seconds for students to find page 49.
- Read aloud with expression.
- Use the pause points and queries from the Literature Study Guide.
- Discuss any literary devices you have selected.

TEACHER NOTES:
Pre-mark your novel before teaching. Select 3 to 4 pause points and 1 to 2 literary devices.

WATCH FOR:
- Students losing place during pauses. Check fingers on the line before continuing.
- Students who notice the close bond between Storm Boy and Mr Percival. Excellent comprehension.""",

    12: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is perched."
- "Perched means to settle or rest on a perch or something similar. It is a verb."
- "Say the word with me: perched."
- "In the story, the pelican found a good spot nearby and perched there heavily to watch and wait."

DO:
- Point to the picture as you give the meaning.
- Choral repeat the word twice.
- Gesture: hands together as if balancing on a branch.

TEACHER NOTES:
First vocabulary word. Picture supports meaning - bird settled on a branch.

WATCH FOR:
- Students who think 'perched' just means 'sat'. Press: 'It is a settled, balanced rest, often up high.'""",

    13: """SAY:
- "Read the sentence with me: 'The bird found a comfortable perch on the highest branch, looking at the world below.'"
- "Ask: where is the bird perched? Expected: on the highest branch."

DO:
- Choral read.
- Cold call one student for the meaning of perch in this sentence.

WATCH FOR:
- Students who say 'fly' or 'rest'. Clarify: 'Perched is sitting, balanced, often elevated.'""",

    14: """SAY:
- "Look at picture A and picture B. Which bird is perched?"
- "Show me A or B with your fingers."

DO:
- Wait time. On cue all show together.
- Reveal the answer and justify.

TEACHER NOTES:
Image discrimination task. The perched bird is the one balanced on something elevated.

WATCH FOR:
- Students confusing 'standing' with 'perched'. Press: 'Perched is on an elevated point, balanced.'""",

    15: """SAY:
- "Look at the two lizards. Which one is perched?"
- "Show me A or B with your fingers."

DO:
- Wait time. On cue all show together.
- Reveal the answer and justify.

TEACHER NOTES:
Second image discrimination. Applies the word to a different animal so students transfer the meaning beyond birds.

WATCH FOR:
- Students who think only birds can perch. Confirm: 'Any animal can perch on something elevated.'""",

    16: """SAY:
- "Read both sentences. Which uses 'perched' correctly?"
- "Sentence A: 'The eagle perched majestically on the cliff's edge.'"
- "Sentence B: 'The cat perched the toy mouse on the floor before pouncing on it.'"
- "Show fingers A or B."

DO:
- Wait time. On cue all show together.
- Cold call one student to justify.

TEACHER NOTES:
Answer: A. Perched describes the action of the subject settling itself, not placing an object.

WATCH FOR:
- Students who choose B because it sounds active. Reteach: 'Perched is what a creature does to itself, not to an object.'""",

    17: """SAY:
- "Build a sentence using the word perched."
- "Frame: 'The mischievous cat was perched on a tree branch, so that it could [blank].'"
- "One example: 'so that it could pounce on the nearby bird.'"
- "On your whiteboards, finish your own version."

DO:
- Whiteboards out. Give 90 seconds.
- Circulate. Read out 4 to 5 examples.

WATCH FOR:
- Students using 'perched' for ground-level rest. Press: 'Think of being elevated and balanced.'""",

    18: """SAY:
- "Look at the picture. Write a sentence about it using the word perched."
- "Example: 'The monarch butterfly perched on the edge of the flower, gazing out at the field.'"

DO:
- Whiteboards out. Give 2 minutes.
- Click to reveal the example.
- Cold call two students to share their sentence.

CFU CHECKPOINT:
Technique: Whiteboard scan
Script:
- Ask: 'Hold up your whiteboard. Sentence about the picture using perched.' Expected: subject + perched + location.
- Scan for: correct use of perched.
PROCEED:
- >=80% correctly use perched. Move to the next vocabulary word.
PIVOT:
- Most likely: students confuse perched with sat or stood.
- Reteach: model two more examples - 'The bird perched on the gate. The cat perched on the windowsill.'
- Re-check: 'Try again with the picture. One sentence using perched.'

WATCH FOR:
- Students who write a sentence without using perched at all. Redirect.""",

    19: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is joyously."
- "Joyously means in a manner that shows you are extremely happy."
- "Say the word with me: joyously."
- "Joyously is an adverb. It tells us how something is done."
- "In the story, the pelican cruised joyously round him."

DO:
- Point to the picture as you give the meaning.
- Choral repeat the word twice.
- Gesture: arms up in celebration.

TEACHER NOTES:
Second vocabulary word. An adverb. Tells how something is done.

WATCH FOR:
- Students who say 'joyous' instead of 'joyously'. Anchor: 'The -ly ending makes it an adverb.'""",

    20: """SAY:
- "Read the sentence with me: 'The children danced joyously in the rain.'"
- "Ask: how did the children dance? Expected: joyously / very happily."

DO:
- Choral read.
- Quick partner share: another time you might do something joyously.

WATCH FOR:
- Students who think joyously is just 'happily'. Press: 'It is more intense, more emphatic.'""",

    21: """SAY:
- "Read both sentences. What is the effect of changing 'happy' to 'joyful'?"
- "Sentence A: 'The children were happy as they ran through the park...'"
- "Sentence B: 'The children were joyful as they ran through the park...'"
- "Turn and tell your partner what changes."

DO:
- Partner talk for 60 seconds.
- Cold call two students.
- Reveal: joyful conveys deeper, more intense happiness; creates a more vivid image.

TEACHER NOTES:
Word-choice analysis. Joyful is more intense and infectious than happy. Stress this for descriptive writing transfer.

WATCH FOR:
- Students who say the meaning is the same. Press: 'Which one paints a stronger picture in your mind?'""",

    22: """SAY:
- "Think of a time you reacted joyously."
- "Watch me first: I was joyous when I got to leave early on a Friday."
- "Now silently act out your joyous reaction for the class."

DO:
- Model an example with facial expression and body language.
- Give 30 seconds think time.
- Invite 3 volunteers to act out.
- Discuss: face, voice tone, body language as evidence of joyous.

TEACHER NOTES:
Kinaesthetic vocabulary practice. Students embody the meaning. Helps storage in long-term memory.

WATCH FOR:
- Students who confuse joyous with calm or content. Press: 'Joyous is intense and obvious.'""",

    23: """SAY:
- "Which words have a similar meaning to joyous?"
- "Thumbs up for similar, thumbs down for different."
- "calm... cheerful... happy... content."

DO:
- Run each word one at a time.
- Wait time on each, then on cue all show.
- Reveal: cheerful (yes), happy (yes), content (yes), calm (no - too quiet).

TEACHER NOTES:
Synonym sort using thumbs. 'Calm' is the weakest match. Press the difference: 'Joyous is loud, calm is quiet.'

WATCH FOR:
- Students who put 'calm' as similar. Reteach: 'Joyous is showing it loudly, calm is quiet inside.'""",

    24: """SAY:
- "When you do something joyously, you..."
- "Options: rush quickly, show extremely upset, show extremely happy, show no emotion."
- "Show fingers 1 to 4."

DO:
- Wait time. On cue all show together.
- Cold call one student to justify.

CFU CHECKPOINT:
Technique: Show Fingers 1-4
Script:
- Ask: 'Show me which option matches joyously.' Expected: 3 (extremely happy).
- Scan for: 3 across the room.
PROCEED:
- >=80% pick 3. Move on.
PIVOT:
- Most likely: students confuse joyously with another emotion.
- Reteach: 'Joyously comes from joy. What is joy? Big happiness.'
- Re-check: 'Show me again.'

WATCH FOR:
- Students who copy. Cover the hand and re-cue.""",

    25: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is waddling."
- "Waddle means to walk like a duck, taking short steps and rolling from side to side. It is a verb."
- "Say the word with me: waddling."
- "In the story, the pelican waddled along at his heels."

DO:
- Point to the picture as you give the meaning.
- Choral repeat the word twice.
- Stand up and demonstrate a brief waddle - short steps, body sway.

TEACHER NOTES:
Third vocabulary word. A physical, observable action. Good chance to model with body movement.

WATCH FOR:
- Students who think waddling means 'walking slowly'. Press: 'Short steps and a rolling sway.'""",

    26: """SAY:
- "Read the sentence with me: 'The penguin waddled clumsily across the icy ground, its squishy body swaying from side to side.'"
- "Ask: how did the penguin walk? Expected: waddled / clumsily."

DO:
- Choral read.
- Make the swaying motion with hands at the same time.

WATCH FOR:
- Students who do not connect 'waddle' with 'sway from side to side'. Re-emphasise the body motion.""",

    27: """SAY:
- "Think of a time you have seen an animal waddling."
- "Example: I have seen geese at the lake waddling along the shore."
- "Now your turn. Tell your partner an example."

DO:
- Partner talk for 30 seconds.
- Cold call three students to share.
- Listen for waddling animals: ducks, geese, penguins, toddlers, ducklings.

WATCH FOR:
- Students who name an animal that does not waddle. Press: 'Which animals walk side to side?'""",

    28: """SAY:
- "Look at the two pictures. Which animal is waddling?"
- "Show me A or B with your fingers."

DO:
- Wait time. On cue all show together.
- Reveal the answer and justify.

WATCH FOR:
- Students unsure. Cue: 'Which one would walk side to side?'""",

    29: """SAY:
- "Which sentence uses waddled correctly?"
- "Sentence A: 'The horse waddled gracefully across the field, its movements smooth and effortless.'"
- "Sentence B: 'The ducks waddled after their mother in a straight line, their tiny legs struggling to keep up.'"
- "Show fingers A or B."

DO:
- Wait time. On cue all show together.
- Cold call to justify.

TEACHER NOTES:
Answer: B. Horses do not waddle. Ducks do. The descriptor 'gracefully' is the giveaway in A.

WATCH FOR:
- Students who choose A. Reteach: 'Waddling is short, side-to-side, and a bit clumsy.'""",

    30: """SAY:
- "Which word fits best with waddle?"
- "Options: wiggles, run, slide, glide."
- "Show fingers 1 to 4."

DO:
- Wait time. On cue all show together.
- Cold call one student to justify.

CFU CHECKPOINT:
Technique: Show Fingers 1-4
Script:
- Ask: 'Show me which word fits best with waddle.' Expected: 1 (wiggles).
- Scan for: 1 on most boards.
PROCEED:
- >=80% pick 1. Move to the booklet task.
PIVOT:
- Most likely: students link waddle with general movement, not the side-to-side sway.
- Reteach: 'Wiggle has the same side-to-side movement. Run, slide and glide are smooth, not swaying.'
- Re-check: 'Show me again.'

WATCH FOR:
- Students unsure. Demonstrate the wiggle and waddle side by side.""",

    31: """SAY:
- "Open your booklet to Lesson 12 Vocabulary."
- "Complete the booklet tasks for perched, joyously and waddling."
- "First: read each prompt. Next: write your sentence. Then: check you used the word correctly."

DO:
- Direct to the booklet page.
- Set timer for 8 minutes.
- Circulate. Check first response of 3 to 5 students.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: provide sentence frames for each word: 'The [animal] [word] across the [place].'
EXTENDING PROMPT:
- Task: write one sentence that uses two of the three words together.

WATCH FOR:
- Students who use the word in the wrong word class. Quick correction: 'Which word class is this?'""",

    33: """SAY:
- "Why do we summarise?"
- "A summary is a brief statement. It does not include extra details, just the main ideas."
- "Summaries help us boost comprehension, focus on the main idea, retain information and analyse information."
- "Watch this first. I will show you what summarising looks like in the next slides."

DO:
- Read the four reasons aloud.
- Cold call two students: 'Why might a summary help when you study?'

TEACHER NOTES:
I Do for summarising. Establishes the purpose before the SPO scaffold is introduced.

WATCH FOR:
- Students who think summarising means 'shorter copy'. Press: 'It is the main ideas in your own words.'""",

    34: """SAY:
- "Let's summarise the story plot so far."
- "Use these question prompts: who, what doing, when, where, why, how."
- "Who are the main characters? What does Mr Percival do after he is set free? Where is Storm Boy? How does Hide-Away feel? Why does Storm Boy care about Mr Percival?"

DO:
- Class discussion. Take 1 to 2 contributions per question.
- Record key ideas on the board.

TEACHER NOTES:
We Do plot summary using question prompts. Use this to identify main ideas before introducing the SPO format.

WATCH FOR:
- Students who give plot details rather than main ideas. Cue: 'What is the big idea, not every event.'""",

    35: """SAY:
- "A Single Paragraph Outline (SPO) puts your writing in an order that makes it easy for readers to understand."
- "We will use an SPO to summarise key ideas in Storm Boy."
- "Read along with me."

DO:
- Choral read both sentences.
- Hold up the SPO graphic if students used one in earlier lessons.

TEACHER NOTES:
Introduces the SPO concept. The SPO scaffold reduces the cognitive load of paragraph writing by separating planning from drafting.

WATCH FOR:
- Students who confuse SPO with a full paragraph. Confirm: 'SPO is the plan, not the writing.'""",

    36: """SAY:
- "The parts of a paragraph. Read with me."
- "Topic sentence: expresses the paragraph's main idea."
- "Supporting details: provide details that support the topic sentence."
- "Concluding sentence: summarises or concludes, echoing the topic sentence without repeating it."

DO:
- Choral read each definition.
- Hold up three fingers. We have three parts.

TEACHER NOTES:
We Do for the SPO parts. Foundation for the modelling that follows.

WATCH FOR:
- Students who try to repeat the topic sentence as the conclusion. Press: 'Echo the idea. Do not repeat the words.'""",

    37: """SAY:
- "Watch me find supporting details."
- "Topic sentence: 'Mr Percival, the pelican he nursed back to health, is Storm Boy's best friend.'"
- "I am going to read the next slide and pull out three short notes that support this idea."
- "Notice my notes are in note form, not full sentences."

DO:
- Display the SPO frame on the board or use the slide.
- Walk through reading and underlining the relevant detail in the next slide.

TEACHER NOTES:
I Do introduction to the SPO frame and supporting details. Note form, not full sentences.

WATCH FOR:
- Students who write full sentences in supporting details. Press: 'Short notes with arrows and abbreviations.'""",

    38: """SAY:
- "Watch me think aloud."
- "Read this extract from page 49 about Mr Percival following Storm Boy everywhere."
- "I need to find a detail that shows their close relationship."
- "I notice 'wherever Storm Boy went, Mr Percival followed'. That tells me they were inseparable, so my note is 'where he went, he followed -> they were inseparable.'"

DO:
- Read the extract aloud with expression.
- Click to reveal the detail.
- Underline the relevant words on the slide.

TEACHER NOTES:
I Do think-aloud for finding supporting details from a text extract. Source: Storm Boy by Colin Thiele, p.49.

WATCH FOR:
- Students who write full sentences. Press: 'Short note form.'""",

    39: """SAY:
- "Together this time."
- "Read the next extract about Hide-Away and Fingerbone."
- "What detail shows their relationship is good? What does Mr Percival do for fun with Storm Boy?"
- "Click to reveal."

DO:
- Read the extract aloud.
- Pause for partner talk for 30 seconds.
- Click to reveal: 'glad he found Mr Percival -> they saw how good they were together' and 'fun on the beach -> they played games and enjoyed each other's company.'

TEACHER NOTES:
We Do for finding supporting details. Source: Storm Boy by Colin Thiele, p.50.

WATCH FOR:
- Students who copy chunks of the extract. Press: 'Turn the words into a short note.'""",

    40: """SAY:
- "Hinge question. Which option contains supporting details that match the topic sentence?"
- "Topic sentence: 'Mr Percival, the pelican he nursed back to health, is Storm Boy's best friend.'"
- "Option A: Storm Boy went -> Mr P follow / FB + H-A = glad / Storm Boy fun @ beach w/ Mr P."
- "Option B: Storm Boy collect shells / Mr P = better than watchdog / Mr P fetch pebble."
- "Show fingers A or B."

DO:
- Wait time. On cue all show together.
- Cold call one student to justify.

CFU CHECKPOINT:
Technique: Show Fingers A or B
Script:
- Ask: 'Show me which option supports the topic sentence.' Expected: A.
- Scan for: A on most boards.
PROCEED:
- >=80% show A. Move to the model task.
PIVOT:
- Most likely: students show B. Misconception: choosing details from the same text without checking they support the main idea.
- Reteach: 'Option B has true facts but they do not all show best friendship. Option A is all about the bond.'
- Re-check: 'Show me again.'

TEACHER NOTES:
Key hinge for supporting details vs unrelated facts. Answer: A.""",

    41: """SAY:
- "Now I will model adding the supporting details to the SPO."
- "I write each detail in note form on a separate line."
- "Detail 1: Storm Boy went -> Mr P follow."
- "Detail 2: FB + H-A = glad Storm Boy has Mr P."
- "Detail 3: Storm Boy / fun @ beach w/ Mr P."

DO:
- Display the SPO frame.
- Add each detail one at a time.
- Choral read the SPO so far.

TEACHER NOTES:
I Do continued. Completes the supporting details on the SPO. Stress the note form, not full sentences.

WATCH FOR:
- Students who try to write the full version on their booklet. Press: 'Stay in note form for the SPO.'""",

    42: """SAY:
- "Now I add a concluding sentence."
- "It should sum up the main idea without copying the topic sentence."
- "My example: 'Although it was unusual, Mr Percival and Storm Boy were such close friends that they had become inseparable.'"
- "Notice it echoes 'best friend' but uses different words like 'close friends' and 'inseparable'."

DO:
- Add the CS to the SPO.
- Choral read the whole SPO from TS to CS.
- Cold call two students for an alternative concluding sentence.

TEACHER NOTES:
I Do for the concluding sentence. CS = summarises or concludes, echoing the main idea without repeating it.

WATCH FOR:
- Students who copy the topic sentence. Press: 'Change the words but keep the idea.'""",

    43: """SAY:
- "Hinge question. When summarising a main idea you should..."
- "Option 1: include as many details as you can."
- "Option 2: only list your opinions."
- "Option 3: only include the main ideas in a brief statement."
- "Option 4: include the main ideas but add extra information for the reader."
- "Show fingers 1 to 4."

DO:
- Wait time. On cue all show together.
- Cold call one student to justify.

CFU CHECKPOINT:
Technique: Show Fingers 1-4
Script:
- Ask: 'Show me what summarising means.' Expected: 3.
- Scan for: 3 on most boards.
PROCEED:
- >=80% show 3. Move to the next We Do.
PIVOT:
- Most likely: students think summary needs all details, includes opinions, or extras.
- Reteach: 'A summary is a brief statement. Brief means short. Just the main ideas.'
- Re-check: 'Show me again.'

TEACHER NOTES:
Key hinge for summarising. Answer: 3.""",

    44: """SAY:
- "Now we use the same process on a different idea."
- "Main idea: how Storm Boy got his name."
- "Watch the extract on the next slide and find the supporting details."

DO:
- Hold up three fingers. We will find three supporting details.

TEACHER NOTES:
I Do bridging slide. Sets up the next We Do extract.""",

    45: """SAY:
- "Read the extract from pages 8 and 9 of Storm Boy."
- "Look for key details that show how Storm Boy got his name."
- "I will read first. Listen for the campers and what happens with the boy."

DO:
- Read the extract aloud.
- Click to reveal.
- Highlight 'wandering down the beach all alone' and 'must be lost'.

TEACHER NOTES:
We Do extract one for the 'how Storm Boy got his name' SPO. Source: Storm Boy by Colin Thiele, pp.8-9.

WATCH FOR:
- Students lost in the long extract. Cue: 'Only listen for what happens to the boy.'""",

    46: """SAY:
- "Read the extract from page 9."
- "Look for the moment Storm Boy gets his name."
- "Listen for what the postmaster says."

DO:
- Read the extract aloud.
- Click to reveal.
- Highlight 'boy in the storm' and 'they called him Storm Boy'.

TEACHER NOTES:
We Do extract two for the 'how Storm Boy got his name' SPO. Source: Storm Boy by Colin Thiele, p.9.

WATCH FOR:
- Students who miss the postmaster's role. Press: 'Who actually gives him the name?'""",

    47: """SAY:
- "Hinge question. Which option contains supporting details that best match the main idea: 'how Storm Boy got his name'?"
- "Option A: shells / molly-hawk / postmaster smiled."
- "Option B: wandering beach alone / disappeared, ran to town / he is 'boy in storm' -> Storm Boy."
- "Show fingers A or B."

DO:
- Wait time. On cue all show together.
- Cold call one student to justify.

CFU CHECKPOINT:
Technique: Show Fingers A or B
Script:
- Ask: 'Show me which option supports the main idea.' Expected: B.
- Scan for: B on most boards.
PROCEED:
- >=80% show B. Move to the SPO building task.
PIVOT:
- Most likely: students show A. Misconception: picking details from the extract that are present but not on-topic.
- Reteach: 'Option A has details but they do not explain how he got his name. Option B is all about the name.'
- Re-check: 'Show me again.'

TEACHER NOTES:
Hinge for matching supporting details to the main idea. Answer: B.""",

    48: """SAY:
- "Now together. List the supporting details in note form on the SPO."
- "Detail 1: wandering beach alone / campers worried."
- "Detail 2: he disappeared -> ran to town to get help."
- "Detail 3: he is 'boy in storm' -> then on Storm Boy."

DO:
- Whiteboards or board.
- Choral read each detail as it goes on the SPO.

TEACHER NOTES:
We Do for adding supporting details. Confirm short note form.

WATCH FOR:
- Students writing full sentences. Press: 'Short notes only.'""",

    49: """SAY:
- "Now we add a topic sentence to match the supporting details."
- "Idea: 'Hide-Away's little chap, Storm Boy, was nicknamed by curious campers.'"
- "Cold call: what main idea do all three details support?"

DO:
- Class discussion. Collect 2 to 3 ideas before revealing.
- Add the suggested topic sentence to the SPO.

TEACHER NOTES:
We Do for writing the topic sentence after the supporting details (reverse engineering).

WATCH FOR:
- Students who write a topic sentence about the storm rather than the name. Press: 'Name is the focus.'""",

    50: """SAY:
- "Now add the concluding sentence to the SPO."
- "Suggested CS: 'Storm Boy, who was calm and happy in the thunderous storm, was only ever known as Storm Boy after that day.'"
- "Notice the CS echoes the topic sentence without repeating it."

DO:
- Add the CS to the SPO.
- Choral read the entire SPO from TS to CS.

TEACHER NOTES:
We Do for the concluding sentence. Closes the second SPO.

WATCH FOR:
- Students who write a CS that introduces a new idea. Press: 'Just sum up.'""",

    51: """SAY:
- "Hinge question. Find the supporting detail that does not fit."
- "Main idea: 'Hide-Away and Storm Boy are connected to their home and the nature that surrounds them.'"
- "Topic sentence: 'Hide-Away and Storm Boy live near the sanctuary in the Coorong.'"
- "Supporting details: 1) Away from the outside world = joyous, 2) Humpy in sand dunes -> wood + iron, 3) Midden = campsite / shells + ancient site."
- "Show fingers 1, 2 or 3."

DO:
- Wait time. On cue all show together.
- Cold call to justify.

CFU CHECKPOINT:
Technique: Show Fingers 1-3
Script:
- Ask: 'Show me which detail does not fit the main idea.' Expected: 3.
- Scan for: 3 on most boards.
PROCEED:
- >=80% show 3. Move to the You Do.
PIVOT:
- Most likely: students show 1 or 2. Misconception: confusing 'in the Coorong' details with 'connection to nature' details.
- Reteach: 'Detail 3 is about middens as ancient campsites. That is a fact about the place but not about Hide-Away and Storm Boy's connection to the nature around them.'
- Re-check: 'Show me again.'

TEACHER NOTES:
Hinge for off-topic supporting details. Answer: 3 (middens as campsites disrupts the thematic coherence).""",

    52: """SAY:
- "Independent task. Complete the SPO on Mr Percival's intelligence."
- "Main idea: Mr Percival's intelligence."
- "Topic sentence already given: 'Mr Percival, Storm Boy's closest friend, was a very intelligent bird.'"
- "Concluding sentence already given: 'Mr Percival was a quick learner and used his intelligence to help Hide-Away and Storm Boy.'"
- "Your job: write 2 supporting details that show Mr Percival's intelligence."

DO:
- Whiteboards or booklet.
- Set timer for 4 minutes.
- Circulate. Check first detail of 3 students.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: provide a detail bank: 'fetched objects', 'understood routines', 'helped with fishing'. Students choose two.
EXTENDING PROMPT:
- Task: write three supporting details, then rewrite as full sentences with adjectives.

TEACHER NOTES:
Scaffolded You Do. Use this for students who need extra scaffolding before the booklet task.

WATCH FOR:
- Students who write full sentences. Press: 'Note form.'
- Students who write details that are not about intelligence. Press: 'Only include details that show smart behaviour.'""",

    53: """SAY:
- "Open your booklet to Lesson 12 Sentence-level writing."
- "Complete the tasks on using a Single Paragraph Outline to summarise a text."
- "First: identify the main idea. Next: find 3 supporting details in note form. Then: write a topic sentence and concluding sentence."

DO:
- Direct to the booklet.
- Set timer for 12 to 15 minutes.
- Circulate. Check first response of 3 to 5 students.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: use the scaffolded SPO from slide 51 first.
EXTENDING PROMPT:
- Task: write a second SPO on a different theme from the novel.

TEACHER NOTES:
You Do task. SC2 exit-ticket-style task. Students apply the SPO format to summarise a section of Storm Boy.

WATCH FOR:
- Students writing full sentences in supporting details. Press: 'Short notes.'
- Students whose CS repeats the TS word-for-word. Press: 'Echo the idea, change the words.'
- Fast finishers. Prompt: 'Try the extending task.'""",

    55: """SAY:
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
    print(f"L12 written: {stats}")


if __name__ == "__main__":
    main()
