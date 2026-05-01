"""Deep-pass teacher notes for Lesson 2 (Storm Boy - because, but, so) per teachernotes.md v2.0."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from polish_notes_lib import polish_deck

SRC = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated/2. Storm Boy - because, but, so - with notes.pptx")
OUT = SRC.with_name(SRC.stem + " v2.pptx")

NOTES = {
    1: """TEACHER NOTES:
Title slide. Begin once materials are ready and the novel is open at page 9.""",

    2: """DO:
- Print the worksheet, one per student.
- Print 4 to 6 enabling scaffolds for students who need conjunction cards and pre-filled stems.
- Keep the answer key with your copy, not in the student stack.

TEACHER NOTES:
Teacher setup slide, not student-facing. Resources are referenced again at the You Do.""",

    3: """TEACHER NOTES:
Read aloud or summarise before today's reading. Lesson 2 introduces Fingerbone Bill, a Ngarrindjeri man.

SENSITIVITY ADVISORY:
- What it is: today's pages introduce Fingerbone Bill, a Ngarrindjeri man who lives near Storm Boy.
- Framing language: 'Fingerbone is a Ngarrindjeri man who lives near Storm Boy. We use respectful language when we talk about him.'
- Watch for: students affected by names or images of deceased persons, especially Aboriginal and Torres Strait Islander students.
- Protocol: pause if a student is upset, offer a quiet break with a peer or aide, follow up at recess and with your wellbeing lead if needed.""",

    4: """TEACHER NOTES:
Teacher orientation only, not for students. Read once before delivering the unit. The Literature Study Guide names the pause points and queries used through the lesson.""",

    5: """TEACHER NOTES:
Teacher reference for the I Do, We Do, You Do badges and the support and extension icons used through the deck. Not student-facing.""",

    6: """TEACHER NOTES:
Teacher reference for the response routines used through the deck: whiteboards, choral, thumbs, show fingers, pair share, cold call. Not student-facing.""",

    7: """TEACHER NOTES:
Teacher reference for the sentence-element colour coding used in the modelling slides: who, what doing, when, where, why, how. Not student-facing.""",

    8: """SAY:
- "Read the learning intention with me."
- "These are the three things we are practising today."
- "Ask: which one will be on your worksheet? Expected: SC2, completing a sentence with because, but or so."
- "If 'conjunction' feels new, that is okay. We will build it together."

DO:
- Choral read the LI, then track each success criterion with your finger.
- Ask one student to say SC1 in their own words.
- Park 'conjunction' on the board for now. It will be unpacked at the I Do.

TEACHER NOTES:
SC1 is the floor. Almost every student should reach it from today's reading. SC3 is the metalinguistic stretch.

WATCH FOR:
- Students unsure about conjunction. Quick gloss: 'a word that joins two ideas in a sentence.'""",

    9: """SAY:
- "Boards out, novel out, booklet ready."
- "Texta in your hand. Lid checked."

DO:
- Scan the room for missing items before reading begins.
- Pair up any students missing a board or a working texta.

TEACHER NOTES:
Material check. Settle this fast so the lesson can start.

WATCH FOR:
- Dry textas. Swap before the first show-me, not during it.""",

    10: """TEACHER NOTES:
Section divider. Today's reading mode is Teacher Read Aloud. Have your novel pre-marked at pages 9 to 11.""",

    11: """SAY:
- "Listen carefully and picture what is happening."
- "Listen for: who is the new character, and what problem does Hide-Away Tom have?"
- "I will pause to check, so be ready to think."

DO:
- Read pages 9 to 11 aloud, slow and clear.
- Use a different voice for Fingerbone if you are comfortable.
- Pause at the marked points (next slide).

TEACHER NOTES:
Introduces Fingerbone and the lookout post. Board notes from this read feed the conjunction sentence work later in the lesson.

WATCH FOR:
- Students who miss Fingerbone's introduction. Pause and point: 'The new character is Fingerbone.'
- Students confused about why Storm Boy might get lost. Anchor: 'Storm Boy cannot read.'""",

    12: """SAY:
- "Pause one, page 10. Ask: what has the author told us about Fingerbone? Expected: Aboriginal man, lives nearby, wrinkled face, tells stories."
- "Pause two, page 11. Ask: why is Hide-Away Tom worried about Storm Boy getting lost? Expected: Storm Boy cannot read signs."
- "Ask: what does the author want us to know about this place? Expected: it is remote, easy to get lost in, no roads or signs."

DO:
- Pause at each point.
- Cold call for the page 10 question.
- Turn and Talk 30 seconds, then cold call 2 pairs for the page 11 question.
- Record key Fingerbone and lookout ideas on the board.

CFU CHECKPOINT:
Technique: Cold Call
Script:
- Ask: '[Name], why does Hide-Away Tom build the lookout?' Expected: so Storm Boy can find his way home because he cannot read signs.
- Scan for: students linking Storm Boy's literacy to the danger of getting lost.
PROCEED:
- >=80% recall key events. Move to vocabulary.
PIVOT:
- Most likely: students cannot sequence events.
- Reteach: re-read pages 9 to 10 slowly. Listen for who Fingerbone is and what he does.
- Re-check: 'What is one thing we know about Fingerbone now?'

TEACHER NOTES:
Cold Call ensures all students are processing. Board notes feed the conjunction task at the I Do.

WATCH FOR:
- Students confusing Fingerbone with Hide-Away Tom. Clarify: 'Fingerbone is a different person. He lives a mile down the beach.'""",

    13: """TEACHER NOTES:
Section divider. Today's words are scribbly and mysterious, both pulled from pages 9 to 11.""",

    14: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is scribbly. Say it with me. Scribbly."
- "Scribbly means untidy and hard to read or understand. It is an adjective."
- "Watch this. The story talks about scribbly patterns on the sand made by beetles. Scribbly tells us the marks are messy."
- "Word family: scribble, scribbled, scribbling."

DO:
- Point to the image as you give the meaning.
- Choral read the word twice.
- Trace a quick scribbly mark in the air with your finger.

TEACHER NOTES:
First meeting with the word. Anchor it in the Storm Boy example before the practice slides.

WATCH FOR:
- Students treating any drawing as scribbly. Correct: 'Careful drawings are not scribbly. Scribbly is messy.'""",

    15: """SAY:
- "Listen: 'Fingerbone could read the writing on the sandhills and beaches, all the scribbly stories made by birds, mice and bandicoots.'"
- "Ask: what does the author mean by 'scribbly stories'? Expected: the messy footprints and tracks left by animals. Fingerbone reads them like words."

DO:
- Read the sentence aloud twice, slowly.
- Turn and Talk for 30 seconds.
- Cold call 1 to 2 pairs.
- Highlight that 'scribbly stories' is a metaphor for the tracks.

WATCH FOR:
- Students taking 'stories' literally. Clarify: 'The tracks are like writing. They tell a story to someone who can read them.'""",

    16: """SAY:
- "Read with me: 'The toddler's drawing was scribbly.'"
- "If this is scribbly, say 'scribbly' together."
- "Expected: scribbly. Toddler drawings are messy lines."

DO:
- Read aloud. Point to the image. Cue the choral response.

WATCH FOR:
- Students staying silent. Re-cue: 'On three, all together.'""",

    17: """SAY:
- "Read with me: 'The artwork was made of scribbly patterns.'"
- "If this is scribbly, say 'scribbly' together."
- "Expected: scribbly. The patterns are messy on purpose."

DO:
- Read aloud. Point to the image. Cue the choral response.

WATCH FOR:
- Students hesitating because the artwork looks intentional. Clarify: 'The marks themselves are still scribbly, even when the artist meant it.'""",

    18: """SAY:
- "Read with me: 'The scribbly writing on the blackboard was hard to read.'"
- "If this is scribbly, say 'scribbly' together."
- "Expected: scribbly. The writing is untidy and hard to read."

DO:
- Read aloud. Point to the image. Cue the choral response.

WATCH FOR:
- Students unsure why hard to read = scribbly. Re-anchor: 'Scribbly always means untidy.'""",

    19: """SAY:
- "Read with me: 'The insect made scribbly marks on the gum tree.'"
- "If this is scribbly, say 'scribbly' together."
- "Expected: scribbly. The marks are messy lines."

DO:
- Read aloud. Point to the image. Cue the choral response.
- Quick fact: this is a scribbly gum tree, named for these marks.

TEACHER NOTES:
Australian context. Pause briefly if students share that they have seen these trees.

WATCH FOR:
- Students wanting to talk about gum trees at length. Acknowledge and move on to keep pace.""",

    20: """SAY:
- "Read with me: 'The pencil drawing of the pelican was scribbly.'"
- "If this is scribbly, say 'scribbly' together."
- "Expected: scribbly."

DO:
- Read aloud. Point to the image. Cue the choral response.
- Brief link: 'Pelicans matter in Storm Boy. We will meet Mr Percival soon.'

WATCH FOR:
- Students wanting to talk about pelicans. Park the interest for later in the unit.""",

    21: """SAY:
- "Read the example: 'The marks on the paper were scribbly because they were made by a little child.'"
- "On your whiteboard: write your own version. 'The marks on the ___ were scribbly because ___.'"
- "1 minute. Show me!"

DO:
- Read the example.
- Time 1 minute.
- Scan boards.
- Cold call 2 to 3 students to share.

WATCH FOR:
- Reasons that do not match scribbly, e.g. 'because they were beautiful'. Redirect: 'Scribbly is messy. Why would something be messy?'""",

    22: """SAY:
- "Read the example: 'The marks on the beach were scribbly because they were made by birds walking on the sand.'"
- "This connects to Fingerbone reading scribbly marks in the story."
- "On your whiteboard: write your own version about the beach. 1 minute. Show me!"

DO:
- Read the example.
- Time 1 minute.
- Scan boards.
- Cold call to share.

WATCH FOR:
- Students writing 'because birds were there'. Push: 'How did the birds make the marks?' Expected: footprints or tracks.""",

    23: """SAY:
- "Sort each one. Thumbs up for scribbly, thumbs down for not scribbly."
- "Really neat handwriting?"
- "Notes written hastily?"
- "Carefully drawn shapes on a whiteboard?"
- "Messy scratching in the sand?"
- "Expected: NEAT down. HASTY up. CAREFUL down. MESSY up."

DO:
- Read each item one at a time.
- Cue thumbs after each.
- Confirm aloud after each so students hear the boundary.

CFU CHECKPOINT:
Technique: Thumbs Up Thumbs Down
Script:
- Ask: 'Carefully drawn shapes - scribbly or not?' Expected: down (not scribbly).
- Scan for: thumbs down across the room.
PROCEED:
- >=80% sort all four correctly. Move to the next vocab word.
PIVOT:
- Most likely: students treat any drawing as scribbly.
- Reteach: 'Scribbly means untidy. Careful drawings are not scribbly.'
- Re-check: 'A neatly painted picture - scribbly or not?'

WATCH FOR:
- Students changing their thumb after seeing peers. Re-cue with a fresh example.""",

    24: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is mysterious. Say it with me. Mysterious."
- "Mysterious means kept secret, unexplained or unknown. It is an adjective."
- "Watch this. The story talks about marks on the sandhills made by mysterious sliding bellies. Mysterious tells us we do not know what made them."
- "Word family: mystery, mysteriously."

DO:
- Point to the image as you give the meaning.
- Choral read the word twice.
- Lower your voice to match the feeling of the word.

TEACHER NOTES:
First meeting with mysterious. Anchor in the story image before the practice slides.

WATCH FOR:
- Students confusing mysterious with scary. Clarify: 'Mysterious means unknown. It can be exciting or scary, but the key idea is that we do not know.'""",

    25: """SAY:
- "Read with me: 'In the mysterious forest, a hidden path led to a magical waterfall.'"
- "Ask: what is the author telling us about the forest? Expected: it is unknown, full of secrets."

DO:
- Read aloud twice.
- Turn and Talk for 20 seconds.
- Cold call one pair.

WATCH FOR:
- Students saying 'the forest is scary'. Affirm and refine: 'It might be scary, but mysterious means unknown.'""",

    26: """SAY:
- "Read with me: 'Fingerbone could read the marks on the sandhills and beaches made by mysterious sliding bellies in the night.'"
- "Ask: why does the author choose 'mysterious' here? Expected: we do not see what makes the marks. They are unknown until Fingerbone reads them."

DO:
- Read aloud twice.
- Turn and Talk for 30 seconds.
- Cold call one pair.

WATCH FOR:
- Students who do not link 'sliding bellies' to animals. Reteach: 'Sliding bellies in the sand at night - what kind of animal? Snakes, lizards.'""",

    27: """SAY:
- "Mysterious describes something unknown."
- "An abandoned house can feel mysterious because we do not know its story."
- "On your whiteboard: draw or write something that feels mysterious to you. 1 minute. Show me!"

DO:
- Time 1 minute.
- Scan boards.
- Cold call 2 to 3 students to share.

WATCH FOR:
- Students writing something obviously not mysterious, e.g. 'school'. Redirect: 'Is it unknown? Try something where you do not know the answer.'""",

    28: """SAY:
- "Look at the padlocked door. It is mysterious."
- "Use because to explain why."
- "Frame: 'The padlocked door is mysterious because ___.'"
- "On your whiteboard. 1 minute. Show me!"

DO:
- Read the frame and example.
- Time 1 minute.
- Scan boards.
- Cold call 2 to 3 students.

WATCH FOR:
- Students writing description, not a reason. Redirect: 'After because we need a reason. What do we not know?'""",

    29: """SAY:
- "New picture: a forest hut. Use because to explain why it is mysterious."
- "Frame: 'The forest hut is mysterious because ___.'"
- "On your whiteboard. 1 minute. Show me!"

DO:
- Read the frame and example.
- Time 1 minute.
- Scan boards.
- Cold call 2 to 3 students.

WATCH FOR:
- Students repeating the example. Redirect: 'Use a different reason from your own thinking.'""",

    30: """SAY:
- "Look at the words. Which mean similar to mysterious?"
- "On your whiteboard: write the matching words. 1 minute. Show me!"
- "Expected: puzzling, baffling, secret, unknown. Not: obvious, usual."

DO:
- Read each option once.
- Time 1 minute.
- Scan boards.
- Confirm answers aloud.

WATCH FOR:
- Students including 'obvious'. Clarify: 'Obvious is the opposite of mysterious.'""",

    31: """SAY:
- "Something mysterious will be... 1 common, 2 seen often, 3 a little interesting, or 4 unfamiliar and unexplored?"
- "Show me with fingers."
- "Expected: 4. Mysterious means unknown."

DO:
- Read each option.
- Cue finger response.
- Scan, then confirm.

CFU CHECKPOINT:
Technique: Show Fingers 1-4
Script:
- Ask: 'Show me 1, 2, 3 or 4.' Expected: 4.
- Scan for: 4 across the room.
PROCEED:
- >=80% pick 4. Move to the You Do.
PIVOT:
- Most likely: students pick 3 (a little interesting).
- Reteach: 'Mysterious is more than interesting. It means unknown.'
- Re-check: 'A glass of water - mysterious or familiar? A hidden cave - mysterious or familiar?'

WATCH FOR:
- Students copying the student next to them. Cover their hand briefly and re-cue.""",

    32: """SAY:
- "In your booklet: write the meaning of scribbly, then the meaning of mysterious. Use your own words."
- "Add an example sentence if you have time."
- "Three minutes."

DO:
- Set 3 minutes.
- Circulate, check 2 to 3 students at the start of writing.
- Hand the scaffold to identified students before they get stuck.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: sentence frame: 'Scribbly means ___. Mysterious means ___.'
EXTENDING PROMPT:
- Task: write one sentence using each word in context. Underline the word.

WATCH FOR:
- Students copying from earlier slides word for word. Prompt: 'Your own words. Try it without looking.'""",

    33: """TEACHER NOTES:
Section divider. The next slides explicitly teach because, but and so as conjunctions, with the comma rules.""",

    34: """SAY:
- "Three new conjunctions: because, but, so."
- "BECAUSE shows why something happened. It gives a reason."
- "BUT shows a change in direction. Something unexpected."
- "SO shows what happens as a result. Cause and effect."
- "Watch this everyday model: 'I was hungry because I did not eat lunch. I was hungry, but I had no food. I was hungry, so I ate an apple.'"

DO:
- Read each definition.
- Write the everyday model on the board.
- Highlight the comma rule: comma before BUT and SO. No comma before BECAUSE.

WATCH FOR:
- Students confusing BECAUSE and SO. Anchor: 'Because answers why. So answers what happened next.'""",

    35: """SAY:
- "Watch this. My stem: 'Fingerbone loved the Coorong.'"
- "BECAUSE version: 'Fingerbone loved the Coorong because he felt a strong connection to the land.' Because tells me why."
- "BUT version: 'Fingerbone loved the Coorong, but he did not love the tiger snakes who lived there.' But shows a change."
- "SO version: 'Fingerbone loved the Coorong, so he taught Storm Boy how to understand and respect it.' So shows the result."
- "Notice: comma before BUT and SO. No comma before BECAUSE."

DO:
- Display the stem.
- Build each version one at a time, writing on the board if possible.
- Point clearly to comma placement on each.
- Ask: 'Did Fingerbone love the Coorong in all three? Yes. What changed? The second part.'

MISCONCEPTIONS:
- Misconception: students put a comma before because.
  Why: overgeneralisation from but and so.
  Impact: incorrect punctuation in most contexts.
  Quick correction: 'No comma before because. Just before but and so.'

TEACHER NOTES:
Core I Do. The repeated stem makes the meaning shift visible.

WATCH FOR:
- Students saying all three sentences mean the same. Re-read the second clauses with extra emphasis.""",

    36: """SAY:
- "A compound sentence has two parts that could each stand alone, joined with a conjunction."
- "BUT and SO are coordinating conjunctions."
- "Comma before BUT and SO."

DO:
- Read the slide.
- Point to each clause and say 'this could stand alone'.
- Underline the comma and the conjunction in the example.

WATCH FOR:
- Students confused by 'independent clause'. Reframe: 'a part of the sentence that could stand on its own.'""",

    37: """SAY:
- "A complex sentence has two parts, but one of them cannot stand alone."
- "BECAUSE is a subordinating conjunction."
- "No comma before BECAUSE in the middle of a sentence."

DO:
- Read the slide.
- Point to each clause: which can stand alone, which cannot.
- Highlight the no-comma rule.

WATCH FOR:
- Students putting a comma before because. Restate the rule and circle the missing comma on the slide example.""",

    38: """SAY:
- "Read each model with me."
- "BECAUSE he had lived there for a long time."
- "BUT was always willing to learn more."
- "SO he taught Storm Boy how to read the tracks and weather."
- "Notice how each conjunction changes the meaning."

DO:
- Cue partner discussion 30 seconds: what did Fingerbone know about the Coorong?
- Read each model aloud with the class.
- Underline each conjunction and its comma (or lack of one) on the slide.

WATCH FOR:
- Students who do not see the meaning shift. Re-read the second clause with extra emphasis on the new information.""",

    39: """SAY:
- "Try this together. New stem: 'Hide-Away Tom worried about Storm Boy.'"
- "On your whiteboard: finish with BECAUSE. 30 seconds. Show me!"
- "Flip your board. Finish with BUT. 30 seconds. Show me!"
- "Now we build the strongest version on the board."

DO:
- Display the stem.
- Run two whiteboard checks: BECAUSE first, then BUT.
- Take 2 to 3 contributions, build the best version on the board.
- Point to the comma before BUT.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Hide-Away Tom worried about Storm Boy because... finish it.' Hold up. Expected: a clear reason, e.g. because Storm Boy might get lost, because he could not read.
- Scan for: a reason that fits the story.
PROCEED:
- >=80% produce both versions correctly. Move to the next stem.
PIVOT:
- Most likely: students use BECAUSE and SO interchangeably.
- Reteach: 'Because answers why. So answers what happened next.'
- Re-check: 'I was tired because... versus I was tired so... what is the difference?'

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: scaffold conjunction cards. Pick the right card and complete the pre-filled stem.
- Extra Notes: sit with these students for the first stem.
EXTENDING PROMPT:
- Task: write all three versions for two more story stems. Choose the strongest and explain why that conjunction works best.

WATCH FOR:
- Students writing the same content for both conjunctions. Pair for explicit reteaching.
- Students forgetting the comma before BUT. Point to the I Do model.""",

    40: """SAY:
- "New stem: 'Hide-Away needed Fingerbone's help moving the timber for the lookout.'"
- "Try with BECAUSE. 30 seconds. Show me!"
- "Try with SO. 30 seconds. Show me!"
- "Comma before SO. No comma before BECAUSE."

DO:
- Display the stem.
- Run two whiteboard checks.
- Take 2 to 3 contributions.
- Build the strongest version on the board.

WATCH FOR:
- Students writing the same idea for BECAUSE and SO. Reteach: 'Because answers why. So answers what happened next.'""",

    41: """SAY:
- "New stem: 'Storm Boy doesn't go to school.'"
- "Try with BECAUSE. Show me!"
- "Try with BUT. Show me!"
- "The BUT version is interesting. It shows that something unexpected is true. Without school, Storm Boy still learns a lot."

DO:
- Display the stem.
- Run two whiteboard checks.
- Build versions on the board.
- Discuss the BUT version: 'What unexpected thing is true about Storm Boy?'

WATCH FOR:
- Students writing reasons not in the text. Link back to the story: 'What reason does the author give?'""",

    42: """SAY:
- "New stem: 'Fingerbone is the only man who lives anywhere near them.'"
- "All three this time. BECAUSE, then BUT, then SO. 1 minute total. Show me!"
- "Expected: BECAUSE they live in a remote area. BUT he lives nearly a mile down the beach. SO they spend a lot of time with him."

DO:
- Display the stem.
- Set 1 minute total.
- Scan boards.
- Build the best of each on the board.
- Highlight commas: before BUT and SO, none before BECAUSE.

TEACHER NOTES:
Final We Do before the CFU. Three conjunctions in one slide tests readiness for the You Do.

WATCH FOR:
- Students who can do BECAUSE but stall on BUT or SO. Reteach the meaning of the harder one before they try.""",

    43: """SAY:
- "Look at the three sentences about Fingerbone's humpy."
- "Choose the correct conjunction: BECAUSE (no comma) or , BUT or , SO."
- "Write the number and conjunction. 2 minutes. Show me!"
- "Expected: 1. , but. 2. , so. 3. because."

DO:
- Read each sentence aloud once.
- Time 2 minutes.
- Scan boards.
- Confirm answers, modelling each comma rule.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Hold up your three answers.' Expected: 1. , but. 2. , so. 3. because.
- Scan for: all three correct.
PROCEED:
- >=80% answer all three correctly. Move to the You Do.
PIVOT:
- Most likely: students mix SO and BECAUSE.
- Reteach with substitution test: 'Try because in item 1. Does it explain why? No. So the conjunction is but or so.'
- Re-check: 'In sentence 1, does the second part explain why or show a change?'

WATCH FOR:
- Students skipping commas. Quick reminder: 'Comma before BUT and SO. No comma before BECAUSE.'""",

    44: """SAY:
- "First: read each comprehension question and stem on your worksheet."
- "Next: complete each sentence using the conjunction shown. Check your commas."
- "Then: for the challenge questions, choose your own conjunction and explain your choice."

DO:
- Distribute the worksheet.
- Set 10 minutes.
- Circulate, check comma placement and meaning fit on the first item for each student.
- Hand the scaffold to identified students.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: scaffold worksheet with conjunction cards and pre-filled starters. Aim to finish all BECAUSE items first.
- Extra Notes: small group at the start to set up.
EXTENDING PROMPT:
- Task: write a 3 to 4 sentence paragraph about Fingerbone using at least one because, one but and one so. Underline each.

TEACHER NOTES:
Watch comma placement and meaning fit. Sentences with no link to the story mean students need a quick re-anchor in the reading.

WATCH FOR:
- Students using SO and BECAUSE interchangeably. Reteach: because = why, so = what next.
- Students forgetting commas before BUT and SO. Point to the I Do model.
- Readiness signal: 4 of 6 sentences correct.""",

    45: """SAY:
- "Read each I can statement with me."
- "SC1: I can recall key events from pages 9 to 11. Thumbs."
- "SC2: I can complete a sentence using because, but or so. Thumbs."
- "SC3: I can explain how the conjunction changes the meaning. Thumbs."
- "Turn and Talk: read your best sentence to your partner. Your partner names the conjunction and says if the meaning makes sense. 30 seconds."

DO:
- Run a thumbs check after each SC.
- Time the Turn and Talk strictly.
- Cold call 1 to 2 pairs to share.

TEACHER NOTES:
Reciprocal Turn and Talk reinforces SC3. Use thumbs data to plan tomorrow's launch.

WATCH FOR:
- Students thumbs down on SC2. Plan small-group reteach at the start of Lesson 3.
- Partners who cannot identify the conjunction. Flag for explicit follow-up.""",

    46: """TEACHER NOTES:
Credits and attribution slide. Not student-facing. End the lesson on the closing reflection slide rather than this one.""",

    47: """SAY:
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
    print(f"L02 written: {stats}")


if __name__ == "__main__":
    main()
