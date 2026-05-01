"""Deep-pass teacher notes for Lesson 3 (Storm Boy - summary sentences) per teachernotes.md v2.0."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from polish_notes_lib import polish_deck

SRC = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated/3. Storm Boy - summary sentences - with notes.pptx")
OUT = SRC.with_name(SRC.stem + " v2.pptx")

NOTES = {
    1: """TEACHER NOTES:
Title slide. Begin once materials are ready and the novel is open at page 12.""",

    2: """DO:
- Print the Session 3 Summary Sentence Worksheet, one per student.
- Print 4 to 6 enabling scaffolds with question prompts pre-filled and frames.
- Keep the answer key with your copy.

TEACHER NOTES:
Teacher setup slide, not student-facing. Resources are referenced again at the You Do.""",

    3: """TEACHER NOTES:
Read aloud or summarise before today's reading.

SENSITIVITY ADVISORY:
- What it is: Storm Boy is set on Ngarrindjeri country and includes Aboriginal characters and language from the period.
- Framing language: 'This story shows respect for the Coorong and the Ngarrindjeri people.'
- Watch for: students affected by names or images of deceased persons.
- Protocol: pause if a student is upset, offer a quiet break with a peer or aide, follow up at recess and with your wellbeing lead if needed.""",

    4: """TEACHER NOTES:
Teacher orientation only, not for students. Read once before delivering the unit.""",

    5: """TEACHER NOTES:
Teacher reference for the I Do, We Do, You Do badges and the support and extension icons. Not student-facing.""",

    6: """TEACHER NOTES:
Teacher reference for the response routines used through the deck: whiteboards, choral, thumbs, show fingers, pair share, cold call. Not student-facing.""",

    7: """TEACHER NOTES:
Teacher reference for the sentence-element colour coding used in the modelling slides. Not student-facing.""",

    8: """SAY:
- "Read the learning intention with me."
- "Today we are squeezing a long passage into one clear sentence."
- "Ask: why is summarising useful? Expected: it helps us remember and retell the main idea quickly."
- "If summarising feels new, that is okay. We will build it together."

DO:
- Choral read the LI, then track each success criterion with your finger.
- Take 2-3 student responses on usefulness.

TEACHER NOTES:
SC2 is the core. SC3 builds directly on the fronted adverbial taught in Lesson 1.

WATCH FOR:
- Students who confuse summarising with retelling. Clarify: 'A summary is the main idea. Retelling is everything.'""",

    9: """SAY:
- "Boards out, novel out, booklet ready."
- "Texta in your hand. Lid checked."

DO:
- Scan the room for missing items before reading begins.
- Pair up any students missing a board or a working texta.

TEACHER NOTES:
Material check. Settle this fast.""",

    10: """TEACHER NOTES:
Section divider. Today's reading mode is Teacher Read Aloud. Have your novel pre-marked at pages 12 to 15.""",

    11: """SAY:
- "Listen carefully and think about what Storm Boy is learning."
- "Listen for: what does Storm Boy learn from living on the land? What makes him happiest?"
- "I will pause to check, so be ready to think."

DO:
- Read pages 12 to 15 aloud, with expression.
- Pause at 'be a little' (p.13) and 'tinted porcelain' (p.14).
- Use Choral Response then Turn and Talk at the pauses.

CFU CHECKPOINT:
Technique: Choral Response
Script:
- Ask: 'What is Storm Boy learning about?' Expected: the land, the animals, the Coorong.
- Scan for: confident class-wide response.
PROCEED:
- >=80% can identify the learning. Continue reading.
PIVOT:
- Most likely: students focus on details and miss the big picture.
- Reteach: 'Storm Boy is learning about the Coorong from Fingerbone and from exploring. He is happy here.'
- Re-check: 'What makes Storm Boy happy?'

TEACHER NOTES:
Choral Response is a low-stakes recall check. Many similes appear in this section. Note any students miss the big picture for follow-up at the I Do.

WATCH FOR:
- Students listing individual events. Guide: 'Those are details. What is the one main thing?'""",

    12: """TEACHER NOTES:
Section divider. Today's words are battered, sedately and roused, all from pages 12 to 15.""",

    13: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is battered. Say it with me. Battered."
- "Battered means old and in poor condition because it has been used a lot. It is an adjective."
- "Watch this. Hide-Away Tom wore a battered old Tom Sawyer hat. Battered tells us the hat is worn out."
- "Word family: batter, batters, battering."

DO:
- Point to the image as you give the meaning.
- Choral read the word twice.

TEACHER NOTES:
First meeting with battered. Anchor in the Storm Boy hat example.

WATCH FOR:
- Students confusing battered with batter (cake mix) or batting (hitting). Clarify: 'Same letters, different words.'""",

    14: """SAY:
- "Read with me: 'The old boot is scruffy and battered.'"
- "If this is battered, say 'battered' together."
- "Expected: battered. Worn out and scruffy."

DO:
- Read aloud. Point to the image. Cue choral response.

WATCH FOR:
- Students staying silent. Re-cue: 'On three, all together.'""",

    15: """SAY:
- "Read with me: 'The car is rusty and battered.'"
- "If this is battered, say 'battered' together."
- "Expected: battered. Old and in poor condition."

DO:
- Read aloud. Point to the image. Cue choral response.

WATCH FOR:
- Students hesitating. Anchor: 'Rusty plus old equals battered.'""",

    16: """SAY:
- "Read with me: 'I poured batter into the patty pans.'"
- "If this is battered, say 'battered' together."
- "Expected: silence. Cake batter is a different word."

DO:
- Read aloud. Point to the image.
- After the silence, confirm: 'Cake batter is fresh, not old and worn.'

WATCH FOR:
- Students still saying battered. Slow down: 'Is the cake mix old and worn?'""",

    17: """SAY:
- "Read with me: 'The boy is batting the ball.'"
- "If this is battered, say 'battered' together."
- "Expected: silence. Batting means hitting with a bat."

DO:
- Read aloud. Point to the image.
- After the silence, confirm: 'Batting is hitting. Battered is old and worn.'

WATCH FOR:
- Students confusing the verbs. Quick anchor: 'A batter hits. A battered thing is worn.'""",

    18: """SAY:
- "Read with me: 'The old battered shed was falling apart.'"
- "If this is battered, say 'battered' together."
- "Expected: battered."

DO:
- Read aloud. Point to the image. Cue choral response.

WATCH FOR:
- Students missing the cue. The double cue 'old battered' makes this an easy yes.""",

    19: """SAY:
- "Ask: would or wouldn't a fence be battered after a storm?"
- "On your whiteboard. Show me!"
- "Expected: would. Storms can damage fences."

DO:
- Read aloud. Signal. Scan.

WATCH FOR:
- Students writing wouldn't. Anchor: 'Strong winds and rain leave a fence damaged - that is battered.'""",

    20: """SAY:
- "Ask: would or wouldn't a ball be battered after one soccer game?"
- "Show me!"
- "Expected: wouldn't. One game is not enough."

DO:
- Read aloud. Signal. Scan.
- Quick reason: 'Battered means old and worn out. One game does not do that.'

WATCH FOR:
- Students writing would. Re-anchor: 'Battered takes time. A new ball after one game is not battered.'""",

    21: """SAY:
- "Look at the words. Which mean similar to battered?"
- "On your whiteboard. 1 minute. Show me!"
- "Expected: falling apart, poor condition, old, run down. Not: well-kept, good condition."

DO:
- Read each option once.
- Time 1 minute.
- Scan boards.
- Confirm answers.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Hold up your sort.' Expected: falling apart, poor condition, old, run down on similar.
- Scan for: well-kept and good condition not included.
PROCEED:
- >=80% sort correctly. Move to the next vocab word.
PIVOT:
- Most likely: students include 'old' but skip 'falling apart'.
- Reteach: 'Battered means old and damaged. Both fit.'
- Re-check: 'Is shabby similar to battered?'

WATCH FOR:
- Students who pick well-kept. Quick anchor: 'Battered is the opposite of well-kept.'""",

    22: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is sedately. Say it with me. Sedately."
- "Sedately means calmly and steadily. It is an adverb. It tells us how."
- "Watch this. The fairy penguin sat sedately on her tail. Sedately tells us how she sat."
- "Word family: sedate (adjective), sedately (adverb)."

DO:
- Point to the image as you give the meaning.
- Choral read the word twice with a calm, slow voice so the meaning matches.

TEACHER NOTES:
First meeting with sedately. Anchor in the penguin example.

WATCH FOR:
- Students who think sedately means asleep. Clarify: 'Awake but very calm and steady.'""",

    23: """SAY:
- "Look at A and B. Which person is doing something sedately?"
- "On your whiteboard: A or B. Show me!"
- "Expected: the calm and steady person."

DO:
- Read prompt. Signal. Scan.
- Confirm with one student.

WATCH FOR:
- Students unable to choose. Reframe: 'Which one is calm? Which one is rushed?'""",

    24: """SAY:
- "Look at A and B. Which animal is doing something sedately?"
- "Show me!"
- "Expected: the calm and still animal."

DO:
- Read prompt. Signal. Scan.

WATCH FOR:
- Students picking the active animal. Anchor: 'Sedately is calm. Not running, not jumping.'""",

    25: """SAY:
- "Look at A and B. Which person is doing something sedately?"
- "Show me!"
- "Expected: the calm and steady person."

DO:
- Read prompt. Signal. Scan.
- Quick reason: 'Sedately is calm. Not rushed, not loud.'

WATCH FOR:
- Students who got the previous slides wrong. Pull alongside in pairs work.""",

    26: """SAY:
- "Reading a book on the couch is something done sedately."
- "On your whiteboard: write something you do sedately. 1 minute. Show me!"

DO:
- Time 1 minute.
- Scan boards.
- Cold call 2-3 students to share.

WATCH FOR:
- Students writing energetic activities like running or playing. Redirect: 'Sedately is calm. Try something quieter.'""",

    27: """SAY:
- "Three sentences. One uses sedately incorrectly. Find it."
- "A: cat sat sedately watching birds. B: children ran sedately around the playground. C: man meditated sedately under a tree."
- "Write A, B or C. Show me!"
- "Expected: B. Running and sedately do not match."

DO:
- Read each sentence aloud once.
- Signal. Scan.
- Confirm: 'Running is rushed. Sedately is calm.'

WATCH FOR:
- Students picking A. Quick correction: 'A cat watching birds is calm. The verb sat fits sedately.'""",

    28: """SAY:
- "Two columns: similar to sedately, different from sedately."
- "Sort the four words. 1 minute. Show me!"
- "Expected: similar are calmly and quietly. Different are noisily and excitedly."

DO:
- Read each word.
- Time 1 minute.
- Scan boards.
- Confirm answers.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Hold up your sort.' Expected: calmly and quietly on similar; noisily and excitedly on different.
- Scan for: correct placement.
PROCEED:
- >=80% sort correctly. Move to the next vocab word.
PIVOT:
- Most likely: students put 'quietly' under different.
- Reteach: 'Sedately is calm and quiet. Quiet things can be sedate.'
- Re-check: 'Is gently similar to sedately?'

WATCH FOR:
- Students placing all four in one column. Stop and re-read the column headings together.""",

    29: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is roused. Say it with me. Roused."
- "If something rouses you, it wakes you up. Roused is a verb."
- "Watch this. Hide-Away's voice roused Storm Boy when he was thinking. Roused tells us he was woken from his thoughts."
- "Word family: rouse, rouses, rousing."

DO:
- Point to each part of the slide.
- Choral read the word twice.

TEACHER NOTES:
First meeting with roused. Note the wider meaning: woken from sleep or from thought.

WATCH FOR:
- Students thinking roused means only physical waking. Clarify: 'From sleep or from thoughts or daydreams.'""",

    30: """SAY:
- "True or false: people are usually roused by yelling."
- "Show me!"
- "Expected: true. Yelling is loud."

DO:
- Read aloud. Signal. Scan.

WATCH FOR:
- Students writing F. Anchor: 'Loud noises wake people. Yelling is loud.'""",

    31: """SAY:
- "True or false: people are usually roused by smells."
- "Show me!"
- "Expected: true. Strong smells can wake people."

DO:
- Read aloud. Signal. Scan.

WATCH FOR:
- Students unsure. Anchor: 'A strong smell like burning toast wakes someone up.'""",

    32: """SAY:
- "True or false: people are usually roused by whispering."
- "Show me!"
- "Expected: false. Whispering is too quiet."

DO:
- Read aloud. Signal. Scan.
- Quick reason: 'Roused needs something noticeable.'

WATCH FOR:
- Students writing T. Reframe: 'A whisper is gentle. Roused needs more.'""",

    33: """SAY:
- "True or false: people are usually roused by touching."
- "Show me!"
- "Expected: true. A tap on the shoulder can wake someone."

DO:
- Read aloud. Signal. Scan.

WATCH FOR:
- Students writing F. Anchor: 'A gentle tap is enough to bring someone back.'""",

    34: """SAY:
- "A: the loud crash of thunder roused the household. B: the soup roused on the stove."
- "Which uses roused correctly? Show me!"
- "Expected: A. Roused happens to people or animals."

DO:
- Read both. Signal. Scan.
- Confirm: 'Soup cannot be roused.'

WATCH FOR:
- Students picking B. Re-anchor: 'Roused needs a living thing being woken.'""",

    35: """SAY:
- "A: the aroma of fresh coffee roused from the kitchen. B: the barking dog roused the boy in the night."
- "Which uses roused correctly? Show me!"
- "Expected: B. The barking woke the boy. Aromas drift, not rouse."

DO:
- Read both. Signal. Scan.

WATCH FOR:
- Students picking A. Anchor: 'An aroma drifts. To rouse you need something that wakes a person.'""",

    36: """SAY:
- "Roused or not?"
- "Listening to music on your headphones?"
- "A teacher tapping you on the shoulder when you are daydreaming?"
- "Waking up after loud fireworks?"
- "Thumbs up for roused, thumbs down for not."
- "Expected: headphones down. Tap up. Fireworks up."

DO:
- Read each item.
- Cue thumbs after each.
- Confirm aloud.

CFU CHECKPOINT:
Technique: Thumbs Up Thumbs Down
Script:
- Ask: 'Tap on the shoulder when daydreaming. Thumbs?' Expected: up.
- Scan for: thumbs up across the room.
PROCEED:
- >=80% answer all three correctly. Move to the You Do.
PIVOT:
- Most likely: students mark headphones as roused because music is loud.
- Reteach: 'Roused means being woken from sleep, daydream or stillness. If you chose to listen, you are not being roused.'
- Re-check: 'An alarm clock - roused or not?'

WATCH FOR:
- Students copying neighbours. Cover their hand briefly and re-cue.""",

    37: """SAY:
- "In your booklet: write the meaning of sedately, then the meaning of roused."
- "Note: the slide says 'sedatedly' - this is a typo. Read it as sedately."
- "Use your own words. Add an example sentence if you have time."

DO:
- Set 3 minutes.
- Circulate, check 2-3 students at the start.
- Hand the scaffold to identified students.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: sentence frame: 'Sedately means ___. Roused means ___.'
EXTENDING PROMPT:
- Task: write one sentence using each word in context. Underline the word.

WATCH FOR:
- Students copying earlier slides word for word. Prompt: 'Your own words.'""",

    38: """TEACHER NOTES:
Section divider. The next slides explicitly teach summary sentence writing using the WHO/WHAT/WHEN/WHERE/WHY prompts from Lesson 1.""",

    39: """SAY:
- "A summary is a short version. Most important parts only. No extra detail."
- "Watch this. I am going to summarise a passage from page 14."
- "I ask: WHO is it about? Storm Boy. WHAT DOING? Wandered, finding treasures. WHEN? After a big blow. WHERE? The beach near the Coorong."
- "I am going to put the WHEN at the front: 'After a big blow, Storm Boy wandered along the beach near the Coorong, finding all kinds of treasure.'"
- "Notice: I did not include every detail. I captured the main idea."

DO:
- Display the source passage.
- Work through the prompts visibly.
- Write the summary on the board step by step.
- Compare the long passage to the short summary.

TEACHER NOTES:
Core I Do. The question prompts from Lesson 1 transfer directly. Students should already be familiar with the WHO/WHAT/WHEN/WHERE/WHY routine.

WATCH FOR:
- Students wanting every detail. Redirect: 'In one sentence. What is the main idea?'
- Students writing too short. Push: 'Add detail. When? Where? What kinds of things?'""",

    40: """SAY:
- "Read this passage from page 14 with me."
- "We are going to summarise it in one sentence."

DO:
- Read the passage aloud, slowly.
- Re-read once if students need it.
- Pause before moving on so students can think.

TEACHER NOTES:
Source passage display. The next two slides build the summary.""",

    41: """SAY:
- "Use the question prompts to break this passage into parts."
- "WHO? Show me! Expected: Storm Boy."
- "WHAT DOING? Show me! Expected: liked to wander along the beach."
- "WHEN? Show me! Expected: after a storm."
- "WHERE? Show me! Expected: along the beach."
- "WHY? Show me! Expected: to collect treasures washed in by wind and waves."

DO:
- Run whiteboard checks for each prompt.
- Display the answers in the columns.
- Build a complete picture before assembling the summary.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'WHAT was Storm Boy doing?' Hold up. Expected: wandering along the beach.
- Scan for: students identifying the main action, not listing every detail.
PROCEED:
- >=80% identify the main action. Move to building the summary.
PIVOT:
- Most likely: students list every action.
- Reteach: 'What is the one main action? Wandering along the beach. The other things are details.'
- Re-check: 'In three words, what is the passage about?'

WATCH FOR:
- Students copying phrases from the text. Prompt: 'Your own words.'
- Students who cannot identify the main action. Simplify: 'What is the one thing he does?'""",

    42: """SAY:
- "We have our parts. Now we build the summary."
- "Starting with WHEN at the front: 'After a storm,...'"
- "Then WHO: '...Storm Boy...'"
- "Then WHAT DOING: '...liked to wander...'"
- "Then WHERE: '...along the beach...'"
- "Then WHY: '...so he could collect treasures.'"
- "Notice the comma after the fronted adverbial."

DO:
- Build the summary on the board, one piece at a time.
- Use the column data from the previous slide.
- Read the final sentence aloud once.
- Compare to the original: 'Did we capture the main idea?'

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: scaffold with multiple-choice question prompts. Circle the right detail and use the frame.
- Extra Notes: pair with a confident partner.
EXTENDING PROMPT:
- Task: after summarising all passages, write a 2-3 sentence summary of pages 5 to 15. Each sentence covers a different section.

WATCH FOR:
- Students forgetting the comma after the fronted adverbial. Point to the I Do model.
- Students writing multi-sentence answers. Refocus: 'One sentence.'""",

    43: """SAY:
- "First: read each passage on your worksheet carefully."
- "Next: answer the who, what doing, when and where questions for each passage."
- "Then: write your summary, starting with the when. Don't forget the comma."
- "One sentence per passage."

DO:
- Distribute the worksheet.
- Set 12 minutes.
- Circulate. Check that students answer the prompts before writing the summary.
- Hand the scaffold to identified students.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: scaffold worksheet with multiple-choice prompts and frames. Aim for 2 of 3 passages.
- Extra Notes: small group at the start, then release.
EXTENDING PROMPT:
- Task: after all passages, write a 2-3 sentence summary of pages 5 to 15. Identify your favourite literary device and explain why in one sentence.

TEACHER NOTES:
Watch the comma rule and the use of text-based detail. Sentences with no link to the passage mean students need a quick re-anchor.

WATCH FOR:
- Students skipping the prompts and going straight to writing. Redirect: 'Answer the questions first.'
- Students writing multi-sentence summaries. Refocus: 'One sentence. Squeeze it down.'
- Readiness signal: 2 of 3 summaries with correct fronted adverbial and comma.""",

    44: """TEACHER NOTES:
Use this slide only if you are not printing the worksheet. Otherwise students work from their printed sheet.""",

    45: """TEACHER NOTES:
Use this slide only if you are not printing the worksheet. Otherwise students work from their printed sheet.""",

    46: """TEACHER NOTES:
Use this slide only if you are not printing the worksheet. Otherwise students work from their printed sheet.""",

    47: """SAY:
- "Read each I can statement with me."
- "SC1: I can identify who, what, when, where details. Thumbs."
- "SC2: I can write a summary that captures the main idea. Thumbs."
- "SC3: I can begin with a when adverbial and include precise details. Thumbs."
- "Turn and Talk: read your best summary to your partner. Partner: thumbs up or down. Did it capture the main idea? 30 seconds."

DO:
- Run a thumbs check after each SC.
- Time the Turn and Talk strictly.
- Cold call 1-2 pairs to share.

TEACHER NOTES:
Partner feedback as a quick formative check. Lessons 4 and 5 shift to non-fiction writing.

WATCH FOR:
- Students thumbs down on SC2. Plan small-group support at the start of Lesson 4.
- Partners giving thumbs up too easily. Listen in to a few pairs.""",

    48: """SAY:
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
    print(f"L03 written: {stats}")


if __name__ == "__main__":
    main()
