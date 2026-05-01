"""Deep-pass teacher notes for Lesson 6 (Storm Boy - Sentence expansion) per teachernotes.md v2.0."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from polish_notes_lib import polish_deck

SRC = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated/6. Storm Boy - Expanding sentences - with notes.pptx")
OUT = SRC.with_name(SRC.stem + " v2.pptx")

NOTES = {
    1: """TEACHER NOTES:
Title slide. Reading focus is pages 15 to 19. Writing focus is sentence expansion using who, when, where, how and why. Begin once materials are out and the novel is open at page 15.""",

    2: """TEACHER NOTES:
Weekly overview. Teacher reference for the week's printing and routines. Not student-facing.""",

    3: """TEACHER NOTES:
Read aloud or summarise before reading from the novel.

SENSITIVITY ADVISORY:
- What it is: the novel and surrounding resources reference First Nations characters from the period.
- Framing language: 'We talk about First Nations characters with respect.'
- Watch for: students unsettled by period language.
- Protocol: pause if a student is upset, name the language as historical, allow a quiet break, follow up at recess if needed.""",

    4: """TEACHER NOTES:
Teacher orientation only, not for students.""",

    5: """TEACHER NOTES:
Teacher reference for I Do, We Do, You Do badges and support and extension icons. Not student-facing.""",

    6: """TEACHER NOTES:
Teacher reference for the response routines used through the deck. Not student-facing.""",

    7: """TEACHER NOTES:
Teacher reference for sentence-element colour coding. Not student-facing.""",

    8: """SAY:
- "Read the learning intention with me."
- "We are reading more of Storm Boy and learning to grow short sentences into bigger ones."
- "A kernel sentence is tiny, like 'Storm Boy walked.'"
- "Ask: what could you add to tell us where? Expected: along the beach, over the sandhills."
- "Some of you may remember sentence expanding from earlier. If it feels new, that is okay."

DO:
- Choral read the LI, then track each success criterion with your finger.
- Cold call one student to repeat SC1 in their own words.

TEACHER NOTES:
SC3 is the writing target the You Do assesses. Refer back to the LI when you start the writing block.

WATCH FOR:
- Students who cannot say SC3 back. Pair them with a confident partner during the You Do.""",

    9: """SAY:
- "Quick check before we start."
- "You need pencil, booklet, novel, mini whiteboard and texta."
- "Thumbs up when all five are on your desk."

DO:
- Scan the room.
- Wait until every thumb is up before moving on.

TEACHER NOTES:
The lesson uses both novel and whiteboard. Don't begin until materials are out.

WATCH FOR:
- Students with no novel. Pair them with a buddy.""",

    10: """TEACHER NOTES:
Section divider into the read-aloud section. Today's reading mode is teacher choice. The Ochre Text Level Reading guide explains options.""",

    11: """SAY:
- "Open Storm Boy to page 15."
- "I will read aloud. Follow with your eyes or finger."
- "We will stop at 'Neither of them liked...' on page 19."
- "Listen for the big idea: the Coorong is a sanctuary, but people still hurt the birds."
- "Ask: why is Storm Boy sad? Expected: hunters trampled the nests and killed birds."

DO:
- Read pages 15 to 19 with clear pacing.
- Pause briefly at any pre-marked vocabulary words.
- Cold call two students for the comprehension question at the end.

TEACHER NOTES:
Pre-mark page 16 for the personification line you need at the I Do. The big idea about hurt and protection sets up the writing focus.

WATCH FOR:
- Students losing place. Quietly tap their book.
- Students who name the sanctuary but not why Storm Boy is sad. They need a re-read of the key paragraph.""",

    12: """TEACHER NOTES:
Section divider into vocabulary teaching for muttered and trampling.""",

    13: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is muttered."
- "Muttered means to speak in a low voice that is hard to understand. It is a verb."
- "Say it with me: muttered."
- "Listen to it in our story: 'Hide-Away muttered angrily when he heard about the birds.'"

DO:
- Say muttered together two times: quiet first, slightly louder second.
- Mime a low grumble under your breath so students hear the meaning.
- Point to the image as you give the kid-friendly meaning.

TEACHER NOTES:
Muttered comes from page 16 when Hide-Away reacts to the destruction. Connect it to the story moment.

WATCH FOR:
- Students confusing muttered with shouted. Emphasise quiet, hard-to-understand voice.""",

    14: """SAY:
- "When might a person mutter?"
- "A: worried about getting an answer wrong in class. B: excited to show your friend your birthday present."
- "Whiteboards: write A or B."

DO:
- Wait ten seconds.
- Cue: 'Hold up. Show me!'
- Scan the room.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Write A or B.' Expected: A.
- Scan for: A on most boards.
PROCEED:
- >=80% show A. Confirm: 'A. Quiet and unsure.'
PIVOT:
- Most likely: students think any quiet voice is muttering, including a happy whisper.
- Reteach: act out an excited birthday whisper versus a worried mumble.
- Re-check: 'Same question. Which one is muttering?'

WATCH FOR:
- Students who pick B. They missed 'hard to understand'.""",

    15: """SAY:
- "Round two. When might a person mutter?"
- "A: teaching your friends a new game. B: waking up and Mum asks if you want breakfast."
- "Whiteboards: A or B."

DO:
- Wait ten seconds.
- Cue: 'Show me!'
- Scan.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Show me which one would be muttered.' Expected: B.
- Scan for: B on most boards.
PROCEED:
- >=80% show B. Confirm: 'B. Sleepy and not really wanting to talk.'
PIVOT:
- Most likely: students pick A because the speaker is talking.
- Reteach: 'When you teach a friend, do you want them to hear you? Yes, so you speak clearly.'
- Re-check: 'Show me again. Which one is muttering?'

WATCH FOR:
- Slow show-me. Students are guessing, not deciding.""",

    16: """SAY:
- "Look at the two emojis."
- "Which one fits with muttered? Whiteboards: write 1 or 2."

DO:
- Wait ten seconds.
- Show me. Scan.
- Reveal the muttering emoji.

WATCH FOR:
- Students who cannot choose. Cue: 'Show me on your face what muttering looks like.'""",

    17: """SAY:
- "Which words mean nearly the same as muttered?"
- "Read the options together: murmured, mumbled, speaking quietly, speaking clearly."
- "Whiteboards: write all that match."

DO:
- Wait ten seconds.
- Show me. Scan for: murmured, mumbled, speaking quietly.
- Confirm: 'Speaking clearly is the opposite.'

WATCH FOR:
- Students who include 'speaking clearly'. Quick demo of quiet vs clear voice.""",

    18: """SAY:
- "Which sentence uses muttering wrong?"
- "A: After the argument, the man walked away muttering angrily."
- "B: The loud muttering of the crowd filled the stadium."
- "C: The student muttered under her breath."
- "Whiteboards: A, B or C."

DO:
- Wait fifteen seconds.
- Show me. Scan.
- Reveal: B.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Show me the letter of the wrong sentence.' Expected: B.
- Scan for: B on most boards.
PROCEED:
- >=80% show B. Confirm: 'B. We would not say a stadium muttered loudly.'
PIVOT:
- Most likely: students pick A because angrily seems wrong.
- Reteach: read each option asking 'Is this voice quiet or loud?'
- Re-check: 'Show me again. Which one is wrong?'

WATCH FOR:
- Students who pick C. 'Under her breath' is a textbook mutter, not wrong.""",

    19: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is trampling."
- "To trample means to damage by stepping heavily on something. Trampling is a verb."
- "Say it with me: trampling."
- "Listen to it in our story: 'The visitors were trampling about, breaking nests and eggs.'"

DO:
- Say trampling together two times.
- Stomp once gently on the floor to model the heavy step.
- Point to the image and name the action.

TEACHER NOTES:
Trampling captures the harm done at the sanctuary. It is the key word for today's comprehension.

WATCH FOR:
- Students who think trampling is only elephants. Emphasise: any heavy step that damages.""",

    20: """SAY:
- "Is this trampling?"
- "If yes, say trampling. If no, stay quiet."

DO:
- Cue 'three, two, one'.
- Listen for chorus.
- Reveal: trampling.

WATCH FOR:
- Lukewarm chorus. Cue an unsure student by name on the next image.""",

    21: """SAY:
- "Next image. Trampling?"
- "If yes, say it. If no, stay quiet."

DO:
- Cue. Listen.
- Reveal: not trampling.

WATCH FOR:
- Students who say trampling out of habit. Silence is the correct response.""",

    22: """SAY:
- "Image three. Trampling?"
- "If yes, say it."

DO:
- Cue. Listen.
- Reveal: trampling.

WATCH FOR:
- Students staying quiet when they should call out. Acknowledge: 'Good if you were not sure, but this one is trampling.'""",

    23: """SAY:
- "Last one. Trampling?"
- "Quick recap: trampling is heavy steps that damage something."

DO:
- Cue. Listen.
- Reveal: trampling.

WATCH FOR:
- A student who got all four right. Call on them on the next slide.""",

    24: """SAY:
- "Read with me: 'The lady was stepping on plants.'"
- "Rephrase the sentence using trampling. Whiteboards."

DO:
- Wait twenty seconds.
- Show me.
- Reveal model: 'The lady was trampling the plants.'

WATCH FOR:
- Students who copy the original word for word. Cue: 'Swap stepping for trampling.'""",

    25: """SAY:
- "Read with me: 'The can was squashed.'"
- "Rephrase using a form of trample. Whiteboards."

DO:
- Wait twenty seconds.
- Show me.
- Reveal model: 'The can was trampled on.'
- Praise students who used the past tense form.

TEACHER NOTES:
This one needs trampled (past), not trampling. Teach the word family.

WATCH FOR:
- Students who write trampling. Cue: 'It already happened. Past tense.'""",

    26: """SAY:
- "Sort these words on your whiteboard."
- "One side: similar to trampling. Other side: different."
- "Words: stomping, tiptoeing, running, crushing."

DO:
- Wait thirty seconds.
- Show me.
- Reveal: similar are stomping and crushing. Different are tiptoeing and running.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Show me your sort.' Expected: stomping and crushing on similar; tiptoeing and running on different.
- Scan for: correct placement on both sides.
PROCEED:
- >=80% correct. Confirm: 'Stomping and crushing are heavy and damaging.'
PIVOT:
- Most likely: students put running with trampling because both are fast.
- Reteach: 'Could you run without damage? Yes. Could you trample without damage? No.'
- Re-check: 'Sort again with damage in mind.'

WATCH FOR:
- Students who sort tiptoeing as similar. They lost the 'heavy' part.""",

    27: """SAY:
- "In your booklet, write the meaning of muttered in your own words."
- "Then write the meaning of trampling in your own words."
- "Five minutes."

DO:
- Open booklets.
- Circulate.
- Check the first definition each student writes.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: sentence frame: 'Muttered means ___. Trampling means ___.'
- Extra Notes: copy from the vocab slide if needed.
EXTENDING PROMPT:
- Task: write one sentence using each word that connects to Storm Boy.
- Extra Notes: model: 'Storm Boy muttered when he saw the trampled nests.'

TEACHER NOTES:
First You Do. Note who needs vocab support during the writing You Do later.

WATCH FOR:
- Students who copy dictionary phrasing. Check they can say it back simply.""",

    28: """TEACHER NOTES:
Section divider from comprehension into the writing focus on sentence expansion.""",

    29: """SAY:
- "This part of the book uses a writing trick called personification."
- "Personification gives human actions or feelings to animals or objects."
- "Read with me: 'the tall birds stood up and clapped and cheered the rising sun.'"
- "Ask: can birds really clap and cheer? Expected: no. The author makes them sound like people celebrating."

DO:
- Point to the example sentence as students read.
- Mime the clapping and cheering briefly.
- Cold call: 'What human action did the author give the birds?'

TEACHER NOTES:
Light literary device touch under three minutes. Anchors why writers expand sentences with vivid detail.

WATCH FOR:
- Students who think personification is only cartoon talking animals. Clarify: small actions count too.""",

    30: """SAY:
- "Watch this first."
- "We expand sentences for three reasons: more information, more interesting, more like real writers."
- "'Storm Boy walked' is a kernel - true but empty."
- "'One morning, Storm Boy walked along the beach because he wanted to see the pelicans' answers when, where and why."
- "This is the trap: just adding adjectives is not enough."

DO:
- Point to each reason as you read it.
- Write kernel and expanded version side by side on the board.
- Cold call: 'Which one tells you more?'

TEACHER NOTES:
This is the I Do anchor. Students must see the contrast before they try.

WATCH FOR:
- Students who think expanding means more adjectives. Emphasise: we answer when, where, why, how.""",

    31: """SAY:
- "Match each phrase to the question word it answers."
- "'Storm Boy' = who. 'Went walking' = what. 'Over the sandhills' = where. 'One morning' = when."
- "As each phrase highlights, write the question word on your board."
- "Then unscramble the sentence."

DO:
- Reveal each phrase one at a time.
- Wait for whiteboard answers.
- Reveal model: 'One morning, Storm Boy went walking over the sandhills.'
- Note the comma after the WHEN starter.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Write the four question words in order.' Expected: who, what, where, when.
- Scan for: correct labels.
PROCEED:
- >=80% correct. Move on: 'Now we add why and how.'
PIVOT:
- Most likely: students confuse where with when.
- Reteach: 'Where is a place. When is a time.'
- Re-check: 'Same sentence. Which is where?'

WATCH FOR:
- Students who write 'who' for everything. They need where vs when split practice first.""",

    32: """SAY:
- "Round two with how and why."
- "'The pelicans' = who. 'Sat' = what. 'Quietly' = how. 'Because they were not afraid' = why."
- "Write the question word on your board as each phrase highlights."
- "Unscramble the sentence."

DO:
- Reveal each phrase.
- Wait for boards.
- Reveal model: 'The pelicans sat quietly because they were not afraid of Storm Boy.'

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Label each phrase with the question word.' Expected: how matched to quietly, why matched to the because clause.
- Scan for: correct labels.
PROCEED:
- >=80% correct. Move into the I Do (or use the optional adverbial revision slides if needed).
PIVOT:
- Most likely: students label 'because they were not afraid' as how.
- Reteach: 'Because is always why. It gives a reason.'
- Re-check: 'Show me which phrase is why.'

WATCH FOR:
- Students writing without thinking. Slow them: 'Read each phrase out loud first.'""",

    33: """TEACHER NOTES:
The next slides are hidden by default. Unhide them only if students need adverbial revision before the writing You Do.""",

    34: """SAY:
- "An adverbial is a group of words that describes a verb."
- "Example one: 'The visitors hurt the birds by sneaking into the sanctuary.' This tells us how."
- "Example two: 'The birds splashed in the early morning.' This tells us when."
- "Example three: 'Storm Boy went walking along the beach.' This tells us where."
- "If you describe how, when or where with a group of words, you are using an adverbial."

DO:
- Highlight the adverbial in each example on the slide.
- Cold call one student per example: 'How, when or where?'

TEACHER NOTES:
Only run if students need adverbial revision. Adverbial of manner = how. Time = when. Place = where.

WATCH FOR:
- Students who confuse adverbial with adjective. Anchor: 'An adverbial describes the action.'""",

    35: """SAY:
- "When you start a sentence with an adverbial, you must put a comma after it."
- "Look: 'During the open season, shooters chased wounded ducks up the Coorong.'"
- "The comma signals 'pause - the main idea is coming.'"

DO:
- Write the example on the board with the comma highlighted.
- Read it aloud with the pause to model the rule.

TEACHER NOTES:
Keep concrete: the comma sits where your voice naturally pauses.

WATCH FOR:
- Students who place commas mid-clause. The next two practice slides will surface this.""",

    36: """SAY:
- "Where should the comma go?"
- "'During the open season shooters chased wounded ducks up the Coorong.'"
- "Write the word that should come right before the comma."

DO:
- Wait. Show me.
- Reveal: comma after season.

WATCH FOR:
- Students who put comma after shooters or ducks. Re-read with no pause to show why it fails.""",

    37: """SAY:
- "Round two. Where does the comma go?"
- "'When they hit a bird the shooters laughed and walked off.'"
- "Write the word that comes right before the comma."

DO:
- Wait. Show me.
- Reveal: comma after bird.

WATCH FOR:
- Students who hear no pause. Read again slowly with an exaggerated pause.""",

    38: """SAY:
- "Is the comma in the right place?"
- "'When Storm Boy saw the birds he ran, back to his father.'"
- "Whiteboards: yes or no."

DO:
- Wait. Show me.
- Reveal: no. 'When Storm Boy saw the birds' is the adverbial.
- Show correct: 'When Storm Boy saw the birds, he ran back to his father.'

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Yes or no. Is the comma right?' Expected: no.
- Scan for: no on most boards.
PROCEED:
- >=80% show no. Confirm and move on.
PIVOT:
- Most likely: students think any comma is fine.
- Reteach: read 'he ran, back to his father'. It sounds odd.
- Re-check: 'Where should the comma go?'

WATCH FOR:
- Students saying yes. They located a comma but did not check where.""",

    39: """SAY:
- "Round two. Is the comma in the right place?"
- "'When Fingerbone heard about the birds, he slapped his blunderbuss.'"
- "Whiteboards: yes or no."

DO:
- Wait. Show me.
- Reveal: yes. The comma is after the adverbial.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Yes or no?' Expected: yes.
- Scan for: yes on most boards.
PROCEED:
- >=80% show yes. Move into the I Do.
PIVOT:
- Most likely: students say no expecting every example to be wrong.
- Reteach: 'Read it out loud. Does the pause sound right?'
- Re-check: 'Yes or no?'

WATCH FOR:
- Students who guess. Cue: 'Read it in your head before you decide.'""",

    40: """SAY:
- "Watch this first."
- "Question: where did Hide-Away and Fingerbone build their humpies, and why?"
- "Kernel: 'Hide-Away and Fingerbone built their humpies.'"
- "I need to add a where: 'some distance away from the sanctuary.'"
- "I need to add a why: 'because the sanctuary was a protected place.'"
- "Put together: 'Hide-Away and Fingerbone built their humpies some distance away from the sanctuary because the sanctuary was a protected place.'"

DO:
- Read the question aloud, pointing to the question word.
- Build the expanded sentence on the board piece by piece.
- Underline where in one colour, why in another.

MISCONCEPTIONS:
- Misconception: students think expanding means adding more adjectives.
  Why: prior 'juicy words' approaches.
  Impact: padded sentences with no new information.
  Quick correction: 'Expanding answers WHO, WHEN, WHERE, WHY, HOW. Not fancy words.'

TEACHER NOTES:
This is the I Do students will copy as a model. Make every think-aloud step visible.

WATCH FOR:
- Students who copy the kernel only. Keep checking they see the expanded version takes more details.""",

    41: """SAY:
- "Quick check."
- "Kernel: 'Storm Boy walked.'"
- "Write ONE detail to tell us WHERE Storm Boy walked. Twenty seconds."

DO:
- Wait time.
- Cue: 'Three, two, one. Show me!'
- Scan. Reveal possible answers: along the beach, through the sanctuary, over the sandhills, near the water.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Show me one where detail.' Expected: any phrase that names a place.
- Scan for: a place phrase on most boards.
PROCEED:
- >=80% show a where phrase. Move into the We Do.
PIVOT:
- Most likely: students write a when detail like 'one morning'.
- Reteach: 'Where is a place. If you can point to it on a map, it is where.'
- Re-check: 'Same kernel. Show me a where detail this time.'

WATCH FOR:
- Students who write a full sentence. Praise but redirect: 'Just the where part.'""",

    42: """SAY:
- "Now we try one together."
- "Question: how did the hunters enter the sanctuary?"
- "Kernel: 'The hunters entered.'"
- "Ask: what HOW detail can we add? Expected: secretly."
- "Using the prompts: during the night, the sanctuary, to shoot the birds."
- "Full sentence: 'During the night, the hunters secretly entered the sanctuary to shoot the birds.'"

DO:
- Build the sentence in pieces on the board.
- Cold call one student per detail.
- Highlight the comma after the when starter.

TEACHER NOTES:
We Do uses a different sentence to the I Do so students are not just copying.

WATCH FOR:
- Students who add details that do not fit the question. Redirect: 'Does that detail help us answer how?'""",

    43: """SAY:
- "Read the sentence: 'At daybreak, the visitors quickly and secretly sneaked out of the sanctuary.'"
- "Write the WHERE detail and the HOW detail on your whiteboard."

DO:
- Wait. Show me.
- Reveal: where = out of the sanctuary. How = quickly and secretly.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Write the where and how detail.' Expected: 'out of the sanctuary' and 'quickly and secretly'.
- Scan for: both phrases.
PROCEED:
- >=80% locate both. Move into the You Do.
PIVOT:
- Most likely: students label 'at daybreak' as how.
- Reteach: 'Daybreak is a time of day. That is when.'
- Re-check: 'Show me where and how only.'

WATCH FOR:
- Students who find one detail but not both. Will get more practice in the You Do.""",

    44: """SAY:
- "Your turn to expand sentences on your own."
- "First: read the question carefully."
- "Next: pick which details (who, when, where, how, why) help answer it."
- "Then: write your expanded sentence."
- "Challenge: try one sentence with a WHEN or WHERE at the front - and the comma."

DO:
- Distribute or open booklets.
- Set fifteen minutes.
- Circulate. Check the first sentence each student writes.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: five-column table with question word labels (who, when, where, how, why) to fill in before writing.
- Extra Notes: stay with three details rather than four.
EXTENDING PROMPT:
- Task: write the expanded sentence and a second version starting with a different question word.
- Extra Notes: praise variety in starters.

TEACHER NOTES:
Main check on SC3. Mark which students get this on the first sentence.

WATCH FOR:
- Students who write the kernel only. They have lost the task.
- Students adding details that do not relate to the question. Reread the question with them.""",

    45: """SAY:
- "Question one: why did the birds feel safe in the Coorong?"
- "Kernel: 'The birds felt safe.'"
- "Plan who, what, where, why on your worksheet, then write your expanded sentence."

DO:
- Set five minutes.
- Circulate.
- Check the planning grid before students write.

TEACHER NOTES:
Model answer: 'The birds felt safe in the Coorong sanctuary because it was a protected place where nobody could hurt them.' Look for a where and a why at minimum.

WATCH FOR:
- Students who only add why and forget where. Cue: 'Where do they feel safe?'""",

    46: """SAY:
- "Question two: how did Storm Boy feel when he saw what the hunters did?"
- "Kernel: 'Storm Boy felt upset.'"
- "Plan who, how, why on your worksheet, then expand."

DO:
- Set five minutes.
- Circulate.
- Check the planning grid.

TEACHER NOTES:
Model answer: 'Storm Boy felt deeply upset when he saw the destruction because the hunters had trampled through the sanctuary.' Look for a how (intensity) and a why.

WATCH FOR:
- Students who write 'Storm Boy felt sad' only. Cue: 'Why? Add the reason.'""",

    47: """SAY:
- "For these two, choose your own details. No slot prompts."
- "Question three: where did the pelicans gather each morning?"
- "Question four: when and where did Storm Boy walk with the birds?"
- "Plan, then write."

DO:
- Set ten minutes.
- Circulate.
- Note one student to share for each question.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: keep the question word headings visible.
- Extra Notes: tell students which two slots to focus on.
EXTENDING PROMPT:
- Task: combine both expanded sentences into one paragraph that flows.
- Extra Notes: encourage a connector like 'Each morning'.

TEACHER NOTES:
Highest level of independence in this lesson. Students choose their own slots.

WATCH FOR:
- Students who skip planning. It slows them but improves the sentence.""",

    48: """SAY:
- "Turn to your partner."
- "Read your best expanded sentence."
- "Your partner says which question prompts you used."
- "Then swap."

DO:
- One minute each way.
- Thumbs self-assessment for each SC.
- SC1: read and discuss pages 15-19. SC2: identify adverbials. SC3: expand a kernel.

TEACHER NOTES:
Use the SC2 and SC3 thumbs to plan tomorrow. If sideways or down dominates SC3, plan a re-teach.

WATCH FOR:
- Students whose partner cannot name the question word used. The expansion was not tied to the question.""",

    49: """TEACHER NOTES:
Credits and attribution slide. Not student-facing. End the lesson on the closing reflection slide rather than this one.""",

    50: """SAY:
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
    print(f"L06 written: {stats}")


if __name__ == "__main__":
    main()
