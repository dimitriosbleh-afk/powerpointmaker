"""Deep-pass teacher notes for Lesson 11 (Storm Boy - Switching the adverbial) per teachernotes.md v2.0."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from polish_notes_lib import polish_deck

SRC = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated/11. Storm Boy - Switching the adverbial  - with notes.pptx")
OUT = SRC.with_name(SRC.stem + " v2.pptx")

NOTES = {
    1: """TEACHER NOTES:
Title slide. Two strands today: read pages 29 to 32 and learn to switch adverbial position. Begin once materials are out and the novel is open at page 29.""",

    7: """SAY:
- "Read the learning intention with me."
- "Two parts today: read about Mr Percival's return, then learn to switch the adverbial in a sentence."
- "Read each success criterion together. Three I can statements. We will check them at the end."

DO:
- Choral read the LI, then track each SC with your finger.
- Cold call one student to repeat SC3 in their own words.

TEACHER NOTES:
SC1 is comprehension. SC2 is finding the adverbial. SC3 is moving it and applying the comma rule. SC3 is the exit-ticket focus.

WATCH FOR:
- Students unsure what 'adverbial' means. Reassure: 'If this feels new, that is okay. We will build it together.'""",

    8: """SAY:
- "Get your mini whiteboard, texta, booklet, novel and pencil ready."
- "Pencil cases on the floor."

DO:
- Scan the room. Wait until every student has the four items.
- Distribute spares for any student missing materials.

TEACHER NOTES:
Set up before the bell where possible so the start of the lesson is not lost to materials checks.""",

    10: """SAY:
- "We are reading pages 29 to 32 of Storm Boy today."
- "Big idea: Mr Percival returns and Storm Boy is overjoyed."
- "Listen out for two new words: 'miserably' and 'reunion'."
- "Find page 29."

DO:
- Give 30 seconds for students to find page 29.
- Read aloud with expression.
- Pause at the marked points: 'with his fishing' (p.30), 'bird of stone' (p.30), 'come back home' (p.31) and 'Storm Boy's side' (p.32).
- Celebrate when students hear 'miserably' or 'reunion'.

TEACHER NOTES:
Confirm pause points against the Literature Study Guide before teaching.

WATCH FOR:
- Students tracking Storm Boy's emotional shift from sad to overjoyed. Excellent comprehension.
- Students who lose place during pauses. Check fingers on the line before continuing.""",

    12: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is refused."
- "Refused means saying you will not do something."
- "Say the word with me: refused."
- "Refused is a verb. It tells us about an action."
- "In the story, Storm Boy refused to believe Mr Percival had gone."

DO:
- Point to the picture as you give the meaning.
- Choral repeat the word twice.
- Gesture: shake head and cross arms.

TEACHER NOTES:
Introduce vocabulary explicitly before students meet it in sentence practice.

WATCH FOR:
- Students confusing 'refused' with passive listening. Anchor: 'Refused is the active 'I will not'.'""",

    13: """SAY:
- "Read the sentence with me: 'Despite his father's insistence on leaving the pelican, Storm Boy refused to abandon his feathered friend.'"
- "Ask: what did Storm Boy refuse to do? Expected: leave the pelican."

DO:
- Choral read the sentence.
- Cold call two students for the meaning of refused in this sentence.

WATCH FOR:
- Students who say 'fight' or 'argue' instead of 'said no to'. Anchor: 'Refused means said no.'""",

    14: """SAY:
- "Read the sentence with me: 'Tim refused to eat his vegetables even though his mum promised he could have dessert afterwards.'"
- "Ask: what did Tim refuse to do? Expected: eat his vegetables."

DO:
- Choral read.
- Quick partner share: another time you refused to do something.

WATCH FOR:
- Students who give a non-example (he ate them anyway). Reteach: 'Refuse means say no to.'""",

    15: """SAY:
- "Think of something you might refuse to do."
- "Example: sing on stage in front of a crowd. I would refuse to do that."
- "Now your turn. Think for ten seconds."

DO:
- Give ten seconds wait time.
- Cold call three non-volunteers for examples.
- Listen for the word 'refused' in the response.

WATCH FOR:
- Examples that do not actually use refused. Rephrase the response back using refused.""",

    16: """SAY:
- "Build a sentence using the word refused."
- "The frame: 'The parents refused [blank], but [blank].'"
- "One example: 'The parents refused to serve ice-cream for dinner, but they did agree to ice-cream for dessert.'"
- "On your whiteboards, finish your own version."

DO:
- Whiteboards out. Give 90 seconds writing time.
- Circulate. Read 4 to 5 examples aloud.
- Praise correct use of refused as a verb.

WATCH FOR:
- Students using 'refuse' (present tense) instead of 'refused' (past tense). Quick correction: 'We are practising the past tense form.'""",

    17: """SAY:
- "Look at the four pictures."
- "Which one would Storm Boy probably refuse to do? Why?"
- "Options: walk along the beach, talk to Fingerbone, train a pelican, invite the hunters to dinner."

DO:
- Show fingers 1 to 4.
- Wait time, then on cue all show together.
- Cold call two students to justify.

CFU CHECKPOINT:
Technique: Show Fingers 1-4
Script:
- Ask: 'Show me with your fingers which one Storm Boy would refuse to do.' Expected: 4.
- Scan for: 4 (invite the hunters to dinner) on most boards.
PROCEED:
- >=80% pick 4. Brief celebration, move on.
PIVOT:
- Most likely: students choose 1 or 2. Misconception: not connecting Storm Boy's love of the birds and his hatred of the hunters from the novel.
- Reteach: re-read the relevant part of pages 29 to 32 and ask: 'Who does Storm Boy not want at his home?'
- Re-check: 'Now show me again. Which one would Storm Boy refuse to do?'

WATCH FOR:
- Students copying neighbours. Cover the hand briefly and re-cue.""",

    18: """SOURCES:
Macquarie Dictionary, 2024. Image: Screen Australia.

SAY:
- "The word is reunion."
- "A reunion is a special meeting, usually of family or people who have not seen each other for a long time."
- "Say the word with me: reunion."
- "Reunion is a noun. It names a thing."
- "In the story, the reunion between Storm Boy and Mr Percival brought much joy."

DO:
- Point to the picture as you give the meaning.
- Choral repeat the word twice.
- Gesture: open arms wide as if hugging.

TEACHER NOTES:
Key lesson word. Students will hear it again in the read-aloud pages.

WATCH FOR:
- Students saying 'reunion' as 'reuniting'. Clarify: 'Reunion is the event. Reuniting is the action.'""",

    19: """SAY:
- "Read the sentence with me: 'When Storm Boy saw Mr Percival again, it was like a grand reunion of two best friends who hadn't seen each other in a long time.'"
- "Ask: what does reunion mean here? Expected: meeting again after time apart."

DO:
- Choral read.
- Cold call one student for the meaning.

WATCH FOR:
- Students who think reunion just means 'meeting'. Press: 'Special meeting after a long time apart.'""",

    20: """SAY:
- "Read the sentence with me: 'Amy was excited about the family reunion, she couldn't wait to see her cousins and play games all day.'"
- "Ask: who is the reunion between? Expected: Amy and her cousins / her family."

DO:
- Choral read.
- Quick partner share: who would be at your family reunion?

WATCH FOR:
- Students who name only one person rather than a group. Anchor: 'A reunion brings people back together.'""",

    21: """SAY:
- "Think of a time you might have had a reunion."
- "Example: coming back to school after the holidays - that is a reunion with your friends."
- "Now your turn. Think for ten seconds."

DO:
- Give ten seconds wait time.
- Cold call three students for examples.
- Listen for 'reunion' used correctly as a noun.

WATCH FOR:
- Students who confuse a one-off meeting with a reunion. Press: 'Have you been apart for a while?'""",

    22: """SAY:
- "Build a sentence using the word reunion."
- "The frame: 'The lost [blank] had a joyful reunion with [blank].'"
- "One example: 'The lost puppy had a joyful reunion with its family.'"
- "On your whiteboards, finish your own version."

DO:
- Whiteboards out. Give 90 seconds.
- Circulate. Read 4 to 5 examples aloud.
- Praise correct use of reunion as a noun.

WATCH FOR:
- Students writing 'reunioned' or 'reunioning'. Anchor: 'Reunion is a noun only. The verb form is reunite.'""",

    23: """SAY:
- "Sort the words into two columns: similar to reunion, and different from reunion."
- "Words: gathering, break up, get together, farewell."

DO:
- Whiteboards split into two columns.
- Give 60 seconds.
- Reveal answers: similar are gathering and get together. Different are break up and farewell.

WATCH FOR:
- Students who put 'gathering' on different. Press: 'People coming together, like a reunion.'""",

    24: """SAY:
- "Which word fits best with reunion?"
- "Options: dispute, falling out, meeting, split."
- "Show fingers 1, 2, 3 or 4."

DO:
- Wait time. On cue all show together.
- Cold call one student to justify.

CFU CHECKPOINT:
Technique: Show Fingers 1-4
Script:
- Ask: 'Show me which word fits best with reunion.' Expected: 3 (meeting).
- Scan for: 3 across the room.
PROCEED:
- >=80% pick 3. Move to the comprehension task.
PIVOT:
- Most likely: students show 1, 2 or 4. Misconception: linking reunion to argument or separation rather than coming back together.
- Reteach: 'Reunion comes from re-unite, which means coming back together. Which option means coming together?'
- Re-check: 'Show me again. Which fits with reunion?'

WATCH FOR:
- Students copying. Cover the hand and re-cue.""",

    25: """SAY:
- "Open your booklet to the Reading Comprehension page for Lesson 11."
- "Write the meaning of refused in your own words."
- "Write the meaning of reunion in your own words."
- "You can use the example sentences if you need."

DO:
- Distribute or direct to the booklet page.
- Set timer for 4 minutes.
- Circulate. Check first response of three students.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: provide a sentence frame: 'Refused means [blank]. A reunion is when [blank].'
EXTENDING PROMPT:
- Task: write your own sentence using each word, in a new context.

WATCH FOR:
- Students who copy directly from the slide. Cue: 'In your own words, not the dictionary's words.'""",

    27: """SAY:
- "Today we are learning about adverbial phrases."
- "An adverbial phrase is a group of words that describes a verb. It tells how, when, where or why."
- "Watch this first. I will show you two examples."
- "Example one: 'Without talking, they just sat in the bobbing boat.' The adverbial 'without talking' tells how."
- "Example two: 'Towards evening, they packed up and set off for home.' The adverbial 'towards evening' tells when."

DO:
- Underline 'without talking' as you say 'how'.
- Underline 'towards evening' as you say 'when'.
- Hold up two fingers. We have two adverbial types so far.

TEACHER NOTES:
I Do for adverbial phrases. Definition first, then two clear examples (how and when). Examples are taken from Storm Boy.

WATCH FOR:
- Students confusing the verb with the adverbial. Quick check: 'What does the adverbial describe?' Answer: a verb.""",

    28: """SAY:
- "Two more types of adverbial phrases."
- "'The sun was flinging a million golden mirrors in a lane across the water.' The adverbial 'across the water' tells where."
- "'Hide-Away and Storm Boy were feeling sad because they missed Mr Percival, Mr Ponder and Mr Proud.' The adverbial 'because they missed Mr Percival, Mr Ponder and Mr Proud' tells why."

DO:
- Underline 'across the water' as you say 'where'.
- Underline the 'because' phrase as you say 'why'.
- Hold up four fingers: how, when, where, why.

TEACHER NOTES:
I Do continued. Completing the four adverbial types: how, when, where, why.

WATCH FOR:
- Students mixing up where and when. Quick check: 'Where is location. When is time.'""",

    29: """SAY:
- "Adverbials can sit at the start or the end of a sentence."
- "Example one: 'They just sat in the bobbing boat without talking.' Adverbial at the end - no comma needed."
- "Example two: 'Without talking, they just sat in the bobbing boat.' Adverbial at the start - comma is required."
- "The comma rule: if the adverbial is at the start, you need a comma."

DO:
- Point to the comma in example two.
- Hold up a thumb when you say 'comma needed', thumb down when you say 'no comma needed'.
- Choral repeat the rule: 'Adverbial at the start needs a comma.'

MISCONCEPTIONS:
- Misconception: a comma is needed for every adverbial.
  Why: students remember the comma but not the position rule.
  Impact: over-commaing in their writing.
  Quick correction: 'Only at the start. End-position adverbials do not get a comma.'

TEACHER NOTES:
This is the core rule for the lesson. Students will apply it in the next slides.

WATCH FOR:
- Students who do not say the rule back fluently. Repeat the choral cue once more.""",

    30: """SAY:
- "Watch me first."
- "Original sentence: 'As the sun began to set, they packed up and headed home.'"
- "The adverbial is 'as the sun began to set' - this tells when."
- "I am going to move it to the end: 'They packed up and headed home as the sun began to set.'"
- "Now there is no comma. The adverbial is at the end."

DO:
- Use the slide to highlight the adverbial chunk first.
- Move it to the end and erase the comma.
- Choral re-read both sentences.

TEACHER NOTES:
First modelled switch. Adverbial of time. Note the comma is removed when the adverbial moves to the end.

WATCH FOR:
- Students missing that the comma is gone. Cue: 'Where did the comma go? Why?'""",

    31: """SAY:
- "Let's do this one together."
- "Original: 'The sun cast a million golden mirrors across the water.'"
- "Ask: what is the adverbial? Expected: across the water."
- "Ask: where is the adverbial? Expected: end."
- "Now move it to the start. What do we add? Expected: a comma."
- "'Across the water, the sun cast a million golden mirrors.'"

DO:
- Whiteboards: students write the new sentence with the comma.
- Scan for the comma after 'water'.
- Reveal the answer.

TEACHER NOTES:
We Do switch from end to start. This is the harder direction - students must add the comma. Adverbial of place.

WATCH FOR:
- Students who write the new sentence without the comma. Targeted reteach using the rule: 'Adverbial at the start needs a comma.'""",

    32: """SAY:
- "Together again."
- "Original: 'All of a sudden, Storm Boy looked up at the sky.'"
- "Ask: what is the adverbial? Expected: all of a sudden."
- "Ask: where is it? Expected: start."
- "Move it to the end. What do we drop? Expected: the comma."
- "'Storm Boy looked up at the sky all of a sudden.'"

DO:
- Whiteboards: students write the new sentence.
- Scan for no comma.
- Reveal the answer.

TEACHER NOTES:
We Do switch from start to end. Adverbial of manner ('all of a sudden' tells how).

WATCH FOR:
- Students who keep the comma after 'sky'. Reteach: 'End-position adverbials do not need a comma.'""",

    33: """SAY:
- "Last together one."
- "Original: 'The shape balanced on top of the post as still as a statue.'"
- "Ask: what is the adverbial? Expected: as still as a statue."
- "Ask: where is it? Expected: end."
- "Move it to the start. What do we add? Expected: a comma."
- "'As still as a statue, the shape balanced on top of the post.'"

DO:
- Whiteboards: students write the new sentence.
- Scan for the comma after 'statue'.
- Reveal.

TEACHER NOTES:
We Do switch from end to start. Adverbial of manner.

WATCH FOR:
- Students who only move part of the adverbial. Press: 'The whole phrase moves together.'""",

    34: """SAY:
- "Hinge question. Which sentence needs a comma?"
- "Sentence 1: 'It was a happy reunion for the bird and boy that night.'"
- "Sentence 2: 'That night it was a happy reunion for the bird and boy.'"
- "Show fingers: 1 or 2."

DO:
- Wait time. On cue all show together.
- Cold call one student to justify.

CFU CHECKPOINT:
Technique: Show Fingers 1 or 2
Script:
- Ask: 'Show me which sentence needs a comma.' Expected: 2.
- Scan for: 2 on most boards.
PROCEED:
- >=80% show 2. Move to the next We Do.
PIVOT:
- Most likely: students show 1. Misconception: thinking commas are about the verb or the noun, not the adverbial position.
- Reteach: highlight 'that night' in both sentences. Where does it sit? Start = comma. End = no comma.
- Re-check: 'Show me again. Which one needs the comma?'

TEACHER NOTES:
Key hinge for the lesson. Answer: Sentence 2 (adverbial 'that night' is at the start). The comma goes after 'night'.""",

    35: """SAY:
- "Complete the sentence by adding an adverbial. Then switch its position."
- "Stem: 'Hide-Away watched Storm Boy fishing [blank].'"
- "One example: 'Hide-Away watched Storm Boy fishing rather miserably.'"
- "Now switch: 'Rather miserably, Hide-Away watched Storm Boy fishing.'"
- "A comma is required at the start."

DO:
- Whiteboards.
- Give 2 minutes. First complete the sentence, then switch.
- Circulate. Read out 3 strong examples.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: provide an adverbial bank: 'rather miserably', 'in the early morning', 'with quiet eyes'. Students choose one.
EXTENDING PROMPT:
- Task: add two adverbials, one of time and one of manner. Switch both.

TEACHER NOTES:
We Do that combines adverbial generation and position-switching. Connects to today's read-aloud word 'miserably'.

WATCH FOR:
- Students who write the same sentence twice without switching. Press: 'The second version must have the adverbial in a new position.'""",

    36: """SAY:
- "Same task again."
- "Stem: 'With a secret sense of happiness, [blank] welcomed Mr Percival home.'"
- "One example: 'With a secret sense of happiness, Hide-Away welcomed Mr Percival home.'"
- "Now switch: 'Hide-Away welcomed Mr Percival home with a secret sense of happiness.'"
- "A comma is not required at the end."

DO:
- Whiteboards.
- Give 2 minutes.
- Circulate. Read out 3 strong examples.

TEACHER NOTES:
We Do switching from start to end. Note the comma drops when the adverbial moves to the end.

WATCH FOR:
- Students who keep the comma after 'home'. Reteach: 'End-position adverbials drop the comma.'""",

    37: """SAY:
- "Hinge question. Is the comma in the correct place?"
- "'With all his might, Storm Boy hugged Mr Percival.'"
- "Thumbs up if yes. Thumbs down if no."

DO:
- Wait time. On cue all show together.
- Reveal: yes. The comma is after the adverbial phrase 'with all his might'.

CFU CHECKPOINT:
Technique: Thumbs Up or Down
Script:
- Ask: 'Thumbs up if the comma is correct, thumbs down if it is wrong.' Expected: thumbs up.
- Scan for: thumbs up on most.
PROCEED:
- >=80% thumbs up. Move to the You Do task.
PIVOT:
- Most likely: students thumbs down. Misconception: not seeing 'with all his might' as the adverbial.
- Reteach: highlight the adverbial chunk. 'Adverbial is at the start, so comma is required.'
- Re-check: 'Show me again. Is the comma correct?'

TEACHER NOTES:
Final CFU before the You Do. Answer: yes, the comma is correct.""",

    38: """SAY:
- "Open your booklet to Lesson 11 sentence-level writing."
- "Q2: switch the adverbial position in two sentences."
- "Q3: finish the sentence with an adverbial, then switch its position."
- "Read each sentence carefully. Check the comma rule each time."
- "On Q3: First, write your adverbial. Next, decide its position. Then check the comma."

DO:
- Direct to the booklet.
- Set timer for 8 to 10 minutes.
- Circulate. Check first response of 3 to 5 students.
- Use 'A comma is required' or 'A comma is not required' as the success cue.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: provide a colour-coded version with the adverbial highlighted. Students only switch and add or remove the comma.
EXTENDING PROMPT:
- Task: Q3 - write two different versions of the same adverbial in different positions.

TEACHER NOTES:
You Do task. SC3 exit-ticket-style task. Q2 sentences: 'Throughout the day, the devoted bird stayed by Storm Boy's side' and 'The three pelicans left the humpy quite happily.'

WATCH FOR:
- Students who switch the adverbial but forget the comma rule.
- Students who change the meaning by moving only some of the adverbial. Press: 'The whole phrase moves together.'
- Fast finishers. Prompt: 'Try a third position by combining adverbials.'""",

    40: """SAY:
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
    print(f"L11 written: {stats}")


if __name__ == "__main__":
    main()
