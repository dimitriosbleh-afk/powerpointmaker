"""Deep-pass teacher notes for Lesson 8 (Storm Boy - note taking) per teachernotes.md v2.0.

Lean approach: only override substantive teaching slides. polish_notes_lib defaults
handle admin / divider / credits / cultural sensitivity / LI placeholder / materials.
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from polish_notes_lib import polish_deck

SRC = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated/8. Storm Boy - note taking - with notes.pptx")
OUT = SRC.with_name(SRC.stem + " v2.pptx")

NOTES = {
    1: """TEACHER NOTES:
Title slide. Reading focus is pages 24 to 28. Writing focus is note-taking - converting sentences into keywords, phrases, abbreviations and symbols (KPAS).""",

    7: """TEACHER NOTES:
Learning objectives reference slide. Not student-facing.""",

    8: """SAY:
- "Read the learning intention with me."
- "We are reading more of Storm Boy and learning a quick way to take notes called KPAS."
- "KPAS means keywords, phrases, abbreviations and symbols."
- "Ask: why might it be useful to take notes quickly? Expected: to listen and write at the same time, to remember more, to plan writing."
- "If KPAS feels new, that is okay. Some of you may already use shorthand in texts."

DO:
- Choral read the LI, then track each success criterion with your finger.
- Cold call two students for the why-take-notes question.

TEACHER NOTES:
SC3 is the writing target the You Do assesses. Refer back during the writing block.

WATCH FOR:
- Students who cannot say what KPAS means back. Sit them next to a strong note-taker.""",

    11: """SAY:
- "Open Storm Boy to page 24."
- "Today's pages are 24 to 28."
- "I will pause at page 25 'woke up'."
- "Ask at the pause: what is the author saying here? Expected: the pelicans are growing up. They are becoming beautiful, dignified birds."

DO:
- Read pages 24 to 28.
- Pause at page 25 'woke up.'
- Cold call two students for the question, then continue reading.

TEACHER NOTES:
The big idea is the transformation of the pelicans from rescued chicks into dignified adults.

WATCH FOR:
- Students who cannot articulate the change. Prompt: 'How are they different from when Storm Boy first found them?'""",

    13: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is impatient."
- "Impatient means unwilling to wait. It is an adjective."
- "Say it with me: impatient."
- "Listen in our story: 'Mr Proud and Mr Ponder began to get impatient after five minutes.'"

DO:
- Say impatient together two times.
- Mime tapping your foot or checking the clock.
- Cold call: 'When were you last impatient?' Take two answers.

TEACHER NOTES:
Universal feeling. Lets students relate emotionally.

WATCH FOR:
- Students confusing impatient with angry. Anchor: 'Impatient is the feeling of waiting. Anger is broader.'""",

    14: """SAY:
- "Read with me: 'Storm Boy grew impatient as he waited for Hide-Away.'"
- "Write what Storm Boy did to show he was impatient."
- "Whiteboards."

DO:
- Wait. Show me.
- Discuss: paced, looked at the door, sighed.

WATCH FOR:
- Students who write 'he waited'. That is not impatient behaviour. Cue: 'How would you show you were impatient?'""",

    15: """SAY:
- "Round two: 'I was growing impatient in the huge traffic jam.'"
- "Write one thing you might do when impatient in traffic."
- "Whiteboards."

DO:
- Wait. Show me.
- Share two answers.

WATCH FOR:
- Students who write something dangerous. Acknowledge but redirect to safe examples.""",

    16: """SAY:
- "When you are impatient, you are unwilling to wait."
- "Think of a time you were unwilling to wait."
- "Turn to your partner. Share one time."

DO:
- One minute partner talk.
- Cold call two pairs.

WATCH FOR:
- Students who freeze. Cue with a neutral example: waiting for a turn or a birthday.""",

    17: """SAY:
- "Read the sentence: 'I had to wait until Mum came home to open my birthday presents, so I ___.'"
- "Fill in the blank with an impatient action."
- "Whiteboards."

DO:
- Wait two minutes. Show me.
- Share two student examples.

WATCH FOR:
- Students who write a calm action like 'read a book'. Prompt: 'That sounds patient.'""",

    18: """SAY:
- "Round two: 'I had to line up for twenty minutes for an ice cream. I was impatient, so I ___.'"
- "Whiteboards."

DO:
- Wait. Show me.
- Share an example: 'snuck to the front.'

WATCH FOR:
- Students who write the same action twice. Prompt for variety.""",

    19: """SAY:
- "Which word fits best with impatient? Irritated, cautious, calm, unexcited."
- "Whiteboards."

DO:
- Wait. Show me. Scan.
- Reveal: irritated.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Show me your answer. One word.' Expected: irritated.
- Scan for: irritated on most boards.
PROCEED:
- >=80% show irritated. Confirm: 'Irritated. That close, ready-to-snap feeling.'
PIVOT:
- Most likely: students pick calm because impatient feels strong.
- Reteach: 'Calm and impatient are opposites. Irritated is what impatient turns into.'
- Re-check: 'Show me again.'

WATCH FOR:
- Students who pick cautious. They have not connected impatient to a feeling.""",

    20: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is dignified."
- "Dignified means controlled and quite formal in your manner. Other people respect you for it."
- "A simpler way: dignified people stay calm and polite even when it would be easy to lose their temper."
- "Listen in our story: 'The baby pelicans took a few dignified steps forward.'"

DO:
- Say dignified together two times.
- Stand tall and walk slowly to model dignified posture.
- Cold call: 'Who in our school looks dignified?'

TEACHER NOTES:
Ties to the transformation of the pelicans. They grow into dignified birds.

WATCH FOR:
- Students who think dignified means rich. Clarify: 'It is about manner and self-control, not money.'""",

    21: """SAY:
- "Read the sentence: 'Sarah handled losing the game in a dignified manner, congratulating the winner with a smile.'"
- "Write what Sarah did that was dignified."
- "Whiteboards."

DO:
- Wait. Show me.
- Discuss: she stayed calm and congratulated the winner even though she lost.

WATCH FOR:
- Students who write 'she lost'. Cue them to look at how she behaved AFTER losing.""",

    22: """SAY:
- "How are dignity and respect connected?"
- "Turn to your partner. Talk about how they go together."

DO:
- One minute partner talk.
- Cold call non-volunteers.
- Confirm: 'When someone acts with dignity, they earn respect from others.'

WATCH FOR:
- Students who cannot make the connection. Offer: 'If someone yells at the umpire, do you respect them more or less?'""",

    23: """SAY:
- "Which shows a dignified response?"
- "Setup: James is arguing on the playground."
- "A: yelling and calling names back. B: waiting patiently to explain his reasoning."
- "Whiteboards."

DO:
- Wait. Show me.
- Reveal: B.

WATCH FOR:
- Students who pick A. They may think dignified means standing up for yourself loudly.""",

    24: """SAY:
- "Round two. James is annoyed at his little brother for playing with his new toy."
- "A: telling his brother they can play together. B: stealing his brother's toys and hiding them."
- "Whiteboards."

DO:
- Wait. Show me.
- Reveal: A.

WATCH FOR:
- Students who pick B. They read it as 'getting back'. Reteach: 'Dignified means solving calmly.'""",

    25: """SAY:
- "Look at the picture."
- "Write a sentence about it using dignified."
- "Two minutes."

DO:
- Wait. Show me.
- Reveal example: 'The team forgot to be dignified and congratulate the other team.'
- Read two student sentences.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Show me your sentence using dignified.' Expected: dignified used correctly as an adjective.
- Scan for: dignified used to describe a person or behaviour.
PROCEED:
- >=80% correct. Move into the You Do.
PIVOT:
- Most likely: students use dignified as a verb ('He dignified the room').
- Reteach: 'Dignified describes a person or way of acting. Use it like calm or polite.'
- Re-check: 'Show me again.'

WATCH FOR:
- Students who write about the picture without using dignified at all. Redirect.""",

    26: """SAY:
- "In your booklet, write the meaning of impatient in your own words."
- "Then write the meaning of dignified."
- "Five minutes."

DO:
- Open booklets.
- Circulate. Check the first definition.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: sentence frame: 'Impatient means ___. Dignified means ___.'
- Extra Notes: copy from the vocab slide if needed.
EXTENDING PROMPT:
- Task: write one sentence using each word that connects to Storm Boy.
- Extra Notes: model: 'The pelicans took dignified steps forward, no longer impatient.'

WATCH FOR:
- Students who copy dictionary phrasing. Check they can say it back simply.""",

    28: """SAY:
- "Watch this first."
- "Note-taking gets better the more you practise."
- "It involves five things: actively listening, single words or short phrases, finding keywords, summarising into your own words, headings to organise."
- "This is the trap: copying everything is not note-taking. It is the opposite."

DO:
- Point to each item as you read it.
- Cold call: 'What do good note-takers do BEFORE they write?' Expected: listen or read actively.

TEACHER NOTES:
Introduce note-taking. Emphasise it is a skill that improves with practice.

WATCH FOR:
- Students who think note-taking means copying everything. Clarify: 'It is the opposite of copying.'""",

    29: """SAY:
- "Why do we take notes?"
- "Reasons: planning writing, organising thoughts, staying focused, clarifying and remembering."
- "Turn to your partner. Pick the reason you think matters most for YOUR writing."

DO:
- One minute partner talk.
- Cold call three pairs.
- Write the strongest reasons on the board.

TEACHER NOTES:
Note-taking helps students retain information by reinforcing what they have heard or read.

WATCH FOR:
- Students who cannot pick a reason. Prompt: 'When have notes helped you write better?'""",

    30: """SAY:
- "What does note-taking involve?"
- "Option 1: past tense verbs."
- "Option 2: full sentences quickly."
- "Option 3: keywords, symbols, abbreviations."
- "Option 4: words you don't understand."
- "Whiteboards: option number."

DO:
- Wait. Show me.
- Reveal: option 3.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Show me the option number.' Expected: 3.
- Scan for: 3 on most boards.
PROCEED:
- >=80% show 3. Confirm: 'Yes. Keywords, symbols, abbreviations.'
PIVOT:
- Most likely: students pick 2 (full sentences). That is what they have been doing.
- Reteach: 'Full sentences are too slow when listening. Notes are short.'
- Re-check: 'Show me again.'

WATCH FOR:
- Students who pick 1 or 4. They have not understood the purpose of notes.""",

    31: """SAY:
- "Keywords and phrases tell us the main idea."
- "They usually answer who/what, what it did, when, where, why and how."
- "The other words are joiners and extras."

DO:
- Point to each piece of information as you read it.
- Write a quick example on the board: 'The dog barked loudly at the postman.' Underline: dog, barked, postman.

TEACHER NOTES:
I Do for keyword identification. Move slowly so students see what gets left out.

WATCH FOR:
- Students who underline every word. They need to see what gets cut.""",

    32: """SAY:
- "Read with me: 'Pelicans are large water birds that have white feathers.'"
- "Which version shows the keywords?"
- "Whiteboards: option number."

DO:
- Wait. Show me.
- Reveal version with: pelicans, large water birds, white feathers.

TEACHER NOTES:
First guided practice. Look for: pelicans (who/what), large water birds (what they are), white feathers (description).

WATCH FOR:
- Students who keep 'are' and 'that have'. Those are joiners.""",

    33: """SAY:
- "Round two: 'Ever since the miracle of Mr Percival's rescue, he has been Storm Boy's favourite.'"
- "Which version shows the keywords?"
- "Whiteboards."

DO:
- Wait. Show me.
- Reveal version with: Mr Percival's rescue, Storm Boy's favourite.

TEACHER NOTES:
Keywords are the unique, content-rich parts. Joiners get cut.

WATCH FOR:
- Students who keep all proper nouns AND all verbs. Prompt: 'Which words could you cut and still understand?'""",

    34: """SAY:
- "Round three: 'At an astounding 194 kilometres, the Coorong is the longest beach in Australia.'"
- "Which version shows the keywords?"
- "Whiteboards."

DO:
- Wait. Show me.
- Reveal: 194 kilometres, Coorong, longest beach, Australia.

TEACHER NOTES:
Numerical facts are usually keywords. Location and superlatives stay.

WATCH FOR:
- Students who drop the number. Specific facts (numbers, dates, names) almost always stay.""",

    35: """SAY:
- "Watch this first."
- "From our reading: 'As Storm Boy was about to run away, he heard a faint rustling and crying, and there under the sticks and grass of the broken nests were three tiny pelicans - still alive.'"
- "Keywords: Storm Boy. Heard rustling. Under the broken nests. Three tiny pelicans still alive."
- "The joiners and extras get cut."

DO:
- Highlight the keywords on the slide as you say them.
- Compare full sentence and keyword version side by side on the board.

TEACHER NOTES:
I Do for keyword identification in story sentences.

WATCH FOR:
- Students who try to keep the descriptive language. Anchor: 'Notes are functional, not pretty.'""",

    36: """SAY:
- "We Do. Read with me: 'The baby pelicans will need to go back to the sanctuary, we can't afford to feed them anymore.'"
- "Write the keywords on your whiteboard."

DO:
- Wait. Show me. Scan.
- Reveal model: baby pelicans, back to sanctuary, can't afford to feed.

TEACHER NOTES:
Guided practice. Accept any version capturing who, what, why.

WATCH FOR:
- Students who keep 'will need to go'. Cue: ''back to sanctuary' is enough.'""",

    37: """SAY:
- "Round two: 'When Storm Boy went walking along the beach or in the sanctuary, the birds were not afraid because they knew he was a friend.'"
- "Write the keywords."
- "Whiteboards."

DO:
- Wait. Show me.
- Reveal model: 'Storm Boy walking beach sanctuary birds not afraid knew he friend.'

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Show me your keyword version.' Expected: Storm Boy, walking, beach, sanctuary, birds, not afraid, friend.
- Scan for: those keywords without the joiners.
PROCEED:
- >=80% capture the keywords. Move into the multiple choice CFU.
PIVOT:
- Most likely: students keep all conjunctions (when, or, because).
- Reteach: 'Joiners get cut. Keep the words that carry meaning.'
- Re-check: 'Same sentence. Rewrite with no joiners.'

WATCH FOR:
- Students writing fewer than four keywords. They cut too much. Notes need to make sense to a future reader.""",

    38: """SAY:
- "Multiple choice CFU."
- "'The baby pelicans were growing stronger each day.'"
- "Option 1: pelicans growing stronger."
- "Option 2: the baby were each day more."
- "Whiteboards."

DO:
- Wait. Show me.
- Reveal: option 1.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Option 1 or 2?' Expected: 1.
- Scan for: 1 on most boards.
PROCEED:
- >=80% show 1. Move into the shorthand block.
PIVOT:
- Most likely: students pick 2 because it has more words.
- Reteach: 'Option 2 is leftover scraps. Option 1 is the main idea.'
- Re-check: 'Show me again.'

WATCH FOR:
- Students who pick 2. They have not understood that keywords need to make sense.""",

    39: """SAY:
- "Now we add shorthand to make notes faster - abbreviations and symbols."
- "Watch this on the next slide."

DO:
- Build anticipation: 'You are about to learn how to take notes faster than you can write.'

TEACHER NOTES:
Brief intro to the shorthand block.

WATCH FOR:
- Students who already use shorthand. Call on them in the next slide.""",

    40: """SAY:
- "Watch this first."
- "Sentence: 'When Storm Boy went walking along the beach or in the sanctuary, the birds were not afraid because they knew he was a friend.'"
- "Keywords plus shorthand: Storm Boy walking beach + sanctuary -> birds not afraid b/c friend."
- "The plus sign means and. The arrow means leads to. b/c is short for because."

DO:
- Write the shorthand on the board.
- Cold call: 'What does the plus sign mean?' Expected: and.
- Cold call: 'What does b/c mean?' Expected: because.

TEACHER NOTES:
Common shorthand: + (and), b/c (because), arrow (leads to), w/ (with), = (is/equals), # (number).

WATCH FOR:
- Students confused by the symbols. Reassure: 'You do not need to use all of them, just the ones that help.'""",

    41: """SAY:
- "Try this with me."
- "'The whole stretch of the Coorong and the land around it had been turned into a sanctuary, but sometimes Storm Boy saw things that made him sad.'"
- "Write the shorthand version. Whiteboards."

DO:
- Wait. Show me.
- Reveal model: Coorong = sanctuary -> Storm Boy saw things = sad.

TEACHER NOTES:
First guided practice with shorthand. Accept any version with at least one symbol or abbreviation.

WATCH FOR:
- Students who write full keywords with no shorthand. Praise the keywords, prompt for one symbol.""",

    42: """SAY:
- "Round two: 'He wrapped up the tiny bruised body in one of Hideaway's scarves and put it by the fire so it could stay warm.'"
- "Whiteboards."

DO:
- Wait. Show me.
- Reveal model: bruised body wrapped Hide-Away's scarf + put by fire -> keep it warm.

WATCH FOR:
- Students who get the shorthand right but lose the meaning. Anchor: 'Notes still need to make sense.'""",

    43: """SAY:
- "Round three: 'Before long the three pelicans were big and strong.'"
- "Whiteboards."

DO:
- Wait. Show me.
- Reveal model: 3 pelicans = big + strong.

TEACHER NOTES:
Short and quick. Use the digit 3 instead of the word three.

WATCH FOR:
- Students who write 'three'. Prompt them to use the digit.""",

    44: """SAY:
- "Round four: 'Hide-Away spoke sternly to Storm Boy, telling him the pelicans would have to go back to the sanctuary.'"
- "Whiteboards."

DO:
- Wait. Show me.
- Reveal model: Hide-Away spoke sternly w/ Storm Boy -> pelicans = back 2 sanctuary.
- Note: w/ means with, 2 is shorthand for to.

TEACHER NOTES:
Last We Do before the CFU. Students should now be using multiple shorthand pieces.

WATCH FOR:
- Students stuck on which symbols to use. Reassure: 'Pick what works for you.'""",

    45: """SAY:
- "Multiple choice CFU."
- "'Mr Proud, Mr Ponder and Mr Percival will have to go back to the sanctuary. We can't afford to feed them any more.'"
- "Option 1: Mr P+P+P = sanctuary + can't feed them."
- "Option 2: Mr P+P+P -> sanctuary b/c can't feed them."
- "Whiteboards."

DO:
- Wait. Show me.
- Reveal: option 2.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Show me option 1 or 2.' Expected: 2.
- Scan for: 2 on most boards.
PROCEED:
- >=80% show 2. Confirm: 'The arrow shows the move is BECAUSE they cannot feed them.'
PIVOT:
- Most likely: students pick 1 because the equals sign feels familiar.
- Reteach: 'Equals means is. Arrow means leads to or because. We need the why here.'
- Re-check: 'Show me again.'

WATCH FOR:
- Students who pick 1 because it has more symbols. Anchor: 'More is not always better.'""",

    46: """SAY:
- "Your turn."
- "Turn to the page titled Lesson 8: Sentence level writing in your booklet."
- "Complete the note-taking tasks."
- "Use keywords, phrases, abbreviations and symbols."

DO:
- Direct students to the booklet page.
- Set fifteen minutes.
- Circulate. Check the first note each student writes.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: keep the shorthand reference visible (+, -, b/c, w/, =, arrow).
- Extra Notes: keyword-only first, then shorthand on a second pass.
EXTENDING PROMPT:
- Task: write notes for an additional sentence using the shortest version that still makes sense.
- Extra Notes: students share their shortest correct note with a partner.

TEACHER NOTES:
Main You Do for SC3. Track which students take usable notes.

WATCH FOR:
- Students who copy whole sentences. Redirect.
- Students whose notes do not make sense to themselves. They cut too much.""",

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
    print(f"L08 written: {stats}")


if __name__ == "__main__":
    main()
