"""Deep-pass teacher notes for Lesson 13 (Note taking) per teachernotes.md v2.0."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from polish_notes_lib import polish_deck

SRC = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated/13. literature_presentation Note taking 4En06V Storm Boy - Information report - with notes.pptx")
OUT = SRC.with_name(SRC.stem + " v2.pptx")

NOTES = {
    1: """TEACHER NOTES:
Title slide. Today's focus: read pages 54 to 57 and learn note-taking techniques (keywords, shorthand) to support summarising.""",

    6: """SAY:
- "These are our learning objectives across the unit."
- "Today we focus on the last one: identifying keywords and converting sentences into notes."
- "Read the objectives aloud with me."

DO:
- Choral read each objective.
- Highlight the note-taking objective verbally.

WATCH FOR:
- Students new to shorthand. Reassure: 'We will revisit each symbol.'""",

    7: """SAY:
- "Read the learning intention with me."
- "Today we read pages 54 to 57 and learn to take notes that capture the main ideas of a text."
- "Ask: why might note-taking help your writing? Expected: it saves time, helps you remember, helps plan."
- "If shorthand feels new, that is okay. We will build it together."

DO:
- Choral read the LI, then track each success criterion with your finger.
- Cold call one student to repeat SC3 in their own words.

TEACHER NOTES:
SC2 and SC3 are the writing targets the You Do task assesses.

WATCH FOR:
- Students unsure what 'keywords' means. Quick gloss: 'The most important words in a sentence.'""",

    10: """SAY:
- "Today we are reading pages 54 to 57 of Storm Boy."
- "Pause when I stop to check our thinking."
- "Find page 54."

DO:
- Give 30 seconds for students to find page 54.
- Read aloud with expression.
- Use the pause points and queries from the Literature Study Guide.

TEACHER NOTES:
Pre-mark your novel before teaching. Select 3 to 4 pause points.

WATCH FOR:
- Students tracking the storm scene. Listen for the words 'bellowed' and 'shivered'.""",

    12: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is bellowed."
- "To bellow is to cry loudly and deeply. It is a verb."
- "Say the word with me: bellowed."
- "In the story, Fingerbone bellowed that the storm was too loud."

DO:
- Point to the picture as you give the meaning.
- Choral repeat the word twice.
- Demonstrate a brief deep, loud call (not at full volume) to model the sound.

TEACHER NOTES:
First vocabulary word. Stress the deep, loud quality. Bellowing is bigger than shouting.

WATCH FOR:
- Students who say bellowed sounds the same as 'yelled'. Press: 'Bellowing is deeper and louder, like a foghorn.'""",

    13: """SAY:
- "Read the sentence with me: 'Don't kick the ball inside the house!' Dad bellowed from the other room."
- "Ask: how did Dad call out? Expected: bellowed (loud, deep)."

DO:
- Choral read with expression.
- Cold call one student for the meaning of bellowed.

WATCH FOR:
- Students who use a soft voice in choral reading. Cue: 'Read it as if Dad is really cross.'""",

    14: """SAY:
- "Which sentence uses bellowed correctly?"
- "Sentence A: 'The smoke from the fire bellowed above us.'"
- "Sentence B: 'The teacher bellowed across the courtyard when my friend was being unsafe.'"
- "Show fingers A or B."

DO:
- Wait time. On cue all show together.
- Reveal: Sentence A is wrong - the correct word for smoke is 'billowed'. Sentence B is correct.

MISCONCEPTIONS:
- Misconception: bellow and billow are the same word.
  Why: they sound similar and both involve big, dramatic action.
  Impact: misuse in writing (smoke bellowing instead of billowing).
  Quick correction: 'Bellow is a sound. Billow is a movement. Smoke billows, people bellow.'

TEACHER NOTES:
Noticing the homophone confusion: bellow (cry loudly) vs billow (smoke or cloth swelling). Answer: B.""",

    15: """SAY:
- "Think of a time you have bellowed at someone."
- "Watch me first: I bellowed at my dog when she chewed my new shoes."
- "Now silently act out your bellowing for the class."

DO:
- Model an example with a deep, loud (mimed) voice.
- Give 30 seconds think time.
- Invite 3 volunteers to act out (silently - face and body only).
- Discuss: facial features, mouth shape, body language as evidence.

TEACHER NOTES:
Kinaesthetic vocabulary practice.

WATCH FOR:
- Students who do small face expressions. Press: 'Bellowing is BIG and obvious.'""",

    16: """SAY:
- "Look at the picture. Write a sentence about it using the word bellowed."
- "One example: 'Make sure you score!' the coach bellowed at us as we ran onto the field."

DO:
- Whiteboards out. Give 2 minutes.
- Click to reveal the example.
- Cold call two students to share their sentence.

WATCH FOR:
- Students who use bellowed for a quiet voice. Press: 'Bellowing is loud and deep.'""",

    17: """SAY:
- "When you bellow, you..."
- "Options: cry loud and high pitched, scream as loud as you can, cry loud and deeply, groan in frustration."
- "Show fingers 1 to 4."

DO:
- Wait time. On cue all show together.
- Cold call one student to justify.

CFU CHECKPOINT:
Technique: Show Fingers 1-4
Script:
- Ask: 'Show me what bellow means.' Expected: 3.
- Scan for: 3 across the room.
PROCEED:
- >=80% pick 3. Move to the next vocabulary word.
PIVOT:
- Most likely: students think bellowed is just any loud noise.
- Reteach: 'Bellowing is loud AND deep. Like a foghorn or a bull. Not high pitched.'
- Re-check: 'Show me again.'

WATCH FOR:
- Students who pick 1 (high pitched). Reteach: 'Deep, not high.'""",

    18: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is shivered."
- "To shiver is to shake with fear, cold or excitement. It is a verb."
- "Say the word with me: shivered."
- "In the story, the humpy shivered and shook in the storm."

DO:
- Point to the picture as you give the meaning.
- Choral repeat the word twice.
- Gesture: hug arms around body and shake briefly.

TEACHER NOTES:
Second vocabulary word. Three causes: fear, cold, excitement.

WATCH FOR:
- Students who only think of cold. Press: 'You can also shiver from fear or excitement.'""",

    19: """SAY:
- "Read the sentence with me: 'I shivered when I heard the wind howling outside my bedroom window.'"
- "Ask: why did I shiver? Expected: from cold or fear."

DO:
- Choral read.
- Cold call one student for the cause.

WATCH FOR:
- Students who only say 'cold'. Press: 'Could it also be fear?'""",

    20: """SAY:
- "Think of a time you shivered from excitement."
- "Example: when my mum brought home a new puppy."
- "Tell your partner your example."

DO:
- Partner talk for 30 seconds.
- Cold call two students to share.
- Listen for excitement-based shiver examples (presents, surprises, big news).

WATCH FOR:
- Students who say 'shiver from cold'. Cue: 'We are looking at shivers from excitement here.'""",

    21: """SAY:
- "Think of a time you shivered with fear."
- "Example: when my brother told me a scary story."
- "Tell your partner your example."

DO:
- Partner talk for 30 seconds.
- Cold call two students to share.
- Listen for fear-based shiver examples (scary movies, dark hallways, storms).

SENSITIVITY ADVISORY:
- What it is: students may share genuine fears (storms, scary events at home).
- Framing language: 'A pretend example or a small fear is fine. We do not need big personal stories.'
- Watch for: a student who looks distressed when sharing.
- Protocol: privately check in if a student seems upset. Redirect to a milder example.

WATCH FOR:
- Students sharing personal trauma. Redirect to mild fears.""",

    22: """SAY:
- "Which words have a similar meaning to shivering?"
- "Thumbs up for similar, thumbs down for different."
- "quivering... trembling... gliding... cold... shaking... chattering."

DO:
- Run each word one at a time.
- Wait time on each, then on cue all show.
- Reveal: quivering (yes), trembling (yes), gliding (no), cold (no - it is a cause not a synonym), shaking (yes), chattering (yes - of teeth).

TEACHER NOTES:
Synonym sort. Discuss with students - 'gliding' is wrong (smooth, not shaky), 'cold' is the cause not the action.

WATCH FOR:
- Students who put 'gliding' as similar. Press: 'Gliding is smooth, shivering is shaky.'""",

    23: """SAY:
- "Sentence building. Finish each sentence using a conjunction."
- "'My sister was shivering because [blank]' - because tells us why."
- "'My sister was shivering, but [blank]' - but tells us a contrast."
- "'My sister was shivering, so [blank]' - so tells us a result."

DO:
- Whiteboards. One sentence at a time.
- Click through the reveals.
- Read out 2 student examples per sentence type.

CFU CHECKPOINT:
Technique: Whiteboard scan
Script:
- Ask: 'Hold up your version of the because sentence.' Expected: a cause that makes sense.
- Scan for: cold, scared, ill.
PROCEED:
- >=80% correctly use the conjunction. Move to the booklet.
PIVOT:
- Most likely: students mix up because and so.
- Reteach: 'because gives a reason, so gives a result.'
- Re-check: 'Hold up the so sentence.'

WATCH FOR:
- Students writing the same content for both conjunctions. Press: 'Different conjunction, different idea.'""",

    24: """SAY:
- "Open your booklet to Lesson 13 Vocabulary."
- "Complete the booklet tasks for bellowed and shivered."
- "First: read each prompt. Next: write your sentence. Then: check you used the word correctly."

DO:
- Direct to the booklet page.
- Set timer for 6 minutes.
- Circulate. Check first response of 3 to 5 students.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: provide sentence frames: 'I bellowed when [blank]. I shivered because [blank].'
EXTENDING PROMPT:
- Task: write one sentence that uses both words together (e.g. 'I shivered when Dad bellowed at the dog').

WATCH FOR:
- Students misusing bellowed for soft sounds, or shivered for non-shaky things.""",

    26: """SAY:
- "Note-taking is part of good planning."
- "It means: actively listening or reading, picking out keywords or main points, using single words or short phrases, summarising in your own words, using headings."
- "Some of you may have done note-taking before. If this feels new, that is okay. We will build it together."
- "Why does it help? It saves time and helps you remember."

DO:
- Read each bullet aloud.
- Cold call: 'When have you taken notes before?'
- Brief class discussion. Capture good responses on the board.

TEACHER NOTES:
I Do for note-taking. Note-taking is a transferable skill. Emphasise it improves with practice.

WATCH FOR:
- Students who say 'I just write everything down'. Press: 'That is copying, not note-taking.'""",

    27: """SAY:
- "Why do we take notes? Tell me some reasons."
- "Reasons: to plan and prepare for writing, to organise thoughts, to stay focused, to clarify and remember information."
- "Watch this first. I will read each reason aloud."

DO:
- Class discussion: collect 3 to 4 student reasons before revealing.
- Read each revealed reason aloud.

TEACHER NOTES:
We Do for the reasons we take notes. Connects today's lesson to its purpose.

WATCH FOR:
- Students who say 'because the teacher told me to'. Press: 'What is the real benefit for you?'""",

    28: """SAY:
- "When note taking, you..."
- "Options: write only what you remember, shorten every word, use keywords/symbols/abbreviations, copy exactly what the text says."
- "Show fingers 1 to 4."

DO:
- Wait time. On cue all show together.
- Cold call one student to justify.

CFU CHECKPOINT:
Technique: Show Fingers 1-4
Script:
- Ask: 'Show me what you do when note taking.' Expected: 3.
- Scan for: 3 on most boards.
PROCEED:
- >=80% show 3. Move to the keyword section.
PIVOT:
- Most likely: students show 4 (copy). Misconception: notes equal copying.
- Reteach: 'Notes are the keywords only. Not full sentences. Not every word.'
- Re-check: 'Show me again.'

TEACHER NOTES:
Quick CFU. Answer: 3.""",

    29: """SAY:
- "Keywords and phrases are the parts of a sentence that tell us the main idea."
- "Often they tell us who, what, when, where, why and how."
- "Watch the next slides. I will show you how to find the keywords."

DO:
- Read each bullet aloud.
- Hold up six fingers: who/what, what doing, when, where, why, how.

TEACHER NOTES:
I Do for keyword identification. The 5W+H question prompts give students a way to find the main ideas.

WATCH FOR:
- Students confused about which words to highlight. Cue: 'Ask each question - who? what doing? when? where? why? how?'""",

    30: """SAY:
- "We Do. Which sentence highlights the keywords correctly?"
- "Sentence: 'He scratched Mr Percival's neck and gave him another piece of fish.'"
- "Look at how each version highlights different words."
- "Show fingers for the correct option."

DO:
- Wait time. On cue all show together.
- Reveal: the correct option highlights 'scratched Mr Percival's neck' and 'gave him another piece of fish' (the actions).

TEACHER NOTES:
We Do for keyword identification. Answer: the option that highlights the actions and direct objects.

WATCH FOR:
- Students who highlight only nouns. Press: 'Keywords include actions too.'""",

    31: """SAY:
- "We Do. Which sentence highlights the keywords correctly?"
- "Sentence: 'Mr Percival and Storm Boy had a very special bond that no one could get between.'"
- "Find the keywords - who, what doing, what kind."
- "Show fingers for the correct option."

DO:
- Wait time. On cue all show together.
- Reveal: the correct option highlights 'Mr Percival and Storm Boy', 'special bond', 'no one could get between'.

WATCH FOR:
- Students who highlight every word. Press: 'Keywords ONLY.'""",

    32: """SAY:
- "Read the following sentence."
- "'Mr Percival, the pelican, found a comfortable spot nearby and perched there to watch and wait until it was over.'"
- "We will work through three steps to find the keywords."

DO:
- Choral read.
- Hold up three fingers. Three steps coming.

TEACHER NOTES:
We Do walk-through. Introducing the three-step keyword routine.""",

    33: """SAY:
- "Step 1: erase the connector words."
- "Connector words like 'and', 'the', 'a', 'until' are not keywords."
- "Watch me. I cross out 'a', 'the', 'and', 'until'."
- "What is left: 'Mr Percival pelican found comfortable spot nearby perched there watch wait it was over'."

DO:
- Use a marker or pen on the slide.
- Cross out the connector words as you read.

TEACHER NOTES:
We Do step 1. Erasing connectors. Connectors are common words that hold the sentence together but do not carry main meaning.

WATCH FOR:
- Students who erase a meaningful word like 'pelican'. Cue: 'Only erase the small filler words.'""",

    34: """SAY:
- "Step 2: identify the key words or phrases."
- "Use the questions: who/what, what doing, when, where, why, how."
- "Who/what? Mr Percival, the pelican."
- "What doing? found a spot, perched, watch and wait."
- "Where? nearby."
- "Why? until it was over."

DO:
- Highlight each chunk in a different colour.
- Cold call students to identify each chunk.

TEACHER NOTES:
We Do step 2. Identifying keywords using the question prompts.

WATCH FOR:
- Students who pick out single words rather than chunks. Press: 'Keep the chunk together.'""",

    35: """SAY:
- "New sentence. Apply the three steps."
- "'In June, the storm flattened the grasses, rooted out some of the bushes that had grown on top of the sandhill for years, and blew out one of the iron sheets from the humpy.'"
- "This is a long sentence. We need keywords only."

DO:
- Choral read.
- Whiteboards: students try step 1 (erase connectors) on their own.
- Give 60 seconds, then move to the next slide.

TEACHER NOTES:
We Do for a longer sentence. Same three steps. Source: Storm Boy by Colin Thiele.""",

    36: """SAY:
- "Step 1 again. Erase the connectors."
- "Words like 'in', 'the', 'and', 'one of', 'on top of', 'for', 'from' come out."
- "Listen for what stays: 'June storm flattened grasses rooted bushes grown sandhill years blew iron sheets humpy.'"

DO:
- Cross out connector words on the slide.
- Read what is left aloud.

WATCH FOR:
- Students leaving in too many connectors. Cue: 'Small filler words go.'""",

    37: """SAY:
- "Step 2: identify the keywords using the questions."
- "When? In June."
- "What? the storm."
- "What doing? flattened the grasses, rooted bushes, blew iron sheets."
- "Where? from the humpy."

DO:
- Highlight each chunk on the slide.
- Cold call students to identify each.

TEACHER NOTES:
We Do step 2. Keywords on the longer sentence.""",

    38: """SAY:
- "Hinge question. Are these keywords correct?"
- "Sentence: 'In the early morning, Storm Boy woke up suddenly with Hide-Away's voice in his ear.'"
- "Highlighted: 'In the of early morning, Storm Boy woke up suddenly with Hide-Away's voice in his ear.'"
- "Notice what is highlighted. Is it correct?"
- "Thumbs up for yes. Thumbs down for no."

DO:
- Wait time. On cue all show together.
- Cold call to justify.

CFU CHECKPOINT:
Technique: Thumbs Up or Down
Script:
- Ask: 'Thumbs up if these keywords are correct, thumbs down if they are wrong.' Expected: thumbs down.
- Scan for: thumbs down on most.
PROCEED:
- >=80% thumbs down. Move to the next slide for the corrected version.
PIVOT:
- Most likely: students thumbs up. Misconception: not noticing the extra word 'of' in the highlight.
- Reteach: 'Look closely. One extra word is highlighted.'
- Re-check: 'Show me again.'

TEACHER NOTES:
Key hinge for keyword precision. Answer: thumbs down (the highlight includes the connector 'of').""",

    39: """SAY:
- "Now the corrected version: 'In the early morning, Storm Boy woke up suddenly with Hide-Away's voice in his ear.'"
- "Thumbs up if these keywords are correct."

DO:
- Wait time. On cue all show together.
- Reveal: thumbs up.

TEACHER NOTES:
Follow-up to slide 38. Shows the correctly highlighted version. Answer: thumbs up.""",

    40: """SAY:
- "Today we use shorthand to take notes faster."
- "These symbols save time."
- "Watch the chart: arrow means leads to / cause and effect, plus or ampersand means and, equals means same."
- "Below: b/c means because, w/o means without, w/ means with."

DO:
- Choral read each symbol and meaning.
- Hold up the chart for reference.

TEACHER NOTES:
I Do for shorthand revision. The Year 3/4 students may have used some of these in earlier units.

WATCH FOR:
- Students new to shorthand. Reassure: 'If this feels new, that is okay. We will use it together.'""",

    41: """SAY:
- "We Do. Which shorthand means cause and effect?"
- "Options: arrow, plus or ampersand, b/c, equals."
- "Show fingers 1 to 4."

DO:
- Wait time. On cue all show together.
- Cold call one student to justify.

TEACHER NOTES:
We Do shorthand recall. Answer: 1 (arrow).""",

    42: """SAY:
- "We Do. Which shorthand means and?"
- "Options: arrow, plus or ampersand, slash, equals."
- "Show fingers 1 to 4."

DO:
- Wait time. On cue all show together.
- Cold call one student to justify.

TEACHER NOTES:
We Do shorthand recall. Answer: 2 (plus or ampersand).""",

    43: """SAY:
- "We Do. Which shorthand means because?"
- "Options: cause, b/c, bcus, b/cs."
- "Show fingers 1 to 4."

DO:
- Wait time. On cue all show together.
- Cold call one student to justify.

TEACHER NOTES:
We Do shorthand recall. Answer: 2 (b/c).""",

    44: """SAY:
- "Now together. Use the shorthand chart to take notes on a sentence."
- "Sentence: 'When Storm Boy went walking along the beach or in the sanctuary, the birds were not afraid because they knew he was a friend.'"
- "Watch me: 'Storm Boy walking beach + sanctuary -> birds not afraid b/c friend.'"

DO:
- Use the shorthand chart on the slide.
- Walk through erasing connectors and applying shorthand.
- Choral read the final note.

WATCH FOR:
- Students writing 'because' instead of 'b/c'. Cue: 'Use the shorthand we learned.'""",

    45: """SAY:
- "New shorthand symbols. Read with me."
- "Up arrow means increase, growth, rise."
- "Down arrow means decrease, decline."
- "Slash means new idea."

DO:
- Read each symbol aloud.
- Choral repeat the meanings.

TEACHER NOTES:
I Do for new shorthand symbols.

WATCH FOR:
- Students confusing up and down arrow direction. Cue: 'Up means more, down means less.'""",

    46: """SAY:
- "Together. Use the new symbols on this sentence."
- "'As the summer heat went away and the winter storms became worse and worse, the birds all hid from the horrible weather.'"
- "Watch me: 'summer heat down arrow / winter storms up arrow / birds hid from weather.'"
- "Notice: down arrow for decrease, up arrow for increase, slash for new idea."

DO:
- Walk through the slide.
- Choral read the final note.

TEACHER NOTES:
We Do combining new and old shorthand.""",

    47: """SAY:
- "Now we use the full chart."
- "Sentence: 'People began to talk about Storm Boy and Mr Percival more and more.'"
- "Watch me: 'Ppl talk about Storm Boy + Mr P up arrow.'"
- "Notice 'people' is shortened to 'Ppl' and 'and' becomes plus."

DO:
- Walk through the slide.
- Choral read the final note.

WATCH FOR:
- Students who do not shorten 'people' to 'ppl'. Cue: 'Shorten common words too.'""",

    48: """SAY:
- "Together. Take notes on this sentence."
- "'When they went on their trips to Goolwa, Mr Percival did not understand what was happening.'"
- "Note version: 'Trips to Goolwa / Mr P didn't understand.'"
- "Notice: slash means new idea, names shortened."

DO:
- Whiteboards. Students attempt their own first, then reveal.
- Compare to the model.

WATCH FOR:
- Students who write the full sentence. Cue: 'Notes only.'""",

    49: """SAY:
- "Together. Take notes on this sentence."
- "'When the storm calmed down, Hide-Away and Storm Boy went outside to check the damage.'"
- "Note version: 'Weather down arrow / Hide-Away + Storm Boy outside -> check damage.'"
- "Notice: down arrow means decrease (storm calming), arrow means leads to."

DO:
- Whiteboards. Students attempt their own first, then reveal.
- Compare to the model.

WATCH FOR:
- Students who use up arrow (storm getting worse). Press: 'Storm calming is decrease.'""",

    50: """SAY:
- "Together. Take notes on this sentence."
- "'When the wind picked up, Mr Percival went inside.'"
- "Note version: 'Wind up arrow -> Mr P inside.'"
- "Notice: up arrow for picking up, arrow for cause and effect."

DO:
- Whiteboards. Students attempt their own first, then reveal.
- Compare to the model.

WATCH FOR:
- Students who write 'cause' or 'because'. Cue: 'Use the arrow shorthand.'""",

    51: """SAY:
- "Hinge question. Which option shows appropriate note taking for this passage?"
- "Passage: 'He flew around them until he saw the town, then he landed and waited patiently near the river.'"
- "Option 1: 'Mr P flew until landed -> patient near river.'"
- "Option 2: 'Mr P flew around until town / landed + waited near river.'"
- "Show fingers 1 or 2."

DO:
- Wait time. On cue all show together.
- Cold call one student to justify.

CFU CHECKPOINT:
Technique: Show Fingers 1 or 2
Script:
- Ask: 'Show me which option shows appropriate note taking.' Expected: 2.
- Scan for: 2 on most boards.
PROCEED:
- >=80% show 2. Move to the booklet task.
PIVOT:
- Most likely: students show 1. Misconception: choosing fewer notes assumes shorter is better.
- Reteach: 'Option 1 misses key info: flew where? Option 2 captures: flew around the town, landed, waited near river. More complete.'
- Re-check: 'Show me again.'

TEACHER NOTES:
Key hinge for note completeness. Answer: 2.""",

    52: """SAY:
- "Open your booklet to Lesson 13 Sentence-level writing."
- "Complete the tasks on note taking."
- "First: read each sentence. Next: erase connectors. Then: identify keywords and apply shorthand."

DO:
- Direct to the booklet.
- Set timer for 10 minutes.
- Circulate. Check first response of 3 to 5 students.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: provide the shorthand chart taped to the desk. Students focus only on identifying keywords first.
EXTENDING PROMPT:
- Task: take notes on a paragraph (not a sentence) from the novel.

TEACHER NOTES:
You Do task. Students apply the three-step process plus shorthand to sentences in their booklet.

WATCH FOR:
- Students writing full sentences. Cue: 'Notes only.'
- Students forgetting shorthand. Cue: 'Use the chart.'""",

    54: """SAY:
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
    print(f"L13 written: {stats}")


if __name__ == "__main__":
    main()
