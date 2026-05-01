"""Deep-pass teacher notes for Lesson 7 (Storm Boy - summary sentences and main idea) per teachernotes.md v2.0."""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from polish_notes_lib import polish_deck

SRC = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated/7. Storm Boy- summary sentences and main idea - with notes.pptx")
OUT = SRC.with_name(SRC.stem + " v2.pptx")

NOTES = {
    1: """TEACHER NOTES:
Title slide. Reading focus is pages 19 to 23. Writing focus is summary sentences using main-idea question prompts.""",

    2: """TEACHER NOTES:
Weekly overview. Teacher reference for the week's printing and routines. Not student-facing.""",

    3: """TEACHER NOTES:
Resources slide. Print the Session 7 Summary Sentence Enabler in advance for students who need prompts on paper. Most students work from the screen.""",

    4: """TEACHER NOTES:
Read aloud or summarise before reading. Today's chapter shows the killing of birds in the sanctuary.

SENSITIVITY ADVISORY:
- What it is: violence against animals; First Nations content in surrounding resources.
- Framing language: 'This part of the story is sad. The author wants us to feel for the birds.'
- Watch for: any student visibly upset by the bird-killing scene.
- Protocol: pause if a student is upset, name the emotion, allow students to step out briefly, follow up at recess.""",

    5: """TEACHER NOTES:
Teacher orientation only, not for students.""",

    6: """TEACHER NOTES:
Teacher reference for I Do, We Do, You Do badges and support and extension icons. Not student-facing.""",

    7: """TEACHER NOTES:
Teacher reference for the response routines used through the deck. Not student-facing.""",

    8: """TEACHER NOTES:
Teacher reference for sentence-element colour coding. Not student-facing.""",

    9: """SAY:
- "Read the learning intention with me."
- "We are reading more of Storm Boy and learning to write a summary sentence."
- "A summary sentence captures the main idea of a part of the text in just one sentence."
- "Ask: what might a summary leave out? Expected: extra details, things that aren't important."
- "If summarising feels new, that is okay."

DO:
- Choral read the LI, then track each success criterion with your finger.
- Cold call one student to say SC3 in their own words.

TEACHER NOTES:
SC3 is the writing target the You Do assesses. Refer back to it during the writing block.

WATCH FOR:
- Students who cannot repeat SC3. Sit them next to a strong writer for the You Do.""",

    10: """SAY:
- "Quick check."
- "You need pencil, booklet, novel, mini whiteboard and texta."
- "Thumbs up when you have all five."

DO:
- Scan the room. Wait for all thumbs.

TEACHER NOTES:
Don't start the read aloud until materials are out.

WATCH FOR:
- Students with no novel. Pair with a buddy.""",

    11: """TEACHER NOTES:
Section divider into the read-aloud section. Today's reading mode is teacher choice.""",

    12: """SAY:
- "Open Storm Boy to page 19."
- "I will read aloud. Follow with your eyes or finger."
- "We stop at 'came to live with Storm Boy.'"
- "Listen for the big ideas: Storm Boy finds the destruction, rescues three baby pelicans, nurses them, names them."
- "Ask: how did Storm Boy feel when he found the destruction? Expected: sad, angry, scared."

DO:
- Read pages 19 to 23 with a softer voice for the sad parts.
- Pre-mark page 20 'a faint rustling and crying' for the vocabulary slide.
- Pre-mark page 19-20 for the personification line you need later.

TEACHER NOTES:
The destruction and the rescue are emotionally heavy. Read them slowly and let students feel the moment.

WATCH FOR:
- Students losing focus during the destruction scene. Pause: 'It's a sad page. The next part is hope.'""",

    13: """TEACHER NOTES:
Section divider. Today's vocabulary words are faint and destruction.""",

    14: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is faint."
- "Faint means weak. Hardly noticeable."
- "Faint is an adjective."
- "Say it with me: faint."
- "Listen to it in our story: 'Storm Boy heard a faint rustling and crying.'"

DO:
- Say faint together two times: speaking volume, then a whisper.
- Cup your hand to your ear to model listening for a faint sound.

TEACHER NOTES:
The word matters because the rustling Storm Boy hears is what saves the baby pelicans.

WATCH FOR:
- Students who only know faint as the verb (passing out). Acknowledge it, then refocus on the adjective meaning.""",

    15: """SAY:
- "If the animal makes a faint sound, say faint."
- "Three, two, one."

DO:
- Listen for chorus.
- Reveal: faint.

WATCH FOR:
- Quiet response. Cue students by name on the next image.""",

    16: """SAY:
- "Next image. Faint? Say it if yes."
- "Three, two, one."

DO:
- Listen.
- Reveal: not faint.

WATCH FOR:
- Students who chorus faint regardless. They are pattern-matching. Anchor: silence is a valid answer.""",

    17: """SAY:
- "Round three. Faint? Say it if yes."
- "Three, two, one."

DO:
- Listen.
- Reveal: not faint.

WATCH FOR:
- Slow chorus. Students unsure. Acknowledge and move on.""",

    18: """SAY:
- "Last one. Faint? Say it if yes."
- "Recap: faint sounds are quiet and hard to notice."

DO:
- Listen.
- Reveal: faint.

WATCH FOR:
- Students who got all four right. Call on them in the CFU next.""",

    19: """SAY:
- "Which word is most similar to a faint sound? Loud, soft, silent, squeaky."
- "Whiteboards: write your answer."

DO:
- Wait. Show me. Scan.
- Reveal: soft.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Show me your answer. One word.' Expected: soft.
- Scan for: soft on most boards.
PROCEED:
- >=80% show soft. Confirm: 'Soft. Quiet but you can still hear it.'
PIVOT:
- Most likely: students pick silent thinking faint = no sound.
- Reteach: 'Silent means no sound. Faint means there is a sound, just very small.'
- Re-check: 'Silent or soft - which matches faint?'

WATCH FOR:
- Students who pick squeaky. They confused pitch with volume.""",

    20: """SAY:
- "The sound would or wouldn't be faint if you covered your ears."
- "Whiteboards: would or wouldn't."

DO:
- Wait. Show me.
- Reveal: would.

WATCH FOR:
- Students who pick wouldn't. They missed that covering ears makes sounds quieter.""",

    21: """SAY:
- "Round two. Would or wouldn't be faint while watching fireworks?"
- "Whiteboards."

DO:
- Wait. Show me.
- Reveal: wouldn't. Fireworks are loud.

WATCH FOR:
- Students who pick would. Check whether they read the question carefully.""",

    22: """SAY:
- "Which would make a faint sound? Mouse, turning a page, radio, dropping a pencil on carpet, filling the sink, clock ticking."
- "Whiteboards: write the faint ones."

DO:
- Wait twenty seconds.
- Show me.
- Discuss: faint = mouse, turning a page, dropping a pencil on carpet, clock ticking. Loud = radio, sink filling.

WATCH FOR:
- Disagreement on dropping a pencil. Acknowledge: 'It depends on the surface.'""",

    23: """SOURCES:
Macquarie Dictionary, 2024.

SAY:
- "The word is destruction."
- "Destruction is the act of destroying or being destroyed. Serious damage."
- "Destruction is a noun."
- "Say it with me: destruction."
- "Listen to it in our story: 'Storm Boy looked at the destruction caused by the men.'"

DO:
- Say destruction together two times.
- Point to the image and name what you see.

TEACHER NOTES:
This word captures what Storm Boy finds at the sanctuary.

WATCH FOR:
- Students who think destruction must be huge (war, bushfire). Emphasise it can be smaller damage too.""",

    24: """SAY:
- "Look at the image."
- "What destruction did Storm Boy find in the story?"
- "Turn to your partner. Describe what was destroyed."

DO:
- One minute partner talk.
- Cold call two pairs.
- Confirm: 'Storm Boy found pelican nests destroyed and two adult pelicans killed.'

TEACHER NOTES:
Link the vocabulary back to the read aloud.

WATCH FOR:
- Students who cannot recall what was destroyed. Quick re-read of page 20.""",

    25: """SAY:
- "Use destruction in a sentence about the picture."
- "Whiteboards. One sentence."

DO:
- Wait two minutes. Show me.
- Reveal model: 'The bushfire caused a huge amount of destruction.'
- Read two student examples.

WATCH FOR:
- Students using destruction as a verb ('The fire destruction the trees'). Quick reteach: 'Destruction is a noun. The verb is destroy.'""",

    26: """SAY:
- "Which sentence uses destruction correctly?"
- "A: Pollution can lead to destruction of coral reefs."
- "B: The pollution was destruction on the coral reef."
- "Whiteboards: A or B."

DO:
- Wait. Show me.
- Reveal: A.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Show me A or B.' Expected: A.
- Scan for: A on most boards.
PROCEED:
- >=80% show A. Confirm: 'A uses destruction as the result. B uses it as a verb.'
PIVOT:
- Most likely: students pick B because both mention pollution and destruction.
- Reteach: 'Read each out loud. Does it sound like real English?' B sounds wrong.
- Re-check: 'Show me again. A or B.'

WATCH FOR:
- Students who pick B. Need more practice with verb vs noun forms.""",

    27: """SAY:
- "Sort the words on your whiteboard."
- "One side: similar to destruction. Other side: opposite."
- "Words: wreck, protect, care, destroy."

DO:
- Wait twenty seconds.
- Show me.
- Reveal: similar = wreck, destroy. Opposite = protect, care.

TEACHER NOTES:
Reinforces meaning by contrast. Link 'protect' to the sanctuary - it was meant to be protected.

WATCH FOR:
- Students who put wreck as opposite. They may know it as the noun (shipwreck), not the verb.""",

    28: """SAY:
- "Destruction or not?"
- "Spilled food, ocean pollution, rubbish at the dump, graffiti on a building."
- "Write the destruction ones on your whiteboard."

DO:
- Wait. Show me. Scan.
- Discuss: pollution and graffiti = destruction. Rubbish at dump is where it belongs. Spilled food is mess unless it ruins something.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Which are destruction? Write them.' Expected: pollution in ocean and graffiti on building.
- Scan for: those two on most boards.
PROCEED:
- >=80% show those two. Confirm: 'Both cause real damage hard to undo.'
PIVOT:
- Most likely: students label any mess as destruction.
- Reteach: 'Destruction means real damage that is hard to fix. A spilled drink wipes up.'
- Re-check: 'Same list. Which are destruction?'

WATCH FOR:
- Students who pick all four. They lost the line between mess and destruction.""",

    29: """SAY:
- "In your booklet, write the meaning of faint in your own words."
- "Then write the meaning of destruction in your own words."
- "Five minutes."

DO:
- Open booklets.
- Circulate.
- Check the first definition.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: sentence frame: 'Faint means ___. Destruction means ___.'
- Extra Notes: copy from the vocab slide if needed.
EXTENDING PROMPT:
- Task: write one sentence using each word that connects to Storm Boy.
- Extra Notes: model: 'Storm Boy heard a faint cry and ran towards the destruction.'

WATCH FOR:
- Students who cannot define in their own words. They may need another example from you before continuing.""",

    30: """TEACHER NOTES:
Section divider into the writing focus block on summary sentences.""",

    31: """SAY:
- "This section uses three writing tricks: simile, personification, metaphor."
- "Simile: 'like a lot of important old men' - compares pelicans to old men using LIKE."
- "Personification: 'sitting up as if he owns the place' - gives Mr Percival a human attitude."
- "Metaphor: 'the ibises cut the air into strips' - wings aren't actually cutting; the author paints a picture."

DO:
- Read each example aloud.
- Mime each: stand tall for pelicans, point for Mr Percival, slice the air for ibises.
- Cold call: 'Which one uses LIKE?' Expected: simile.

TEACHER NOTES:
Light touch under five minutes. The lesson focus is summary sentences, not literary devices.

WATCH FOR:
- Students confusing simile and metaphor. Anchor: 'Simile uses like or as. Metaphor does not.'""",

    32: """SAY:
- "Watch this first."
- "A summary is a short version of a story or event with the most important parts only."
- "It does not include every detail. Just the main ideas."
- "Ask: what does a summary leave out? Expected: extra details."

DO:
- Point to the definition as you read.
- Cold call: 'What does a summary keep?' Expected: the main idea.

TEACHER NOTES:
Key idea: a summary keeps the main idea, not every detail. This is the I Do anchor.

WATCH FOR:
- Students who think a summary is a retelling of everything. Clarify: 'Just the most important point.'""",

    33: """SAY:
- "Let's try one together. This is from page 19 and 20."
- "Read with me: the long passage about Storm Boy walking with the birds."
- "We will find the main idea, then write one summary sentence."

DO:
- Choral read.
- Then teacher reads with expression.
- Ask: 'In one sentence - what is this about?' Take two responses.

TEACHER NOTES:
Lead-in to the I Do. Do not write the summary yet. The next slide builds it.

WATCH FOR:
- Students who try to retell every detail. Redirect: 'Just the main idea.'""",

    34: """SAY:
- "Watch this first."
- "Now we use question prompts to find the main idea."
- "Who? The pelicans. What did they do? Stood still and greeted him."
- "When? When Storm Boy walked. Where? Along the beach."
- "Why? Because they were not afraid. How? Drily."

DO:
- Reveal each label one at a time.
- Cold call to confirm each piece.
- Highlight: 'These pieces become our summary sentence next.'

TEACHER NOTES:
Move slowly so students see how each prompt picks out one piece.

WATCH FOR:
- Students confused about who/what. They may write 'Storm Boy' for who. Clarify: 'Who or what is the doer of this part.'""",

    35: """SAY:
- "Now we put it into one sentence."
- "Watch me build it: 'When Storm Boy walked along the beach, the pelicans stood still and greeted him drily because they were not afraid.'"
- "Notice the order: when, who, what, how, why."

DO:
- Build the sentence on the board piece by piece.
- Cold call: 'Which question word does our sentence start with?' Expected: when.
- Note the comma after the when starter.

TEACHER NOTES:
The I Do model. Show exactly how the prompts turn into a summary.

WATCH FOR:
- Students who think the summary should include every detail. Re-read original and summary side by side to show the difference.""",

    36: """SAY:
- "Now we try a different extract."
- "Read with me: the passage about the men entering the sanctuary."
- "Who? Three or four young men. What? Killed two nesting pelicans."
- "When? One morning. Where? In the sanctuary."

DO:
- Choral read.
- Reveal each label as students answer.
- Cold call to confirm each piece.

TEACHER NOTES:
We Do. Don't write the summary yet. That is the next slide.

WATCH FOR:
- Students who say Storm Boy is the who. He is not the actor here. Anchor: 'Who is doing the action in this part?'""",

    37: """SAY:
- "Now you put it together."
- "Write a summary sentence using these pieces: three or four young men, killed two nesting pelicans, in the sanctuary, one night."
- "One sentence. Thirty seconds."

DO:
- Wait. Show me.
- Reveal model: 'One night, three or four young men killed two nesting pelicans in the sanctuary.'
- Compare student answers.

CFU CHECKPOINT:
Technique: Show Me Boards
Script:
- Ask: 'Write the main idea of this passage in one sentence. Thirty seconds.' Expected: a single sentence with who, what, where, when.
- Scan for: a single sentence with all four pieces.
PROCEED:
- >=80% write a clear summary. Move into the You Do.
PIVOT:
- Most likely: students write more than one sentence.
- Reteach: 'One sentence only. Cut anything that is not needed.'
- Re-check: 'Same passage. One sentence on your board.'

WATCH FOR:
- Students writing two or three sentences. They have not grasped the one-sentence rule yet.""",

    38: """SAY:
- "Your turn to write summary sentences on your own."
- "First: read the extract carefully."
- "Next: answer the prompts on the side - who, what, when, where, why."
- "Then: write your summary sentence. One per extract."
- "Try to start with WHEN if you can."

DO:
- Set fifteen minutes.
- Distribute or open booklets.
- Circulate. Check the prompts before the sentence.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: print the Session 7 Summary Sentence Enabler. Prompts pre-laid out.
- Extra Notes: do one extract together with the student first.
EXTENDING PROMPT:
- Task: print the Session 7 Summary Sentence Extender. Two more extracts.
- Extra Notes: ask them to start two summaries with different question words.

TEACHER NOTES:
Main You Do for SC3. Track which students write a clear single sentence.

WATCH FOR:
- Students who copy the extract into the summary. Redirect: 'Your own words.'
- Students who write three sentences. Cut to one.""",

    39: """SAY:
- "Read the extract on pages 20-21 with me."
- "It is the moment Storm Boy finds the three baby pelicans alive."
- "Write the key information: who, what, where, how."

DO:
- Choral read once.
- Read again at normal pace while students take notes.
- Set five minutes.

TEACHER NOTES:
First You Do extract. Take notes first, then write the summary.

WATCH FOR:
- Students who jump to the summary. Cue: 'Use the prompts first.'""",

    40: """SAY:
- "Fill in the table on your worksheet."
- "Who/what is the subject. What did they do. Where. How."
- "Five minutes."

DO:
- Circulate.
- Check the table BEFORE students move to the summary.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: pre-print the Session 7 Enabler with prompts laid out for this extract.
- Extra Notes: suggest one piece of evidence per prompt.
EXTENDING PROMPT:
- Task: students who finish add why and when, with extra evidence.
- Extra Notes: extension for early finishers.

TEACHER NOTES:
Note-taking step before the summary. Skipping it is the most common cause of weak summaries.

WATCH FOR:
- Students who write 'Storm Boy' for everything. Each prompt has a different answer.""",

    41: """SAY:
- "Now use your notes to write the summary."
- "One sentence. Capture the main idea: Storm Boy carefully picks up the baby pelicans and takes them home."

DO:
- Set five minutes.
- Circulate. Read each summary as students finish.
- Pull two strong examples to share.

TEACHER NOTES:
Sample: 'Storm Boy carefully picked up the three tiny pelicans under the broken nests and hurried back to Hide-Away.' Look for one sentence with who, what, how, where.

WATCH FOR:
- Students who include the destruction detail but not the rescue. The rescue is the main idea here.""",

    42: """TEACHER NOTES:
The next slides are hidden by default. Unhide them only if students need extra scaffolded practice. Includes one more extract from page 22.""",

    43: """SAY:
- "Read the extract on page 22 with me."
- "It describes the baby pelicans three days after the rescue."
- "Ask: what were the pelicans doing after three days? Expected: had their beaks open hungrily."

DO:
- Choral read once.
- Read again at normal pace.
- Cold call one student to answer.

TEACHER NOTES:
Optional extra-practice extract for students who need more guided reps.

WATCH FOR:
- Students who cannot separate the brothers from the rescued one. Re-read the relevant sentence.""",

    44: """SAY:
- "Fill in the table on your whiteboard."
- "Who/what: the three baby pelicans. What: had their beaks open. When: after three days."
- "How: hungrily. Why: too young to feed themselves."

DO:
- Reveal each piece as students complete it.
- Cold call to confirm each.

TEACHER NOTES:
Guided table fill. Models the note-taking step.

WATCH FOR:
- Students unsure how to fit 'after three days'. Show that 'after three days' is when.""",

    45: """SAY:
- "Write the summary on your whiteboard using the notes."
- "One sentence."

DO:
- Set two minutes. Show me. Scan.
- Reveal model: 'After three days, the three baby pelicans had their beaks open hungrily because they were too young to feed themselves.'

TEACHER NOTES:
Final hidden-extension example. Students should now be ready for any extract.

WATCH FOR:
- Students writing multiple sentences. Cue: 'Combine into one.'""",

    46: """SAY:
- "Turn to your partner."
- "Read your best summary sentence."
- "Your partner says whether it captures the main idea or includes too much detail."
- "Then swap."

DO:
- One minute each way.
- Thumbs self-assessment for each SC.

TEACHER NOTES:
Use the SC2 and SC3 thumbs to plan tomorrow. If sideways or down dominates SC3, plan a re-teach.

WATCH FOR:
- Students whose partner says 'too much detail'. That is a writing-revision target.""",

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
    print(f"L07 written: {stats}")


if __name__ == "__main__":
    main()
