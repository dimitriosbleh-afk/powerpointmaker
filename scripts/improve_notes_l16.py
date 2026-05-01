"""Deep-pass teacher notes for Lesson 16 (Sentence Combining - Compound Sentences) per teachernotes.md v2.0.

L16 has 76 slides with already-strong existing notes (Stage tags, full structure).
Lean override: only LI/SC, key teaching slides and closing - polish_lib preserves the rest.
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from polish_notes_lib import polish_deck

SRC = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated/16. literature_presentation Sentence Combining - Compound Sentences 4En06V Storm Boy - Information report - with notes.pptx")
OUT = SRC.with_name(SRC.stem + " v2.pptx")

NOTES = {
    1: """TEACHER NOTES:
Title slide. Reading focus is pages 57 to 62. Writing focus is combining two simple sentences into a compound sentence using a comma and a FANBOYS coordinating conjunction.""",

    7: """SAY:
- "These are our learning objectives across the unit."
- "We are reading more of Storm Boy, learning three new words, and learning to combine sentences using FANBOYS conjunctions."
- "If joining sentences feels new, that is okay. We will build it together."

DO:
- Point to each SC.
- Tap the FANBOYS criterion twice so students hold it as today's writing focus.

TEACHER NOTES:
The FANBOYS criterion is what the booklet task assesses. Comprehension and vocab feed into the writing.

WATCH FOR:
- Students who cannot say the FANBOYS goal back in their own words. Flag for extra modelling later.""",

    8: """SAY:
- "Read the learning intention with me."
- "Today we read pages 57 to 62 of Storm Boy and learn to combine two simple sentences into a compound sentence using a FANBOYS conjunction."
- "Read each success criterion together."
- "If FANBOYS or 'compound sentence' feels new, that is okay. We will build it together."

DO:
- Choral read the LI, then track each SC with your finger.
- Cold call one student to repeat SC3 in their own words.

TEACHER NOTES:
SC3 is the writing target the booklet task assesses (combining with a comma and a FANBOYS conjunction).

WATCH FOR:
- Students unsure what 'compound sentence' means. Reassure: 'We will model it step by step.'""",

    49: """SAY:
- "Watch this first."
- "A compound sentence is two simple sentences joined together."
- "The recipe: independent clause, comma, FANBOYS word, independent clause."
- "I need to check that I have two complete sentences before I join them."
- "Each side could stand alone."

DO:
- Read the structure twice, pointing to each part.
- Write a live example: 'The waves were huge, and the boat was sinking.'
- Underline both clauses, circle the comma and the FANBOYS word.

TEACHER NOTES:
Core teaching point. Keep your live example up for the rest of the lesson.

WATCH FOR:
- Students who write compound sentences without a comma. The comma plus FANBOYS is the rule.""",

    50: """SAY:
- "FANBOYS is a memory trick. Each letter stands for a coordinating conjunction."
- "F-A-N-B-O-Y-S. For, And, Nor, But, Or, Yet, So."
- "Ask: say it with me. Expected: 'For, and, nor, but, or, yet, so.'"

DO:
- Lead a choral chant of FANBOYS twice.
- Write the seven words vertically on the board.

TEACHER NOTES:
Display the FANBOYS list nearby for the rest of the lesson.

WATCH FOR:
- Students confusing 'for' the conjunction with 'for' the preposition. Clarify quickly only if it surfaces.""",

    58: """SAY:
- "Watch this first. I will combine these two sentences."
- "Step one: I find the noun that repeats. The storm appears in both."
- "Step two: I replace one noun with a pronoun. The storm becomes it."
- "Step three: I join with a coordinating conjunction. So fits because it shows cause and effect."
- "Step four: I add a comma before the FANBOYS word."
- "Final sentence: 'The storm was too wild, so it made it hard for the men to swim.'"

DO:
- Walk through the four steps slowly, pointing at each.
- Underline the comma and FANBOYS word in the final sentence.

TEACHER NOTES:
This is the model for the four steps. Keep them visible for every We Do that follows.

MISCONCEPTIONS:
- Misconception: students join with FANBOYS but no comma.
  Why: shorter sentences feel like they don't need the comma.
  Impact: incorrect compound punctuation.
  Quick correction: 'Comma plus FANBOYS, always together.'

WATCH FOR:
- Students who copy without watching the steps. Emphasise: 'Watch first.'""",

    74: """SAY:
- "Open your booklet to Lesson 16 Sentence Level Writing."
- "Complete the sentence combining tasks."
- "Use the four steps every time: find the repeated noun, replace one with a pronoun, join with a FANBOYS word, add a comma."

DO:
- Set the timer.
- Circulate. Check the first 2 to 3 combinations per student.
- Pull a small group for re-modelling.

ENABLING & EXTENDING:
ENABLING PROMPT:
- Task: provide the FANBOYS word for the first three combinations so students focus on the comma and pronoun.
- Extra Notes: pair with a peer.
EXTENDING PROMPT:
- Task: write three of your own pairs of simple sentences from today's reading, then combine each one. Use a different FANBOYS word each time.

TEACHER NOTES:
The booklet is the You Do for the FANBOYS objective. Use student work to plan tomorrow's review.

WATCH FOR:
- Missing commas before FANBOYS words.
- Pronoun mismatches.
- Students using only 'and' or 'but'.""",

    76: """SAY:
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
    print(f"L16 written: {stats}")


if __name__ == "__main__":
    main()
