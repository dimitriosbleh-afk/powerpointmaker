"""Deep-pass teacher notes for Lesson 17 (SPO to summarise a text) per teachernotes.md v2.0.

Lean override: only LI/SC and key new teaching slides; polish_lib preserves existing strong notes.
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from polish_notes_lib import polish_deck

SRC = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated/17. literature_presentation Single paragraph outline (SPO) to summarise a text 4En06V Storm Boy - Information report - with notes.pptx")
OUT = SRC.with_name(SRC.stem + " v2.pptx")

NOTES = {
    1: """TEACHER NOTES:
Title slide. Reading focus is pages 62 to 66. Writing focus is using a Single Paragraph Outline (SPO) to summarise main ideas of a section of the text.""",

    7: """SAY:
- "These are our learning objectives across the unit."
- "We are reading more of Storm Boy, learning two new words, and writing a Single Paragraph Outline - we call it an SPO - to summarise part of the story."
- "If SPOs feel new, that is okay. We will build it together."

DO:
- Point to each SC.
- Tap the SPO criterion twice so students hold it as today's writing focus.

TEACHER NOTES:
The SPO criterion is what the booklet task assesses.

WATCH FOR:
- Students who cannot say the SPO goal back. They will need extra modelling later in the lesson.""",

    8: """SAY:
- "Read the learning intention with me."
- "Today we read pages 62 to 66 of Storm Boy and use a Single Paragraph Outline to summarise the main ideas."
- "Read each success criterion together."
- "If SPO feels new, that is okay. We will model it step by step."

DO:
- Choral read the LI, then track each SC with your finger.
- Cold call one student to repeat SC3 in their own words.

TEACHER NOTES:
SC3 is the writing target the booklet task assesses.

WATCH FOR:
- Students unsure what an SPO is. Quick gloss: 'A paragraph plan with a topic sentence, three supporting details and a concluding sentence.'""",

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
    print(f"L17 written: {stats}")


if __name__ == "__main__":
    main()
