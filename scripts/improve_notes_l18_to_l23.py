"""Lean deep-pass for L18, L19, L20, L22, L23 per teachernotes.md v2.0.

These five decks share Ochre's Stage-tagged notes structure. Existing notes are
already strong - polish_lib preserves them. This script overrides only the
title, LI/SC, and final closing slide on each deck.
"""
from pathlib import Path
import sys
sys.path.insert(0, str(Path(__file__).parent))
from polish_notes_lib import polish_deck


SRC_DIR = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated")


CLOSING = """SAY:
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
- Patterns where SC2 or SC3 is mostly thumbs down. Flag for reteach."""


LESSONS = [
    {
        "n": 18,
        "filename": "18. literature_presentation Note taking 4En06V Storm Boy - Information report - with notes.pptx",
        "slides": 55,
        "title_note": "Title slide. Reading focus is pages 67 to 69. Writing focus is note-taking with keywords, phrases, abbreviations and symbols (KPAS).",
        "li_note": """SAY:
- "Read the learning intention with me."
- "Today we read pages 67 to 69 of Storm Boy and take notes that capture the main ideas of a text."
- "Read each success criterion together."
- "If KPAS feels new, that is okay. We will model it together."

DO:
- Choral read the LI, then track each SC with your finger.
- Cold call one student to repeat SC3 in their own words.

TEACHER NOTES:
SC2 and SC3 are the writing targets the booklet task assesses.

WATCH FOR:
- Students unsure what 'keywords' are. Quick gloss: 'The most important words in a sentence.'""",
    },
    {
        "n": 19,
        "filename": "19. literature_presentation Plan a body paragraph for an information report 4En06V Storm Boy - Information report - with notes.pptx",
        "slides": 50,
        "title_note": "Title slide. Today's focus: plan body paragraph 3 of the Coorong information report (flora and fauna) using a Single Paragraph Outline.",
        "li_note": """SAY:
- "Read the learning intention with me."
- "Today we plan body paragraph 3 of the Coorong report - the flora and fauna paragraph."
- "Flora means plants. Fauna means animals."
- "If the SPO format feels new again, that is okay. We have used it in earlier lessons."

DO:
- Choral read the LI, then track each SC with your finger.
- Cold call one student to repeat SC3 in their own words.
- Briefly remind students of the previous SPO they planned in Lesson 14.

TEACHER NOTES:
SC3 is what the booklet task assesses.

WATCH FOR:
- Students unsure of flora and fauna. Anchor: 'Flora plants. Fauna animals.'""",
    },
    {
        "n": 20,
        "filename": "20. literature_presentation Write a body paragraph for an information report 4En06V Storm Boy - Information report - with notes.pptx",
        "slides": 47,
        "title_note": "Title slide. Today's focus: turn the SPO from Lesson 19 into a full body paragraph (flora and fauna) using information report language features.",
        "li_note": """SAY:
- "Read the learning intention with me."
- "Today is writing day for body paragraph 3."
- "We turn our SPO from yesterday into a full paragraph."
- "Read each success criterion."

DO:
- Choral read the LI, then track each SC with your finger.
- Brief check: 'Hold up your SPO from Lesson 19.' Scan.
- Cold call one student to repeat SC3 in their own words.

TEACHER NOTES:
SC1, SC2 and SC3 are all writing targets. The booklet task is the You Do.

WATCH FOR:
- Students without their SPO from Lesson 19. Provide a model SPO as a backup.""",
    },
    {
        "n": 22,
        "filename": "22. literature_presentation Single paragraph outline (SPO) to summarise a text 4En06V Storm Boy - Information report - with notes.pptx",
        "slides": 55,
        "title_note": "Title slide. Reading focus is pages 72 to 73. Writing focus is using a Single Paragraph Outline (SPO) to summarise main ideas.",
        "li_note": """SAY:
- "Read the learning intention with me."
- "Today we read pages 72 to 73 of Storm Boy and use a Single Paragraph Outline to summarise the main ideas."
- "Read each success criterion together."
- "If SPO feels new, that is okay. We have used it before and will model it again."

DO:
- Choral read the LI, then track each SC with your finger.
- Cold call one student to repeat SC3 in their own words.

TEACHER NOTES:
SC3 is the writing target the booklet task assesses.

WATCH FOR:
- Students unsure what an SPO is. Quick gloss: 'Topic sentence, three supporting details, concluding sentence.'""",
    },
    {
        "n": 23,
        "filename": "23. literature_presentation Note taking 4En06V Storm Boy - Information report - with notes.pptx",
        "slides": 54,
        "title_note": "Title slide. Reading focus is pages 73 to 76. Writing focus is note-taking with keywords, phrases, abbreviations and symbols (KPAS).",
        "li_note": """SAY:
- "Read the learning intention with me."
- "Today we read pages 73 to 76 of Storm Boy and take notes that capture the main ideas of a text."
- "Read each success criterion together."
- "If KPAS feels new again, that is okay. We will model it together."

DO:
- Choral read the LI, then track each SC with your finger.
- Cold call one student to repeat SC3 in their own words.

TEACHER NOTES:
SC2 and SC3 are the writing targets the booklet task assesses.

WATCH FOR:
- Students unsure what 'keywords' are. Quick gloss: 'The most important words in a sentence.'""",
    },
]


def main():
    for lesson in LESSONS:
        src = SRC_DIR / lesson["filename"]
        out = src.with_name(src.stem + " v2.pptx")
        # LI/SC is slide 8 in these decks; closing is the last slide.
        notes = {
            1: f"TEACHER NOTES:\n{lesson['title_note']}",
            8: lesson["li_note"],
            lesson["slides"]: CLOSING,
        }
        stats = polish_deck(src, out, overrides=notes)
        print(f"L{lesson['n']:02d} written: {stats}")


if __name__ == "__main__":
    main()
