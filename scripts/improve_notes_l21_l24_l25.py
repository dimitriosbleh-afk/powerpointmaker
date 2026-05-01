"""Lean deep-pass for L21, L24, L25 per teachernotes.md v2.0.

Existing notes on these decks are strong. Override only title, LI/SC, closing.
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
        "n": 21,
        "filename": "21. literature_presentation Punctuate direct speech 4En06V Storm Boy - Information report - with notes.pptx",
        "slides": 65,
        "title_note": "Title slide. Reading focus is pages 69 to 72. Writing focus is punctuating direct speech with speech marks, a comma and a capital letter.",
        "li_note": """SAY:
- "Read the learning intention with me."
- "Today we read pages 69 to 72 of Storm Boy and learn to punctuate direct speech accurately."
- "Read each success criterion together."
- "If speech marks feel new, that is okay. We will build it together."

DO:
- Choral read the LI, then track each SC with your finger.
- Cold call one student to repeat SC3 in their own words.

TEACHER NOTES:
SC3 is what the booklet task assesses (speech marks, comma, capital letter rule).

WATCH FOR:
- Students who confuse speech marks with quotation marks for emphasis. Anchor: 'Speech marks go around the words a person says.'""",
    },
    {
        "n": 24,
        "filename": "24. literature_presentation Plan a concluding paragraph for an information report 4En06V Storm Boy - Information report - with notes.pptx",
        "slides": 39,
        "title_note": "Title slide. Today's focus: plan the concluding paragraph for the Coorong information report using the TSG formula (Thesis statement, Summary of body paragraphs, General statement).",
        "li_note": """SAY:
- "Read the learning intention with me."
- "Today we plan the concluding paragraph for our Coorong report using the TSG formula."
- "T is for thesis statement. S is for summary of body paragraphs. G is for general statement."
- "If TSG feels new, that is okay. We have used GST for the introduction; this is its mirror."

DO:
- Choral read the LI, then track each SC with your finger.
- Cold call one student to repeat SC3 in their own words.

TEACHER NOTES:
SC3 is what the booklet task assesses. The TSG formula closes the report; GST opened it.

WATCH FOR:
- Students who try to add new facts in the conclusion. Anchor: 'Conclusions sum up. They do not add new information.'""",
    },
    {
        "n": 25,
        "filename": "25. literature_presentation Write a concluding paragraph for an information report 4En06V Storm Boy - Information report - with notes.pptx",
        "slides": 38,
        "title_note": "Title slide. Today's focus: turn the TSG plan from Lesson 24 into a full concluding paragraph for the Coorong information report.",
        "li_note": """SAY:
- "Read the learning intention with me."
- "Today is the final writing day. We turn our TSG plan into a full concluding paragraph."
- "Read each success criterion together."

DO:
- Choral read the LI, then track each SC with your finger.
- Brief check: 'Hold up your TSG plan from Lesson 24.' Scan.
- Cold call one student to repeat SC3 in their own words.

TEACHER NOTES:
This is the final writing lesson of the unit. SC1, SC2 and SC3 are all writing targets.

WATCH FOR:
- Students without their TSG plan from Lesson 24. Provide a model TSG as a backup.
- Students who repeat the introduction word for word. Press: 'Echo the idea, change the words.'""",
    },
]


def main():
    for lesson in LESSONS:
        src = SRC_DIR / lesson["filename"]
        out = src.with_name(src.stem + " v2.pptx")
        notes = {
            1: f"TEACHER NOTES:\n{lesson['title_note']}",
            8: lesson["li_note"],
            lesson["slides"]: CLOSING,
        }
        stats = polish_deck(src, out, overrides=notes)
        print(f"L{lesson['n']:02d} written: {stats}")


if __name__ == "__main__":
    main()
