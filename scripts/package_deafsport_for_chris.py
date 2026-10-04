"""Package the Deaf Sports term into two upload-ready folders for Chris.

    python scripts/package_deafsport_for_chris.py

Writes output/Auslan Term 4 for Chris/Challenge and .../Enrichment, each holding
everything that cohort needs for the term: the slides (all eight lessons plus
the Week 9 and 10 reviews), the printables, the student journal template, and the
planning documents. Rebuild the decks first (build_unit.py and
build_deafsport_journal.js); this only copies. The printables folder keeps the
name Resources because the deck's links point at it.
"""
import os
import shutil

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "output")
DEST = os.path.join(OUT, "Auslan Term 4 for Chris")
UNIT = os.path.join(OUT, "Deaf_Sports_Term4_2026")
LAYER = os.path.join(UNIT, "Teacher Layer")

# The landscape pick-up documents ship as Word only: LibreOffice on the Mac
# paginates them differently from Word (a lesson page spills onto two), and
# Drive opens the Word file directly.
SHARED_PLANNING = [
    (UNIT, "Deaf Sports In Australia Unit.docx"),
    (UNIT, "Deaf Sports In Australia Unit.pdf"),
    (LAYER, "Deaf Sports At A Glance.docx"),
    (LAYER, "Deaf Sports Lesson Pages.docx"),
    (LAYER, "Deaf Sports CRT Review Pack.docx"),
]

COHORTS = {
    "Challenge": {
        "unit": "Deaf_Sports_Challenge_Unit",
        "deck": "Deaf Sports in Australia Challenge.pptx",
        "journal": "Deaf Sports Challenge Student Journal Template.pptx",
        "tracker": "Deaf Sports Challenge Evidence Tracker.xlsx",
    },
    "Enrichment": {
        "unit": "Deaf_Sports_Enrichment_Unit",
        "deck": "Deaf Sports in Australia Enrichment.pptx",
        "journal": "Deaf Sports Enrichment Student Journal Template.pptx",
        "tracker": "Deaf Sports Enrichment Evidence Tracker.xlsx",
    },
}


def copy(src, dst):
    if not os.path.exists(src):
        raise SystemExit("Missing: " + src)
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(src, dst)


def main():
    if os.path.isdir(DEST):
        shutil.rmtree(DEST)
    for name, c in COHORTS.items():
        base = os.path.join(DEST, name)
        unit = os.path.join(OUT, c["unit"])
        copy(os.path.join(unit, c["deck"]), os.path.join(base, "Deaf Sports " + name + " Slides.pptx"))
        copy(os.path.join(unit, c["journal"]), os.path.join(base, c["journal"]))
        shutil.copytree(os.path.join(unit, "Resources"), os.path.join(base, "Resources"))
        planning = os.path.join(base, "Planning")
        for folder, f in SHARED_PLANNING:
            copy(os.path.join(folder, f), os.path.join(planning, f))
        copy(os.path.join(LAYER, c["tracker"]),
             os.path.join(planning, "Deaf Sports " + name + " Evidence Tracker.xlsx"))
        files = sum(len(fs) for _, _, fs in os.walk(base))
        print(name + ": " + str(files) + " files -> " + base)


if __name__ == "__main__":
    main()
