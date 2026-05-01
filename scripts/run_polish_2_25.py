"""Run baseline polish across lessons 2-25.

Reads each '... - with notes.pptx' source and writes '... - with notes v2.pptx'
beside it. Lesson 1 is skipped (it was polished manually).
"""
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent))

from polish_notes_lib import polish_deck  # noqa: E402


SRC_DIR = Path(r"c:/Users/09560329/Downloads/Challenge Slides Updated")


def main():
    files = sorted(
        [f for f in SRC_DIR.glob("*.pptx")
         if "with notes" in f.name.lower() and " v2" not in f.name],
        key=lambda p: int(p.name.split(".")[0]),
    )
    for f in files:
        n = int(f.name.split(".")[0])
        if n == 1:
            continue
        out = f.with_name(f.stem + " v2.pptx")
        stats = polish_deck(f, out)
        print(f"L{n:02d}: slides={stats['slides']:>3}  added={stats['added']:>2}  fixed={stats['fixed']:>2}  -> {out.name[:70]}...")


if __name__ == "__main__":
    main()
