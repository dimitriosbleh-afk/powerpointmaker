"""Place sections of a lyric file onto the lyric slides of a built deck.

    python scripts/inject_lyrics.py lyrics/<song>.txt "output/<Folder>/<Deck>.pptx"

The lyric file is a plain text copy of the song, with the usual bracketed
section headers the lyric sites use:

    [Verse 1]
    ...lines...

    [Chorus]
    ...lines...

Lines starting with # are ignored. PLACEMENTS below says which section lands
on which slide, and how many of its lines to use. Edit that list, not the code.

The script replaces the text of the widest body box on each named slide and
keeps its font, colour and card. Everything else in the deck is untouched:
images, notes, animations, hyperlinks. Re-running it is safe - it overwrites
whatever is on those slides, so it also fixes a hand-typed transcription slip.

Nothing is printed except line counts, so the lyric text never leaves the file.
"""
import math
import re
import sys
from pptx import Presentation
from pptx.util import Pt, Emu

# slide is 1-based, as shown in PowerPoint's slide pane.
# lines: None means the whole section; a number means the first N lines.
PLACEMENTS = [
    {"slide": 11, "section": "Chorus", "occurrence": 1, "lines": None},
    {"slide": 12, "section": "Verse 2", "occurrence": 1, "lines": 3},
]

MIN_PT, MAX_PT = 15.0, 30.0


def parse_sections(path):
    """-> list of (name, [lines]) in file order."""
    out, name, buf = [], None, []
    for raw in open(path, encoding="utf-8-sig"):
        line = raw.rstrip("\n").rstrip()
        if line.lstrip().startswith("#"):
            continue
        m = re.match(r"^\s*\[([^\]]+)\]\s*$", line)
        if m:
            if name is not None:
                out.append((name, buf))
            name, buf = m.group(1).strip(), []
            continue
        if name is not None and line.strip():
            buf.append(line.strip())
    if name is not None:
        out.append((name, buf))
    return out


def pick(sections, want, occurrence):
    seen = 0
    for name, lines in sections:
        if name.lower() == want.lower():
            seen += 1
            if seen == occurrence:
                return lines
    have = ", ".join(sorted({n for n, _ in sections}))
    raise SystemExit(
        f"section [{want}] occurrence {occurrence} not found in the lyric file.\n"
        f"Sections present: {have}"
    )


def body_box(slide):
    """The extract card's body: the largest text frame in the content band
    that actually holds text. The card behind it also has a text frame and is
    bigger, so "has text" is what separates them."""
    best, best_area = None, 0
    for sh in slide.shapes:
        if not sh.has_text_frame or sh.top is None:
            continue
        if not sh.text_frame.text.strip():
            continue
        top = Emu(sh.top).inches
        if not (1.25 <= top <= 4.0):
            continue
        area = Emu(sh.width).inches * Emu(sh.height).inches
        if area > best_area:
            best, best_area = sh, area
    if best is None:
        raise SystemExit("no body text box found on that slide")
    return best


# Georgia at N pt: about N * 0.00823 inches per character, N * 0.0176 per line.
CHAR_W, LINE_H = 0.00823, 0.0176


def fit_size(lines, box_w, box_h):
    """Largest size at which the lines fit the box.

    First pass insists no line wraps: on a lyric slide a wrapped line reads as
    two lines, which is confusing when the lesson is about counting lines.
    Second pass allows wrapping rather than shrinking past MIN_PT.
    """
    usable_w, usable_h = box_w - 0.15, box_h - 0.12

    def fits(size, allow_wrap):
        per_line = max(8, int(usable_w / (size * CHAR_W)))
        rendered = sum(max(1, math.ceil(len(t) / per_line)) for t in lines)
        if not allow_wrap and rendered != len(lines):
            return False
        return rendered * size * LINE_H <= usable_h

    for allow_wrap in (False, True):
        half = int(MAX_PT * 2)
        while half >= int(MIN_PT * 2):
            size = half / 2.0
            if fits(size, allow_wrap):
                return size
            half -= 1
    return MIN_PT


def write_lines(shape, lines):
    tf = shape.text_frame
    tf.word_wrap = True
    proto = None
    for p in tf.paragraphs:
        if p.runs:
            proto = p.runs[0].font
            break
    size = fit_size(lines, Emu(shape.width).inches, Emu(shape.height).inches)

    for p in list(tf.paragraphs)[1:]:
        p._p.getparent().remove(p._p)
    first = tf.paragraphs[0]
    for r in list(first.runs):
        r._r.getparent().remove(r._r)

    for i, text in enumerate(lines):
        p = first if i == 0 else tf.add_paragraph()
        r = p.add_run()
        r.text = text
        r.font.size = Pt(round(size, 1))
        if proto is not None:
            if proto.name:
                r.font.name = proto.name
            r.font.bold = proto.bold
            try:
                if proto.color and proto.color.rgb:
                    r.font.color.rgb = proto.color.rgb
            except (AttributeError, TypeError):
                pass
    return round(size, 1)


def main():
    if len(sys.argv) != 3:
        raise SystemExit(__doc__)
    lyrics_path, deck_path = sys.argv[1], sys.argv[2]
    sections = parse_sections(lyrics_path)
    if not sections:
        raise SystemExit(
            f"{lyrics_path} has no [Section] headers. Paste the whole song, "
            "including its bracketed headers, and save."
        )
    prs = Presentation(deck_path)
    for job in PLACEMENTS:
        lines = pick(sections, job["section"], job.get("occurrence", 1))
        if job.get("lines"):
            lines = lines[: job["lines"]]
        if not lines:
            raise SystemExit(f"[{job['section']}] is empty in the lyric file")
        slide = prs.slides[job["slide"] - 1]
        pt = write_lines(body_box(slide), lines)
        print(f"slide {job['slide']:>2}: placed {len(lines)} line(s) "
              f"from [{job['section']}] at {pt} pt")
    prs.save(deck_path)
    print(f"saved {deck_path}")


if __name__ == "__main__":
    main()
