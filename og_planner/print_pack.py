"""The week's print pack: one PDF the teacher prints from at the start of the week
(OG_MEGA_PROMPT 1c, v1.9).

James (30 Oct 2026): better outcomes for students without more organising for
staff.  So every printable the builder makes for the week lands in one file, in
day order, behind a one-page cover that says what each page is for, who it is
for and which pages to print - plus what each day needs on the desks.  Teacher
copies (answers, follow-up) sit at the back.

build_og_week.py calls build_pack() after the last session of a v1.9 week.
"""
from pathlib import Path

import fitz

FONTS = Path(r"C:\Windows\Fonts")
REG, BOLD = FONTS / "arial.ttf", FONTS / "arialbd.ttf"
PAGE_W, PAGE_H = fitz.paper_size("a4")
M = 40.0
INK = (0.13, 0.13, 0.15)
GREY = (0.4, 0.4, 0.45)
LINE = (0.78, 0.75, 0.82)
MAGENTA = (0xA6 / 255, 0x4D / 255, 0x79 / 255)
TINT = (0.965, 0.92, 0.945)


def file_name(week):
    return f"0. Print pack - Term {week['term']} Week {week['week']}.pdf"


def _wrap(font, text, size, width):
    lines, cur = [], ""
    for word in str(text).split():
        trial = (cur + " " + word).strip()
        if font.text_length(trial, fontsize=size) <= width or not cur:
            cur = trial
        else:
            lines.append(cur)
            cur = word
    return lines + ([cur] if cur else [])


def _routines(obj):
    """Every `routine` chip on the session's slides ("Whiteboards - show me")."""
    if isinstance(obj, dict):
        for k, v in obj.items():
            if k == "routine" and isinstance(v, str):
                yield v
            else:
                yield from _routines(v)
    elif isinstance(obj, list):
        for v in obj:
            yield from _routines(v)


def needs(session):
    """What a session needs on the desks, read from its own routine chips."""
    routines = " ".join(_routines(session)).lower()
    out = ["OG book"]
    if "whiteboard" in routines:
        out.append("mini whiteboard and marker")
    if session.get("friday_review"):
        out = ["a pencil", "Friday Review pages 1-2"]
    return out


def build_pack(week, week_dir, decks, friday):
    """decks: [(session, deck file name)] in session order.
    friday: {"worksheet": path, "student_pages": n, "follow_up": path} or {}.
    Returns (pack path, page count)."""
    week_dir = Path(week_dir)
    student, teacher = [], []      # (day, what, who, path, from_page, to_page)
    for n, (session, _deck) in enumerate(decks, 1):
        day = session["day"]
        en = week_dir / "Enablers" / f"{n}. {day} enabler.pdf"
        ex = week_dir / "Extension" / f"{n}. {day} extension.pdf"
        if en.exists():
            student.append((day, "Enabler sheet", "Students who need support", en, 0, 0))
        if ex.exists():
            student.append((day, "Extension sheet", "Students ready for more", ex, 0, 0))
            teacher.append((day, "Extension answers", "Teacher", ex, 1, 1))
    if friday:
        student.append(("Friday", "Friday Review (2 pages)", "Every student",
                        friday["worksheet"], 0, friday["student_pages"] - 1))
        with fitz.open(friday["worksheet"]) as d:
            last = d.page_count - 1
        teacher.append(("Friday", "Friday Review answers", "Teacher", friday["worksheet"],
                        friday["student_pages"], last))
        teacher.append(("Friday", "Teacher Follow-up sheet", "Teacher",
                        friday["follow_up"], 0, 0))

    pack = fitz.open()
    pack.new_page(width=PAGE_W, height=PAGE_H)          # cover, drawn last
    placed = []
    for group in (student, teacher):
        for day, what, who, path, a, b in group:
            start = pack.page_count + 1
            with fitz.open(path) as src:
                pack.insert_pdf(src, from_page=a, to_page=b)
            end = pack.page_count
            placed.append((day, what, who, start if start == end else f"{start}-{end}"))

    page = pack[0]
    page.insert_font(fontname="r", fontfile=str(REG))
    page.insert_font(fontname="b", fontfile=str(BOLD))
    fr, fb = fitz.Font(fontfile=str(REG)), fitz.Font(fontfile=str(BOLD))
    y = M + 22
    page.insert_text((M, y), f"OG print pack - Term {week['term']} Week {week['week']}",
                     fontname="b", fontsize=20, color=MAGENTA)
    y += 20
    for ln in _wrap(fr, "Print this file once at the start of the week. Each row says what "
                        "the page is for, who it is for and which pages to print. Teacher "
                        "copies are at the back.", 10.5, PAGE_W - 2 * M):
        page.insert_text((M, y), ln, fontname="r", fontsize=10.5, color=GREY)
        y += 14
    y += 8
    cols = [("Day", 70), ("Sheet", 160), ("Who it is for", 170), ("Pages", 115)]

    def row(cells, bold=False, fill=None):
        nonlocal y
        h = 20
        if fill:
            page.draw_rect(fitz.Rect(M, y, PAGE_W - M, y + h), color=None, fill=fill)
        x = M
        for (label, w), cell in zip(cols, cells):
            page.insert_text((x + 6, y + 14), str(cell), fontname="b" if bold else "r",
                             fontsize=10, color=(1, 1, 1) if fill == MAGENTA else INK)
            x += w
        page.draw_line((M, y + h), (PAGE_W - M, y + h), color=LINE, width=0.6)
        y += h

    row([c for c, _ in cols], bold=True, fill=MAGENTA)
    for i, (day, what, who, pages) in enumerate(placed):
        row([day, what, who, pages], fill=TINT if who == "Teacher" else None)
    y += 22
    page.insert_text((M, y), "Each day at a glance", fontname="b", fontsize=13, color=MAGENTA)
    y += 8
    for n, (session, deck) in enumerate(decks, 1):
        y += 18
        page.insert_text((M, y), session["day"], fontname="b", fontsize=10.5, color=INK)
        text = f"Deck: {deck}.  On desks: {', '.join(needs(session))}."
        for ln in _wrap(fr, text, 10, PAGE_W - 2 * M - 80):
            page.insert_text((M + 80, y), ln, fontname="r", fontsize=10, color=INK)
            y += 13
        y -= 13
    y += 30
    for ln in _wrap(fr, "Enabler and extension sheets: one copy for each student who needs "
                        "one that day. The Friday Review: one copy of pages 1-2 for every "
                        "student, printed double-sided if you like.", 10, PAGE_W - 2 * M):
        page.insert_text((M, y), ln, fontname="r", fontsize=10, color=GREY)
        y += 13
    out = week_dir / file_name(week)
    pack.save(str(out), garbage=4, deflate=True)
    count = pack.page_count
    pack.close()
    return out, count
