#!/usr/bin/env python3
"""Friday Review worksheet (student pages + answers page) and Teacher Follow-up sheet.

Ported on 30 Oct 2026 from the 5/6 Challenge team's package, whose worksheets copy
the school's "Friday OG Review Term 3 - Challenge" sheets: purple Part banners,
numbered question blocks, a labelled Teacher check panel on every question,
Example / Tip / Word Bank / Challenge boxes, and generous writing space.  The
drawing engine is theirs, unchanged; what goes IN the worksheet is pitched for
our Grade 5/6 week by OG_MEGA_PROMPT section 1a (v1.9), and build_og_week.py
gates it.

The layout is a keep-together flow: every question block is measured first and
moved to a new page whole, so no question ever splits across a page break.
Writing space is fixed by the constants below, never shrunk to fit.
"""

import re
from pathlib import Path

import fitz

MM = 72 / 25.4
PAGE_W, PAGE_H = fitz.paper_size("a4")
MARGIN = 13 * MM
CONTENT_W = PAGE_W - 2 * MARGIN
LINE_GAP = 10 * MM            # distance between writing lines (spec minimum 9 mm)
WORD_LINE_MIN = 42 * MM       # one-word answer line (spec minimum 40 mm)
BODY = 11.5
SMALL = 10
HEAD = 13
BANNER = 17

PURPLE = (0x5B / 255, 0x3F / 255, 0xA0 / 255)
PURPLE_DARK = (0x3E / 255, 0x2A / 255, 0x75 / 255)
TINT = (0xF1 / 255, 0xEC / 255, 0xFA / 255)
OUTLINE = (0xC9 / 255, 0xBE / 255, 0xE6 / 255)
INK = (0.13, 0.13, 0.15)
GREY = (0.45, 0.45, 0.5)
LINE_INK = (0.55, 0.55, 0.6)
WHITE = (1, 1, 1)
GOLD = (0xE8 / 255, 0xA8 / 255, 0x20 / 255)

_FONT_FILES = {
    "reg": [r"C:\Windows\Fonts\arial.ttf", "/Library/Fonts/Arial.ttf",
            "/System/Library/Fonts/Supplemental/Arial.ttf"],
    "bold": [r"C:\Windows\Fonts\arialbd.ttf", "/Library/Fonts/Arial Bold.ttf",
             "/System/Library/Fonts/Supplemental/Arial Bold.ttf"],
}


def _font_file(kind):
    for candidate in _FONT_FILES[kind]:
        if Path(candidate).exists():
            return candidate
    return None


class Fonts:
    def __init__(self):
        self.files = {k: _font_file(k) for k in _FONT_FILES}
        self.objs = {}
        for kind, path in self.files.items():
            self.objs[kind] = (fitz.Font(fontfile=path) if path
                               else fitz.Font("helv" if kind == "reg" else "hebo"))

    def width(self, text, size, bold=False):
        return self.objs["bold" if bold else "reg"].text_length(text, fontsize=size)

    def name(self, bold):
        return "ogb" if bold else "ogr"

    def install(self, page):
        for kind, fname in (("reg", "ogr"), ("bold", "ogb")):
            path = self.files[kind]
            if path:
                page.insert_font(fontname=fname, fontfile=path)
            else:
                page.insert_font(fontname=fname, fontbuffer=self.objs[kind].buffer)


def wrap(fonts, text, size, width, bold=False):
    lines = []
    for para in str(text).split("\n"):
        words = para.split(" ")
        cur = ""
        for word in words:
            trial = word if not cur else cur + " " + word
            if fonts.width(trial, size, bold) <= width or not cur:
                cur = trial
            else:
                lines.append(cur)
                cur = word
        lines.append(cur)
    return lines


GAP_RE = re.compile(r"_{3,}")


def widen_gaps(text):
    """Every answer gap in a cloze line becomes at least WORD_LINE_MIN long."""
    return GAP_RE.sub("_" * 22, str(text))


class Canvas:
    """Draws onto a page, or only measures when page is None."""

    def __init__(self, fonts, page=None):
        self.fonts = fonts
        self.page = page

    def text(self, x, y, s, size=BODY, bold=False, color=INK):
        if self.page is not None and s:
            self.page.insert_text((x, y), s, fontsize=size,
                                  fontname=self.fonts.name(bold), color=color)

    def rect(self, r, fill=None, stroke=None, width=0.8, radius=None):
        if self.page is None:
            return
        shape = self.page.new_shape()
        if radius:
            shape.draw_rect(fitz.Rect(*r), radius=radius)
        else:
            shape.draw_rect(fitz.Rect(*r))
        shape.finish(color=stroke, fill=fill, width=width)
        shape.commit()

    def line(self, x1, y, x2, color=LINE_INK, width=0.7):
        if self.page is None:
            return
        self.page.draw_line((x1, y), (x2, y), color=color, width=width)

    def circle(self, cx, cy, r, fill):
        if self.page is None:
            return
        self.page.draw_circle((cx, cy), r, color=fill, fill=fill)

    def star(self, cx, cy, r, fill=PURPLE):
        if self.page is None:
            return
        import math
        pts = []
        for i in range(10):
            rad = r if i % 2 == 0 else r * 0.45
            ang = math.pi / 2 + i * math.pi / 5
            pts.append((cx + rad * math.cos(ang), cy - rad * math.sin(ang)))
        shape = self.page.new_shape()
        shape.draw_polyline(pts + [pts[0]])
        shape.finish(color=fill, fill=fill, width=0.5, closePath=True)
        shape.commit()

    def paragraph(self, x, y, width, s, size=BODY, bold=False, color=INK, leading=1.35):
        """Wrapped text whose FIRST baseline sits at y + size. Returns height used."""
        lines = wrap(self.fonts, s, size, width, bold)
        for i, ln in enumerate(lines):
            self.text(x, y + size + i * size * leading, ln, size, bold, color)
        return len(lines) * size * leading


# ------------------------------------------------------------------ question block
def check_panel(cv, right, top, labels):
    """'Teacher check:' panel with one labelled tick box per skill. Returns height."""
    labels = [str(label) for label in labels]
    box = 8.5
    item_w = [box + 4 + cv.fonts.width(label, 8.5) + 10 for label in labels]
    width = max(cv.fonts.width("Teacher check:", 9, True) + 12, sum(item_w) + 10)
    width = min(width, CONTENT_W * 0.55)
    height = 24
    left = right - width
    cv.rect((left, top, right, top + height), fill=WHITE, stroke=OUTLINE, width=1, radius=0.15)
    cv.text(left + 6, top + 9.5, "Teacher check:", 9, True, PURPLE_DARK)
    x = left + 6
    for label, w in zip(labels, item_w):
        cv.rect((x, top + 12.5, x + box, top + 12.5 + box), fill=WHITE, stroke=INK, width=0.7)
        cv.text(x + box + 4, top + 20, label, 8.5, False, INK)
        x += w
    return width, height


def draw_items(cv, x, y, width, task, number_offset=0):
    """Render the task body. Returns height used."""
    kind = task.get("type", "sentences")
    items = task.get("items", [])
    start_y = y
    if kind in ("sentences", "cloze"):
        extra_lines = int(task.get("answer_lines", 0 if kind == "cloze" else 1))
        for i, item in enumerate(items, 1):
            text = widen_gaps(item) if kind == "cloze" else str(item)
            label = f"{i + number_offset}."
            cv.text(x, y + BODY, label, BODY, True, PURPLE_DARK)
            if extra_lines:
                y += cv.paragraph(x + 18, y, width - 18, text, BODY) + 2
                for _ in range(extra_lines):
                    y += LINE_GAP
                    cv.line(x + 18, y, x + width)
                y += 8
            else:
                # circle/underline and cloze items: room above and below to mark or write
                used = cv.paragraph(x + 18, y, width - 18, text, BODY, leading=1.9)
                y += max(used, LINE_GAP) + 2
    elif kind == "meaning_lines":
        for i, item in enumerate(items, 1):
            label = f"{i}.  {item}"
            cv.text(x, y + BODY, label, BODY, False, INK)
            lx = x + cv.fonts.width(label, BODY) + 8
            cv.line(max(lx, x + 120), y + BODY + 2, x + width)
            y += LINE_GAP + 2
    elif kind == "clue_answer":
        clue_w = width * 0.62
        for i, item in enumerate(items, 1):
            cv.text(x, y + BODY, f"{i}.", BODY, True, PURPLE_DARK)
            used = cv.paragraph(x + 18, y, clue_w - 24, str(item), BODY)
            cv.line(x + clue_w, y + BODY + 2, x + width)
            y += max(used, LINE_GAP) + 8
    elif kind == "codes":
        codes = task.get("codes", "")
        cols = max(1, int(task.get("columns", 1)))
        if cols > 1:
            # the 3 x 2 grid the spec allows: the code key lives in the Tip box,
            # so each cell still gets a full WORD_LINE_MIN answer line
            col_w = width / cols
            for i, item in enumerate(items, 1):
                col, row = (i - 1) % cols, (i - 1) // cols
                cx, cy = x + col * col_w, y + row * (LINE_GAP + 2)
                label = f"{i}.  {item} = "
                cv.text(cx, cy + BODY, label, BODY, False, INK)
                lx = cx + cv.fonts.width(label, BODY)
                cv.line(lx, cy + BODY + 2, lx + WORD_LINE_MIN)
            y += ((len(items) + cols - 1) // cols) * (LINE_GAP + 2)
        else:
            for i, item in enumerate(items, 1):
                label = f"{i}.  {item} = "
                cv.text(x, y + BODY, label, BODY, False, INK)
                lx = x + cv.fonts.width(label, BODY)
                cv.line(lx, y + BODY + 2, lx + WORD_LINE_MIN)
                cv.text(lx + WORD_LINE_MIN + 10, y + BODY, f"({codes})", BODY,
                        False, GREY)
                y += LINE_GAP
    elif kind == "match":
        meanings = task.get("meanings", [])
        left_w = width * 0.38
        right_x = x + left_w + 14
        right_w = width - left_w - 14
        rows = max(len(items), len(meanings))
        for i in range(rows):
            row_h = LINE_GAP
            if i < len(items):
                label = f"{i + 1}.  {items[i]}"
                cv.text(x, y + BODY, label, BODY, False, INK)
                lx = x + cv.fonts.width(label, BODY) + 6
                cv.line(max(lx, x + left_w - 60), y + BODY + 2, x + left_w)
            if i < len(meanings):
                used = cv.paragraph(right_x, y, right_w, meanings[i], BODY)
                row_h = max(row_h, used)
            y += row_h + 6
    elif kind == "table":
        headers = task.get("headers", [])
        rows = task.get("rows", [])
        widths = task.get("widths") or [1] * len(headers)
        total = sum(widths)
        col_w = [width * w / total for w in widths]
        head_h = 20
        cv.rect((x, y, x + width, y + head_h), fill=PURPLE, stroke=PURPLE, radius=None)
        cx = x
        for text, w in zip(headers, col_w):
            cv.text(cx + 5, y + 14, str(text), SMALL, True, WHITE)
            cx += w
        y += head_h
        for row in rows:
            heights = []
            for cell, w in zip(row, col_w):
                heights.append(len(wrap(cv.fonts, str(cell), BODY, w - 10)) * BODY * 1.35)
            row_h = max(10.5 * MM, max(heights) + 10)
            cv.rect((x, y, x + width, y + row_h), fill=None, stroke=OUTLINE, width=0.8)
            cx = x
            for cell, w in zip(row, col_w):
                if str(cell).strip():
                    cv.paragraph(cx + 5, y + 4, w - 10, str(cell), BODY)
                else:
                    cv.line(cx + 8, y + row_h - 7, cx + w - 8)
                cx += w
            y += row_h
    elif kind == "syllables":
        cols = int(task.get("columns", 3))
        cell_w = width / cols
        cell_h = 16 * MM
        for i, word in enumerate(items):
            col, row = i % cols, i // cols
            cx, cy = x + col * cell_w, y + row * cell_h
            cv.rect((cx, cy, cx + cell_w, cy + cell_h), stroke=OUTLINE, width=0.8)
            cv.text(cx + 6, cy + 15, f"{i + 1}.  {word}", BODY, False, INK)
            cv.line(cx + 10, cy + cell_h - 9, cx + cell_w - 10)
        y += ((len(items) + cols - 1) // cols) * cell_h
    elif kind == "build":
        if task.get("prompt"):
            y += cv.paragraph(x, y, width, task["prompt"], BODY, True) + 4
        for _ in range(max(6, int(task.get("lines", 6)))):
            y += LINE_GAP
            cv.line(x, y, x + width)
        y += 4
    elif kind == "lines":
        for _ in range(int(task.get("lines", 2))):
            y += LINE_GAP
            cv.line(x, y, x + width)
        y += 4
    else:
        raise ValueError(f"unknown Friday Review task type {kind!r}")
    return y - start_y


def draw_block(cv, top, block, number):
    """One question block in its outlined panel. Returns the height used."""
    x = MARGIN + 10
    width = CONTENT_W - 20
    y = top + 6
    heading = block.get("heading", "")
    is_challenge = block.get("challenge", False)
    checks = block.get("checks", [])
    panel_w, panel_h = check_panel(cv, MARGIN + CONTENT_W - 8, y, checks) if checks else (0, 0)
    head_w = width - panel_w - 12
    if is_challenge:
        cv.star(x + 8, y + 9, 8, GOLD)
        cv.text(x + 22, y + 14, heading or "Challenge", HEAD, True, PURPLE_DARK)
        head_used = 20
    elif heading:
        cv.circle(x + 9, y + 9, 9, PURPLE)
        cv.text(x + 5.5 if number < 10 else x + 2.5, y + 13, str(number), 11, True, WHITE)
        head_used = cv.paragraph(x + 24, y - 2, head_w - 24, heading, HEAD, True, PURPLE_DARK)
    else:
        head_used = 0
    y += max(head_used, panel_h) + 6
    instruction = block.get("instruction", "")
    if instruction:
        cv.star(x + 5, y + 7, 5)
        y += cv.paragraph(x + 16, y, width - 16, instruction, BODY, True) + 4
    def info_box(bx, by, bw, title, lines, bold_words):
        inner = []
        for ln in lines:
            inner += wrap(cv.fonts, str(ln), BODY, bw - 24, bold_words)
        box_h = 18 + len(inner) * BODY * 1.4 + 4
        return inner, box_h

    if block.get("tip") and block.get("example"):
        # side by side, as on the school's Week 5 sheet
        col_w = (width - 10) / 2
        parts = []
        for title, key, bx in (("Tip:", "tip", x), ("Example:", "example", x + col_w + 10)):
            value = block[key]
            lines = value if isinstance(value, list) else [value]
            inner, box_h = info_box(bx, y, col_w, title, lines, False)
            parts.append((title, bx, inner, box_h))
        tallest = max(p[3] for p in parts)
        for title, bx, inner, _ in parts:
            cv.rect((bx, y, bx + col_w, y + tallest), fill=TINT, stroke=OUTLINE, width=0.8, radius=0.12)
            cv.text(bx + 10, y + 13, title, SMALL, True, PURPLE_DARK)
            for i, ln in enumerate(inner):
                cv.text(bx + 12, y + 18 + BODY + i * BODY * 1.4, ln, BODY, False, INK)
        y += tallest + 6
        box_keys = (("Word Bank:", "word_bank"),)
    else:
        box_keys = (("Tip:", "tip"), ("Example:", "example"), ("Word Bank:", "word_bank"))
    for title, key in box_keys:
        value = block.get(key)
        if not value:
            continue
        lines = value if isinstance(value, list) else [value]
        if key == "word_bank":
            lines = ["      ".join(str(v) for v in lines)]
        inner = []
        for ln in lines:
            inner += wrap(cv.fonts, str(ln), BODY, width - 24, key == "word_bank")
        title_w = cv.fonts.width(title, SMALL, True) + 14
        one_row = (len(inner) == 1 and cv.fonts.width(
            inner[0], BODY, key == "word_bank") <= width - 24 - title_w)
        if one_row:
            # label and value share a row: the box is reference text, not writing space
            box_h = 5 + BODY * 1.4 + 4
            cv.rect((x, y, x + width, y + box_h), fill=TINT, stroke=OUTLINE, width=0.8,
                    radius=0.12)
            cv.text(x + 10, y + 5 + BODY, title, SMALL, True, PURPLE_DARK)
            cv.text(x + 10 + title_w, y + 5 + BODY, inner[0], BODY,
                    key == "word_bank", INK)
        else:
            box_h = 18 + len(inner) * BODY * 1.4 + 4
            cv.rect((x, y, x + width, y + box_h), fill=TINT, stroke=OUTLINE, width=0.8,
                    radius=0.12)
            cv.text(x + 10, y + 13, title, SMALL, True, PURPLE_DARK)
            for i, ln in enumerate(inner):
                cv.text(x + 12, y + 18 + BODY + i * BODY * 1.4, ln, BODY,
                        key == "word_bank", INK)
        y += box_h + 6
    task = block.get("task")
    if task:
        y += draw_items(cv, x, y, width, task)
    height = y - top + 4
    stroke = GOLD if is_challenge else OUTLINE
    cv.rect((MARGIN, top, MARGIN + CONTENT_W, top + height), fill=None, stroke=stroke,
            width=1.1, radius=0.03)
    return height


def banner(cv, top, text):
    w = cv.fonts.width(text, BANNER, True) + 40
    cv.rect((MARGIN, top, MARGIN + w, top + 25), fill=PURPLE, stroke=PURPLE, radius=0.25)
    cv.text(MARGIN + 12, top + 18, text, BANNER, True, WHITE)
    cv.star(MARGIN + w - 14, top + 12.5, 7, GOLD)
    return 28


PART_TITLES = {"part_a": "Part A", "part_b": "Part B", "part_c": "Part C", "part_d": "Part D"}


def build_worksheet(review, out_path, week_label, dictation):
    """Write the PDF. Returns a report dict used by the builder's gate."""
    fonts = Fonts()
    doc = fitz.open()
    report = {"blocks": [], "student_pages": 0, "answer_pages": 0, "problems": []}
    state = {"page": None, "y": 0}

    def new_page():
        page = doc.new_page(width=PAGE_W, height=PAGE_H)
        fonts.install(page)
        state["page"] = page
        state["y"] = MARGIN
        foot = f"Friday Review - {week_label}"
        Canvas(fonts, page).text(MARGIN, PAGE_H - 7 * MM, foot, 8, False, GREY)
        return page

    def room(height):
        return state["y"] + height <= PAGE_H - MARGIN - 6 * MM

    new_page()
    cv = Canvas(fonts, state["page"])
    name_x = MARGIN + CONTENT_W * 0.55
    title = f"Friday Review - {week_label}"
    # a long label (e.g. "Week 4 Extension") shrinks rather than running into Name:
    title_size = min(20, 20 * (name_x - MARGIN - 12) / fonts.width(title, 20, True))
    cv.text(MARGIN, state["y"] + 20, title, title_size, True, PURPLE_DARK)
    cv.text(name_x, state["y"] + 20, "Name:", 12, True, INK)
    cv.line(name_x + 42, state["y"] + 22, MARGIN + CONTENT_W, INK, 0.8)
    state["y"] += 36

    for key in ("part_a", "part_b", "part_c", "part_d"):
        part = review.get(key) or {}
        title = f"{PART_TITLES[key]}: {part.get('title', '')}".rstrip(": ")
        blocks = list(part.get("subblocks", []))
        for challenge in review.get("challenges", []) or []:
            if challenge.get("part") == key:
                blocks.append({"heading": "Challenge", "challenge": True,
                               "instruction": challenge.get("text", ""),
                               "checks": challenge.get("checks", []),
                               "task": {"type": "lines", "lines": challenge.get("lines", 2)},
                               "answers": challenge.get("answers", [])})
        number = 0
        for i, block in enumerate(blocks):
            if not block.get("challenge"):
                number += 1
            show_number = number if (len([b for b in blocks if not b.get("challenge")]) > 1
                                     or block.get("heading")) else number
            measure = draw_block(Canvas(fonts), 0, block, show_number)
            need = measure + 5 + (28 if i == 0 else 0)
            if not room(need):
                if measure + 28 > PAGE_H - 2 * MARGIN - 6 * MM:
                    report["problems"].append(
                        f"{title}: block {block.get('heading') or i + 1!r} is taller than a "
                        "page; cut items rather than shrinking the writing space")
                new_page()
            cv = Canvas(fonts, state["page"])
            if i == 0:
                state["y"] += banner(cv, state["y"], title)
            used = draw_block(cv, state["y"], block, show_number)
            report["blocks"].append({"part": key, "heading": block.get("heading", ""),
                                     "challenge": bool(block.get("challenge")),
                                     "checks": list(block.get("checks", [])),
                                     "page": doc.page_count})
            state["y"] += used + 5

    # Part E: dictation lines
    lines = max(5, int(review.get("dictation_lines", 6)))
    need = 28 + lines * LINE_GAP + 10
    if not room(need):
        new_page()
    cv = Canvas(fonts, state["page"])
    state["y"] += banner(cv, state["y"], "Part E: Dictation")
    for _ in range(lines):
        state["y"] += LINE_GAP
        cv.line(MARGIN, state["y"], MARGIN + CONTENT_W, LINE_INK, 0.8)
    report["student_pages"] = doc.page_count

    # Answers page(s)
    new_page()
    cv = Canvas(fonts, state["page"])
    cv.text(MARGIN, state["y"] + 20, f"Friday Review - {week_label}: Answers", 18, True,
            PURPLE_DARK)
    cv.text(MARGIN, state["y"] + 36, "Teacher copy - do not print for students.", 10,
            False, GREY)
    state["y"] += 50

    def answer_line(text, size=BODY, bold=False, color=INK, indent=0):
        height = len(wrap(fonts, text, size, CONTENT_W - indent, bold)) * size * 1.35 + 3
        if not room(height):
            new_page()
        page_cv = Canvas(fonts, state["page"])
        state["y"] += page_cv.paragraph(MARGIN + indent, state["y"], CONTENT_W - indent,
                                        text, size, bold, color) + 3

    for key in ("part_a", "part_b", "part_c", "part_d"):
        part = review.get(key) or {}
        answer_line(f"{PART_TITLES[key]}: {part.get('title', '')}", 13, True, PURPLE_DARK)
        for block in part.get("subblocks", []):
            if block.get("heading"):
                answer_line(block["heading"], BODY, True)
            for n, ans in enumerate(block.get("answers", []), 1):
                answer_line(f"{n}. {ans}", BODY, False, INK, 12)
        for challenge in review.get("challenges", []) or []:
            if challenge.get("part") == key and challenge.get("answers"):
                answer_line("Challenge", BODY, True)
                for ans in challenge["answers"]:
                    answer_line(f"e.g. {ans}", BODY, False, INK, 12)
        state["y"] += 6
    answer_line("Part E: Dictation (read each sentence whole, then in chunks, then whole again)", 13, True, PURPLE_DARK)
    for n, entry in enumerate(dictation, 1):
        answer_line(f"{n}. {entry.get('sentence', '')}", BODY, False, INK, 12)
    report["answer_pages"] = doc.page_count - report["student_pages"]

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path), garbage=3, deflate=True)
    doc.close()
    return report


FOLLOW_UP_HEADERS = ("Question", "What it checks", "Add to a student's program if...",
                     "Revisit it next week (no slide edits)")
FOLLOW_UP_WIDTHS = (0.12, 0.20, 0.27, 0.41)


def build_follow_up(rows, out_path, title):
    """One-page A4 Teacher Follow-up sheet (OG_MEGA_PROMPT 1a). Returns the page count.

    Landscape A4 so four columns of short phrases fit one page without shrinking the
    type; rows that will not fit spill to a second page and the builder's gate fails.
    """
    fonts = Fonts()
    doc = fitz.open()
    width, height = fitz.paper_size("a4-l")
    content_w = width - 2 * MARGIN
    col_w = [content_w * f for f in FOLLOW_UP_WIDTHS]
    size = 10

    def page_setup():
        page = doc.new_page(width=width, height=height)
        fonts.install(page)
        cv = Canvas(fonts, page)
        cv.text(MARGIN, MARGIN + 16, title, 16, True, PURPLE_DARK)
        cv.text(MARGIN, MARGIN + 31, "Teacher copy - do not print for students. Labels match "
                "the worksheet's Teacher check boxes.", 9.5, False, GREY)
        y = MARGIN + 42
        cv.rect((MARGIN, y, MARGIN + content_w, y + 20), fill=PURPLE, stroke=PURPLE)
        x = MARGIN
        for text, w in zip(FOLLOW_UP_HEADERS, col_w):
            cv.text(x + 5, y + 14, text, SMALL, True, WHITE)
            x += w
        return cv, y + 20

    cv, y = page_setup()
    for n, row in enumerate(rows):
        checks = row.get("checks", [])
        checks = ", ".join(checks) if isinstance(checks, list) else str(checks)
        cells = [row.get("question", ""), checks, row.get("flag_when", ""),
                 row.get("next_move", "")]
        lines = [wrap(fonts, str(c), size, w - 10) for c, w in zip(cells, col_w)]
        row_h = max(len(ls) for ls in lines) * size * 1.3 + 8
        if y + row_h > height - MARGIN:
            cv, y = page_setup()
        if n % 2:
            cv.rect((MARGIN, y, MARGIN + content_w, y + row_h), fill=TINT, stroke=None)
        cv.rect((MARGIN, y, MARGIN + content_w, y + row_h), fill=None, stroke=OUTLINE, width=0.8)
        x = MARGIN
        for ls, w, bold in zip(lines, col_w, (True, True, False, False)):
            for i, ln in enumerate(ls):
                cv.text(x + 5, y + 4 + size + i * size * 1.3, ln, size, bold, INK)
            x += w
        y += row_h
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    pages = doc.page_count
    doc.save(str(out_path), garbage=3, deflate=True)
    doc.close()
    return pages
