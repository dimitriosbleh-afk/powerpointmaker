"""One-page A4 enabler sheets on dotted thirds for v1.9 OG weeks (OG_MEGA_PROMPT 1b).

Ported from the 5/6 Challenge team's package (Oct 2026), whose sheets copied a
3/4 teacher's hand-made Word sheets.  Re-pitched for our Grade 5/6 week:

- Victorian Modern letters, SEPARATED, not joined (James, 30 Oct 2026: "Vic modern
  but not cursive, just regular separated"): VicDots to trace, VicRegular for the
  bold morpheme.  The same family as VicCursive, unjoined.  Never VICCursiveDotted
  / VICCursiveSolid: their widths are only right inside Word, so in a PDF the
  letters pile up.
- Smaller thirds (about 4.9 mm, not 6 mm): Year 5/6 handwriting is smaller.
- Our day types, not new/review: Monday writes the morpheme and a sentence,
  Tuesday carries the spelling and dictation, Thursday the learned words.
- The scaffold leaves the TEACHING blank: the week's morpheme is the gap in the
  Words to Spell New words, and every dictation target is a gap.

Every row is a solid baseline with two dotted lines above it: the x-height fills
the lowest third, ascenders reach the top line, descenders drop a third into the
next row.  Fonts are read from the installed fonts (Windows or Mac) and never
copied into the repo; without VicDots/VicRegular the free Department font is
used, traced in grey (see FALLBACK below).

build_og_week.py calls build_session_sheet() for each session; run it alone with:
    python og_planner/enabler_pdf.py <week spec>.json <out dir>
"""
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path

import fitz

# Windows first, then the Mac font folders (Font Book installs to ~/Library/Fonts).
FONT_DIRS = [Path(r"C:\Windows\Fonts"), Path.home() / "Library" / "Fonts",
             Path("/Library/Fonts"), Path("/System/Library/Fonts/Supplemental")]


def find_font(*names):
    """The first installed copy of any of `names`; a missing path otherwise,
    which fonts_available() reports."""
    for folder in FONT_DIRS:
        for name in names:
            if (folder / name).exists():
                return folder / name
    return FONT_DIRS[0] / names[0]


DOTTED = find_font("VICDOTS_.TTF")    # VicDots: dotted Victorian Modern, separated
SOLID = find_font("VICREGUL.TTF")     # VicRegular: the same letters, solid
# VicDots and VicRegular are paid School Fonts fonts.  Where they are not installed
# (James's Mac, 31 Oct 2026), fall back to the Department's free Victorian Modern
# Cursive (Vicmorg_.ttf, separated letters) and trace in pale grey instead of dots.
FREE = find_font("Vicmorg_.ttf")
FALLBACK = not (DOTTED.exists() and SOLID.exists()) and FREE.exists()
if FALLBACK:
    DOTTED = SOLID = FREE
HEAD = find_font("arialbd.ttf", "Arial Bold.ttf")
BODY = find_font("arial.ttf", "Arial.ttf")
ITALIC = find_font("ariali.ttf", "Arial Italic.ttf")

PAGE_W, PAGE_H = fitz.paper_size("a4")
LEFT, RIGHT = 36.0, PAGE_W - 36.0
# 24 pt (8.5 mm) at the foot: Term 4 Week 5's Tuesday-type sheet, with guarantee
# and mortgage both on review, ran 6 pt past a 30 pt margin at 4.2 mm thirds.
TOP, BOTTOM = 34.0, PAGE_H - 24.0
GREY = (0.5, 0.5, 0.5)
SIZE_PER_THIRD = 2.7     # x-height 0.33 em = 0.9 of a third (measured, both fonts)
TRACE = (0, 0, 0)        # colour of the letters students trace
if FALLBACK:
    SIZE_PER_THIRD = 2.25    # Vicmorg_ x-height is 0.40 em (measured)
    TRACE = (0.62, 0.62, 0.62)
MIN_GAP = 36.0
WORD_GAP = 0.45          # word space in a traced sentence, as a share of the font size
# Year 5/6 thirds: 14 pt = 4.9 mm (a 14.8 mm row) down to 12 pt = 4.2 mm.
THIRDS = (14.0, 13.5, 13.0, 12.5, 12.0)
SMALL_THIRD_NOTE = 12.5

_FONTS = {}


def font(path):
    if path not in _FONTS:
        _FONTS[path] = fitz.Font(fontfile=str(path))
    return _FONTS[path]


def fonts_available():
    return DOTTED.exists() and SOLID.exists()


class Sheet:
    def __init__(self, third):
        self.T = third
        self.size = third * SIZE_PER_THIRD
        self.ops = []           # deferred drawing so we can measure first
        self.y = TOP            # top solid line of the next row
        self.first = True

    def item_width(self, item, T=None):
        size = (T or self.T) * SIZE_PER_THIRD
        w = 0.0
        for kind, val in item:
            if kind == "t":
                w += font(DOTTED).text_length(val, fontsize=size)
            elif kind == "bold":
                w += font(SOLID).text_length(val, fontsize=size) + 1.5
            elif kind == "b":
                w += blank_w(val, size)
            elif kind == "line":
                w += val
            elif kind == "dash":
                w += size * 0.32
            elif kind == "sp":
                w += val
            elif kind == "s":
                w += font(SOLID).text_length(val, fontsize=size)
        return w

    def row(self, items, align="spread", third=None):
        T = third or self.T
        top = self.y
        base = top + 3 * T
        self.ops.append(("lines", top, T))
        self.ops.append(("items", items, base, T, align))
        self.y = base

    def divider(self):
        # clear the last row's descenders before the double rule
        self.y += self.T
        self.ops.append(("divider", self.y))
        self.y += 6.0

    def heading(self, title):
        """Section heading in its own strip, so no descender ever runs through it."""
        if not self.first:
            self.divider()
        self.first = False
        self.ops.append(("head", title, self.y + 11.0))
        self.y += 14.0

    def para(self, text, size=11.0, color=(0.13, 0.13, 0.15), path=None, fontname="bd"):
        """Printed instructions above the writing rows (Arial, wrapped)."""
        f = font(path or BODY)
        lines, cur = [], ""
        for word in str(text).split():
            trial = (cur + " " + word).strip()
            if f.text_length(trial, fontsize=size) <= RIGHT - LEFT - 6 or not cur:
                cur = trial
            else:
                lines.append(cur)
                cur = word
        if cur:
            lines.append(cur)
        self.ops.append(("para", lines, self.y, size, color, fontname))
        self.y += len(lines) * size * 1.3 + 4

    def section(self, title, rows, align="spread", T=None):
        self.heading(title)
        if len(rows) > 1 and align == "spread":
            align = "pack"      # wrapped rows line up on the left, same gaps
        for r in rows:
            self.row(r, align=align, third=T)

    def pack(self, items, T=None, gap=MIN_GAP):
        """Greedy-pack items into rows that fit the width; never squash."""
        rows, cur, cur_w = [], [], 0.0
        for it in items:
            w = self.item_width(it, T)
            need = w if not cur else cur_w + gap + w
            if cur and need > (RIGHT - LEFT) - 8:
                rows.append(cur)
                cur, cur_w = [it], w
            else:
                cur.append(it)
                cur_w = need
        if cur:
            rows.append(cur)
        return rows

    def wrap_words(self, items, T):
        space = T * SIZE_PER_THIRD * WORD_GAP
        rows, cur, w = [], [], 0.0
        for it in items:
            iw = self.item_width(it, T)
            need = iw if not cur else w + space + iw
            if cur and need > (RIGHT - LEFT) - 12:
                rows.append(cur)
                cur, w = [it], iw
            else:
                cur.append(it)
                w = need
        if cur:
            rows.append(cur)
        return rows

    def render(self, page):
        page.insert_font(fontname="vd", fontfile=str(DOTTED))
        page.insert_font(fontname="vs", fontfile=str(SOLID))
        page.insert_font(fontname="hd", fontfile=str(HEAD))
        page.insert_font(fontname="bd", fontfile=str(BODY))
        page.insert_font(fontname="it", fontfile=str(ITALIC))
        for op in self.ops:
            if op[0] == "lines":
                _, top, T = op
                for k, style in ((0, "solid"), (1, "dot"), (2, "dot"), (3, "solid")):
                    y = top + k * T
                    if style == "solid":
                        page.draw_line((LEFT, y), (RIGHT, y), color=(0, 0, 0), width=0.5)
                    else:
                        page.draw_line((LEFT, y), (RIGHT, y), color=GREY, width=0.5,
                                       dashes="[1.5 0.6] 0")
            elif op[0] == "divider":
                y = op[1]
                page.draw_line((LEFT, y), (RIGHT, y), color=(0, 0, 0), width=0.7)
                page.draw_line((LEFT, y + 2.6), (RIGHT, y + 2.6), color=(0, 0, 0), width=2.6)
                page.draw_line((LEFT, y + 5.2), (RIGHT, y + 5.2), color=(0, 0, 0), width=0.7)
            elif op[0] == "head":
                _, text, base = op
                # a long heading shrinks to fit the strip rather than run off the page
                size = min(10.5, 10.5 * (RIGHT - LEFT - 4)
                           / max(1.0, font(HEAD).text_length(text, fontsize=10.5)))
                page.insert_text((LEFT + 2, base), text, fontname="hd", fontsize=size,
                                 color=(0.25, 0.25, 0.3))
            elif op[0] == "para":
                _, lines, top, size, color, fontname = op
                for i, ln in enumerate(lines):
                    page.insert_text((LEFT + 2, top + size + i * size * 1.3), ln,
                                     fontname=fontname, fontsize=size, color=color)
            elif op[0] == "title":
                _, left_text, base = op[:3]
                page.insert_text((LEFT, base), left_text, fontname="hd", fontsize=11,
                                 color=(0.25, 0.25, 0.3))
                if len(op) > 3 and not op[3]:
                    continue        # a teacher copy has no Name line
                page.insert_text((RIGHT - 210, base), "Name:", fontname="hd", fontsize=11,
                                 color=(0.25, 0.25, 0.3))
                page.draw_line((RIGHT - 172, base + 1), (RIGHT, base + 1), color=(0, 0, 0),
                               width=0.6)
            elif op[0] == "items":
                _, items, base, T, align = op
                size = T * SIZE_PER_THIRD
                widths = [self.item_width(it, T) for it in items]
                total = sum(widths)
                if align == "left":
                    gap, x = size * WORD_GAP, LEFT + 6
                elif align == "pack":
                    gap, x = MIN_GAP, LEFT + 6
                elif len(items) == 1:
                    gap, x = 0, LEFT + ((RIGHT - LEFT) - total) / 2
                else:
                    gap = min(((RIGHT - LEFT - 12) - total) / (len(items) - 1), 70)
                    x = LEFT + ((RIGHT - LEFT) - (total + gap * (len(items) - 1))) / 2
                for it, w in zip(items, widths):
                    draw_item(page, it, x, base, size)
                    x += w + gap


def blank_w(text, size):
    return font(DOTTED).text_length(text, fontsize=size) * 1.15 + 10


def draw_item(page, item, x, base, size):
    for kind, val in item:
        if kind == "t":
            page.insert_text((x, base), val, fontname="vd", fontsize=size, color=TRACE)
            x += font(DOTTED).text_length(val, fontsize=size)
        elif kind == "bold":
            page.insert_text((x, base), val, fontname="vs", fontsize=size, render_mode=2,
                             border_width=0.035, color=(0, 0, 0), fill=(0, 0, 0))
            x += font(SOLID).text_length(val, fontsize=size) + 1.5
        elif kind == "dash":
            # VicDots' own hyphen is a heavy solid bar; draw a dotted one at
            # mid x-height instead
            y = base - size * 0.17
            page.draw_line((x + 2, y), (x + size * 0.32 - 2, y), color=(0, 0, 0),
                           width=0.9, dashes="[1.2 1.0] 0")
            x += size * 0.32
        elif kind == "sp":
            x += val
        elif kind == "s":
            page.insert_text((x, base), val, fontname="vs", fontsize=size)
            x += font(SOLID).text_length(val, fontsize=size)
        elif kind in ("b", "line"):
            w = blank_w(val, size) if kind == "b" else val
            page.draw_line((x + 2, base), (x + w - 2, base), color=(0, 0, 0), width=1.1,
                           dashes="[0.9 1.3] 0")
            x += w


# ---------------------------------------------------------------- content
def morph_forms(label):
    """Every spelling a word may carry: 'grad / gress' -> grad, gress; '-ity' -> ity;
    'spir(e)' -> spir, spire; 'caus-/cus(e)' -> caus, cus, cuse.  gap_part tries
    the longest first, so 'inspiring' gaps spir and 'excuse' gaps cuse."""
    forms = set()
    for part in str(label).casefold().split("/"):
        bare = re.sub(r"\([^)]*\)", "", part)
        full = part.replace("(", "").replace(")", "")
        for f in (bare, full):
            f = re.sub(r"[^a-z]", "", f)
            if f:
                forms.add(f)
    return sorted(forms, key=len, reverse=True)


def partial(word, keep):
    """The first `keep` letters shown, the rest a blank."""
    return [("t", word[:keep]), ("b", word[keep:])] if keep < len(word) else [("t", word)]


def gap_part(word, forms):
    """The week's morpheme is the gap, the rest of the word is there to trace."""
    low = word.casefold()
    for f in sorted(forms, key=len, reverse=True):
        i = low.find(f)
        if i >= 0:
            out = [("t", word[:i])] if i else []
            out.append(("b", word[i:i + len(f)]))
            if word[i + len(f):]:
                out.append(("t", word[i + len(f):]))
            return out
    return partial(word, 1)


def unfair_gap(word, unfair):
    i = word.find(unfair) if unfair else -1
    if i < 0:
        return partial(word, 1)
    out = [("t", word[:i])] if i else []
    out.append(("b", unfair))
    if word[i + len(unfair):]:
        out.append(("t", word[i + len(unfair):]))
    return out


def bank_item(morph):
    """The Sound Bank is DICTATED (2c), so the sheet gives the first letter only."""
    label = str(morph).split("/")[0].strip()
    lead = "-" if label.startswith("-") else ""
    tail = "-" if label.endswith("-") and len(label) > 1 else ""
    core = label.strip("-")
    out = [("dash", None)] if lead else []
    if core[:1]:
        out.append(("t", core[:1]))
    if core[1:]:
        out.append(("b", core[1:]))
    if tail:
        out.append(("dash", None))
    return out


def spell_items(words):
    """Gradual release, 5/6 shape: word 1 to trace, the next three half shown, the
    rest first letter only.  Our list runs most-recent-first, not easy-to-hard,
    so there is no full 'hardest last' word as on the 3/4 sheets."""
    out = []
    for i, w in enumerate(words):
        if i == 0:
            out.append([("t", w)])
        elif i <= 3:
            out.append(partial(w, (len(w) + 1) // 2))
        else:
            out.append(partial(w, 1))
    return out


TOKEN_RE = re.compile(r"^([\"'(]?)([A-Za-z][A-Za-z'-]*?)([.,!?;:\"')]*)$")


def dictation_items(sentence, targets):
    """Every dictation target is a gap; the rest of the sentence is there to trace."""
    targets = {str(t).casefold() for t in targets}
    items = []
    for tok in str(sentence).split():
        m = TOKEN_RE.match(tok)
        core, tail = (m.group(2), m.group(3)) if m else ("", "")
        # a possessive target (ecosystem's) keeps its 's to trace, the word is the gap
        if core.casefold() not in targets and re.search(r"'s?$", core):
            stem = re.sub(r"'s?$", "", core)
            if stem.casefold() in targets:
                core, tail = stem, core[len(stem):] + tail
        if m and core.casefold() in targets:
            item = [("t", m.group(1))] if m.group(1) else []
            item.append(("b", core))
            if tail:
                item.append(("t", tail))
            items.append(item)
        else:
            items.append([("t", tok)])
    return items


def week_morph(week):
    return next(((s.get("new_morphology") or {}) for s in week.get("sessions", [])
                 if (s.get("new_morphology") or {}).get("morph")), {})


def wants_sheet(session):
    """Which sessions get a sheet: the ones with written spelling or a sentence.
    Wednesday's build is boards and books; a Friday with the worksheet has none.
    A 3-session week's `combined` day carries Thursday's learned words and suffix
    cards, so it gets Thursday's sheet."""
    stype = session.get("type")
    if stype == "week_review":
        return not session.get("friday_review")
    if stype == "combined":
        stype = "apply"
    return stype in ("teach", "spell", "apply") and (
        stype != "apply" or session.get("learned_words") or session.get("dictation"))


def build(week, session, third):
    s = Sheet(third)
    s.ops.append(("title", f"{session['day']} - Term {week.get('term')} Week "
                           f"{week.get('week')} - OG", s.y + 11.0))
    s.y += 22.0
    nm = week_morph(week)
    forms = morph_forms(nm.get("morph", ""))
    stype = session.get("type")
    # the VIC fonts' colon reads as a semicolon, so the card is written "ject = throw"
    head = [("bold", nm.get("morph", "")), ("t", " = " + str(nm.get("meaning", "")))]
    if stype == "teach":
        morph = nm.get("morph", "")
        s.section("New Morphology", [[head], [[("t", morph)], [("t", morph)], [("t", morph)]]])
        s.section("My keyword and what it means", [[]], align="left")
        s.section("My sentence with a " + morph + " word", [[], [], []], align="left")
        return s
    if stype == "spell":
        words = [gap_part(str(w), forms) for w in session.get("words_to_spell_new", [])]
        s.section("New Morphology", [[head]] + s.pack(words))
    if session.get("sound_bank") and stype in ("spell", "week_review"):
        s.section("Sound Bank", s.pack([bank_item(b["morph"]) for b in session["sound_bank"]]))
    wsr = session.get("words_to_spell_review") or []
    if wsr and stype in ("spell", "week_review"):
        s.section("Words to Spell", s.pack(spell_items([e["word"] for e in wsr])))
    lw = session.get("learned_words") or {}
    if lw.get("review"):
        # each word and its gapped copy stay together on one row
        s.section("Learned Words (Review)",
                  s.pack([[("t", w["word"]), ("sp", MIN_GAP * 0.6)]
                          + unfair_gap(w["word"], w.get("unfair")) for w in lw["review"]]))
    if lw.get("new"):
        it = lw["new"]
        word, gap = [("t", it["word"])], unfair_gap(it["word"], it.get("unfair"))
        line = [("line", s.item_width(word) + 10)]
        s.section("Learned Words (New)", s.pack([word, gap, line]))
    dictation = session.get("dictation") or []
    if dictation:
        s.heading("Dictation")
    for dt in dictation:
        for r in s.wrap_words(dictation_items(dt["sentence"], dt.get("targets", [])), s.T):
            s.row(r, align="left")
    return s


def make(week, session, out_path):
    """Write the sheet at the largest 5/6 size that fits one page.
    Returns (third_pt, None) or raises ValueError when nothing fits."""
    for third in THIRDS:
        s = build(week, session, third)
        if s.y + s.T <= BOTTOM:     # room for the last row's descenders
            break
    else:
        raise ValueError(f"{session['day']} enabler does not fit one page even at "
                         f"{THIRDS[-1]} pt thirds - trim the session, never the size")
    doc = fitz.open()
    page = doc.new_page(width=PAGE_W, height=PAGE_H)
    s.render(page)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path), garbage=4, deflate=True)
    return third


# ---------------------------------------------------------------- extension sheets
# v1.9 (James, 30 Oct 2026: "make an extension page equivalent"): the enabler's
# mirror for students ready for more - the same sections in the same order on
# the same lines, each pushed a step (OG_MEGA_PROMPT 3e: same concept, further).
# The enabler gives the word with a gap to trace; the extension gives only the
# meaning or the sound, and the student writes the whole word, then goes one
# step further (a word for each bank part, the parts marked, a sentence).
# Written from the spec, so nobody authors it.

def grid_sum(entry):
    """'pro- means forward + JECT means throw + -or means a thing that: ...' ->
    'pro (forward) + ject (throw) + or (a thing that)'."""
    head = str(entry.get("meaning", "")).split(":")[0]
    parts = []
    for chunk in head.split(" + "):
        m = re.match(r"\s*(\S+)\s+means\s+(.+)", chunk)
        if not m:
            return ""
        parts.append(f"{m.group(1).strip('-').lower()} ({m.group(2).strip()})")
    return " + ".join(parts)


def week_words(week, forms):
    """Every word this week's sessions showed that carries the morpheme."""
    seen = []
    for s_ in week.get("sessions", []):
        pools = [s_.get("words_to_spell_new") or [],
                 [e["word"] for e in s_.get("words_to_spell_review") or []],
                 (s_.get("words_to_read_review") or {}).get("words", []),
                 [w.get("word") if isinstance(w, dict) else w
                  for w in s_.get("words_to_read_new") or []]]
        for pool in pools:
            for w in pool:
                w = str(w)
                if any(f in w.lower() for f in forms) and w not in seen:
                    seen.append(w)
    return seen


def suffix_word(card):
    """The word in a parts-of-speech card's sentence that ends in its suffix."""
    tail = re.sub(r"[^a-z]", "", str(card.get("morph", "")).split("/")[0].lower())
    for tok in re.findall(r"[A-Za-z]+", str(card.get("sentence", ""))):
        if tail and tok.lower().endswith(tail) and len(tok) > len(tail) + 1:
            return tok.lower()
    return ""


def grid_meanings(week):
    """word -> its plain meaning from any teach-day grid ('to throw in')."""
    out = {}
    for s_ in week.get("sessions", []):
        for g in s_.get("words_to_read_new", []):
            if isinstance(g, dict) and ":" in str(g.get("meaning", "")):
                plain = str(g["meaning"]).split(":", 1)[1].split(",")[0].strip()
                if plain:
                    out[str(g["word"]).lower()] = plain
    return out


def example_word(part, week, skip=()):
    """A word the class met this week that carries `part`, for the answers page."""
    forms = morph_forms(part)
    for w in week_words(week, forms):
        if w.lower() not in skip:
            return w
    return ""


def extension_sheet(week, session):
    """The enabler's sections, in the same order, each pushed a step further, on
    the school worksheet layout (MEGA_PROMPT 86): numbered cards, ruled lines and
    tables, not dotted thirds (James, 31 Oct 2026: extension students write in
    their own hand, so the sheet need not look like the enabler).  Where the
    enabler prints the word to trace with a gap, the extension prints only what
    the student needs to start.  Answers ride on the items, so the same spec
    draws the answer key.  Returns the worksheet resource."""
    nm = week_morph(week)
    morph, meaning = nm.get("morph", ""), str(nm.get("meaning", ""))
    stype = session.get("type")
    if stype == "combined":
        stype = "apply"     # the same written work as Thursday (wants_sheet)
    sections = []

    def blank_table(prompt, n, answers):
        """One row of n numbered boxes: parts or words written as they are read."""
        return {"kind": "table", "prompt": prompt, "span": "full",
                "columns": [str(i) for i in range(1, n + 1)], "rows": [[""] * n], "rowH": 28,
                "answers": [answers]}

    if stype == "teach":
        grid = [g for g in session.get("words_to_read_new", []) if isinstance(g, dict)]
        sums = sorted(((g["word"], grid_sum(g)) for g in grid if grid_sum(g)),
                      key=lambda ws: len(ws[0]))
        means = grid_meanings(week)
        sections.append({"title": "New Morphology", "columns": 1,
                         "intro": f"{morph} = {meaning}", "items": [
            {"kind": "table", "prompt": f"Write three {morph} words and what each means.",
             "columns": [f"{morph} word", "What it means"], "widths": [1, 2],
             "rows": [["", ""] for _ in range(3)], "rowH": 28,
             "answers": [[w, means.get(w.lower(), "")] for w, _ in sums[:3]]},
            {"prompt": "My keyword and what it means.", "lines": 1,
             "answer": "Any real " + morph + " word with a true meaning."},
            {"prompt": f"Build the longest {morph} word you can. Write its word sum.",
             "lines": 2, "answer": "e.g. " + "; ".join(f"{w} = {sm}" for w, sm in sums[-2:])},
            {"kind": "open", "prompt": f"Write one sentence with TWO {morph} words.",
             "lines": 3, "answer": f"Check: two {morph} words used correctly, and the "
                                   f"sentence shows the meaning ({morph} = {meaning})."},
        ]})
        return sections

    if stype == "spell":
        means = grid_meanings(week)
        new = [str(w) for w in session.get("words_to_spell_new", [])]
        sections.append({"title": "New Morphology", "columns": 1,
                         "intro": f"{morph} = {meaning}", "items": [
            {"kind": "table", "prompt": f"Write the {morph} word for each meaning.",
             "columns": ["Meaning", "Word"], "widths": [3, 2],
             "rows": [[means.get(w.lower(), w[:1] + "..."), ""] for w in new], "rowH": 28,
             "answers": [[None, w] for w in new]}]})
    if session.get("sound_bank") and stype in ("spell", "week_review"):
        bank = [str(b["morph"]) for b in session["sound_bank"]]
        examples = [f"{b} (e.g. {example_word(b, week)})" for b in bank
                    if example_word(b, week)]
        sections.append({"title": "Sound Bank", "columns": 1, "items": [
            blank_table("Write each part as it is read.", len(bank), bank),
            {"prompt": "Pick two parts. Write a word that uses each.", "lines": 1,
             "answer": "; ".join(examples)}]})
    wsr = [e["word"] for e in session.get("words_to_spell_review") or []]
    if wsr and stype in ("spell", "week_review"):
        sections.append({"title": "Words to Spell", "columns": 1, "items": [
            {"kind": "table", "prompt": "Spell each word as it is read. Then write it "
                                        "again with its parts marked with /.",
             "columns": ["", "Spell it", "Its parts, marked with /"], "widths": [1, 6, 6],
             "rows": [[str(i), "", ""] for i in range(1, len(wsr) + 1)], "rowH": 28,
             "answers": [[None, w, ""] for w in wsr]}]})
    cards = (session.get("pos_drill") or {}).get("cards", []) if stype == "apply" else []
    if cards:
        # Thursday's real work: a word for each suffix card, and the job it does (3g)
        sections.append({"title": "Suffix jobs", "columns": 1, "items": [
            {"kind": "table", "prompt": "Write a new word for each suffix. Then write its "
                                        "job: noun, verb or adjective.",
             "columns": ["Suffix", "My word", "Its job"], "widths": [1, 3, 2],
             "rows": [[str(c.get("morph", "")), "", ""] for c in cards], "rowH": 28,
             "answers": [[None, f"e.g. {suffix_word(c)}", c.get("pos_answer", "")]
                         for c in cards]},
            {"kind": "open", "prompt": "Write one sentence using two of your suffix words.",
             "lines": 2, "answer": "Check: two of the student's suffix words, each doing "
                                   "the job they wrote. Any real word with the right job counts."}]})
    lw = session.get("learned_words") or {}
    words = [w["word"] for w in lw.get("review", [])]
    if lw.get("new"):
        words.append(lw["new"]["word"])
    if words:
        new_word = (lw.get("new") or {}).get("word")
        grammar = ""
        ov = (week.get("overview") or {}).get(session.get("day"), {})
        if stype == "apply" and isinstance(ov.get("grammar"), str) \
                and ov["grammar"] not in ("Not today", ""):
            grammar = ov["grammar"].split(":")[0].strip().lower()
        items = [blank_table("Write each word from memory as it is read.", len(words), words)]
        if new_word:
            items.append({"kind": "open", "lines": 1 if stype == "spell" else 2,
                          "prompt": f"Use {new_word} in a sentence"
                                    + (f" with a {grammar}." if grammar else "."),
                          "answer": f"Check: {new_word} spelled right"
                                    + (f" and the {grammar} used correctly." if grammar else ".")})
        sections.append({"title": "Learned Words", "columns": 1, "items": items})
    dictation = session.get("dictation") or []
    if dictation:
        sections.append({"title": "Dictation", "columns": 1, "items": [
            {"prompt": f"Write both sentences as they are read. Then add a {morph} word "
                       "to sentence 1 with a ^.", "lines": 2 * len(dictation),
             "answer": " / ".join(dt["sentence"] for dt in dictation)
                       + f" Check: the added {morph} word fits and is spelled right."}]})
    return sections


def make_extension(week, session, out_path):
    """The student sheet (one page, or two as one double-sided sheet), then its
    answer key (teacher copy), on the school worksheet layout rendered by
    render_extension.js.  Returns True, or None when the day has no sheet."""
    if not wants_sheet(session):
        return None
    term, wk = week.get("term"), week.get("week")
    job = {"resource": {"title": f"{session['day']} OG Extension",
                        "label": f"{session['day']} OG Extension",
                        "sections": extension_sheet(week, session)},
           "out": str(Path(out_path).resolve()), "week": wk,
           "info": f"Term {term} Week {wk}", "footer": f"OG Extension | Term {term} Week {wk}"}
    out_path.parent.mkdir(parents=True, exist_ok=True)
    root = Path(__file__).resolve().parent
    with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False) as fh:
        json.dump(job, fh)
    try:
        run = subprocess.run(["node", str(root / "render_extension.js"), fh.name],
                             cwd=root.parent, capture_output=True, text=True)
    finally:
        Path(fh.name).unlink(missing_ok=True)
    if run.returncode != 0:
        raise ValueError(f"{session['day']} extension sheet: "
                         f"{(run.stderr or run.stdout).strip().splitlines()[-1]}")
    return True


def extension_file_name(number, day):
    return f"{number}. {day} extension.pdf"


def file_name(number, day):
    return f"{number}. {day} enabler.pdf"


def week_sheet_names(week):
    """[(day, file name)] for every session that gets a sheet, in session order."""
    return [(s["day"], file_name(n, s["day"]))
            for n, s in enumerate(week.get("sessions", []), 1) if wants_sheet(s)]


if __name__ == "__main__":
    spec = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    for n, session in enumerate(spec["sessions"], 1):
        if wants_sheet(session):
            name = file_name(n, session["day"])
            third = make(spec, session, Path(sys.argv[2]) / name)
            print(f"{name}: thirds {third} pt ({third / 72 * 25.4:.1f} mm)")
