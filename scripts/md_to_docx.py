"""Convert a unit Markdown document into a Word document.

Written for AUSLAN_1_UNIT_PROMPT.md, which delivers as a Word file. Handles
# / ## headings, \newpage page breaks, markdown tables (with <br> and
**bold** inside cells), - bullets, inline bold, and a Word TOC field.

    python scripts/md_to_docx.py in.md out.docx ["Title"] ["Subtitle"] ["Meta"]
                                 [--landscape] [--plain] [--fontsize=9]

Title defaults to the output filename stem. Subtitle and Meta are optional
title-page lines. --landscape gives A4 landscape (for the teacher layer's
at-a-glance, lesson pages and weekly planner). --plain drops the title page and
the contents field, for short documents that start at their first heading.
--fontsize sets the body point size; column widths scale with it. There is no pandoc in this environment; python-docx is the
route. QA the result with LibreOffice + PyMuPDF (pdftoppm is not installed):

    soffice --headless --convert-to pdf out.docx
    python -c "import fitz; [p.get_pixmap(dpi=95).save(f'p{i}.png') for i,p in enumerate(fitz.open('out.pdf'))]"
"""
import os
import re
import sys
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_BREAK, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

argv = [a for a in sys.argv[1:] if not a.startswith('--')]
FLAGS = {a for a in sys.argv[1:] if a.startswith('--')}
LANDSCAPE = '--landscape' in FLAGS
PLAIN = '--plain' in FLAGS
FONT_PT = 11.0
for f in FLAGS:
    if f.startswith('--fontsize'):
        FONT_PT = float(f.split('=', 1)[1])

SRC = argv[0]
OUT = argv[1]
TITLE = argv[2] if len(argv) > 2 else os.path.splitext(os.path.basename(OUT))[0]
SUBTITLE = argv[3] if len(argv) > 3 else ''
META = argv[4] if len(argv) > 4 else ''

# usable text width, 1.5cm margins each side
USABLE_CM = 26.7 if LANDSCAPE else 18.0
NEWPAGE = chr(92) + 'newpage'


def add_toc_field(par):
    r = par.add_run()
    fld = OxmlElement('w:fldChar')
    fld.set(qn('w:fldCharType'), 'begin')
    r._r.append(fld)
    r2 = par.add_run()
    it = OxmlElement('w:instrText')
    it.set(qn('xml:space'), 'preserve')
    it.text = r'TOC \o "1-2" \h \z \u'
    r2._r.append(it)
    r3 = par.add_run()
    sep = OxmlElement('w:fldChar')
    sep.set(qn('w:fldCharType'), 'separate')
    r3._r.append(sep)
    r4 = par.add_run('Right-click here and choose Update Field to build the contents list.')
    r4.italic = True
    r5 = par.add_run()
    end = OxmlElement('w:fldChar')
    end.set(qn('w:fldCharType'), 'end')
    r5._r.append(end)


INLINE = re.compile(r'\*\*(.+?)\*\*|`([^`]+)`')


def emit_inline(par, text):
    """Write text into a paragraph, honouring **bold**, `code` and <br> line breaks."""
    for bi, chunk in enumerate(text.split('<br>')):
        if bi:
            par.add_run().add_break(WD_BREAK.LINE)
        pos = 0
        for m in INLINE.finditer(chunk):
            if m.start() > pos:
                par.add_run(chunk[pos:m.start()])
            if m.group(1) is not None:
                par.add_run(m.group(1)).bold = True
            else:
                run = par.add_run(m.group(2))
                run.font.name = 'Consolas'
                run.font.size = Pt(FONT_PT - 1)
            pos = m.end()
        if pos < len(chunk):
            par.add_run(chunk[pos:])


def split_row(line):
    return [c.strip() for c in line.strip().strip('|').split('|')]


def is_sep(line):
    return bool(re.match(r'^\|[\s:|-]+\|$', line.strip())) and '-' in line


CHAR_CM = 0.20 * FONT_PT / 11.0   # approx width of one Calibri character at FONT_PT
PAD_CM = 0.45                     # left + right cell padding


def col_widths(rows, ncol):
    """Give every column at least the width of its longest unbreakable word,
    then share what is left in proportion to how much text the column holds."""
    floors, weights = [], []
    for c in range(ncol):
        longest_word, longest_cell = 1, 1
        for ri, r in enumerate(rows):
            txt = r[c].replace('<br>', ' ').replace('**', '').replace('`', '')
            longest_cell = max(longest_cell, len(txt))
            # the header row renders bold, which is about 15 per cent wider
            scale = 1.15 if ri == 0 else 1.0
            for w in txt.split():
                longest_word = max(longest_word, len(w) * scale)
        floors.append(min(longest_word * CHAR_CM + PAD_CM, USABLE_CM / 2))
        weights.append(min(longest_cell, 120))

    spare = USABLE_CM - sum(floors)
    if spare <= 0:                       # floors alone overflow: scale them down
        k = USABLE_CM / sum(floors)
        return [Cm(f * k) for f in floors]
    total = sum(weights)
    return [Cm(f + spare * w / total) for f, w in zip(floors, weights)]


def shade(cell, colour):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement('w:shd')
    sh.set(qn('w:val'), 'clear')
    sh.set(qn('w:color'), 'auto')
    sh.set(qn('w:fill'), colour)
    tcPr.append(sh)


def set_repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement('w:tblHeader')
    el.set(qn('w:val'), 'true')
    trPr.append(el)


doc = Document()

sec = doc.sections[0]
if LANDSCAPE:
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
else:
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
sec.left_margin = sec.right_margin = Cm(1.5)
sec.top_margin = sec.bottom_margin = Cm(1.0 if LANDSCAPE else 1.8)

normal = doc.styles['Normal']
normal.font.name = 'Calibri'
normal.font.size = Pt(FONT_PT)
normal.paragraph_format.space_after = Pt(6 if FONT_PT >= 11 else 3)
normal.paragraph_format.space_before = Pt(0)

for name, size, colour in (('Heading 1', FONT_PT + 9, RGBColor(0x1F, 0x3B, 0x63)),
                           ('Heading 2', FONT_PT + 4, RGBColor(0x1F, 0x3B, 0x63)),
                           ('Heading 3', FONT_PT + 1, RGBColor(0x1F, 0x3B, 0x63))):
    st = doc.styles[name]
    st.font.name = 'Calibri'
    st.font.size = Pt(size)
    st.font.bold = True
    st.font.color.rgb = colour
    st.paragraph_format.space_before = Pt(12)
    st.paragraph_format.space_after = Pt(6)
    st.paragraph_format.keep_with_next = True

# ---- title page ----
def build_title_page():
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t.add_run(TITLE)
    r.bold = True
    r.font.size = Pt(28)
    for text, size in ((SUBTITLE, 16), (META, 12)):
        if not text:
            continue
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(text).font.size = Pt(size)
    t4 = doc.add_paragraph()
    t4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = t4.add_run('(c) 2026 James Hooke. Confidential. Internal use only. Not for redistribution.')
    r.font.size = Pt(9)
    r.italic = True
    doc.add_paragraph()
    h = doc.add_paragraph()
    r = h.add_run('Contents')
    r.bold = True
    r.font.size = Pt(16)
    add_toc_field(doc.add_paragraph())
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


if not PLAIN:
    build_title_page()

lines = open(SRC, encoding='utf-8').read().split('\n')
i = 0
pending_break = False

while i < len(lines):
    line = lines[i]
    stripped = line.strip()

    if stripped == NEWPAGE:
        pending_break = True
        i += 1
        continue

    if not stripped:
        i += 1
        continue

    if pending_break:
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        pending_break = False

    m = re.match(r'^(#{1,3}) (.+)$', stripped)
    if m:
        doc.add_heading(m.group(2).strip(), level=len(m.group(1)))
        i += 1
        continue

    if stripped.startswith('|'):
        block = []
        while i < len(lines) and lines[i].strip().startswith('|'):
            if not is_sep(lines[i]):
                block.append(split_row(lines[i]))
            i += 1
        ncol = max(len(r) for r in block)
        block = [r + [''] * (ncol - len(r)) for r in block]
        widths = col_widths(block, ncol)
        tbl = doc.add_table(rows=0, cols=ncol)
        tbl.style = 'Table Grid'
        tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
        tbl.autofit = False
        layout = OxmlElement('w:tblLayout')
        layout.set(qn('w:type'), 'fixed')
        tbl._tbl.tblPr.append(layout)
        grid = tbl._tbl.find(qn('w:tblGrid'))
        for gc, w in zip(grid.findall(qn('w:gridCol')), widths):
            gc.set(qn('w:w'), str(int(w.twips)))
        for ri, rowdata in enumerate(block):
            row = tbl.add_row()
            for ci, celltext in enumerate(rowdata):
                cell = row.cells[ci]
                cell.width = widths[ci]
                par = cell.paragraphs[0]
                par.paragraph_format.space_after = Pt(2)
                par.paragraph_format.space_before = Pt(2)
                emit_inline(par, celltext)
                if ri == 0:
                    shade(cell, 'DCE4F0')
                    for rr in par.runs:
                        rr.bold = True
            if ri == 0:
                set_repeat_header(row)
        # no spacer paragraph when a page break or the end of file follows:
        # on a table that exactly fills the page it would create a blank page
        nxt = next((x.strip() for x in lines[i:] if x.strip()), '')
        if nxt and nxt != NEWPAGE:
            doc.add_paragraph()
        continue

    if stripped.startswith('- ') or stripped.startswith('* '):
        while i < len(lines) and lines[i].strip()[:2] in ('- ', '* '):
            par = doc.add_paragraph(style='List Bullet')
            par.paragraph_format.space_after = Pt(2)
            emit_inline(par, lines[i].strip()[2:])
            i += 1
        continue

    par = doc.add_paragraph()
    emit_inline(par, stripped)
    i += 1

doc.save(OUT)
print('written', OUT)
