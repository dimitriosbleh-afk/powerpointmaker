#!/usr/bin/env python3
"""Build Orton-Gillingham session decks from the master template.

Usage:
    python3 og_planner/build_og_week.py og_planner/weeks/<week>.json [--only Monday] [--outdir output]

Reads the week spec JSON (schema documented in IMPORTANT/OG_MEGA_PROMPT.md),
clones slides from og_planner/OG_MASTER_TEMPLATE.pptx at raw-XML level so that
fonts, positions, backgrounds and click animations are preserved exactly, and
writes one PPTX per session into output/<unit_folder>/.

The master template must not be edited. All content decisions live in the spec.
"""

import copy
import json
import math
import re
import shutil
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

from lxml import etree

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def q(tag):
    pfx, local = tag.split(":")
    return "{%s}%s" % (NS[pfx], local)


ROOT_DIR = Path(__file__).resolve().parent
MASTER = ROOT_DIR / "OG_MASTER_TEMPLATE.pptx"

# ---------------------------------------------------------------- template map
# 1-based slide indices in the master template.
T = {
    "title": 1,
    "overview": 2,
    "hdr_morph_review": 3,
    "card_review_root": 4,
    "card_review_prefix": 5,
    "card_review_suffix": 6,
    "hdr_wtr_review": 7,
    "wtr_review_table": 8,
    "hdr_sound_bank": 9,
    "sound_bank": 10,
    "hdr_wts_review": 11,
    "spell_word": 12,
    "hdr_new_morph": 13,
    "card_new_root": 14,
    "card_new_prefix": 15,
    "card_new_suffix": 16,
    "hdr_wtr_new": 17,
    "wtr_new_grid": 18,
    "hdr_wts_new": 19,
    "wts_new_grid": 20,
    "hdr_review_lw": 21,
    "learned_word": 22,
    "hdr_new_lw": 24,
    "learned_word_new": 25,
    "hdr_dictation": 26,
    "dictation": 27,
    "hdr_grammar": 29,
    "gradual_release": {"i_do": 30, "we_do": 31, "you_do": 32},
}

# Shape ids inside each template slide (stable - taken from the master).
SHAPE = {
    "title_heading": 6391,       # 'OG Planner'
    "title_sub": 6392,           # 'Enrichment / Term 3 Week XYZ'
    "overview_table": 6406,
    "card_review_word": {4: 6431, 5: 6439, 6: 6447},
    "wtr_review_tbl": 6465,
    "sound_boxes": [6482, 6483, 6484, 6486, 6485, 6487, 6488, 6489, 6490],
    # row-major order: (r1: x0.74,3.84,6.93) (r2: 0.74=6486, 3.84=6485, 6.93=6487) (r3)
    "spell_word_box": 6505,
    "hdr_new_morph_txt": 6512,
    "card_new_word": {14: 6523, 15: 6531, 16: 6539},
    "hdr_wtr_new_txt": 6544,
    "wtr_cols": [6556, 6555, 6554],  # left, middle, right
    "hdr_wts_new_txt": 6561,
    "wts_boxes": [6571, 6572],       # left (2 paras), right (2 paras)
    "lw_word": {22: 6598, 25: 6634},
    "dictation_title": 6660,
    "dictation_sentence": 6663,
    "gr_title": {30: 6687, 31: 6693, 32: 6699},
    "gr_body": {30: 6688, 31: 6694, 32: 6700},
}

MORPH_FILL = {"root": "FFF2CC", "prefix": "ABEA90", "suffix": "E06666"}
CARD_REVIEW = {"root": 4, "prefix": 5, "suffix": 6}
CARD_NEW = {"root": 14, "prefix": 15, "suffix": 16}

GRAMMAR_PURPLE = "5E58A6"   # sampled from the template's Grammar section header
GRAMMAR_TINT = "EFEEF8"
MORPH_MAGENTA = "A64D79"    # sampled from the New Morphology section header
MORPH_TINT = "F6EBF1"
CHARCOAL = "26262B"
PUNCT_RED = "C00000"

WARNINGS = []


def warn(msg):
    WARNINGS.append(msg)
    print(f"  WARN: {msg}")


# ---------------------------------------------------------------- text helpers
def fit_font(text, width_in, base_pt, min_pt=20, char_em=0.62):
    """Largest size <= base_pt at which `text` fits on one line in width_in."""
    n = max(1, len(text))
    fit = int(width_in * 72 / (char_em * n))
    return max(min_pt, min(base_pt, fit))


def fit_font_block(text, width_in, height_in, base_pt, min_pt=20, char_em=0.55,
                   line_factor=1.32):
    """Largest size <= base_pt at which wrapped `text` fits in the box."""
    for sz in range(base_pt, min_pt - 1, -2):
        cpl = max(1, int(width_in * 72 / (char_em * sz)))
        lines = math.ceil(len(text) / cpl)
        if lines * sz * line_factor / 72 <= height_in:
            return sz
    return min_pt


def find_sp(root, spid):
    for sp in root.iter(q("p:sp")):
        nv = sp.find(q("p:nvSpPr") + "/" + q("p:cNvPr"))
        if nv is not None and nv.get("id") == str(spid):
            return sp
    raise KeyError(f"shape id {spid} not found")


def txbody(sp):
    tb = sp.find(q("p:txBody"))
    if tb is None:
        raise KeyError("shape has no txBody")
    return tb


def set_runs(sp, paragraphs, size_pt=None):
    """Replace a shape's paragraphs.

    paragraphs: list of paragraphs; each paragraph is a list of runs;
    each run is (text, color_hex_or_None, bold_or_None, underline_or_None).
    The first existing run's rPr is cloned as base formatting for every run;
    the first existing paragraph's pPr is kept for all paragraphs.
    """
    tb = txbody(sp)
    paras = tb.findall(q("a:p"))
    if not paras:
        raise KeyError("no paragraphs in shape")
    proto_p = paras[0]
    proto_ppr = proto_p.find(q("a:pPr"))
    proto_rpr = None
    for r_el in proto_p.findall(q("a:r")):
        proto_rpr = r_el.find(q("a:rPr"))
        if proto_rpr is not None:
            break
    if proto_rpr is None:  # look in later paragraphs
        for p_el in paras[1:]:
            for r_el in p_el.findall(q("a:r")):
                proto_rpr = r_el.find(q("a:rPr"))
                break
            if proto_rpr is not None:
                break
    for p_el in paras:
        tb.remove(p_el)
    for runs in paragraphs:
        p_el = etree.SubElement(tb, q("a:p"))
        if proto_ppr is not None:
            p_el.append(copy.deepcopy(proto_ppr))
        for text, color, bold, underline in runs:
            r_el = etree.SubElement(p_el, q("a:r"))
            if proto_rpr is not None:
                rpr = copy.deepcopy(proto_rpr)
            else:
                rpr = etree.Element(q("a:rPr"), lang="en-US")
            if size_pt is not None:
                rpr.set("sz", str(int(size_pt * 100)))
            if bold is not None:
                rpr.set("b", "1" if bold else "0")
            if underline:
                rpr.set("u", "sng")
            if color is not None:
                for fill in rpr.findall(q("a:solidFill")):
                    rpr.remove(fill)
                fill = etree.Element(q("a:solidFill"))
                clr = etree.SubElement(fill, q("a:srgbClr"))
                clr.set("val", color)
                fill.insert(0, clr)
                rpr.insert(0, fill)
            r_el.append(rpr)
            t_el = etree.SubElement(r_el, q("a:t"))
            t_el.text = text
        if not runs:
            if proto_rpr is not None:
                end = copy.deepcopy(proto_rpr)
                end.tag = q("a:endParaRPr")
                p_el.append(end)


def set_text(sp, text, size_pt=None, color=None):
    set_runs(sp, [[(text, color, None, None)]], size_pt=size_pt)


def replace_run_text(sp, old, new):
    """Replace `old` with `new` inside existing runs, keeping formatting."""
    hit = False
    for t_el in sp.iter(q("a:t")):
        if t_el.text and old in t_el.text:
            t_el.text = t_el.text.replace(old, new)
            hit = True
    return hit


def set_solid_fill(sp, hexval):
    sppr = sp.find(q("p:spPr"))
    for fill in sppr.findall(q("a:solidFill")):
        sppr.remove(fill)
    fill = etree.Element(q("a:solidFill"))
    clr = etree.SubElement(fill, q("a:srgbClr"))
    clr.set("val", hexval)
    ln = sppr.find(q("a:ln"))
    if ln is not None:
        ln.addprevious(fill)
    else:
        sppr.append(fill)


def add_textbox(root, sid, x, y, w, h, lines, size_pt, color=CHARCOAL, bold=False,
                fill=None, rounded=False, align="ctr", anchor="ctr", italic=False):
    """Add a generated text box/card to a slide. `**bold**` markup supported."""
    a, p = NS["a"], NS["p"]
    geom = "roundRect" if rounded else "rect"
    fill_xml = (f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
                if fill else "<a:noFill/>")
    paras = []
    for ln in lines:
        runs = ""
        for i, seg in enumerate(re.split(r"\*\*(.*?)\*\*", ln)):
            if seg == "":
                continue
            b = "1" if (bold or i % 2 == 1) else "0"
            ital = ' i="1"' if italic else ""
            runs += (f'<a:r><a:rPr lang="en-AU" sz="{int(size_pt * 100)}" b="{b}"{ital}>'
                     f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
                     f'<a:latin typeface="Lexend"/></a:rPr><a:t>{escape(seg)}</a:t></a:r>')
        if not runs:
            runs = f'<a:endParaRPr lang="en-AU" sz="{int(size_pt * 100)}"/>'
        paras.append(f'<a:p><a:pPr algn="{align}"/>{runs}</a:p>')
    xml = (
        f'<p:sp xmlns:p="{p}" xmlns:a="{a}">'
        f'<p:nvSpPr><p:cNvPr id="{sid}" name="OGGen{sid}"/><p:cNvSpPr txBox="1"/>'
        f'<p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{int(x * 914400)}" y="{int(y * 914400)}"/>'
        f'<a:ext cx="{int(w * 914400)}" cy="{int(h * 914400)}"/></a:xfrm>'
        f'<a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>{fill_xml}'
        f'<a:ln><a:noFill/></a:ln></p:spPr>'
        f'<p:txBody><a:bodyPr anchor="{anchor}" wrap="square" lIns="91440" '
        f'tIns="45720" rIns="91440" bIns="45720"/><a:lstStyle/>{"".join(paras)}'
        f'</p:txBody></p:sp>')
    sptree = root.find(q("p:cSld") + "/" + q("p:spTree"))
    sptree.append(etree.fromstring(xml))


def table_cells(root, spid):
    for gf in root.iter(q("p:graphicFrame")):
        nv = gf.find(q("p:nvGraphicFramePr") + "/" + q("p:cNvPr"))
        if nv is not None and nv.get("id") == str(spid):
            return gf.findall(".//" + q("a:tr"))
    raise KeyError(f"table {spid} not found")


def set_cell_paras(tc, lines, size_pt=None):
    """Replace a table cell's paragraphs with `lines` (list of strings)."""
    tx = tc.find(q("a:txBody"))
    paras = tx.findall(q("a:p"))
    proto_p = paras[0]
    proto_ppr = proto_p.find(q("a:pPr"))
    proto_rpr = None
    for p_el in paras:
        for r_el in p_el.findall(q("a:r")):
            proto_rpr = r_el.find(q("a:rPr"))
            break
        if proto_rpr is not None:
            break
    for p_el in paras:
        tx.remove(p_el)
    for line in lines:
        p_el = etree.SubElement(tx, q("a:p"))
        if proto_ppr is not None:
            p_el.append(copy.deepcopy(proto_ppr))
        r_el = etree.SubElement(p_el, q("a:r"))
        rpr = copy.deepcopy(proto_rpr) if proto_rpr is not None else etree.Element(q("a:rPr"), lang="en-US")
        if size_pt is not None:
            rpr.set("sz", str(int(size_pt * 100)))
        r_el.append(rpr)
        t_el = etree.SubElement(r_el, q("a:t"))
        t_el.text = line


# ---------------------------------------------------------------- timing regen
def regen_timing(root, targets):
    """Rebuild the click-to-reveal sequence.

    targets: list of (spid, para_index_or_None). One click per entry, in order.
    Uses the slide's existing timing block as the structural pattern.
    """
    timing = root.find(q("p:timing"))
    if timing is None:
        raise KeyError("slide has no timing block to use as pattern")
    seq_ctn = timing.find(".//" + q("p:seq") + "/" + q("p:cTn"))
    child_lst = seq_ctn.find(q("p:childTnLst"))
    pars = child_lst.findall(q("p:par"))
    pattern = copy.deepcopy(pars[0])
    for p_el in pars:
        child_lst.remove(p_el)
    ids = [2000]

    def renumber(el):
        for ctn in el.iter(q("p:cTn")):
            if ctn.get("id"):
                ids[0] += 1
                ctn.set("id", str(ids[0]))

    for spid, para in targets:
        node = copy.deepcopy(pattern)
        renumber(node)
        tgt = node.find(".//" + q("p:spTgt"))
        tgt.set("spid", str(spid))
        txel = tgt.find(q("p:txEl"))
        if para is None:
            if txel is not None:
                tgt.remove(txel)
        else:
            if txel is None:
                txel = etree.SubElement(tgt, q("p:txEl"))
                etree.SubElement(txel, q("p:pRg"))
            prg = txel.find(q("p:pRg"))
            if prg is None:
                for ch in list(txel):
                    txel.remove(ch)
                prg = etree.SubElement(txel, q("p:pRg"))
            prg.set("st", str(para))
            prg.set("end", str(para))
        child_lst.append(node)
    # bldLst: ensure one bldP per shape used, paragraph builds
    bld = timing.find(q("p:bldLst"))
    if bld is not None:
        for b in list(bld):
            bld.remove(b)
        for spid in dict.fromkeys(s for s, _ in targets):
            bp = etree.SubElement(bld, q("p:bldP"))
            bp.set("spid", str(spid))
            bp.set("grpId", "0")
            if any(p is not None for s, p in targets if s == spid):
                bp.set("build", "p")


# ---------------------------------------------------------------- notes
def set_notes(notes_root, text):
    """Replace the body placeholder text of a notes slide.

    Lightweight markup: `**bold**` segments render as bold runs. Notes are read
    live on an iPad, so bold section labels + blank lines between blocks are the
    only structure - never markdown bullets.
    """
    body = None
    for sp in notes_root.iter(q("p:sp")):
        ph = sp.find(".//" + q("p:ph"))
        if ph is not None and ph.get("type") == "body":
            body = sp
            break
    if body is None:
        raise KeyError("notes body placeholder not found")
    tb = txbody(body)
    for p_el in tb.findall(q("a:p")):
        tb.remove(p_el)
    for line in text.split("\n"):
        line = re.sub(r" {2,}", " ", line.rstrip())
        p_el = etree.SubElement(tb, q("a:p"))
        # kill the notes master's default bullets - they read as dot-point
        # walls on the iPad speaker-notes view
        ppr = etree.SubElement(p_el, q("a:pPr"))
        etree.SubElement(ppr, q("a:buNone"))
        for i, seg in enumerate(re.split(r"\*\*(.*?)\*\*", line)):
            if seg == "":
                continue
            r_el = etree.SubElement(p_el, q("a:r"))
            rpr = etree.SubElement(r_el, q("a:rPr"), lang="en-AU")
            if i % 2 == 1:  # odd split segments were inside ** **
                rpr.set("b", "1")
            t_el = etree.SubElement(r_el, q("a:t"))
            t_el.text = seg


def tidy_notes_whitespace(notes_root):
    """Normalise fixed template notes (double spaces read badly on the iPad)."""
    for t_el in notes_root.iter(q("a:t")):
        if t_el.text:
            t_el.text = re.sub(r" {2,}", " ", t_el.text)


# ---------------------------------------------------------------- package
class Package:
    def __init__(self, path):
        with zipfile.ZipFile(path) as z:
            self.parts = {n: z.read(n) for n in z.namelist()}
        self.pres = etree.fromstring(self.parts["ppt/presentation.xml"])
        self.pres_rels = etree.fromstring(self.parts["ppt/_rels/presentation.xml.rels"])
        # ordered template slide part names
        rid_to_target = {
            rel.get("Id"): rel.get("Target")
            for rel in self.pres_rels.findall(q("rel:Relationship"))
        }
        self.slide_parts = []
        for sld in self.pres.findall(".//" + q("p:sldId")):
            rid = sld.get(q("r:id"))
            self.slide_parts.append("ppt/" + rid_to_target[rid].lstrip("/").replace("../", ""))

    def slide_xml(self, idx1):
        return self.parts[self.slide_parts[idx1 - 1]]

    def slide_rels(self, idx1):
        name = self.slide_parts[idx1 - 1]
        rn = name.rsplit("/", 1)
        return self.parts.get(f"{rn[0]}/_rels/{rn[1]}.rels")

    def notes_name(self, idx1):
        rels = self.slide_rels(idx1)
        if not rels:
            return None
        root = etree.fromstring(rels)
        for rel in root.findall(q("rel:Relationship")):
            if rel.get("Type").endswith("/notesSlide"):
                tgt = rel.get("Target").replace("../", "ppt/")
                return tgt
        return None


class DeckBuilder:
    """Assembles one output PPTX from template slide clones."""

    def __init__(self, pkg):
        self.pkg = pkg
        self.slides = []  # (slide_xml_root, rels_root, notes_root_or_None)

    def add(self, template_idx, fill=None, notes=None):
        """Clone template slide `template_idx`; apply fill(root) mutations."""
        root = etree.fromstring(self.pkg.slide_xml(template_idx))
        rels_bytes = self.pkg.slide_rels(template_idx)
        rels = etree.fromstring(rels_bytes) if rels_bytes else None
        notes_name = self.pkg.notes_name(template_idx)
        notes_root = None
        if notes is not None and notes_name is None:
            # borrow a notes slide from a template slide that has one
            notes_name = self.pkg.notes_name(T["spell_word"])
        if notes_name:
            notes_root = etree.fromstring(self.pkg.parts[notes_name])
            notes_rels = etree.fromstring(
                self.pkg.parts[notes_name.replace("notesSlides/", "notesSlides/_rels/") + ".rels"]
            )
        else:
            notes_rels = None
        if fill:
            fill(root)
        if notes is not None and notes_root is not None:
            set_notes(notes_root, notes)
        elif notes_root is not None:
            tidy_notes_whitespace(notes_root)
        self.slides.append((root, rels, notes_root, notes_rels))

    def write(self, out_path):
        parts = {}
        # copy every part except slides/notesSlides and their rels and ctypes/pres
        skip_prefixes = ("ppt/slides/", "ppt/notesSlides/")
        for name, data in self.pkg.parts.items():
            if name.startswith(skip_prefixes):
                continue
            if name in ("[Content_Types].xml", "ppt/presentation.xml",
                        "ppt/_rels/presentation.xml.rels"):
                continue
            parts[name] = data

        # slides + notes
        for i, (root, rels, notes_root, notes_rels) in enumerate(self.slides, 1):
            sname = f"ppt/slides/slide{i}.xml"
            parts[sname] = etree.tostring(root, xml_declaration=True,
                                          encoding="UTF-8", standalone=True)
            if rels is not None:
                rels2 = copy.deepcopy(rels)
                for rel in rels2.findall(q("rel:Relationship")):
                    if rel.get("Type").endswith("/notesSlide"):
                        if notes_root is not None:
                            rel.set("Target", f"../notesSlides/notesSlide{i}.xml")
                        else:
                            rels2.remove(rel)
                parts[f"ppt/slides/_rels/slide{i}.xml.rels"] = etree.tostring(
                    rels2, xml_declaration=True, encoding="UTF-8", standalone=True)
            if notes_root is not None:
                nname = f"ppt/notesSlides/notesSlide{i}.xml"
                parts[nname] = etree.tostring(notes_root, xml_declaration=True,
                                              encoding="UTF-8", standalone=True)
                nrels = copy.deepcopy(notes_rels)
                for rel in nrels.findall(q("rel:Relationship")):
                    if rel.get("Type").endswith("/slide"):
                        rel.set("Target", f"../slides/slide{i}.xml")
                parts[f"ppt/notesSlides/_rels/notesSlide{i}.xml.rels"] = etree.tostring(
                    nrels, xml_declaration=True, encoding="UTF-8", standalone=True)

        # presentation.xml
        pres = copy.deepcopy(self.pkg.pres)
        sldlst = pres.find(q("p:sldIdLst"))
        for sld in list(sldlst):
            sldlst.remove(sld)
        for i in range(1, len(self.slides) + 1):
            sld = etree.SubElement(sldlst, q("p:sldId"))
            sld.set("id", str(255 + i))
            sld.set(q("r:id"), f"rIdSl{i}")
        parts["ppt/presentation.xml"] = etree.tostring(
            pres, xml_declaration=True, encoding="UTF-8", standalone=True)

        # presentation rels
        prels = copy.deepcopy(self.pkg.pres_rels)
        for rel in prels.findall(q("rel:Relationship")):
            if rel.get("Type").endswith("/slide"):
                prels.remove(rel)
        for i in range(1, len(self.slides) + 1):
            rel = etree.SubElement(prels, q("rel:Relationship"))
            rel.set("Id", f"rIdSl{i}")
            rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide")
            rel.set("Target", f"slides/slide{i}.xml")
        parts["ppt/_rels/presentation.xml.rels"] = etree.tostring(
            prels, xml_declaration=True, encoding="UTF-8", standalone=True)

        # content types
        ct = etree.fromstring(self.pkg.parts["[Content_Types].xml"])
        for ov in ct.findall(q("ct:Override")):
            pn = ov.get("PartName")
            if pn.startswith("/ppt/slides/") or pn.startswith("/ppt/notesSlides/"):
                ct.remove(ov)
        for i in range(1, len(self.slides) + 1):
            ov = etree.SubElement(ct, q("ct:Override"))
            ov.set("PartName", f"/ppt/slides/slide{i}.xml")
            ov.set("ContentType",
                   "application/vnd.openxmlformats-officedocument.presentationml.slide+xml")
            if self.slides[i - 1][2] is not None:
                ov = etree.SubElement(ct, q("ct:Override"))
                ov.set("PartName", f"/ppt/notesSlides/notesSlide{i}.xml")
                ov.set("ContentType",
                       "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml")
        parts["[Content_Types].xml"] = etree.tostring(
            ct, xml_declaration=True, encoding="UTF-8", standalone=True)

        out_path.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
            for name, data in sorted(parts.items()):
                z.writestr(name, data)


# ---------------------------------------------------------------- fills
def draw_structured_slide(root, blk, accent, tint, ctx):
    """Rule banner / hero card / support lines / routine chip / footer layout.

    Shared by grammar I-We-You do slides and the new-morphology You Do activity.
    """
    sid, y = 90100, 1.15
    if blk.get("rule"):
        rsz = fit_font_block(blk["rule"], 8.6, 0.46, 15, min_pt=11)
        add_textbox(root, sid, 0.5, y, 9.0, 0.52, [blk["rule"]], rsz,
                    color="FFFFFF", bold=True, fill=accent, rounded=True)
        sid += 1
        y += 0.66
    if blk.get("example"):
        ex_h = 1.65 if blk.get("items") else 2.45
        sz = fit_font_block(blk["example"], 7.7, ex_h - 0.25, 30, min_pt=18,
                            char_em=0.58)
        add_textbox(root, sid, 0.9, y, 8.2, ex_h, [blk["example"]], sz,
                    fill=tint, rounded=True)
        sid += 1
        y += ex_h + 0.12
    for it in blk.get("items", [])[:3]:
        add_textbox(root, sid, 0.9, y, 8.2, 0.40, [it], 14, align="l")
        sid += 1
        y += 0.44
    if y > 4.60:
        warn(f"{ctx}: content reaches y={y:.2f} - trim items")
    if blk.get("routine"):
        csz = fit_font(blk["routine"], 3.55, 13, min_pt=10, char_em=0.58)
        add_textbox(root, sid, 0.5, 4.66, 3.9, 0.44, [blk["routine"]], csz,
                    color="FFFFFF", bold=True, fill=accent, rounded=True)
        sid += 1
    if blk.get("footer"):
        add_textbox(root, sid, 4.55, 4.66, 4.95, 0.44, [blk["footer"]], 12,
                    align="r", italic=True, color="595959")


def sanitize(s):
    if s is None:
        return s
    for a, b in [("–", "-"), ("—", "-"), ("‘", "'"), ("’", "'"),
                 ("“", '"'), ("”", '"'), ("…", "...")]:
        s = s.replace(a, b)
    return s


def deep_sanitize(obj):
    if isinstance(obj, str):
        return sanitize(obj)
    if isinstance(obj, list):
        return [deep_sanitize(x) for x in obj]
    if isinstance(obj, dict):
        return {k: deep_sanitize(v) for k, v in obj.items()}
    return obj


PUNCT_RE = re.compile(r"([,.;:!?\"'-]+)")


def punct_runs(text):
    """Split plain text into runs with punctuation coloured red (CUPS marking)."""
    runs = []
    for i, seg in enumerate(PUNCT_RE.split(text)):
        if not seg:
            continue
        runs.append((seg, PUNCT_RED if i % 2 == 1 else None, None, None))
    return runs


def dictation_runs(sentence, targets):
    """Runs for the reveal: targets bold+underlined, punctuation red."""
    if not targets:
        return [punct_runs(sentence)]
    pattern = re.compile(
        r"\b(" + "|".join(re.escape(t) for t in sorted(targets, key=len, reverse=True)) + r")\b",
        re.IGNORECASE)
    runs = []
    pos = 0
    for m in pattern.finditer(sentence):
        if m.start() > pos:
            runs.extend(punct_runs(sentence[pos:m.start()]))
        runs.append((m.group(0), None, True, True))
        pos = m.end()
    if pos < len(sentence):
        runs.extend(punct_runs(sentence[pos:]))
    return [runs]


def lw_runs(word, unfair):
    """White word with the unfair part in yellow."""
    i = word.lower().find(unfair.lower()) if unfair else -1
    if unfair and i < 0:
        warn(f"learned word '{word}': unfair part '{unfair}' not found; highlighting nothing")
    if i < 0:
        return [[(word, "FFFFFF", None, None)]]
    runs = []
    if i > 0:
        runs.append((word[:i], "FFFFFF", None, None))
    runs.append((word[i:i + len(unfair)], "FFFF00", None, None))
    if i + len(unfair) < len(word):
        runs.append((word[i + len(unfair):], "FFFFFF", None, None))
    return [runs]


def build_session(pkg, week, session, out_path):
    d = DeckBuilder(pkg)
    day = session["day"]
    stype = session.get("type", "new")  # new | review | week_review
    term, wknum = week["term"], week["week"]
    cohort = week.get("cohort", "Enrichment")

    # --- 1 title
    def fill_title(root):
        sp = find_sp(root, SHAPE["title_sub"])
        set_runs(sp, [[(cohort, None, None, None)],
                      [(f"Term {term} Week {wknum} - {day}", None, None, None)]])
        # widen the subtitle box (kept centred) so the day fits on one line
        xfrm = sp.find(q("p:spPr") + "/" + q("a:xfrm"))
        off, ext = xfrm.find(q("a:off")), xfrm.find(q("a:ext"))
        new_w = int(4.6 * 914400)
        cx = int(off.get("x")) + int(ext.get("cx")) // 2
        off.set("x", str(cx - new_w // 2))
        ext.set("cx", str(new_w))
    d.add(T["title"], fill_title, notes=f"{cohort} OG session deck. Term {term} Week {wknum}, {day}.")

    # --- 2 weekly overview
    ov = week["overview"]  # dict day -> {morphology, grammar, learned_words}
    days = week.get("days", ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"])

    def fill_overview(root):
        rows = table_cells(root, SHAPE["overview_table"])
        # rows[0] headers; rows[1] morphology; rows[2] grammar; rows[3] learned words
        for ci, dy in enumerate(days, start=1):
            cells_m = rows[1].findall(q("a:tc"))
            cells_g = rows[2].findall(q("a:tc"))
            cells_l = rows[3].findall(q("a:tc"))
            if ci >= len(cells_m):
                break
            o = ov.get(dy, {})
            set_cell_paras(cells_m[ci], o.get("morphology", [""]) if isinstance(o.get("morphology"), list) else [o.get("morphology", "")])
            set_cell_paras(cells_g[ci], [o.get("grammar", "")])
            set_cell_paras(cells_l[ci], o.get("learned_words", [""]) if isinstance(o.get("learned_words"), list) else [o.get("learned_words", "")])
    d.add(T["overview"], fill_overview,
          notes="\n".join(f"{dy} - {ov.get(dy, {}).get('morphology','')} | Grammar: {ov.get(dy, {}).get('grammar','')} | LW: {ov.get(dy, {}).get('learned_words','')}" for dy in days))

    # --- 3 morphology review header (fixed drill notes stay)
    d.add(T["hdr_morph_review"])

    # --- review cards x10
    for card in session["morphology_review"]:
        tidx = CARD_REVIEW[card["type"]]

        def fill_card(root, card=card, tidx=tidx):
            sp = find_sp(root, SHAPE["card_review_word"][tidx])
            set_text(sp, card["morph"],
                     size_pt=fit_font(card["morph"], 5.75, 56))
        card_note = (f"**{card['type'].capitalize()}:** {card['morph']}\n"
                     f"**Keyword:** {card.get('keyword', '')}\n"
                     f"**Meaning:** {card['meaning']}")
        if card.get("derivative_ask"):
            card_note += f"\n**Derivative ask:** {card['derivative_ask']}"
        d.add(tidx, fill_card, notes=card_note)

    # --- words to read review
    d.add(T["hdr_wtr_review"])
    wtr = session["words_to_read_review"]
    if len(wtr["words"]) != 15:
        warn(f"{day}: words_to_read_review has {len(wtr['words'])} words (template holds 15)")

    def fill_wtr_table(root):
        rows = table_cells(root, SHAPE["wtr_review_tbl"])
        n = 0
        for row in rows:
            for tc in row.findall(q("a:tc")):
                if n < len(wtr["words"]):
                    set_cell_paras(tc, [f"{n + 1}) {wtr['words'][n]}"])
                else:
                    set_cell_paras(tc, [""])
                n += 1
    d.add(T["wtr_review_table"], fill_wtr_table, notes=wtr["notes"])

    # --- sound bank
    d.add(T["hdr_sound_bank"])
    sb = session["sound_bank"]
    if len(sb) > 9:
        warn(f"{day}: sound bank has {len(sb)} entries; only 9 boxes - extras dropped")

    def fill_sound_bank(root):
        for i, spid in enumerate(SHAPE["sound_boxes"]):
            sp = find_sp(root, spid)
            if i < len(sb):
                item = sb[i]
                set_text(sp, item["morph"],
                         size_pt=fit_font(item["morph"], 2.33, 45, min_pt=24))
                set_solid_fill(sp, MORPH_FILL[item["type"]])
            else:
                set_text(sp, "")
                set_solid_fill(sp, "FFFFFF")
    d.add(T["sound_bank"], fill_sound_bank,
          notes="Students copy sounds into books before spelling words. Box colour = card type: yellow root, green prefix, red suffix.")

    # --- words to spell review (one reveal slide per word)
    d.add(T["hdr_wts_review"])
    wsr = session["words_to_spell_review"]
    if not 8 <= len(wsr) <= 12:
        warn(f"{day}: words_to_spell_review has {len(wsr)} words (expected 10)")
    for item in wsr:
        def fill_spell(root, item=item):
            sp = find_sp(root, SHAPE["spell_word_box"])
            set_text(sp, item["word"], size_pt=fit_font(item["word"], 4.96, 60, min_pt=28))
        if not item.get("prompt"):
            warn(f"{day}: spell review word '{item['word']}' has no after-check prompt")
        elif not item.get("answer"):
            warn(f"{day}: spell review word '{item['word']}' prompt has no answer")
        note = (f"**Word to spell:** {item['word']}\n"
                f"**Sentence:** {item['sentence']}\n"
                "Click to reveal answer.")
        if item.get("prompt"):
            note += f"\n**After checking:** {item['prompt']}"
        if item.get("answer"):
            note += f"\n**Answer:** {item['answer']}"
        d.add(T["spell_word"], fill_spell, notes=note)

    # --- new/review morphology section (skipped on week_review days)
    if stype != "week_review":
        section = "New Morphology" if stype == "new" else "Review Morphology"
        nm = session["new_morphology"]

        def fill_nm_hdr(root):
            sp = find_sp(root, SHAPE["hdr_new_morph_txt"])
            replace_run_text(sp, "New Morphology", section)
        d.add(T["hdr_new_morph"], fill_nm_hdr)

        tidx = CARD_NEW[nm["type"]]

        def fill_nm_card(root, tidx=tidx):
            sp = find_sp(root, SHAPE["card_new_word"][tidx])
            set_text(sp, nm["morph"], size_pt=fit_font(nm["morph"], 5.75, 56))
        d.add(tidx, fill_nm_card,
              notes=(f"**{nm['type'].capitalize()}:** {nm['morph']}\n"
                     f"**Keyword:** {nm.get('keyword', '')}\n"
                     f"**Meaning:** {nm['meaning']}"))

        def fill_wtr_new_hdr(root):
            sp = find_sp(root, SHAPE["hdr_wtr_new_txt"])
            replace_run_text(sp, "New Morphology", section)
        d.add(T["hdr_wtr_new"], fill_wtr_new_hdr)

        words = session["words_to_read_new"]  # list of {word, meaning}
        if not 6 <= len(words) <= 12:
            warn(f"{day}: words_to_read_new has {len(words)} words (aim 9-12, min 6)")

        def fill_wtr_grid(root):
            # always the template's 3-column grid; with 10-12 words the columns
            # widen and stretch down the page so a 4th row still reads big
            n = len(words)
            col_spids = SHAPE["wtr_cols"]
            per = math.ceil(n / 3)
            cols = [words[0:per], words[per:2 * per], words[2 * per:]]
            if n > 9:
                for spid, x in zip(col_spids, (0.30, 3.45, 6.60)):
                    sp = find_sp(root, spid)
                    xfrm = sp.find(q("p:spPr") + "/" + q("a:xfrm"))
                    xfrm.find(q("a:off")).set("x", str(int(x * 914400)))
                    xfrm.find(q("a:off")).set("y", str(int(1.15 * 914400)))
                    xfrm.find(q("a:ext")).set("cx", str(int(3.10 * 914400)))
                    xfrm.find(q("a:ext")).set("cy", str(int(3.95 * 914400)))
                    bodypr = sp.find(q("p:txBody") + "/" + q("a:bodyPr"))
                    if bodypr is not None:  # centre the 4-row block vertically
                        bodypr.set("anchor", "ctr")
                col_w, box_h = 3.10, 3.95
            else:
                col_w, box_h = 3.0, 3.0
            longest = max(len(w["word"]) for w in words)
            rows_max = max(len(c) for c in cols)
            sz = fit_font("x" * longest, col_w, 50 if rows_max <= 3 else 44, min_pt=22)
            sz = min(sz, int(box_h / rows_max * 72 / 1.5))
            targets = []
            for spid, col in zip(col_spids, cols):
                sp = find_sp(root, spid)
                set_runs(sp, [[(w["word"], None, None, None)] for w in col] or [[]],
                         size_pt=sz)
                for pi in range(len(col)):
                    targets.append((spid, pi))
            regen_timing(root, targets)
        d.add(T["wtr_new_grid"], fill_wtr_grid,
              notes=session.get(
                  "wtr_new_notes",
                  "\n".join(f"**{w['word']}** - {w['meaning']}" for w in words)))

        def fill_wts_new_hdr(root):
            sp = find_sp(root, SHAPE["hdr_wts_new_txt"])
            replace_run_text(sp, "New Morphology", section)
        d.add(T["hdr_wts_new"], fill_wts_new_hdr)

        spell = session["words_to_spell_new"]
        if len(spell) != 4:
            warn(f"{day}: words_to_spell_new has {len(spell)} words (template holds 4)")

        def fill_wts_grid(root):
            pairs = [spell[0:2], spell[2:4]]
            longest = max((len(w) for w in spell), default=1)
            sz = fit_font("x" * longest, 4.8, 52, min_pt=28)
            targets = []
            for spid, pair in zip(SHAPE["wts_boxes"], pairs):
                sp = find_sp(root, spid)
                set_runs(sp, [[(w, None, None, None)] for w in pair] or [[]], size_pt=sz)
                for pi in range(len(pair)):
                    targets.append((spid, pi))
            regen_timing(root, targets)
        d.add(T["wts_new_grid"], fill_wts_grid,
              notes=session.get("extension",
                                "Once finished, students write a fun sentence with one or more of the words."))

        # --- Yoshimoto You Do activity (magenta header + designed task slide)
        act = session.get("new_morph_activity")
        if not act:
            warn(f"{day}: no new_morph_activity (Yoshimoto You Do) provided")
        else:
            def fill_act_hdr(root, section=section):
                sp = find_sp(root, SHAPE["hdr_wts_new_txt"])
                replace_run_text(sp, "Words to Spell- New Morphology",
                                 f"{section} - You Do")
                replace_run_text(sp, "2 mins", act.get("time", "3 mins"))
            d.add(T["hdr_wts_new"], fill_act_hdr)

            act_tidx = T["gradual_release"]["you_do"]

            def fill_act(root, act=act, act_tidx=act_tidx):
                tsp = find_sp(root, SHAPE["gr_title"][act_tidx])
                set_text(tsp, "You do - " + act.get("title", ""))
                bsp = find_sp(root, SHAPE["gr_body"][act_tidx])
                set_runs(bsp, [[]])
                draw_structured_slide(root, act, MORPH_MAGENTA, MORPH_TINT,
                                      f"{day} new_morph_activity")
            d.add(act_tidx, fill_act, notes=act.get("notes", ""))

    # --- learned words
    lw = session["learned_words"]
    d.add(T["hdr_review_lw"])
    for item in lw.get("review", []):
        def fill_lw(root, item=item):
            sp = find_sp(root, SHAPE["lw_word"][22])
            sz = fit_font(item["word"], 7.78, 67, min_pt=32)
            set_runs(sp, lw_runs(item["word"], item.get("unfair", "")), size_pt=sz)
        d.add(T["learned_word"], fill_lw, notes=item["notes"])
    if lw.get("new"):
        item = lw["new"]
        d.add(T["hdr_new_lw"])

        def fill_lw_new(root, item=item):
            sp = find_sp(root, SHAPE["lw_word"][25])
            sz = fit_font(item["word"], 7.78, 67, min_pt=32)
            set_runs(sp, lw_runs(item["word"], item.get("unfair", "")), size_pt=sz)
        d.add(T["learned_word_new"], fill_lw_new, notes=item["notes"])

    # --- dictation
    d.add(T["hdr_dictation"])
    for dt in session["dictation"]:
        cups = dt.get("cups")
        targets = dt.get("targets", [])
        score = dt.get("score")
        if score is None and cups:
            score = (len(cups.get("capitals", [])) + len(cups.get("punctuation", []))
                     + len(targets) + 1)  # +1 = Understanding (dictated order)
        if not cups:
            warn(f"{day}: dictation '{dt['sentence'][:30]}...' has no cups marking block")

        def fill_dict(root, dt=dt, score=score):
            sp = find_sp(root, SHAPE["dictation_sentence"])
            sz = fit_font_block(dt["sentence"], 9.0, 2.6, 48, min_pt=28)
            set_runs(sp, dictation_runs(dt["sentence"], dt.get("targets", [])), size_pt=sz)
            if score:
                add_textbox(root, 90001, 6.8, 4.55, 2.7, 0.45,
                            [f"Score: ___ /{score}"], 20, align="r", anchor="ctr")
        lines = [dt["sentence"], ""]
        if cups:
            caps = cups.get("capitals", [])
            punct = cups.get("punctuation", [])
            lines += ["Mark with CUPS, tick each line:",
                      f"**C - Capitals:** {', '.join(caps)} ({len(caps)})",
                      "**U - Understanding:** words written in the dictated order (1)",
                      f"**P - Punctuation, shown red:** {', '.join(punct)} ({len(punct)})",
                      f"**S - Spelling targets, underlined:** {', '.join(targets)} ({len(targets)})",
                      f"**Score:** /{score}"]
        elif targets:
            lines.append(f"**Targets:** {', '.join(targets)}")
        if dt.get("focus"):
            lines.append(f"**Focus:** {dt['focus']}")
        d.add(T["dictation"], fill_dict, notes="\n".join(lines))

    # --- grammar
    gr = session["grammar"]

    def fill_gr_hdr(root):
        pass
    d.add(T["hdr_grammar"], fill_gr_hdr,
          notes=gr.get("header_notes", ""))
    for key, label in [("i_do", "I do - "), ("we_do", "We do - "), ("you_do", "You do - ")]:
        tidx = T["gradual_release"][key]
        blk = gr[key]

        def fill_gr(root, tidx=tidx, blk=blk, label=label):
            tsp = find_sp(root, SHAPE["gr_title"][tidx])
            set_text(tsp, label + blk.get("title", ""))
            bsp = find_sp(root, SHAPE["gr_body"][tidx])
            if "lines" in blk:  # legacy plain layout
                set_runs(bsp, [[(ln, None, None, None)] for ln in blk["lines"]] or [[]])
                return
            set_runs(bsp, [[]])  # clear the body placeholder; we draw instead
            draw_structured_slide(root, blk, GRAMMAR_PURPLE, GRAMMAR_TINT,
                                  f"{day} grammar {key}")
        d.add(tidx, fill_gr, notes=blk.get("notes", ""))

    d.write(out_path)
    return len(d.slides)


# ---------------------------------------------------------------- QA gate
def qa_deck(path):
    """Reopen the deck: parse check + leftover-placeholder check."""
    problems = []
    try:
        from pptx import Presentation
        pres = Presentation(path)
        for i, slide in enumerate(pres.slides, 1):
            for sh in slide.shapes:
                if sh.has_text_frame and "XYZ" in sh.text_frame.text:
                    problems.append(f"slide {i}: leftover XYZ placeholder")
            if slide.has_notes_slide and "XYZ" in slide.notes_slide.notes_text_frame.text:
                problems.append(f"slide {i}: leftover XYZ in notes")
    except Exception as e:
        problems.append(f"failed to reopen: {e}")
    return problems


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    only = None
    outdir = Path("output")
    for i, a in enumerate(sys.argv[1:]):
        if a == "--only":
            only = sys.argv[1:][i + 1]
        if a == "--outdir":
            outdir = Path(sys.argv[1:][i + 1])
    if not args:
        sys.exit("usage: build_og_week.py <week_spec.json> [--only Day]")
    spec_path = Path(args[0])
    week = deep_sanitize(json.loads(spec_path.read_text()))
    pkg = Package(MASTER)
    unit = week.get("unit_folder", f"OG_Term{week['term']}_Week{week['week']}")
    built = []
    failed = False
    morph_idx = 0
    for session in week["sessions"]:
        # team filename convention: "1a. Monday (-eer).pptx", "1b. Tuesday
        # (-eer review).pptx", ..., "3. Friday (week review).pptx"
        stype = session.get("type", "new")
        if stype == "new":
            morph_idx += 1
            label = f"{morph_idx}a"
        elif stype == "review":
            label = f"{max(morph_idx, 1)}b"
        else:
            morph_idx += 1
            label = str(morph_idx)
        morph = session.get("new_morphology", {}).get("morph", "")
        morph = morph.replace(" / ", "-").replace("/", "-")
        if stype == "week_review" or not morph:
            paren = "week review"
        else:
            paren = morph + (" review" if stype == "review" else "")
        fname = session.get("file_name", f"{label}. {session['day']} ({paren}).pptx")
        if only and session["day"].lower() != only.lower():
            continue
        out_path = outdir / unit / fname
        print(f"Building {out_path} ...")
        n = build_session(pkg, week, session, out_path)
        problems = qa_deck(out_path)
        for p in problems:
            print(f"  GATE FAIL: {p}")
        if problems:
            failed = True
        print(f"  {n} slides written.")
        built.append(out_path)
    if WARNINGS:
        print(f"\n{len(WARNINGS)} warning(s) - review above.")
    if failed:
        sys.exit("QA gate failed.")
    print("\nDone:")
    for b in built:
        print(f"  {b}")


if __name__ == "__main__":
    main()
