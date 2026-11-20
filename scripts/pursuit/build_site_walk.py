#!/usr/bin/env python3
"""scripts/pursuit/build_site_walk.py — 27001D one-page site walk checklist.

Writes, into a staging folder for scripts/publish/publish_outputs.py:
  * a one-page landscape PDF drawn at exact coordinates (PyMuPDF), the
    field copy; the script asserts one page, no label overflow, and the
    minimum write-in / sketch / notes space;
  * a matching .docx (python-docx) for transcription into internal templates;
  * a matching .md.

    python scripts/pursuit/build_site_walk.py [--stage _scratch/stage_sitewalk]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
REPO_ROOT = SCRIPT_DIR.parent.parent
sys.path.insert(0, str(REPO_ROOT / "scripts" / "lib"))
from guard_cwd import require_proposal_or_engine_root  # noqa: E402

import fitz  # noqa: E402  (PyMuPDF)
from docx import Document  # noqa: E402
from docx.enum.section import WD_ORIENT, WD_SECTION  # noqa: E402
from docx.enum.table import WD_ROW_HEIGHT_RULE  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Inches, Pt, RGBColor  # noqa: E402

P = "27001D"
VERSION = "v02"
FONT = "Arial Narrow"
FONT_FILE = "C:/Windows/Fonts/ARIALN.TTF"
FONT_FILE_B = "C:/Windows/Fonts/ARIALNB.TTF"
SZ = 7.0
BOX = "☐"

TITLE = "SITE WALK CHECKLIST — 27001D · N4008027R0004 FY27 CIP MINCON: Substation Generators · NSF Dahlgren, VA"
META = ("Tue 17 Nov 2026, 1300 EST · Gov site POC: Brian Hornaman · Encon: T. Colvin / M. Jackson / D. Ajmera · "
        "BFA: ______________________     Bring: ID + base access, PPE, laser + 100 ft tape, phone GPS, camera (confirm photo approval), "
        "RFP field packet (E-101–E-108 11×17, SOW pp 24, 29, 44, 195–199), PPI list, this sheet.     "
        "By: K = Encon (KTR) data · A = BFA A/E data · Q = ask Gov, record answer. Verbal answers are not binding; submit as PPIs by 2 Dec.")

# Sections: (title, mode, rows). Row = (label, by). mode "site" → Sub #1 / Sub #6 write-in columns.
COL1 = [
    ("1  EXISTING 4.16 kV SWITCHGEAR  (E-103–E-106)", "site", [
        ("Indoor/outdoor, bldg no., door width; front/rear working clearance", "A"),
        ("Mfr / model / year; V · bus A · kA int · BIL (photo nameplates)", "A"),
        ("Spare / 'Feeder Future' cell no., breaker present?; room for new section?", "A"),
        ("Existing relays (7SJ80?), CT ratios/class; 87B present?", "A"),
        ("Bus / sync PTs available?  Main-tie interlocks (kirk key)?", "A"),
        ("DC control power: V, battery / charger age", "A"),
        ("480 V black-start source + continuous aux power for trailer standby", "A"),
        ("System grounding (solid / resistance / ungrounded); ground grid tie pt", "A"),
        ("Cable entry top/bottom; floor trench; spare conduit stubs", "A"),
        ("Arc-flash labels; date of last study / relay test", "A"),
        ("SCADA / comm panel; fiber or copper; nearest network node + dist", "A"),
        ("Who operates switchgear; outage window + lead time", "K"),
    ]),
    ("2  TRAILER PARK POSITION + CONDUIT ROUTE  (E-101/102, E-107/108)", "site", [
        ("Usable L × W (need ≈52 × 8.5 ft + 4 ft clear); surface, thickness", "A"),
        ("Level / slope; ponding, drainage; signs of fill or soft ground", "A"),
        ("Route LF: switchgear → JB → trailer; JB location", "A"),
        ("Crossings: roads, walks; utility marks, manholes, duct banks", "A"),
        ("Overhead lines / limbs, clearance > 13.5 ft?; area lighting", "A"),
        ("Delivery path, turning room, gate width (WB-62); fence/bollard mods", "K"),
        ("Shoreline / wetland edge, high-water marks (Sub #1); flood zone", "A"),
        ("Survey monuments seen; GPS elevation (RFP says 112 ft AMSL)", "A"),
    ]),
]
COL2 = [
    ("3  FIRE · ENVIRONMENTAL · MECHANICAL", "site", [
        ("Distance to bldgs, doors, windows, air intakes (NFPA 37)", "A"),
        ("Exhaust direction vs intakes (prevailing wind S); noise receptors", "A"),
        ("Fire alarm panel / base reporting for CO2 system; nearest hydrant", "A"),
        ("Storm inlets near fuel / DEF area; existing ASTs, containment", "A"),
        ("DEF tank spot; water source for purge + distance", "A"),
        ("Fuel truck access for refill; fill-point height", "K"),
    ]),
    ("4  CONTRACTOR LOGISTICS", "site", [
        ("Laydown / staging / parking; temp power source", "K"),
        ("Escort, work hours, dig-permit and locate process", "K"),
        ("Hazmat seen (lead paint, asbestos, PCB tags)", "K"),
    ]),
    ("5  ASK THE GOVERNMENT  (record answer; feeds PPIs due 2 Dec)", "answer", [
        ("PPI-01  Subs #1A and #4 in scope, or #1 and #6 only?", "Q"),
        ("PPI-15  Spare cubicle for gen breaker + 87B at each sub?", "Q"),
        ("PPI-16/19  Island only or grid-parallel? System grounding?", "Q"),
        ("PPI-14  Conduit above grade (metal) or underground?", "Q"),
        ("PPI-13  New concrete pad required? Sub #6 pad capacity?", "Q"),
        ("PPI-07  SCR/DEF really required? Air permit holder?", "Q"),
        ("PPI-17  SCADA system/owner; RMF/ATO by whom; ISSM?", "Q"),
        ("PPI-20  Black-start 480 V source at each sub?", "Q"),
        ("Record dwgs, relay settings, last arc-flash study available?", "Q"),
    ]),
]
PHOTOS = ["switchgear front + every nameplate", "spare cell open", "rear / cable entry", "relays + PTs",
          "park area ×4 directions", "full conduit route", "gate / fence", "overhead", "adjacent bldgs / intakes",
          "hydrant / inlets / FACP", "wetland edge (Sub #1)"]
SKETCH = ("SKETCH Sub #1 — swgr → JB → trailer, dims, N arrow",
          "SKETCH Sub #6 — swgr → JB → trailer, dims, N arrow")
NOTES = "NOTES — surprises, scope gaps, items for Encon pricing, follow-ups (owner / date)"

# ── geometry (inches) — shared by PDF and docx ─────────────────────────
PAGE_W, PAGE_H = 11.0, 8.5
ML, MT, MB = 0.35, 0.3, 0.3
COLW, GAP = 5.0, 0.3
SITE_W = [0.14, 2.62, 0.16, 1.04, 1.04]
ANS_W = [0.14, 2.62, 0.16, 2.08]
HDR_H = 0.17          # section header bar
ROW_H = 0.23          # write-in row (Sub #1 / Sub #6)
ANS_H = 0.26          # Government answer row
SEC_GAP = 0.07
MIN_SKETCH, MIN_NOTES = 1.6, 1.6   # asserted minimum box heights

NAVY, RED, TINT = (0x1F / 255, 0x38 / 255, 0x64 / 255), (0.75, 0, 0), (0xD9 / 255, 0xE1 / 255, 0xF2 / 255)
GRID = (0.55, 0.55, 0.55)
FAINT = (0.82, 0.82, 0.82)


# ── PDF (field copy) ───────────────────────────────────────────────────
class Pdf:
    def __init__(self):
        self.doc = fitz.open()
        self.pg = self.doc.new_page(width=PAGE_W * 72, height=PAGE_H * 72)
        self.pg.insert_font(fontname="an", fontfile=FONT_FILE)
        self.pg.insert_font(fontname="anb", fontfile=FONT_FILE_B)
        self.f = {"an": fitz.Font(fontfile=FONT_FILE), "anb": fitz.Font(fontfile=FONT_FILE_B)}
        self.overflow: list[str] = []

    def width(self, text, size=SZ, bold=False):
        return self.f["anb" if bold else "an"].text_length(text, fontsize=size) / 72

    def text(self, x, y, text, size=SZ, bold=False, color=(0, 0, 0), maxw=None):
        if maxw is not None and self.width(text, size, bold) > maxw:
            self.overflow.append(text)
        self.pg.insert_text((x * 72, y * 72), text, fontname="anb" if bold else "an", fontsize=size, color=color)

    def rect(self, x, y, w, h, fill=None, color=GRID, width=0.5):
        self.pg.draw_rect(fitz.Rect(x * 72, y * 72, (x + w) * 72, (y + h) * 72), color=color, fill=fill, width=width)

    def checkbox(self, x, y, s=0.075):
        self.rect(x, y, s, s, color=(0, 0, 0), width=0.6)

    def wrap(self, x, y, text, w, size, bold=False, lead=1.2):
        words, line, lines = text.split(" "), "", []
        for wd in words:
            t = (line + " " + wd).strip()
            if self.width(t, size, bold) > w and line:
                lines.append(line); line = wd
            else:
                line = t
        lines.append(line)
        for i, ln in enumerate(lines):
            self.text(x, y + i * size * lead / 72, ln, size, bold)
        return y + (len(lines) - 1) * size * lead / 72

    def header_bar(self, x, y, widths, cells):
        """cells: list of (text, span) across widths."""
        cx, i = x, 0
        for txt, span in cells:
            w = sum(widths[i:i + span])
            self.rect(cx, y, w, HDR_H, fill=NAVY, color=NAVY)
            self.text(cx + 0.04, y + HDR_H - 0.05, txt, SZ, True, (1, 1, 1), maxw=w - 0.06)
            cx += w; i += span

    def section(self, x, y, title, mode, rows):
        widths = SITE_W if mode == "site" else ANS_W
        h = ROW_H if mode == "site" else ANS_H
        if mode == "site":
            self.header_bar(x, y, widths, [(title, 3), ("Sub #1", 1), ("Sub #6", 1)])
        else:
            self.header_bar(x, y, widths, [(title, 3), ("Answer / who answered", 1)])
        y += HDR_H
        for label, by in rows:
            cx = x
            for w in widths:
                self.rect(cx, y, w, h); cx += w
            base = y + h / 2 + SZ * 0.35 / 72
            self.checkbox(x + 0.035, y + h / 2 - 0.0375)
            self.text(x + widths[0] + 0.03, base, label, SZ, maxw=widths[1] - 0.05)
            self.text(x + widths[0] + widths[1] + 0.035, base, by, SZ, True, RED if by == "Q" else NAVY)
            y += h
        return y + SEC_GAP

    def box(self, x, y, titles, h, ruled=False):
        n = len(titles)
        w = COLW / n
        for i, t in enumerate(titles):
            bx = x + i * w
            self.rect(bx, y, w, HDR_H, fill=TINT)
            self.text(bx + 0.04, y + HDR_H - 0.05, t, SZ, True, maxw=w - 0.06)
            self.rect(bx, y + HDR_H, w, h)
            if ruled:      # faint writing lines every 0.25 in
                ly = y + HDR_H + 0.25
                while ly < y + HDR_H + h - 0.05:
                    self.pg.draw_line(((bx + 0.05) * 72, ly * 72), ((bx + w - 0.05) * 72, ly * 72), color=FAINT, width=0.4)
                    ly += 0.25
            else:          # faint 0.25 in dot grid for dimensioned sketches
                gy = y + HDR_H + 0.125
                while gy < y + HDR_H + h - 0.05:
                    gx = bx + 0.125
                    while gx < bx + w - 0.05:
                        self.pg.draw_circle((gx * 72, gy * 72), 0.35, color=FAINT, fill=FAINT)
                        gx += 0.25
                    gy += 0.25
        return y + HDR_H + h

    def photos(self, x, y, w):
        self.text(x, y + 0.1, "PHOTOS", SZ, True, NAVY)
        cx, cy = x + self.width("PHOTOS ", SZ, True) + 0.03, y
        for item in PHOTOS:
            iw = 0.11 + self.width(item) + 0.07
            if cx + iw > x + w:
                cx, cy = x, cy + 0.14
            self.checkbox(cx, cy + 0.035)
            self.text(cx + 0.1, cy + 0.1, item)
            cx += iw
        return cy + 0.14 + SEC_GAP


def build_pdf(out):
    pdf = Pdf()
    x1, x2 = ML, ML + COLW + GAP
    pdf.text(ML, MT + 0.15, TITLE, 10.5, True, NAVY)
    top = pdf.wrap(ML, MT + 0.32, META, PAGE_W - 2 * ML, 6.8) + 0.12
    bottom = PAGE_H - MB
    # column 1
    y = top
    for t, m, r in COL1:
        y = pdf.section(x1, y, t, m, r)
    y = pdf.photos(x1, y, COLW)
    sketch_h = bottom - y - HDR_H
    pdf.box(x1, y, SKETCH, sketch_h)
    # column 2
    y = top
    for t, m, r in COL2:
        y = pdf.section(x2, y, t, m, r)
    notes_h = bottom - y - HDR_H
    pdf.box(x2, y, (NOTES,), notes_h, ruled=True)
    pdf.text(PAGE_W - ML - 1.6, PAGE_H - 0.12, f"27001D_09 Site Walk Checklist {VERSION} · BFA / Encon internal", 5.5,
             color=(0.4, 0.4, 0.4))
    pdf.doc.set_metadata({"title": f"27001D Site Walk Checklist {VERSION}", "author": "Bowman, Foster & Associates"})
    pages = pdf.doc.page_count
    pdf.doc.save(out)
    report = {"pages": pages, "body_top_in": round(top, 2), "row_h_in": ROW_H, "answer_row_h_in": ANS_H,
              "write_cell_w_in": SITE_W[3], "answer_cell_w_in": ANS_W[3],
              "sketch_box_in": f"{COLW / 2:.2f} x {sketch_h:.2f} (each of 2)", "notes_box_in": f"{COLW:.2f} x {notes_h:.2f}",
              "label_overflow": pdf.overflow}
    assert pages == 1, report
    assert not pdf.overflow, report
    assert sketch_h >= MIN_SKETCH and notes_h >= MIN_NOTES, report
    return report, sketch_h, notes_h


# ── docx (transcription copy) ──────────────────────────────────────────
def _font(run, size=SZ, bold=False, color=None):
    run.font.name = FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def _para(p, text="", size=SZ, bold=False, color=None):
    pf = p.paragraph_format
    pf.space_before = pf.space_after = Pt(0)
    pf.line_spacing = 1.0
    if text:
        _font(p.add_run(text), size, bold, color)
    return p


def _shade(cell, hexfill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto"); shd.set(qn("w:fill"), hexfill)
    tcPr.append(shd)


def _tight(table):
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side, w in (("top", 0), ("left", 35), ("bottom", 0), ("right", 35)):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:w"), str(w)); e.set(qn("w:type"), "dxa")
        mar.append(e)
    tblPr.append(mar)
    lay = OxmlElement("w:tblLayout"); lay.set(qn("w:type"), "fixed"); tblPr.append(lay)


def _widths(table, widths):
    for row in table.rows:
        for c, w in zip(row.cells, widths):
            c.width = Inches(w)


def _height(row, inches, exact=False):
    row.height = Inches(inches)
    row.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY if exact else WD_ROW_HEIGHT_RULE.AT_LEAST


def section_table(doc, title, mode, rows):
    widths = SITE_W if mode == "site" else ANS_W
    t = doc.add_table(rows=1, cols=len(widths))
    t.style = "Table Grid"
    _tight(t)
    hdr = t.rows[0].cells
    h = hdr[0].merge(hdr[2])
    _para(h.paragraphs[0], title, SZ, True, "FFFFFF"); _shade(h, "1F3864")
    if mode == "site":
        for c, txt in ((hdr[3], "Sub #1"), (hdr[4], "Sub #6")):
            _para(c.paragraphs[0], txt, SZ, True, "FFFFFF"); _shade(c, "1F3864")
    else:
        _para(hdr[3].paragraphs[0], "Answer / who answered", SZ, True, "FFFFFF"); _shade(hdr[3], "1F3864")
    for label, by in rows:
        r = t.add_row()
        _height(r, ROW_H if mode == "site" else ANS_H, exact=True)
        _para(r.cells[0].paragraphs[0], BOX, SZ)
        _para(r.cells[1].paragraphs[0], label, SZ)
        _para(r.cells[2].paragraphs[0], by, SZ, True, "C00000" if by == "Q" else "1F3864")
    _widths(t, widths)
    _para(doc.add_paragraph(), "", 2)


def notes_box(doc, titles, height):
    cols = len(titles)
    t = doc.add_table(rows=2, cols=cols)
    t.style = "Table Grid"
    _tight(t)
    for c, txt in zip(t.rows[0].cells, titles):
        _para(c.paragraphs[0], txt, SZ, True); _shade(c, "D9E1F2")
    _height(t.rows[1], height, exact=True)
    _widths(t, [COLW / cols] * cols)


def build_docx(out, sketch_h, notes_h):
    # Word's line metrics run slightly taller than the PDF's; trim boxes so the docx stays on one page.
    trim = 0.25
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = FONT; st.font.size = Pt(SZ)
    s = doc.sections[0]
    s.orientation = WD_ORIENT.LANDSCAPE
    s.page_width, s.page_height = Inches(PAGE_W), Inches(PAGE_H)
    s.left_margin = s.right_margin = Inches(ML)
    s.top_margin, s.bottom_margin = Inches(MT), Inches(MB)
    s.header_distance = s.footer_distance = Inches(0.2)
    _para(doc.paragraphs[0] if doc.paragraphs else doc.add_paragraph(), TITLE, 10.5, True, "1F3864")
    _para(doc.add_paragraph(), META, 6.8)
    sec = doc.add_section(WD_SECTION.CONTINUOUS)
    cols = sec._sectPr.xpath("./w:cols")
    c = cols[0] if cols else OxmlElement("w:cols")
    c.set(qn("w:num"), "2"); c.set(qn("w:space"), str(int(GAP * 1440)))
    if not cols:
        sec._sectPr.append(c)
    for title, mode, rows in COL1:
        section_table(doc, title, mode, rows)
    _para(doc.add_paragraph(), "PHOTOS  " + "  ".join(f"{BOX} {p}" for p in PHOTOS), SZ)
    _para(doc.add_paragraph(), "", 2)
    notes_box(doc, SKETCH, sketch_h - trim)
    br = doc.add_paragraph()
    _para(br)
    b = OxmlElement("w:br"); b.set(qn("w:type"), "column"); br.add_run()._r.append(b)
    for title, mode, rows in COL2:
        section_table(doc, title, mode, rows)
    notes_box(doc, (NOTES,), notes_h - trim)
    doc.save(out)


def build_md(out):
    L = [f"# Site Walk Checklist — {P} ({VERSION})", "",
         "**N4008027R0004 FY27 CIP MINCON: Substation Generators, NSF Dahlgren, VA**  ",
         META.replace("     ", "  \n"), "",
         "<!-- INTERNAL: source for the one-page field copy. The PDF is the print copy; the .docx is for transcription. -->"]
    for title, mode, rows in COL1 + COL2:
        L += ["", f"## {title}", ""]
        if mode == "site":
            L += ["| ☐ | Data point | By | Sub #1 | Sub #6 |", "|---|---|---|---|---|"]
            L += [f"| ☐ | {lbl} | {by} |  |  |" for lbl, by in rows]
        else:
            L += ["| ☐ | Question | By | Answer / who answered |", "|---|---|---|---|"]
            L += [f"| ☐ | {lbl} | {by} |  |" for lbl, by in rows]
    L += ["", "## Route sketches", "", "Sub #1 and Sub #6: switchgear → JB → trailer, with dimensions and a north arrow.",
          "", "## Photos", "", " ".join(f"☐ {p}" for p in PHOTOS), "", "## Notes", "", NOTES.split(" — ", 1)[1], ""]
    out.write_text("\n".join(L), encoding="utf-8")


def main():
    require_proposal_or_engine_root(REPO_ROOT)
    ap = argparse.ArgumentParser()
    ap.add_argument("--stage", default="_scratch/stage_sitewalk")
    a = ap.parse_args()
    st = REPO_ROOT / a.stage
    (st / "artifacts").mkdir(parents=True, exist_ok=True)
    (st / "published").mkdir(parents=True, exist_ok=True)
    report, sketch_h, notes_h = build_pdf(st / "published" / f"{P}_09_Site_Walk_Checklist.pdf")
    build_docx(st / "published" / f"{P}_09_Site_Walk_Checklist.docx", sketch_h, notes_h)
    build_md(st / "artifacts" / f"{P}_09_Site_Walk_Checklist.md")
    for k, v in report.items():
        print(f"{k:18} {v}")
    print("built", st)


if __name__ == "__main__":
    main()
