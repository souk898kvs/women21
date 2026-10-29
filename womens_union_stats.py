#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ບົດສະຫຼຸບສະຖິຕິ ສະຫະພັນແມ່ຍິງ ໜ່ວຍຄູ  (ໂປຣແກຣມຈັດການລາຍຊື່ + ພິມ Legal)

ຕິດຕັ້ງ:   pip install python-docx        (ຈຳເປັນສະເພາະການສົ່ງອອກ Word)
ຣັນ:       python womens_union_stats.py
ຂໍ້ມູນເກັບໃນໄຟລ໌ womens_union_data.json ຢູ່ຂ້າງໆໂປຣແກຣມ.

ການພິມ:  ປຸ່ມ "ພິມ (HTML)" ຈະເປີດໃນບຼາວເຊີ -> ກົດ Ctrl+P (Legal, ປິດ Headers and footers)
         ຫຼື ປຸ່ມ "ສົ່ງອອກ Word" ແລ້ວເປີດໃນ Word ເພື່ອພິມ.
ຟອນ:     Phetsarath OT ຫຼື Times New Roman (ຕົວເລກ/ລາຕິນ) + Phetsarath OT (ລາວ)
         -- ຟອນຕ້ອງຕິດຕັ້ງໃນເຄື່ອງ.
"""
import csv
import html
import json
import os
import re
import tempfile
import webbrowser

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, "womens_union_data.json")

# (key, ຫົວຂໍ້, ຄວາມກວ້າງ %, ຊິດຊ້າຍ?)
COLS = [
    ("rank",   "ຊັ້ນ",                       4,    False),
    ("name",   "ຊື່ ແລະ ນາມສະກຸນ",          11,   True),
    ("ethnic", "ຊົນເຜົ່າ",                  4,    False),
    ("dob",    "ວັນເດືອນ ປີເກີດ",           5.5,  False),
    ("vill",   "ບ້ານຢູ່ ປະຈຸບັນ",           6,    False),
    ("dist",   "ເມືອງ",                     5,    False),
    ("prov",   "ແຂວງ",                      5.5,  False),
    ("bvill",  "ບ້ານເກີດ",                  6,    False),
    ("bdist",  "ເມືອງ",                     5,    False),
    ("bprov",  "ແຂວງ",                      5.5,  False),
    ("army",   "ວັນເດືອນປີ ເຂົ້າທະຫານ",      5.2,  False),
    ("resv",   "ວັນເດືອນປີ ເຂົ້າພັກສຳຮອງ",   5.2,  False),
    ("party",  "ວັນເດືອນປີ ເຂົ້າພັກສົມບູນ",  5.2,  False),
    ("women",  "ວັນເດືອນປີ ເຂົ້າແມ່ຍິງ",     5.2,  False),
    ("youth",  "ວັນເດືອນປີ ເຂົ້າຊາວໜຸ່ມ",    5.2,  False),
    ("edu",    "ລະດັບ ການສຶກສາ",            3.5,  False),
    ("spec",   "ລະດັບ ວິຊາສະເພາະ",          4.5,  False),
    ("phone",  "ເບີໂທຕິດຕໍ່",               6,    False),
]
KEYS = [c[0] for c in COLS]
NO_W = 2.5  # ຄວາມກວ້າງຖັນ ລ.ດ (%)

META_LABELS = [
    ("country",  "ປະເທດ"),
    ("motto",    "ຄຳຂັ້ນ"),
    ("title",    "ຊື່ບົດສະຫຼຸບ"),
    ("year",     "ປະຈຳປີ"),
    ("dept",     "ພະແນກ"),
    ("unit",     "ໜ່ວຍ"),
    ("docno",    "ເລກທີ"),
    ("place",    "ສະຖານທີ່, ວັນທີ"),
    ("signTitle", "ຕຳແໜ່ງຜູ້ລົງລາຍເຊັນ"),
    ("signName", "ຊື່ຜູ້ລົງລາຍເຊັນ"),
]

DEFAULT_META = {
    "country": "ສາທາລະນະລັດ ປະຊາທິປະໄຕ ປະຊາຊົນລາວ",
    "motto": "ສັນຕິພາບ ເອກະລາດ ປະຊາທິປະໄຕ ເອກະພາບ ວັດທະນາຖາວອນ.",
    "title": "ບົດສະຫຼຸບ ສະຖິຕິ ຂອງ ສະຫະພັນແມ່ຍິງ ໜ່ວຍຄູ",
    "year": "2026",
    "dept": "ພະແນກສະຫະພັນແມ່ຍິງ ວຕກ.",
    "unit": "ສະຫະພັນແມ່ຍິງ ໜ່ວຍຄູ",
    "docno": "ເລກທີ ....../ສພໜຄ.",
    "place": "ນະຄອນຫຼວງວຽງຈັນ, ວັນທີ ...... ຕຸລາ 2026",
    "signTitle": "ເລຂາສະຫະພັນແມ່ຍິງ ໜ່ວຍຄູ.",
    "signName": "",
}
DEFAULT_PREFS = {"font": "tnr", "fs": 7.5, "hfs": 11, "orient": "landscape"}

FONT_LABELS = {
    "ph": "Phetsarath OT",
    "tnr": "Times New Roman + Phetsarath OT",
}


def blank_row():
    return {k: "" for k in KEYS}


class Store:
    """ເກັບ/ໂຫຼດຂໍ້ມູນ JSON"""

    def __init__(self):
        self.meta = dict(DEFAULT_META)
        self.prefs = dict(DEFAULT_PREFS)
        self.rows = []
        self.load()

    def load(self, path=None):
        path = path or DATA_FILE
        if not os.path.exists(path):
            return
        with open(path, "r", encoding="utf-8") as f:
            d = json.load(f)
        self.meta = {**DEFAULT_META, **d.get("meta", {})}
        self.prefs = {**DEFAULT_PREFS, **d.get("prefs", {})}
        self.rows = [{**blank_row(), **r} for r in d.get("rows", [])]

    def save(self, path=None):
        path = path or DATA_FILE
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"meta": self.meta, "prefs": self.prefs, "rows": self.rows},
                      f, ensure_ascii=False, indent=1)
        os.replace(tmp, path)


# ----------------------------------------------------------------------------
#  ສົ່ງອອກ HTML (ສຳລັບພິມ Legal ຜ່ານບຼາວເຊີ -- ຈັດຮູບແບບພາສາລາວຖືກຕ້ອງທີ່ສຸດ)
# ----------------------------------------------------------------------------
HTML_TEMPLATE = """<!doctype html>
<html lang="lo"><head><meta charset="utf-8">
<title>%%PLAIN_TITLE%%</title>
<style>
@page{size:Legal %%ORIENT%%;margin:8mm}
*{box-sizing:border-box}
body{margin:0;font-family:%%FONT%%;color:#000}
.head{text-align:center;font-weight:700;font-size:%%HFS%%pt;line-height:1.55}
.title{text-align:center;font-weight:700;font-size:calc(%%HFS%%pt + 3pt);margin:3mm 0 4mm;line-height:1.5}
.meta{display:flex;justify-content:space-between;font-size:%%HFS%%pt;margin-bottom:2mm;line-height:1.6}
.meta .r{text-align:right}
table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:%%FS%%pt;line-height:1.35}
thead{display:table-header-group}
th,td{border:1px solid #000;padding:1.2mm 1mm;vertical-align:middle;overflow-wrap:anywhere}
th{font-weight:400;text-align:center}
td{text-align:center;height:7mm}
td.left{text-align:left}
td.red{color:#e00000}
tr{page-break-inside:avoid}
.sign{display:flex;justify-content:flex-end;margin-top:6mm;font-size:%%HFS%%pt;font-weight:700}
.sign div{text-align:center;min-width:70mm;line-height:1.6}
.space{height:22mm}
*{-webkit-print-color-adjust:exact;print-color-adjust:exact}
</style></head><body>
<div class="head">
%%LOGO%%
<div>%%COUNTRY%%</div><div>%%MOTTO%%</div><div>***********</div></div>
<div class="title">%%TITLE%%<br>ປະຈຳປີ %%YEAR%%.</div>
<div class="meta"><div><div>%%DEPT%%</div><div>%%UNIT%%</div></div>
<div class="r"><div>%%DOCNO%%</div><div>%%PLACE%%</div></div></div>
<table><colgroup>%%COLGROUP%%</colgroup>
<thead><tr>%%THEAD%%</tr></thead><tbody>
%%BODY%%
</tbody></table>
<div class="sign"><div><div>%%SIGNTITLE%%</div><div class="space"></div><div>%%SIGNNAME%%</div></div></div>
<script>window.addEventListener('load',()=>setTimeout(()=>window.print(),400));</script>
</body></html>
"""


def html_font_stack(mode):
    if mode == "tnr":
        return '"Times New Roman",Times,"Phetsarath OT","Phetsarath",serif'
    return '"Phetsarath OT","Phetsarath","Saysettha OT","Noto Sans Lao Looped",sans-serif'


def build_html(store, auto_print=True):
    e = html.escape
    m, p = store.meta, store.prefs
    colgroup = f'<col style="width:{NO_W}%">' + "".join(f'<col style="width:{w}%">' for _, _, w, _ in COLS)
    thead = "<th>ລ.ດ</th>" + "".join(f"<th>{e(h)}</th>" for _, h, _, _ in COLS)
    body = []
    for i, r in enumerate(store.rows, 1):
        tds = [f"<td>{i}</td>"]
        for k, _, _, left in COLS:
            v = r.get(k, "")
            cls = []
            if left:
                cls.append("left")
            if k == "youth" and "ອອກແລ້ວ" in v:
                cls.append("red")
            tds.append(f'<td class="{" ".join(cls)}">{e(v)}</td>')
        body.append("<tr>" + "".join(tds) + "</tr>")
    out = HTML_TEMPLATE
    rep = {
        "%%PLAIN_TITLE%%": e(m["title"]),
        "%%ORIENT%%": p["orient"],
        "%%FONT%%": html_font_stack(p["font"]),
        "%%HFS%%": str(p["hfs"]),
        "%%FS%%": str(p["fs"]),
        "%%LOGO%%": "",
        "%%COUNTRY%%": e(m["country"]), "%%MOTTO%%": e(m["motto"]),
        "%%TITLE%%": e(m["title"]), "%%YEAR%%": e(m["year"]),
        "%%DEPT%%": e(m["dept"]), "%%UNIT%%": e(m["unit"]),
        "%%DOCNO%%": e(m["docno"]), "%%PLACE%%": e(m["place"]),
        "%%COLGROUP%%": colgroup, "%%THEAD%%": thead, "%%BODY%%": "\n".join(body),
        "%%SIGNTITLE%%": e(m["signTitle"]), "%%SIGNNAME%%": e(m["signName"]),
    }
    for k, v in rep.items():
        out = out.replace(k, v)
    if not auto_print:
        out = out.replace("<script>window.addEventListener('load',()=>setTimeout(()=>window.print(),400));</script>", "")
    return out


# ----------------------------------------------------------------------------
#  ສົ່ງອອກ Word (.docx) -- ຕ້ອງ pip install python-docx
# ----------------------------------------------------------------------------
def export_docx(store, path):
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Mm, Pt, RGBColor

    m, p = store.meta, store.prefs
    latin = "Times New Roman" if p["font"] == "tnr" else "Phetsarath OT"
    lao = "Phetsarath OT"

    def style_run(run, size, bold=False, color=None):
        run.font.size = Pt(size)
        run.font.bold = bold
        if color:
            run.font.color.rgb = RGBColor(*color)
        rpr = run._element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.append(rf)
        rf.set(qn("w:ascii"), latin)
        rf.set(qn("w:hAnsi"), latin)
        rf.set(qn("w:eastAsia"), lao)
        rf.set(qn("w:cs"), lao)      # ຕົວອັກສອນລາວເປັນ complex script
        for tag, val in (("w:szCs", int(size * 2)), ("w:bCs", None)):
            el = OxmlElement(tag)
            if val is not None:
                el.set(qn("w:val"), str(val))
            elif not bold:
                continue
            rpr.append(el)

    def para(doc_or_cell, text, size, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER,
             color=None, first=None):
        par = first if first is not None else doc_or_cell.add_paragraph()
        par.alignment = align
        par.paragraph_format.space_after = Pt(0)
        par.paragraph_format.space_before = Pt(0)
        style_run(par.add_run(text), size, bold, color)
        return par

    doc = Document()
    sec = doc.sections[0]
    if p["orient"] == "landscape":
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = Mm(297), Mm(210)
    else:
        sec.orientation = WD_ORIENT.PORTRAIT
        sec.page_width, sec.page_height = Mm(210), Mm(297)
    for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(sec, side, Mm(8))

    hfs, fs = float(p["hfs"]), float(p["fs"])
    para(doc, m["country"], hfs, True, first=doc.paragraphs[0] if doc.paragraphs else None)
    para(doc, m["motto"], hfs, True)
    para(doc, "***********", hfs, True)
    para(doc, "", 4)
    para(doc, m["title"], hfs + 3, True)
    para(doc, f"ປະຈຳປີ {m['year']}.", hfs + 3, True)
    para(doc, "", 4)

    # ແຖວ meta ຊ້າຍ-ຂວາ ດ້ວຍຕາຕະລາງບໍ່ມີເສັ້ນ
    mt = doc.add_table(rows=2, cols=2)
    mt.autofit = True
    for r, (l, rr) in enumerate([(m["dept"], m["docno"]), (m["unit"], m["place"])]):
        para(mt.cell(r, 0), l, hfs, False, WD_ALIGN_PARAGRAPH.LEFT, first=mt.cell(r, 0).paragraphs[0])
        para(mt.cell(r, 1), rr, hfs, False, WD_ALIGN_PARAGRAPH.RIGHT, first=mt.cell(r, 1).paragraphs[0])
    para(doc, "", 4)

    total_w = (297 if p["orient"] == "landscape" else 210) - 16
    widths = [NO_W] + [c[2] for c in COLS]
    scale = total_w / sum(widths)

    tbl = doc.add_table(rows=1, cols=len(widths))
    tbl.style = "Table Grid"
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    # ຄວາມກວ້າງຄົງທີ່
    tblPr = tbl._tbl.tblPr
    lay = OxmlElement("w:tblLayout")
    lay.set(qn("w:type"), "fixed")
    tblPr.append(lay)

    def set_row_flags(row, header=False):
        trPr = row._tr.get_or_add_trPr()
        cs = OxmlElement("w:cantSplit")
        trPr.append(cs)
        if header:
            th = OxmlElement("w:tblHeader")
            trPr.append(th)

    heads = ["ລ.ດ"] + [c[1] for c in COLS]
    set_row_flags(tbl.rows[0], header=True)
    for i, h in enumerate(heads):
        c = tbl.rows[0].cells[i]
        c.width = Mm(widths[i] * scale)
        para(c, h, fs, False, first=c.paragraphs[0])

    for n, r in enumerate(store.rows, 1):
        row = tbl.add_row()
        set_row_flags(row)
        vals = [str(n)] + [r.get(k, "") for k in KEYS]
        for i, v in enumerate(vals):
            c = row.cells[i]
            c.width = Mm(widths[i] * scale)
            key = KEYS[i - 1] if i else None
            align = WD_ALIGN_PARAGRAPH.LEFT if (i and COLS[i - 1][3]) else WD_ALIGN_PARAGRAPH.CENTER
            color = (0xE0, 0, 0) if (key == "youth" and "ອອກແລ້ວ" in v) else None
            para(c, v, fs, False, align, color, first=c.paragraphs[0])

    para(doc, "", 6)
    sg = doc.add_table(rows=1, cols=2)
    cell = sg.cell(0, 1)
    para(cell, m["signTitle"], hfs, True, first=cell.paragraphs[0])
    for _ in range(3):
        para(cell, "", hfs)
    para(cell, m["signName"], hfs, True)

    doc.save(path)


def run_tk_app():
    """ເວີຊັນ Tkinter (ເກົ່າ) -- ໃຊ້ແທນໄດ້ຖ້າບໍ່ມີ PySide6; ການພິມລາວອາດບໍ່ເຮັດວຽກໃນບາງເຄື່ອງ"""
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from tkinter import font as tkfont

    # ----------------------------------------------------------------------------
    #  ໜ້າຈໍ GUI
    # ----------------------------------------------------------------------------
    _LAO = re.compile(r"[\u0e80-\u0eff]")
    _RUN = re.compile(r"[^\x00-\x7f\u0e80-\u0eff]+")


    def fix_mojibake(s):
        """ແກ້ຂໍ້ຄວາມລາວທີ່ຖືກປ້ອນເຂົ້າມາເປັນ UTF-8 bytes ຜິດ (ເຫັນເປັນ 'àº...')"""
        if not s:
            return s

        def repl(m):
            run = m.group(0)
            for enc in ("latin-1", "cp1252"):
                try:
                    t = run.encode(enc).decode("utf-8")
                except (UnicodeEncodeError, UnicodeDecodeError):
                    continue
                if _LAO.search(t):
                    return t
            return run

        return _RUN.sub(repl, s)


    def attach_fix(var, entry):
        """ຜູກຕົວແກ້ເຂົ້າກັບຊ່ອງປ້ອນຂໍ້ມູນ: ແກ້ທັນທີທີ່ພິມ ຫຼື ວາງ"""
        def cb(*_):
            s = var.get()
            t = fix_mojibake(s)
            if t != s:
                var.set(t)
                entry.icursor("end")
        var.trace_add("write", cb)


    class RowDialog(tk.Toplevel):
        """ຟອມເພີ່ມ/ແກ້ໄຂ 1 ລາຍຊື່"""

        def __init__(self, master, row, title, ui_font):
            super().__init__(master)
            self.title(title)
            self.transient(master)
            self.grab_set()
            self.result = None
            self.vars = {}
            frm = ttk.Frame(self, padding=10)
            frm.pack(fill="both", expand=True)
            for idx, (k, label, _, _) in enumerate(COLS):
                r, c = divmod(idx, 2)
                ttk.Label(frm, text=label, font=ui_font).grid(row=r, column=c * 2, sticky="e", padx=(8, 4), pady=3)
                v = tk.StringVar(value=row.get(k, ""))
                self.vars[k] = v
                ent = ttk.Entry(frm, textvariable=v, width=30, font=ui_font)
                ent.grid(row=r, column=c * 2 + 1, sticky="we", pady=3)
                attach_fix(v, ent)
            btns = ttk.Frame(frm)
            btns.grid(row=len(COLS) // 2 + 1, column=0, columnspan=4, pady=(10, 0))
            ttk.Button(btns, text="ບັນທຶກ", command=self.ok).pack(side="left", padx=5)
            ttk.Button(btns, text="ຍົກເລີກ", command=self.destroy).pack(side="left", padx=5)
            self.bind("<Return>", lambda e: self.ok())
            self.bind("<Escape>", lambda e: self.destroy())
            self.wait_visibility()
            self.focus_set()

        def ok(self):
            self.result = {k: fix_mojibake(v.get()).strip() for k, v in self.vars.items()}
            self.destroy()


    class SettingsDialog(tk.Toplevel):
        def __init__(self, master, store, ui_font):
            super().__init__(master)
            self.title("ຕັ້ງຄ່າ ແລະ ຫົວເອກະສານ")
            self.transient(master)
            self.grab_set()
            self.store = store
            self.ok_pressed = False
            frm = ttk.Frame(self, padding=10)
            frm.pack(fill="both", expand=True)
            self.mv = {}
            for i, (k, label) in enumerate(META_LABELS):
                ttk.Label(frm, text=label, font=ui_font).grid(row=i, column=0, sticky="e", padx=4, pady=3)
                v = tk.StringVar(value=store.meta.get(k, ""))
                self.mv[k] = v
                ent = ttk.Entry(frm, textvariable=v, width=60, font=ui_font)
                ent.grid(row=i, column=1, sticky="we", pady=3)
                attach_fix(v, ent)
            n = len(META_LABELS)
            ttk.Separator(frm).grid(row=n, column=0, columnspan=2, sticky="we", pady=8)
            ttk.Label(frm, text="ຟອນ", font=ui_font).grid(row=n + 1, column=0, sticky="e", padx=4)
            self.font_var = tk.StringVar(value=FONT_LABELS[store.prefs["font"]])
            ttk.Combobox(frm, textvariable=self.font_var, values=list(FONT_LABELS.values()),
                         state="readonly", width=40).grid(row=n + 1, column=1, sticky="w")
            ttk.Label(frm, text="ຂະໜາດຕາຕະລາງ (pt)", font=ui_font).grid(row=n + 2, column=0, sticky="e", padx=4, pady=3)
            self.fs = tk.StringVar(value=str(store.prefs["fs"]))
            ttk.Spinbox(frm, from_=5, to=14, increment=0.5, textvariable=self.fs, width=8).grid(row=n + 2, column=1, sticky="w")
            ttk.Label(frm, text="ຂະໜາດຫົວເອກະສານ (pt)", font=ui_font).grid(row=n + 3, column=0, sticky="e", padx=4, pady=3)
            self.hfs = tk.StringVar(value=str(store.prefs["hfs"]))
            ttk.Spinbox(frm, from_=8, to=18, increment=0.5, textvariable=self.hfs, width=8).grid(row=n + 3, column=1, sticky="w")
            ttk.Label(frm, text="ໜ້າເຈ້ຍ A4", font=ui_font).grid(row=n + 4, column=0, sticky="e", padx=4, pady=3)
            self.orient = tk.StringVar(value="ແນວນອນ" if store.prefs["orient"] == "landscape" else "ແນວຕັ້ງ")
            ttk.Combobox(frm, textvariable=self.orient, values=["ແນວນອນ", "ແນວຕັ້ງ"],
                         state="readonly", width=12).grid(row=n + 4, column=1, sticky="w")
            b = ttk.Frame(frm)
            b.grid(row=n + 5, column=0, columnspan=2, pady=(10, 0))
            ttk.Button(b, text="ບັນທຶກ", command=self.ok).pack(side="left", padx=5)
            ttk.Button(b, text="ຍົກເລີກ", command=self.destroy).pack(side="left", padx=5)

        def ok(self):
            try:
                fs, hfs = float(self.fs.get()), float(self.hfs.get())
            except ValueError:
                messagebox.showerror("ຜິດພາດ", "ຂະໜາດຕ້ອງເປັນຕົວເລກ", parent=self)
                return
            for k, v in self.mv.items():
                self.store.meta[k] = v.get()
            inv = {v: k for k, v in FONT_LABELS.items()}
            self.store.prefs.update(font=inv[self.font_var.get()], fs=fs, hfs=hfs,
                                    orient="landscape" if self.orient.get() == "ແນວນອນ" else "portrait")
            self.ok_pressed = True
            self.destroy()


    class App(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("ສະຖິຕິ ສະຫະພັນແມ່ຍິງ ໜ່ວຍຄູ")
            self.geometry("1350x700")
            self.store = Store()
            self.dirty = False

            families = set(tkfont.families())
            name = "Phetsarath OT" if "Phetsarath OT" in families else \
                   ("Saysettha OT" if "Saysettha OT" in families else "TkDefaultFont")
            self.ui_font = (name, 11) if name != "TkDefaultFont" else "TkDefaultFont"
            style = ttk.Style(self)
            style.configure("Treeview", font=self.ui_font, rowheight=26)
            style.configure("Treeview.Heading", font=self.ui_font)

            self._toolbar()
            self._tree()
            self.status = tk.StringVar()
            ttk.Label(self, textvariable=self.status, anchor="w", relief="sunken").pack(fill="x", side="bottom")
            if not self.store.rows:
                self.store.rows.append(blank_row())
            self.refresh()
            self.protocol("WM_DELETE_WINDOW", self.on_close)
            self.bind("<Control-s>", lambda e: self.save())
            self.bind("<Delete>", lambda e: self.delete())

        # ---------- UI ----------
        def _toolbar(self):
            tb = ttk.Frame(self, padding=4)
            tb.pack(fill="x")
            items = [
                ("＋ ເພີ່ມ", self.add), ("ແຊກເທິງ", lambda: self.insert(0)), ("ແຊກລຸ່ມ", lambda: self.insert(1)),
                ("✎ ແກ້ໄຂ", self.edit), ("🗑 ລຶບ", self.delete), ("▲ ຂຶ້ນ", lambda: self.move(-1)), ("▼ ລົງ", lambda: self.move(1)),
                None,
                ("💾 ບັນທຶກ", self.save), ("ຕັ້ງຄ່າ/ຫົວເອກະສານ", self.settings),
                None,
                ("ນຳເຂົ້າ JSON", self.import_json), ("ສົ່ງອອກ CSV", self.export_csv),
                ("ສົ່ງອອກ Word", self.export_word), ("🖨 ພິມ (HTML)", self.print_html),
            ]
            for it in items:
                if it is None:
                    ttk.Separator(tb, orient="vertical").pack(side="left", fill="y", padx=6)
                else:
                    ttk.Button(tb, text=it[0], command=it[1]).pack(side="left", padx=2)

        def _tree(self):
            frm = ttk.Frame(self)
            frm.pack(fill="both", expand=True)
            cols = ["no"] + KEYS
            self.tv = ttk.Treeview(frm, columns=cols, show="headings", selectmode="extended")
            self.tv.heading("no", text="ລ.ດ")
            self.tv.column("no", width=45, anchor="center", stretch=False)
            for k, h, w, left in COLS:
                self.tv.heading(k, text=h.replace(" ", "\n", 1) if False else h)
                self.tv.column(k, width=int(w * 16), anchor="w" if left else "center", stretch=False)
            ys = ttk.Scrollbar(frm, orient="vertical", command=self.tv.yview)
            xs = ttk.Scrollbar(frm, orient="horizontal", command=self.tv.xview)
            self.tv.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
            ys.pack(side="right", fill="y")
            xs.pack(side="bottom", fill="x")
            self.tv.pack(fill="both", expand=True)
            self.tv.bind("<Double-1>", lambda e: self.edit())
            self.tv.bind("<Return>", lambda e: self.edit())
            self.tv.tag_configure("odd", background="#f5f7fa")

        # ---------- helpers ----------
        def refresh(self, select=()):
            self.tv.delete(*self.tv.get_children())
            for i, r in enumerate(self.store.rows):
                self.tv.insert("", "end", iid=str(i), values=[i + 1] + [r[k] for k in KEYS],
                               tags=("odd",) if i % 2 else ())
            if select:
                self.tv.selection_set([str(i) for i in select if 0 <= i < len(self.store.rows)])
                if select:
                    self.tv.see(str(select[0]))
            self.set_status()

        def set_status(self, extra=""):
            self.status.set(f" ທັງໝົດ {len(self.store.rows)} ຄົນ   {'* ຍັງບໍ່ໄດ້ບັນທຶກ' if self.dirty else 'ບັນທຶກແລ້ວ'}   {extra}")

        def selected(self):
            return sorted(int(i) for i in self.tv.selection())

        def mark(self):
            self.dirty = True
            self.set_status()

        # ---------- actions ----------
        def add(self):
            d = RowDialog(self, blank_row(), "ເພີ່ມລາຍຊື່ໃໝ່", self.ui_font)
            self.wait_window(d)
            if d.result:
                self.store.rows.append(d.result)
                self.mark()
                self.refresh([len(self.store.rows) - 1])

        def insert(self, after):
            s = self.selected()
            if not s:
                return messagebox.showinfo("ແຈ້ງເຕືອນ", "ກະລຸນາເລືອກແຖວກ່ອນ")
            pos = s[-1] + 1 if after else s[0]
            d = RowDialog(self, blank_row(), "ແຊກລາຍຊື່ໃໝ່", self.ui_font)
            self.wait_window(d)
            if d.result:
                self.store.rows.insert(pos, d.result)
                self.mark()
                self.refresh([pos])

        def edit(self):
            s = self.selected()
            if not s:
                return messagebox.showinfo("ແຈ້ງເຕືອນ", "ກະລຸນາເລືອກແຖວກ່ອນ")
            i = s[0]
            d = RowDialog(self, self.store.rows[i], f"ແກ້ໄຂລາຍຊື່ ລຳດັບ {i + 1}", self.ui_font)
            self.wait_window(d)
            if d.result:
                self.store.rows[i] = d.result
                self.mark()
                self.refresh([i])

        def delete(self):
            s = self.selected()
            if not s:
                return
            if not messagebox.askyesno("ຢືນຢັນ", f"ລຶບ {len(s)} ແຖວທີ່ເລືອກ?"):
                return
            self.store.rows = [r for i, r in enumerate(self.store.rows) if i not in set(s)]
            self.mark()
            self.refresh([min(s[0], len(self.store.rows) - 1)])

        def move(self, d):
            s = self.selected()
            rows = self.store.rows
            if not s:
                return
            if d < 0:
                if s[0] == 0:
                    return
                for i in s:
                    rows[i - 1], rows[i] = rows[i], rows[i - 1]
                new = [i - 1 for i in s]
            else:
                if s[-1] == len(rows) - 1:
                    return
                for i in reversed(s):
                    rows[i + 1], rows[i] = rows[i], rows[i + 1]
                new = [i + 1 for i in s]
            self.mark()
            self.refresh(new)

        def save(self):
            self.store.save()
            self.dirty = False
            self.set_status("✔ ບັນທຶກແລ້ວ")

        def settings(self):
            d = SettingsDialog(self, self.store, self.ui_font)
            self.wait_window(d)
            if d.ok_pressed:
                self.mark()

        def import_json(self):
            p = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
            if not p:
                return
            try:
                self.store.load(p)
            except Exception as ex:
                return messagebox.showerror("ຜິດພາດ", f"ອ່ານໄຟລ໌ບໍ່ໄດ້: {ex}")
            if not self.store.rows:
                self.store.rows.append(blank_row())
            self.mark()
            self.refresh()

        def export_csv(self):
            p = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")],
                                             initialfile="womens_union_stats.csv")
            if not p:
                return
            with open(p, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(f)
                w.writerow(["ລ.ດ"] + [c[1] for c in COLS])
                for i, r in enumerate(self.store.rows, 1):
                    w.writerow([i] + [r[k] for k in KEYS])
            messagebox.showinfo("ສຳເລັດ", "ສົ່ງອອກ CSV ແລ້ວ (ເປີດດ້ວຍ Excel ໄດ້)")

        def export_word(self):
            try:
                import docx  # noqa: F401
            except ImportError:
                return messagebox.showerror("ຂາດ python-docx", "ກະລຸນາຕິດຕັ້ງກ່ອນ:\n\npip install python-docx")
            p = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")],
                                             initialfile="womens_union_stats.docx")
            if not p:
                return
            try:
                export_docx(self.store, p)
            except Exception as ex:
                return messagebox.showerror("ຜິດພາດ", str(ex))
            messagebox.showinfo("ສຳເລັດ", "ສົ່ງອອກ Word ແລ້ວ ເປີດໃນ Word ເພື່ອພິມ A4")
            try:
                os.startfile(p)  # Windows
            except Exception:
                pass

        def print_html(self):
            path = os.path.join(tempfile.gettempdir(), "womens_union_print.html")
            with open(path, "w", encoding="utf-8") as f:
                f.write(build_html(self.store, auto_print=True))
            webbrowser.open("file://" + path)

        def on_close(self):
            if self.dirty:
                a = messagebox.askyesnocancel("ອອກຈາກໂປຣແກຣມ", "ມີການປ່ຽນແປງທີ່ຍັງບໍ່ໄດ້ບັນທຶກ. ບັນທຶກກ່ອນອອກ?")
                if a is None:
                    return
                if a:
                    self.store.save()
            self.destroy()



    App().mainloop()


if __name__ == "__main__":
    run_tk_app()
