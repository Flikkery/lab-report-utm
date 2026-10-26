#!/usr/bin/env python3
"""Сборщик отчётов по лабораторным работам (УТМ, ф-т FCIM) в формате .docx.

    python build_report.py spec.json out.docx [--no-render] [--png-dir DIR] [--keep-pass1]

spec.json описан в references/spec_format.md. Скрипт:
  1. собирает docx по нормам тех. редактирования (references/norms.md);
  2. рендерит его в PDF (Word на Windows, иначе LibreOffice), находит страницы
     заголовков и пересобирает docx с настоящим полем «Оглавление» и верными номерами;
  3. сохраняет рядом PDF и (по --png-dir) PNG-страницы для визуальной проверки.
"""
import argparse
import datetime
import json
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH as AL, WD_TAB_ALIGNMENT, WD_TAB_LEADER, WD_BREAK
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.shared import Pt, Cm, Emu, RGBColor

SKILL_DIR = Path(__file__).resolve().parent.parent
ASSETS = SKILL_DIR / "assets"

TEXT_W_CM = 16.5          # 21 - 3 (слева) - 1.5 (справа)
BODY_FONT = "Times New Roman"
HEAD_FONT = "Arial"

LABELS = {
    "ru": dict(lab="по лабораторной работе № {n}", disc="по дисциплине «{d}»", topic="Тема: «{t}»",
               done="Выполнил студент: ", checked="Проверил: ", group="ст. гр. ",
               toc="СОДЕРЖАНИЕ", figure="Рисунок", table="Таблица", listing="Листинг",
               appendix="Приложение", lang="ru-RU", toc_empty="Обновите оглавление (F9)"),
    "ro": dict(lab="Lucrarea de laborator nr. {n}", disc="la disciplina «{d}»", topic="Tema: «{t}»",
               done="A efectuat studentul: ", checked="A controlat: ", group="gr. ",
               toc="CUPRINS", figure="Figura", table="Tabelul", listing="Listing",
               appendix="Anexa", lang="ro-RO", toc_empty="Actualizați cuprinsul (F9)"),
    "en": dict(lab="Laboratory work No. {n}", disc="Discipline: «{d}»", topic="Topic: «{t}»",
               done="Performed by: ", checked="Checked by: ", group="gr. ",
               toc="CONTENTS", figure="Figure", table="Table", listing="Listing",
               appendix="Appendix", lang="en-US", toc_empty="Update the table of contents (F9)"),
}
APPENDIX_LETTERS = {
    "ru": "АБВГДЕЖИКЛМНПРСТУФХЦШЩЭЮЯ",
    "ro": "ABCDEFGHIJKLMNOPQRSTUVWXYZ",
    "en": "ABCDEFGHIJKLMNOPQRSTUVWXYZ",
}
FIRST_PERSON = {
    "ru": r"\b(я|мы|мне|мной|меня|нам|нами|наш\w*|мой|моя|моё|мои)\b",
    "ro": r"\b(eu|noi|meu|mea|noastră|nostru)\b",
    "en": r"\b(I|we|my|our)\b",
}


# --------------------------------------------------------------------------- низкоуровневые помощники
def set_rfonts(rpr, name):
    rf = rpr.find(qn("w:rFonts"))
    if rf is not None:
        rpr.remove(rf)
    rf = OxmlElement("w:rFonts")
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rf.set(qn(a), name)
    rpr.insert(0, rf)


def fmt(run, font=BODY_FONT, size=None, bold=None, italic=None, caps=None):
    set_rfonts(run._r.get_or_add_rPr(), font)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if caps is not None:
        run.font.all_caps = caps
    return run


def pfmt(p, align=None, before=0, after=0, line=1.0, first=None, left=None, right=None,
         keep_next=None, keep_lines=None, page_break_before=None):
    f = p.paragraph_format
    if align is not None:
        p.alignment = align
    f.space_before = Pt(before)
    f.space_after = Pt(after)
    f.line_spacing = line
    if first is not None:
        f.first_line_indent = Cm(first)
    if left is not None:
        f.left_indent = Cm(left)
    if right is not None:
        f.right_indent = Cm(right)
    if keep_next is not None:
        f.keep_with_next = keep_next
    if keep_lines is not None:
        f.keep_together = keep_lines
    if page_break_before is not None:
        f.page_break_before = page_break_before
    return p


def mark_size(p, size_pt, font=BODY_FONT):
    """Размер шрифта знака абзаца — задаёт высоту пустой строки."""
    ppr = p._p.get_or_add_pPr()
    rpr = ppr.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        ppr.append(rpr)
    set_rfonts(rpr, font)
    for tag in ("w:sz", "w:szCs"):
        el = OxmlElement(tag)
        el.set(qn("w:val"), str(int(size_pt * 2)))
        rpr.append(el)


def fld_run(par, kind):
    r = par.add_run()
    fc = OxmlElement("w:fldChar")
    fc.set(qn("w:fldCharType"), kind)
    r._r.append(fc)
    return r


def instr_run(par, text):
    r = par.add_run()
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = text
    r._r.append(it)
    return r


# --------------------------------------------------------------------------- разметка внутри текста
TOKEN = re.compile(
    r"(\*\*[^*]+?\*\*|(?<![\w*])\*(?!\s)[^*]+?(?<!\s)\*(?![\w*])|`[^`]+`|\^\{[^}]+\}|_\{[^}]+\}"
    r"|\{(?:fig|tab|eq|lst|src):[^}]+\})")


class Report:
    def __init__(self, spec, pages=None):
        self.spec = spec
        self.meta = spec["meta"]
        self.lang = self.meta.get("lang", "ru")
        self.L = LABELS[self.lang]
        self.profile = json.loads((ASSETS / "profile.json").read_text(encoding="utf-8"))
        self.pages = pages or {}
        self.warnings = []
        self.labels = {}        # id -> номер для {fig:id} и т.п.
        self.src_index = {}
        self.headings = []      # (level, text) — для оглавления
        self.doc = Document()

    # ----- нумерация -----------------------------------------------------------------------------
    class Counters:
        def __init__(self, letters):
            self.ch = self.sub = self.subsub = 0
            self.app = None
            self.app_i = -1
            self.letters = letters
            self.n = {}

        def scope(self):
            if self.app is not None:
                return self.app
            if self.ch == 0:
                return ""
            return f"{self.ch}.{self.sub}" if self.sub else f"{self.ch}"

        def next(self, kind):
            sc = self.scope()
            self.n[(sc, kind)] = self.n.get((sc, kind), 0) + 1
            k = self.n[(sc, kind)]
            return f"{sc}.{k}" if sc else str(k)

        def h1(self, numbered):
            self.app = None
            self.sub = self.subsub = 0
            if numbered:
                self.ch += 1
                return str(self.ch)
            return ""

        def h2(self):
            self.sub += 1
            self.subsub = 0
            return f"{self.ch}.{self.sub}"

        def h3(self):
            self.subsub += 1
            return f"{self.ch}.{self.sub}.{self.subsub}"

        def appendix(self):
            self.app_i += 1
            self.app = self.letters[self.app_i]
            return self.app

    def prepass(self):
        c = self.Counters(APPENDIX_LETTERS[self.lang])
        last = None
        for i, b in enumerate(self.spec["blocks"]):
            t = b["type"]
            if t == "heading":
                lv = b.get("level", 1)
                if lv == 1:
                    c.h1(b.get("numbered", True))
                elif lv == 2:
                    c.h2()
                else:
                    c.h3()
            elif t == "appendix":
                c.appendix()
            elif t in ("figure", "table", "code", "equation"):
                kind = {"figure": "fig", "table": "tab", "code": "lst", "equation": "eq"}[t]
                if t == "code" and not b.get("caption"):
                    pass
                else:
                    num = c.next(kind)
                    if t == "equation":
                        num = f"({num})"
                    if b.get("id"):
                        self.labels[(kind, b["id"])] = num.strip("()") if t == "equation" else num
            elif t == "sources":
                for j, it in enumerate(b["items"], 1):
                    if isinstance(it, dict) and it.get("id"):
                        self.src_index[it["id"]] = j
            # правило «две таблицы / два рисунка подряд — нельзя»
            if t in ("figure", "table") and last == t:
                self.warnings.append(
                    f"БЛОК {i}: два подряд «{t}» — между ними по нормам должен быть текст. "
                    f"Добавьте поясняющий абзац.")
            last = t if t in ("figure", "table") else None
            if t == "code" and b.get("caption") and last is None:
                pass

    # ----- inline ---------------------------------------------------------------------------------
    def resolve(self, kind, key):
        if kind == "src":
            out = []
            for k in key.split(","):
                k = k.strip()
                if k not in self.src_index:
                    self.warnings.append(f"Ссылка {{src:{k}}}: источник не найден в блоке sources.")
                    out.append("?")
                else:
                    out.append(str(self.src_index[k]))
            return "[" + ", ".join(out) + "]"
        v = self.labels.get((kind, key.strip()))
        if v is None:
            self.warnings.append(f"Ссылка {{{kind}:{key}}}: id не найден.")
            return "?"
        return v

    def inline(self, par, text, font=BODY_FONT, size=12, bold=None, italic=None):
        for part in TOKEN.split(text):
            if not part:
                continue
            r = None
            if part.startswith("**") and part.endswith("**") and len(part) > 4:
                r = fmt(par.add_run(part[2:-2]), font, size, True, italic)
            elif part.startswith("`") and part.endswith("`"):
                r = fmt(par.add_run(part[1:-1]), "Consolas", max(size - 1, 9), bold, italic)
            elif part.startswith("^{"):
                r = fmt(par.add_run(part[2:-1]), font, size, bold, italic)
                r.font.superscript = True
            elif part.startswith("_{"):
                r = fmt(par.add_run(part[2:-1]), font, size, bold, italic)
                r.font.subscript = True
            elif part.startswith("{") and ":" in part and part.endswith("}"):
                kind, key = part[1:-1].split(":", 1)
                r = fmt(par.add_run(self.resolve(kind, key)), font, size, bold, italic)
            elif part.startswith("*") and part.endswith("*") and len(part) > 2:
                r = fmt(par.add_run(part[1:-1]), font, size, bold, True)
            else:
                r = fmt(par.add_run(part), font, size, bold, italic)
        return par

    # ----- документ -------------------------------------------------------------------------------
    def setup(self):
        doc = self.doc
        sec = doc.sections[0]
        sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
        sec.top_margin = sec.bottom_margin = Cm(2)
        sec.left_margin, sec.right_margin = Cm(3), Cm(1.5)
        sec.header_distance = sec.footer_distance = Cm(1.25)

        st = doc.styles["Normal"]
        st.font.size = Pt(12)
        set_rfonts(st.element.get_or_add_rPr(), BODY_FONT)
        lang = OxmlElement("w:lang")
        lang.set(qn("w:val"), self.L["lang"])
        lang.set(qn("w:eastAsia"), "ru-RU")
        st.element.get_or_add_rPr().append(lang)
        pf = st.paragraph_format
        pf.alignment = AL.JUSTIFY
        pf.space_before = pf.space_after = Pt(0)
        pf.line_spacing = 1.5
        pf.widow_control = True

        for name, size, align, before, after in (
                ("Heading 1", 14, AL.CENTER, 12, 8),
                ("Heading 2", 12, AL.LEFT, 12, 8),
                ("Heading 3", 11, AL.LEFT, 12, 6)):
            h = doc.styles[name]
            h.font.size = Pt(size)
            h.font.bold = True
            h.font.italic = False
            h.font.color.rgb = RGBColor(0, 0, 0)
            set_rfonts(h.element.get_or_add_rPr(), HEAD_FONT)
            h.paragraph_format.alignment = align
            h.paragraph_format.space_before = Pt(before)
            h.paragraph_format.space_after = Pt(after)
            h.paragraph_format.line_spacing = 1.0
            h.paragraph_format.keep_with_next = True
            h.paragraph_format.first_line_indent = Cm(0)
            h.paragraph_format.left_indent = Cm(0)

        for lvl, indent in ((1, 0.0), (2, 0.5), (3, 1.0)):
            s = doc.styles.add_style(f"toc {lvl}", WD_STYLE_TYPE.PARAGRAPH)
            s.base_style = doc.styles["Normal"]
            s.font.size = Pt(12)
            set_rfonts(s.element.get_or_add_rPr(), BODY_FONT)
            s.paragraph_format.alignment = AL.LEFT
            s.paragraph_format.left_indent = Cm(indent)
            s.paragraph_format.space_after = Pt(4)
            s.paragraph_format.line_spacing = 1.25
            s.paragraph_format.tab_stops.add_tab_stop(Cm(TEXT_W_CM), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

        student = self.meta.get("student") or self.profile["student"].get(self.lang) or self.profile["student"]["ru"]
        cp = doc.core_properties
        cp.author = student
        cp.title = self.meta.get("topic", "")
        cp.language = self.L["lang"]

    # ----- титульный лист ------------------------------------------------------------------------
    def logo_anchor(self, run):
        """Логотип УТМ: плавающий, слева, текст обтекает справа (как в эталоне)."""
        from PIL import Image
        logo = ASSETS / "utm_logo.png"
        w_emu, h_emu = 1687989, 1446848
        run.add_picture(str(logo), width=Emu(w_emu), height=Emu(h_emu))
        drawing = run._r.find(qn("w:drawing"))
        inline = drawing.find(qn("wp:inline"))
        graphic = inline.find(qn("a:graphic"))
        anchor = parse_xml(
            f'<wp:anchor {nsdecls("wp", "a", "pic", "r")} distT="114300" distB="114300" distL="114300" '
            f'distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" '
            f'layoutInCell="1" allowOverlap="1">'
            f'<wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="column"><wp:posOffset>-295274</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>114300</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{w_emu}" cy="{h_emu}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:wrapSquare wrapText="bothSides"/>'
            f'<wp:docPr id="1" name="UTM logo" descr="Logo UTM"/><wp:cNvGraphicFramePr/></wp:anchor>')
        anchor.append(graphic)
        drawing.remove(inline)
        drawing.append(anchor)

    def title_page(self):
        doc, m, L, prof = self.doc, self.meta, self.L, self.profile
        student = m.get("student") or prof["student"].get(self.lang) or prof["student"]["ru"]
        group = m.get("group") or prof["group"]
        year = str(m.get("year") or datetime.date.today().year)
        faculty = m.get("faculty") or prof["faculty"]
        department = m.get("department")

        # 1. «шапка»: логотип слева + центрированный текст справа
        p = doc.add_paragraph()
        pfmt(p, AL.CENTER, line=1.15)
        self.logo_anchor(p.add_run())
        lines = [(prof["ministry"], 10.5), (prof["university"], 15), (faculty, 11.5)]
        if department:
            lines.append((department, 11.5))
        for i, (txt, sz) in enumerate(lines):
            r = fmt(p.add_run(txt), HEAD_FONT, sz, True)
            r.add_break()       # как в эталоне: после кафедры остаётся пустая строка
        # 2. воздух под логотипом (две пустые строки, как в эталоне)
        e1 = doc.add_paragraph()
        pfmt(e1, AL.CENTER, before=8, after=19, line=1.0)
        e2 = doc.add_paragraph()
        pfmt(e2, AL.CENTER, before=10, after=15, line=1.25)
        mark_size(e2, 18, HEAD_FONT)
        # 3. вид работы, номер, дисциплина, тема
        p = doc.add_paragraph()
        pfmt(p, AL.CENTER, before=10, after=15, line=1.25)
        rows = [(m.get("report_title", "RAPORT"), 18)]
        if m.get("lab_number") is not None:
            rows.append((L["lab"].format(n=m["lab_number"]), 13))
        rows.append((L["disc"].format(d=m["discipline"]), 12))
        if m.get("discipline_translation"):
            rows.append((f"(«{m['discipline_translation']}»)", 12))
        rows += [("", 12), (L["topic"].format(t=m["topic"]), 12), ("", 12)]
        for i, (txt, sz) in enumerate(rows):
            r = fmt(p.add_run(txt), HEAD_FONT, sz, True)
            if i < len(rows) - 1:
                r.add_break()
        # 4. «Выполнил / Проверил» — прижат вправо
        p = doc.add_paragraph()
        pfmt(p, AL.RIGHT, before=35, after=50, line=1.3, left=7.62, right=0.89)
        fmt(p.add_run(L["done"]), BODY_FONT, 12)
        r = fmt(p.add_run(f"{L['group']}{group}"), BODY_FONT, 12, True)
        r.add_break()
        r = fmt(p.add_run(student), BODY_FONT, 12, True)
        r.add_break()
        r.add_break()
        fmt(p.add_run(L["checked"]), BODY_FONT, 12)
        fmt(p.add_run(m["teacher"]), BODY_FONT, 12, True)
        # 5. «CHIȘINĂU – год» — прибит к низу страницы рамкой
        p = doc.add_paragraph()
        pfmt(p, AL.CENTER, line=1.0)
        ppr = p._p.get_or_add_pPr()
        frame = parse_xml(
            f'<w:framePr {nsdecls("w")} w:w="{int(TEXT_W_CM * 567)}" w:hSpace="0" w:wrap="around" '
            f'w:vAnchor="margin" w:hAnchor="margin" w:xAlign="center" w:y="13920"/>')
        ppr.insert_element_before(frame, "w:widowControl", "w:numPr", "w:pBdr", "w:shd", "w:tabs",
                                  "w:suppressAutoHyphens", "w:spacing", "w:ind", "w:jc", "w:rPr")
        fmt(p.add_run(f"{prof['city']} – {year}"), HEAD_FONT, 12, True)
        tail = doc.add_paragraph()      # обычный абзац — держит рамку на первой странице
        pfmt(tail, AL.CENTER, line=1.0)

    # ----- оглавление ----------------------------------------------------------------------------
    def toc_placeholder(self):
        """Заголовок + настоящее поле TOC (кэш заполняется на втором проходе)."""
        doc = self.doc
        h = doc.add_paragraph()
        pfmt(h, AL.CENTER, before=12, after=8, line=1.0, page_break_before=True, keep_next=True)
        fmt(h.add_run(self.L["toc"]), HEAD_FONT, 14, True)
        self._toc_anchor = doc.add_paragraph()      # сюда вставим поле после сборки тела

    def fill_toc(self):
        doc = self.doc
        anchor = self._toc_anchor
        entries = self.headings or [(1, self.L["toc_empty"])]
        pars = []
        for lvl, text in entries:
            p = doc.add_paragraph(style=f"toc {min(lvl, 3)}")
            fmt(p.add_run(text), BODY_FONT, 12)
            r = p.add_run()
            r.add_tab()
            fmt(r, BODY_FONT, 12)
            fmt(p.add_run(str(self.pages.get(text, 0)) if self.headings else ""), BODY_FONT, 12)
            pars.append(p)
        fld_run(pars[0], "begin")
        instr_run(pars[0], ' TOC \\o "1-3" \\u \\z ')
        fld_run(pars[0], "separate")
        # begin/instr/separate должны идти ПЕРЕД текстом первого абзаца
        first = pars[0]._p
        runs = first.findall(qn("w:r"))
        for rr in runs[-3:][::-1]:
            first.remove(rr)
            first.insert(1 if first.find(qn("w:pPr")) is not None else 0, rr)
        fld_run(pars[-1], "end")
        for p in pars:
            anchor._p.addprevious(p._p)
        anchor._p.getparent().remove(anchor._p)

    # ----- блоки ---------------------------------------------------------------------------------
    def body_par(self, text, first_indent=1.5):
        p = self.doc.add_paragraph()
        pfmt(p, AL.JUSTIFY, line=1.5, first=first_indent)
        self.inline(p, text)
        fp = FIRST_PERSON[self.lang]
        if re.search(fp, text, re.I | re.U):
            self.warnings.append(f"Возможное первое лицо (нужен безличный стиль): «{text[:70]}…»")
        return p

    def heading(self, b, c):
        lv = b.get("level", 1)
        text = b["text"]
        if lv == 1:
            num = c.h1(b.get("numbered", True))
        elif lv == 2:
            num = c.h2()
        else:
            num = c.h3()
        display = f"{num} {text}" if (num and b.get("numbered", True)) else text
        self.add_heading(display, lv)

    def add_heading(self, display, lv):
        p = self.doc.add_paragraph(style=f"Heading {lv}")
        fmt(p.add_run(display), HEAD_FONT, {1: 14, 2: 12, 3: 11}[lv], True, caps=(lv == 1))
        if lv == 1:
            p.paragraph_format.page_break_before = not self._first_h1
            self._first_h1 = False
        self.headings.append((lv, display))
        return p

    def caption(self, kind_label, number, text, align, size, above=False):
        p = self.doc.add_paragraph()
        if above:
            pfmt(p, align, before=6, after=3, line=1.0, keep_next=True)
        else:
            pfmt(p, align, before=3, after=12, line=1.0)
        fmt(p.add_run(f"{kind_label} {number}"), BODY_FONT, size, True)
        self.inline(p, f" — {text}" if text else "", size=size)
        return p

    def figure(self, b, c):
        from PIL import Image
        path = Path(b["image"])
        if not path.is_absolute():
            path = self.base_dir / path
        if not path.exists():
            raise FileNotFoundError(f"Нет изображения: {path}")
        with Image.open(path) as im:
            w_px, h_px = im.size
        width = float(b.get("width_cm", min(TEXT_W_CM, 14.0)))
        height = width * h_px / w_px
        max_h = float(b.get("max_height_cm", 20.0))
        if height > max_h:
            width, height = width * max_h / height, max_h
        p = self.doc.add_paragraph()
        pfmt(p, AL.CENTER, before=6, line=1.0, keep_next=True)
        p.add_run().add_picture(str(path), width=Cm(width))
        num = c.next("fig")
        self.caption(self.L["figure"], num, b.get("caption", ""), AL.CENTER, 11)

    def table(self, b, c):
        header, rows = b["header"], b["rows"]
        n = len(header)
        widths = b.get("col_widths_cm") or [TEXT_W_CM / n] * n
        aligns = b.get("align") or ["l"] * n
        amap = {"l": AL.LEFT, "c": AL.CENTER, "r": AL.RIGHT}
        num = c.next("tab")
        cap_align = {"left": AL.LEFT, "right": AL.RIGHT, "center": AL.CENTER}[
            self.meta.get("table_caption_align", "right")]
        self.caption(self.L["table"], num, b.get("caption", ""), cap_align, 11, above=True)
        t = self.doc.add_table(rows=1 + len(rows), cols=n)
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        tblPr = t._tbl.tblPr
        lay = OxmlElement("w:tblLayout")
        lay.set(qn("w:type"), "fixed")
        tblPr.append(lay)
        for ri, row in enumerate([header] + rows):
            tr = t.rows[ri]
            trPr = tr._tr.get_or_add_trPr()
            cs = OxmlElement("w:cantSplit")
            trPr.append(cs)
            if ri == 0:
                trPr.append(OxmlElement("w:tblHeader"))
            for ci in range(n):
                cell = tr.cells[ci]
                cell.width = Cm(widths[ci])
                tcPr = cell._tc.get_or_add_tcPr()
                tcPr.append(parse_xml(
                    f'<w:tcBorders {nsdecls("w")}>' + "".join(
                        f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="D9D9D9"/>'
                        for s in ("top", "left", "bottom", "right")) + "</w:tcBorders>"))
                if ri == 0:
                    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="E7E6E6"/>'))
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                par = cell.paragraphs[0]
                pfmt(par, AL.CENTER if ri == 0 and aligns[ci] != "l" else amap[aligns[ci]],
                     before=3, after=3, line=1.0)
                self.inline(par, str(row[ci]), size=11, bold=True if ri == 0 else None)
        gap = self.doc.add_paragraph()      # «1 enter» после таблицы
        pfmt(gap, AL.LEFT, line=1.0)
        mark_size(gap, 12)

    def code(self, b, c, console=False):
        text = b["text"].rstrip("\n").replace("\t", "    ")
        lines = text.split("\n")
        size = float(b.get("size", 9.5))
        for i, ln in enumerate(lines):
            p = self.doc.add_paragraph()
            pfmt(p, AL.LEFT, line=1.0, first=0, left=0.3)
            p.paragraph_format.keep_together = True
            ppr = p._p.get_or_add_pPr()
            shd = parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="F2F2F2"/>')
            bdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:left w:val="single" w:sz="18" w:space="6" w:color="A6A6A6"/></w:pBdr>')
            ppr.insert_element_before(bdr, "w:shd", "w:tabs", "w:spacing", "w:ind", "w:jc", "w:rPr")
            ppr.insert_element_before(shd, "w:tabs", "w:spacing", "w:ind", "w:jc", "w:rPr")
            fmt(p.add_run(ln if ln else " "), "Consolas", size)
        if b.get("caption"):
            num = c.next("lst")
            self.caption(self.L["listing"], num, b["caption"], AL.CENTER, 11)
        else:
            gap = self.doc.add_paragraph()
            pfmt(gap, AL.LEFT, line=1.0)
            mark_size(gap, 8)

    def equation(self, b, c):
        num = c.next("eq")
        p = self.doc.add_paragraph()
        pfmt(p, AL.LEFT, before=6, after=6, line=1.5, first=0)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(TEXT_W_CM / 2), WD_TAB_ALIGNMENT.CENTER)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(TEXT_W_CM), WD_TAB_ALIGNMENT.RIGHT)
        p.add_run().add_tab()
        self.inline(p, b["text"], italic=b.get("italic", True))
        p.add_run().add_tab()
        fmt(p.add_run(f"({num})"), BODY_FONT, 12)

    def listing(self, items, numbered):
        for i, it in enumerate(items, 1):
            p = self.doc.add_paragraph()
            pfmt(p, AL.JUSTIFY, line=1.5, left=2.0, first=-0.5)
            self.inline(p, f"{i}) " if numbered else "– ")
            self.inline(p, it)

    def sources(self, b):
        self.add_heading(b.get("title", "БИБЛИОГРАФИЯ"), 1)
        for j, it in enumerate(b["items"], 1):
            text = it["text"] if isinstance(it, dict) else it
            p = self.doc.add_paragraph()
            pfmt(p, AL.JUSTIFY, line=1.5, left=1.5, first=-1.0)
            fmt(p.add_run(f"{j}. "), BODY_FONT, 12)
            self.inline(p, text)

    # ----- сборка --------------------------------------------------------------------------------
    def build(self, out, base_dir):
        self.base_dir = Path(base_dir)
        self.prepass()
        self.setup()
        self._first_h1 = True
        self.title_page()
        self.toc_placeholder()
        # новый раздел: титул и оглавление не нумеруются, но входят в нумерацию страниц
        self.doc.add_section(WD_SECTION.NEW_PAGE)
        sec2 = self.doc.sections[1]
        sec2.footer.is_linked_to_previous = False
        fp = sec2.footer.paragraphs[0]
        pfmt(fp, AL.CENTER, line=1.0)
        fld_run(fp, "begin")
        instr_run(fp, " PAGE ")
        fld_run(fp, "separate")
        fmt(fp.add_run("1"), BODY_FONT, 12)
        fld_run(fp, "end")

        c = self.Counters(APPENDIX_LETTERS[self.lang])
        for b in self.spec["blocks"]:
            t = b["type"]
            if t == "heading":
                self.heading(b, c)
            elif t == "appendix":
                letter = c.appendix()
                self.add_heading(f"{self.L['appendix']} {letter}. {b['title']}", 1)
            elif t in ("p", "paragraph"):
                self.body_par(b["text"], b.get("indent", 1.5))
            elif t == "bullets":
                self.listing(b["items"], False)
            elif t == "numbered":
                self.listing(b["items"], True)
            elif t == "figure":
                self.figure(b, c)
            elif t == "table":
                self.table(b, c)
            elif t == "code":
                self.code(b, c)
            elif t == "console":
                self.code(b, c, console=True)
            elif t == "equation":
                self.equation(b, c)
            elif t == "sources":
                self.sources(b)
            elif t == "pagebreak":
                self.doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
            else:
                raise ValueError(f"Неизвестный тип блока: {t}")
        self.fill_toc()
        self.doc.save(out)


# --------------------------------------------------------------------------- рендер и номера страниц
def render_pdf(docx_path, pdf_path):
    docx_path, pdf_path = Path(docx_path).resolve(), Path(pdf_path).resolve()
    if sys.platform == "win32":
        ps1 = SKILL_DIR / "scripts" / "render_docx.ps1"
        r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps1),
                            "-In", str(docx_path), "-Out", str(pdf_path)],
                           capture_output=True, text=True, encoding="utf-8", errors="replace")
        if r.returncode == 0 and pdf_path.exists():
            return True
        print("Word-рендер не удался:", (r.stderr or r.stdout).strip()[:400], file=sys.stderr)
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(pdf_path.parent),
                        str(docx_path)], capture_output=True)
        produced = pdf_path.parent / (docx_path.stem + ".pdf")
        if produced.exists():
            if produced != pdf_path:
                shutil.move(str(produced), str(pdf_path))
            return True
    return False


def norm(s):
    return re.sub(r"\s+", " ", s.replace(" ", " ")).strip().casefold()


def find_pages(pdf_path, headings):
    import fitz
    doc = fitz.open(str(pdf_path))
    pages_lines = [[norm(x) for x in pg.get_text().split("\n") if x.strip()] for pg in doc]
    # страницы оглавления: со 2-й, пока на них есть точки-заполнители
    toc_end = 1
    for i in range(1, len(pages_lines)):
        if sum(1 for ln in pages_lines[i] if "....." in ln or "…" in ln) >= 2:
            toc_end = i
        else:
            break
    result, cur = {}, toc_end + 1
    for lv, text in headings:
        target = norm(text)
        found = None
        for pg in range(cur, len(pages_lines)):
            lines = pages_lines[pg]
            for k in range(len(lines)):
                joined = lines[k]
                if joined == target:
                    found = pg
                    break
                for k2 in range(k + 1, min(k + 4, len(lines))):
                    joined = f"{joined} {lines[k2]}"
                    if joined == target:
                        found = pg
                        break
                if found is not None:
                    break
            if found is not None:
                break
        if found is None:
            print(f"! Не найден заголовок в PDF: {text}", file=sys.stderr)
            continue
        result[text] = found + 1
        cur = found
    return result, len(pages_lines)


def export_png(pdf_path, out_dir, dpi=70):
    import fitz
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    d = fitz.open(str(pdf_path))
    for i, pg in enumerate(d, 1):
        pg.get_pixmap(dpi=dpi).save(str(out / f"page_{i:02d}.png"))
    return len(d)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("spec")
    ap.add_argument("out")
    ap.add_argument("--no-render", action="store_true", help="не рендерить PDF (оглавление без номеров)")
    ap.add_argument("--png-dir", help="сохранить страницы PDF как PNG для проверки глазами")
    ap.add_argument("--allow-consecutive", action="store_true",
                    help="не считать ошибкой два рисунка/таблицы подряд")
    args = ap.parse_args()

    spec_path = Path(args.spec).resolve()
    spec = json.loads(spec_path.read_text(encoding="utf-8"))
    out = Path(args.out).resolve()
    out.parent.mkdir(parents=True, exist_ok=True)
    pdf = out.with_suffix(".pdf")

    rep = Report(spec)
    rep.build(out, spec_path.parent)
    hard = [w for w in rep.warnings if "подряд" in w]
    if hard and not args.allow_consecutive:
        print("ОШИБКА НОРМ ОФОРМЛЕНИЯ:\n  " + "\n  ".join(hard), file=sys.stderr)
        sys.exit(2)

    if not args.no_render:
        if render_pdf(out, pdf):
            pages, total = find_pages(pdf, rep.headings)
            rep2 = Report(spec, pages)
            rep2.build(out, spec_path.parent)
            render_pdf(out, pdf)
            print(f"Страниц: {total}; заголовков с номерами: {len(pages)}/{len(rep.headings)}")
            if args.png_dir:
                export_png(pdf, args.png_dir)
                print(f"PNG страниц: {args.png_dir}")
        else:
            print("Рендер PDF недоступен (нет Word/LibreOffice): оглавление сохранено без номеров "
                  "страниц — откройте docx в Word и нажмите F9 по оглавлению.", file=sys.stderr)
    for w in sorted(set(rep.warnings)):
        if "подряд" not in w:
            print("ПРЕДУПРЕЖДЕНИЕ:", w, file=sys.stderr)
    print(f"Готово: {out}")


if __name__ == "__main__":
    main()
