"""Создает .docx по ГОСТу для СРС/отчета. Использование:
python3 create_docx.py --json input.json --out result.docx
input.json: {"title_contents": "СОДЕРЖАНИЕ", "sections": [{"heading": "...", "level": 1, "pagebreak_before": true, "paragraphs": ["..."]}], "bibliography": ["1. ..."]}
Если входного json нет — скрипт создает демо-файл чтобы проверить что python-docx работает.
"""
import json, sys, argparse
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_margins(section):
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    section.left_margin = Mm(30)
    section.right_margin = Mm(10)
    section.top_margin = Mm(20)
    section.bottom_margin = Mm(20)

def add_page_number(run):
    fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = 'PAGE'
    fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'end')
    run._r.append(fld1); run._r.append(instr); run._r.append(fld2)

def add_toc(paragraph):
    run = paragraph.add_run()
    fld1 = OxmlElement('w:fldChar'); fld1.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve')
    instr.text = 'TOC \\o "1-2" \\h \\z \\u'
    fld2 = OxmlElement('w:fldChar'); fld2.set(qn('w:fldCharType'), 'separate')
    fld3 = OxmlElement('w:fldChar'); fld3.set(qn('w:fldCharType'), 'end')
    run._r.append(fld1); run._r.append(instr); run._r.append(fld2); run._r.append(fld3)

def style_document(doc):
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(14)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    pf = style.paragraph_format
    pf.space_after = Pt(0); pf.space_before = Pt(0)
    pf.line_spacing = 1.5
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for i in (1, 2):
        hs = doc.styles[f'Heading {i}']
        hs.font.name = 'Times New Roman'
        hs.font.size = Pt(14)
        hs.font.bold = True
        hs.font.color.rgb = RGBColor(0, 0, 0)
        hs.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')

def add_heading_styled(doc, text, level, break_before=False):
    if break_before:
        doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
    h = doc.add_heading(level=level)
    run = h.add_run(text)
    run.bold = True; run.font.size = Pt(14); run.font.name = 'Times New Roman'
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER if level == 1 else WD_ALIGN_PARAGRAPH.LEFT
    h.paragraph_format.space_after = Pt(6)
    h.paragraph_format.first_line_indent = Mm(0)
    return h

def add_body(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Mm(12.5)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(text)
    run.font.size = Pt(14); run.font.name = 'Times New Roman'
    return p

def build(doc_data, out):
    doc = Document()
    set_margins(doc.sections[0])
    style_document(doc)
    # footer page number
    footer = doc.sections[0].footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = footer.add_run(); r.font.size = Pt(12); r.font.name = 'Times New Roman'
    add_page_number(r)
    # contents
    add_heading_styled(doc, 'СОДЕРЖАНИЕ', level=1)
    toc_p = doc.add_paragraph()
    add_toc(toc_p)
    # ВАЖНО: никаких подсказок внутрь документа не добавляем.
    # python-docx вставляет только код поля TOC, сами номера страниц посчитает Word
    # при открытии. Подсказку про "Обновить поле" пиши только в чат при сдаче, не в docx.
    for sec in doc_data.get('sections', []):
        add_heading_styled(doc, sec['heading'], level=sec.get('level', 1), break_before=True)
        for par in sec.get('paragraphs', []):
            add_body(doc, par)
    if doc_data.get('bibliography'):
        add_heading_styled(doc, 'СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ', level=1, break_before=True)
        for i, src in enumerate(doc_data['bibliography'], 1):
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Mm(0)
            r = p.add_run(src); r.font.size = Pt(14); r.font.name = 'Times New Roman'
    doc.save(out)
    print(f'saved to {out}')

if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--json', default='')
    ap.add_argument('--out', default='/tmp/srs_demo.docx')
    a = ap.parse_args()
    if a.json:
        with open(a.json, encoding='utf-8') as f:
            data = json.load(f)
    else:
        data = {'sections': [{'heading': 'ВВЕДЕНИЕ', 'level': 1, 'pagebreak_before': True, 'paragraphs': ['Демо-текст. Проверка ГОСТ-форматирования.']}], 'bibliography': []}
    build(data, a.out)
