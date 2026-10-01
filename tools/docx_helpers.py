# -*- coding: utf-8 -*-
"""Вспомогательные функции оформления отчёта строго по требованиям нормоконтроля (ГОСТ и ЕСКД)."""
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

FONT = "Times New Roman"
SIZE_MAIN = Pt(14)
SIZE_TABLE = Pt(14)
SIZE_CODE = Pt(10)

COUNTERS = {"табл": 0, "рис": 0, "форм": 0}
SHIFR = "09.02.07.ИС.Р-25-29.2/70.14.ПЗ"


def enable_auto_update_fields(doc):
    """Включает автоматическое обновление полей (TOC, PAGE, NUMPAGES) при открытии в MS Word."""
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is None:
        uf = OxmlElement("w:updateFields")
        uf.set(qn("w:val"), "true")
        settings.append(uf)
    # Отключение автоматического переноса слов на уровне документа (п. 1.3 нормоконтроля)
    if settings.find(qn("w:autoHyphenation")) is None:
        ah = OxmlElement("w:autoHyphenation")
        ah.set(qn("w:val"), "false")
        settings.append(ah)


def _set_run_font(run, font_name=FONT, size=SIZE_MAIN, bold=False, italic=False):
    run.font.name = font_name
    run.font.size = size
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = RGBColor(0, 0, 0)
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    rf.set(qn("w:ascii"), font_name)
    rf.set(qn("w:hAnsi"), font_name)
    rf.set(qn("w:eastAsia"), font_name)
    rf.set(qn("w:cs"), font_name)


def set_base_style(doc):
    enable_auto_update_fields(doc)
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = SIZE_MAIN
    st.font.color.rgb = RGBColor(0, 0, 0)
    rpr = st.element.get_or_add_rPr()
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    rf.set(qn("w:ascii"), FONT)
    rf.set(qn("w:hAnsi"), FONT)
    rf.set(qn("w:eastAsia"), FONT)
    rf.set(qn("w:cs"), FONT)
    pf = st.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.first_line_indent = Cm(1.25)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    suppress = OxmlElement("w:suppressAutoHyphens")
    st.element.get_or_add_pPr().append(suppress)


def set_margins(section, left=2.5, right=1.5, top=1.5, bottom=2.5, footer_dist=0.5):
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(left)
    section.right_margin = Cm(right)
    section.top_margin = Cm(top)
    section.bottom_margin = Cm(bottom)
    section.footer_distance = Cm(footer_dist)
    section.header_distance = Cm(0.5)


def add_page_borders(section, sz=12):
    """Рамка ЕСКД по ГОСТ 2.104 (п. 1.12 нормоконтроля: отступ слева 20 мм, с остальных сторон 5 мм)."""
    sectPr = section._sectPr
    old = sectPr.find(qn("w:pgBorders"))
    if old is not None:
        sectPr.remove(old)
    b = OxmlElement("w:pgBorders")
    b.set(qn("w:offsetFrom"), "page")
    spaces = {"top": "14", "bottom": "14", "right": "14", "left": "28"}
    for edge, sp in spaces.items():
        e = OxmlElement("w:" + edge)
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), str(sz))
        e.set(qn("w:space"), sp)
        e.set(qn("w:color"), "000000")
        b.append(e)
    sectPr.append(b)


def add_field(paragraph, instr, default_val="3", size=None, bold=False, font="Arial"):
    run = paragraph.add_run()
    _set_run_font(run, font_name=font, size=size or SIZE_MAIN, bold=bold)
    fc1 = OxmlElement("w:fldChar")
    fc1.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = " %s " % instr
    fc2 = OxmlElement("w:fldChar")
    fc2.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t")
    t.text = str(default_val)
    fc3 = OxmlElement("w:fldChar")
    fc3.set(qn("w:fldCharType"), "end")
    for el in (fc1, it, fc2, t, fc3):
        run._r.append(el)
    return run


def _set_outline_level(paragraph, level: int):
    pPr = paragraph._p.get_or_add_pPr()
    ol = pPr.find(qn("w:outlineLvl"))
    if ol is None:
        ol = OxmlElement("w:outlineLvl")
        pPr.append(ol)
    ol.set(qn("w:val"), str(level))


# ---------------------------------------------------------------------------
# Штампы (основные надписи ГОСТ 2.104 Форма 2 и Форма 2а, п. 1.12 нормоконтроля)
# ---------------------------------------------------------------------------
def _cell_text(cell, text, size=Pt(8.5), font="Arial", bold=False, align="center"):
    cell.text = ""
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.alignment = {"left": WD_ALIGN_PARAGRAPH.LEFT,
                   "center": WD_ALIGN_PARAGRAPH.CENTER,
                   "right": WD_ALIGN_PARAGRAPH.RIGHT}[align]
    if text:
        r = p.add_run(text)
        _set_run_font(r, font_name=font, size=size, bold=bold)
    return p


def _zero_cell_margins(table):
    tblPr = table._tbl.tblPr
    tblCellMar = OxmlElement("w:tblCellMar")
    for side, val in (("top", "15"), ("bottom", "15"), ("left", "40"), ("right", "40")):
        node = OxmlElement("w:" + side)
        node.set(qn("w:w"), val)
        node.set(qn("w:type"), "dxa")
        tblCellMar.append(node)
    tblPr.append(tblCellMar)


def add_stamp_15(doc, section, shifr=SHIFR):
    """
    Основная надпись для листов основного текста и приложений (ГОСТ 2.104 Форма 2а, высота 15 мм,
    ширина 170 мм, 3 строки по 5 мм, 7 столбцов: 7 | 10 | 23 | 15 | 10 | 95 | 10 мм — п. 1.12.1).
    """
    footer = section.footer
    footer.is_linked_to_previous = False
    for p in list(footer.paragraphs):
        p._element.getparent().remove(p._element)
    for tbl in list(footer.tables):
        tbl._element.getparent().remove(tbl._element)

    t = doc.add_table(rows=3, cols=7)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _zero_cell_margins(t)

    widths = [Cm(0.7), Cm(1.0), Cm(2.3), Cm(1.5), Cm(1.0), Cm(9.5), Cm(1.0)]
    for r in t.rows:
        r.height = Cm(0.5)
        r.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        for i, w in enumerate(widths):
            r.cells[i].width = w

    # Объединение центральной графы шифра (столбец 5, строки 0..2)
    c_shifr = t.cell(0, 5).merge(t.cell(2, 5))
    c_shifr.width = Cm(9.5)
    _cell_text(c_shifr, shifr, size=Pt(14), font="Arial", bold=False, align="center")

    # Правая графа «Лист» (строка 0) и номер страницы (строки 1..2)
    _cell_text(t.cell(0, 6), "Лист", size=Pt(8.5), font="Arial", align="center")
    c_page = t.cell(1, 6).merge(t.cell(2, 6))
    c_page.width = Cm(1.0)
    p_page = _cell_text(c_page, "", size=Pt(12), font="Arial", align="center")
    add_field(p_page, "PAGE", default_val="3", size=Pt(12), font="Arial")

    # Нижняя строка левого блока: Изм. | Лист | № докум. | Подп. | Дата
    labels = ["Изм.", "Лист", "№ докум.", "Подп.", "Дата"]
    for j in range(5):
        _cell_text(t.cell(0, j), "", size=Pt(8), font="Arial")
        _cell_text(t.cell(1, j), "", size=Pt(8), font="Arial")
        _cell_text(t.cell(2, j), labels[j], size=Pt(8), font="Arial", align="center")

    anchor = footer.add_paragraph()
    anchor.paragraph_format.space_before = Pt(0)
    anchor.paragraph_format.space_after = Pt(0)
    anchor.paragraph_format.first_line_indent = Cm(0)
    anchor._p.addnext(t._tbl)
    return footer


def add_stamp_40(doc, section, shifr=SHIFR,
                 title="Разработка и администрирование базы данных «Поликлиника»",
                 total_pages=34):
    """
    Основная надпись для листа СОДЕРЖАНИЕ (ГОСТ 2.104 Форма 2, высота 40 мм,
    8 строк по 5 мм, 9 столбцов: 7 | 10 | 23 | 15 | 10 | 55 | 15 | 15 | 20 мм — п. 1.12.2).
    """
    footer = section.footer
    footer.is_linked_to_previous = False
    for p in list(footer.paragraphs):
        p._element.getparent().remove(p._element)
    for tbl in list(footer.tables):
        tbl._element.getparent().remove(tbl._element)

    t = doc.add_table(rows=8, cols=9)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    _zero_cell_margins(t)

    widths = [Cm(0.7), Cm(1.0), Cm(2.3), Cm(1.5), Cm(1.0), Cm(5.5), Cm(1.5), Cm(1.5), Cm(2.0)]
    for r in t.rows:
        r.height = Cm(0.5)
        r.height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
        for i, w in enumerate(widths):
            r.cells[i].width = w

    # Строка 2: Изм. | Лист | № докум. | Подпись | Дата
    hdr_left = ["Изм.", "Лист", "№ докум.", "Подпись", "Дата"]
    for j, h in enumerate(hdr_left):
        _cell_text(t.cell(2, j), h, size=Pt(8), font="Arial", align="center")

    # Строки 3..7 левого блока: должности и фамилии (п. 1.12.2 — Arial 9 или 8)
    left_rows = [
        (3, "Разраб.", "Сагадиев А.Р."),
        (4, "Провер.", "Габитова А.И."),
        (5, "Реценз.", ""),
        (6, "Н. Контр.", "Габитова А.И."),
        (7, "Утверд.", ""),
    ]
    for r_idx, role_lbl, person in left_rows:
        c_role = t.cell(r_idx, 0).merge(t.cell(r_idx, 1))
        c_role.width = Cm(1.7)
        _cell_text(c_role, role_lbl, size=Pt(8.5), font="Arial", align="left")
        _cell_text(t.cell(r_idx, 2), person, size=Pt(8.5), font="Arial", align="left")
        _cell_text(t.cell(r_idx, 3), "", size=Pt(8.5), font="Arial")
        _cell_text(t.cell(r_idx, 4), "", size=Pt(8.5), font="Arial")

    # Верхнее правое поле шифра (строки 0..2, столбцы 5..8) — Arial 14
    c_shifr = t.cell(0, 5).merge(t.cell(2, 8))
    c_shifr.width = Cm(10.5)
    _cell_text(c_shifr, shifr, size=Pt(14), font="Arial", align="center")

    # Среднее поле наименования темы (строки 3..7, столбец 5) — Arial 9.5
    c_title = t.cell(3, 5).merge(t.cell(7, 5))
    c_title.width = Cm(5.5)
    _cell_text(c_title, title, size=Pt(9.5), font="Arial", align="center")

    # Поля «Лит.», «Лист», «Листов» (строка 3, столбцы 6..8)
    _cell_text(t.cell(3, 6), "Лит.", size=Pt(8.5), font="Arial", align="center")
    _cell_text(t.cell(3, 7), "Лист", size=Pt(8.5), font="Arial", align="center")
    _cell_text(t.cell(3, 8), "Листов", size=Pt(8.5), font="Arial", align="center")

    # Значения «Лит.», «Лист» (2), «Листов» (NUMPAGES) — Arial 12 (п. 1.12.2)
    _cell_text(t.cell(4, 6), "У", size=Pt(11), font="Arial", align="center")
    p_pg = _cell_text(t.cell(4, 7), "", size=Pt(12), font="Arial", align="center")
    add_field(p_pg, "PAGE", default_val="2", size=Pt(12), font="Arial")
    p_tot = _cell_text(t.cell(4, 8), "", size=Pt(12), font="Arial", align="center")
    add_field(p_tot, "NUMPAGES", default_val=str(total_pages), size=Pt(12), font="Arial")

    # Нижнее правое поле организации: ГБПОУ УГКТИД (строки 5..7, столбцы 6..8) — Arial 12
    c_org = t.cell(5, 6).merge(t.cell(7, 8))
    c_org.width = Cm(5.0)
    _cell_text(c_org, "ГБПОУ УГКТИД", size=Pt(12), font="Arial", align="center")

    anchor = footer.add_paragraph()
    anchor.paragraph_format.space_before = Pt(0)
    anchor.paragraph_format.space_after = Pt(0)
    anchor.paragraph_format.first_line_indent = Cm(0)
    anchor._p.addnext(t._tbl)
    return footer


# ---------------------------------------------------------------------------
# Элементы основного текста
# ---------------------------------------------------------------------------
def para(doc, text="", align="justify", size=None, bold=False, italic=False,
         first=Cm(1.25), space_before=0, space_after=0, line=1.5, style=None,
         font_name=FONT):
    p = doc.add_paragraph(style=style)
    pf = p.paragraph_format
    pf.first_line_indent = first
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if line == 1.5:
        pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    elif line == 1.0:
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    else:
        pf.line_spacing = line
    p.alignment = {"justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
                   "center": WD_ALIGN_PARAGRAPH.CENTER,
                   "left": WD_ALIGN_PARAGRAPH.LEFT,
                   "right": WD_ALIGN_PARAGRAPH.RIGHT}[align]
    if text:
        r = p.add_run(text)
        _set_run_font(r, font_name=font_name, size=size or SIZE_MAIN, bold=bold, italic=italic)
    return p


def heading1(doc, text, has_sub=False, page_break_before=True):
    """
    Заголовок раздела (ГЛАВЫ) строго по п. 1.2.1 нормоконтроля:
    Times New Roman 14, начертание обычное, ВСЕ ПРОПИСНЫЕ, по центру (без отступа),
    междустрочный интервал одинарный, интервал после 12 пт (если дальше нет параграфа)
    или 0 пт (если дальше сразу идёт параграф — Рисунок 1 и Рисунок 2 нормоконтроля).
    """
    p = doc.add_paragraph()
    p.paragraph_format.page_break_before = page_break_before
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.left_indent = Cm(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0 if has_sub else 12)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.keep_with_next = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_outline_level(p, 0)
    r = p.add_run(text.rstrip(".").upper())
    _set_run_font(r, font_name=FONT, size=Pt(14), bold=False)
    return p


def heading2(doc, text):
    """
    Заголовок подраздела (параграфа) строго по п. 1.2.2 нормоконтроля:
    Times New Roman 14, начертание обычное, строчные буквы (кроме 1-ой прописной),
    по центру (без отступа), интервалы перед и после 12 пт, междустрочный одинарный.
    """
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.left_indent = Cm(0)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.keep_with_next = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _set_outline_level(p, 1)
    r = p.add_run(text.rstrip("."))
    _set_run_font(r, font_name=FONT, size=Pt(14), bold=False)
    return p


def heading3(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.left_indent = Cm(0)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.paragraph_format.keep_with_next = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text.rstrip("."))
    _set_run_font(r, font_name=FONT, size=Pt(14), bold=False)
    return p


def bullet(doc, text):
    """
    Маркированный список по п. 1.3 нормоконтроля (стр. 7):
    маркер «⎯», текст после маркера начинается с отступом 2 см через табуляцию.
    """
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(2.0)
    pf.first_line_indent = Cm(-0.75)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.tab_stops.add_tab_stop(Cm(2.0), WD_TAB_ALIGNMENT.LEFT)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run("⎯\t" + text)
    _set_run_font(r, font_name=FONT, size=SIZE_MAIN)
    return p


def numbered_item(doc, num: int, text: str):
    """
    Нумерованное перечисление по п. 1.3 нормоконтроля (стр. 7):
    1), 2), 3) с отступом текста 2 см через табуляцию.
    """
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.left_indent = Cm(2.0)
    pf.first_line_indent = Cm(-0.75)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.tab_stops.add_tab_stop(Cm(2.0), WD_TAB_ALIGNMENT.LEFT)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = p.add_run(f"{num})\t{text}")
    _set_run_font(r, font_name=FONT, size=SIZE_MAIN)
    return p


def table_caption(doc, text):
    """
    Заголовок таблицы строго по п. 1.4.1 нормоконтроля (стр. 11–12):
    «Таблица N – Название», Times New Roman 14, обычное, выравнивание по левому краю,
    без абзацного отступа, интервал перед 12 пт, после 6 пт, междустрочный одинарный, без точки.
    """
    COUNTERS["табл"] += 1
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.left_indent = Cm(0)
    pf.space_before = Pt(12)
    pf.space_after = Pt(6)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.keep_with_next = True
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run("Таблица %d – %s" % (COUNTERS["табл"], text.rstrip(".")))
    _set_run_font(r, font_name=FONT, size=Pt(14), bold=False)
    return p, COUNTERS["табл"]


def make_table(doc, headers, rows, widths=None, size=Pt(14), header_bold=False,
               align_first_col="left", add_col_numbers=True):
    """
    Таблица строго по п. 1.4.2 нормоконтроля (стр. 12–14):
    Times New Roman 14 пт (или 12 пт для многографных справочных таблиц),
    начертание обычное, интервалы перед и после 0 пт, междустрочный одинарный,
    строка нумерации граф 1..N под головкой таблицы.
    """
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr = t.rows[0]
    for i, h in enumerate(headers):
        c = hdr.cells[i]
        c.text = ""
        c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        p = c.paragraphs[0]
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(str(h).rstrip("."))
        _set_run_font(r, font_name=FONT, size=size, bold=header_bold)

    trPr = hdr._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))

    if add_col_numbers:
        num_row = t.add_row()
        num_trPr = num_row._tr.get_or_add_trPr()
        num_trPr.append(OxmlElement("w:tblHeader"))
        for i in range(len(headers)):
            c = num_row.cells[i]
            c.text = ""
            c.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = c.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(str(i + 1))
            _set_run_font(r, font_name=FONT, size=size, bold=False)

    for row_vals in rows:
        row_obj = t.add_row()
        # запрет разрыва строки таблицы между страницами
        r_trPr = row_obj._tr.get_or_add_trPr()
        r_trPr.append(OxmlElement("w:cantSplit"))
        cells = row_obj.cells
        for i, v in enumerate(row_vals):
            cells[i].text = ""
            cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cells[i].paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            if i == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if align_first_col == "left" else WD_ALIGN_PARAGRAPH.CENTER
            else:
                # Короткие числовые/кодовые значения по центру, длинный текст — по левому краю
                s_val = str(v)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if len(s_val) <= 18 else WD_ALIGN_PARAGRAPH.LEFT
            r = p.add_run(str(v))
            _set_run_font(r, font_name=FONT, size=size, bold=False)

    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    return t


def figure(doc, path, caption, width_cm=15.5, is_chart=False):
    """
    Иллюстрация и подпись строго по п. 1.6.1, 1.6.2 и 1.7 нормоконтроля (стр. 17–19):
    - сам рисунок: по центру, интервалы перед и после 6 пт (для графика/диаграммы п. 1.7: 12 пт, одинарный);
    - подпись: «Рисунок N – Название», Times New Roman 14, обычное, по центру,
      интервалы до и после 6 пт, междустрочный полуторный (1,5), без точки в конце.
    """
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.left_indent = Cm(0)
    p.paragraph_format.space_before = Pt(12 if is_chart else 6)
    p.paragraph_format.space_after = Pt(12 if is_chart else 6)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE if is_chart else WD_LINE_SPACING.ONE_POINT_FIVE
    p.paragraph_format.keep_with_next = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(path, width=Cm(width_cm))

    COUNTERS["рис"] += 1
    cp = doc.add_paragraph()
    cpf = cp.paragraph_format
    cpf.first_line_indent = Cm(0)
    cpf.left_indent = Cm(0)
    cpf.space_before = Pt(6)
    cpf.space_after = Pt(6)
    cpf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = cp.add_run("Рисунок %d – %s" % (COUNTERS["рис"], caption.rstrip(".")))
    _set_run_font(r, font_name=FONT, size=Pt(14), bold=False)
    return COUNTERS["рис"]


def formula(doc, text, number=None, size=Pt(14)):
    """
    Текст формулы строго по п. 1.5.2 нормоконтроля (стр. 15–16):
    Times New Roman 14, обычное, выравнивание по левому краю, отступ слева 3 см,
    интервалы перед и после 0 пт, междустрочный полуторный (1,5), номер формулы
    с правой стороны листа на уровне формулы в круглых скобках.
    """
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.left_indent = Cm(3.0)
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.tab_stops.add_tab_stop(Cm(17.0), WD_TAB_ALIGNMENT.RIGHT)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(text)
    _set_run_font(r, font_name=FONT, size=size, bold=False, italic=False)
    if number:
        rp = p.add_run("\t(%s)" % number)
        _set_run_font(rp, font_name=FONT, size=size, bold=False)
        COUNTERS["форм"] += 1
    return p


def formula_note(doc, items, size=Pt(14)):
    """
    Пояснения символов под формулой строго по п. 1.5.3 нормоконтроля (стр. 15–16):
    Times New Roman 14, обычное, выравнивание по левому краю,
    отступ слева 1-го коэффициента («где N – ...») 0 см, остальных — 0,8 см,
    интервалы перед и после 0 пт, междустрочный полуторный (1,5), без двоеточия после «где».
    """
    for i, it in enumerate(items):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.first_line_indent = Cm(0)
        pf.left_indent = Cm(0.0 if i == 0 else 0.8)
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(("где " if i == 0 else "") + it)
        _set_run_font(r, font_name=FONT, size=size, bold=False)
    return p


def code_block(doc, lines, size=Pt(9.5), font="Courier New"):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.left_indent = Cm(0.5)
    pf.space_before = Pt(3)
    pf.space_after = Pt(3)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for i, ln in enumerate(lines):
        if i:
            p.add_run().add_break()
        r = p.add_run(ln)
        _set_run_font(r, font_name=font, size=size)
    return p


def page_break(doc):
    p = doc.add_paragraph()
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.add_run().add_break(WD_BREAK.PAGE)
    return p
