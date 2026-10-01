# -*- coding: utf-8 -*-
"""Генерация иллюстраций для отчёта (PNG, 200 dpi)."""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Rectangle

from db_schema import build_data

OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "Рисунки"))
os.makedirs(OUT, exist_ok=True)

plt.rcParams["font.family"] = "DejaVu Sans"
plt.rcParams["font.size"] = 8
plt.rcParams["axes.linewidth"] = 1.0

INK = "#000000"
GREY = "#4a4a4a"
LIGHT = "#f2f2f2"


def new_fig(w_cm=17.0, h_cm=11.5):
    fig, ax = plt.subplots(figsize=(w_cm / 2.54, h_cm / 2.54), dpi=200)
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.set_aspect("auto")
    ax.axis("off")
    fig.subplots_adjust(left=0.01, right=0.99, top=0.98, bottom=0.02)
    return fig, ax


def save(fig, name):
    path = os.path.join(OUT, name)
    fig.savefig(path, dpi=200, facecolor="white", bbox_inches="tight", pad_inches=0.08)
    plt.close(fig)
    print("  рисунок:", name)


# --------------------------------------------------------------------------
# Рис. 1 — функциональная модель (IDEF0-контекст + процессы)
# --------------------------------------------------------------------------
def fig_functional():
    fig, ax = new_fig(17.0, 12.0)

    def block(x, y, w, h, title, lines, fc="white"):
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.3,rounding_size=0.8",
                                    linewidth=1.2, edgecolor=INK, facecolor=fc))
        ax.text(x + w / 2, y + h - 2.4, title, ha="center", va="center",
                fontsize=8.2, fontweight="bold")
        ax.plot([x + 0.5, x + w - 0.5], [y + h - 4.4, y + h - 4.4], color=INK, lw=0.8)
        for i, ln in enumerate(lines):
            ax.text(x + 0.9, y + h - 6.8 - i * 2.8, ln, ha="left", va="center", fontsize=6.9)

    def arrow(x1, y1, x2, y2, label=""):
        ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>",
                                     mutation_scale=10, linewidth=1.0, color=GREY))
        if label:
            ax.text((x1 + x2) / 2, (y1 + y2) / 2 + 2.2, label, ha="center", va="center",
                    fontsize=6.8, color=INK,
                    bbox=dict(boxstyle="round,pad=0.15", fc="white", ec="none"))

    ax.text(50, 97, "Контекстная диаграмма A-0", ha="center", va="center",
            fontsize=9.5, fontweight="bold")
    block(28, 66, 44, 20, "Администрирование БД «Поликлиника»",
          ["СУБД MS Access (Jet / ACE)", "Исполнитель: Сагадиев А. Р., гр. Р-25-29",
           "Вариант 14"], "#e8e8e8")

    arrow(3, 76, 28, 76, "Заявка пациента\n(запись на приём)")
    arrow(72, 76, 97, 76, "Талон, диагноз,\nназначения, отчёты")
    ax.text(50, 92.5, "Управление: регламент поликлиники, расписание врачей, МКБ-10",
            ha="center", va="center", fontsize=6.9, color=GREY)
    ax.add_patch(FancyArrowPatch((50, 90.5), (50, 86), arrowstyle="-|>", mutation_scale=9, color=GREY))
    ax.text(50, 59.5, "Механизмы: СУБД MS Access, SQL, VBA, макросы",
            ha="center", va="center", fontsize=6.9, color=GREY)
    ax.add_patch(FancyArrowPatch((50, 61.5), (50, 66), arrowstyle="-|>", mutation_scale=9, color=GREY))

    ax.text(50, 53.5, "Декомпозиция A0: функциональные блоки автоматизированной системы",
            ha="center", va="center", fontsize=9.2, fontweight="bold")

    block(2, 31, 22, 19, "A1. Регистратура",
          ["Ввод пациентов", "Сетка расписания", "Контроль времени", "Выдача талона"])
    block(27, 31, 22, 19, "A2. Приём врача",
          ["Карта приёма", "Жалобы и анамнез", "Диагнозы МКБ-10", "Статус приёма"])
    block(52, 31, 22, 19, "A3. Назначения",
          ["Услуги и лечение", "Дозировки, курс", "Расчёт стоимости", "Отметка выполнения"])
    block(77, 31, 21, 19, "A4. Администрир.",
          ["Загрузка врачей", "История болезни", "Роли и аудит", "Резерв. копии"])

    arrow(24, 40.5, 27, 40.5)
    arrow(49, 40.5, 52, 40.5)
    arrow(74, 40.5, 77, 40.5)

    block(2, 4, 47, 21, "База данных «Поликлиника» (13 таблиц, 3НФ)",
          ["Специальности, Врачи, Пациенты, Услуги, МКБ,",
           "РасписаниеПриёма, ЗаписиНаПриём, Диагнозы,",
           "Назначения, Роли, Пользователи, АрхивЗаписей,",
           "ЖурналСобытий (ссылочная целостность)"], LIGHT)
    block(51, 4, 47, 21, "Объекты приложения MS Access",
          ["Запросы (16): выборка, с параметрами, итоговые,",
           "на добавление, обновление и удаление;",
           "Формы (7): главная с подчинёнными, ввод, ленточная;",
           "Отчёты (4); макросы (6); модули VBA (3)"], LIGHT)

    arrow(13, 31, 13, 25)
    arrow(63, 31, 63, 25)
    save(fig, "рис01_функциональная_модель.png")


# --------------------------------------------------------------------------
# Рис. 2 — ER-модель
# --------------------------------------------------------------------------
def fig_er():
    fig, ax = new_fig(17.0, 12.5)

    def ent(x, y, w, h, title, attrs, fc="white"):
        ax.add_patch(Rectangle((x, y), w, h, linewidth=1.1, edgecolor=INK, facecolor=fc))
        ax.add_patch(Rectangle((x, y + h - 4.4), w, 4.4, linewidth=1.1,
                               edgecolor=INK, facecolor="#dcdcdc"))
        ax.text(x + w / 2, y + h - 2.2, title, ha="center", va="center",
                fontsize=8.0, fontweight="bold")
        for i, a in enumerate(attrs):
            mark = "PK " if i == 0 else "     "
            ax.text(x + 0.7, y + h - 6.8 - i * 2.7, mark + a, ha="left", va="center", fontsize=6.6)

    def rel(x1, y1, x2, y2, lbl="1:M"):
        ax.plot([x1, x2], [y1, y2], color=INK, lw=0.9)
        ax.text((x1 + x2) / 2, (y1 + y2) / 2 + 1.4, lbl, ha="center", va="center",
                fontsize=6.4, fontweight="bold",
                bbox=dict(boxstyle="round,pad=0.12", fc="white", ec=INK, lw=0.5))

    ax.text(50, 97, "ER-модель предметной области «Поликлиника»",
            ha="center", va="center", fontsize=9.5, fontweight="bold")

    ent(2, 66, 28, 25, "СПЕЦИАЛЬНОСТИ",
        ["КодСпециальности", "НазваниеСпециальности", "Описание", "ПорядокОтображения"])
    ent(2, 35, 28, 24, "ВРАЧИ",
        ["КодВрача", "Фамилия, Имя, Отчество", "КодСпециальности (FK)", "Кабинет, Телефон",
         "ДатаПриёмаНаРаботу, Категория"])
    ent(2, 4, 28, 24, "ПАЦИЕНТЫ",
        ["КодПациента", "Фамилия, Имя, Отчество", "ДатаРождения, Пол", "Адрес, Телефон",
         "НомерПолиса, Участок"])

    ent(36, 66, 28, 25, "УСЛУГИ",
        ["КодУслуги", "КодСпециальности (FK)", "НазваниеУслуги", "Стоимость", "ДлительностьМинут"])
    ent(36, 35, 28, 24, "РАСПИСАНИЕ ПРИЁМА",
        ["КодРасписания", "КодВрача (FK)", "ДатаПриёма", "ВремяНачала, ВремяОкончания",
         "Кабинет, ПризнакПриёма"])
    ent(36, 4, 28, 24, "ЗАПИСИ НА ПРИЁМ",
        ["КодЗаписи", "НомерТалона (UNIQUE)", "КодПациента (FK), КодВрача (FK)",
         "КодРасписания (FK)", "ДатаПриёма, ВремяПриёма", "СтатусЗаписи, Жалобы"])

    ent(70, 66, 28, 25, "МКБ-10",
        ["КодМКБ", "Наименование", "КлассЗаболеваний"])
    ent(70, 35, 28, 24, "ДИАГНОЗЫ",
        ["КодДиагноза", "КодЗаписи (FK)", "КодМКБ (FK)", "ТипДиагноза", "ДатаУстановки"])
    ent(70, 4, 28, 24, "НАЗНАЧЕНИЯ",
        ["КодНазначения", "КодЗаписи (FK)", "КодУслуги (FK)", "Количество", "Дозировка, Кратность"])

    rel(16, 66, 16, 59, "1:M")
    rel(30, 78, 36, 78, "1:M")
    rel(30, 47, 36, 47, "1:M")
    rel(30, 16, 36, 16, "1:M")
    rel(30, 39, 36, 24, "1:M")
    rel(50, 35, 50, 28, "1:M")
    rel(64, 22, 70, 44, "1:M")
    rel(64, 14, 70, 14, "1:M")
    rel(84, 66, 84, 59, "1:M")
    rel(64, 72, 70, 24, "1:M")
    save(fig, "рис02_ER_модель.png")


# --------------------------------------------------------------------------
# Рис. 3 — схема связей (окно «Схема данных» MS Access)
# --------------------------------------------------------------------------
def fig_schema():
    fig, ax = new_fig(17.0, 13.5)
    w, h = 29, 17.5
    boxes = {
        "Специальности": (2, 76),
        "Врачи": (2, 53),
        "Пациенты": (2, 30),
        "Роли": (2, 7),
        "Услуги": (35.5, 76),
        "РасписаниеПриёма": (35.5, 53),
        "ЗаписиНаПриём": (35.5, 30),
        "Пользователи": (35.5, 7),
        "МКБ": (69, 76),
        "Диагнозы": (69, 53),
        "Назначения": (69, 30),
        "ЖурналСобытий": (69, 15),
        "АрхивЗаписей": (69, 2),
    }

    def tbl(name, x, y, fields, h_box=None):
        hb = h_box or h
        ax.add_patch(Rectangle((x, y), w, hb, linewidth=1.0, edgecolor=INK, facecolor="white"))
        ax.add_patch(Rectangle((x, y + hb - 4.0), w, 4.0, linewidth=1.0,
                               edgecolor=INK, facecolor="#d6d6d6"))
        ax.text(x + w / 2, y + hb - 2.0, name, ha="center", va="center",
                fontsize=7.6, fontweight="bold")
        for i, f in enumerate(fields[:6]):
            key = "PK " if i == 0 else "     "
            ax.text(x + 0.7, y + hb - 6.0 - i * 2.0, key + f, ha="left", va="center", fontsize=6.2)

    tbl("Специальности", 2, 76, ["КодСпециальности", "НазваниеСпециальности", "Описание", "ПорядокОтображения"])
    tbl("Врачи", 2, 53, ["КодВрача", "Фамилия", "Имя, Отчество", "КодСпециальности", "Кабинет, Телефон", "Категория"])
    tbl("Пациенты", 2, 30, ["КодПациента", "Фамилия, Имя", "ДатаРождения, Пол", "НомерПолиса", "Телефон, Адрес", "Участок"])
    tbl("Роли", 2, 7, ["КодРоли", "НазваниеРоли", "Описание", "УровеньДоступа"], h_box=15)
    tbl("Услуги", 35.5, 76, ["КодУслуги", "КодСпециальности", "НазваниеУслуги", "Стоимость", "ДлительностьМинут"])
    tbl("РасписаниеПриёма", 35.5, 53, ["КодРасписания", "КодВрача", "ДатаПриёма", "ВремяНачала", "ВремяОкончания", "Кабинет"])
    tbl("ЗаписиНаПриём", 35.5, 30, ["КодЗаписи", "НомерТалона", "КодПациента", "КодВрача", "КодРасписания", "ДатаПриёма, Время"])
    tbl("Пользователи", 35.5, 7, ["КодПользователя", "Логин, Пароль", "ФИО", "КодРоли, КодВрача", "Активен"], h_box=15)
    tbl("МКБ", 69, 76, ["КодМКБ", "Наименование", "КлассЗаболеваний"])
    tbl("Диагнозы", 69, 53, ["КодДиагноза", "КодЗаписи", "КодМКБ", "ТипДиагноза", "ДатаУстановки"])
    tbl("Назначения", 69, 30, ["КодНазначения", "КодЗаписи", "КодУслуги", "Количество", "Дозировка, Кратность"])
    tbl("ЖурналСобытий", 69, 15, ["КодСобытия", "ДатаВремя, Логин", "Действие, Результат"], h_box=11)
    tbl("АрхивЗаписей", 69, 2, ["КодАрхива", "КодЗаписи, Талон", "ДатаПриёма, Статус"], h_box=11)

    def link(a, b, side="h", label="1:M"):
        xa, ya = boxes[a]
        xb, yb = boxes[b]
        if side == "h":
            x1, y1 = xa + w, ya + 8
            x2, y2 = xb, yb + 8
        else:
            x1, y1 = xa + w / 2, ya
            x2, y2 = xb + w / 2, yb + h
        ax.plot([x1, x2], [y1, y2], color=INK, lw=0.85)
        ax.text((x1 + x2) / 2, (y1 + y2) / 2 + 1.1, label, fontsize=5.8, ha="center",
                bbox=dict(boxstyle="round,pad=0.08", fc="white", ec="none"))

    link("Специальности", "Врачи", "v")
    link("Специальности", "Услуги", "h")
    link("Врачи", "РасписаниеПриёма", "h")
    link("Пациенты", "ЗаписиНаПриём", "h")
    link("РасписаниеПриёма", "ЗаписиНаПриём", "v")
    link("ЗаписиНаПриём", "Диагнозы", "h")
    link("ЗаписиНаПриём", "Назначения", "h")
    link("МКБ", "Диагнозы", "v")
    link("Роли", "Пользователи", "h")

    ax.text(50, 97.2, "Схема данных MS Access (13 таблиц, 12 связей, целостность данных)",
            ha="center", va="center", fontsize=9.2, fontweight="bold")
    save(fig, "рис03_схема_связей.png")


# --------------------------------------------------------------------------
# Рис. 4 — нормализация
# --------------------------------------------------------------------------
def fig_normal():
    fig, ax = new_fig(17.0, 11.5)
    ax.text(50, 96.5, "Нормализация отношений базы данных «Поликлиника»: 1НФ -> 2НФ -> 3НФ",
            ha="center", va="center", fontsize=9.2, fontweight="bold")

    def table(x, y, w, title, header, rows, fc="white", fs=6.3):
        row_h = 3.6
        h = 5.2 + row_h * (len(rows) + 1)
        ax.add_patch(Rectangle((x, y - h), w, h, linewidth=1.0, edgecolor=INK, facecolor=fc))
        ax.add_patch(Rectangle((x, y - 5.2), w, 5.2, linewidth=1.0, edgecolor=INK, facecolor="#dddddd"))
        ax.text(x + w / 2, y - 2.6, title, ha="center", va="center", fontsize=7.6, fontweight="bold")
        for j, cname in enumerate(header):
            cx = x + (w / len(header)) * (j + 0.5)
            ax.text(cx, y - 7.0, cname, ha="center", va="center", fontsize=fs, fontweight="bold")
            if j:
                ax.plot([x + (w / len(header)) * j, x + (w / len(header)) * j],
                        [y - 5.2, y - h], color=INK, lw=0.6)
        ax.plot([x, x + w], [y - 8.8, y - 8.8], color=INK, lw=0.6)
        for i, r in enumerate(rows):
            for j, cell in enumerate(r):
                cx = x + (w / len(header)) * (j + 0.5)
                ax.text(cx, y - 10.8 - i * row_h, cell, ha="center", va="center", fontsize=fs)
        return h

    table(2, 91, 46, "Исходное отношение (не нормализовано)",
          ["Талон", "Пациент", "Врач", "Спец.", "Диагнозы", "Услуги"],
          [["Т-001", "Закиров В.", "Валиева А.", "Терапевт", "J06.9, I10", "ЭКГ, Осмотр"],
           ["Т-002", "Юсупова А.", "Гареев Р.", "Терапевт", "E11.9", "Глюкозотест"],
           ["Т-003", "Сафин М.", "Хайруллина", "Кардиолог", "I10", "ЭКГ"]],
          "#f8f2f2", fs=5.8)

    table(52, 91, 46, "1НФ: атомарность атрибутов (без групп)",
          ["Талон", "Пациент", "Врач", "Спец.", "Диагноз", "Услуга"],
          [["Т-001", "Закиров В.", "Валиева А.", "Терапевт", "J06.9", "Осмотр"],
           ["Т-001", "Закиров В.", "Валиева А.", "Терапевт", "I10", "ЭКГ"],
           ["Т-002", "Юсупова А.", "Гареев Р.", "Терапевт", "E11.9", "Глюкозотест"]],
          "#f8f6f0", fs=5.8)

    table(2, 52, 46, "2НФ: устранение частичных зависимостей",
          ["КодЗаписи", "Талон", "КодПац.", "КодВрача", "Дата", "Время"],
          [["1", "Т-001", "1", "1", "01.09.2026", "09:00"],
           ["2", "Т-002", "2", "2", "01.09.2026", "09:30"],
           ["3", "Т-003", "3", "3", "02.09.2026", "10:00"]],
          "#f0f5f8")

    table(52, 52, 46, "3НФ: устранение транзитивных зависимостей",
          ["КодДиаг.", "КодЗаписи", "КодМКБ", "ТипДиагноза"],
          [["1", "1", "J06.9", "Основной"],
           ["2", "1", "I10", "Сопутствующий"],
           ["3", "2", "E11.9", "Основной"]],
          "#f0f8f0")

    ax.add_patch(FancyArrowPatch((48.2, 80), (51.8, 80), arrowstyle="-|>", mutation_scale=11, color=INK))
    ax.add_patch(FancyArrowPatch((25, 68), (25, 54), arrowstyle="-|>", mutation_scale=11, color=INK))
    ax.add_patch(FancyArrowPatch((75, 68), (75, 54), arrowstyle="-|>", mutation_scale=11, color=INK))

    ax.text(50, 12,
            "Результат декомпозиции: 13 таблиц в третьей нормальной форме (3НФ).\n"
            "Каждый неключевой атрибут полностью зависит от первичного ключа и не зависит "
            "транзитивно от других неключевых атрибутов.",
            ha="center", va="center", fontsize=7.2,
            bbox=dict(boxstyle="round,pad=0.5", fc=LIGHT, ec=INK, lw=0.8))
    save(fig, "рис04_нормализация.png")


# --------------------------------------------------------------------------
# Рис. 5-8 — макеты форм и отчёта
# --------------------------------------------------------------------------
def mock_form(title, blocks, footer, filename, w_cm=16.5, h_cm=10.5):
    fig, ax = new_fig(w_cm, h_cm)
    ax.add_patch(Rectangle((2, 6), 96, 90, linewidth=1.2, edgecolor=INK, facecolor="#fbfbfb"))
    ax.add_patch(Rectangle((2, 87), 96, 9, linewidth=1.2, edgecolor=INK, facecolor="#d9d9d9"))
    ax.text(50, 91.5, title, ha="center", va="center", fontsize=9.2, fontweight="bold")

    for (x, y, w, h, label, kind) in blocks:
        if kind == "button":
            ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.2,rounding_size=0.5",
                                        linewidth=0.9, edgecolor=INK, facecolor="#e2e2e2"))
            ax.text(x + w / 2, y + h / 2, label, ha="center", va="center", fontsize=6.8, fontweight="bold")
        elif kind == "subform":
            ax.text(x, y + h + 1.8, label, ha="left", va="center", fontsize=7.0, fontweight="bold")
            ax.add_patch(Rectangle((x, y), w, h, linewidth=1.0, edgecolor=INK, facecolor="white"))
            ax.add_patch(Rectangle((x, y + h - 3.6), w, 3.6, linewidth=0.8,
                                   edgecolor=INK, facecolor="#e8e8e8"))
            for j in range(1, 5):
                cx = x + w * j / 5
                ax.plot([cx, cx], [y, y + h], color="#cccccc", lw=0.6)
            ax.text(x + 1.0, y + h - 1.8, "Услуга / Код МКБ   |   Кол-во / Тип   |   Цена / Дата   |   Сумма   |   Кратность",
                    ha="left", va="center", fontsize=6.3)
        elif kind == "grid_hdr":
            ax.add_patch(Rectangle((x, y), w, h, linewidth=0.8, edgecolor=INK, facecolor="#e4e4e4"))
            ax.text(x + w / 2, y + h / 2, label, ha="center", va="center", fontsize=6.4, fontweight="bold")
        elif kind == "grid_cell":
            ax.add_patch(Rectangle((x, y), w, h, linewidth=0.7, edgecolor="#888888", facecolor="white"))
            ax.text(x + 0.8, y + h / 2, label, ha="left", va="center", fontsize=6.2)
        else:
            ax.text(x, y + h + 1.6, label, ha="left", va="center", fontsize=6.6)
            ax.add_patch(Rectangle((x, y), w, h, linewidth=0.8, edgecolor=INK, facecolor="white"))
            if kind == "combo":
                ax.plot([x + w - 2.2, x + w - 2.2], [y, y + h], color=INK, lw=0.7)
                ax.text(x + w - 1.1, y + h / 2, "v", ha="center", va="center", fontsize=6.0)

    ax.plot([2, 98], [13, 13], color=INK, lw=0.8)
    ax.text(50, 9.5, footer, ha="center", va="center", fontsize=6.6, color=GREY)
    save(fig, filename)


def fig_forms():
    mock_form("Форма «Ф_ПриёмПациента» — КАРТА ПРИЁМА ПАЦИЕНТА",
              [(5, 74, 38, 5.5, "Пациент (Поле со списком):", "combo"),
               (48, 74, 22, 5.5, "Номер талона:", "text"),
               (74, 74, 20, 5.5, "Статус записи:", "combo"),
               (5, 62, 38, 5.5, "Врач (Поле со списком):", "combo"),
               (48, 62, 22, 5.5, "Дата приёма:", "text"),
               (74, 62, 20, 5.5, "Время приёма:", "text"),
               (5, 49, 89, 7.0, "Жалобы и анамнез:", "memo"),
               (5, 31, 89, 13.0, "Подчинённая форма «ПФ_Назначения» (связь по КодЗаписи):", "subform"),
               (5, 15, 89, 11.5, "Подчинённая форма «ПФ_Диагнозы» (связь по КодЗаписи):", "subform")],
              "Примечание формы: Стоимость приёма =DSum([Сумма])  |  [Обновить]  [Печать талона]  [Закрыть]",
              "рис05_макет_форма_приём.png")

    mock_form("Форма «Ф_ЗаписьНаПриём» — ЗАПИСЬ ПАЦИЕНТА НА ПРИЁМ (РЕГИСТРАТУРА)",
              [(6, 73, 28, 6.0, "Номер талона (авто):", "text"),
               (38, 73, 26, 6.0, "Дата приёма:", "text"),
               (68, 73, 26, 6.0, "Время приёма:", "combo"),
               (6, 59, 58, 6.0, "Пациент (ФИО, полис ОМС):", "combo"),
               (68, 59, 26, 6.0, "Статус записи:", "combo"),
               (6, 45, 58, 6.0, "Врач (ФИО, специальность, кабинет):", "combo"),
               (68, 45, 26, 6.0, "Свободные талоны", "button"),
               (6, 28, 88, 11.0, "Жалобы (первичный опрос регистратора):", "memo"),
               (6, 17, 26, 6.5, "Сохранить запись", "button"),
               (36, 17, 26, 6.5, "Новая запись", "button"),
               (66, 17, 28, 6.5, "Закрыть форму", "button")],
              "Контроль Form_BeforeUpdate: проверка ВремяСвободно(КодВрача, ДатаПриёма, ВремяПриёма)",
              "рис06_макет_форма_запись.png")

    cols = [("Пациент", 4, 24), ("Дата рожд.", 28, 12), ("Возр.", 40, 7), ("Пол", 47, 5),
            ("Группа", 52, 15), ("Полис ОМС", 67, 14), ("Телефон", 81, 13)]
    sample_rows = [
        ("Закиров В. С.", "07.01.1961", "65", "М", "Старшая", "5294 7334 8972", "8-962-521-56-14"),
        ("Миннибаева Ю. А.", "18.10.2013", "12", "Ж", "Детская", "5272 4164 9893", "8-963-948-75-51"),
        ("Соколов В. Д.", "12.04.1996", "30", "М", "Трудоспособная", "2260 8976 8238", "8-916-211-42-86"),
        ("Каримова А. Р.", "25.08.1984", "42", "Ж", "Трудоспособная", "4188 3012 7741", "8-917-405-19-22"),
    ]
    blocks = []
    for name, xx, ww in cols:
        blocks.append((xx, 76, ww, 6.5, name, "grid_hdr"))
    for r_i, r_vals in enumerate(sample_rows):
        yy = 68 - r_i * 8.5
        for (_, xx, ww), val in zip(cols, r_vals):
            blocks.append((xx, yy, ww, 6.5, val, "grid_cell"))
    blocks.append((40, 17, 26, 6.5, "Карта приёмов", "button"))
    blocks.append((70, 17, 24, 6.5, "Закрыть", "button"))
    mock_form("Форма «Ф_Пациенты» — КАРТОТЕКА ПАЦИЕНТОВ (ЛЕНТОЧНАЯ ФОРМА)",
              blocks,
              "Режим DefaultView = 1 (непрерывные формы), источник записей: Q04_ПациентыСВозрастом",
              "рис07_макет_форма_пациенты.png")

    r_cols = [("Специальность / Врач", 5, 34), ("Всего талонов", 39, 15),
              ("Принято", 54, 13), ("Неявка", 67, 13), ("Отменено", 80, 14)]
    r_blocks = []
    for name, xx, ww in r_cols:
        r_blocks.append((xx, 76, ww, 6.5, name, "grid_hdr"))
    r_rows = [
        ("Группа: Терапевт", "", "", "", ""),
        ("   Валиева А. Р.", "5", "3", "0", "0"),
        ("   Гареев Р. М.", "5", "2", "1", "0"),
        ("Итого по специальности Терапевт:", "10", "5", "1", "0"),
        ("Группа: Кардиолог", "", "", "", ""),
        ("   Хайруллина Д. И.", "5", "3", "0", "0"),
        ("ВСЕГО ПО ПОЛИКЛИНИКЕ:", "54", "28", "3", "4"),
    ]
    for r_i, r_vals in enumerate(r_rows):
        yy = 68.5 - r_i * 7.2
        is_hdr = r_vals[0].startswith(("Группа", "Итого", "ВСЕГО"))
        for (_, xx, ww), val in zip(r_cols, r_vals):
            r_blocks.append((xx, yy, ww, 6.2, val, "grid_hdr" if is_hdr else "grid_cell"))
    mock_form("Отчёт «О_ЗагрузкаВрачей» — ЗАГРУЗКА ВРАЧЕЙ ПО СПЕЦИАЛЬНОСТЯМ",
              r_blocks,
              "Источник данных: Q06_ЗагрузкаВрачейЗаПериод; группировка по полю [Специальность], итоги Sum()",
              "рис08_макет_отчёт_загрузка.png")


# --------------------------------------------------------------------------
# Рис. 9 — диаграмма (Excel-стиль, монохром, строго по п. 1.7 нормоконтроля)
# --------------------------------------------------------------------------
def fig_chart():
    data = build_data()
    spec_names = {s["КодСпециальности"]: s["НазваниеСпециальности"] for s in data["Специальности"]}
    doc_spec = {d["КодВрача"]: d["КодСпециальности"] for d in data["Врачи"]}
    stats = {}
    for z in data["ЗаписиНаПриём"]:
        name = spec_names[doc_spec[z["КодВрача"]]]
        st = stats.setdefault(name, {"Проведён": 0, "Запланирован": 0, "Неявка": 0, "Отменён": 0})
        st[z["СтатусЗаписи"]] += 1
    order = [s["НазваниеСпециальности"] for s in sorted(data["Специальности"],
                                                        key=lambda x: x["ПорядокОтображения"])]
    prived = [stats.get(n, {}).get("Проведён", 0) for n in order]
    plans = [stats.get(n, {}).get("Запланирован", 0) for n in order]
    short = [n if len(n) < 16 else n.split()[0] for n in order]

    fig, ax = plt.subplots(figsize=(16.0 / 2.54, 10.5 / 2.54), dpi=200)
    x = list(range(len(order)))
    ax.bar([i - 0.18 for i in x], prived, width=0.36, label="Принято пациентов",
           color="#3a3a3a", edgecolor=INK, linewidth=1.8)
    ax.bar([i + 0.18 for i in x], plans, width=0.36, label="Запланировано приёмов",
           color="#b8b8b8", edgecolor=INK, linewidth=1.8)
    ax.set_xticks(x)
    ax.set_xticklabels(short, rotation=20, ha="right", fontsize=8)
    ax.tick_params(axis="y", labelsize=8)
    ax.set_ylabel("Количество талонов", fontsize=10, fontweight="bold")
    ax.set_xlabel("Специальность врача", fontsize=10, fontweight="bold", labelpad=6)
    ax.set_title("Загрузка врачей поликлиники по специальностям",
                 fontsize=12, fontweight="bold", pad=10)
    ax.legend(fontsize=8, loc="upper center", bbox_to_anchor=(0.5, -0.26), ncol=2, frameon=False)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    for s in ("left", "bottom"):
        ax.spines[s].set_linewidth(1.0)
    ax.grid(axis="y", color="#d0d0d0", lw=0.6)
    ax.set_axisbelow(True)
    fig.subplots_adjust(bottom=0.28, top=0.88, left=0.10, right=0.97)
    save(fig, "рис09_диаграмма_загрузка.png")
    return order, prived, plans


# --------------------------------------------------------------------------
# Рис. 10 — разграничение доступа
# --------------------------------------------------------------------------
def fig_roles():
    fig, ax = new_fig(17.0, 11.0)
    ax.text(50, 96, "Модель разграничения доступа и администрирования БД «Поликлиника»",
            ha="center", va="center", fontsize=9.2, fontweight="bold")

    def role(x, y, w, h, name, lvl, items, fc):
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.3,rounding_size=0.8",
                                    linewidth=1.1, edgecolor=INK, facecolor=fc))
        ax.text(x + w / 2, y + h - 2.8, name, ha="center", va="center",
                fontsize=8.4, fontweight="bold")
        ax.text(x + w / 2, y + h - 5.8, "Уровень доступа: %d" % lvl, ha="center", va="center",
                fontsize=6.8, color=GREY)
        ax.plot([x + 0.6, x + w - 0.6], [y + h - 7.4, y + h - 7.4], color=INK, lw=0.7)
        for i, t in enumerate(items):
            ax.text(x + 1.0, y + h - 10.0 - i * 3.3, "- " + t, ha="left", va="center", fontsize=6.4)

    role(2, 58, 30, 32, "АДМИНИСТРАТОР", 3,
         ["Структура таблиц и связей", "Учётные записи и роли",
          "Резервное копирование", "Импорт и экспорт в Excel/PDF",
          "Анализ журнала событий"], "#dcdcdc")

    role(35, 58, 30, 32, "ВРАЧ", 2,
         ["Карта приёма пациента", "Диагнозы по МКБ-10",
          "Лист назначений и услуг", "Изменение статуса приёма",
          "Отчёт «История посещений»"], "#e8e8e8")

    role(68, 58, 30, 32, "РЕГИСТРАТОР", 1,
         ["Регистрация пациентов", "Поиск свободных талонов",
          "Запись пациента на приём", "Просмотр справочников",
          "Удаление записей запрещено"], "#f2f2f2")

    objs = ["ПАЦИЕНТЫ\nЗАПИСИ НА ПРИЁМ", "ДИАГНОЗЫ\nНАЗНАЧЕНИЯ",
            "ПОЛЬЗОВАТЕЛИ\nРОЛИ", "ЖУРНАЛ СОБЫТИЙ\nАРХИВ ЗАПИСЕЙ"]
    for i, o in enumerate(objs):
        x = 2 + i * 24.5
        ax.add_patch(Rectangle((x, 26), 22.5, 16, linewidth=1.0, edgecolor=INK, facecolor="white"))
        ax.text(x + 11.25, 34, o, ha="center", va="center", fontsize=7.0, fontweight="bold")

    for i in range(4):
        x = 2 + i * 24.5 + 11.25
        ax.plot([17, x], [58, 42], color="#333333", lw=0.9, ls="--")
    ax.plot([50, 13.25], [58, 42], color="#333333", lw=0.9)
    ax.plot([50, 37.75], [58, 42], color="#333333", lw=0.9)
    ax.plot([83, 13.25], [58, 42], color="#333333", lw=0.9)

    ax.text(50, 12,
            "Средства защиты: авторизация при входе (форма «Ф_Вход»), проверка уровня доступа "
            "ПроверитьПраво(),\nрегистрация действий в таблице «ЖурналСобытий», шифрование БД "
            "паролем и ежедневное резервное копирование.",
            ha="center", va="center", fontsize=6.8,
            bbox=dict(boxstyle="round,pad=0.4", fc=LIGHT, ec=INK, lw=0.7))
    save(fig, "рис10_доступ.png")


if __name__ == "__main__":
    fig_functional()
    fig_er()
    fig_schema()
    fig_normal()
    fig_forms()
    fig_chart()
    fig_roles()
    print("готово:", sorted(os.listdir(OUT)))
