# -*- coding: utf-8 -*-
"""Генератор SQL-скриптов для MS Access из единого описания схемы."""
import os
from datetime import date, datetime, time

from db_schema import TABLES, INDEXES, FOREIGN_KEYS, build_data
from queries import QUERIES

OUT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "SQL"))

HEAD = """' ============================================================================
'  БАЗА ДАННЫХ «ПОЛИКЛИНИКА»
'  Отчёт по производственной практике
'  Дисциплина: Администрирование баз данных (СУБД MS Access)
'  Вариант 14, предметная область «Поликлиника»
'  Выполнил: студент группы Р-25-29 Сагадиев А. Р.
'  Проверил: преподаватель Габитова А. И.                           Уфа, 2026
' ============================================================================
'  ВНИМАНИЕ! Microsoft Access выполняет только ОДИН оператор SQL за один
'  запрос. Скрипт разбит на отдельные операторы, каждый из которых
'  заканчивается точкой с запятой. Порядок выполнения: 01 -> 02 -> 03 -> 04,
'  внутри файла — сверху вниз.
'  Способ 1 (автомат, 1 клик): запустить скрипт Создать_БД_Поликлиника.vbs.
'  Способ 2 (через VBA): импортировать VBA\\AutoBuild.bas и запустить СоздатьБД.
'  Способ 3 (вручную): Создание -> Конструктор запросов -> Режим SQL ->
'                      вставить один оператор -> Выполнить.
' ============================================================================
"""

ORDER = ["Специальности", "Врачи", "Пациенты", "Услуги", "МКБ", "РасписаниеПриёма",
         "ЗаписиНаПриём", "Диагнозы", "Назначения", "Роли", "Пользователи",
         "ЖурналСобытий"]


def esc(v):
    """Значение -> SQL-литерал Access."""
    if v is None:
        return "Null"
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, (int, float)):
        return str(v)
    if isinstance(v, datetime):
        return "#%02d/%02d/%04d %02d:%02d:%02d#" % (v.month, v.day, v.year, v.hour, v.minute, v.second)
    if isinstance(v, date):
        return "#%02d/%02d/%04d#" % (v.month, v.day, v.year)
    if isinstance(v, time):
        return "#%02d:%02d:%02d#" % (v.hour, v.minute, v.second)
    s = str(v).replace("'", "''")
    return "'%s'" % s


def fix_quotes(sql: str) -> str:
    """SQL-литералы в одинарных кавычках (совместимость с VBA)."""
    return sql.replace('"', "'")


def ddl_tables():
    out = ["' --- 01. СОЗДАНИЕ ТАБЛИЦ (Jet/ACE SQL DDL) -------------------------", ""]
    for t in TABLES:
        lines = ["' %s. %s" % (t["name"], t["comment"])]
        for cname, cexpr in t.get("checks", []):
            lines.append("' Ограничение %s: CHECK (%s)" % (cname, cexpr))
        cols = []
        for fld, ftype, extra in t["fields"]:
            frag = "    [%s] %s" % (fld, ftype)
            if "PK" in extra:
                frag += " CONSTRAINT PK_%s PRIMARY KEY" % t["name"]
            if "NOT NULL" in extra:
                frag += " NOT NULL"
            if "UNIQUE" in extra:
                frag += " UNIQUE"
            cols.append(frag)
        lines.append("CREATE TABLE [%s] (" % t["name"])
        lines.append(",\n".join(cols))
        lines.append(");")
        lines.append("")
        out.extend(lines)
    return out


def ddl_props():
    out = ["' --- 01.2. СВОЙСТВА ПОЛЕЙ (DAO, выполняется из VBA) ---------------",
           "' Маски ввода, подписи столбцов, значения по умолчанию и условия на",
           "' значение устанавливаются процедурой УстановитьСвойстваПолей() модуля",
           "' AutoBuild.bas через объектную модель DAO.",
           ""]
    for t in TABLES:
        props = t.get("props", {})
        if not props:
            continue
        out.append("' Таблица %s:" % t["name"])
        for fld in props:
            for pname, pval in props[fld].items():
                out.append("'   [%s].%s = %s" % (fld, pname, pval))
        out.append("")
    return out


def ddl_relations():
    out = ["' --- 02. СВЯЗИ (ВНЕШНИЕ КЛЮЧИ) И ИНДЕКСЫ --------------------------", ""]
    for name, child, cfield, parent, pfield, ondel in FOREIGN_KEYS:
        out.append("' %s -> %s (при удалении: %s)" % (child, parent, ondel))
        out.append("ALTER TABLE [%s] ADD CONSTRAINT %s FOREIGN KEY ([%s]) REFERENCES [%s] ([%s]);"
                   % (child, name, cfield, parent, pfield))
    out.append("")
    for iname, tbl, cols, uniq, cmt in INDEXES:
        out.append("' %s" % cmt)
        out.append("CREATE %sINDEX %s ON [%s] (%s);"
                   % ("UNIQUE " if uniq else "", iname, tbl, ", ".join("[%s]" % c for c in cols)))
    out.append("")
    return out


def dml_data():
    data = build_data()
    out = ["' --- 03. ЗАПОЛНЕНИЕ ТАБЛИЦ ТЕСТОВЫМИ ДАННЫМИ ---------------------", ""]
    for tbl in ORDER:
        rows = data[tbl]
        out.append("' %s: %d записей" % (tbl, len(rows)))
        for r in rows:
            cols = ", ".join("[%s]" % c for c in r.keys())
            vals = ", ".join(esc(v) for v in r.values())
            out.append("INSERT INTO [%s] (%s) VALUES (%s);" % (tbl, cols, vals))
        out.append("")
    return out


def ddl_queries():
    out = ["' --- 04. ЗАПРОСЫ --------------------------------------------------", ""]
    for q in QUERIES:
        out.append("' %s — %s" % (q["name"], q["kind"]))
        out.append("' Параметры: %s" % q["params"])
        out.append("' %s" % q["comment"])
        out.append(fix_quotes(q["sql"]).rstrip().rstrip(";") + ";")
        out.append("")
    return out


def write(name, header, lines):
    path = os.path.join(OUT, name)
    with open(path, "w", encoding="utf-8") as f:
        f.write(header + "\n" + "\n".join(lines).rstrip() + "\n")
    print("  записан %s (%d строк)" % (name, len(lines)))
    return lines


def main():
    os.makedirs(OUT, exist_ok=True)
    parts = {}
    parts["01"] = write("01_Создание_Таблиц.sql", HEAD, ddl_tables() + ddl_props())
    parts["02"] = write("02_Связи_и_Индексы.sql", HEAD, ddl_relations())
    parts["03"] = write("03_Тестовые_Данные.sql", HEAD, dml_data())
    parts["04"] = write("04_Запросы.sql", HEAD, ddl_queries())
    write("Поликлиника_Полный_Скрипт.sql", HEAD,
          parts["01"] + parts["02"] + parts["03"] + parts["04"])


if __name__ == "__main__":
    main()
