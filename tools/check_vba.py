# -*- coding: utf-8 -*-
"""Статическая проверка сгенерированных VBA-модулей и SQL-скриптов.

Проверяет:
  1) отсутствие недопустимых имён в квадратных скобках ([Имя, с точкой.] -> ошибка 3126);
  2) отсутствие конфликтов имён процедур и глобальных переменных между модулями;
  3) баланс блоков Sub/Function/If/Select/For/Do в каждом модуле;
  4) длину строк VBA (не более 1020 символов);
  5) уникальность ключа (КодВрача, ДатаПриёма, ВремяПриёма) в тестовых данных;
  6) наличие всех полей, используемых в SQL, в описании схемы.
"""
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
sys.path.insert(0, HERE)

from db_schema import TABLES, INDEXES, build_data          # noqa: E402
from queries import QUERIES                                # noqa: E402

FORBIDDEN = ".,!`[]"
problems = []


def check_bracket_names(text, origin):
    for m in re.finditer(r"\[([^\]\n]+)\]", text):
        name = m.group(1)
        bad = [c for c in name if c in FORBIDDEN]
        if bad:
            problems.append("%s: недопустимые символы %r в имени [%s]" % (origin, bad, name))


def strip_comments(line):
    in_str = False
    i = 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            return line[:i]
        i += 1
    return line


def check_module(path):
    origin = os.path.basename(path)
    text = io.open(path, encoding="utf-8").read()
    lines = text.splitlines()
    check_bracket_names(text, origin)
    for n, line in enumerate(lines, 1):
        if len(line) > 1020:
            problems.append("%s:%d: слишком длинная строка (%d символов)" % (origin, n, len(line)))

    # склеиваем продолжения строк (символ подчёркивания в конце строки)
    joined = []
    buf = ""
    for line in lines:
        body = strip_comments(line)
        if buf:
            body = buf + " " + body.strip()
            buf = ""
        if body.rstrip().endswith("_"):
            buf = body.rstrip()[:-1]
            continue
        joined.append(body)
    code = "\n".join(joined)
    pairs = [("Sub", "End Sub"), ("Function", "End Function"),
             ("If ", "End If"), ("Select Case", "End Select"),
             ("For ", "Next"), ("Do ", "Loop"), ("With ", "End With")]
    for open_kw, close_kw in pairs:
        opened = len(re.findall(r"(?m)^\s*(?:Public |Private |Friend )?%s" % open_kw, code))
        closed = len(re.findall(r"(?m)^\s*%s" % close_kw, code))
        # однострочные If без End If
        opened -= len(re.findall(r"(?m)^\s*(?:If|For|Do)\b.*\bThen\b\s+\S+.*$", code)) if open_kw == "If " else 0
        if open_kw == "If ":
            opened = len(re.findall(
                r"(?m)^\s*(?:If\b.*\bThen\s*$|ElseIf\b)", code))
            closed = len(re.findall(r"(?m)^\s*End If\b", code))
        if open_kw == "For ":
            opened = len(re.findall(r"(?m)^\s*For\b", code))
            closed = len(re.findall(r"(?m)^\s*Next\b", code))
        if open_kw == "Do ":
            opened = len(re.findall(r"(?m)^\s*Do\b", code))
            closed = len(re.findall(r"(?m)^\s*Loop\b", code))
        if open_kw == "Sub":
            opened = len(re.findall(r"(?m)^\s*(?:Public |Private |Friend )?Sub\s+\w+", code))
            closed = len(re.findall(r"(?m)^\s*End Sub\b", code))
        if open_kw == "Function":
            opened = len(re.findall(r"(?m)^\s*(?:Public |Private |Friend )?Function\s+\w+", code))
            closed = len(re.findall(r"(?m)^\s*End Function\b", code))
        if open_kw == "With ":
            continue
        if opened != closed:
            problems.append("%s: несбалансированные блоки %s/%s: %d != %d"
                            % (origin, open_kw.strip(), close_kw, opened, closed))

    decls = set()
    for m in re.finditer(r"(?m)^\s*(?:Public |Private |Friend )?(?:Sub|Function)\s+(\w+)", code):
        name = m.group(1)
        if name in decls:
            problems.append("%s: повторное объявление процедуры %s" % (origin, name))
        decls.add(name)
    publics = set()
    for m in re.finditer(r"(?m)^\s*Public\s+(?:Const\s+)?(\w+)\s+As", code):
        name = m.group(1)
        if name in publics:
            problems.append("%s: повторное объявление Public %s" % (origin, name))
        publics.add(name)
    return decls, publics


def _proc_bodies(text, name_re):
    """Возвращает карту «имя процедуры -> текст её тела»."""
    lines = text.splitlines()
    bodies, cur, buf = {}, None, []
    for line in lines:
        m = re.match(r"\s*(?:Private|Public|Friend)\s+(?:Sub|Function)\s+(\w+)", line)
        if m and re.match(name_re, m.group(1)):
            cur, buf = m.group(1), []
            continue
        if cur is not None:
            if re.match(r"\s*End (?:Sub|Function)", line):
                bodies[cur] = buf
                cur = None
            else:
                buf.append(line)
    return bodies


def _module_text(body):
    """Склеивает VBA-литералы из строк вида  s = s & "..."  (модуль формы/отчёта)."""
    parts = []
    for line in body:
        for lit in re.finditer(r'"(?:[^"]|"")*"', line):
            s = lit.group(0)
            if s.startswith('""'):
                continue
            parts.append(s[1:-1].replace('""', '"'))
    return "\n".join(parts)


def check_form_modules(text, table_fields, query_aliases):
    """Проверяет модули, которые сборщик формирует для форм динамически:
    баланс блоков, а также ссылки Me!<Имя> на существующие элементы управления
    или поля источника записей."""
    creators = _proc_bodies(text, r"^СоздатьФорму")
    codes = _proc_bodies(text, r"^FormCode")
    # какая процедура создания формы подключает какой модуль
    links = {}
    for cname, body in creators.items():
        joined = "\n".join(body)
        for ccode in codes:
            if "AddFromString %s()" % ccode in joined:
                links[ccode] = cname
    for ccode, cname in sorted(links.items()):
        code = _module_text(codes[ccode])
        for open_kw, close_kw in (("Sub", "End Sub"), ("Function", "End Function")):
            opened = len(re.findall(r"(?m)^\s*(?:Private |Public )?%s \w+" % open_kw, code))
            closed = len(re.findall(r"(?m)^\s*%s" % close_kw, code))
            if opened != closed:
                problems.append("%s -> %s: несбалансированные блоки %s/%s: %d != %d"
                                % (cname, ccode, open_kw, close_kw, opened, closed))
        opened = len(re.findall(r"(?m)^\s*If\b.*\bThen\s*$", code))
        closed = len(re.findall(r"(?m)^\s*End If\b", code))
        if opened != closed:
            problems.append("%s -> %s: несбалансированные блоки If/End If: %d != %d"
                            % (cname, ccode, opened, closed))
        src_text = "\n".join(creators[cname])
        ctl_names = set(re.findall(r'ctl\.Name = "([^"]+)"', src_text))
        # элементы, имена которых формируются в цикле (ПолеН1, ПолеД2 и т. п.)
        loop_names = re.findall(r'ctl\.Name = "([^"]*)" & CStr\(i \+ 1\)', src_text)
        # поля источника записей формы (таблица или запрос с псевдонимами полей)
        known_fields = set(table_fields)
        for m in re.finditer(r'frm\.RecordSource = "([^"]*)"', src_text):
            known_fields |= query_aliases.get(m.group(1), set())
        for ref in set(re.findall(r"Me!\s*(?:\[)?([A-Za-zА-Яа-яЁё_][\w]*)", code)):
            if ref in ctl_names or ref in known_fields:
                continue
            if any(ref.startswith(prefix) for prefix in loop_names):
                continue
            problems.append("%s -> %s: ссылка Me!%s не найдена среди элементов управления формы"
                            % (cname, ccode, ref))


def main():
    modules = ["VBA/AutoBuild_UTF8.bas", "VBA/ModSecurity_UTF8.bas", "VBA/ModAdmin_UTF8.bas"]
    all_proc, all_pub = {}, {}
    for mod in modules:
        procs, pubs = check_module(os.path.join(ROOT, mod))
        for name in procs:
            if name in all_proc:
                problems.append("Конфликт имён процедур: %s (%s и %s)" % (name, all_proc[name], mod))
            all_proc[name] = mod
        for name in pubs:
            if name in all_pub:
                problems.append("Конфликт глобальных переменных: %s (%s и %s)"
                                % (name, all_pub[name], mod))
            all_pub[name] = mod

    # SQL: имена в скобках и ссылки на поля схемы
    fields = {t["name"]: {f[0] for f in t["fields"]} for t in TABLES}
    for q in QUERIES:
        check_bracket_names(q["sql"], q["name"])
        for tname, fname in re.findall(r"([А-Яа-яЁёA-Za-z0-9_]+)\.([А-Яа-яЁёA-Za-z0-9_*]+)", q["sql"]):
            if fname == "*" or tname.lower() in ("dd", "mm", "yyyy", "hh", "nn", "ss", "date", "time"):
                continue
            if tname not in fields:
                problems.append("%s: неизвестная таблица %s" % (q["name"], tname))
            elif fname not in fields[tname]:
                problems.append("%s: неизвестное поле %s.%s" % (q["name"], tname, fname))
    for path in ("SQL/01_Создание_Таблиц.sql", "SQL/02_Связи_и_Индексы.sql",
                 "SQL/03_Тестовые_Данные.sql", "SQL/04_Запросы.sql",
                 "SQL/Поликлиника_Полный_Скрипт.sql"):
        check_bracket_names(io.open(os.path.join(ROOT, path), encoding="utf-8").read(), path)

    # контроль: имена полей в ControlSource/RowSource существуют в схеме или запросах
    known = set()
    for tbl in TABLES:
        known.update(f[0] for f in tbl["fields"])
    for q in QUERIES:
        alias = re.search(r"\bFROM\b", q["sql"], re.IGNORECASE)
        select = q["sql"][:alias.start()] if alias else q["sql"]
        for m in re.finditer(r"\bAS\s+\[([^\]]+)\]", select, re.IGNORECASE):
            known.add(m.group(1))
        for m in re.finditer(r"\bAS\s+([\wА-Яа-яЁё]+)(?!\s*\[)", select, re.IGNORECASE):
            known.add(m.group(1))
        for m in re.finditer(r"(?<![\w.\]])\b([А-ЯЁ][А-Яа-яЁё0-9_]+)\b(?!\s*[.(])", select):
            known.add(m.group(1))
    all_fields = set()
    for tbl in TABLES:
        all_fields.update(f[0] for f in tbl["fields"])
    query_aliases = {}
    for q in QUERIES:
        from_pos = re.search(r"\bFROM\b", q["sql"], re.IGNORECASE)
        select = q["sql"][:from_pos.start()] if from_pos else q["sql"]
        names = set()
        for m in re.finditer(r"\bAS\s+\[([^\]]+)\]", select, re.IGNORECASE):
            names.add(m.group(1))
        for m in re.finditer(r"\bAS\s+([\wА-Яа-яЁё]+)(?!\s*\[)", select, re.IGNORECASE):
            names.add(m.group(1))
        query_aliases[q["name"]] = names
    vba = io.open(os.path.join(ROOT, "VBA", "AutoBuild_UTF8.bas"), encoding="utf-8").read()
    check_form_modules(vba, all_fields, query_aliases)
    for m in re.finditer(r'(?:ControlSource|RowSource)\s*=\s*("[^"]*")', vba):
        literal = m.group(1)
        if literal.startswith('"="='):            # выражение-вычисление, разбираем ниже
            pass
        for b in re.finditer(r"\[([^\]]+)\]", literal):
            name = b.group(1)
            if name not in known and not name.startswith("="):
                problems.append("AutoBuild: неизвестное имя поля [%s] в ControlSource/RowSource" % name)

    # уникальность ключа уникального индекса в тестовых данных
    data = build_data()
    for iname, tbl, cols, uniq, _cmt in INDEXES:
        if not uniq:
            continue
        seen = set()
        for row in data.get(tbl, []):
            key = tuple(str(row.get(c)) for c in cols)
            if key in seen:
                problems.append("Дублирование уникального ключа %s в таблице %s: %s"
                                % (iname, tbl, key))
            seen.add(key)

    if problems:
        print("ОБНАРУЖЕНЫ ПРОБЛЕМЫ (%d):" % len(problems))
        for p in problems:
            print("  -", p)
        return 1
    print("Проверка пройдена: модулей %d, процедур %d, глобальных переменных %d, запросов %d."
          % (len(modules), len(all_proc), len(all_pub), len(QUERIES)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
