# -*- coding: utf-8 -*-
"""Генератор VBA-модулей, текстовых дампов макросов и VBS-сборщика для MS Access."""
import os
import textwrap

from db_schema import TABLES, INDEXES, FOREIGN_KEYS, build_data
from queries import QUERIES
from gen_sql import esc, fix_quotes, ORDER

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
VBA_DIR = os.path.join(ROOT, "VBA")
MACRO_DIR = os.path.join(ROOT, "Макросы")

TAB = "    "


def vba_literal_chunks(s: str, width: int = 840):
    """Делит текст на фрагменты, не разрывая экранированные пары кавычек."""
    s = s.replace('"', '""')
    chunks = []
    i = 0
    while i < len(s):
        end = min(i + width, len(s))
        if end < len(s):
            tail = 0
            while end - 1 - tail >= i and s[end - 1 - tail] == '"':
                tail += 1
            if tail % 2 == 1:          # фрагмент заканчивается одиночной кавычкой
                end -= 1
                while end > i and s[end - 1] == '"':
                    end -= 1
                if end <= i:
                    end = min(i + width, len(s))
        chunks.append(s[i:end])
        i = end
    return chunks or [""]


def vba_expr(sql: str, indent: int = 2) -> str:
    chunks = vba_literal_chunks(sql)
    pad = TAB * indent
    lines = []
    for i, ch in enumerate(chunks):
        tail = " & _" if i < len(chunks) - 1 else ""
        lines.append(('"' + ch + '"' + tail) if i == 0 else (pad + '"' + ch + '"' + tail))
    return "\n".join(lines)


def exec_call(sql: str, func: str = "Exec") -> str:
    sql = " ".join(line.strip() for line in sql.splitlines() if line.strip()).strip()
    sql = fix_quotes(sql)
    if len(sql) <= 840:
        return TAB + '%s "%s"' % (func, sql.replace('"', '""'))
    return TAB + func + " " + vba_expr(sql, indent=2)


def comment(text: str, width: int = 92) -> str:
    return "\n".join(TAB + "' " + l for l in textwrap.wrap(text, width))


HEADER = """Attribute VB_Name = "AutoBuild"
' ============================================================================
'  МОДУЛЬ АВТОМАТИЧЕСКОГО ПОСТРОЕНИЯ БАЗЫ ДАННЫХ «ПОЛИКЛИНИКА»
'  Дисциплина : Администрирование баз данных (СУБД MS Access)
'  Вариант 14 : предметная область «Поликлиника»
'  Выполнил   : студент группы Р-25-29 Сагадиев А. Р.
'  Проверил   : преподаватель Габитова А. И.                      Уфа, 2026 г.
' ----------------------------------------------------------------------------
'  НАЗНАЧЕНИЕ. Одна процедура СоздатьБД() полностью собирает базу данных:
'    этап 1  - очистка предыдущей версии базы (связи, таблицы, запросы,
'              формы, отчёты, макросы);
'    этап 2  - создание 13 таблиц со всеми типами данных и первичными ключами;
'    этап 3  - установка свойств полей через DAO: значения по умолчанию,
'              условия на значение, маски ввода, подписи столбцов, форматы;
'    этап 4  - создание 12 связей (внешних ключей) с обеспечением ссылочной
'              целостности, каскадным удалением и обновлением;
'    этап 5  - создание 10 индексов, в том числе уникального составного
'              индекса (КодВрача + ДатаПриёма + ВремяПриёма), запрещающего
'              пересечение времени приёма у одного врача;
'    этап 6  - заполнение таблиц тестовыми данными;
'    этап 7  - создание 16 запросов (выборка, параметрические, с вычисляемыми
'              полями, итоговые, на добавление, на обновление, на удаление);
'    этап 8  - создание 7 форм (главная с двумя подчинёнными, форма ввода,
'              ленточная, две подчинённые, авторизация, кнопочная);
'    этап 9  - создание 4 отчётов с группировкой и итогами;
'    этап 10 - загрузка макросов и настройка параметров запуска приложения.
' ----------------------------------------------------------------------------
'  КАК ПОЛЬЗОВАТЬСЯ
'  Вариант А (в 1 клик): запустить скрипт Создать_БД_Поликлиника.vbs.
'  Вариант Б (вручную из Access):
'    1. Создать пустую базу «Поликлиника.accdb».
'    2. Alt+F11 -> Файл -> Импорт файла -> AutoBuild.bas, ModSecurity.bas,
'       ModAdmin.bas.
'    3. Нажать Alt+F8 -> выбрать «СоздатьБД» -> «Выполнить».
' ============================================================================
Option Compare Database
Option Explicit

Private mOk As Long          ' счётчик успешно выполненных операторов
Private mOkForm As Long      ' счётчик созданных форм
Private mOkReport As Long    ' счётчик созданных отчётов
Private mOkQuery As Long     ' счётчик созданных запросов
Private mOkMacro As Long     ' счётчик загруженных макросов
Private mErr As Long         ' счётчик ошибок
Private mLog As String       ' журнал построения

Public gLogin As String      ' логин текущего пользователя
Public gRole As String       ' роль текущего пользователя
Public gLevel As Long        ' уровень доступа: 1 - регистратор, 2 - врач, 3 - администратор
Public gKodVracha As Long    ' код врача (для роли «Врач»)
"""


def part_tables():
    """Этап 2: создание таблиц (совместимо с DAO Jet/ACE SQL)."""
    out = ["Private Sub СоздатьТаблицы()",
           comment("Создание таблиц. Оператор DROP удаляет предыдущую версию таблицы; "
                   "ошибка удаления (таблицы нет) игнорируется.")]
    for t in TABLES:
        out.append("")
        out.append(comment("%s — %s" % (t["name"], t["comment"])))
        out.append(exec_call("DROP TABLE [%s]" % t["name"], "ExecSilent"))
        cols = []
        for fld, ftype, extra in t["fields"]:
            frag = "[%s] %s" % (fld, ftype)
            if "PK" in extra:
                frag += " CONSTRAINT PK_%s PRIMARY KEY" % t["name"]
            if "NOT NULL" in extra:
                frag += " NOT NULL"
            if "UNIQUE" in extra:
                frag += " UNIQUE"
            cols.append(frag)
        ddl = "CREATE TABLE [%s] (%s);" % (t["name"], ", ".join(cols))
        out.append(exec_call(ddl))
    out.append("    LogLine \"  создано таблиц: " + str(len(TABLES)) + "\"")
    out.append("End Sub")
    return out


def part_props():
    """Этап 3: свойства полей через DAO."""
    out = ["Private Sub УстановитьСвойстваПолей()",
           comment("Значения по умолчанию, условия и сообщения проверки, маски ввода, "
                   "подписи и форматы полей через объектную модель DAO."),
           "    Dim db As DAO.Database",
           "    Dim tdf As DAO.TableDef",
           "    Dim fld As DAO.Field",
           "    Set db = CurrentDb",
           "    On Error Resume Next",
           ""]
    for t in TABLES:
        props = t.get("props", {})
        if not props:
            continue
        out.append(comment("Таблица «%s»" % t["name"]))
        out.append('    Set tdf = db.TableDefs("%s")' % t["name"])
        for fld, plist in props.items():
            out.append('    Set fld = tdf.Fields("%s")' % fld)
            for pname, pval in plist.items():
                val_esc = str(pval).replace('"', '""')
                if pname in ("DefaultValue", "ValidationRule", "ValidationText"):
                    out.append('    fld.%s = "%s"' % (pname, val_esc))
                else:
                    typ = "dbBoolean" if isinstance(pval, bool) else "dbText"
                    out.append('    AddProp fld, "%s", %s, "%s"' % (pname, typ, val_esc))
        out.append("")
    out.append("    On Error GoTo 0")
    out.append('    LogLine "  свойства полей (маски ввода, подписи, проверки) установлены"')
    out.append("End Sub")
    return out


def part_relations():
    """Этап 4: внешние ключи."""
    out = ["Private Sub СоздатьСвязи()",
           comment("Связи между таблицами с обеспечением ссылочной целостности."),
           "    Dim db As DAO.Database",
           "    Dim rel As DAO.Relation",
           "    Set db = CurrentDb",
           "    On Error Resume Next",
           ""]
    for name, child, cfield, parent, pfield, ondel in FOREIGN_KEYS:
        out.append(comment("Связь %s: %s.[%s] -> %s.[%s] (при удалении: %s)"
                           % (name, child, cfield, parent, pfield, ondel)))
        out.append("    Set rel = db.CreateRelation(\"%s\", \"%s\", \"%s\", dbRelationUpdateCascade)" %
                   (name, parent, child))
        out.append('    rel.Fields.Append rel.CreateField("%s")' % pfield)
        out.append('    rel.Fields("%s").ForeignName = "%s"' % (pfield, cfield))
        if ondel == "CASCADE":
            out.append("    rel.Attributes = rel.Attributes Or dbRelationDeleteCascade")
        out.append("    db.Relations.Append rel")
        out.append("")
    out.append("    On Error GoTo 0")
    out.append('    LogLine "  создано связей: ' + str(len(FOREIGN_KEYS)) + '"')
    out.append("End Sub")
    return out


def part_indexes():
    """Этап 5: индексы."""
    out = ["Private Sub СоздатьИндексы()",
           comment("Индексы. Уникальный составной индекс IX_Zap_Vrach_Time "
                   "(КодВрача + ДатаПриёма + ВремяПриёма) реализует требование "
                   "«время приёма не может пересекаться у одного врача».")]
    for iname, tbl, cols, uniq, cmt in INDEXES:
        kw = "UNIQUE " if uniq else ""
        out.append(comment(cmt))
        out.append(exec_call("CREATE %sINDEX %s ON [%s] (%s)"
                             % (kw, iname, tbl, ", ".join("[%s]" % c for c in cols))))
    out.append('    LogLine "  создано индексов: ' + str(len(INDEXES)) + '"')
    out.append("End Sub")
    return out


def part_data():
    """Этап 6: тестовые данные (разбиты на подпроцедуры по 140 операторов)."""
    data = build_data()
    stmts = []
    counts = {}
    for tbl in ORDER:
        rows = data[tbl]
        counts[tbl] = len(rows)
        stmts.append(("' %s (%d записей)" % (tbl, len(rows)), None))
        for r in rows:
            cols = ", ".join("[%s]" % c for c in r.keys())
            vals = ", ".join(esc(v) for v in r.values())
            stmts.append(("INSERT INTO [%s] (%s) VALUES (%s);" % (tbl, cols, vals), None))

    chunk_size = 140
    chunks = [stmts[i:i + chunk_size] for i in range(0, len(stmts), chunk_size)]
    out = []
    names = ["ЗаполнитьДанные_%d" % i for i in range(1, len(chunks) + 1)]
    out.append("Private Sub ЗаполнитьДанные()")
    for n in names:
        out.append(TAB + n)
    out.append('    LogLine "  записей загружено: ' +
               " + ".join([str(counts[t]) for t in ORDER]) + '"')
    out.append("End Sub")
    out.append("")
    for n, ch in zip(names, chunks):
        out.append("Private Sub %s()" % n)
        for txt, _ in ch:
            if txt.startswith("'"):
                out.append(TAB + txt if txt.startswith("' ") else TAB + "' " + txt.lstrip("'"))
            else:
                out.append(exec_call(txt))
        out.append("End Sub")
        out.append("")
    return out


def part_queries():
    """Этап 7: запросы."""
    out = ["Private Sub СоздатьЗапросы()",
           "    Dim db As DAO.Database",
           "    Dim qd As DAO.QueryDef",
           "    Set db = CurrentDb",
           ""]
    for q in QUERIES:
        sql = " ".join(line.strip() for line in q["sql"].splitlines() if line.strip()).strip().rstrip(";")
        sql = fix_quotes(sql)
        out.append(comment("%s — %s. %s Параметры: %s."
                           % (q["name"], q["kind"], q["comment"], q["params"])))
        out.append('    On Error Resume Next')
        out.append('    db.QueryDefs.Delete "' + q["name"] + '"')
        out.append('    Err.Clear')
        if len(sql) <= 840:
            out.append('    Set qd = db.CreateQueryDef("' + q["name"] + '", "' + sql.replace('"', '""') + '")')
        else:
            out.append('    Set qd = db.CreateQueryDef("' + q["name"] + '", ' + vba_expr(sql, indent=2) + ")")
        out.append('    If Err.Number <> 0 Then')
        out.append('        mErr = mErr + 1')
        out.append('        mLog = mLog & "ОШИБКА запроса ' + q["name"] + ': " & Err.Description & vbCrLf')
        out.append('        Err.Clear')
        out.append('    Else')
        out.append('        AddProp qd, "Description", dbText, "' + q["comment"].replace('"', "'") + '"')
        out.append('        qd.Close')
        out.append('        mOk = mOk + 1')
        out.append('        mOkQuery = mOkQuery + 1')
        out.append('    End If')
        out.append('    On Error GoTo 0')
        out.append("")
    out.append('    LogLine "  создано запросов: ' + str(len(QUERIES)) + '"')
    out.append("End Sub")
    return out


FORM_PATIENTS = '''
' ----------------------------------------------------------------------------
'  Форма «Ф_Пациенты» — ленточная форма картотеки
' ----------------------------------------------------------------------------
Private Sub СоздатьФормуПациенты()
    Dim frmName As String, tmpName As String
    Dim frm As Form
    Dim ctl As Object
    Dim flds As Variant, cols As Variant, w As Variant
    Dim i As Integer, x As Long
    frmName = "Ф_Пациенты"

    On Error Resume Next
    DoCmd.Close acForm, frmName, acSaveNo
    DoCmd.DeleteObject acForm, frmName
    Err.Clear
    On Error GoTo 0

    Set frm = Application.CreateForm()
    tmpName = frm.Name
    On Error Resume Next
    DoCmd.RunCommand acCmdFormHdrFtr
    On Error GoTo 0

    frm.RecordSource = "Q04_ПациентыСВозрастом"
    frm.Caption = "Пациенты поликлиники"
    frm.DefaultView = 1                  ' 1 = непрерывные формы (ленточная)
    frm.NavigationButtons = True
    frm.DividingLines = True
    frm.InsideWidth = 14200
    frm.Section(acDetail).Height = 420
    frm.Section(acHeader).Height = 1000
    frm.Section(acFooter).Height = 700

    Set ctl = CreateFormControl(tmpName, acLabel, acHeader, , , 200, 120, 7000, 420)
    ctl.Caption = "КАРТОТЕКА ПАЦИЕНТОВ ПОЛИКЛИНИКИ"
    ctl.FontName = "Tahoma": ctl.FontSize = 14: ctl.FontBold = True

    flds = Array("Пациент", "Дата рождения", "Возраст", "Пол", "Группа", "Полис ОМС", "Телефон", "Участок")
    cols = Array("Пациент", "Дата рождения", "Возраст", "Пол", "Группа", "Полис ОМС", "Телефон", "Участок")
    w = Array(4200, 1700, 900, 700, 2100, 2300, 2000, 900)

    x = 200
    For i = 0 To UBound(flds)
        Set ctl = CreateFormControl(tmpName, acLabel, acHeader, , , x, 620, CLng(w(i)), 300)
        ctl.Caption = CStr(flds(i))
        ctl.FontBold = True
        Set ctl = CreateFormControl(tmpName, acTextBox, acDetail, , , x, 60, CLng(w(i)), 320)
        ctl.ControlSource = "[" & CStr(cols(i)) & "]"
        ctl.Name = "Поле" & CStr(i + 1)
        x = x + CLng(w(i)) + 60
    Next i

    Set ctl = CreateFormControl(tmpName, acCommandButton, acHeader, , , 11400, 120, 2400, 420)
    ctl.Name = "КнопкаКарта"
    ctl.Caption = "Карта приёмов"
    ctl.OnClick = "[Event Procedure]"
    Set ctl = CreateFormControl(tmpName, acCommandButton, acFooter, , , 11400, 100, 2400, 460)
    ctl.Name = "КнопкаЗакрыть"
    ctl.Caption = "Закрыть"
    ctl.OnClick = "[Event Procedure]"

    frm.HasModule = True
    frm.Module.AddFromString FormCodePatients()

    DoCmd.Close acForm, tmpName, acSaveYes
    DoCmd.Rename frmName, acForm, tmpName
    LogLine "  форма «Ф_Пациенты» (ленточная) создана"
End Sub

Private Function FormCodePatients() As String
    Dim s As String
    s = "Private Sub Form_Load()" & vbCrLf
    s = s & "    Me.Caption = ""Картотека пациентов: "" & DCount(""*"", ""Пациенты"") & "" чел.""" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub КнопкаКарта_Click()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    DoCmd.OpenForm ""Ф_ПриёмПациента"", , , ""КодПациента="" & Nz(Me![Код], 0)" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub КнопкаЗакрыть_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FormCodePatients = s
End Function
'''

FORM_VISIT = '''
' ----------------------------------------------------------------------------
'  Форма «Ф_ЗаписьНаПриём» — форма ввода новой записи (регистратура)
' ----------------------------------------------------------------------------
Private Sub СоздатьФормуЗапись()
    Dim frmName As String, tmpName As String
    Dim frm As Form
    Dim ctl As Object
    Dim y As Long
    frmName = "Ф_ЗаписьНаПриём"

    On Error Resume Next
    DoCmd.Close acForm, frmName, acSaveNo
    DoCmd.DeleteObject acForm, frmName
    Err.Clear
    On Error GoTo 0

    Set frm = Application.CreateForm()
    tmpName = frm.Name
    On Error Resume Next
    DoCmd.RunCommand acCmdFormHdrFtr
    On Error GoTo 0

    frm.RecordSource = "ЗаписиНаПриём"
    frm.Caption = "Запись на приём"
    frm.DefaultView = 0
    frm.InsideWidth = 12400
    frm.Section(acDetail).Height = 6200
    frm.Section(acHeader).Height = 700

    Set ctl = CreateFormControl(tmpName, acLabel, acHeader, , , 200, 120, 8000, 420)
    ctl.Caption = "ЗАПИСЬ ПАЦИЕНТА НА ПРИЁМ"
    ctl.FontName = "Tahoma": ctl.FontSize = 14: ctl.FontBold = True

    y = 200
    Set ctl = CreateFormControl(tmpName, acLabel, acDetail, , , 300, y, 3000, 320)
    ctl.Caption = "Номер талона:"
    Set ctl = CreateFormControl(tmpName, acTextBox, acDetail, , , 3500, y, 3600, 320)
    ctl.ControlSource = "НомерТалона": ctl.Name = "ПолеТалон"
    y = y + 600

    Set ctl = CreateFormControl(tmpName, acLabel, acDetail, , , 300, y, 3000, 320)
    ctl.Caption = "Пациент:"
    Set ctl = CreateFormControl(tmpName, acComboBox, acDetail, , , 3500, y, 6200, 320)
    ctl.Name = "ПолеПациент"
    ctl.ControlSource = "КодПациента"
    ctl.RowSourceType = "Table/Query"
    ctl.RowSource = "SELECT Пациенты.КодПациента, [Пациенты].[Фамилия] & "" "" & [Пациенты].[Имя] & "" "" & [Пациенты].[Отчество] AS ФИО, Пациенты.НомерПолиса FROM Пациенты ORDER BY Пациенты.Фамилия, Пациенты.Имя;"
    ctl.ColumnCount = 3
    ctl.ColumnWidths = "0;4200;2200"
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    y = y + 600

    Set ctl = CreateFormControl(tmpName, acLabel, acDetail, , , 300, y, 3000, 320)
    ctl.Caption = "Врач:"
    Set ctl = CreateFormControl(tmpName, acComboBox, acDetail, , , 3500, y, 6200, 320)
    ctl.Name = "ПолеВрач"
    ctl.ControlSource = "КодВрача"
    ctl.RowSourceType = "Table/Query"
    ctl.RowSource = "SELECT Врачи.КодВрача, [Врачи].[Фамилия] & "" "" & [Врачи].[Имя] & "" "" & [Врачи].[Отчество] AS ФИО, Специальности.НазваниеСпециальности FROM Врачи INNER JOIN Специальности ON Врачи.КодСпециальности = Специальности.КодСпециальности ORDER BY Врачи.Фамилия;"
    ctl.ColumnCount = 3
    ctl.ColumnWidths = "0;4200;2200"
    ctl.BoundColumn = 1
    ctl.LimitToList = True
    ctl.AfterUpdate = "[Event Procedure]"
    y = y + 600

    Set ctl = CreateFormControl(tmpName, acLabel, acDetail, , , 300, y, 3000, 320)
    ctl.Caption = "Дата приёма:"
    Set ctl = CreateFormControl(tmpName, acTextBox, acDetail, , , 3500, y, 2200, 320)
    ctl.Name = "ПолеДата"
    ctl.ControlSource = "ДатаПриёма"
    ctl.Format = "dd.mm.yyyy"
    y = y + 600

    Set ctl = CreateFormControl(tmpName, acLabel, acDetail, , , 300, y, 3000, 320)
    ctl.Caption = "Время приёма:"
    Set ctl = CreateFormControl(tmpName, acComboBox, acDetail, , , 3500, y, 2400, 320)
    ctl.Name = "ПолеВремя"
    ctl.ControlSource = "ВремяПриёма"
    ctl.RowSourceType = "Value List"
    ctl.RowSource = "09:00:00;09:30:00;10:00:00;10:30:00;11:00:00;11:30:00;14:00:00;14:30:00;15:00:00;15:30:00;16:00:00;16:30:00"
    ctl.Format = "hh:nn"
    ctl.LimitToList = True
    Set ctl = CreateFormControl(tmpName, acCommandButton, acDetail, , , 6100, y, 3600, 380)
    ctl.Name = "КнопкаСвободные"
    ctl.Caption = "Свободные талоны"
    ctl.OnClick = "[Event Procedure]"
    y = y + 640

    Set ctl = CreateFormControl(tmpName, acLabel, acDetail, , , 300, y, 3000, 320)
    ctl.Caption = "Статус записи:"
    Set ctl = CreateFormControl(tmpName, acComboBox, acDetail, , , 3500, y, 2400, 320)
    ctl.Name = "ПолеСтатус"
    ctl.ControlSource = "СтатусЗаписи"
    ctl.RowSourceType = "Value List"
    ctl.RowSource = "Запланирован;Проведён;Отменён;Неявка"
    ctl.DefaultValue = """Запланирован"""
    y = y + 600

    Set ctl = CreateFormControl(tmpName, acLabel, acDetail, , , 300, y, 3000, 320)
    ctl.Caption = "Жалобы (анамнез):"
    Set ctl = CreateFormControl(tmpName, acTextBox, acDetail, , , 3500, y, 8000, 1300)
    ctl.Name = "ПолеЖалобы"
    ctl.ControlSource = "Жалобы"
    ctl.EnterKeyBehavior = True
    y = y + 1500

    Set ctl = CreateFormControl(tmpName, acCommandButton, acDetail, , , 3500, y, 2600, 480)
    ctl.Name = "КнопкаСохранить"
    ctl.Caption = "Сохранить запись"
    ctl.OnClick = "[Event Procedure]"
    Set ctl = CreateFormControl(tmpName, acCommandButton, acDetail, , , 6300, y, 2600, 480)
    ctl.Name = "КнопкаНовая"
    ctl.Caption = "Новая запись"
    ctl.OnClick = "[Event Procedure]"
    Set ctl = CreateFormControl(tmpName, acCommandButton, acDetail, , , 9100, y, 2400, 480)
    ctl.Name = "КнопкаЗакрыть"
    ctl.Caption = "Закрыть"
    ctl.OnClick = "[Event Procedure]"

    frm.HasModule = True
    frm.Module.AddFromString FormCodeVisit()

    DoCmd.Close acForm, tmpName, acSaveYes
    DoCmd.Rename frmName, acForm, tmpName
    LogLine "  форма «Ф_ЗаписьНаПриём» (форма ввода) создана"
End Sub

Private Function FormCodeVisit() As String
    Dim s As String
    s = "Private Sub Form_BeforeUpdate(Cancel As Integer)" & vbCrLf
    s = s & "    ' КОНТРОЛЬ: время приёма не может пересекаться у одного врача" & vbCrLf
    s = s & "    If Not ВремяСвободно(Me!КодВрача, Me!ДатаПриёма, Me!ВремяПриёма, Me!КодЗаписи) Then" & vbCrLf
    s = s & "        Cancel = True" & vbCrLf
    s = s & "        MsgBox ""У этого врача уже есть запись на указанные дату и время!"" & vbCrLf & _" & vbCrLf
    s = s & "               ""Выберите свободный талон."", vbExclamation, ""Контроль расписания""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub ПолеВрач_AfterUpdate()" & vbCrLf
    s = s & "    ' Автоматическое формирование номера талона" & vbCrLf
    s = s & "    If Me.NewRecord Or Len(Nz(Me!НомерТалона, """")) = 0 Then" & vbCrLf
    s = s & "        Me!НомерТалона = СформироватьНомерТалона(Me!КодВрача, Nz(Me!ДатаПриёма, Date))" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub КнопкаСвободные_Click()" & vbCrLf
    s = s & "    DoCmd.OpenQuery ""Q01_СвободныеТалоны""" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub КнопкаСохранить_Click()" & vbCrLf
    s = s & "    If Me.Dirty Then DoCmd.RunCommand acCmdSaveRecord" & vbCrLf
    s = s & "    ЗаписатьВЖурнал Nz(gLogin, ""registrator""), ""Запись на приём"", Nz(Me!НомерТалона, """"), ""Успешно""" & vbCrLf
    s = s & "    MsgBox ""Запись сохранена. Номер талона: "" & Me!НомерТалона, vbInformation, ""Регистратура""" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub КнопкаНовая_Click()" & vbCrLf
    s = s & "    DoCmd.GoToRecord , , acNewRec" & vbCrLf
    s = s & "    Me!ДатаПриёма = Date" & vbCrLf
    s = s & "    Me!СтатусЗаписи = ""Запланирован""" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub КнопкаЗакрыть_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FormCodeVisit = s
End Function
'''

FORM_SUBS = '''
' ----------------------------------------------------------------------------
'  Подчинённые формы «Ф_Назначения» и «Ф_Диагнозы»
' ----------------------------------------------------------------------------
Private Sub СоздатьФормуНазначения()
    Dim frmName As String, tmpName As String
    Dim frm As Form
    Dim ctl As Object
    Dim w As Variant, cap As Variant, src As Variant
    Dim i As Integer, x As Long
    frmName = "Ф_Назначения"

    On Error Resume Next
    DoCmd.Close acForm, frmName, acSaveNo
    DoCmd.DeleteObject acForm, frmName
    Err.Clear
    On Error GoTo 0

    Set frm = Application.CreateForm()
    tmpName = frm.Name
    On Error Resume Next
    DoCmd.RunCommand acCmdFormHdrFtr
    On Error GoTo 0

    frm.RecordSource = "Q13_НазначенияПриёма"
    frm.Caption = "Назначения"
    frm.DefaultView = 2                  ' 2 = табличное представление
    frm.InsideWidth = 13200
    frm.Section(acDetail).Height = 340
    frm.Section(acHeader).Height = 400

    cap = Array("Услуга", "Кол-во", "Цена", "Сумма", "Кратность", "Выполнено")
    src = Array("Услуга", "Количество", "Цена", "Сумма", "Кратность", "Выполнено")
    w = Array(4400, 900, 1200, 1400, 2200, 700)
    x = 100
    For i = 0 To UBound(cap)
        Set ctl = CreateFormControl(tmpName, acLabel, acHeader, , , x, 60, CLng(w(i)), 280)
        ctl.Caption = CStr(cap(i))
        ctl.FontBold = True
        Set ctl = CreateFormControl(tmpName, acTextBox, acDetail, , , x, 40, CLng(w(i)), 300)
        ctl.ControlSource = "[" & CStr(src(i)) & "]"
        ctl.Name = "ПолеН" & CStr(i + 1)
        x = x + CLng(w(i)) + 40
    Next i

    DoCmd.Close acForm, tmpName, acSaveYes
    DoCmd.Rename frmName, acForm, tmpName
    LogLine "  подчинённая форма «Ф_Назначения» создана"
End Sub

Private Sub СоздатьФормуДиагнозы()
    Dim frmName As String, tmpName As String
    Dim frm As Form
    Dim ctl As Object
    Dim cap As Variant, w As Variant
    Dim i As Integer, x As Long
    frmName = "Ф_Диагнозы"

    On Error Resume Next
    DoCmd.Close acForm, frmName, acSaveNo
    DoCmd.DeleteObject acForm, frmName
    Err.Clear
    On Error GoTo 0

    Set frm = Application.CreateForm()
    tmpName = frm.Name
    On Error Resume Next
    DoCmd.RunCommand acCmdFormHdrFtr
    On Error GoTo 0

    frm.RecordSource = "Q12_ДиагнозыПриёма"
    frm.Caption = "Диагнозы"
    frm.DefaultView = 2
    frm.InsideWidth = 13200
    frm.Section(acDetail).Height = 340
    frm.Section(acHeader).Height = 400

    cap = Array("Код МКБ", "Наименование", "Тип", "Дата", "Примечание")
    w = Array(1500, 5200, 1800, 1600, 2600)
    x = 100
    For i = 0 To UBound(cap)
        Set ctl = CreateFormControl(tmpName, acLabel, acHeader, , , x, 60, CLng(w(i)), 280)
        ctl.Caption = CStr(cap(i))
        ctl.FontBold = True
        If i = 0 Then
            Set ctl = CreateFormControl(tmpName, acComboBox, acDetail, , , x, 40, CLng(w(i)), 300)
            ctl.RowSourceType = "Table/Query"
            ctl.RowSource = "SELECT МКБ.КодМКБ, МКБ.Наименование FROM МКБ ORDER BY МКБ.КодМКБ;"
            ctl.ColumnCount = 2
            ctl.ColumnWidths = "1400;6000"
            ctl.BoundColumn = 1
            ctl.ListWidth = 7400
        Else
            Set ctl = CreateFormControl(tmpName, acTextBox, acDetail, , , x, 40, CLng(w(i)), 300)
        End If
        ctl.ControlSource = "[" & CStr(cap(i)) & "]"
        ctl.Name = "ПолеД" & CStr(i + 1)
        x = x + CLng(w(i)) + 40
    Next i

    DoCmd.Close acForm, tmpName, acSaveYes
    DoCmd.Rename frmName, acForm, tmpName
    LogLine "  подчинённая форма «Ф_Диагнозы» создана"
End Sub
'''

FORM_MAIN_VISIT = '''
' ----------------------------------------------------------------------------
'  Главная форма «Ф_ПриёмПациента» с двумя подчинёнными формами
' ----------------------------------------------------------------------------
Private Sub СоздатьФормуПриём()
    Dim frmName As String, tmpName As String
    Dim frm As Form
    Dim ctl As Object
    Dim y As Long
    frmName = "Ф_ПриёмПациента"

    On Error Resume Next
    DoCmd.Close acForm, frmName, acSaveNo
    DoCmd.DeleteObject acForm, frmName
    Err.Clear
    On Error GoTo 0

    Set frm = Application.CreateForm()
    tmpName = frm.Name
    On Error Resume Next
    DoCmd.RunCommand acCmdFormHdrFtr
    On Error GoTo 0

    frm.RecordSource = "ЗаписиНаПриём"
    frm.Caption = "Приём пациента"
    frm.DefaultView = 0
    frm.AllowDeletions = False
    frm.InsideWidth = 15000
    frm.Section(acDetail).Height = 7600
    frm.Section(acHeader).Height = 800
    frm.Section(acFooter).Height = 700

    Set ctl = CreateFormControl(tmpName, acLabel, acHeader, , , 300, 150, 9000, 500)
    ctl.Caption = "КАРТА ПРИЁМА ПАЦИЕНТА"
    ctl.FontName = "Tahoma": ctl.FontSize = 16: ctl.FontBold = True

    y = 200
    Set ctl = CreateFormControl(tmpName, acLabel, acDetail, , , 300, y, 2800, 320)
    ctl.Caption = "Пациент:"
    Set ctl = CreateFormControl(tmpName, acComboBox, acDetail, , , 3200, y, 6400, 320)
    ctl.Name = "ПолеПациент"
    ctl.ControlSource = "КодПациента"
    ctl.RowSourceType = "Table/Query"
    ctl.RowSource = "SELECT Пациенты.КодПациента, [Пациенты].[Фамилия] & "" "" & [Пациенты].[Имя] & "" "" & [Пациенты].[Отчество] AS ФИО FROM Пациенты ORDER BY Пациенты.Фамилия;"
    ctl.ColumnCount = 2
    ctl.ColumnWidths = "0;6200"
    ctl.BoundColumn = 1
    y = y + 550

    Set ctl = CreateFormControl(tmpName, acLabel, acDetail, , , 300, y, 2800, 320)
    ctl.Caption = "Врач:"
    Set ctl = CreateFormControl(tmpName, acComboBox, acDetail, , , 3200, y, 6400, 320)
    ctl.Name = "ПолеВрач"
    ctl.ControlSource = "КодВрача"
    ctl.RowSourceType = "Table/Query"
    ctl.RowSource = "SELECT Врачи.КодВрача, [Врачи].[Фамилия] & "" "" & [Врачи].[Имя] & "" "" & [Врачи].[Отчество] AS ФИО, Специальности.НазваниеСпециальности FROM Врачи INNER JOIN Специальности ON Врачи.КодСпециальности = Специальности.КодСпециальности ORDER BY Врачи.Фамилия;"
    ctl.ColumnCount = 3
    ctl.ColumnWidths = "0;3800;2500"
    ctl.BoundColumn = 1
    y = y + 550

    Set ctl = CreateFormControl(tmpName, acLabel, acDetail, , , 300, y, 2800, 320)
    ctl.Caption = "Дата и время приёма:"
    Set ctl = CreateFormControl(tmpName, acTextBox, acDetail, , , 3200, y, 2000, 320)
    ctl.ControlSource = "ДатаПриёма": ctl.Name = "ПолеДата": ctl.Format = "dd.mm.yyyy"
    Set ctl = CreateFormControl(tmpName, acTextBox, acDetail, , , 5400, y, 1600, 320)
    ctl.ControlSource = "ВремяПриёма": ctl.Name = "ПолеВремя": ctl.Format = "hh:nn"
    y = y + 550

    Set ctl = CreateFormControl(tmpName, acLabel, acDetail, , , 300, y, 2800, 320)
    ctl.Caption = "Статус:"
    Set ctl = CreateFormControl(tmpName, acComboBox, acDetail, , , 3200, y, 2400, 320)
    ctl.ControlSource = "СтатусЗаписи": ctl.Name = "ПолеСтатус"
    ctl.RowSourceType = "Value List"
    ctl.RowSource = "Запланирован;Проведён;Отменён;Неявка"
    Set ctl = CreateFormControl(tmpName, acLabel, acDetail, , , 6000, y, 1800, 320)
    ctl.Caption = "Номер талона:"
    Set ctl = CreateFormControl(tmpName, acTextBox, acDetail, , , 7900, y, 2600, 320)
    ctl.ControlSource = "НомерТалона": ctl.Name = "ПолеТалон"
    y = y + 550

    Set ctl = CreateFormControl(tmpName, acLabel, acDetail, , , 300, y, 2800, 320)
    ctl.Caption = "Жалобы:"
    Set ctl = CreateFormControl(tmpName, acTextBox, acDetail, , , 3200, y, 11200, 1100)
    ctl.ControlSource = "Жалобы": ctl.Name = "ПолеЖалобы"
    ctl.EnterKeyBehavior = True
    y = y + 1250

    Set ctl = CreateFormControl(tmpName, acLabel, acDetail, , , 300, y, 5000, 300)
    ctl.Caption = "Назначенные услуги и препараты:"
    ctl.FontBold = True
    y = y + 340
    Set ctl = CreateFormControl(tmpName, acSubform, acDetail, , , 300, y, 11800, 1800)
    ctl.Name = "ПФ_Назначения"
    ctl.SourceObject = "Form.Ф_Назначения"
    ctl.LinkMasterFields = "КодЗаписи"
    ctl.LinkChildFields = "КодЗаписи"
    y = y + 1900

    Set ctl = CreateFormControl(tmpName, acLabel, acDetail, , , 300, y, 5000, 300)
    ctl.Caption = "Диагнозы (МКБ-10):"
    ctl.FontBold = True
    y = y + 340
    Set ctl = CreateFormControl(tmpName, acSubform, acDetail, , , 300, y, 11800, 1600)
    ctl.Name = "ПФ_Диагнозы"
    ctl.SourceObject = "Form.Ф_Диагнозы"
    ctl.LinkMasterFields = "КодЗаписи"
    ctl.LinkChildFields = "КодЗаписи"

    Set ctl = CreateFormControl(tmpName, acLabel, acFooter, , , 300, 120, 3400, 320)
    ctl.Caption = "Стоимость приёма, руб.:"
    Set ctl = CreateFormControl(tmpName, acTextBox, acFooter, , , 3800, 120, 2000, 320)
    ctl.Name = "ПолеИтог"
    ctl.ControlSource = "=DSum(""[Сумма]"", ""Q13_НазначенияПриёма"", ""КодЗаписи="" & Nz([КодЗаписи], 0))"
    ctl.Format = "# ##0.00"
    Set ctl = CreateFormControl(tmpName, acCommandButton, acFooter, , , 6400, 100, 2400, 460)
    ctl.Name = "КнопкаОбновить"
    ctl.Caption = "Обновить"
    ctl.OnClick = "[Event Procedure]"
    Set ctl = CreateFormControl(tmpName, acCommandButton, acFooter, , , 9000, 100, 2500, 460)
    ctl.Name = "КнопкаПечать"
    ctl.Caption = "Печать талона"
    ctl.OnClick = "[Event Procedure]"
    Set ctl = CreateFormControl(tmpName, acCommandButton, acFooter, , , 11700, 100, 1800, 460)
    ctl.Name = "КнопкаЗакрыть"
    ctl.Caption = "Закрыть"
    ctl.OnClick = "[Event Procedure]"

    frm.HasModule = True
    frm.Module.AddFromString FormCodePriem()

    DoCmd.Close acForm, tmpName, acSaveYes
    DoCmd.Rename frmName, acForm, tmpName
    LogLine "  главная форма «Ф_ПриёмПациента» с подчинёнными формами создана"
End Sub

Private Function FormCodePriem() As String
    Dim s As String
    s = "Private Sub Form_Current()" & vbCrLf
    s = s & "    On Error Resume Next" & vbCrLf
    s = s & "    Me!ПФ_Назначения.Requery" & vbCrLf
    s = s & "    Me!ПФ_Диагнозы.Requery" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub КнопкаОбновить_Click()" & vbCrLf
    s = s & "    Me.Requery" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub КнопкаПечать_Click()" & vbCrLf
    s = s & "    DoCmd.OpenReport ""О_ТалоныНаДату"", acViewPreview" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub КнопкаЗакрыть_Click()" & vbCrLf
    s = s & "    DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FormCodePriem = s
End Function
'''

FORM_LOGIN_MENU = '''
' ----------------------------------------------------------------------------
'  Форма авторизации «Ф_Вход» и кнопочная форма «Ф_Главная»
' ----------------------------------------------------------------------------
Private Sub СоздатьФормуВход()
    Dim frmName As String, tmpName As String
    Dim frm As Form
    Dim ctl As Object
    frmName = "Ф_Вход"

    On Error Resume Next
    DoCmd.Close acForm, frmName, acSaveNo
    DoCmd.DeleteObject acForm, frmName
    Err.Clear
    On Error GoTo 0

    Set frm = Application.CreateForm()
    tmpName = frm.Name
    On Error Resume Next
    DoCmd.RunCommand acCmdFormHdrFtr
    On Error GoTo 0

    frm.Caption = "Вход в систему"
    frm.DefaultView = 0
    frm.RecordSource = ""
    frm.NavigationButtons = False
    frm.DividingLines = False
    frm.InsideWidth = 8200
    frm.Section(acDetail).Height = 3200
    frm.Section(acHeader).Height = 700

    Set ctl = CreateFormControl(tmpName, acLabel, acHeader, , , 300, 150, 7400, 450)
    ctl.Caption = "БАЗА ДАННЫХ «ПОЛИКЛИНИКА»"
    ctl.FontName = "Tahoma": ctl.FontSize = 14: ctl.FontBold = True

    Set ctl = CreateFormControl(tmpName, acLabel, acDetail, , , 400, 400, 2600, 320)
    ctl.Caption = "Логин:"
    Set ctl = CreateFormControl(tmpName, acTextBox, acDetail, , , 3200, 400, 3600, 320)
    ctl.Name = "ПолеЛогин"
    Set ctl = CreateFormControl(tmpName, acLabel, acDetail, , , 400, 900, 2600, 320)
    ctl.Caption = "Пароль:"
    Set ctl = CreateFormControl(tmpName, acTextBox, acDetail, , , 3200, 900, 3600, 320)
    ctl.Name = "ПолеПароль"
    ctl.InputMask = "PASSWORD"
    Set ctl = CreateFormControl(tmpName, acCommandButton, acDetail, , , 3200, 1600, 3600, 500)
    ctl.Name = "КнопкаВход"
    ctl.Caption = "Войти"
    ctl.OnClick = "[Event Procedure]"

    frm.HasModule = True
    frm.Module.AddFromString FormCodeLogin()

    DoCmd.Close acForm, tmpName, acSaveYes
    DoCmd.Rename frmName, acForm, tmpName
    LogLine "  форма авторизации «Ф_Вход» создана"
End Sub

Private Function FormCodeLogin() As String
    Dim s As String
    s = "Private Sub КнопкаВход_Click()" & vbCrLf
    s = s & "    If ВойтиВСистему(Nz(Me!ПолеЛогин, """"), Nz(Me!ПолеПароль, """")) Then" & vbCrLf
    s = s & "        DoCmd.Close acForm, Me.Name" & vbCrLf
    s = s & "        DoCmd.OpenForm ""Ф_Главная""" & vbCrLf
    s = s & "    Else" & vbCrLf
    s = s & "        MsgBox ""Неверный логин или пароль!"", vbExclamation, ""Доступ запрещён""" & vbCrLf
    s = s & "    End If" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FormCodeLogin = s
End Function

Private Sub СоздатьФормуГлавная()
    Dim frmName As String, tmpName As String
    Dim frm As Form
    Dim ctl As Object
    Dim captions As Variant
    Dim i As Integer, y As Long
    frmName = "Ф_Главная"

    On Error Resume Next
    DoCmd.Close acForm, frmName, acSaveNo
    DoCmd.DeleteObject acForm, frmName
    Err.Clear
    On Error GoTo 0

    Set frm = Application.CreateForm()
    tmpName = frm.Name
    On Error Resume Next
    DoCmd.RunCommand acCmdFormHdrFtr
    On Error GoTo 0

    frm.Caption = "Поликлиника — главная панель"
    frm.RecordSource = ""
    frm.NavigationButtons = False
    frm.InsideWidth = 9000
    frm.Section(acDetail).Height = 5400
    frm.Section(acHeader).Height = 800

    Set ctl = CreateFormControl(tmpName, acLabel, acHeader, , , 300, 150, 8000, 500)
    ctl.Caption = "ПОЛИКЛИНИКА: РАБОЧЕЕ МЕСТО"
    ctl.FontName = "Tahoma": ctl.FontSize = 16: ctl.FontBold = True

    captions = Array("Запись пациента на приём", "Приём пациента (врач)", _
                     "Картотека пациентов", "Отчёт «Загрузка врачей»", _
                     "Отчёт «История посещений»", "Отчёт «Стоимость лечения»", _
                     "Сервис: резервная копия", "Выход")
    y = 300
    For i = 0 To UBound(captions)
        Set ctl = CreateFormControl(tmpName, acCommandButton, acDetail, , , 600, y, 6800, 520)
        ctl.Name = "Кнопка" & CStr(i + 1)
        ctl.Caption = CStr(captions(i))
        ctl.OnClick = "[Event Procedure]"
        y = y + 600
    Next i

    frm.HasModule = True
    frm.Module.AddFromString FormCodeMenu()

    DoCmd.Close acForm, tmpName, acSaveYes
    DoCmd.Rename frmName, acForm, tmpName
    LogLine "  кнопочная форма «Ф_Главная» создана"
End Sub

Private Function FormCodeMenu() As String
    Dim s As String
    s = "Private Sub Form_Open(Cancel As Integer)" & vbCrLf
    s = s & "    Me.Caption = ""Поликлиника — "" & Nz(gRole, ""гость"") & "" ("" & Nz(gLogin, """") & "")""" & vbCrLf
    s = s & "    Me!Кнопка2.Enabled = (Nz(gLevel, 0) >= 2)" & vbCrLf
    s = s & "    Me!Кнопка7.Enabled = (Nz(gLevel, 0) >= 3)" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub Кнопка1_Click()" & vbCrLf
    s = s & "    DoCmd.OpenForm ""Ф_ЗаписьНаПриём"", , , , acFormAdd" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub Кнопка2_Click()" & vbCrLf
    s = s & "    DoCmd.OpenForm ""Ф_ПриёмПациента""" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub Кнопка3_Click()" & vbCrLf
    s = s & "    DoCmd.OpenForm ""Ф_Пациенты""" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub Кнопка4_Click()" & vbCrLf
    s = s & "    DoCmd.OpenReport ""О_ЗагрузкаВрачей"", acViewPreview" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub Кнопка5_Click()" & vbCrLf
    s = s & "    DoCmd.OpenReport ""О_ИсторияПосещений"", acViewPreview" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub Кнопка6_Click()" & vbCrLf
    s = s & "    DoCmd.OpenReport ""О_СтоимостьЛечения"", acViewPreview" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub Кнопка7_Click()" & vbCrLf
    s = s & "    СоздатьРезервнуюКопию" & vbCrLf
    s = s & "End Sub" & vbCrLf & vbCrLf
    s = s & "Private Sub Кнопка8_Click()" & vbCrLf
    s = s & "    ЗаписатьВЖурнал Nz(gLogin, ""гость""), ""Выход из системы"", ""Сеанс"", ""Успешно""" & vbCrLf
    s = s & "    Application.Quit acQuitSaveNone" & vbCrLf
    s = s & "End Sub" & vbCrLf
    FormCodeMenu = s
End Function
'''


def part_forms():
    return [FORM_PATIENTS.strip(), FORM_VISIT.strip(), FORM_SUBS.strip(),
            FORM_MAIN_VISIT.strip(), FORM_LOGIN_MENU.strip(),
            """Private Sub СоздатьФормы()
    СоздатьФормуБезопасно "Ф_Назначения"
    СоздатьФормуБезопасно "Ф_Диагнозы"
    СоздатьФормуБезопасно "Ф_Пациенты"
    СоздатьФормуБезопасно "Ф_ЗаписьНаПриём"
    СоздатьФормуБезопасно "Ф_ПриёмПациента"
    СоздатьФормуБезопасно "Ф_Вход"
    СоздатьФормуБезопасно "Ф_Главная"
    LogLine "  форм создано: " & CStr(mOkForm)
End Sub

Private Sub СоздатьФормуБезопасно(ByVal ИмяФормы As String)
    ' Каждая форма создаётся независимо: сбой одной формы не прерывает сборку.
    On Error Resume Next
    Select Case ИмяФормы
        Case "Ф_Назначения":     СоздатьФормуНазначения
        Case "Ф_Диагнозы":       СоздатьФормуДиагнозы
        Case "Ф_Пациенты":       СоздатьФормуПациенты
        Case "Ф_ЗаписьНаПриём":  СоздатьФормуЗапись
        Case "Ф_ПриёмПациента":  СоздатьФормуПриём
        Case "Ф_Вход":           СоздатьФормуВход
        Case "Ф_Главная":        СоздатьФормуГлавная
    End Select
    If Err.Number <> 0 Then
        mErr = mErr + 1
        mLog = mLog & "ОШИБКА при создании формы «" & ИмяФормы & "»: " & _
               Err.Description & vbCrLf
        Err.Clear
        УбратьВременныеОбъекты
    Else
        mOk = mOk + 1
        mOkForm = mOkForm + 1
    End If
    On Error GoTo 0
End Sub

Private Sub УбратьВременныеОбъекты()
    ' Закрывает и удаляет незавершённые (временные) формы и отчёты.
    Dim ao As AccessObject
    On Error Resume Next
    For Each ao In CurrentProject.AllForms
        If Left(ao.Name, 2) <> "Ф_" Then
            DoCmd.Close acForm, ao.Name, acSaveNo
            DoCmd.DeleteObject acForm, ao.Name
        End If
    Next ao
    For Each ao In CurrentProject.AllReports
        If Left(ao.Name, 2) <> "О_" Then
            DoCmd.Close acReport, ao.Name, acSaveNo
            DoCmd.DeleteObject acReport, ao.Name
        End If
    Next ao
    On Error GoTo 0
End Sub"""]


REPORTS_CODE = '''Private Sub СоздатьОтчёты()
    Dim fldsЗагрузка As Variant, fldsИстория As Variant
    Dim fldsТалоны As Variant, fldsСтоимость As Variant
    fldsЗагрузка = Array("Специальность", "Врач", "Всего талонов", "Принято", "Неявка", "Отменено")
    fldsИстория = Array("Пациент", "№ талона", "Дата", "Время", "Врач", "Специальность", "Статус")
    fldsТалоны = Array("№ талона", "Дата", "Время", "Пациент", "Возраст", "Врач", "Специальность", "Статус")
    fldsСтоимость = Array("№ талона", "Пациент", "Врач", "Дата", "Услуг назначено", "Сумма руб")
    On Error Resume Next
    СоздатьОтчётЗагрузка fldsЗагрузка
    ОтметитьОтчёт "О_ЗагрузкаВрачей"
    СоздатьОтчёт "О_ИсторияПосещений", "ИСТОРИЯ ПОСЕЩЕНИЙ ПАЦИЕНТА", _
                 "Q02_ИсторияПосещенийПациента", "Пациент", fldsИстория, False
    ОтметитьОтчёт "О_ИсторияПосещений"
    СоздатьОтчёт "О_ТалоныНаДату", "ЖУРНАЛ ТАЛОНОВ НА ДАТУ", _
                 "Q03_ЖурналЗаписей", "Дата", fldsТалоны, False
    ОтметитьОтчёт "О_ТалоныНаДату"
    СоздатьОтчёт "О_СтоимостьЛечения", "СТОИМОСТЬ ЛЕЧЕНИЯ ПО ПРИЁМАМ", _
                 "Q07_СтоимостьЛечения", "Врач", fldsСтоимость, True
    ОтметитьОтчёт "О_СтоимостьЛечения"
    On Error GoTo 0
    LogLine "  отчётов создано: " & CStr(mOkReport)
End Sub

Private Sub ОтметитьОтчёт(ByVal ИмяОтчёта As String)
    ' Фиксирует успех или сбой создания очередного отчёта.
    If Err.Number <> 0 Then
        mErr = mErr + 1
        mLog = mLog & "ОШИБКА при создании отчёта «" & ИмяОтчёта & "»: " & _
               Err.Description & vbCrLf
        Err.Clear
        УбратьВременныеОбъекты
    Else
        mOk = mOk + 1
        mOkReport = mOkReport + 1
    End If
End Sub

Private Sub СоздатьОтчётЗагрузка(ByVal flds As Variant)
    Dim rptName As String, tmpName As String
    Dim rpt As Report
    Dim ctl As Object
    Dim w As Variant
    Dim i As Integer, x As Long, grpLvl As Variant
    rptName = "О_ЗагрузкаВрачей"

    On Error Resume Next
    DoCmd.Close acReport, rptName, acSaveNo
    DoCmd.DeleteObject acReport, rptName
    Err.Clear
    On Error GoTo 0

    Set rpt = Application.CreateReport()
    tmpName = rpt.Name
    On Error Resume Next
    DoCmd.RunCommand acCmdReportHdrFtr
    On Error GoTo 0

    rpt.RecordSource = "Q06_ЗагрузкаВрачейЗаПериод"
    rpt.Caption = "Загрузка врачей по специальностям"
    rpt.Width = 14400
    rpt.Section(acHeader).Height = 1100
    rpt.Section(acPageHeader).Height = 480
    rpt.Section(acDetail).Height = 420
    rpt.Section(acFooter).Height = 700
    rpt.Section(acPageFooter).Height = 460

    Set ctl = CreateReportControl(tmpName, acLabel, acHeader, , , 300, 100, 12000, 520)
    ctl.Caption = "ЗАГРУЗКА ВРАЧЕЙ ПО СПЕЦИАЛЬНОСТЯМ"
    ctl.FontName = "Tahoma": ctl.FontSize = 16: ctl.FontBold = True
    Set ctl = CreateReportControl(tmpName, acLabel, acHeader, , , 300, 640, 12000, 320)
    ctl.Caption = "Отчёт сформирован: " & Format(Date, "dd.mm.yyyy") & " г."

    grpLvl = Application.CreateGroupLevel(tmpName, "Специальность", True, True)
    rpt.Section(acGroupLevel1Header).Height = 480
    rpt.Section(acGroupLevel1Footer).Height = 520
    Set ctl = CreateReportControl(tmpName, acTextBox, acGroupLevel1Header, , , 300, 60, 6400, 360)
    ctl.ControlSource = "Специальность"
    ctl.Name = "ПолеГруппа"
    ctl.FontBold = True: ctl.FontSize = 12

    w = Array(4200, 2200, 2000, 1600, 1600)
    x = 600
    For i = 1 To UBound(flds)
        Set ctl = CreateReportControl(tmpName, acLabel, acPageHeader, , , x, 120, CLng(w(i - 1)), 320)
        ctl.Caption = CStr(flds(i))
        ctl.FontBold = True
        Set ctl = CreateReportControl(tmpName, acTextBox, acDetail, , , x, 60, CLng(w(i - 1)), 340)
        ctl.ControlSource = "[" & CStr(flds(i)) & "]"
        ctl.Name = "ПолеД" & CStr(i)
        x = x + CLng(w(i - 1)) + 60
    Next i

    Set ctl = CreateReportControl(tmpName, acLabel, acGroupLevel1Footer, , , 800, 80, 4200, 340)
    ctl.Caption = "Итого по специальности:"
    Set ctl = CreateReportControl(tmpName, acTextBox, acGroupLevel1Footer, , , 5000, 80, 2400, 340)
    ctl.ControlSource = "=Sum([Принято])"
    ctl.Name = "ИтогПринято"
    ctl.FontBold = True
    Set ctl = CreateReportControl(tmpName, acTextBox, acGroupLevel1Footer, , , 7600, 80, 2400, 340)
    ctl.ControlSource = "=Sum([Всего талонов])"
    ctl.Name = "ИтогТалонов"
    ctl.FontBold = True

    Set ctl = CreateReportControl(tmpName, acLabel, acFooter, , , 800, 100, 4200, 340)
    ctl.Caption = "ВСЕГО ПО ПОЛИКЛИНИКЕ:"
    Set ctl = CreateReportControl(tmpName, acTextBox, acFooter, , , 5000, 100, 2400, 340)
    ctl.ControlSource = "=Sum([Принято])"
    ctl.Name = "ВсегоПринято"
    ctl.FontBold = True
    Set ctl = CreateReportControl(tmpName, acTextBox, acFooter, , , 7600, 100, 2400, 340)
    ctl.ControlSource = "=Sum([Всего талонов])"
    ctl.Name = "ВсегоТалонов"
    ctl.FontBold = True

    Set ctl = CreateReportControl(tmpName, acTextBox, acPageFooter, , , 11000, 80, 2400, 320)
    ctl.ControlSource = "=""Стр. "" & [Page] & "" из "" & [Pages]"
    ctl.Name = "НумерацияСтраниц"

    DoCmd.Close acReport, tmpName, acSaveYes
    DoCmd.Rename rptName, acReport, tmpName
    LogLine "  отчёт «О_ЗагрузкаВрачей» создан"
End Sub

Private Sub СоздатьОтчёт(ByVal rptName As String, ByVal Заголовок As String, _
                         ByVal Источник As String, ByVal ПолеГруппы As String, _
                         ByVal flds As Variant, ByVal ЕстьСумма As Boolean)
    Dim tmpName As String
    Dim rpt As Report
    Dim ctl As Object
    Dim i As Integer, x As Long, w As Long, grpLvl As Variant

    On Error Resume Next
    DoCmd.Close acReport, rptName, acSaveNo
    DoCmd.DeleteObject acReport, rptName
    Err.Clear
    On Error GoTo 0

    Set rpt = Application.CreateReport()
    tmpName = rpt.Name
    On Error Resume Next
    DoCmd.RunCommand acCmdReportHdrFtr
    On Error GoTo 0

    rpt.RecordSource = Источник
    rpt.Caption = Заголовок
    rpt.Width = 14400
    rpt.Section(acHeader).Height = 900
    rpt.Section(acPageHeader).Height = 480
    rpt.Section(acDetail).Height = 420
    rpt.Section(acFooter).Height = 700
    rpt.Section(acPageFooter).Height = 460

    Set ctl = CreateReportControl(tmpName, acLabel, acHeader, , , 300, 100, 12000, 520)
    ctl.Caption = Заголовок
    ctl.FontName = "Tahoma": ctl.FontSize = 16: ctl.FontBold = True

    grpLvl = Application.CreateGroupLevel(tmpName, ПолеГруппы, True, True)
    rpt.Section(acGroupLevel1Header).Height = 480
    rpt.Section(acGroupLevel1Footer).Height = 520
    Set ctl = CreateReportControl(tmpName, acTextBox, acGroupLevel1Header, , , 300, 60, 6400, 360)
    ctl.ControlSource = "[" & ПолеГруппы & "]"
    ctl.Name = "ПолеГруппа"
    ctl.FontBold = True: ctl.FontSize = 12

    w = 2000
    x = 600
    For i = 0 To UBound(flds)
        Set ctl = CreateReportControl(tmpName, acLabel, acPageHeader, , , x, 120, w, 320)
        ctl.Caption = CStr(flds(i))
        ctl.FontBold = True
        Set ctl = CreateReportControl(tmpName, acTextBox, acDetail, , , x, 60, w, 340)
        ctl.ControlSource = "[" & CStr(flds(i)) & "]"
        ctl.Name = "ПолеД" & CStr(i + 1)
        x = x + w + 40
        If x > 13000 Then Exit For
    Next i

    Set ctl = CreateReportControl(tmpName, acLabel, acGroupLevel1Footer, , , 600, 80, 4200, 340)
    ctl.Caption = "Итого по группе:"
    Set ctl = CreateReportControl(tmpName, acTextBox, acGroupLevel1Footer, , , 5000, 80, 2400, 340)
    If ЕстьСумма Then
        ctl.ControlSource = "=Sum([Сумма руб])"
    Else
        ctl.ControlSource = "=Count(*)"
    End If
    ctl.Name = "ИтогГруппы"
    ctl.FontBold = True

    Set ctl = CreateReportControl(tmpName, acLabel, acFooter, , , 600, 100, 4200, 340)
    ctl.Caption = "ВСЕГО ПО ОТЧЁТУ:"
    Set ctl = CreateReportControl(tmpName, acTextBox, acFooter, , , 5000, 100, 2400, 340)
    If ЕстьСумма Then
        ctl.ControlSource = "=Sum([Сумма руб])"
    Else
        ctl.ControlSource = "=Count(*)"
    End If
    ctl.Name = "ИтогОтчёта"
    ctl.FontBold = True

    Set ctl = CreateReportControl(tmpName, acTextBox, acPageFooter, , , 11000, 80, 2400, 320)
    ctl.ControlSource = "=""Стр. "" & [Page] & "" из "" & [Pages]"
    ctl.Name = "НумерацияСтраниц"

    DoCmd.Close acReport, tmpName, acSaveYes
    DoCmd.Rename rptName, acReport, tmpName
    LogLine "  отчёт «" & rptName & "» создан"
End Sub
'''

MACROS_CODE = '''Private Sub СоздатьМакросы()
    Dim p As String
    Dim names As Variant
    Dim i As Integer
    Dim okCount As Integer
    p = CurrentProject.Path & "\\Макросы\\"
    names = Array("AutoExec", "М_ОткрытьПриём", "М_ОбновитьСтатусы", _
                  "М_ЭкспортОтчётаExcel", "М_РезервнаяКопия", "М_ЗакрытьПриложение")
    okCount = 0
    For i = 0 To UBound(names)
        If Dir(p & CStr(names(i)) & ".txt") <> "" Then
            On Error Resume Next
            DoCmd.DeleteObject acMacro, CStr(names(i))
            Err.Clear
            Application.LoadFromText acMacro, CStr(names(i)), p & CStr(names(i)) & ".txt"
            If Err.Number = 0 Then
                okCount = okCount + 1
            Else
                mLog = mLog & "Макрос " & CStr(names(i)) & _
                       " не загружен из файла (создайте вручную по прил. Г)." & vbCrLf
                Err.Clear
            End If
            On Error GoTo 0
        Else
            mLog = mLog & "Файл дампа макроса не найден: " & p & CStr(names(i)) & ".txt" & vbCrLf
        End If
    Next i
    mOkMacro = okCount
    LogLine "  макросов загружено: " & CStr(okCount) & " из " & CStr(UBound(names) + 1)
End Sub
'''

STARTUP_CODE = '''Private Sub НастроитьЗапуск()
    Dim db As DAO.Database
    Set db = CurrentDb
    On Error Resume Next
    AddProp db, "StartupForm", dbText, "Ф_Вход"
    AddProp db, "AppTitle", dbText, "Поликлиника — база данных (Сагадиев А. Р., гр. Р-25-29)"
    On Error GoTo 0
    Application.RefreshTitleBar
    LogLine "  параметры запуска настроены (стартовая форма «Ф_Вход»)"
End Sub
'''

MAIN_CODE = '''Public Sub СоздатьБД(Optional ByVal БезДиалогов As Boolean = False)
    Dim t0 As Double
    On Error GoTo КритическаяОшибка
    t0 = Timer
    mOk = 0
    mErr = 0
    mOkForm = 0
    mOkReport = 0
    mOkQuery = 0
    mOkMacro = 0
    mLog = ""

    If Not БезДиалогов Then
        If MsgBox("Будет создана база данных «Поликлиника»." & vbCrLf & _
                  "Существующие таблицы, запросы, формы, отчёты и макросы с теми же " & _
                  "именами будут удалены!" & vbCrLf & vbCrLf & "Продолжить?", _
                  vbQuestion + vbYesNo, "Создание БД «Поликлиника»") = vbNo Then Exit Sub
    End If

    LogLine "=== НАЧАЛО ПОСТРОЕНИЯ: " & Now() & " ==="
    LogLine "Этап 1/10: очистка предыдущей версии базы данных"
    ОчиститьБД
    LogLine "Этап 2/10: создание таблиц"
    СоздатьТаблицы
    LogLine "Этап 3/10: установка свойств полей (маски ввода, подписи, проверки)"
    УстановитьСвойстваПолей
    LogLine "Этап 4/10: создание связей (внешних ключей)"
    СоздатьСвязи
    LogLine "Этап 5/10: создание индексов"
    СоздатьИндексы
    LogLine "Этап 6/10: заполнение таблиц тестовыми данными"
    ЗаполнитьДанные
    LogLine "Этап 7/10: создание запросов"
    СоздатьЗапросы
    LogLine "Этап 8/10: создание форм"
    СоздатьФормы
    LogLine "Этап 9/10: создание отчётов"
    СоздатьОтчёты
    LogLine "Этап 10/10: загрузка макросов и настройка запуска"
    СоздатьМакросы
    НастроитьЗапуск

    LogLine "=== ОКОНЧАНИЕ: " & Now() & " ==="
    LogLine "Успешно выполнено операторов SQL: " & CStr(mOk)
    LogLine "Ошибок: " & CStr(mErr)

    If Not БезДиалогов Then
        MsgBox "БАЗА ДАННЫХ «ПОЛИКЛИНИКА» СОЗДАНА!" & vbCrLf & vbCrLf & _
               "Успешно выполнено операторов SQL: " & CStr(mOk) & vbCrLf & _
               "Ошибок: " & CStr(mErr) & vbCrLf & _
               "Время построения, с: " & Format(Timer - t0, "0.0") & vbCrLf & vbCrLf & _
               "Таблиц: 13, запросов: " & CStr(mOkQuery) & ", форм: " & CStr(mOkForm) & _
               ", отчётов: " & CStr(mOkReport) & ", макросов: " & CStr(mOkMacro) & ".", _
               vbInformation, "Готово"

        If mErr > 0 Then
            If MsgBox("Построение завершено, но с замечаниями (ошибок: " & CStr(mErr) & ")." & vbCrLf & _
                      "Показать журнал построения? Он также доступен в окне отладки VBA (Ctrl+G).", _
                      vbYesNo + vbExclamation, "Построение завершено") = vbYes Then
                MsgBox mLog, vbInformation, "Журнал построения"
            End If
        End If
    End If
    Exit Sub
КритическаяОшибка:
    mErr = mErr + 1
    mLog = mLog & vbCrLf & "КРИТИЧЕСКАЯ ОШИБКА: " & Err.Description & vbCrLf & _
           "  Последний выполненный этап указан в журнале выше." & vbCrLf
    If Not БезДиалогов Then
        MsgBox "Построение прервано из-за ошибки:" & vbCrLf & vbCrLf & _
               Err.Description & vbCrLf & vbCrLf & mLog, vbCritical, "Ошибка построения"
    End If
End Sub

Public Sub ПоказатьСтатистику()
    Dim s As String
    s = "Состав базы данных «Поликлиника»:" & vbCrLf & vbCrLf
    s = s & "Специальностей:    " & DCount("*", "Специальности") & vbCrLf
    s = s & "Врачей:            " & DCount("*", "Врачи") & vbCrLf
    s = s & "Пациентов:         " & DCount("*", "Пациенты") & vbCrLf
    s = s & "Услуг:             " & DCount("*", "Услуги") & vbCrLf
    s = s & "Слотов расписания: " & DCount("*", "РасписаниеПриёма") & vbCrLf
    s = s & "Записей на приём:  " & DCount("*", "ЗаписиНаПриём") & vbCrLf
    s = s & "Диагнозов:         " & DCount("*", "Диагнозы") & vbCrLf
    s = s & "Назначений:        " & DCount("*", "Назначения") & vbCrLf
    s = s & "Пользователей:     " & DCount("*", "Пользователи") & vbCrLf
    MsgBox s, vbInformation, "Статистика базы данных"
End Sub
'''

FOOTER = '''
' ============================================================================
'  СЛУЖЕБНЫЕ ПРОЦЕДУРЫ
' ============================================================================
Private Function DB() As DAO.Database
    Set DB = CurrentDb
End Function

Private Function Exec(ByVal sqlText As String) As Boolean
    On Error GoTo EH
    DB.Execute sqlText, dbFailOnError
    mOk = mOk + 1
    Exec = True
    Exit Function
EH:
    mErr = mErr + 1
    mLog = mLog & "ОШИБКА: " & Err.Description & vbCrLf & "  SQL: " & Left(sqlText, 200) & vbCrLf & vbCrLf
End Function

Private Function ExecSilent(ByVal sqlText As String) As Boolean
    On Error Resume Next
    DB.Execute sqlText, dbFailOnError
    ExecSilent = (Err.Number = 0)
    Err.Clear
End Function

Private Sub AddProp(ByRef obj As Object, ByVal pName As String, _
                    ByVal pType As Integer, ByVal pValue As Variant)
    Dim p As DAO.Property
    On Error Resume Next
    Set p = obj.Properties(pName)
    If Err.Number = 0 Then
        p.Value = pValue
    Else
        Err.Clear
        Set p = obj.CreateProperty(pName, pType, pValue)
        obj.Properties.Append p
    End If
    On Error GoTo 0
End Sub

Private Sub LogLine(ByVal s As String)
    Debug.Print s
    mLog = mLog & s & vbCrLf
End Sub

Private Sub ОчиститьБД()
    Dim db As DAO.Database
    Dim i As Long
    Dim ao As AccessObject
    Set db = CurrentDb
    On Error Resume Next
    For i = db.Relations.Count - 1 To 0 Step -1
        db.Relations.Delete db.Relations(i).Name
    Next i
    For i = db.QueryDefs.Count - 1 To 0 Step -1
        If Left(db.QueryDefs(i).Name, 4) <> "~sq_" Then db.QueryDefs.Delete db.QueryDefs(i).Name
    Next i
    For i = db.TableDefs.Count - 1 To 0 Step -1
        If Left(db.TableDefs(i).Name, 4) <> "MSys" Then db.TableDefs.Delete db.TableDefs(i).Name
    Next i
    For Each ao In CurrentProject.AllMacros
        DoCmd.DeleteObject acMacro, ao.Name
    Next ao
    For Each ao In CurrentProject.AllReports
        DoCmd.DeleteObject acReport, ao.Name
    Next ao
    For Each ao In CurrentProject.AllForms
        DoCmd.DeleteObject acForm, ao.Name
    Next ao
    On Error GoTo 0
End Sub

Public Sub ЗаписатьВЖурнал(ByVal Логин As String, ByVal Действие As String, _
                           ByVal Объект As String, ByVal Результат As String)
    Dim sqlText As String
    On Error Resume Next
    sqlText = "INSERT INTO ЖурналСобытий (ДатаВремя, Логин, Действие, ОбъектДействия, Результат) " & _
              "VALUES (Now(), '" & Replace(Логин, "'", "''") & "', '" & _
              Replace(Действие, "'", "''") & "', '" & Replace(Объект, "'", "''") & _
              "', '" & Результат & "')"
    DB.Execute sqlText, dbFailOnError
End Sub

Public Function СформироватьНомерТалона(ByVal КодВрача As Variant, _
                                        ByVal ДатаПриёма As Variant) As String
    Dim n As Long
    On Error Resume Next
    n = DCount("*", "ЗаписиНаПриём", "ДатаПриёма=#" & Format(ДатаПриёма, "yyyy\\-mm\\-dd") & "#")
    If Err.Number <> 0 Then n = 0
    Err.Clear
    On Error GoTo 0
    СформироватьНомерТалона = "Т-" & Format(Nz(ДатаПриёма, Date), "ddmmyy") & "-" & _
                              Format(Nz(КодВрача, 0), "00") & "-" & Format(n + 1, "000")
End Function

Public Function ВремяСвободно(ByVal КодВрача As Variant, ByVal ДатаПриёма As Variant, _
                              ByVal ВремяПриёма As Variant, ByVal ТекущаяЗапись As Variant) As Boolean
    Dim n As Long
    Dim cond As String
    If IsNull(КодВрача) Or IsNull(ДатаПриёма) Or IsNull(ВремяПриёма) Then
        ВремяСвободно = True
        Exit Function
    End If
    cond = "КодВрача=" & CLng(КодВрача) & _
           " AND ДатаПриёма=#" & Format(ДатаПриёма, "yyyy\\-mm\\-dd") & "#" & _
           " AND ВремяПриёма=#" & Format(ВремяПриёма, "hh:nn:ss") & "#" & _
           " AND КодЗаписи<>" & CLng(Nz(ТекущаяЗапись, 0))
    n = DCount("*", "ЗаписиНаПриём", cond)
    ВремяСвободно = (n = 0)
End Function
'''

MOD_SECURITY = '''Attribute VB_Name = "ModSecurity"
' ============================================================================
'  МОДУЛЬ РАЗГРАНИЧЕНИЯ ДОСТУПА, АУДИТА И АДМИНИСТРИРОВАНИЯ
'  База данных «Поликлиника» (вариант 14, Сагадиев А. Р., гр. Р-25-29)
' ============================================================================
Option Compare Database
Option Explicit

' Глобальные переменные сеанса (gLogin, gRole, gLevel, gKodVracha) объявлены
' в модуле AutoBuild.bas, чтобы не возникало конфликта имён при компиляции.

Public Function ВойтиВСистему(ByVal Логин As String, ByVal Пароль As String) As Boolean
    Dim rs As DAO.Recordset
    Dim sqlText As String
    ВойтиВСистему = False
    If Len(Логин) = 0 Or Len(Пароль) = 0 Then Exit Function

    sqlText = "SELECT Пользователи.Логин, Пользователи.ФИО, Роли.НазваниеРоли, " & _
              "Роли.УровеньДоступа, Пользователи.КодВрача " & _
              "FROM Пользователи INNER JOIN Роли ON Пользователи.КодРоли = Роли.КодРоли " & _
              "WHERE Пользователи.Логин = '" & Replace(Логин, "'", "''") & _
              "' AND Пользователи.Пароль = '" & Replace(Пароль, "'", "''") & _
              "' AND Пользователи.Активен = True"

    Set rs = CurrentDb.OpenRecordset(sqlText, dbOpenSnapshot)
    If Not rs.EOF Then
        gLogin = rs!Логин
        gRole = rs!НазваниеРоли
        gLevel = Nz(rs!УровеньДоступа, 1)
        gKodVracha = Nz(rs!КодВрача, 0)
        ВойтиВСистему = True
        ЗаписатьВЖурнал gLogin, "Вход в систему", "Сеанс", "Успешно"
    Else
        ЗаписатьВЖурнал Логин, "Вход в систему", "Сеанс", "Отказано"
    End If
    rs.Close
End Function

Public Sub ВыйтиИзСистемы()
    ЗаписатьВЖурнал Nz(gLogin, "гость"), "Выход из системы", "Сеанс", "Успешно"
    gLogin = "": gRole = "": gLevel = 0: gKodVracha = 0
End Sub

Public Function ИметьПраво(ByVal ТребуемыйУровень As Long) As Boolean
    ИметьПраво = (Nz(gLevel, 0) >= ТребуемыйУровень)
End Function

Public Sub ПроверитьПраво(ByVal ТребуемыйУровень As Long, ByVal Операция As String)
    If Nz(gLevel, 0) < ТребуемыйУровень Then
        ЗаписатьВЖурнал Nz(gLogin, "гость"), "Отказ в доступе", Операция, "Отказано"
        MsgBox "Недостаточно прав для операции «" & Операция & "»." & vbCrLf & _
               "Ваша роль: " & Nz(gRole, "гость") & _
               " (уровень " & Nz(gLevel, 0) & ").", vbExclamation, "Доступ запрещён"
        Err.Raise vbObjectError + 513, , "Доступ запрещён"
    End If
End Sub

Public Sub СменитьПароль(ByVal Логин As String, ByVal СтарыйПароль As String, _
                         ByVal НовыйПароль As String)
    Dim sqlText As String
    If Len(НовыйПароль) < 5 Then
        MsgBox "Пароль должен содержать не менее 5 символов.", vbExclamation
        Exit Sub
    End If
    sqlText = "UPDATE Пользователи SET Пользователи.Пароль = '" & _
              Replace(НовыйПароль, "'", "''") & "' WHERE Пользователи.Логин = '" & _
              Replace(Логин, "'", "''") & "' AND Пользователи.Пароль = '" & _
              Replace(СтарыйПароль, "'", "''") & "'"
    CurrentDb.Execute sqlText, dbFailOnError
    If CurrentDb.RecordsAffected = 0 Then
        MsgBox "Пароль не изменён: неверный логин или старый пароль.", vbExclamation
        ЗаписатьВЖурнал Логин, "Смена пароля", "Пользователи", "Отказано"
    Else
        MsgBox "Пароль успешно изменён.", vbInformation
        ЗаписатьВЖурнал Логин, "Смена пароля", "Пользователи", "Успешно"
    End If
End Sub

Public Sub БлокироватьПользователя(ByVal Логин As String, ByVal Заблокировать As Boolean)
    ПроверитьПраво 3, "Блокировка учётной записи"
    CurrentDb.Execute "UPDATE Пользователи SET Пользователи.Активен = " & _
                      IIf(Заблокировать, "False", "True") & _
                      " WHERE Пользователи.Логин = '" & Replace(Логин, "'", "''") & "'", dbFailOnError
    ЗаписатьВЖурнал Nz(gLogin, "admin"), IIf(Заблокировать, "Блокировка", "Разблокировка"), _
                    "Пользователи: " & Логин, "Успешно"
End Sub

Public Sub ПоказатьЖурнал(Optional ByVal Записей As Long = 50)
    Dim rs As DAO.Recordset
    Dim s As String
    Dim i As Long
    Set rs = CurrentDb.OpenRecordset("SELECT TOP " & Записей & " * FROM ЖурналСобытий " & _
                                     "ORDER BY ЖурналСобытий.ДатаВремя DESC", dbOpenSnapshot)
    s = ""
    i = 0
    Do Until rs.EOF
        i = i + 1
        s = s & Format(rs!ДатаВремя, "dd.mm.yyyy hh:nn") & "  " & rs!Логин & _
            "  " & rs!Действие & "  (" & Nz(rs!Результат, "") & ")" & vbCrLf
        rs.MoveNext
        If i >= 25 Then Exit Do
    Loop
    rs.Close
    MsgBox s, vbInformation, "Журнал событий (последние записи)"
End Sub
'''

MOD_ADMIN = '''Attribute VB_Name = "ModAdmin"
' ============================================================================
'  МОДУЛЬ АДМИНИСТРИРОВАНИЯ: РЕЗЕРВНОЕ КОПИРОВАНИЕ, ИМПОРТ/ЭКСПОРТ, СЕРВИС
'  База данных «Поликлиника» (вариант 14, Сагадиев А. Р., гр. Р-25-29)
' ============================================================================
Option Compare Database
Option Explicit

Private Function ПокаПапка(ByVal Имя As String) As String
    Dim p As String
    p = CurrentProject.Path & "\\" & Имя
    If Dir(p, vbDirectory) = "" Then MkDir p
    ПокаПапка = p & "\\"
End Function

Public Sub СоздатьРезервнуюКопию()
    Dim src As String, dst As String
    On Error GoTo EH
    src = CurrentDb.Name
    dst = ПокаПапка("Резервные_копии") & "Поликлиника_" & Format(Now(), "yyyy-mm-dd_hhnn") & ".accdb"
    FileCopy src, dst
    ЗаписатьВЖурнал Nz(gLogin, "admin"), "Резервное копирование", dst, "Успешно"
    MsgBox "Резервная копия создана:" & vbCrLf & dst, vbInformation, "Резервное копирование"
    Exit Sub
EH:
    ЗаписатьВЖурнал Nz(gLogin, "admin"), "Резервное копирование", src, "Ошибка"
    MsgBox "Не удалось создать резервную копию: " & Err.Description, vbCritical
End Sub

Public Sub ЭкспортВExcel(Optional ByVal ИмяОбъекта As String = "Q03_ЖурналЗаписей")
    Dim p As String
    On Error GoTo EH
    p = ПокаПапка("Экспорт") & ИмяОбъекта & "_" & Format(Date, "yyyy-mm-dd") & ".xlsx"
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, ИмяОбъекта, p, True
    ЗаписатьВЖурнал Nz(gLogin, "admin"), "Экспорт в Excel", ИмяОбъекта, "Успешно"
    MsgBox "Данные выгружены в файл:" & vbCrLf & p, vbInformation, "Экспорт"
    Exit Sub
EH:
    MsgBox "Ошибка экспорта: " & Err.Description, vbCritical
End Sub

Public Sub ЭкспортОтчётаВPDF(Optional ByVal ИмяОтчёта As String = "О_ЗагрузкаВрачей")
    Dim p As String
    On Error GoTo EH
    p = ПокаПапка("Экспорт") & ИмяОтчёта & "_" & Format(Date, "yyyy-mm-dd") & ".pdf"
    DoCmd.OutputTo acOutputReport, ИмяОтчёта, acFormatPDF, p, False
    ЗаписатьВЖурнал Nz(gLogin, "admin"), "Экспорт отчёта в PDF", ИмяОтчёта, "Успешно"
    MsgBox "Отчёт сохранён в файл:" & vbCrLf & p, vbInformation, "Экспорт PDF"
    Exit Sub
EH:
    MsgBox "Ошибка экспорта PDF: " & Err.Description, vbCritical
End Sub

Public Sub ИмпортИзExcel(ByVal ИмяТаблицы As String, Optional ByVal Путь As String = "")
    On Error GoTo EH
    If Len(Путь) = 0 Then
        With Application.FileDialog(1)
            .Title = "Выберите книгу Microsoft Excel"
            .Filters.Clear
            .Filters.Add "Книги Excel", "*.xlsx; *.xls"
            If .Show = -1 Then Путь = .SelectedItems(1) Else Exit Sub
        End With
    End If
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, ИмяТаблицы, Путь, True
    ЗаписатьВЖурнал Nz(gLogin, "admin"), "Импорт из Excel", ИмяТаблицы, "Успешно"
    MsgBox "Данные загружены в таблицу «" & ИмяТаблицы & "».", vbInformation, "Импорт"
    Exit Sub
EH:
    MsgBox "Ошибка импорта: " & Err.Description, vbCritical
End Sub

Public Sub СжатьИВосстановить()
    Dim s As String
    s = CurrentDb.Name
    ЗаписатьВЖурнал Nz(gLogin, "admin"), "Сжатие и восстановление БД", s, "Успешно"
    Application.SetOption "Auto Compact", True
    MsgBox "Автоматическое сжатие включено." & vbCrLf & _
           "Полное сжатие выполните командой:" & vbCrLf & _
           "Файл -> Сведения -> Сжать и восстановить базу данных.", vbInformation
End Sub

Public Sub Обслуживание()
    On Error GoTo EH
    CurrentDb.Execute "UPDATE ЗаписиНаПриём SET СтатусЗаписи = 'Проведён' " & _
                      "WHERE ДатаПриёма < Date() AND СтатусЗаписи = 'Запланирован'", dbFailOnError
    CurrentDb.Execute "INSERT INTO АрхивЗаписей (КодЗаписи, НомерТалона, КодПациента, " & _
                      "КодВрача, ДатаПриёма, ВремяПриёма, СтатусЗаписи, Жалобы, ДатаАрхивации) " & _
                      "SELECT КодЗаписи, НомерТалона, КодПациента, КодВрача, ДатаПриёма, " & _
                      "ВремяПриёма, СтатусЗаписи, Жалобы, Date() FROM ЗаписиНаПриём " & _
                      "WHERE ДатаПриёма < Date() AND СтатусЗаписи In ('Проведён','Неявка')", dbFailOnError
    CurrentDb.Execute "DELETE FROM ЗаписиНаПриём WHERE СтатусЗаписи = 'Отменён' " & _
                      "AND ДатаПриёма < Date() - 30", dbFailOnError
    ЗаписатьВЖурнал Nz(gLogin, "admin"), "Регламентное обслуживание", "Все таблицы", "Успешно"
    MsgBox "Регламентное обслуживание выполнено:" & vbCrLf & _
           "обновлены статусы, проведённые приёмы перенесены в архив, " & _
           "устаревшие отменённые записи удалены.", vbInformation
    Exit Sub
EH:
    ЗаписатьВЖурнал Nz(gLogin, "admin"), "Регламентное обслуживание", "Все таблицы", "Ошибка"
    MsgBox "Ошибка обслуживания: " & Err.Description, vbCritical
End Sub
'''


def macro_text(actions: list) -> str:
    """Формирует текстовый дамп макроса Access (для LoadFromText acMacro)."""
    lines = ["Version =196611", "ColumnsShown =0"]
    for act, args in actions:
        lines.append("Begin")
        lines.append('    Action ="%s"' % act)
        for a in args:
            lines.append('    Argument ="%s"' % a)
        lines.append("End")
    lines.append("")
    return "\r\n".join(lines)


MACRO_FILES = {
    "AutoExec": [
        ("MsgBox", ["База данных «Поликлиника» загружена.", "-1", "0", "Поликлиника"]),
        ("OpenForm", ["Ф_Вход", "0", "", "", "-1", "0"]),
    ],
    "М_ОткрытьПриём": [
        ("OpenForm", ["Ф_ПриёмПациента", "0", "", "", "-1", "0"]),
    ],
    "М_ОбновитьСтатусы": [
        ("OpenQuery", ["Q09_ОбновлениеСтатусаПроведён", "0", "1"]),
        ("MsgBox", ["Статусы прошедших приёмов обновлены.", "-1", "0", "Обслуживание"]),
    ],
    "М_ЭкспортОтчётаExcel": [
        ("RunCode", ["ЭкспортВExcel()"]),
    ],
    "М_РезервнаяКопия": [
        ("RunCode", ["СоздатьРезервнуюКопию()"]),
    ],
    "М_ЗакрытьПриложение": [
        ("Quit", ["1"]),
    ],
}


VBS_BUILDER = """' ============================================================================
'  АВТОМАТИЧЕСКИЙ СБОРЩИК БАЗЫ ДАННЫХ MS ACCESS «ПОЛИКЛИНИКА» (В 1 КЛИК)
'  Вариант 14 — Сагадиев А. Р., группа Р-25-29
'  Запустите этот файл двойным щелчком мыши в Windows (где установлен MS Access).
'  Скрипт автоматически создаст файл «Поликлиника.accdb», импортирует модули VBA
'  и соберёт все 13 таблиц, 12 связей, данные, 16 запросов, 7 форм, 4 отчёта и 6 макросов.
' ============================================================================
Option Explicit

Dim fso, baseDir, dbPath, accApp
Set fso = CreateObject("Scripting.FileSystemObject")
baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
dbPath = baseDir & "\\Поликлиника.accdb"

If fso.FileExists(dbPath) Then
    If MsgBox("Файл Поликлиника.accdb уже существует." & vbCrLf & _
              "Пересоздать базу данных с нуля?", vbQuestion + vbYesNo, _
              "Сборка БД «Поликлиника»") = vbNo Then
        WScript.Quit 0
    End If
    On Error Resume Next
    fso.DeleteFile dbPath, True
    If Err.Number <> 0 Then
        MsgBox "Закройте файл Поликлиника.accdb в MS Access и повторите запуск.", vbExclamation
        WScript.Quit 1
    End If
    On Error GoTo 0
End If

On Error Resume Next
Set accApp = CreateObject("Access.Application")
If Err.Number <> 0 Then
    MsgBox "Не удалось запустить Microsoft Access (код: " & Err.Description & ")." & vbCrLf & _
           "Убедитесь, что Microsoft Access установлен на компьютере.", vbCritical, "Ошибка"
    WScript.Quit 1
End If
On Error GoTo 0

accApp.Visible = True
accApp.NewCurrentDatabase dbPath

' Загрузка модулей VBA через LoadFromText (работает без включения VBOM в реестре)
Sub LoadBasModule(app, modName, basPath)
    Dim tsIn, tsOut, tmpPath, line
    tmpPath = fso.GetSpecialFolder(2) & "\\" & modName & "_tmp.txt"
    Set tsIn = fso.OpenTextFile(basPath, 1, False, 0)
    Set tsOut = fso.CreateTextFile(tmpPath, True, False)
    Do Until tsIn.AtEndOfStream
        line = tsIn.ReadLine
        If Left(Trim(line), 9) <> "Attribute" Then
            tsOut.WriteLine line
        End If
    Loop
    tsIn.Close
    tsOut.Close
    app.LoadFromText 5, modName, tmpPath
    On Error Resume Next
    fso.DeleteFile tmpPath, True
    On Error GoTo 0
End Sub

LoadBasModule accApp, "ModSecurity", baseDir & "\\VBA\\ModSecurity.bas"
LoadBasModule accApp, "ModAdmin", baseDir & "\\VBA\\ModAdmin.bas"
LoadBasModule accApp, "AutoBuild", baseDir & "\\VBA\\AutoBuild.bas"

' Запуск автоматической сборки всех объектов БД
accApp.Run "СоздатьБД", True

accApp.CloseCurrentDatabase
accApp.Quit
Set accApp = Nothing

MsgBox "Готовая база данных «Поликлиника.accdb» успешно создана!" & vbCrLf & _
       "Путь: " & dbPath & vbCrLf & vbCrLf & _
       "Учётные записи для входа:" & vbCrLf & _
       "  admin / Admin2026 (Администратор)" & vbCrLf & _
       "  dr_valieva / Doc2026 (Врач)" & vbCrLf & _
       "  registrator / Reg2026 (Регистратор)", vbInformation, "Сборка завершена"
"""


def write_macros():
    os.makedirs(MACRO_DIR, exist_ok=True)
    for name, acts in MACRO_FILES.items():
        with open(os.path.join(MACRO_DIR, name + ".txt"), "w", encoding="cp1251") as f:
            f.write(macro_text(acts))
        print("  макрос:", name)


def main():
    os.makedirs(VBA_DIR, exist_ok=True)
    parts = [HEADER]
    parts.append(MAIN_CODE)
    parts.append("' ============================================================================\n"
                 "'  ЭТАП 2. СОЗДАНИЕ ТАБЛИЦ\n"
                 "' ============================================================================")
    parts.append("\n".join(part_tables()))
    parts.append("\n' ============================================================================\n"
                 "'  ЭТАП 3. СВОЙСТВА ПОЛЕЙ\n"
                 "' ============================================================================")
    parts.append("\n".join(part_props()))
    parts.append("\n' ============================================================================\n"
                 "'  ЭТАП 4. СВЯЗИ (ВНЕШНИЕ КЛЮЧИ)\n"
                 "' ============================================================================")
    parts.append("\n".join(part_relations()))
    parts.append("\n' ============================================================================\n"
                 "'  ЭТАП 5. ИНДЕКСЫ\n"
                 "' ============================================================================")
    parts.append("\n".join(part_indexes()))
    parts.append("\n' ============================================================================\n"
                 "'  ЭТАП 6. ТЕСТОВЫЕ ДАННЫЕ\n"
                 "' ============================================================================")
    parts.append("\n".join(part_data()))
    parts.append("\n' ============================================================================\n"
                 "'  ЭТАП 7. ЗАПРОСЫ\n"
                 "' ============================================================================")
    parts.append("\n".join(part_queries()))
    parts.append("\n' ============================================================================\n"
                 "'  ЭТАП 8. ФОРМЫ\n"
                 "' ============================================================================")
    parts.append("\n\n".join(part_forms()))
    parts.append("\n' ============================================================================\n"
                 "'  ЭТАП 9. ОТЧЁТЫ\n"
                 "' ============================================================================")
    parts.append(REPORTS_CODE.strip())
    parts.append("\n' ============================================================================\n"
                 "'  ЭТАП 10. МАКРОСЫ И ПАРАМЕТРЫ ЗАПУСКА\n"
                 "' ============================================================================")
    parts.append(MACROS_CODE.strip())
    parts.append(STARTUP_CODE.strip())
    parts.append(FOOTER.strip())

    text = "\n\n".join(parts).rstrip() + "\n"

    with open(os.path.join(VBA_DIR, "AutoBuild.bas"), "w", encoding="cp1251", newline="\r\n") as f:
        f.write(text)
    with open(os.path.join(VBA_DIR, "AutoBuild_UTF8.bas"), "w", encoding="utf-8", newline="\r\n") as f:
        f.write(text)

    for name, body in (("ModSecurity.bas", MOD_SECURITY), ("ModAdmin.bas", MOD_ADMIN)):
        with open(os.path.join(VBA_DIR, name), "w", encoding="cp1251", newline="\r\n") as f:
            f.write(body.rstrip() + "\n")
        with open(os.path.join(VBA_DIR, name.replace(".bas", "_UTF8.bas")), "w",
                  encoding="utf-8", newline="\r\n") as f:
            f.write(body.rstrip() + "\n")
        print("  модуль:", name)

    vbs_path = os.path.join(ROOT, "Создать_БД_Поликлиника.vbs")
    with open(vbs_path, "w", encoding="cp1251", newline="\r\n") as f:
        f.write(VBS_BUILDER.lstrip())
    print("  скрипт сборки в 1 клик:", vbs_path)

    bad = [(i, len(l)) for i, l in enumerate(text.splitlines(), 1) if len(l) > 1020]
    if bad:
        print("  ! слишком длинные строки:", bad[:5])
    print("  AutoBuild.bas: %d строк, %d символов" % (len(text.splitlines()), len(text)))
    write_macros()


if __name__ == "__main__":
    main()
