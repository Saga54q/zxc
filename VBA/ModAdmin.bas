Attribute VB_Name = "ModAdmin"
' ============================================================================
'  МОДУЛЬ АДМИНИСТРИРОВАНИЯ: РЕЗЕРВНОЕ КОПИРОВАНИЕ, ИМПОРТ/ЭКСПОРТ, СЕРВИС
'  База данных «Поликлиника» (вариант 14, Сагадиев А. Р., гр. Р-25-29)
' ============================================================================
Option Compare Database
Option Explicit

Private Function ПокаПапка(ByVal Имя As String) As String
    Dim p As String
    p = CurrentProject.Path & "\" & Имя
    If Dir(p, vbDirectory) = "" Then MkDir p
    ПокаПапка = p & "\"
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
