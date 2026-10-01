Attribute VB_Name = "ModSecurity"
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
