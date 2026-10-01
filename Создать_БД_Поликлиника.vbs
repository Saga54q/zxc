' ============================================================================
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
dbPath = baseDir & "\Поликлиника.accdb"

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
    tmpPath = fso.GetSpecialFolder(2) & "\" & modName & "_tmp.txt"
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

LoadBasModule accApp, "ModSecurity", baseDir & "\VBA\ModSecurity.bas"
LoadBasModule accApp, "ModAdmin", baseDir & "\VBA\ModAdmin.bas"
LoadBasModule accApp, "AutoBuild", baseDir & "\VBA\AutoBuild.bas"

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
