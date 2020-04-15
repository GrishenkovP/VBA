Option Explicit

'****************************************************************************************************************************************
Public Sub BankStatements(strPath, strUser, strPassword, strOrganization, strDateStart, strDateFinish, shData)
' Подключение к БД средствами OLE Automation

bBoolean = False
On Error GoTo ErrorFileDb

Set objFileDb = CreateObject("" & strVersion & "")
objFileDb.Connect ("File=""" & strPath & """;Usr=""" & strUser & """;Pwd=""" & strPassword & """")

On Error GoTo 0

' Установка видимости приложения
objFileDb.Visible = False
' Задаем текст запроса
strText = "SELECT БанковскиеВыписки.Дата, БанковскиеВыписки.Номер, БанковскиеВыписки.БанковскийСчет,БанковскиеВыписки.Контрагент,БанковскиеВыписки.ВидОперации,БанковскиеВыписки.Поступление, БанковскиеВыписки.Списание, БанковскиеВыписки.Валюта " & vbNewLine & _
          "FROM ЖурналДокументов.БанковскиеВыписки " & vbNewLine & _
          "AS БанковскиеВыписки " & vbNewLine & _
          "WHERE БанковскиеВыписки.Организация = &Организация AND БанковскиеВыписки.Дата >= &Дата1 AND БанковскиеВыписки.Дата <= &Дата2"

' Создаем объект Запрос
Set objQuery = objFileDb.NewObject("Запрос")
objQuery.Text = strText

'Передача параметров
Set objOrg = objFileDb.Справочники.Организации.НайтиПоНаименованию("" & strOrganization & "")
objQuery.SetParameter "Организация", objOrg
dt1 = CDate(strDateStart)
objQuery.SetParameter "Дата1", dt1
dt2 = CDate(strDateFinish & " 23:59:59")
objQuery.SetParameter "Дата2", dt2

' Выполнение запроса
Set objResult = objQuery.Execute().Choose()


        'Отключаем автоматический пересчет формул
        Application.Calculation = xlCalculationManual
        'Отключаем обновление экрана
        Application.ScreenUpdating = False
        'Отключаем отслеживание событий
        Application.EnableEvents = False
        'Отключаем отображение информации в строке статуса Excel
        Application.DisplayStatusBar = False
                    

' Выборка результата выполнения запроса
Do While objResult.Next()

    With shData
    
        'Формируем шапку отчета
        If .Cells(1, 1) = "" Then
        .Cells(1, 1).Value = "Дата"
        .Cells(1, 2).Value = "Номер"
        .Cells(1, 3).Value = "БанковскийСчет"
        .Cells(1, 4).Value = "ВидОперации"
        .Cells(1, 5).Value = "Поступление"
        .Cells(1, 6).Value = "Списание"
        .Cells(1, 7).Value = "Валюта"
        End If
        
        
        'Определяем последнюю заполненную строку на листе
        lStart = .Cells(Rows.Count, "A").End(xlUp).Row + 1
        
        ' Заполняем ячейки
        .Cells(lStart, 1).Value = objFileDb.String(objResult.Дата)
        .Cells(lStart, 2).Value = objFileDb.String(objResult.Номер)
        .Cells(lStart, 3).Value = objFileDb.String(objResult.БанковскийСчет)
        .Cells(lStart, 4).Value = objFileDb.String(objResult.ВидОперации)
        
        If objFileDb.String(objResult.Поступление) = "" Then
            .Cells(lStart, 5).Value = ""
        Else
            .Cells(lStart, 5).Value = CCur(objFileDb.String(objResult.Поступление))
        End If
    
        If objFileDb.String(objResult.Списание) = "" Then
            .Cells(lStart, 6).Value = ""
        Else
            .Cells(lStart, 6).Value = CCur(objFileDb.String(objResult.Списание))
        End If
        
        
        .Cells(lStart, 7).Value = objFileDb.String(objResult.Валюта)
        lStart = lStart + 1
    
    End With
    
Loop


        'Включаем автоматический пересчет формул
        Application.Calculation = xlCalculationAutomatic
        'Включаем обновление экрана
        Application.ScreenUpdating = True
        'Включаем отслеживание событий
        Application.EnableEvents = True
        'Включаем отображение информации в строке статуса Excel
        Application.DisplayStatusBar = True
        
        
bBoolean = True

Exit Sub

ErrorFileDb:
MsgBox "Проверьте правильность пути к базе данных, имя пользователя и указанный пароль!"


End Sub

'****************************************************************************************************************************************

'****************************************************************************************************************************************
Public Sub CashDocuments(strPath, strUser, strPassword, strOrganization, strDateStart, strDateFinish, shData)
' Подключение к БД средствами OLE Automation

bBoolean = False
On Error GoTo ErrorFileDb

Set objFileDb = CreateObject("" & strVersion & "")
objFileDb.Connect ("File=""" & strPath & """;Usr=""" & strUser & """;Pwd=""" & strPassword & """")

On Error GoTo 0

' Установка видимости приложения
objFileDb.Visible = False
' Задаем текст запроса
strText = "SELECT КассовыеДокументы.Дата, КассовыеДокументы.Номер, КассовыеДокументы.Тип,КассовыеДокументы.ВидОперации,КассовыеДокументы.Контрагент,КассовыеДокументы.Приход, КассовыеДокументы.Расход, КассовыеДокументы.Валюта " & vbNewLine & _
          "FROM ЖурналДокументов.КассовыеДокументы " & vbNewLine & _
          "AS КассовыеДокументы " & vbNewLine & _
          "WHERE КассовыеДокументы.Организация = &Организация AND КассовыеДокументы.Дата >= &Дата1 AND КассовыеДокументы.Дата <= &Дата2"

' Создаем объект Запрос
Set objQuery = objFileDb.NewObject("Запрос")
objQuery.Text = strText

'Передача параметров
Set objOrg = objFileDb.Справочники.Организации.НайтиПоНаименованию("" & strOrganization & "")
objQuery.SetParameter "Организация", objOrg
dt1 = CDate(strDateStart)
objQuery.SetParameter "Дата1", dt1
dt2 = CDate(strDateFinish & " 23:59:59")
objQuery.SetParameter "Дата2", dt2

' Выполнение запроса
Set objResult = objQuery.Execute().Choose()


        'Отключаем автоматический пересчет формул
        Application.Calculation = xlCalculationManual
        'Отключаем обновление экрана
        Application.ScreenUpdating = False
        'Отключаем отслеживание событий
        Application.EnableEvents = False
        'Отключаем отображение информации в строке статуса Excel
        Application.DisplayStatusBar = False
                

' Выборка результата выполнения запроса
Do While objResult.Next()

    With shData
    
        'Формируем шапку отчета
        If .Cells(1, 1) = "" Then
        .Cells(1, 1).Value = "Дата"
        .Cells(1, 2).Value = "Номер"
        .Cells(1, 3).Value = "Тип"
        .Cells(1, 4).Value = "ВидОперации"
        .Cells(1, 5).Value = "Контрагент"
        .Cells(1, 6).Value = "Приход"
        .Cells(1, 7).Value = "Расход"
        .Cells(1, 8).Value = "Валюта"
        End If
        
        
        'Определяем последнюю заполненную строку на листе
        lStart = .Cells(Rows.Count, "A").End(xlUp).Row + 1
        
        ' Заполняем ячейки
        .Cells(lStart, 1).Value = objFileDb.String(objResult.Дата)
        .Cells(lStart, 2).Value = objFileDb.String(objResult.Номер)
        .Cells(lStart, 3).Value = objFileDb.String(objResult.Тип)
        .Cells(lStart, 4).Value = objFileDb.String(objResult.ВидОперации)
        .Cells(lStart, 5).Value = objFileDb.String(objResult.Контрагент)
        
        If objFileDb.String(objResult.Приход) = "" Then
            .Cells(lStart, 6).Value = ""
        Else
            .Cells(lStart, 6).Value = CCur(objFileDb.String(objResult.Приход))
        End If
    
        If objFileDb.String(objResult.Расход) = "" Then
            .Cells(lStart, 7).Value = ""
        Else
            .Cells(lStart, 7).Value = CCur(objFileDb.String(objResult.Расход))
        End If
        
        
        .Cells(lStart, 8).Value = objFileDb.String(objResult.Валюта)
        lStart = lStart + 1
    
    End With
    
Loop

        'Включаем автоматический пересчет формул
        Application.Calculation = xlCalculationAutomatic
        'Включаем обновление экрана
        Application.ScreenUpdating = True
        'Включаем отслеживание событий
        Application.EnableEvents = True
        'Включаем отображение информации в строке статуса Excel
        Application.DisplayStatusBar = True

bBoolean = True

Exit Sub

ErrorFileDb:
MsgBox "Проверьте правильность пути к базе данных, имя пользователя и указанный пароль!"


End Sub
