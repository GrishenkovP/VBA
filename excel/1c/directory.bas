Option Explicit

'****************************************************************************************************************************************
Public Sub Nomenclature(strPath, strUser, strPassword, shData)
' Подключение к БД средствами OLE Automation

bBoolean = False
On Error GoTo ErrorFileDb

Set objFileDb = CreateObject("" & strVersion & "")
objFileDb.Connect ("File=""" & strPath & """;Usr=""" & strUser & """;Pwd=""" & strPassword & """")

On Error GoTo 0

' Установка видимости приложения
objFileDb.Visible = False
' Задаем текст запроса
strText = "SELECT Номенклатура.Код, Номенклатура.НоменклатурнаяГруппа, Номенклатура.Наименование, Номенклатура.ЕдиницаИзмерения, Номенклатура.Услуга, Номенклатура.ЭтоГруппа " & vbNewLine & _
          "FROM Справочник.Номенклатура " & vbNewLine & _
          "AS Номенклатура "
'

' Создаем объект Запрос
Set objQuery = objFileDb.NewObject("Запрос")
objQuery.Text = strText

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
        .Cells(1, 1).Value = "Код"
        .Cells(1, 2).Value = "НоменклатурнаяГруппа"
        .Cells(1, 3).Value = "Наименование"
        .Cells(1, 4).Value = "ЕдиницаИзмерения"
        .Cells(1, 5).Value = "Услуга"
        .Cells(1, 6).Value = "ЭтоГруппа"
        End If
        
        
        'Определяем последнюю заполненную строку на листе
        lStart = .Cells(Rows.Count, "A").End(xlUp).Row + 1
                       
                
        ' Заполняем ячейки
        .Cells(lStart, 1).Value = objFileDb.String(objResult.Код)
        .Cells(lStart, 2).Value = objFileDb.String(objResult.НоменклатурнаяГруппа)
        .Cells(lStart, 3).Value = objFileDb.String(objResult.Наименование)
        .Cells(lStart, 4).Value = objFileDb.String(objResult.ЕдиницаИзмерения)
        .Cells(lStart, 5).Value = objFileDb.String(objResult.Услуга)
        .Cells(lStart, 6).Value = objFileDb.String(objResult.ЭтоГруппа)
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

'************************************************************************************************************************************************


'****************************************************************************************************************************************
Public Sub Counterparties(strPath, strUser, strPassword, shData)
' Подключение к БД средствами OLE Automation

bBoolean = False
On Error GoTo ErrorFileDb

Set objFileDb = CreateObject("" & strVersion & "")
objFileDb.Connect ("File=""" & strPath & """;Usr=""" & strUser & """;Pwd=""" & strPassword & """")

On Error GoTo 0

' Установка видимости приложения
objFileDb.Visible = False
' Задаем текст запроса
strText = "SELECT Контрагенты.Код, Контрагенты.НаименованиеПолное, Контрагенты.ОбособленноеПодразделение, Контрагенты.ЮридическоеФизическоеЛицо, Контрагенты.ОсновнойДоговорКонтрагента, Контрагенты.ЭтоГруппа " & vbNewLine & _
          "FROM Справочник.Контрагенты " & vbNewLine & _
          "AS Контрагенты "
'

' Создаем объект Запрос
Set objQuery = objFileDb.NewObject("Запрос")
objQuery.Text = strText

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
        .Cells(1, 1).Value = "Код"
        .Cells(1, 2).Value = "НаименованиеПолное"
        .Cells(1, 3).Value = "ОбособленноеПодразделение"
        .Cells(1, 4).Value = "ЮридическоеФизическоеЛицо"
        .Cells(1, 5).Value = "ОсновноеДоговорКонтрагента"
        .Cells(1, 6).Value = "ЭтоГруппа"
        End If
        
        
        'Определяем последнюю заполненную строку на листе
        lStart = .Cells(Rows.Count, "A").End(xlUp).Row + 1
        
               
                
        ' Заполняем ячейки
        .Cells(lStart, 1).Value = objFileDb.String(objResult.Код)
        .Cells(lStart, 2).Value = objFileDb.String(objResult.НаименованиеПолное)
        .Cells(lStart, 3).Value = objFileDb.String(objResult.ОбособленноеПодразделение)
        .Cells(lStart, 4).Value = objFileDb.String(objResult.ЮридическоеФизическоеЛицо)
        .Cells(lStart, 5).Value = objFileDb.String(objResult.ОсновнойДоговорКонтрагента)
        .Cells(lStart, 6).Value = objFileDb.String(objResult.ЭтоГруппа)
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

'************************************************************************************************************************************************
