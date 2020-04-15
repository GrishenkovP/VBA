Option Explicit

'****************************************************************************************************************************************
Public Sub ProduceProduct(strPath, strUser, strPassword, strOrganization, strDateStart, strDateFinish, shData)
' Подключение к БД средствами OLE Automation

bBoolean = False
On Error GoTo ErrorFileDb


Set objFileDb = CreateObject("" & strVersion & "")
objFileDb.Connect ("File=""" & strPath & """;Usr=""" & strUser & """;Pwd=""" & strPassword & """")

On Error GoTo 0

' Установка видимости приложения
objFileDb.Visible = False
' Задаем текст запроса
strText = "SELECT ВыпускПродукцииУслуг.Регистратор, ВыпускПродукцииУслуг.Подразделение, ВыпускПродукцииУслуг.НоменклатурнаяГруппа,ВыпускПродукцииУслуг.Продукция,ВыпускПродукцииУслуг.Количество AS КоличествоИзд,ВыпускПродукцииУслуг.ПлановаяСтоимость " & vbNewLine & _
          "FROM РегистрНакопления.ВыпускПродукцииУслуг " & vbNewLine & _
          "AS ВыпускПродукцииУслуг " & vbNewLine & _
          "WHERE ВыпускПродукцииУслуг.Организация = &Организация AND ВыпускПродукцииУслуг.Период >= &Дата1 AND ВыпускПродукцииУслуг.Период <= &Дата2"

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
        .Cells(1, 1).Value = "Регистратор"
        .Cells(1, 2).Value = "Подразделение"
        .Cells(1, 3).Value = "НомеклатурнаяГруппа"
        .Cells(1, 4).Value = "Продукция"
        .Cells(1, 5).Value = "Количество"
        .Cells(1, 6).Value = "ПлановаяСтоимость"
        End If
        
        
        'Определяем последнюю заполненную строку на листе
        lStart = .Cells(Rows.Count, "A").End(xlUp).Row + 1
        
        ' Заполняем ячейки
        .Cells(lStart, 1).Value = objFileDb.String(objResult.Регистратор)
        .Cells(lStart, 2).Value = objFileDb.String(objResult.Подразделение)
        .Cells(lStart, 3).Value = objFileDb.String(objResult.НоменклатурнаяГруппа)
        .Cells(lStart, 4).Value = objFileDb.String(objResult.Продукция)
        .Cells(lStart, 5).Value = CLng(objFileDb.String(objResult.КоличествоИзд))
        .Cells(lStart, 6).Value = CCur(objFileDb.String(objResult.ПлановаяСтоимость))
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

