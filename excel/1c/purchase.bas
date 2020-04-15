Option Explicit

'****************************************************************************************************************************************
Public Sub ReceiptProductService(strPath, strUser, strPassword, strOrganization, strDateStart, strDateFinish, shData)
' Подключение к БД средствами OLE Automation

bBoolean = False
On Error GoTo ErrorFileDb

Set objFileDb = CreateObject("" & strVersion & "")
objFileDb.Connect ("File=""" & strPath & """;Usr=""" & strUser & """;Pwd=""" & strPassword & """")

On Error GoTo 0

' Установка видимости приложения
objFileDb.Visible = False
' Задаем текст запроса
strText = "SELECT ХозрасчетныйДвиженияССубконто.Период, ХозрасчетныйДвиженияССубконто.Регистратор,ХозрасчетныйДвиженияССубконто.СубконтоДт1,ХозрасчетныйДвиженияССубконто.СубконтоДт3,ХозрасчетныйДвиженияССубконто.СубконтоКт1,ХозрасчетныйДвиженияССубконто.СубконтоКт2,ХозрасчетныйДвиженияССубконто.СубконтоКт3, ХозрасчетныйДвиженияССубконто.КоличествоДт,ХозрасчетныйДвиженияССубконто.Сумма " & vbNewLine & _
          "FROM РегистрБухгалтерии.Хозрасчетный.ДвиженияССубконто(&Дата1,&Дата2,Организация =&Организация И СчетДт = &СчетДт И СчетКт = &СчетКт,,) " & vbNewLine & _
          "AS ХозрасчетныйДвиженияССубконто "

' Создаем объект Запрос
Set objQuery = objFileDb.NewObject("Запрос")
objQuery.Text = strText

'Передача параметров
Set objOrg = objFileDb.Справочники.Организации.НайтиПоНаименованию("" & strOrganization & "")
objQuery.SetParameter "Организация", objOrg
Set objАccount1 = objFileDb.ПланыСчетов.Хозрасчетный.НайтиПоКоду("41.01")
objQuery.SetParameter "СчетДт", objАccount1
Set objАccount2 = objFileDb.ПланыСчетов.Хозрасчетный.НайтиПоКоду("60.01")
objQuery.SetParameter "СчетКт", objАccount2

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
        .Cells(1, 1).Value = "Период"
        .Cells(1, 2).Value = "Регистратор"
        .Cells(1, 3).Value = "СубконтоДт1"
        .Cells(1, 4).Value = "СубконтоДт3"
        .Cells(1, 5).Value = "СубконтоКт1"
        .Cells(1, 6).Value = "СубконтоКт2"
        .Cells(1, 7).Value = "СубконтоКт3"
        .Cells(1, 8).Value = "КоличествоДт"
        .Cells(1, 9).Value = "Сумма"
        
        End If
        
        
        'Определяем последнюю заполненную строку на листе
        lStart = .Cells(Rows.Count, "A").End(xlUp).Row + 1
        
        ' Заполняем ячейки
        .Cells(lStart, 1).Value = objFileDb.String(objResult.Период)
        .Cells(lStart, 2).Value = objFileDb.String(objResult.Регистратор)
        .Cells(lStart, 3).Value = objFileDb.String(objResult.СубконтоДт1)
        .Cells(lStart, 4).Value = objFileDb.String(objResult.СубконтоДт3)
        .Cells(lStart, 5).Value = objFileDb.String(objResult.СубконтоКт1)
        .Cells(lStart, 6).Value = objFileDb.String(objResult.СубконтоКт2)
        .Cells(lStart, 7).Value = objFileDb.String(objResult.СубконтоКт3)
        .Cells(lStart, 8).Value = CLng(objFileDb.String(objResult.КоличествоДт))
        .Cells(lStart, 9).Value = CCur(objFileDb.String(objResult.Сумма))
        
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





