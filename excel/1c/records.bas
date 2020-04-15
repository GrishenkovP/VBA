Option Explicit

'****************************************************************************************************************************************
Public Sub AccountingRecords(strPath, strUser, strPassword, strOrganization, strDateStart, strDateFinish, shData)
' Подключение к БД средствами OLE Automation

bBoolean = False
On Error GoTo ErrorFileDb


Set objFileDb = CreateObject("" & strVersion & "")
objFileDb.Connect ("File=""" & strPath & """;Usr=""" & strUser & """;Pwd=""" & strPassword & """")

On Error GoTo 0

' Установка видимости приложения
objFileDb.Visible = False
' Задаем текст запроса
strText = "SELECT РегБух.Период,РегБух.Регистратор, РегБух.Организация, РегБух.СчетДт, РегБух.ВидСубконтоДт1,РегБух.СубконтоДт1,РегБух.ВидСубконтоДт2,РегБух.СубконтоДт2,РегБух.ВидСубконтоДт3,РегБух.СубконтоДт3,РегБух.СчетКт,РегБух.ВидСубконтоКт1,РегБух.СубконтоКт1,РегБух.ВидСубконтоКт2,РегБух.СубконтоКт2,РегБух.ВидСубконтоКт3,РегБух.СубконтоКт3,РегБух.Содержание,РегБух.Сумма " & vbNewLine & _
          "FROM РегистрБухгалтерии.Хозрасчетный.ДвиженияССубконто " & vbNewLine & _
          "AS РегБух " & vbNewLine & _
          "WHERE РегБух.Организация = &Организация AND РегБух.Период >= &Дата1 AND РегБух.Период <= &Дата2"

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
        .Cells(1, 1).Value = "Период"
        .Cells(1, 2).Value = "Регистратор"
        .Cells(1, 3).Value = "Организация"
        .Cells(1, 4).Value = "СчетДт"
        .Cells(1, 5).Value = "ВидСубконтоДт1"
        .Cells(1, 6).Value = "СубконтоДт1"
        .Cells(1, 7).Value = "ВидСубконтоДт2"
        .Cells(1, 8).Value = "СубконтоДт2"
        .Cells(1, 9).Value = "ВидСубконтоДт3"
        .Cells(1, 10).Value = "СубконтоДт3"
        .Cells(1, 11).Value = "СчетКт"
        .Cells(1, 12).Value = "ВидСубконтоКт1"
        .Cells(1, 13).Value = "СубконтоКт1"
        .Cells(1, 14).Value = "ВидСубконтоКт2"
        .Cells(1, 15).Value = "СубконтоКт2"
        .Cells(1, 16).Value = "ВидСубконтоКт3"
        .Cells(1, 17).Value = "СубконтоКт3"
        .Cells(1, 18).Value = "Содержание"
        .Cells(1, 19).Value = "Сумма"
        End If
        
        
        'Определяем последнюю заполненную строку на листе
        lStart = .Cells(Rows.Count, "A").End(xlUp).Row + 1
        
        ' Заполняем ячейки
        .Cells(lStart, 1).Value = objFileDb.String(objResult.Период)
        .Cells(lStart, 2).Value = objFileDb.String(objResult.Регистратор)
        .Cells(lStart, 3).Value = objFileDb.String(objResult.Организация)
        
        .Cells(lStart, 4).Value = "'" & objFileDb.String(objResult.СчетДт)
        .Cells(lStart, 5).Value = objFileDb.String(objResult.ВидСубконтоДт1)
        .Cells(lStart, 6).Value = objFileDb.String(objResult.СубконтоДт1)
        .Cells(lStart, 7).Value = objFileDb.String(objResult.ВидСубконтоДт2)
        .Cells(lStart, 8).Value = objFileDb.String(objResult.СубконтоДт2)
        .Cells(lStart, 9).Value = objFileDb.String(objResult.ВидСубконтоДт3)
        .Cells(lStart, 10).Value = objFileDb.String(objResult.СубконтоДт3)
        
        .Cells(lStart, 11).Value = "'" & objFileDb.String(objResult.СчетКт)
        .Cells(lStart, 12).Value = objFileDb.String(objResult.ВидСубконтоКт1)
        .Cells(lStart, 13).Value = objFileDb.String(objResult.СубконтоКт1)
        .Cells(lStart, 14).Value = objFileDb.String(objResult.ВидСубконтоКт2)
        .Cells(lStart, 15).Value = objFileDb.String(objResult.СубконтоКт2)
        .Cells(lStart, 16).Value = objFileDb.String(objResult.ВидСубконтоКт3)
        .Cells(lStart, 17).Value = objFileDb.String(objResult.СубконтоКт3)
             
        .Cells(lStart, 18).Value = objFileDb.String(objResult.Содержание)
        .Cells(lStart, 19).Value = CCur(objFileDb.String(objResult.Сумма))
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



