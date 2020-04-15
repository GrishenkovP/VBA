'****************************************************************************************************************************************
'Product: Коннектор 1СБух 8.3 (файловый вариант)
'Author : Павел Гришенков
'Information: Демо-версия программы
'****************************************************************************************************************************************
Option Explicit

Public Const strVersion As String = "v83.Application"

Public shParameter As Worksheet
Public shData As Worksheet
Public strPath As String
Public strUser As String
Public strPassword As String
Public strOrganization As String
Public strDateStart As String
Public strDateFinish As String
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public bBoolean As Boolean
Public objFileDb As Object
Public strText As String
Public objQuery As Object
Public objOrg As Object
Public objАccount1 As Object
Public objАccount2 As Object
Public dt1 As Date
Public dt2 As Date
Public objResult As Object
Public lStart As Long

'****************************************************************************************************************************************

Sub MacroMain()

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Set shParameter = ThisWorkbook.Sheets("Параметры")

strPath = shParameter.Cells(3, 2).Value
        If strPath = "" Then
        MsgBox "Укажите путь к БД!"
        Exit Sub
        End If

strUser = shParameter.Cells(4, 2).Value
        If strUser = "" Then
        MsgBox "Укажите имя пользователя!"
        Exit Sub
        End If
        
strPassword = shParameter.Cells(5, 2).Value

strOrganization = shParameter.Cells(7, 2).Value
        If strOrganization = "" Then
        MsgBox "Укажите название организации!"
        Exit Sub
        End If

strDateStart = shParameter.Cells(8, 2).Value
        If strDateStart = "" Then
        MsgBox "Заполните корректно поле даты!"
        Exit Sub
        End If
        

strDateFinish = shParameter.Cells(9, 2).Value
        If strDateFinish = "" Then
        MsgBox "Заполните корректно поле даты!"
        Exit Sub
        End If
        
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Set shData = ThisWorkbook.Sheets("ЗагрузкаДанных")
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Select Case shParameter.Cells(11, 2).Value

    Case "Выпуск продукции и услуг"
    Call ProduceProduct(strPath, strUser, strPassword, strOrganization, strDateStart, strDateFinish, shData)
    
    Case "Номенклатура"
    Call Nomenclature(strPath, strUser, strPassword, shData)
    
    Case "Контрагенты"
    Call Counterparties(strPath, strUser, strPassword, shData)
    
    Case "Банковские выписки"
    Call BankStatements(strPath, strUser, strPassword, strOrganization, strDateStart, strDateFinish, shData)
    
    Case "Кассовые документы"
    Call CashDocuments(strPath, strUser, strPassword, strOrganization, strDateStart, strDateFinish, shData)
 
    Case "Реализация товаров и услуг"
    Call SalesProductService(strPath, strUser, strPassword, strOrganization, strDateStart, strDateFinish, shData)
    
    Case "Поступление товаров и услуг"
    Call ReceiptProductService(strPath, strUser, strPassword, strOrganization, strDateStart, strDateFinish, shData)
    
'    Case "План счетов"
'    Call Accounts(strPath, strUser, strPassword, shData)
    
    Case "Журнал проводок"
    Call AccountingRecords(strPath, strUser, strPassword, strOrganization, strDateStart, strDateFinish, shData)

    Case Else
    MsgBox "Укажите источник данных!"
    Exit Sub

End Select

If bBoolean Then
    MsgBox "Загрузка успешно завершена!"
    Else
    MsgBox "Загрузка не произошла из-за ошибки!"
End If

End Sub

