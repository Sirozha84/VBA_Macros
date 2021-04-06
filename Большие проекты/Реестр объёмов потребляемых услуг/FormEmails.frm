VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormEmails 
   Caption         =   "Рассылка почты"
   ClientHeight    =   2640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5520
   OleObjectBlob   =   "FormEmails.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormEmails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Last change: 06.04.2021 08:22

Private Sub UserForm_Activate()
    On Error GoTo er
    TextBoxAdr = Sheets("Рассылка").name
    Exit Sub
er:
    MsgBox ("Нет таблицы ""Рассылка""!")
End Sub

'Кнопка рассылки почты
Private Sub ButtonOK_Click()
    On Error GoTo er
    Set eml = Sheets(TextBoxAdr.text)
    On Error GoTo 0
    i = 2
    s = 0
    Do While eml.Cells(i, 1) <> ""
        adr = eml.Cells(i, 2)
        If adr <> "" And LCase(eml.Cells(i, 3).text) = "да" Then
            cmp = Cells(i, 1).text
            For Each sht In Sheets
                'Debug.Print sht.name
                If sht.name = cmp Then
                    If sendmail(sht, adr) Then s = s + 1
                End If
            Next
        End If
        i = i + 1
    Loop
    
    If s > 0 Then
        MsgBox "Рассылка завершена, отправлено писем: " + CStr(s)
    Else
        MsgBox "Ни одного письма не отправлено"
    End If
    End
er:
    MsgBox "Указанная таблица не существует"
End Sub

'Отправка страницы sht по почте
Function sendmail(ByVal sht As Variant, ByVal adr As String) As Boolean
    On Error GoTo er
    sht.Copy
    'ActiveWorkbook.sendmail adr, "Объёмы"
    With ActiveWorkbook
        .sendmail Recipients:=adr, Subject:="Объёмы"
        .Close SaveChanges:=False
    End With
    
    sendmail = True
    Exit Function
er:
    ActiveWorkbook.Close False
    sendmail = False
End Function

'Кнопка отмены
Private Sub ButtonClose_Click()
    End
End Sub

'******************** End of File ********************