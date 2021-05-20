VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Rows(1).Row = 2 And Target.Columns(1).Column = 11 Then
        On Error GoTo er
        CreateHead Cells(2, 11)
    End If
    Exit Sub
er:
    MsgBox "Дату, пожалуйста..."
End Sub

Sub CreateHead(dat As Date)
    first = DateSerial(Year(dat), Month(dat), 1)
    last = DateSerial(Year(dat), Month(dat) + 1, 1) - 1
    Application.ScreenUpdating = False
    For i = 8 To 17
        'Cells(5, i) = NumberFormat = "@"
        Cells(5, i) = " "
    Next
    k = 0
    For d = first To last
        If Weekday(d, vbMonday) = 2 Or Weekday(d, vbMonday) = 4 Then
            Cells(5, 8 + k) = d
            k = k + 1
        End If
    Next
End Sub
