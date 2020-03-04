Attribute VB_Name = "Uslugi"
'Выборка услуг и схлопывание строк

Const Max = 194311
Const Col = 7
Const Usl = 8

Sub Uslugi()
    Application.ScreenUpdating = False
    last = 0
    For i = 2 To Max
        
        If i Mod 50 = 0 Then Call ProgressBar("Обработка", i, Max)
        
        If last = Cells(i, Col) Then
            Cells(i - 1, Usl) = Cells(i, Usl)
            Rows(i).EntireRow.Delete
            i = i - 1
        Else
            last = Cells(i, Col)
        End If
        
        If Cells(i, Usl) = "ХВС" Then Cells(i, 10) = "ХВС"
        If Cells(i, Usl) = "ГВС ТН" Then Cells(i, 11) = "ГВС ТН"
        If Cells(i, Usl) = "ВО" Then Cells(i, 12) = "ВО"
        If Cells(i, Usl) = "Отопление" Then Cells(i, 13) = "Отопление"

        
        If Cells(i, 1) = "" Then Exit For
    Next
End Sub

'Рисование прогресса, text - имя, cur - текущее значение, all - всего, отображать каждые over штук
Private Sub ProgressBar(text As String, ByVal cur As Long, ByVal all As Long)
    Application.ScreenUpdating = True
    Application.StatusBar = text + ":" + Str(cur) + " из" + Str(all) + _
        " (" + Str(Int(cur / all * 100)) + "% )"
        DoEvents
    Application.ScreenUpdating = False
End Sub


Private Function Adress(i As Long) As String
    Adress = CStr(Cells(i, 1)) + CStr(Cells(i, 2)) + CStr(Cells(i, 3)) + CStr(Cells(i, 4))
    Adress = LCase(Adress)
    Adress = Replace(Adress, "ё", "е")
End Function

Private Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    DoEvents
    Application.ScreenUpdating = False
End Sub





