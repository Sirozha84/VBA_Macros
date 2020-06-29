Attribute VB_Name = "Splitter"
Sub Spliter()
    ii = 4
    For i = 4 To 1276
        st = Sheets("TDSheet").Cells(i, 5)
        s = Split(st, Chr(10))
        For j = 0 To UBound(s)
            For k = 1 To 4
                Sheets("Result").Cells(ii, k) = Sheets("TDSheet").Cells(i, k)
            Next
            ss = Split(s(j), ";")
            For k = 0 To UBound(ss)
                Sheets("Result").Cells(ii, 5 + k) = ss(k)
            Next
            ii = ii + 1
        Next
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


