Attribute VB_Name = "VolumeFind"
Const mainTab = "Result"

Sub VolumeFind()

    max = 108813
    For i = 2 To max
        If i Mod 100 = 0 Then Call ProgressBar("Обработка", i, max)
        If Sheets(mainTab).Cells(i, 21) = "ГВС ТН" Then Call search("12 ГВС ТН", i, 38519)
        If Sheets(mainTab).Cells(i, 21) = "ХВС" Then Call search("12 ХВС", i, 38894)
    Next

    Message ("Готово")
    
End Sub

Private Sub search(sht, i, last)
    first = 2
    iskomoe = Sheets(mainTab).Cells(i, 7)
    Find = 0
        
    Do
        middle = first + Int((last - first) / 2)
        If Sheets(sht).Cells(first, 1) = iskomoe Then Find = first
        If Sheets(sht).Cells(middle, 1) = iskomoe Then Find = middle
        If Sheets(sht).Cells(last, 1) = iskomoe Then Find = last
        If Sheets(sht).Cells(middle, 1) > iskomoe Then last = middle
        If Sheets(sht).Cells(middle, 1) < iskomoe Then first = middle
    Loop Until last - first < 2 Or Find > 0
        
    If Find > 0 Then
        Sheets(mainTab).Cells(i, 12) = Sheets(sht).Cells(Find, 2)
        Sheets(mainTab).Cells(i, 13) = min(Sheets(mainTab).Cells(i, 11), Sheets(mainTab).Cells(i, 12))
        Sheets(mainTab).Cells(i, 15) = Sheets(sht).Cells(Find, 3)
        Sheets(mainTab).Cells(i, 16) = min(Sheets(mainTab).Cells(i, 14), Sheets(mainTab).Cells(i, 15))
    Else
        Sheets(mainTab).Cells(i, 12) = "-"
        Sheets(mainTab).Cells(i, 13) = "-"
        Sheets(mainTab).Cells(i, 15) = "-"
        Sheets(mainTab).Cells(i, 16) = "-"
    End If
        
End Sub

Function min(a, b)
    If a <= b Then min = a Else min = b
End Function

'Рисование прогресса, text - имя, cur - текущее значение, all - всего, отображать каждые over штук
Private Sub ProgressBar(text As String, ByVal cur As Long, ByVal all As Long)
    Application.ScreenUpdating = True
    Application.StatusBar = text + ":" + Str(cur) + " из" + Str(all) + _
        " (" + Str(Int(cur / all * 100)) + "% )"
        DoEvents
    Application.ScreenUpdating = False
End Sub


Private Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    DoEvents
    Application.ScreenUpdating = False
End Sub


