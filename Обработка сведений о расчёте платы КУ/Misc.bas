Attribute VB_Name = "Misc"
'Дедупликация домов
Sub Dedupl()
    Application.ScreenUpdating = False
    j = 1
    last = ""
    For i = 2 To 42283
        Cells(j, 5) = Cells(i, 1)
        Cells(j, 6) = Cells(i, 2)
        If Cells(i, 1) <> last Then
            last = Cells(i, 1)
            
            j = j + 1
        End If
    Next
End Sub

'Подбор коэффициентов
Sub Koef()
    col = 25
    max = 170147
    For i = 1 To max
        If i Mod 1000 = 0 Then Call ProgressBar("Подбор коэффициентов", i, max)
        adr = Cells(i, 2) + CStr(Cells(i, 3)) + Cells(i, 4)
        For j = 1 To 1147
            If adr = Sheets("Counter").Cells(j, 5) Then
                Cells(i, col) = Sheets("Counter").Cells(j, 6)
                Exit For
            End If
        Next
    Next
    Message "Готово!"
End Sub
