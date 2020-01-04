Attribute VB_Name = "SummaPotr"
'Версия 1.0 (26.12.2019)
'Версия 1.1 (03.01.2020) - Исправлена ошибка суммирования
'Версия 1.2 (04.01.2020) - Исправлена ошибка последней строки, улучшено сопоставление адреса

Const v1 = 11   'Объём потребления ИПУ
Const v2 = 12   'Объём потребления Норматив
Const v3 = 13   'Поле для вывода суммы
Const sf = 20   'Поле для вывода суммы

Sub SummaPotr()
    
    Message "Обработка..."
    
    Sum = 0
    Dim i As Long
    
    i = 2
    last = Adress(i)
    first = 2
    Do While Cells(i, 1) <> ""
        Sum = Sum + Cells(i, v1) + Cells(i, v2) + Cells(i, v3)
        If last <> Adress(i + 1) Then
            For j = first To i
                Cells(j, sf) = Sum
            Next
            last = Adress(i + 1)
            Sum = 0
            first = i + 1
        End If
        i = i + 1
    Loop
    
    Message "Готово!"

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

