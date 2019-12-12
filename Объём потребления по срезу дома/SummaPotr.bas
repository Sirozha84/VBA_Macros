Attribute VB_Name = "SummaPotr"
Sub SummaPotr()
    
    res = "ВО"  'Страница
    v1 = 10     'Объём потребления ИПУ
    v2 = 11     'Объём потребления Норматив
    v3 = 12     'Объём потребления РО
    sf = 20     'Поле для вывода суммы
    Max = 999999 'Максимальное количество записей
    
    Sum = 0
    first = 2
    a1 = Sheets(res).Cells(2, 1)
    a2 = Sheets(res).Cells(2, 2)
    a3 = Sheets(res).Cells(2, 3)
    a4 = Sheets(res).Cells(2, 3)
    For i = 2 To Max
        If Sheets(res).Cells(i, 1) = "" Then Exit For
        If a1 = Sheets(res).Cells(i, 1) And _
           a2 = Sheets(res).Cells(i, 2) And _
           a3 = Sheets(res).Cells(2, 3) And _
           a4 = Sheets(res).Cells(2, 4) Then
            Sum = Sum + Sheets(res).Cells(i, v1) + _
                Sheets(res).Cells(i, v2) + _
                Sheets(res).Cells(i, v3)
        Else
            a1 = Sheets(res).Cells(i, 1)
            a2 = Sheets(res).Cells(i, 2)
            a3 = Sheets(res).Cells(2, 3)
            a4 = Sheets(res).Cells(2, 4)
            For j = first To i - 1
                Sheets(res).Cells(j, sf) = Sum
            Next
            Sum = 0
            first = i
            i = i - 1
        End If
    Next
    
    MsgBox ("Готово!")

End Sub
