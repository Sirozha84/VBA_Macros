Attribute VB_Name = "СхлопываниеСтрок"
'Версия 1.1 (01.09.2020)

Const CountColumn = 0   'Колонка для подсчёта количества, 0 - если не нужно

Sub СхлопываниеСтрок()
    i = 2
    Do While Cells(i, 1) <> ""
            
        If CountColumn > 0 Then
            If Cells(i, CountColumn) = "" Then Cells(i, CountColumn) = 1
        End If
        
        If Cells(i, 1) = Cells(i + 1, 1) Then
               
            'Сумма
            c = 2: Cells(i, c) = Cells(i, c) + Cells(i + 1, c)
            
            'Количество
            If CountColumn > 0 Then
                Cells(i, CountColumn) = Cells(i, CountColumn) + 1
            End If
            
            Rows(i + 1).EntireRow.Delete
        Else
            i = i + 1
        End If
    Loop
End Sub