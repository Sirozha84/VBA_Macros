Attribute VB_Name = "Summary"
'Сводная таблица
'Автор: Сергей гордеев
'Дата изменения: 31.01.2017

Const First = 5     'Первая строка таблицы
Const Days = 36     'Количество дней
Const NameCols = 7  'Количество колонок в наименовании
Const Result = 18   'Ячейка с остатком в накладной

Dim n As Integer
Dim d As Integer

Sub Refresh()
    'В случае покупки убрать
    If Not WorkIt Then Exit Sub
    
    n = 1
    d = 1
    Cells.Clear
    Cells(First, 1) = "Обработка..."
    Application.ScreenUpdating = False
    'Рисуем таблицу
    Table
    'Заполняем
    Calc
    'Считаем сумму остатков
    All = 0
    For i = 1 To n * 2 - 2
        Sum = 0
        For j = 3 To Days + 2
            Sum = Sum + Cells(i + First + 1, NameCols + j)
        Next
        Cells(i + First + 1, NameCols + Days + 2) = Sum
        All = All + Sum
        'Затемняем ночные строчки
        If i / 2 = i \ 2 Then
            For j = 3 To Days + 2
                Cells(i + First + 1, j + NameCols - 1).Interior.Color = &HE0E0E0
            Next
        End If
    Next
    'Итог
    Cells(First + n * 2, NameCols + Days + 2) = All
    Bottom
    Application.ScreenUpdating = True
End Sub

Sub Calc()
    For dy = 27 To 31
        AddList "-" + Trim(Str(dy)) + "д", False
        AddList "-" + Trim(Str(dy)) + "н", True
    Next
    For dy = 1 To 31
        AddList Trim(Str(dy)) + "д", False
        AddList Trim(Str(dy)) + "н", True
    Next
End Sub

Sub AddList(sh As String, Night As Boolean)
    'MsgBox """" + sh + """"
    Dim st(NameCols) As String
    'Строка в накладной
    Dim ost             'Остаток в накладной
    'Колонка (дата)
    If Not Night Then Cells(First + 1, 1 + NameCols + d) = Left(sh, Len(sh) - 1)
    For i = 6 To 16 '25
        'Берём с накладной для проверки
        For C = 1 To NameCols
            st(C) = Sheets(sh).Cells(i, C + 1)
        Next
        ost = Sheets(sh).Cells(i, Result)
        If st(1) <> "" Then
            en = 0
            'Проверяем, есть ли строка в отчёте
            For j = 1 To n
                complate = True
                For C = 1 To NameCols
                    If Cells(First + j * 2, 1 + C) <> st(C) Then complate = False
                Next
                If complate Then en = j
            Next
            If en = 0 Then
                sn = n
                n = n + 1
            Else
                sn = en
            End If
            Cells(First + sn * 2, 1) = sn       'Номер
            For C = 1 To NameCols               'Наименование
                Cells(First + sn * 2, 1 + C) = "'" + st(C)
            Next
            Cells(First + sn * 2 - Night, 1 + NameCols + d) = ost  'Остаток
        End If
    Next
    If Night Then d = d + 1
End Sub

Sub Table()
    'Объединение ячеек (делаем это перед заполнением, что бы не изменялся размер)
    For C = 1 To NameCols + 1
        range(Cells(First, C), Cells(First + 1, C)).Merge
        Cells(First, C).HorizontalAlignment = xlCenter
        Cells(First, C).VerticalAlignment = xlCenter
        Cells(First, C).WrapText = True
    Next
    'Шапка таблицы
    Cells(First, 1) = "№"
    Cells(First, 2 + NameCols) = "Дата"
    For C = 1 To NameCols
        Cells(First, 1 + C) = Sheets("1д").Cells(4, 1 + C)
    Next
    Cells(First, NameCols + Days + 2) = "Итого"
    range(Cells(First, 1), Cells(First + 1, NameCols + Days + 2)).Interior.Color = &HE0E0E0
    'Дата
    range(Cells(First, NameCols + 2), Cells(First, NameCols + Days + 1)).Merge
    Cells(First, NameCols + 2).HorizontalAlignment = xlCenter
    Cells(First, NameCols + 2).VerticalAlignment = xlCenter
End Sub

Sub Bottom()
    'Подвал
    'Рамка
    last = First + (n - 1) * 2 + 2
    range(Cells(First, 1), Cells(last, NameCols + Days + 2)).Borders.Weight = xlThin
    'Красота в ячейках наименования
    For i = First + 2 To First + (n - 1) * 2 Step 2
        For C = 1 To NameCols + 1
            range(Cells(i, C), Cells(i + 1, C)).Merge
            Cells(i, C).HorizontalAlignment = xlCenter
            Cells(i, C).VerticalAlignment = xlCenter
        Next
    Next
    'Итого
    range(Cells(First, NameCols + Days + 2), Cells(First + 1, NameCols + Days + 2)).Merge
    Cells(First, NameCols + Days + 2).HorizontalAlignment = xlCenter
    Cells(First, NameCols + Days + 2).VerticalAlignment = xlCenter
    'Итогогого
    Cells(last, 1) = "Итого:"
    range(Cells(last, 1), Cells(last, NameCols + Days + 1)).Merge
    Cells(last, 1).HorizontalAlignment = xlRight
End Sub
