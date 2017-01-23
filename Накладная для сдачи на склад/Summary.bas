Attribute VB_Name = "Summary"
'Сводная таблица
'Автор: Сергей гордеев
'Дата изменения: 23.01.2017

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
    Call AddList("-27д", False)
    Call AddList("-27н", True)
    Call AddList("-28д", False)
    Call AddList("-28н", True)
    Call AddList("-29д", False)
    Call AddList("-29н", True)
    Call AddList("-30д", False)
    Call AddList("-30н", True)
    Call AddList("-31д", False)
    Call AddList("-31н", True)
    Call AddList("1д", False)
    Call AddList("1н", True)
    Call AddList("2д", False)
    Call AddList("2н", True)
    Call AddList("3д", False)
    Call AddList("3н", True)
    Call AddList("4д", False)
    Call AddList("4н", True)
    Call AddList("5д", False)
    Call AddList("5н", True)
    Call AddList("6д", False)
    Call AddList("6н", True)
    Call AddList("7д", False)
    Call AddList("7н", True)
    Call AddList("8д", False)
    Call AddList("8н", True)
    Call AddList("9д", False)
    Call AddList("9н", True)
    Call AddList("10д", False)
    Call AddList("10н", True)
    Call AddList("11д", False)
    Call AddList("11н", True)
    Call AddList("12д", False)
    Call AddList("12н", True)
    Call AddList("13д", False)
    Call AddList("13н", True)
    Call AddList("14д", False)
    Call AddList("14н", True)
    Call AddList("15д", False)
    Call AddList("15н", True)
    Call AddList("16д", False)
    Call AddList("16н", True)
    Call AddList("17д", False)
    Call AddList("17н", True)
    Call AddList("18д", False)
    Call AddList("18н", True)
    Call AddList("19д", False)
    Call AddList("19н", True)
    Call AddList("20д", False)
    Call AddList("20н", True)
    Call AddList("21д", False)
    Call AddList("21н", True)
    Call AddList("22д", False)
    Call AddList("22н", True)
    Call AddList("23д", False)
    Call AddList("23н", True)
    Call AddList("24д", False)
    Call AddList("24н", True)
    Call AddList("25д", False)
    Call AddList("25н", True)
    Call AddList("26д", False)
    Call AddList("26н", True)
    Call AddList("27д", False)
    Call AddList("27н", True)
    Call AddList("28д", False)
    Call AddList("28н", True)
    Call AddList("29д", False)
    Call AddList("29н", True)
    Call AddList("30д", False)
    Call AddList("30н", True)
    Call AddList("31д", False)
    Call AddList("31н", True)
End Sub

Sub AddList(sh As String, Night As Boolean)
    Dim st(NameCols) As String
    'Строка в накладной
    Dim ost             'Остаток в накладной
    'Колонка (дата)
    If Not Night Then Cells(First + 1, 1 + NameCols + d) = Left(sh, Len(sh) - 1)
    For i = 6 To 16 '25
        'Берём с накладной для проверки
        For c = 1 To NameCols
            st(c) = Sheets(sh).Cells(i, c + 1)
        Next
        ost = Sheets(sh).Cells(i, Result)
        If st(1) <> "" Then
            en = 0
            'Проверяем, есть ли строка в отчёте
            For j = 1 To n
                complate = True
                For c = 1 To NameCols
                    If Cells(First + j * 2, 1 + c) <> st(c) Then complate = False
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
            For c = 1 To NameCols               'Наименование
                Cells(First + sn * 2, 1 + c) = "'" + st(c)
            Next
            Cells(First + sn * 2 - Night, 1 + NameCols + d) = ost  'Остаток
        End If
    Next
    If Night Then d = d + 1
End Sub

Sub Table()
    'Объединение ячеек (делаем это перед заполнением, что бы не изменялся размер)
    For c = 1 To NameCols + 1
        range(Cells(First, c), Cells(First + 1, c)).Merge
        Cells(First, c).HorizontalAlignment = xlCenter
        Cells(First, c).VerticalAlignment = xlCenter
        Cells(First, c).WrapText = True
    Next
    'Шапка таблицы
    Cells(First, 1) = "№"
    Cells(First, 2 + NameCols) = "Дата"
    For c = 1 To NameCols
        Cells(First, 1 + c) = Sheets("1д").Cells(4, 1 + c)
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
        For c = 1 To NameCols + 1
            range(Cells(i, c), Cells(i + 1, c)).Merge
            Cells(i, c).HorizontalAlignment = xlCenter
            Cells(i, c).VerticalAlignment = xlCenter
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
