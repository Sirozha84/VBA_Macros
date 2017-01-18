Attribute VB_Name = "Module1"
Const First = 6     'Первая строка таблицы
Const Days = 36     'Количество дней
Const Result = 18   'Ячейка с остатком в накладной

Dim n As Integer
Dim d As Integer

Sub Обновить()
    n = 1
    d = 1
    'Рисуем таблицу
    Table
    'Заполняем
    Calc
    'Считаем сумму остатков
    All = 0
    For i = 1 To n * 2 - 2
        Sum = 0
        For j = 3 To Days + 2
            Sum = Sum + Cells(i + First + 1, j)
        Next
        Cells(i + First + 1, Days + 3) = Sum
        All = All + Sum
        If i / 2 = i \ 2 Then
            For j = 3 To Days + 2
                Cells(i + First + 1, j).Interior.Color = &HE0E0E0
            Next
        End If
    Next
    Cells(First + n * 2 + 1, Days + 3) = All
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
    'Колонка
    If Not Night Then Cells(First, d + 2) = Left(sh, Len(sh) - 1)
    For i = 6 To 16 '25
        'Берём строку, и проверяем, есть ли она уже в таблице
        st = Sheets(sh).Cells(i, 2)
        ost = Sheets(sh).Cells(i, Result)
        If st <> "" Then
            en = 0
            For j = 1 To n
                If Cells(First + j * 2, 2) = st Then en = j
            Next
            If en = 0 Then
                sn = n
                n = n + 1
            Else
                sn = en
            End If
            Cells(First + sn * 2, 1) = sn       'Номер
            Cells(First + sn * 2, 2) = st       'Наименование
            Cells(First + sn * 2 - Night, 2 + d) = ost  'Остаток
        End If
    Next
    If Night Then d = d + 1
    'd = d + 1
End Sub

Sub Table()
    Cells.Clear
    Cells(First, Days + 3) = "Итого"
End Sub
