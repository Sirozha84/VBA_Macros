Attribute VB_Name = "Module1"
Const First = 5     'Ïåðâàÿ ñòðîêà òàáëèöû
Const Days = 36     'Êîëè÷åñòâî äíåé
Const NameCols = 7  'Êîëè÷åñòâî êîëîíîê â íàèìåíîâàíèè
Const Result = 18   'ß÷åéêà ñ îñòàòêîì â íàêëàäíîé

Dim n As Integer
Dim d As Integer

Sub Îáíîâèòü()
'Çäåñü íàõîäèòñÿ Âàø êîä
    n = 1
    d = 1
    'Ðèñóåì òàáëèöó
    
    Cells.Clear
    Cells(First, 1) = "Îáðàáîòêà..."
    Application.ScreenUpdating = False
    Table
    'Çàïîëíÿåì
    Calc
    'Ñ÷èòàåì ñóììó îñòàòêîâ
    All = 0
    For i = 1 To n * 2 - 2
        Sum = 0
        For j = 3 To Days + 2
            Sum = Sum + Cells(i + First + 1, NameCols + j)
        Next
        Cells(i + First + 1, NameCols + Days + 2) = Sum
        All = All + Sum
        'Çàòåìíÿåì íî÷íûå ñòðî÷êè
        If i / 2 = i \ 2 Then
            For j = 3 To Days + 2
                Cells(i + First + 1, j + NameCols - 1).Interior.Color = &HE0E0E0
            Next
        End If
    Next
    'Èòîã
    Cells(First + n * 2, NameCols + Days + 2) = All
    Bottom
    Application.ScreenUpdating = True
End Sub

Sub Calc()
    Call AddList("-27ä", False)
    Call AddList("-27í", True)
    Call AddList("-28ä", False)
    Call AddList("-28í", True)
    Call AddList("-29ä", False)
    Call AddList("-29í", True)
    Call AddList("-30ä", False)
    Call AddList("-30í", True)
    Call AddList("-31ä", False)
    Call AddList("-31í", True)
    Call AddList("1ä", False)
    Call AddList("1í", True)
    Call AddList("2ä", False)
    Call AddList("2í", True)
    Call AddList("3ä", False)
    Call AddList("3í", True)
    Call AddList("4ä", False)
    Call AddList("4í", True)
    Call AddList("5ä", False)
    Call AddList("5í", True)
    Call AddList("6ä", False)
    Call AddList("6í", True)
    Call AddList("7ä", False)
    Call AddList("7í", True)
    Call AddList("8ä", False)
    Call AddList("8í", True)
    Call AddList("9ä", False)
    Call AddList("9í", True)
    Call AddList("10ä", False)
    Call AddList("10í", True)
    Call AddList("11ä", False)
    Call AddList("11í", True)
    Call AddList("12ä", False)
    Call AddList("12í", True)
    Call AddList("13ä", False)
    Call AddList("13í", True)
    Call AddList("14ä", False)
    Call AddList("14í", True)
    Call AddList("15ä", False)
    Call AddList("15í", True)
    Call AddList("16ä", False)
    Call AddList("16í", True)
    Call AddList("17ä", False)
    Call AddList("17í", True)
    Call AddList("18ä", False)
    Call AddList("18í", True)
    Call AddList("19ä", False)
    Call AddList("19í", True)
    Call AddList("20ä", False)
    Call AddList("20í", True)
    Call AddList("21ä", False)
    Call AddList("21í", True)
    Call AddList("22ä", False)
    Call AddList("22í", True)
    Call AddList("23ä", False)
    Call AddList("23í", True)
    Call AddList("24ä", False)
    Call AddList("24í", True)
    Call AddList("25ä", False)
    Call AddList("25í", True)
    Call AddList("26ä", False)
    Call AddList("26í", True)
    Call AddList("27ä", False)
    Call AddList("27í", True)
    Call AddList("28ä", False)
    Call AddList("28í", True)
    Call AddList("29ä", False)
    Call AddList("29í", True)
    Call AddList("30ä", False)
    Call AddList("30í", True)
    Call AddList("31ä", False)
    Call AddList("31í", True)
End Sub

Sub AddList(sh As String, Night As Boolean)
    Dim st(NameCols)    'Ñòðîêà â íàêëàäíîé
    Dim ost             'Îñòàòîê â íàêëàäíîé
    'Êîëîíêà (äàòà)
    If Not Night Then Cells(First + 1, 1 + NameCols + d) = Left(sh, Len(sh) - 1)
    For i = 6 To 16 '25
        'Áåð¸ì ñ íàêëàäíîé äëÿ ïðîâåðêè
        For c = 1 To NameCols
            st(c) = Sheets(sh).Cells(i, c + 1)
        Next
        ost = Sheets(sh).Cells(i, Result)
        If st(1) <> "" Then
            en = 0
            'Ïðîâåðÿåì, åñòü ëè ñòðîêà â îò÷¸òå
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
            Cells(First + sn * 2, 1) = sn       'Íîìåð
            For c = 1 To NameCols               'Íàèìåíîâàíèå
                Cells(First + sn * 2, 1 + c) = st(c)
            Next
            Cells(First + sn * 2 - Night, 1 + NameCols + d) = ost  'Îñòàòîê
        End If
    Next
    If Night Then d = d + 1
End Sub

Sub Table()
    'Îáúåäèíåíèå ÿ÷ååê (äåëàåì ýòî ïåðåä çàïîëíåíèåì, ÷òî áû íå èçìåíÿëñÿ ðàçìåð)
    For c = 1 To NameCols + 1
        range(Cells(First, c), Cells(First + 1, c)).Merge
        Cells(First, c).HorizontalAlignment = xlCenter
        Cells(First, c).VerticalAlignment = xlCenter
        Cells(First, c).WrapText = True
    Next
    'Øàïêà òàáëèöû
    Cells(First, 1) = "¹"
    Cells(First, 2 + NameCols) = "Äàòà"
    For c = 1 To NameCols
        Cells(First, 1 + c) = Sheets("1ä").Cells(4, 1 + c)
    Next
    Cells(First, NameCols + Days + 2) = "Èòîãî"
    range(Cells(First, 1), Cells(First + 1, NameCols + Days + 2)).Interior.Color = &HE0E0E0
    'Äàòà
    range(Cells(First, NameCols + 2), Cells(First, NameCols + Days + 1)).Merge
    Cells(First, NameCols + 2).HorizontalAlignment = xlCenter
    Cells(First, NameCols + 2).VerticalAlignment = xlCenter
End Sub

Sub Bottom()
    'Ïîäâàë
    'Ðàìêà
    last = First + (n - 1) * 2 + 2
    range(Cells(First, 1), Cells(last, NameCols + Days + 2)).Borders.Weight = xlThin
    'Êðàñîòà â ÿ÷åéêàõ íàèìåíîâàíèÿ
    For i = First + 2 To First + (n - 1) * 2 Step 2
        For c = 1 To NameCols + 1
            range(Cells(i, c), Cells(i + 1, c)).Merge
            Cells(i, c).HorizontalAlignment = xlCenter
            Cells(i, c).VerticalAlignment = xlCenter
        Next
    Next
    'Èòîãî
    range(Cells(First, NameCols + Days + 2), Cells(First + 1, NameCols + Days + 2)).Merge
    Cells(First, NameCols + Days + 2).HorizontalAlignment = xlCenter
    Cells(First, NameCols + Days + 2).VerticalAlignment = xlCenter
    'Èòîãîãîãî
    Cells(last, 1) = "Èòîãî:"
    range(Cells(last, 1), Cells(last, NameCols + Days + 1)).Merge
    Cells(last, 1).HorizontalAlignment = xlRight
End Sub
