Attribute VB_Name = "Summary"
'������� �������
'�����: ������ �������
'���� ���������: 23.01.2017

Const First = 5     '������ ������ �������
Const Days = 36     '���������� ����
Const NameCols = 7  '���������� ������� � ������������
Const Result = 18   '������ � �������� � ���������

Dim n As Integer
Dim d As Integer

Sub Refresh()
    '� ������ ������� ������
    If Not WorkIt Then Exit Sub
    
    n = 1
    d = 1
    Cells.Clear
    Cells(First, 1) = "���������..."
    Application.ScreenUpdating = False
    '������ �������
    Table
    '���������
    Calc
    '������� ����� ��������
    All = 0
    For i = 1 To n * 2 - 2
        Sum = 0
        For j = 3 To Days + 2
            Sum = Sum + Cells(i + First + 1, NameCols + j)
        Next
        Cells(i + First + 1, NameCols + Days + 2) = Sum
        All = All + Sum
        '��������� ������ �������
        If i / 2 = i \ 2 Then
            For j = 3 To Days + 2
                Cells(i + First + 1, j + NameCols - 1).Interior.Color = &HE0E0E0
            Next
        End If
    Next
    '����
    Cells(First + n * 2, NameCols + Days + 2) = All
    Bottom
    Application.ScreenUpdating = True
End Sub

Sub Calc()
    Call AddList("-27�", False)
    Call AddList("-27�", True)
    Call AddList("-28�", False)
    Call AddList("-28�", True)
    Call AddList("-29�", False)
    Call AddList("-29�", True)
    Call AddList("-30�", False)
    Call AddList("-30�", True)
    Call AddList("-31�", False)
    Call AddList("-31�", True)
    Call AddList("1�", False)
    Call AddList("1�", True)
    Call AddList("2�", False)
    Call AddList("2�", True)
    Call AddList("3�", False)
    Call AddList("3�", True)
    Call AddList("4�", False)
    Call AddList("4�", True)
    Call AddList("5�", False)
    Call AddList("5�", True)
    Call AddList("6�", False)
    Call AddList("6�", True)
    Call AddList("7�", False)
    Call AddList("7�", True)
    Call AddList("8�", False)
    Call AddList("8�", True)
    Call AddList("9�", False)
    Call AddList("9�", True)
    Call AddList("10�", False)
    Call AddList("10�", True)
    Call AddList("11�", False)
    Call AddList("11�", True)
    Call AddList("12�", False)
    Call AddList("12�", True)
    Call AddList("13�", False)
    Call AddList("13�", True)
    Call AddList("14�", False)
    Call AddList("14�", True)
    Call AddList("15�", False)
    Call AddList("15�", True)
    Call AddList("16�", False)
    Call AddList("16�", True)
    Call AddList("17�", False)
    Call AddList("17�", True)
    Call AddList("18�", False)
    Call AddList("18�", True)
    Call AddList("19�", False)
    Call AddList("19�", True)
    Call AddList("20�", False)
    Call AddList("20�", True)
    Call AddList("21�", False)
    Call AddList("21�", True)
    Call AddList("22�", False)
    Call AddList("22�", True)
    Call AddList("23�", False)
    Call AddList("23�", True)
    Call AddList("24�", False)
    Call AddList("24�", True)
    Call AddList("25�", False)
    Call AddList("25�", True)
    Call AddList("26�", False)
    Call AddList("26�", True)
    Call AddList("27�", False)
    Call AddList("27�", True)
    Call AddList("28�", False)
    Call AddList("28�", True)
    Call AddList("29�", False)
    Call AddList("29�", True)
    Call AddList("30�", False)
    Call AddList("30�", True)
    Call AddList("31�", False)
    Call AddList("31�", True)
End Sub

Sub AddList(sh As String, Night As Boolean)
    Dim st(NameCols) As String
    '������ � ���������
    Dim ost             '������� � ���������
    '������� (����)
    If Not Night Then Cells(First + 1, 1 + NameCols + d) = Left(sh, Len(sh) - 1)
    For i = 6 To 16 '25
        '���� � ��������� ��� ��������
        For c = 1 To NameCols
            st(c) = Sheets(sh).Cells(i, c + 1)
        Next
        ost = Sheets(sh).Cells(i, Result)
        If st(1) <> "" Then
            en = 0
            '���������, ���� �� ������ � ������
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
            Cells(First + sn * 2, 1) = sn       '�����
            For c = 1 To NameCols               '������������
                Cells(First + sn * 2, 1 + c) = "'" + st(c)
            Next
            Cells(First + sn * 2 - Night, 1 + NameCols + d) = ost  '�������
        End If
    Next
    If Night Then d = d + 1
End Sub

Sub Table()
    '����������� ����� (������ ��� ����� �����������, ��� �� �� ��������� ������)
    For c = 1 To NameCols + 1
        range(Cells(First, c), Cells(First + 1, c)).Merge
        Cells(First, c).HorizontalAlignment = xlCenter
        Cells(First, c).VerticalAlignment = xlCenter
        Cells(First, c).WrapText = True
    Next
    '����� �������
    Cells(First, 1) = "�"
    Cells(First, 2 + NameCols) = "����"
    For c = 1 To NameCols
        Cells(First, 1 + c) = Sheets("1�").Cells(4, 1 + c)
    Next
    Cells(First, NameCols + Days + 2) = "�����"
    range(Cells(First, 1), Cells(First + 1, NameCols + Days + 2)).Interior.Color = &HE0E0E0
    '����
    range(Cells(First, NameCols + 2), Cells(First, NameCols + Days + 1)).Merge
    Cells(First, NameCols + 2).HorizontalAlignment = xlCenter
    Cells(First, NameCols + 2).VerticalAlignment = xlCenter
End Sub

Sub Bottom()
    '������
    '�����
    last = First + (n - 1) * 2 + 2
    range(Cells(First, 1), Cells(last, NameCols + Days + 2)).Borders.Weight = xlThin
    '������� � ������� ������������
    For i = First + 2 To First + (n - 1) * 2 Step 2
        For c = 1 To NameCols + 1
            range(Cells(i, c), Cells(i + 1, c)).Merge
            Cells(i, c).HorizontalAlignment = xlCenter
            Cells(i, c).VerticalAlignment = xlCenter
        Next
    Next
    '�����
    range(Cells(First, NameCols + Days + 2), Cells(First + 1, NameCols + Days + 2)).Merge
    Cells(First, NameCols + Days + 2).HorizontalAlignment = xlCenter
    Cells(First, NameCols + Days + 2).VerticalAlignment = xlCenter
    '���������
    Cells(last, 1) = "�����:"
    range(Cells(last, 1), Cells(last, NameCols + Days + 1)).Merge
    Cells(last, 1).HorizontalAlignment = xlRight
End Sub
