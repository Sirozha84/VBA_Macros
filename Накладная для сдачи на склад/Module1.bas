Attribute VB_Name = "Module1"
Const First = 6     '������ ������ �������
Const Days = 36     '���������� ����
Const Result = 18   '������ � �������� � ���������

Dim n As Integer
Dim d As Integer

Sub ��������()
    n = 1
    d = 1
    '������ �������
    Table
    '���������
    Calc
    '������� ����� ��������
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
    '�������
    If Not Night Then Cells(First, d + 2) = Left(sh, Len(sh) - 1)
    For i = 6 To 16 '25
        '���� ������, � ���������, ���� �� ��� ��� � �������
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
            Cells(First + sn * 2, 1) = sn       '�����
            Cells(First + sn * 2, 2) = st       '������������
            Cells(First + sn * 2 - Night, 2 + d) = ost  '�������
        End If
    Next
    If Night Then d = d + 1
    'd = d + 1
End Sub

Sub Table()
    Cells.Clear
    Cells(First, Days + 3) = "�����"
End Sub
