Attribute VB_Name = "Summary"
'������� �������
'�����: ������ �������
'���� ���������: 31.01.2017

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
    For dy = 27 To 31
        AddList "-" + Trim(Str(dy)) + "�", False
        AddList "-" + Trim(Str(dy)) + "�", True
    Next
    For dy = 1 To 31
        AddList Trim(Str(dy)) + "�", False
        AddList Trim(Str(dy)) + "�", True
    Next
End Sub

Sub AddList(sh As String, Night As Boolean)
    'MsgBox """" + sh + """"
    Dim st(NameCols) As String
    '������ � ���������
    Dim ost             '������� � ���������
    '������� (����)
    If Not Night Then Cells(First + 1, 1 + NameCols + d) = Left(sh, Len(sh) - 1)
    For i = 6 To 16 '25
        '���� � ��������� ��� ��������
        For C = 1 To NameCols
            st(C) = Sheets(sh).Cells(i, C + 1)
        Next
        ost = Sheets(sh).Cells(i, Result)
        If st(1) <> "" Then
            en = 0
            '���������, ���� �� ������ � ������
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
            Cells(First + sn * 2, 1) = sn       '�����
            For C = 1 To NameCols               '������������
                Cells(First + sn * 2, 1 + C) = "'" + st(C)
            Next
            Cells(First + sn * 2 - Night, 1 + NameCols + d) = ost  '�������
        End If
    Next
    If Night Then d = d + 1
End Sub

Sub Table()
    '����������� ����� (������ ��� ����� �����������, ��� �� �� ��������� ������)
    For C = 1 To NameCols + 1
        range(Cells(First, C), Cells(First + 1, C)).Merge
        Cells(First, C).HorizontalAlignment = xlCenter
        Cells(First, C).VerticalAlignment = xlCenter
        Cells(First, C).WrapText = True
    Next
    '����� �������
    Cells(First, 1) = "�"
    Cells(First, 2 + NameCols) = "����"
    For C = 1 To NameCols
        Cells(First, 1 + C) = Sheets("1�").Cells(4, 1 + C)
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
        For C = 1 To NameCols + 1
            range(Cells(i, C), Cells(i + 1, C)).Merge
            Cells(i, C).HorizontalAlignment = xlCenter
            Cells(i, C).VerticalAlignment = xlCenter
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
