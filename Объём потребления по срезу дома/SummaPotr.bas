Attribute VB_Name = "SummaPotr"
'������ 1.1 (03.01.2020) - ���������� ������ �������������

Sub SummaPotr()
    
    res = "Temp"  '��������
    v1 = 11     '����� ����������� ���
    v2 = 12     '����� ����������� ��������
    v3 = 13     '����� ����������� ��
    sf = 20     '���� ��� ������ �����
    max = 37837 '������������ ���������� �������
    
    Sum = 0
    first = 2
    a1 = Sheets(res).Cells(2, 1)
    a2 = Sheets(res).Cells(2, 2)
    a3 = Sheets(res).Cells(2, 3)
    a4 = Sheets(res).Cells(2, 4)
    For i = 2 To max
        If Sheets(res).Cells(i, 1) = "" Then Exit For
        If a1 = Sheets(res).Cells(i, 1) And _
           a2 = Sheets(res).Cells(i, 2) And _
           a3 = Sheets(res).Cells(i, 3) And _
           a4 = Sheets(res).Cells(i, 4) Then
            Sum = Sum + Sheets(res).Cells(i, v1) + _
                Sheets(res).Cells(i, v2) + _
                Sheets(res).Cells(i, v3)
        Else
            a1 = Sheets(res).Cells(i, 1)
            a2 = Sheets(res).Cells(i, 2)
            a3 = Sheets(res).Cells(i, 3)
            a4 = Sheets(res).Cells(i, 4)
            For j = first To i - 1
                Sheets(res).Cells(j, sf) = Sum
            Next
            Sum = 0
            first = i
            i = i - 1
        End If
    Next
    
    MsgBox ("������!")

End Sub
