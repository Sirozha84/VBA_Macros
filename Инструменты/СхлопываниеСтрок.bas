Attribute VB_Name = "����������������"
'������ 1.1 (01.09.2020)

Const CountColumn = 0   '������� ��� �������� ����������, 0 - ���� �� �����

Sub ����������������()
    i = 2
    Do While Cells(i, 1) <> ""
            
        If CountColumn > 0 Then
            If Cells(i, CountColumn) = "" Then Cells(i, CountColumn) = 1
        End If
        
        If Cells(i, 1) = Cells(i + 1, 1) Then
               
            '�����
            c = 2: Cells(i, c) = Cells(i, c) + Cells(i + 1, c)
            
            '����������
            If CountColumn > 0 Then
                Cells(i, CountColumn) = Cells(i, CountColumn) + 1
            End If
            
            Rows(i + 1).EntireRow.Delete
        Else
            i = i + 1
        End If
    Loop
End Sub