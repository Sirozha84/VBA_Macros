Attribute VB_Name = "����������������"
Sub ����������������()
    i = 2
    Do While Cells(i, 1) <> ""
            
        If Cells(i, 2) = "" Then Cells(i, 2) = 1
        
        If Cells(i, 1) = Cells(i + 1, 1) Then
               
            '�����
            Cells(i, 3) = Cells(i, 3) + Cells(i + 1, 3)
            Cells(i, 4) = Cells(i, 4) + Cells(i + 1, 4)
            
            '����������
            Cells(i, 2) = Cells(i, 2) + 1
            
            Rows(i + 1).EntireRow.Delete
        Else
            i = i + 1
        End If
    Loop
End Sub

