Attribute VB_Name = "Find"
'������ 1.0 (11.12.2010)
'������ 1.1 (12.12.2019) - ����������� �������� �������� ������

Sub Start()
    
    Filter

End Sub

'���������� �� ����������
Private Sub Filter()
    
    rang = 200000   '������������ ������� � ���������
    src = "10.2019" '�������� � ����������
    res = "Result"  '�������� � �����������
    c_adr = 8       '���� � �������
    c_usl = 17      '���� � �������
    c_vip = 22      '���������� �����
    c_fld = 24      '���������� ����� � �������
    
    ReDim tmp(rang) As String
    
    Sheets(res).Cells.Clear
    
    '������ ������������� ��������� ������
    Sheets(res).Cells(1, 1) = "����������..."
    rec = 1
    Max = 1
    For i = 2 To rang
        If Sheets(src).Cells(i, 1) = "" Then
            rec = i
            Exit For
        End If
        If Sheets(src).Cells(i, c_usl) = "���������" And Sheets(src).Cells(i, c_vip) <> 0 Then
            tmp(Max) = Sheets(src).Cells(i, c_adr)
            Max = Max + 1
        End If
    Next
   
    
    '���� ������, ������� �������� ��� ������
    Sheets(res).Cells(1, 1) = "����������..."
    Application.ScreenUpdating = False
    f = 1
    For i = 2 To rang
        If Sheets(src).Cells(i, 1) = "" Then Exit For
        Call StatusBar("����������", i, rec)
        
        'fnd = False 'IsInArray(Sheets(src).Cells(i, 8), tmp)
        fnd = False
        adr = Sheets(src).Cells(i, 8)
        If adr = last Then
            fnd = True
        Else
            For j = 1 To Max
                If adr = tmp(j) Then
                    fnd = True
                    last = adr
                    Exit For
                End If
            Next
        End If
        If fnd Then
            For c = 1 To c_fld
                Sheets(res).Cells(f + 1, c) = Sheets(src).Cells(i, c)
            Next
            f = f + 1
        End If
    Next
    
    '�������� �����
    For i = 1 To c_fld
        Sheets(res).Cells(1, i) = Sheets(src).Cells(1, i)
    Next
    
End Sub

Private Sub ProgressBar(text As String, cur As Integer, all As Integer)
    If cur Mod 50 = 0 Then
        Message text + ":" + cur + "��" + all + "(" + Str(Int(cur / all * 100)) + "%"
    End If
End Sub

Private Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    Application.ScreenUpdating = False
End Sub
