Attribute VB_Name = "Misc"
'���������� ������, ����������� �� Access
Sub PrepareTabFromAccess()
    
    Message "������� �����..."
    max = 1
    
    Do Until Cells(max + 1, 1) = ""
        max = max + 1
    Loop
    
    For i = 2 To max
        
        Call ProgressBar("���������", i, max)
        
        '������� � �����
        Cells(i, 1).NumberFormat = "@"
        
        '�������� �� ������ �������� ������� ������
        Cells(i, 1) = Replace(Cells(i, 1), "��������������� �������� �� � ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "�������  �� � ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "������� � ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "������� �� �  ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "������� �� � ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "������� �� � ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "������� �� �", "")
        Cells(i, 1) = Replace(Cells(i, 1), "�������� �� �  ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "�������� �� � ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "������������� �������� �� � ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "������������� ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "������������� ", "")
        
        '�������� �����
        s = Cells(i, 1)
        If s <> "" Then
            Do While Left(s, 1) = "0"
                s = Right(s, Len(s) - 1)
            Loop
            Cells(i, 1) = s
        End If
        
        '�������� "������"
        s = Replace(Cells(i, 3), "�.", "")
        If s <> "" Then Cells(i, 3) = CDbl(s)
        
        '������������ ��� (���� ����)
        If Cells(1, 4) = -1 Then
            Cells(i, 3) = Cells(i, 3) + Cells(i, 4)
            Cells(i, 4) = ""
        End If

    Next	

    For i = 2 To max + 5        

        '�����������
        If Cells(i, 1) <> "" And Cells(i, 1) <> "." And Cells(i, 1) = Cells(i + 1, 1) Then
           Cells(i, 3) = Cells(i, 3) + Cells(i + 1, 3)
           Rows(i + 1).EntireRow.Delete
           i = i - 1
        End If
        
        '������ �������
        If Cells(i, 1) = "" Then Rows(i).EntireRow.Delete
        
    Next
    
    Message "������!"
    
End Sub

'��������� ���� �� ������ � ��������� ��� � ��������� ������
Sub CodeEject()
    
    codeFrom = 1    '���� ��� ����� ���
    codeTo = 5      '���� ���� �������� ������������

    Message "������� �����..."
    max = 1
    Do Until Cells(max + 1, 1) = ""
        max = max + 1
    Loop
    
    For i = 2 To max
        Call ProgressBar("���������", i, max)
        c = Cells(i, codeFrom)
        If InStr(1, c, "���: ", vbTextCompare) Then
            ss = Split(c, "���: ")
            If UBound(ss) > 0 Then
                c = Replace(ss(1), " ", "")
                kv = InStr(1, c, """", vbTextCompare)
                If kv > 0 Then c = Left(c, kv - 1)
            End If
        End If
        If InStr(1, c, "��� �������: ", vbTextCompare) Then
            ss = Split(c, "��� �������: ")
            If UBound(ss) > 0 Then
                c = Replace(ss(1), " ", "")
                kv = InStr(1, c, """", vbTextCompare)
                If kv > 0 Then c = Left(s, kv - 1)
            End If
        End If
        If InStr(1, c, "(", vbTextCompare) Then
            ss = Split(c, "(")
            If UBound(ss) > 0 Then
                c = Replace(ss(1), " ", "")
                kv = InStr(1, c, "(", vbTextCompare)
                If kv > 0 Then c = Left(c, kv - 1)
                kv = InStr(1, c, ")", vbTextCompare)
                If kv > 0 Then c = Left(c, kv - 1)
            End If
        End If
        
        '����� ������� ���� ��������� �� �� �� ��� ����� �����
        d = True
        For j = 1 To Len(c)
            code = Asc(Mid(c, j, 1))
            If code < 48 Or code > 57 Then d = False
        Next
        If Not d Then c = ""
        Cells(i, codeTo) = c
    Next
    
    Message "������!"

End Sub

'����� ������� ���������
Private Sub ProgressBar(text As String, ByVal cur As Integer, ByVal all As Integer)
    If cur Mod 100 = 0 Then
        Application.ScreenUpdating = True
        Application.StatusBar = text + ":" + _
            Str(cur) + " ��" + Str(all) + " (" + Str(Int(cur / all * 100)) + "% )"
        Application.ScreenUpdating = False
    End If
End Sub

'����� ��������� � ���������
Private Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    Application.ScreenUpdating = False
End Sub


