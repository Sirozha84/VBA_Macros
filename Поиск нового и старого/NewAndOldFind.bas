Attribute VB_Name = "NewAndOldFind"
'������ 1.3 (23.12.2019) - ��������� ���� �� ������
'������ 1.4 (24.12.2019) - ����������� ���������
'������ 1.5 (24.12.2019) - ����� ���������
'������ 1.6 (26.12.2019) - �����������

Const newTab = "������"     '����� �������
Const fNew = 1              '������� � "�����" �������
Const oldTab = "��������"   '������ ������� �������
Const fOld = 1              '������� � "������" ��������
Const fields = 2            '���������� ������� ��� ��������
Const maxTabs = 8           '�������� �����

Global sn As Integer        '������� ����� �����
Global max As Integer       '������� ����� �����
Global oldTabs As Integer   '������� ������ ������


Sub NewAndOldFind()

    MakeCopy
    
    '����� �������� � ����� �������
    AddDead (oldTab)
    '��� ������ � ������ �������� �������� � ������ ����:
    'AddNew ("")
    
    '����� ����� �������
    FindNew (oldTab)
    '��� ������ � ������ �������� �������� � ������ ����:
    'FindDead ("")
    
    DeadSum
    Message "������!"
    
End Sub

'���������� �������� �������
Private Sub MakeCopy()
    
    Message "����������..."
    
    sn = 0
    mx = 0
    oldTabs = 0
    
    
    Sheets("Res").Cells.Clear
    maxStrings = 0
    max = 0
    Do While Sheets(newTab).Cells(max + 1, 1) <> ""
        max = max + 1
        For i = 1 To maxTabs
            Sheets("Res").Cells(max, i) = Sheets(newTab).Cells(max, i)
        Next
    Loop

End Sub

'����� � ���������� �������� �����
Private Sub AddDead(sheet)
    
    '������� �������� � ������ �������
    Message "������� �����..."
    maxOld = 0
    Do Until Sheets(sheet).Cells(maxOld + 1, 1) = ""
        maxOld = maxOld + 1
    Loop
    
    changed = 0
    Find = False
    
    For i = 2 To maxOld
        Call ProgressBar("����� ��������", i, maxOld)
        For j = 2 To max
            Find = True
            For k = 0 To fields - 1
                If Sheets("Res").Cells(j, fNew + k) <> Sheets(sheet).Cells(i, fOld + k) Then
                    Find = False
                    Exit For
                End If
            Next
            If Find Then
                Exit For
            End If
        Next
        If Find Then
            '�������������� �������� ��� ����������
                
                
                
            '���������
            Change = False
            For k = 3 To 8
                'If Sheets("Res").Cells(j, k) = "" Then Sheets("Res").Cells(j, k) = "_"
                'If Sheets(sheet).Cells(i, k) = "" Then Sheets(sheet).Cells(i, k) = "_"
                If Sheets("Res").Cells(j, k) <> Sheets(sheet).Cells(i, k) Then
                    Sheets("Res").Cells(j, k).Interior.Color = vbRed
                    Change = True
                End If
            Next
            If Change Then
                Sheets("Res").Cells(j, maxTabs + 3) = "������"
                changed = changed + 1
            End If
            
            
                
        Else
            For c = 1 To maxTabs
                Sheets("Res").Cells(max, c) = Sheets(sheet).Cells(i, c)
            Next
            Sheets("Res").Cells(max, maxTabs + 2) = "����� (��� �" + sheet + ", �� �� ����� � " + newTab + ")"
            max = max + 1
            sn = sn + 1
        End If
    Next
    
    Sheets("Res").Cells(max + 1, maxTabs + 2) = "�������:" + Str(sn)
    Sheets("Res").Cells(max + 1, maxTabs + 3) = "��������:" + Str(changed)
    
End Sub

'����� ����� �����
Private Sub FindNew(sheet)
    
    oldTabs = oldTabs + 1
    
    '������� �������� � ������ �������
    Message "������� �����..."
    maxOld = 0
    Do Until Sheets(sheet).Cells(maxOld + 1, 1) = ""
        maxOld = maxOld + 1
    Loop
    
    
    For i = 1 To max - sn - 1
        
        Call ProgressBar("����� �����", i, max - sn - 1)
        
        Find = False
        For j = 1 To maxOld
            If Sheets("Res").Cells(i, fNew) = Sheets(sheet).Cells(j, fOld) Then
                Find = True
                Exit For
            End If
        Next
        If Not Find Then
            If Sheets("Res").Cells(i, maxTabs + 1) = "" Then
                Sheets("Res").Cells(i, maxTabs + 1) = 1
            Else
                Sheets("Res").Cells(i, maxTabs + 1) = Val(Sheets("Res").Cells(i, maxTabs + 1)) + 1
            End If
        End If
    Next
    
    
End Sub

'���� �� �����
Private Sub DeadSum()
    
    Message "����������..."
    
    s = 0
    For i = 1 To max
        If Sheets("Res").Cells(i, maxTabs + 1) = oldTabs Then
            Sheets("Res").Cells(i, maxTabs + 1) = "�����! (�������� � " + newTab + ", ������ �� ����������)"
            s = s + 1
        Else
            Sheets("Res").Cells(i, maxTabs + 1) = ""
        End If
    Next
    Sheets("Res").Cells(max + 1, maxTabs + 1) = "�����:" + Str(s)
    
End Sub

Private Sub ProgressBar(text As String, ByVal cur As Integer, ByVal all As Integer)
    If cur Mod 50 = 0 Then
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



'��������� ���� �� ������ � ��������� ��� � ��������� ������
Sub CodeEject()
    
    codeFrom = 1
    codeTo = 5

    Application.ScreenUpdating = True
    Application.StatusBar = "������� �����..."
    Application.ScreenUpdating = False
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
