Attribute VB_Name = "SearchDifference"
'������ 1.0 (19.12.2019)
'������ 1.2 (20.12.2019)
'������ 1.3 (23.12.2019) - ��������� ���� �� ������
'������ 1.4 (24.12.2019) - ����������� ���������
'������ 1.5 (24.12.2019) - ����� ���������
'������ 1.6 (26.12.2019) - �����������
'������ 1.7 (10.01.2020) - ���������� �������

Const newTab = "���"     '����� �������
Const fNew = 1              '������� � "�����" �������
Const oldTab = "Access"     '������ ������� �������
Const fOld = 1              '������� � "������" ��������
Const fields = 1            '���������� ������� ��� ��������
Const maxTabs = 3           '�������� �����

Global sn As Integer        '������� ����� �����
Global max As Integer       '������� ����� �����
Global oldTabs As Integer   '������� ������ ������


Sub NewAndOldFind()

    MakeCopy
    AddDead (oldTab)
    FindNew (oldTab)
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
    
    mached = 0
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
            s = 3
            Sheets("Res").Cells(j, s + 1) = Sheets(sheet).Cells(i, s)
            i1 = Sheets("Res").Cells(j, s)
            i2 = Sheets(sheet).Cells(i, s)
            rz = Round(i1 - i2, 2)
            Sheets("Res").Cells(j, s + 2) = rz
            If Abs(rz) = 0 Then
                Sheets("Res").Cells(j, s + 0).Interior.Color = RGB(128, 255, 128)
                Sheets("Res").Cells(j, s + 1).Interior.Color = RGB(128, 255, 128)
                Sheets("Res").Cells(j, s + 2).Interior.Color = RGB(196, 255, 196)
                mached = mached + 1
                Sheets("Res").Cells(j, maxTabs + 5) = "������"
            End If
            If Abs(rz) > 0 Then
                Sheets("Res").Cells(j, s + 0).Interior.Color = RGB(255, 255, 128)
                Sheets("Res").Cells(j, s + 1).Interior.Color = RGB(255, 255, 128)
                Sheets("Res").Cells(j, s + 2).Interior.Color = RGB(255, 255, 196)
                mached = mached + 1
                Sheets("Res").Cells(j, maxTabs + 5) = "������ (�����)"
            End If
            If Abs(rz) > 10 Then
                Sheets("Res").Cells(j, s + 0).Interior.Color = RGB(255, 128, 128)
                Sheets("Res").Cells(j, s + 1).Interior.Color = RGB(255, 128, 128)
                Sheets("Res").Cells(j, s + 2).Interior.Color = RGB(255, 196, 196)
                changed = changed + 1
                Sheets("Res").Cells(j, maxTabs + 5) = "������"
            End If
            
        Else
            max = max + 1
            For c = 1 To maxTabs
                Sheets("Res").Cells(max, c) = Sheets(sheet).Cells(i, c)
            Next
            Sheets("Res").Cells(max, maxTabs + 5) = "����� (��� � " + sheet + ", �� �� ����� � " + newTab + ")"
            sn = sn + 1
        End If
    Next
    
    Sheets("Res").Cells(max + 3, maxTabs + 5) = "�������:" + Str(sn)
    Sheets("Res").Cells(max + 4, maxTabs + 5) = "�������:" + Str(mached)
    Sheets("Res").Cells(max + 5, maxTabs + 5) = "��������:" + Str(changed)
    
End Sub

'����� ����� �����
Private Sub FindNew(sheet)
    
    '������� �������� � ������ �������
    Message "������� �����..."
    maxOld = 0
    Do Until Sheets(sheet).Cells(maxOld + 1, 1) = ""
        maxOld = maxOld + 1
    Loop
    
    s = 0
    For i = 2 To max - sn
        
        Call ProgressBar("����� �����", i, max - sn - 1)
        
        Find = False
        For j = 2 To maxOld
            If Sheets("Res").Cells(i, fNew) = Sheets(sheet).Cells(j, fOld) Then
                Find = True
                Exit For
            End If
        Next
        If Not Find Then
            Sheets("Res").Cells(i, maxTabs + 5) = "�����! (�������� � " + newTab + ")"
            s = s + 1
        End If
    Next
    Sheets("Res").Cells(max + 2, maxTabs + 5) = "�����:" + Str(s)
    
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
