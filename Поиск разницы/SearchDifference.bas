Attribute VB_Name = "SearchDifference"
'������ 1.0 (19.12.2019)
'������ 1.2 (20.12.2019)
'������ 1.3 (23.12.2019) - ��������� ���� �� ������
'������ 1.4 (24.12.2019) - ����������� ���������
'������ 1.5 (24.12.2019) - ����� ���������
'������ 1.6 (26.12.2019) - �����������
'������ 1.7 (10.01.2020) - ���������� �������
'������ 1.8 (17.01.2020) - ������ �������� �������


Const newTab = "���"        '����� �������
Const oldTab = "Access"     '������ ������� �������
Const maxTabs = 3           '�������� �����
Const fCom = 8

Global sn As Integer        '������� ����� �����
Global max As Integer       '������� ����� �����
Global oldTabs As Integer   '������� ������ ������

Global subNew As Long
Global sumOld As Long
Global sumDif As Long

Sub Start()

    MakeCopy
    AddDead
    FindNew
    
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
    max = 1
    Do While Sheets(newTab).Cells(max + 1, 1) <> ""
        max = max + 1
        For i = 1 To 3
            Sheets("Res").Cells(max, i) = Sheets(newTab).Cells(max, i)
        Next
    Loop
    
    '������������ �����
    Sheets("Res").Cells(1, 1) = "�������"
    Sheets("Res").Cells(1, 3) = newTab
    Sheets("Res").Cells(1, 5) = oldTab
    Sheets("Res").Cells(1, 6) = "�������"
    
End Sub

'����� � ���������� �������� �����
Private Sub AddDead()
    
    '������� �������� � ������ �������
    Message "������� �����..."
    maxOld = 0
    Do Until Sheets(oldTab).Cells(maxOld + 1, 1) = ""
        maxOld = maxOld + 1
    Loop
    
    mached = 0
    changed = 0
    
    Find = False
    For i = 2 To maxOld
        Call ProgressBar("����� ��������", i, maxOld)
        For j = 2 To max
            Find = False
            If Sheets("Res").Cells(j, 1) = Sheets(oldTab).Cells(i, 1) Then
                Find = True
                Exit For
            End If
        Next
        If Find Then
            '���� �����, ����� ����������
            Sheets("Res").Cells(j, 4) = Sheets(oldTab).Cells(i, 2)
            Sheets("Res").Cells(j, 5) = Sheets(oldTab).Cells(i, 3)
            rz = Round(Sheets("Res").Cells(j, 3) - Sheets("Res").Cells(j, 5), 2)
            Sheets("Res").Cells(j, 6) = rz
            If Abs(rz) = 0 Then
                Sheets("Res").Cells(j, 3).Interior.Color = RGB(128, 255, 128)
                Sheets("Res").Cells(j, 5).Interior.Color = RGB(128, 255, 128)
                Sheets("Res").Cells(j, 6).Interior.Color = RGB(196, 255, 196)
                mached = mached + 1
                Sheets("Res").Cells(j, fCom) = "������"
            End If
            If Abs(rz) > 0 Then
                Sheets("Res").Cells(j, 3).Interior.Color = RGB(255, 255, 128)
                Sheets("Res").Cells(j, 5).Interior.Color = RGB(255, 255, 128)
                Sheets("Res").Cells(j, 6).Interior.Color = RGB(255, 255, 196)
                mached = mached + 1
                Sheets("Res").Cells(j, fCom) = "������ (�����)"
            End If
            If Abs(rz) > 10 Then
                Sheets("Res").Cells(j, 3).Interior.Color = RGB(255, 128, 128)
                Sheets("Res").Cells(j, 5).Interior.Color = RGB(255, 128, 128)
                Sheets("Res").Cells(j, 6).Interior.Color = RGB(255, 196, 196)
                changed = changed + 1
                Sheets("Res").Cells(j, fCom) = "������"
            End If
            sumDif = sumDif + rz
        Else
            '� ���� �� �����, ��������
            max = max + 1
            Sheets("Res").Cells(max, 1) = Sheets(oldTab).Cells(i, 1)
            Sheets("Res").Cells(max, 4) = Sheets(oldTab).Cells(i, 2)
            Sheets("Res").Cells(max, 5) = Sheets(oldTab).Cells(i, 3)
            Sheets("Res").Cells(max, fCom) = "���� � " + oldTab + ", �� ��� � " + newTab
            
            '��������� �������
            'Sheets("Res").Cells(max, 3).Cells = 0
            Sheets("Res").Cells(max, 6).Cells = -Sheets("Res").Cells(max, 5).Cells
            Sheets("Res").Cells(max, 3).Interior.Color = RGB(255, 128, 128)
            Sheets("Res").Cells(max, 5).Interior.Color = RGB(255, 128, 128)
            Sheets("Res").Cells(max, 6).Interior.Color = RGB(255, 196, 196)
            
            sumDif = sumDif - Sheets("Res").Cells(max, 3).Cells
            
            sn = sn + 1
        End If
    Next
    
    Sheets("Res").Cells(max + 3, maxTabs + 5) = "���� ������ � " + oldTab + ":" + Str(sn)
    Sheets("Res").Cells(max + 4, maxTabs + 5) = "�������:" + Str(mached)
    Sheets("Res").Cells(max + 5, maxTabs + 5) = "��������:" + Str(changed)
    
End Sub

'����� ����� �����
Private Sub FindNew()
    
    '������� �������� � ������ �������
    Message "������� �����..."
    maxOld = 0
    Do Until Sheets(oldTab).Cells(maxOld + 1, 1) = ""
        maxOld = maxOld + 1
    Loop
    
    s = 0
    For i = 2 To max - sn
        
        Call ProgressBar("����� �����", i, max - sn - 1)
        
        Find = False
        For j = 2 To maxOld
            If Sheets("Res").Cells(i, 1) = Sheets(oldTab).Cells(j, 1) Then
                Find = True
                Exit For
            End If
        Next
        If Not Find Then
            
            '��������� "�������"
            Sheets("Res").Cells(i, 6).Cells = Sheets("Res").Cells(i, maxTabs).Cells
            Sheets("Res").Cells(i, 3).Interior.Color = RGB(255, 128, 128)
            Sheets("Res").Cells(i, 5).Interior.Color = RGB(255, 128, 128)
            Sheets("Res").Cells(i, 6).Interior.Color = RGB(255, 196, 196)
            
            Sheets("Res").Cells(i, fCom) = "���� � " + newTab + ", �� ��� � " + oldTab
            sumDif = sumDif + Sheets("Res").Cells(i, maxTabs).Cells
            
            s = s + 1
            
        End If
    Next
    
    Sheets("Res").Cells(max + 2, 3) = sumNew
    Sheets("Res").Cells(max + 2, 4) = sumOld
    Sheets("Res").Cells(max + 2, 5) = sumDif
    
    Sheets("Res").Cells(max + 2, maxTabs + 5) = "���� ������ � " + newTab + ":" + Str(s)
    
End Sub

'����� ������� ���������
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


