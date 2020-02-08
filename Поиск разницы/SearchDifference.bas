Attribute VB_Name = "SearchDifference"
'������ 1.0 (19.12.2019)
'������ 1.2 (20.12.2019)
'������ 1.3 (23.12.2019) - ��������� ���� �� ������
'������ 1.4 (24.12.2019) - ����������� ���������
'������ 1.5 (24.12.2019) - ����� ���������
'������ 1.6 (26.12.2019) - �����������
'������ 1.7 (10.01.2020) - ���������� �������
'������ 1.8 (17.01.2020) - ������ �������� �������
'������ 1.9 (27.01.2020) - ���������� ������ � �������� �������������� ������� ����������
'������ 1.10 (04.02.2020) - ��������� ������� ���������
'������ 1.11 (07.02.2020) - �������� ������ ��������

Const firstTab = "Access"   '������ �������
Const secondTab = "���"     '������ �������
Const resultTab = "Result"  '������� � �����������
Const maxTabs = 3           '�������� �����
Const machTabs = 1          '���������� ����� ��� ������ �����
Const compareTab = 3        '������� ��� ��������� ��������
Const calculateDif = True   '������� ������� �������� (��� ����� true, ��� ����� false)
Const fCom = 8              '���� ��� �����������

Global sn As Integer        '������� ����� �����
Global max As Integer       '������� ����� �����


Sub Start()

    PrepareTabFromAccess (firstTab)
    PrepareTabFromAccess (secondTab)

    CreateSheet (resultTab)

    MakeCopy
    AddDead
    FindNew
    
    Message "������!"
    
End Sub

'�������� �������������� �������
Private Sub CreateSheet(name As String)
    If Not SheetExist(name) Then
        Sheets.Add(, Sheets(Sheets.Count)).name = name
    End If
    Sheets(name).Cells.Clear
End Sub

'�������� �� ������������� �����
Function SheetExist(name As String) As Boolean
    Dim objSheet As Object
    
    On Error GoTo HandleError
    ThisWorkbook.Worksheets(name).Activate
    SheetExist = True
    Exit Function
    
HandleError:
    SheetExist = False
End Function

'���������� ������, ����������� �� Access
Sub PrepareTabFromAccess(shit As String)
    
    Sheets(shit).Select
    
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
        
        '�������� ������ ��������
        Do While Right(Cells(i, 1), 1) = " "
            Cells(i, 1) = Left(Cells(i, 1), Len(Cells(i, 1)) - 1)
        Loop
        
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

'���������� �������� �������
Private Sub MakeCopy()
    
    Message "����������..."
    
    sn = 0
    mx = 0
    
    Sheets(resultTab).Cells.Clear
    maxStrings = 0
    max = 1
    Do While Sheets(secondTab).Cells(max + 1, 1) <> ""
        max = max + 1
        For i = 1 To 3
            Sheets(resultTab).Cells(max, i) = Sheets(secondTab).Cells(max, i)
        Next
    Loop
    
    '������������ �����
    Sheets(resultTab).Cells(1, 1) = "�������"
    Sheets(resultTab).Cells(1, 3) = secondTab
    Sheets(resultTab).Cells(1, 5) = firstTab
    Sheets(resultTab).Cells(1, 6) = "�������"
    
End Sub

'����� � ���������� �������� �����
Private Sub AddDead()
    
    '������� �������� � ������ �������
    Message "������� �����..."
    maxOld = 0
    Do Until Sheets(firstTab).Cells(maxOld + 1, 1) = ""
        maxOld = maxOld + 1
    Loop
    
    mached = 0
    changed = 0
    
    Find = False
    For i = 2 To maxOld
        Call ProgressBar("����� ��������", i, maxOld)
        For j = 2 To max
            Find = False
            If Sheets(resultTab).Cells(j, 1) = Sheets(firstTab).Cells(i, 1) Then
                Find = True
                Exit For
            End If
        Next
        If Find Then
            '���� �����, ����� ����������
            Sheets(resultTab).Cells(j, 4) = Sheets(firstTab).Cells(i, 2)
            Sheets(resultTab).Cells(j, 5) = Sheets(firstTab).Cells(i, 3)
            rz = Round(Sheets(resultTab).Cells(j, 3) - Sheets(resultTab).Cells(j, 5), 2)
            Sheets(resultTab).Cells(j, 6) = rz
            If Abs(rz) = 0 Then
                Sheets(resultTab).Cells(j, 3).Interior.Color = RGB(128, 255, 128)
                Sheets(resultTab).Cells(j, 5).Interior.Color = RGB(128, 255, 128)
                Sheets(resultTab).Cells(j, 6).Interior.Color = RGB(196, 255, 196)
                Sheets(resultTab).Cells(j, fCom) = "������"
            End If
            If Abs(rz) > 0 Then
                Sheets(resultTab).Cells(j, 3).Interior.Color = RGB(255, 255, 128)
                Sheets(resultTab).Cells(j, 5).Interior.Color = RGB(255, 255, 128)
                Sheets(resultTab).Cells(j, 6).Interior.Color = RGB(255, 255, 196)
                Sheets(resultTab).Cells(j, fCom) = "������ (�����)"
            End If
            If Abs(rz) > 10 Then
                Sheets(resultTab).Cells(j, 3).Interior.Color = RGB(255, 128, 128)
                Sheets(resultTab).Cells(j, 5).Interior.Color = RGB(255, 128, 128)
                Sheets(resultTab).Cells(j, 6).Interior.Color = RGB(255, 196, 196)
                changed = changed + 1
                Sheets(resultTab).Cells(j, fCom) = "������"
            End If
            If Abs(rz) < 10 Then mached = mached + 1
            sumDif = sumDif + rz
        Else
            '� ���� �� �����, ��������
            max = max + 1
            Sheets(resultTab).Cells(max, 1) = Sheets(firstTab).Cells(i, 1)
            Sheets(resultTab).Cells(max, 4) = Sheets(firstTab).Cells(i, 2)
            Sheets(resultTab).Cells(max, 5) = Sheets(firstTab).Cells(i, 3)
            Sheets(resultTab).Cells(max, fCom) = "���� � " + firstTab + ", �� ��� � " + secondTab
            
            '��������� �������
            'Sheets(resultTab).Cells(max, 3).Cells = 0
            Sheets(resultTab).Cells(max, 6).Cells = -Sheets(resultTab).Cells(max, 5).Cells
            Sheets(resultTab).Cells(max, 3).Interior.Color = RGB(255, 128, 128)
            Sheets(resultTab).Cells(max, 5).Interior.Color = RGB(255, 128, 128)
            Sheets(resultTab).Cells(max, 6).Interior.Color = RGB(255, 196, 196)
            
            sumDif = sumDif - Sheets(resultTab).Cells(max, 3).Cells
            
            sn = sn + 1
        End If
    Next
    
    Sheets(resultTab).Cells(max + 3, maxTabs + 5) = "���� ������ � " + firstTab + ":" + Str(sn)
    Sheets(resultTab).Cells(max + 4, maxTabs + 5) = "�������:" + Str(mached)
    Sheets(resultTab).Cells(max + 5, maxTabs + 5) = "��������:" + Str(changed)
    
End Sub

'����� ����� �����
Private Sub FindNew()
    
    '������� �������� � ������ �������
    Message "������� �����..."
    maxOld = 0
    Do Until Sheets(firstTab).Cells(maxOld + 1, 1) = ""
        maxOld = maxOld + 1
    Loop
    
    s = 0
    For i = 2 To max - sn
        
        Call ProgressBar("����� �����", i, max - sn - 1)
        
        Find = False
        For j = 2 To maxOld
            If Sheets(resultTab).Cells(i, 1) = Sheets(firstTab).Cells(j, 1) Then
                Find = True
                Exit For
            End If
        Next
        If Not Find Then
            
            '��������� "�������"
            Sheets(resultTab).Cells(i, 6).Cells = Sheets(resultTab).Cells(i, maxTabs).Cells
            Sheets(resultTab).Cells(i, 3).Interior.Color = RGB(255, 128, 128)
            Sheets(resultTab).Cells(i, 5).Interior.Color = RGB(255, 128, 128)
            Sheets(resultTab).Cells(i, 6).Interior.Color = RGB(255, 196, 196)
            
            Sheets(resultTab).Cells(i, fCom) = "���� � " + secondTab + ", �� ��� � " + firstTab
            sumDif = sumDif + Sheets(resultTab).Cells(i, maxTabs).Cells
            
            s = s + 1
            
        End If
    Next
    
    Sheets(resultTab).Cells(max + 2, maxTabs + 5) = "���� ������ � " + secondTab + ":" + Str(s)
    
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


