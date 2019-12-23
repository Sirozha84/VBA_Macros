Attribute VB_Name = "NewAndOldFind"
Const fNew = 1          '������� � "�����" �������
Const fOld = 2          '������� � "������" ��������
Const method = 2        '����� ���������
Const newTab = "���"    '"�����" �������
'Const newTab = "��� ����"    '"�����" �������
Const maxTabs = 4       '�������� �����

Global sn As Integer        '������� ����� �����
Global max As Integer       '������� ����� �����
Global oldTabs As Integer   '������� ������ ������

Sub NewAndOldFind()

    MakeCopy
    
    'AddNew("������")
    'AddNew("�����")
    AddNew ("Access")
    'AddNew ("Assecc ����")
    
    'FindDead("������")
    'FindDead("�����")
    'FindDead ("Assecc ����")
    FindDead ("Access")
    
    DeadSum
    
End Sub

'���������� �������� �������
Private Sub MakeCopy()
    
    Application.ScreenUpdating = True
    Application.StatusBar = "����������..."
    Application.ScreenUpdating = False
    
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

'����� � ���������� ����� �����
Private Sub AddNew(sheet)
    
    '������� �������� � ������ �������
    Application.ScreenUpdating = True
    Application.StatusBar = "������� �����..."
    Application.ScreenUpdating = False
    maxOld = 0
    Do Until Sheets(sheet).Cells(maxOld + 1, 1) = ""
        maxOld = maxOld + 1
    Loop
    
    For i = 1 To maxOld
    
        If i Mod 50 = 0 Then
            Application.ScreenUpdating = True
            Application.StatusBar = "����� �����:" + Persent(i, maxOld)
            Application.ScreenUpdating = False
        End If
    
        Find = False
        
        For j = 1 To max
            If Sheets("Res").Cells(j, 2) <> "" Then
                If Compare(Sheets("Res").Cells(j, fNew), Sheets(sheet).Cells(i, fOld)) Then
                    Find = True
                End If
            Else
                Exit For
            End If
        Next
        If Not Find Then
            For c = 1 To maxTabs
                Sheets("Res").Cells(max, c) = Sheets(sheet).Cells(i, c)
            Next
            Sheets("Res").Cells(max, maxTabs + 1) = "����� �� " + sheet
            max = max + 1
            sn = sn + 1
        End If
    Next
    
    Sheets("Res").Cells(max, maxTabs + 1) = "�����:" + Str(sn)
    
End Sub

'����� �������� �����
Private Sub FindDead(sheet)
    
    oldTabs = oldTabs + 1
    
    '������� �������� � ������ �������
    Application.ScreenUpdating = True
    Application.StatusBar = "������� �����..."
    Application.ScreenUpdating = False
    maxOld = 0
    Do Until Sheets(sheet).Cells(maxOld + 1, 1) = ""
        maxOld = maxOld + 1
    Loop
    
    
    For i = 1 To max - sn - 1
        
        If i Mod 50 = 0 Then
            Application.ScreenUpdating = True
            Application.StatusBar = "����� ��������:" + Persent(i, max - sn - 1)
            Application.ScreenUpdating = False
        End If
        
        Find = False
        For j = 1 To maxOld
            If Compare(Sheets("Res").Cells(i, fNew), Sheets(sheet).Cells(j, fOld)) Then
                Find = True
                Exit For
            End If
        Next
        If Not Find Then
            If Sheets("Res").Cells(i, maxTabs + 2) = "" Then
                Sheets("Res").Cells(i, maxTabs + 2) = 1
            Else
                Sheets("Res").Cells(i, maxTabs + 2) = Val(Sheets("Res").Cells(i, maxTabs + 2)) + 1
            End If
        End If
    Next
    
    
End Sub

'���� �� ��������
Private Sub DeadSum()
        
    Application.StatusBar = "����������..."
    
    s = 0
    For i = 1 To max
        If Sheets("Res").Cells(i, maxTabs + 2) = oldTabs Then
            Sheets("Res").Cells(i, maxTabs + 2) = "�����!"
            s = s + 1
        Else
            'Sheets("Res").Cells(i, maxTabs + 2) = ""
        End If
    Next
    Sheets("Res").Cells(max, maxTabs + 2) = "�������:" + Str(s)
    
    Application.StatusBar = "������!"
End Sub

'��������� ���������
Private Function Compare(newCell As String, oldCell As String) As Boolean
    
    Compare = False
    
    '����� ��������� 0 - ���� ���������
    If method = 0 Then
        Compare = (newCell = oldCell)
    End If
    
    '����� ��������� 1 - ��������� ��������� 5-� ��������
    If method = 1 Then
        Compare = (Left(newCell, 5) = Left(oldCell, 5))
    End If
    
    '����� ��������� 2 - ��������� �� ���� � ������ ���� "aaaaa ���: � ��" � "xxxx"
    If method = 2 Then
        s = Split(newCell, "���: ")
        If UBound(s) > 0 Then k = Replace(s(1), " ", "")
        Compare = (k = oldCell)
    End If
    
End Function

Private Function Persent(ByVal cur As Integer, ByVal all As Integer) As String
    Persent = Str(cur) + " ��" + Str(all) + " (" + Str(Int(cur / all * 100)) + "% )"
End Function


