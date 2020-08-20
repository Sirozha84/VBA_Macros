Attribute VB_Name = "CalcCompare"
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
'������ 2.0 (21.07.2020) - ��������� ��� ����� �����, ������ ������ ������ ���� ������
'������ 2.1 (22.07.2020) - ��������� ��������� "���" � ������ ������� � ������ ���� ������� (�������� ��������� �������)
'������ 2.2 (23.07.2020) - ��� ���� �������� ��� ��������� "��"
'������ 2.3 (20.08.2020) - �������� ����� �������� ��� ���������

Const Luft = 5              '����, ������������ ����� �������� ����� �������� ��� "����� �������"
Const fCom = 7              '���� ��� �����������
Global Tab1C As String      '������ �������
Global TabAccess As String  '������ �������
Global TabResult As String  '������� � �����������
Global max As Integer       '������� ����� �����


Sub Start()
    
    Tab1C = "TDSheet"
    
    TabAccess = "�����"
    TabResult = "����� R"
    MakeCopy
    Compare
    
    TabAccess = "����"
    TabResult = "���� R"
    MakeCopy
    Compare
    
    TabAccess = "��"
    TabResult = "�� R"
    MakeCopy
    Compare
    
    Message "������!"
    
End Sub

'���������� �������� �������
Private Sub MakeCopy()
    
    Message TabAccess + ": �����������..."
    
    If Not SheetExist(TabResult) Then
        Sheets.Add(Sheets(Sheets.Count)).name = TabResult
    End If
    
    Sheets(TabResult).Select
    Cells.Clear
    maxStrings = 0
    
    '�������� ������ �� Access
    max = 1
    Do While Sheets(TabAccess).Cells(max + 1, 1) <> ""
        max = max + 1
        For i = 1 To 3
            Sheets(TabResult).Cells(max, i) = Sheets(TabAccess).Cells(max, i)
        Next
    Loop
    
    '������������ �����
    Cells(1, 1) = "�������"
    Cells(1, 2) = "Access"
    Cells(1, 3) = "�����"
    Cells(1, 4) = "���� (1�)"
    Cells(1, 5) = "�����"
    Cells(1, 6) = "�������"
    Cells(1, 7) = "����"
    Columns(2).ColumnWidth = 10
    Columns(2).ColumnWidth = 30
    Columns(4).ColumnWidth = 30
    Columns(7).ColumnWidth = 18
    
End Sub

'���������
Private Sub Compare()
    
    Message TabAccess + ": ��������� � ����..."
    Sheets(TabResult).Select
    
    i = 2
    Dim Mached As Long
    Dim Pochti As Long
    Dim Changed As Long
    
    Do While Sheets(Tab1C).Cells(i, 1) <> ""
        
        Find = False
        For j = 2 To max
            Find = False
            If Trim(Cells(j, 1)) = Trim(Sheets(Tab1C).Cells(i, 1)) Then
                Find = True
                Exit For
            End If
        Next
        If Find Then
            '���� �����, ����� ����������
            Cells(j, 4) = Sheets(Tab1C).Cells(i, 2)
            Cells(j, 5) = Sheets(Tab1C).Cells(i, 3)
            rz = Round(Sheets(TabResult).Cells(j, 5) - Cells(j, 3), 2)
            Cells(j, 6) = rz
            If Abs(rz) = 0 Then
                Cells(j, 3).Interior.Color = RGB(196, 255, 196)
                Cells(j, 5).Interior.Color = RGB(196, 255, 196)
                Cells(j, 6).Interior.Color = RGB(128, 255, 128)
                Cells(j, fCom) = "������"
                Mached = Mached + 1
            End If
            If Abs(rz) > 0 And Abs(rz) <= Luft Then
                Cells(j, 3).Interior.Color = RGB(255, 255, 196)
                Cells(j, 5).Interior.Color = RGB(255, 255, 196)
                Cells(j, 6).Interior.Color = RGB(255, 255, 128)
                Cells(j, fCom) = "�����"
                Pochti = Pochti + 1
            End If
            If Abs(rz) > Luft Then
                Cells(j, 3).Interior.Color = RGB(255, 196, 196)
                Cells(j, 5).Interior.Color = RGB(255, 196, 196)
                Cells(j, 6).Interior.Color = RGB(255, 128, 128)
                Changed = Changed + 1
                Cells(j, fCom) = "�� ������"
            End If
            'sumDif = sumDif + rz
        End If
        i = i + 1
    Loop
    
    Message TabAccess + ": ����������..."
    
    '����� �� ��������
    Dim NonFind As Long
    For i = 2 To max
        If Cells(i, 6) = "" Then
            Cells(i, fCom) = "�� �������"
            NonFind = NonFind + 1
        End If
    Next
    
    '�����
    Dim all As Long
    all = Mached + Pochti + Changed + NonFind
    Cells(max + 2, fCom) = "�����: " + CStr(all)
    Cells(max + 3, fCom) = "�������: " + CStr(Mached) + " + " + CStr(Pochti) + Prc(Mached + Pochti, all)
    Cells(max + 4, fCom) = "�� �������: " + CStr(Changed) + Prc(Changed, all)
    Cells(max + 5, fCom) = "�� �������: " + CStr(NonFind) + Prc(NonFind, all)
    
End Sub

'������� ���������
Private Function Prc(val As Long, all As Long) As String
    Prc = " (" + CStr(Round(val / all * 100, 1)) + "%)"
End Function

'����� ��������� � ���������
Private Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    Application.ScreenUpdating = False
End Sub

'�������� �� ������������� �����
Private Function SheetExist(name As String) As Boolean
    Dim objSheet As Object
    
    On Error GoTo HandleError
    ThisWorkbook.Worksheets(name).Activate
    SheetExist = True
    Exit Function
    
HandleError:
    SheetExist = False
End Function
