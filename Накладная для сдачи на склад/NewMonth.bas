Attribute VB_Name = "NewMonth"
'����� �����
'�����: ������ �������
'���� ���������: 31.01.2017

Sub NewMonth()
    '� ������ ������� ������
    If Not WorkIt Then Exit Sub
    
    Dim data As String
    data = InputBox("������� ���� (� ������� �����.���, �������� ������ 2017 ����� ��������� ��� 1.17)." + Chr(13))
    If data = "" Then Exit Sub
    On Error GoTo error1
    dt = Split(data, ".")
    mnth = dt(0)
    yer = dt(1)
    '������ �������� ���
    
    On Error GoTo 0
    If MsgBox("������� ����� ������?" + Chr(13) + "��������!!! ��� ���� ��� ������������ ������ ��������, " + _
        "� ��������� 5 ���� ��������� �����.", vbYesNo) = 6 Then
        Application.ScreenUpdating = False
        '������ �����
        For d = 27 To 31
            CopyPage Trim(Str(d)) + "�", "-" + Trim(Str(d)) + "�"
            CopyPage Trim(Str(d)) + "�", "-" + Trim(Str(d)) + "�"
        Next
        '������� ����� �����
        For d = 1 To 31
            CLearPage Trim(Str(d)) + "�"
            CLearPage Trim(Str(d)) + "�"
        Next
        Application.ScreenUpdating = True
    End If
    '��������� ����
    FillDates mnth, yer
    
    Exit Sub
error1:
    MsgBox ("������. ��������� ������� ��������.")
End Sub

Sub FillDates(m, y)
    nn = 1
    'Dim dd As Date '���������� ������� ���, �� �� ���� ��� :( VBA - ����� ������
    For d = 1 To 31
        'dd = System.DateTime(1993, 5, 31, 12, 14, 0)
        'dd = 3 / 17 / 1984
        Sheets(Trim(Str(d)) + "�").Cells(1, 6) = "��������� �" + Str(nn)
        Sheets(Trim(Str(d)) + "�").Cells(2, 6) = Trim(Str(d)) + "." + m + "." + y
        'Sheets(Trim(Str(d)) + "�").Cells(2, 6) = dd
        nn = nn + 1
        Sheets(Trim(Str(d)) + "�").Cells(1, 6) = "��������� �" + Str(nn)
        Sheets(Trim(Str(d)) + "�").Cells(2, 6) = Trim(Str(d)) + "." + m + "." + y
        'Sheets(Trim(Str(d)) + "�").Cells(2, 6) = dd
        nn = nn + 1
    Next
End Sub

'����������� ��������
Sub CopyPage(sorce, dist)
    'MsgBox """" + sorce + """ - """ + dist + """"
    For s = 6 To 25
        For C = 2 To 17
            Sheets(dist).Cells(s, C) = Sheets(sorce).Cells(s, C)
        Next
    Next
    Sheets(dist).Cells(1, 6) = Sheets(sorce).Cells(1, 6)
    Sheets(dist).Cells(2, 6) = Sheets(sorce).Cells(2, 6)
End Sub

'������� ��������
Sub CLearPage(page)
    For s = 6 To 25
        For C = 2 To 17
            Sheets(page).Cells(s, C) = ""
        Next
    Next
End Sub
