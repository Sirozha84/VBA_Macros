Attribute VB_Name = "NewMonth"
'����� �����
'�����: ������ �������
'���� ���������: 20.01.2017

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
    
    
    If MsgBox("������� ����� ������?" + Chr(13) + "��������!!! ��� ���� ��� ������������ ������ ��������, " + _
        "� ��������� 5 ���� ��������� �����.", vbYesNo) = 6 Then
        '������ ����� � ������� ����� �����
    End If
    Exit Sub
error1:
    MsgBox ("������. ��������� ������� ��������.")
End Sub
