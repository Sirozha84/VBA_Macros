Attribute VB_Name = "Trial"
'�������� �� ����� ������������� ���������
'�����: ������ �������
'���� ���������: 20.01.2017

Function WorkIt() As Boolean
    Days = DateDiff("d", Now, "1.4.2017")
    If (Days <= 0) Then
        MsgBox ("������ ���������� ������������� ��������� ����." + _
            Chr(13) + "���������� ���������� ��������.")
        WorkIt = False
        Exit Function
    End If
    If (Days < 30) Then MsgBox ("��������!" + Chr(13) + "������ ���������� ������������� ��������� ��������." + _
        Chr(13) + "�������� ����: " + Str(Days))
    WorkIt = True
End Function
