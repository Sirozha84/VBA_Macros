Attribute VB_Name = "Starter"
'Last change: 19.04.2021 13:20

Type Adress
    UK As String    '����������� ��������
    ul As String
    dom As String
    korp As String
    kv As String
    ls As String    '����� �������� �����
    index As Long
    t1 As Long
    t2 As Long
End Type

Public Sub Start()
    FormReport.Show
End Sub

Public Sub SendEmails()
    FormEmails.Show
End Sub

Public Sub Instruction()
    MsgBox "��������� ������ �� ��������� � �������: ""������� ����"", ""�������� ����"" � �������� " + _
    """������"". �� ������ ����������� �������� �������� ��������� ����� �� ��������� �������." + Chr(13) + _
    "��������! ������� � ������������ ���������� ������ ���� ������������� �� ������� ""������������ ����������"", " + _
    "� ��� ���� ������ (��������� �����, ���������, �����, ����� ����) ������ ���� ���������.", _
    vbInformation, "����������"
    '�����, � ����������� ���������� � �������� ������ ������ � ��������...
End Sub

'******************** End of File ********************