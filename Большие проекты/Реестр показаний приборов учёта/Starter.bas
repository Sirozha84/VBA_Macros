Attribute VB_Name = "Starter"
'Last change: 01.06.2021 09:17

Type Adress
    UK As String    '����������� ��������
    ul As String
    dom As String
    korp As String
    kv As String
    potr As String  '�����������
    prop As Long    '���������
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
    MsgBox "��������� ������ �� ��������� � �������: ""������� ����"", ""�������� ����"" � �������� ""������""." + Chr(13) + _
    "�� ������ ����������� �������� �������� ��������� ����� �� ��������� �������.", vbInformation
End Sub

'******************** End of File ********************