Attribute VB_Name = "��������������"
'������ 1.0 (02.09.2020)

Const Col = 1		'����� �������
Const Rows = 10000	'���������� �����
Const spl = " �� "	'������, ����� ������� �� ���������� ������� � ������� ������� ���� ������

Sub cutt()
    For i = 1 To Rows
        s = Cells(i, Col)
        If s <> "" Then
            m = Split(s, spl)
            Cells(i, Col) = m(0)
        End If
    Next
End Sub

