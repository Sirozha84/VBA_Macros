Attribute VB_Name = "Calculate"
'��������� �����
Const NP = "������"

Const cNP = 1       '��������� �����
Const cReg = 9      '������������������
Const cOwner = 10   '������������

Dim Humans As Long

Sub Calculate()
    Humans = 0
    i = 2
    Do While Cells(i, cNP) <> ""
        If Cells(i, cNP) = NP Then
            If Cells(i, cReg) > 0 Then
                Humans = Humans + Cells(i, cReg)
            Else
                Humans = Humans + Cells(i, cOwner)
            End If
        End If
        i = i + 1
    Loop
    e = Chr(10)
    MsgBox "������ ��������!" + e + _
            e + "��������� = " + CStr(Humans)
End Sub
