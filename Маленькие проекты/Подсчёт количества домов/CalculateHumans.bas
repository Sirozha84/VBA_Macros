Attribute VB_Name = "Calculate"
'Населённый пункт
Const NP = "Тартат"

Const cNP = 1       'Населённый пункт
Const cReg = 9      'Зарегистрированные
Const cOwner = 10   'Собственники

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
    MsgBox "Расчёт закончен!" + e + _
            e + "Население = " + CStr(Humans)
End Sub
