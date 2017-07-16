Attribute VB_Name = "Module1"
Sub Поиск()
    i = 1
    Do While Cells(i, 1) <> ""
        cnt = 1
        Do
        s1 = Split(Cells(i, 1), ",")(0)
        If Cells(i + cnt, 1) <> "" Then
            s2 = Split(Cells(i + cnt, 1), ",")(0)
        Else
            s2 = ""
        End If
        fnd = (s1 = s2)
        If fnd Then
            cnt = cnt + 1
        Else
            Cells(i, 2) = cnt
            If cnt = 1 Then Cells(i, 1).Interior.Color = &H8080FF
            i = i + cnt - 1
        End If
        Loop Until Not fnd
        i = i + 1
    Loop
End Sub
