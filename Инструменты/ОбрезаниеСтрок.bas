Attribute VB_Name = "ОбрезаниеСтрок"
'Версия 1.0 (02.09.2020)

Const Col = 1		'Номер колонки
Const Rows = 10000	'Количество строк
Const spl = " от "	'Строка, после которой всё обрезается начиная с первого символа этой строки

Sub cutt()
    For i = 1 To Rows
        s = Cells(i, Col)
        If s <> "" Then
            m = Split(s, spl)
            Cells(i, Col) = m(0)
        End If
    Next
End Sub

