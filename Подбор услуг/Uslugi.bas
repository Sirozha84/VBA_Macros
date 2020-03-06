Attribute VB_Name = "Uslugi"
'������� ����� � ����������� �����


Const Col = 3
Const Usl = 4

Sub Uslugi()
    Message "������..."
    Max = 194311
    sh = 0
    last = 0
    Application.ScreenUpdating = False
    For i = 2 To Max
        sh = sh + 1
        If sh > 100 Then sh = 0: Call ProgressBar("���������", i, Max)
        
        If last = Cells(i, Col) Then
            Cells(i - 1, Usl) = Cells(i, Usl)
            Rows(i).EntireRow.Delete
            i = i - 1
            Max = Max - 1
        Else
            last = Cells(i, Col)
        End If
        
        If Cells(i, Usl) = "���" Then Cells(i, 5) = "+"
        If Cells(i, Usl) = "��� ��" Then Cells(i, 6) = "+"
        If Cells(i, Usl) = "��" Then Cells(i, 7) = "+"
        If Cells(i, Usl) = "���������" Then Cells(i, 8) = "+"

        
        If Cells(i, 1) = "" Then Exit For
    Next
    Message "������!"
End Sub

'��������� ���������, text - ���, cur - ������� ��������, all - �����, ���������� ������ over ����
Private Sub ProgressBar(text As String, ByVal cur As Long, ByVal all As Long)
    Application.ScreenUpdating = True
    Application.StatusBar = text + ":" + Str(cur) + " ��" + Str(all) + _
        " (" + Str(Int(cur / all * 100)) + "% )"
        DoEvents
    Application.ScreenUpdating = False
End Sub


Private Function Adress(i As Long) As String
    Adress = CStr(Cells(i, 1)) + CStr(Cells(i, 2)) + CStr(Cells(i, 3)) + CStr(Cells(i, 4))
    Adress = LCase(Adress)
    Adress = Replace(Adress, "�", "�")
End Function

Private Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    DoEvents
    Application.ScreenUpdating = False
End Sub





