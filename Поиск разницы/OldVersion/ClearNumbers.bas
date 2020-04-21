Attribute VB_Name = "ClearNumbers"
'������ ������� ��������� �� ��������

Sub Start()
    
    Message "������� �����..."
    max = 1
    
    Do Until Cells(max + 1, 1) = ""
        max = max + 1
    Loop
    
    For i = 2 To max
        
        Call ProgressBar("���������", i, max)
        
        '������� � �����
        Cells(i, 1).NumberFormat = "@"
        
        '�������� �� ������ �������� ������� ������
        Cells(i, 1) = Replace(Cells(i, 1), "��������������� �������� �� � ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "�������  �� � ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "������� � ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "������� �� �  ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "������� �� � ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "������� �� � ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "������� �� �", "")
        Cells(i, 1) = Replace(Cells(i, 1), "�������� �� �  ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "�������� �� � ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "������������� �������� �� � ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "������������� ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "������������� ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "������������ ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "������������ ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "�������� �� � ", "")

    Next
    
    Message "������!"
    
End Sub


'����� ������� ���������
Private Sub ProgressBar(text As String, ByVal cur As Integer, ByVal all As Integer)
    If cur Mod 100 = 0 Then
        Application.ScreenUpdating = True
        Application.StatusBar = text + ":" + _
            Str(cur) + " ��" + Str(all) + " (" + Str(Int(cur / all * 100)) + "% )"
        Application.ScreenUpdating = False
    End If
End Sub

'����� ��������� � ���������
Private Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    Application.ScreenUpdating = False
End Sub


