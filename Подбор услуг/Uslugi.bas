Attribute VB_Name = "Uslugi"
'������� ����� � ����������� �����

Const Col = 3
Const Usl = 4

Sub Uslugi1()
    Message "������..."
    Max = 194311
    sh = 0
    last = 0
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

Sub Uslugi2()

    ALlMoney = 12 '������� �� ����� ��������
    Money = 13    '������ ������� � ��������
    Volumes = 17  '������ ������� � ��������
    
    Message "������..."
    Max = 3938
    sh = 0
    last = 0
    For i = 2 To Max
        sh = sh + 1
        If sh > 100 Then sh = 0: Call ProgressBar("���������", i, Max)
        
        '����������� �������������
        
        '��-��
        If Cells(i, Volumes) <> "" Then Cells(i, Money) = Cells(i, ALlMoney)
        '���������
        If Cells(i, Volumes + 2) <> "" Then Cells(i, Money + 2) = Cells(i, ALlMoney)
        '���
        If Cells(i, Volumes + 3) <> "" Then Cells(i, Money + 3) = Cells(i, ALlMoney)
        
        If last = Cells(i, 1) Then
            
            '�����������
            
            '��-��
            If Cells(i, Volumes) <> "" Then
                If Cells(i - 1, Volumes) = "" Then
                    '������ �����, ������ ��� ������
                    Cells(i - 1, Money) = Cells(i, ALlMoney)
                    Cells(i - 1, Volumes) = Cells(i, Volumes)
                Else
                    '������ �� �����...
                    If Cells(i - 1, Volumes) > Cells(i, Volumes) Then
                        Cells(i - 1, Money + 1) = Cells(i, ALlMoney)
                        Cells(i - 1, Volumes + 1) = Cells(i, Volumes)
                    Else
                    
                        Cells(i - 1, Money) = Cells(i, ALlMoney)
                        Cells(i - 1, Volumes + 1) = Cells(i - 1, Volumes)
                        Cells(i - 1, Volumes) = Cells(i, Volumes)
                    End If
                End If
            End If
            '���������
            If Cells(i, Volumes + 2) <> "" Then
                Cells(i - 1, Money + 2) = Cells(i, ALlMoney)
                Cells(i - 1, Volumes + 2) = Cells(i, Volumes + 2)
            End If
            '���
            If Cells(i, Volumes + 3) <> "" Then
                Cells(i - 1, Money + 3) = Cells(i, ALlMoney)
                Cells(i - 1, Volumes + 3) = Cells(i, Volumes + 3)
            End If
            
            Rows(i).EntireRow.Delete
            i = i - 1
            Max = Max - 1
        Else
            last = Cells(i, 1)
        End If
        
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





