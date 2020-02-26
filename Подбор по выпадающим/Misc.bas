'������ 1.0 (13.01.2020) - ����� �������������� �������� � ��������� ����
'������ 1.1 (26.02.2020) - ��� ����� ���������: OneToTwo � Vipad

Attribute VB_Name = "Misc"
'������ ������� �� ������� ������
'������ ��������� �� ����� �������, ���������� �� ������ ����� ��, ������ ����������� � ����� � ������
Sub AdresesByLS()
    Const ft = 1    '������ � �� � ������ �������
    Const st = 25   '������ � �� �� ������ �������
    Const atb = 6   '���������� �������� � �������
    i = 2
    Do While Cells(i, ft) <> ""
        j = 2
        Do While Cells(j, st) <> ""
            If Cells(i, ft) = Cells(j, st) Then
                For k = 1 To atb
                    Cells(i, ft + k) = Cells(j, st + k)
                Next
                Exit Do
            End If
            j = j + 1
        Loop
        i = i + 1
    Loop
    
End Sub

'������������ �����
Sub Dedupl()
    Application.ScreenUpdating = False
    j = 1
    last = ""
    For i = 2 To 42283
        Cells(j, 5) = Cells(i, 1)
        Cells(j, 6) = Cells(i, 2)
        If Cells(i, 1) <> last Then
            last = Cells(i, 1)
            
            j = j + 1
        End If
    Next
End Sub

'������ �������������
Sub Koef()
    col = 25
    Max = 153519
    Sheets("Result").Select
    For i = 1 To Max
        If i Mod 1000 = 0 Then Call ProgressBar("������ �������������", i, Max)
        adr = Cells(i, 2) + CStr(Cells(i, 3)) + Cells(i, 4)
        For j = 1 To 1148
            If adr = Sheets("Adresses").Cells(j, 15) Then
                Cells(i, col) = Sheets("Adresses").Cells(j, 16)
                Exit For
            End If
        Next
    Next
    Message "������!"
End Sub

'������ ������� ������� �� "������"
Sub OneToTwo()
    Max = 39689
    For i = 4 To Max
        If i Mod 1000 = 0 Then Call ProgressBar("���������", i, Max)
        Cells(i, 1) = "2" + Right(Cells(i, 1), Len(Cells(i, 1)) - 1)
    Next
    Message "������!"
End Sub

'����������� ����������� � ������ ������� (���������� �� ����� �������� �����)
Sub Vipad()
    Max = 39163
    last = 4
    For i = 2 To Max
        If i Mod 100 = 0 Then Call ProgressBar("���������", i, Max)
        For j = last To 39689
            Find = False
            If Sheets("���������").Cells(i, 7) = Sheets("Vip").Cells(j, 1) Then
                Sheets("���������").Cells(i, 17) = Sheets("Vip").Cells(j, 10)
                last = j
                Find = True
                Exit For
            End If
            If Not Find Then last = 4
        Next
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

Private Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    DoEvents
    Application.ScreenUpdating = False
End Sub
