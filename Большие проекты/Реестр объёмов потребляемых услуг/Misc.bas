Attribute VB_Name = "Misc"
'Last change: 06.04.2021 09:36

Dim startTime As Date
Private SearchMethod As Byte

'��������� � ������ �������
Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    DoEvents
    Application.ScreenUpdating = False
End Sub

'�������� �������������� �������
Sub NewTab(name As String, create As Boolean)
    name = Left(name, 31)
    If create Then
        If Not SheetExist(name) Then
            Sheets.Add(, Sheets(Sheets.Count)).name = name
        End If
    End If
    If Not SheetExist(name) Then
        LabelStatus = "������: �������������� ������� �� ����������"
    End If
    Sheets(name).Cells.Clear
End Sub

'�������� �� ������������� �����
Private Function SheetExist(name As String) As Boolean
    Dim objSheet As Object
    On Error GoTo HandleError
    ThisWorkbook.Worksheets(name).Activate
    SheetExist = True
    Exit Function
HandleError:
    SheetExist = False
End Function

'����� ������������� ���������� ����� (��������� ������ ��� ������ ������� ���������� ���������)
Function FindMax(tb As Variant) As Long
    i = 0
    Do While tb.Cells(i + 1000, 2) <> ""
        i = i + 1000
    Loop
    Do While tb.Cells(i + 1, 2) <> ""
        i = i + 1
    Loop
    FindMax = i
End Function

'������ ��������� (������� �� ������� + �������)
Function Progress(ByVal cur As Long, ByVal all As Long)
    Progress = text + ":" + str(cur) + " ��" + str(all) + " (" + str(Int(cur / all * 100)) + "% )"
End Function

'����� ������ ������ (�������� �� ����������, ���� ����������� - ��������, ���� ��� - �������
Sub MethodSelect(ByVal name As String, ByVal first As Long, ByVal last As Long)
    For i = first To last - 1
        If StrComp(Sheets(name).Cells(i, 1), Sheets(name).Cells(i + 1, 1), vbTextCompare) > 0 Then
            SearchMethod = 1
            Exit For
        End If
    Next
End Sub

'����� �������� � �������
Function Search(ByVal name As String, ByVal str As String, ByVal first As Long, ByVal last As Long) As Long
    
    If SearchMethod = 0 Then
        
        '�������� �����
        Find = 0
        Do
            middle = first + Int((last - first) / 2)
            If StrComp(str, Sheets(name).Cells(first, 1), vbTextCompare) = 0 Then Find = first
            If StrComp(str, Sheets(name).Cells(last, 1), vbTextCompare) = 0 Then Find = last
            If StrComp(str, Sheets(name).Cells(middle, 1), vbTextCompare) = 0 Then Find = middle
            If StrComp(str, Sheets(name).Cells(middle, 1), vbTextCompare) < 0 Then last = middle
            If StrComp(str, Sheets(name).Cells(middle, 1), vbTextCompare) > 0 Then first = middle
        Loop Until Find > 0 Or last - first < 2
    
    Else
        
        '����� ���������
        For i = first To last
            If str = Sheets(name).Cells(i, 1) Then
                Find = i
                Exit For
            End If
        Next
    
    End If
    Search = Find

End Function

'������ �������� (��� �������� ����������)
Function StartProcess()
    startTime = Time
End Function

'��������������� ������� ����������
Function TimePredict(ByVal complate As Integer, ByVal all As Integer) As String
    timepassed = Time - startTime
    timeforone = timepassed / complate
    alltime = timeforone * (all - complate)
    st = CStr(Time + alltime)
    TimePredict = "   ������� ���������� �������� � " + Left(st, Len(st) - 3)
End Function

'******************** End of File ********************