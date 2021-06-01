Attribute VB_Name = "Misc"
'Last change: 01.06.2021 09:17

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
Function FindMax(ByVal name As String) As Long
    i = 0
    Do While Sheets(name).Cells(i + 1000, 1) <> ""
        i = i + 1000
    Loop
    Do While Sheets(name).Cells(i + 1, 1) <> ""
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

'******************** End of File ********************