Attribute VB_Name = "OSRPKU"
'������ 1.0 (11.12.2010)
'������ 1.1 (12.12.2019) - ����������� �������� �������� ������

Const adrSh = "Adresses"
Const tempSh = "Temp"
Const resultSh = "Result"

Public tabs As Integer      '���������� �������
Public max As Long          '���������� ����� �����

Sub Start()
    
    tabs = 17 '��������, �������
    max = 105147 '��������, �������
    
    'Prepare
    'Filter
    YearsFloors
    
    Message "������!"
    
End Sub

'����������
Private Sub Prepare()
    CreateSheet tempSh
    Dim iWS As Worksheet
    s = 1
    max = 2
    For Each iWS In ThisWorkbook.Worksheets
        Call ProgressBar("����������� ������ �� ���������", s, ThisWorkbook.Worksheets.Count)
        If (iWS.name <> tempSh And iWS.name <> adrSh) Then
            '����������� ����� �� ������ ��������
            If Sheets(tempSh).Cells(1, 1) = "" Then
                tabs = 1
                Do While iWS.Cells(1, tabs) <> ""
                    Sheets(tempSh).Cells(1, tabs) = iWS.Cells(1, tabs)
                    tabs = tabs + 1
                Loop
                Sheets(tempSh).Cells(1, tabs) = "������"
            End If
            '����������� ������
            i = 2
            Do While iWS.Cells(i, 1) <> ""
                For j = 1 To tabs - 1
                    Sheets(tempSh).Cells(max, j) = iWS.Cells(i, j)
                Next
                Sheets(tempSh).Cells(max, tabs) = iWS.name
                i = i + 1
                max = max + 1
            Loop
        End If
        s = s + 1
    Next iWS
    max = max - 1
    CreateSheet resultSh
End Sub

'���������� �� ����������
Private Sub Filter()
   
    c_adr = 8       '���� � �������
    c_usl = 17      '���� � �������
    c_vip = 15      '���������� �����
    
    ReDim tmp(max) As String
    
    Sheets(resultSh).Cells.Clear
    
    '������ ��������� ������������� ������
    mx = 1 '����� ��������������� �����
    For i = 2 To max
        If i Mod 1000 = 0 Then Call ProgressBar("����������", i, max)
        If Sheets(tempSh).Cells(i, c_usl) = "���������" And Sheets(tempSh).Cells(i, c_vip) <> 0 Then
            tmp(mx) = Sheets(tempSh).Cells(i, 1) + "," + _
                      Sheets(tempSh).Cells(i, 2) + "," + _
                      Sheets(tempSh).Cells(i, 3)
            mx = mx + 1
        End If
    Next
    mx = mx - 1
    
    '���� ������, ������� �������� ��� ������
    f = 1
    For i = 2 To max
        If Sheets(tempSh).Cells(i, 1) = "" Then Exit For
        If i Mod 100 = 0 Then Call ProgressBar("����������", i, max)
        fnd = False
        adr = Sheets(tempSh).Cells(i, 1) + "," + _
              Sheets(tempSh).Cells(i, 2) + "," + _
              Sheets(tempSh).Cells(i, 3)
        If adr = last Then
            fnd = True
        Else
            For j = 2 To mx
                If adr = tmp(j) Then
                    fnd = True
                    last = adr
                    Exit For
                End If
            Next
        End If
        If fnd Then
            For c = 1 To tabs
                Sheets(resultSh).Cells(f + 1, c) = Sheets(tempSh).Cells(i, c)
            Next
            f = f + 1
        End If
    Next
    
    '�������� �����
    For i = 1 To tabs
        Sheets(resultSh).Cells(1, i) = Sheets(tempSh).Cells(1, i)
    Next
    
End Sub

'����������� ���� ��������� � ���������
Private Sub YearsFloors()
    Sheets(resultSh).Cells(1, tabs + 1) = "��� ���������"
    Sheets(resultSh).Cells(1, tabs + 2) = "���������"
    last = ""
    yr = 0
    floors = 0
    For i = 2 To max
        If i Mod 100 = 0 Then Call ProgressBar("����������", i, max)
        adr = Sheets(resultSh).Cells(i, 1) + "," + _
              Sheets(resultSh).Cells(i, 2) + "," + _
              Sheets(resultSh).Cells(i, 3)
        If adr <> last Then
            last = adr
            
            '�����
            j = 4
            Do While Sheets(adrSh).Cells(j, 1) <> ""
                inBook = Sheets(adrSh).Cells(j, 1) + "," + _
                       Sheets(adrSh).Cells(j, 2) + "," + _
                       Right(Str(Sheets(adrSh).Cells(j, 3)), Len(Str(Sheets(adrSh).Cells(j, 3))) - 1)
                If adr = inBook Then
                    yr = Sheets(adrSh).Cells(j, 9)
                    floors = Sheets(adrSh).Cells(j, 8)
                    Exit Do
                End If
                j = j + 1
            Loop
        End If
        Sheets(resultSh).Cells(i, tabs + 1) = yr
        Sheets(resultSh).Cells(i, tabs + 2) = floors
    Next
End Sub

'��������� ���������, text - ���, cur - ������� ��������, all - �����, ���������� ������ over ����
Private Sub ProgressBar(text As String, ByVal cur As Long, ByVal all As Long)
    Application.ScreenUpdating = True
    Application.StatusBar = text + ":" + Str(cur) + " ��" + Str(all) + _
        " (" + Str(Int(cur / all * 100)) + "% )"
    Application.ScreenUpdating = False
End Sub

Private Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    Application.ScreenUpdating = False
End Sub

Private Sub CreateSheet(name As String)
    If Not SheetExist(name) Then
        Sheets.Add(, Sheets(Sheets.Count)).name = name
    End If
    Sheets(name).Cells.Clear
End Sub

'�������� �� ������������� �����
Function SheetExist(name As String) As Boolean
    Dim objSheet As Object
    
    On Error GoTo HandleError
    ThisWorkbook.Worksheets(name).Activate
    SheetExist = True
    Exit Function
    
HandleError:
    SheetExist = False
End Function
