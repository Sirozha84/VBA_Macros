Attribute VB_Name = "OSRPKU"
'������ 1.0 (11.12.2010)
'������ 1.1 (12.12.2019) - ����������� �������� �������� ������
'������ 1.2 (30.12.2019) - ����������� ��������� � ���� ��������� �� �����������
'������ 1.3 (03.01.2020) - ����������� ��������� ������ �� ����������� "������",
'   ����������� ������ (�������� �����������, �� �� �� ��� ��� �������� ��
'������ 1.4 (04.01.2020) - ����������� � ��������� ������������� ("�"/"�", �������/��������� �����)

Const adrSh = "Adresses"
Const tempSh = "Temp"
Const resultSh = "Result"

Public tabs As Integer      '���������� �������
Public max As Long          '���������� ����� �����

Private Type adrIndex
    adr As String
    iStart As Long
    iEnd As Long
End Type

Sub Start()

'max = 175407
'tabs = 24

    DataCollection
    AdressMatching
    Filter
    
    Message "������!"
    
End Sub

'���� ������ �� ���� ������
Private Sub DataCollection()
    
    CreateSheet tempSh
    Dim iWS As Worksheet
    s = 1
    max = 2
    For Each iWS In ThisWorkbook.Worksheets
        Call ProgressBar("���� 1: ����������� ������", s, ThisWorkbook.Worksheets.Count)
        If (iWS.name <> tempSh And iWS.name <> adrSh And iWS.name <> resultSh) Then
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

'������������� �� ������������ �������
Private Sub AdressMatching()

    '����������
    
    Message "���������� �������"
    
    Sheets(adrSh).Select
    iMax = 1
    i = 4
    ReDim indexes(iMax) As adrIndex
    indexes(iMax).adr = Cells(i, 1) + Cells(i, 2) + CStr(Cells(i, 3))
    indexes(iMax).iStart = i
    indexes(iMax).iEnd = i
    Do While Cells(i, 1) <> ""
        If indexes(iMax).adr = Cells(i, 1) + Cells(i, 2) + CStr(Cells(i, 3)) Then
            indexes(iMax).iEnd = i
        Else
            iMax = iMax + 1
            ReDim Preserve indexes(iMax) As adrIndex
            indexes(iMax).adr = PrepareAdress(Cells(i, 1) + Cells(i, 2) + CStr(Cells(i, 3)))
            indexes(iMax).iStart = i
            indexes(iMax).iEnd = i
        End If
        i = i + 1
    Loop
    'For i = 1 To iMax
    '    Cells(i, 25) = indexes(i).adr
    '    Cells(i, 26) = indexes(i).iStart
    '    Cells(i, 27) = indexes(i).iEnd
    'Next
    
    '�����
    
    Sheets(tempSh).Select
    For j = 0 To 6
        Cells(1, tabs + 1 + j) = Sheets(adrSh).Cells(2, 8 + j)
    Next
    For i = 2 To max
        If i Mod 500 = 0 Then Call ProgressBar("���� 2: ������������� �� ������������", i, max)
        Find = False
        For j = 1 To iMax
            adr = PrepareAdress(Cells(i, 1) + Cells(i, 2) + CStr(Cells(i, 3)))
            If adr = indexes(j).adr Then
                Find = True
                Exit For
            End If
        Next
        If Find Then
            For k = indexes(j).iStart To indexes(j).iEnd
                If CStr(Cells(i, 4)) = CStr(Sheets(adrSh).Cells(k, 4)) And _
                   CStr(Cells(i, 5)) = CStr(Sheets(adrSh).Cells(k, 5)) And _
                   CStr(Cells(i, 6)) = CStr(Sheets(adrSh).Cells(k, 6)) Then
                    For l = 0 To 6
                        Cells(i, tabs + 1 + l) = Sheets(adrSh).Cells(k, 8 + l)
                    Next
                    Exit For
                End If
            Next
        End If
    Next
    tabs = tabs + 7
    
End Sub

Private Function PrepareAdress(adr As String) As String
    adr = LCase(adr)
    PrepareAdress = Replace(adr, "�", "�")
End Function

'���������� �� ����������
Private Sub Filter()
   
    c_adr = 8       '���� � �������
    c_usl = 17      '���� � �������
    c_vip = 15      '���������� �����
    
    ReDim tmp(max) As String
    Sheets(resultSh).Select
    Sheets(resultSh).Cells.Clear
    
    '������ ��������� ������������� ������
    
    mx = 1 '����� ��������������� �����
    For i = 2 To max
        If i Mod 5000 = 0 Then Call ProgressBar("���� 3: ����������", i, max)
        If Sheets(tempSh).Cells(i, c_usl) = "���������" And Sheets(tempSh).Cells(i, c_vip) <> 0 Then
            tmp(mx) = Sheets(tempSh).Cells(i, 1) + "," + _
                      Sheets(tempSh).Cells(i, 2) + "," + _
                      CStr(Sheets(tempSh).Cells(i, 3))
            'Sheets(tempSh).Cells(i, 20) = tmp(mx)
            mx = mx + 1
        End If
    Next
    mx = mx - 1
    
    '������ ������, ������� �������� ��� ������
    
    '�����
    For i = 1 To tabs
        Cells(1, i) = Sheets(tempSh).Cells(1, i)
    Next
    f = 1
    Dim adr As String
    For i = 2 To max
        If Sheets(tempSh).Cells(i, 1) = "" Then Exit For
        If i Mod 1000 = 0 Then Call ProgressBar("���� 4: ������", i, max)
        fnd = False
        adr = Sheets(tempSh).Cells(i, 1) + "," + _
              Sheets(tempSh).Cells(i, 2) + "," + _
              CStr(Sheets(tempSh).Cells(i, 3))
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
                Cells(f + 1, c) = Sheets(tempSh).Cells(i, c)
            Next
            f = f + 1
        End If
    Next

    
End Sub





'����������� ���� ��������� � ���������
Private Sub YearsFloorsSlow()
    
    Message "����������"
    
    Sheets(adrSh).Select
    i = 4
    Do While Cells(i, 1) <> ""
        i = i + 1
    Loop
    aMax = i - 1
    
    ReDim adresses(1 To aMax + 1) As String
    
    For i = 4 To aMax
        adresses(i) = Cells(i, 1) + Cells(i, 2) + CStr(Cells(i, 3)) + Cells(i, 4) + "," + _
        CStr(Cells(i, 5)) + "," + CStr(Cells(i, 6))
    Next
    
    Sheets(resultSh).Select
    For j = 0 To 6
        Cells(1, tabs + 2 + j) = Sheets(adrSh).Cells(2, 8 + j)
    Next
    For i = 2 To max
        If i Mod 100 = 0 Then Call ProgressBar("���� 4: ������������� �� ������������", i, max)
        adr = Cells(i, 1) + Cells(i, 2) + CStr(Cells(i, 3)) + Cells(i, 4) + "," + _
        CStr(Cells(i, 5)) + "," + CStr(Cells(i, 6))

        adresses(aMax + 1) = adr
       
        z = WorksheetFunction.Match(adr, adresses, 0)
        If z < max + 1 Then
            For j = 0 To 6
                Sheets(resultSh).Cells(i, tabs + 2 + j) = Sheets(adrSh).Cells(z, 8 + j)
            Next
        End If
    Next
    
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
