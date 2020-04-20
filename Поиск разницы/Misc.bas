Attribute VB_Name = "Misc"
'Создание результирующей таблицы
Sub NewTab(name As String, create As Boolean)
    If create Then
        If Not SheetExist(name) Then
            Sheets.Add(, Sheets(Sheets.Count)).name = name
        End If
    End If
    If Not SheetExist(name) Then
        LabelStatus = "Ошибка: Результирующей таблицы не существует"
        Err.Raise (0)
    End If
    Sheets(name).Cells.Clear
End Sub

'Проверка на существование листа
Private Function SheetExist(name As String) As Boolean
    Dim objSheet As Object
    On Error GoTo HandleError
    ThisWorkbook.Worksheets(name).Activate
    SheetExist = True
    Exit Function
HandleError:
    SheetExist = False
End Function

'Поиск максимального количества строк (последняя строка где первая колонка непрерывно заполнена)
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

'Расчёт прогресса (сколько из скольки + процент)
Function Progress(ByVal cur As Integer, ByVal all As Integer)
    Progress = text + ":" + str(cur) + " из" + str(all) + " (" + str(Int(cur / all * 100)) + "% )"
End Function

'Бинарный поиск
Function Search(ByVal name As String, ByVal str As String, ByVal first As Long, ByVal last As Long)
    Find = False
    Do
        middle = first + Int((last - first) / 2)
        If Sheets(Tab2).Cells(ii, 1) = Sheets(TabRes).Cells(first, 1) Then Find = True
        If Sheets(Tab2).Cells(ii, 1) = Sheets(TabRes).Cells(middle, 1) Then Find = True
        If Sheets(Tab2).Cells(ii, 1) = Sheets(TabRes).Cells(last, 1) Then Find = True
        If StrComp(Sheets(Tab2).Cells(ii, 1), Sheets(TabRes).Cells(middle, 1), vbTextCompare) < 0 Then last = middle
        If StrComp(Sheets(Tab2).Cells(ii, 1), Sheets(TabRes).Cells(middle, 1), vbTextCompare) > 0 Then first = middle
    Loop Until Find Or last - first <= 2
    Search = Find
End Function
