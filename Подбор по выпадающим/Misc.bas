'Версия 1.0 (13.01.2020) - Вынос дополнииельных процедур в отдельный файл
'Версия 1.1 (26.02.2020) - Две новых процедуры: OneToTwo и Vipad

Attribute VB_Name = "Misc"
'Подбор адресов по лицевым счетам
'Данные находятся на одной таблице, объединяет их только номер лс, адреса переносятся с одной в другую
Sub AdresesByLS()
    Const ft = 1    'Ячейка с ЛС в первая таблица
    Const st = 25   'Ячейка с ЛС во второй таблице
    Const atb = 6   'Количество столбцов с адресом
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

'Дедупликация домов
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

'Подбор коэффициентов
Sub Koef()
    col = 25
    Max = 153519
    Sheets("Result").Select
    For i = 1 To Max
        If i Mod 1000 = 0 Then Call ProgressBar("Подбор коэффициентов", i, Max)
        adr = Cells(i, 2) + CStr(Cells(i, 3)) + Cells(i, 4)
        For j = 1 To 1148
            If adr = Sheets("Adresses").Cells(j, 15) Then
                Cells(i, col) = Sheets("Adresses").Cells(j, 16)
                Exit For
            End If
        Next
    Next
    Message "Готово!"
End Sub

'Замена первого символа на "двойку"
Sub OneToTwo()
    Max = 39689
    For i = 4 To Max
        If i Mod 1000 = 0 Then Call ProgressBar("Обработка", i, Max)
        Cells(i, 1) = "2" + Right(Cells(i, 1), Len(Cells(i, 1)) - 1)
    Next
    Message "Готово!"
End Sub

'Подстановка выпадающего с другой таблицы (ориентация на номер лицевого счёта)
Sub Vipad()
    Max = 39163
    last = 4
    For i = 2 To Max
        If i Mod 100 = 0 Then Call ProgressBar("Обработка", i, Max)
        For j = last To 39689
            Find = False
            If Sheets("Отопление").Cells(i, 7) = Sheets("Vip").Cells(j, 1) Then
                Sheets("Отопление").Cells(i, 17) = Sheets("Vip").Cells(j, 10)
                last = j
                Find = True
                Exit For
            End If
            If Not Find Then last = 4
        Next
    Next
    Message "Готово!"
End Sub

'Рисование прогресса, text - имя, cur - текущее значение, all - всего, отображать каждые over штук
Private Sub ProgressBar(text As String, ByVal cur As Long, ByVal all As Long)
    Application.ScreenUpdating = True
    Application.StatusBar = text + ":" + Str(cur) + " из" + Str(all) + _
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
