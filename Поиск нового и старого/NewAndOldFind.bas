Attribute VB_Name = "NewAndOldFind"
'Версия 1.3 (23.12.2019) - Выделение кода из строки
'Версия 1.4 (24.12.2019) - Оптимизация сообщений
'Версия 1.5 (24.12.2019) - Поиск изменений
'Версия 1.6 (26.12.2019) - Рефакторинг

Const newTab = "Ноябрь"     'Новая таблица
Const fNew = 1              'Колонка в "Новой" таблице
Const oldTab = "Сентябрь"   'Старая таблица таблица
Const fOld = 1              'Колонка в "старых" таблицах
Const fields = 2            'Количество колонок для сравнени
Const maxTabs = 8           'Максимум полей

Global sn As Integer        'Счётчик новых строк
Global max As Integer       'Счётчик строк всего
Global oldTabs As Integer   'Счётчик старых таблиц


Sub NewAndOldFind()

    MakeCopy
    
    'Поиск удалённых в новой таблице
    AddDead (oldTab)
    'Для сверки с другой таблицей добавить с именем сюда:
    'AddNew ("")
    
    'Поиск новых записей
    FindNew (oldTab)
    'Для сверки с другой таблицей добавить с именем сюда:
    'FindDead ("")
    
    DeadSum
    Message "Готово!"
    
End Sub

'Подготовка итоговой таблицы
Private Sub MakeCopy()
    
    Message "Подготовка..."
    
    sn = 0
    mx = 0
    oldTabs = 0
    
    
    Sheets("Res").Cells.Clear
    maxStrings = 0
    max = 0
    Do While Sheets(newTab).Cells(max + 1, 1) <> ""
        max = max + 1
        For i = 1 To maxTabs
            Sheets("Res").Cells(max, i) = Sheets(newTab).Cells(max, i)
        Next
    Loop

End Sub

'Поиск и добавления удалённых строк
Private Sub AddDead(sheet)
    
    'Находим максимум в старой таблице
    Message "Подсчёт строк..."
    maxOld = 0
    Do Until Sheets(sheet).Cells(maxOld + 1, 1) = ""
        maxOld = maxOld + 1
    Loop
    
    changed = 0
    Find = False
    
    For i = 2 To maxOld
        Call ProgressBar("Поиск удалённых", i, maxOld)
        For j = 2 To max
            Find = True
            For k = 0 To fields - 1
                If Sheets("Res").Cells(j, fNew + k) <> Sheets(sheet).Cells(i, fOld + k) Then
                    Find = False
                    Exit For
                End If
            Next
            If Find Then
                Exit For
            End If
        Next
        If Find Then
            'Дополнительные действия при нахождении
                
                
                
            'Сравнение
            Change = False
            For k = 3 To 8
                'If Sheets("Res").Cells(j, k) = "" Then Sheets("Res").Cells(j, k) = "_"
                'If Sheets(sheet).Cells(i, k) = "" Then Sheets(sheet).Cells(i, k) = "_"
                If Sheets("Res").Cells(j, k) <> Sheets(sheet).Cells(i, k) Then
                    Sheets("Res").Cells(j, k).Interior.Color = vbRed
                    Change = True
                End If
            Next
            If Change Then
                Sheets("Res").Cells(j, maxTabs + 3) = "Изменён"
                changed = changed + 1
            End If
            
            
                
        Else
            For c = 1 To maxTabs
                Sheets("Res").Cells(max, c) = Sheets(sheet).Cells(i, c)
            Next
            Sheets("Res").Cells(max, maxTabs + 2) = "Удалён (Был в" + sheet + ", но не стало в " + newTab + ")"
            max = max + 1
            sn = sn + 1
        End If
    Next
    
    Sheets("Res").Cells(max + 1, maxTabs + 2) = "Удалено:" + Str(sn)
    Sheets("Res").Cells(max + 1, maxTabs + 3) = "Изменено:" + Str(changed)
    
End Sub

'Поиск новых строк
Private Sub FindNew(sheet)
    
    oldTabs = oldTabs + 1
    
    'Находим максимум в старой таблице
    Message "Подсчёт строк..."
    maxOld = 0
    Do Until Sheets(sheet).Cells(maxOld + 1, 1) = ""
        maxOld = maxOld + 1
    Loop
    
    
    For i = 1 To max - sn - 1
        
        Call ProgressBar("Поиск новых", i, max - sn - 1)
        
        Find = False
        For j = 1 To maxOld
            If Sheets("Res").Cells(i, fNew) = Sheets(sheet).Cells(j, fOld) Then
                Find = True
                Exit For
            End If
        Next
        If Not Find Then
            If Sheets("Res").Cells(i, maxTabs + 1) = "" Then
                Sheets("Res").Cells(i, maxTabs + 1) = 1
            Else
                Sheets("Res").Cells(i, maxTabs + 1) = Val(Sheets("Res").Cells(i, maxTabs + 1)) + 1
            End If
        End If
    Next
    
    
End Sub

'Итог по новым
Private Sub DeadSum()
    
    Message "Завершение..."
    
    s = 0
    For i = 1 To max
        If Sheets("Res").Cells(i, maxTabs + 1) = oldTabs Then
            Sheets("Res").Cells(i, maxTabs + 1) = "Новый! (Появился в " + newTab + ", раньше не встречался)"
            s = s + 1
        Else
            Sheets("Res").Cells(i, maxTabs + 1) = ""
        End If
    Next
    Sheets("Res").Cells(max + 1, maxTabs + 1) = "Новых:" + Str(s)
    
End Sub

Private Sub ProgressBar(text As String, ByVal cur As Integer, ByVal all As Integer)
    If cur Mod 50 = 0 Then
        Application.ScreenUpdating = True
        Application.StatusBar = text + ":" + _
            Str(cur) + " из" + Str(all) + " (" + Str(Int(cur / all * 100)) + "% )"
        Application.ScreenUpdating = False
    End If
End Sub

'Вывод сообщение в статусбар
Private Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    Application.ScreenUpdating = False
End Sub



'Выделение кода из строки и помещение его в отдельную ячейку
Sub CodeEject()
    
    codeFrom = 1
    codeTo = 5

    Application.ScreenUpdating = True
    Application.StatusBar = "Подсчёт строк..."
    Application.ScreenUpdating = False
    max = 1
    Do Until Cells(max + 1, 1) = ""
        max = max + 1
    Loop
    
    For i = 2 To max
        Call ProgressBar("Обработка", i, max)
        c = Cells(i, codeFrom)
        If InStr(1, c, "код: ", vbTextCompare) Then
            ss = Split(c, "код: ")
            If UBound(ss) > 0 Then
                c = Replace(ss(1), " ", "")
                kv = InStr(1, c, """", vbTextCompare)
                If kv > 0 Then c = Left(c, kv - 1)
            End If
        End If
        If InStr(1, c, "Код Объекта: ", vbTextCompare) Then
            ss = Split(c, "Код Объекта: ")
            If UBound(ss) > 0 Then
                c = Replace(ss(1), " ", "")
                kv = InStr(1, c, """", vbTextCompare)
                If kv > 0 Then c = Left(s, kv - 1)
            End If
        End If
        If InStr(1, c, "(", vbTextCompare) Then
            ss = Split(c, "(")
            If UBound(ss) > 0 Then
                c = Replace(ss(1), " ", "")
                kv = InStr(1, c, "(", vbTextCompare)
                If kv > 0 Then c = Left(c, kv - 1)
                kv = InStr(1, c, ")", vbTextCompare)
                If kv > 0 Then c = Left(c, kv - 1)
            End If
        End If
        
        'Перед записью надо проверить всё ли то что нашли цифры
        d = True
        For j = 1 To Len(c)
            code = Asc(Mid(c, j, 1))
            If code < 48 Or code > 57 Then d = False
        Next
        If Not d Then c = ""
        Cells(i, codeTo) = c
    Next
    
    Message "Готово!"

End Sub
