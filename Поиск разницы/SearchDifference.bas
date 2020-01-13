Attribute VB_Name = "SearchDifference"
'Версия 1.0 (19.12.2019)
'Версия 1.2 (20.12.2019)
'Версия 1.3 (23.12.2019) - Выделение кода из строки
'Версия 1.4 (24.12.2019) - Оптимизация сообщений
'Версия 1.5 (24.12.2019) - Поиск изменений
'Версия 1.6 (26.12.2019) - Рефакторинг
'Версия 1.7 (10.01.2020) - Вычисление разницы

Const newTab = "УФА"     'Новая таблица
Const fNew = 1              'Колонка в "Новой" таблице
Const oldTab = "Access"     'Старая таблица таблица
Const fOld = 1              'Колонка в "старых" таблицах
Const fields = 1            'Количество колонок для сравнени
Const maxTabs = 3           'Максимум полей

Global sn As Integer        'Счётчик новых строк
Global max As Integer       'Счётчик строк всего
Global oldTabs As Integer   'Счётчик старых таблиц


Sub Start()

    MakeCopy
    AddDead
    FindNew
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
Private Sub AddDead()
    
    'Находим максимум в старой таблице
    Message "Подсчёт строк..."
    maxOld = 0
    Do Until Sheets(oldTab).Cells(maxOld + 1, 1) = ""
        maxOld = maxOld + 1
    Loop
    
    mached = 0
    changed = 0
    
    Find = False
    For i = 2 To maxOld
        Call ProgressBar("Поиск удалённых", i, maxOld)
        For j = 2 To max
            Find = True
            For k = 0 To fields - 1
                If Sheets("Res").Cells(j, fNew + k) <> Sheets(oldTab).Cells(i, fOld + k) Then
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
            s = 3
            Sheets("Res").Cells(j, s + 1) = Sheets(oldTab).Cells(i, s)
            i1 = Sheets("Res").Cells(j, s)
            i2 = Sheets(oldTab).Cells(i, s)
            rz = Round(i1 - i2, 2)
            Sheets("Res").Cells(j, s + 2) = rz
            If Abs(rz) = 0 Then
                Sheets("Res").Cells(j, s + 0).Interior.Color = RGB(128, 255, 128)
                Sheets("Res").Cells(j, s + 1).Interior.Color = RGB(128, 255, 128)
                Sheets("Res").Cells(j, s + 2).Interior.Color = RGB(196, 255, 196)
                mached = mached + 1
                Sheets("Res").Cells(j, maxTabs + 5) = "Совпал"
            End If
            If Abs(rz) > 0 Then
                Sheets("Res").Cells(j, s + 0).Interior.Color = RGB(255, 255, 128)
                Sheets("Res").Cells(j, s + 1).Interior.Color = RGB(255, 255, 128)
                Sheets("Res").Cells(j, s + 2).Interior.Color = RGB(255, 255, 196)
                mached = mached + 1
                Sheets("Res").Cells(j, maxTabs + 5) = "Совпал (почти)"
            End If
            If Abs(rz) > 10 Then
                Sheets("Res").Cells(j, s + 0).Interior.Color = RGB(255, 128, 128)
                Sheets("Res").Cells(j, s + 1).Interior.Color = RGB(255, 128, 128)
                Sheets("Res").Cells(j, s + 2).Interior.Color = RGB(255, 196, 196)
                changed = changed + 1
                Sheets("Res").Cells(j, maxTabs + 5) = "Изменён"
            End If
            
        Else
            max = max + 1
            For c = 1 To maxTabs
                Sheets("Res").Cells(max, c) = Sheets(oldTab).Cells(i, c)
            Next
            Sheets("Res").Cells(max, maxTabs + 5) = "Удалён (Был в " + oldTab + ", но не стало в " + newTab + ")"
            sn = sn + 1
        End If
    Next
    
    Sheets("Res").Cells(max + 3, maxTabs + 5) = "Удалено:" + Str(sn)
    Sheets("Res").Cells(max + 4, maxTabs + 5) = "Совпало:" + Str(mached)
    Sheets("Res").Cells(max + 5, maxTabs + 5) = "Изменено:" + Str(changed)
    
End Sub

'Поиск новых строк
Private Sub FindNew()
    
    'Находим максимум в старой таблице
    Message "Подсчёт строк..."
    maxOld = 0
    Do Until Sheets(oldTab).Cells(maxOld + 1, 1) = ""
        maxOld = maxOld + 1
    Loop
    
    s = 0
    For i = 2 To max - sn
        
        Call ProgressBar("Поиск новых", i, max - sn - 1)
        
        Find = False
        For j = 2 To maxOld
            If Sheets("Res").Cells(i, fNew) = Sheets(oldTab).Cells(j, fOld) Then
                Find = True
                Exit For
            End If
        Next
        If Not Find Then
            Sheets("Res").Cells(i, maxTabs + 5) = "Новый! (Появился в " + newTab + ")"
            s = s + 1
        End If
    Next
    Sheets("Res").Cells(max + 2, maxTabs + 5) = "Новых:" + Str(s)
    
End Sub

'Вывод статуса обработки
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


