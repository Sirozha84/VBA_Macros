Attribute VB_Name = "SearchDifference"
'Версия 1.0 (19.12.2019)
'Версия 1.2 (20.12.2019)
'Версия 1.3 (23.12.2019) - Выделение кода из строки
'Версия 1.4 (24.12.2019) - Оптимизация сообщений
'Версия 1.5 (24.12.2019) - Поиск изменений
'Версия 1.6 (26.12.2019) - Рефакторинг
'Версия 1.7 (10.01.2020) - Вычисление разницы
'Версия 1.8 (17.01.2020) - Правка итоговой таблицы


Const newTab = "УФА"        'Новая таблица
Const oldTab = "Access"     'Старая таблица таблица
Const maxTabs = 3           'Максимум полей
Const fCom = 8

Global sn As Integer        'Счётчик новых строк
Global max As Integer       'Счётчик строк всего
Global oldTabs As Integer   'Счётчик старых таблиц

Global subNew As Long
Global sumOld As Long
Global sumDif As Long

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
    max = 1
    Do While Sheets(newTab).Cells(max + 1, 1) <> ""
        max = max + 1
        For i = 1 To 3
            Sheets("Res").Cells(max, i) = Sheets(newTab).Cells(max, i)
        Next
    Loop
    
    'Дорисовываем шапку
    Sheets("Res").Cells(1, 1) = "Договор"
    Sheets("Res").Cells(1, 3) = newTab
    Sheets("Res").Cells(1, 5) = oldTab
    Sheets("Res").Cells(1, 6) = "Разница"
    
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
            Find = False
            If Sheets("Res").Cells(j, 1) = Sheets(oldTab).Cells(i, 1) Then
                Find = True
                Exit For
            End If
        Next
        If Find Then
            'Если нашли, тогда сравниваем
            Sheets("Res").Cells(j, 4) = Sheets(oldTab).Cells(i, 2)
            Sheets("Res").Cells(j, 5) = Sheets(oldTab).Cells(i, 3)
            rz = Round(Sheets("Res").Cells(j, 3) - Sheets("Res").Cells(j, 5), 2)
            Sheets("Res").Cells(j, 6) = rz
            If Abs(rz) = 0 Then
                Sheets("Res").Cells(j, 3).Interior.Color = RGB(128, 255, 128)
                Sheets("Res").Cells(j, 5).Interior.Color = RGB(128, 255, 128)
                Sheets("Res").Cells(j, 6).Interior.Color = RGB(196, 255, 196)
                mached = mached + 1
                Sheets("Res").Cells(j, fCom) = "Совпал"
            End If
            If Abs(rz) > 0 Then
                Sheets("Res").Cells(j, 3).Interior.Color = RGB(255, 255, 128)
                Sheets("Res").Cells(j, 5).Interior.Color = RGB(255, 255, 128)
                Sheets("Res").Cells(j, 6).Interior.Color = RGB(255, 255, 196)
                mached = mached + 1
                Sheets("Res").Cells(j, fCom) = "Совпал (почти)"
            End If
            If Abs(rz) > 10 Then
                Sheets("Res").Cells(j, 3).Interior.Color = RGB(255, 128, 128)
                Sheets("Res").Cells(j, 5).Interior.Color = RGB(255, 128, 128)
                Sheets("Res").Cells(j, 6).Interior.Color = RGB(255, 196, 196)
                changed = changed + 1
                Sheets("Res").Cells(j, fCom) = "Изменён"
            End If
            sumDif = sumDif + rz
        Else
            'А если не нашли, копируем
            max = max + 1
            Sheets("Res").Cells(max, 1) = Sheets(oldTab).Cells(i, 1)
            Sheets("Res").Cells(max, 4) = Sheets(oldTab).Cells(i, 2)
            Sheets("Res").Cells(max, 5) = Sheets(oldTab).Cells(i, 3)
            Sheets("Res").Cells(max, fCom) = "Есть в " + oldTab + ", но нет в " + newTab
            
            'Рисование разницы
            'Sheets("Res").Cells(max, 3).Cells = 0
            Sheets("Res").Cells(max, 6).Cells = -Sheets("Res").Cells(max, 5).Cells
            Sheets("Res").Cells(max, 3).Interior.Color = RGB(255, 128, 128)
            Sheets("Res").Cells(max, 5).Interior.Color = RGB(255, 128, 128)
            Sheets("Res").Cells(max, 6).Interior.Color = RGB(255, 196, 196)
            
            sumDif = sumDif - Sheets("Res").Cells(max, 3).Cells
            
            sn = sn + 1
        End If
    Next
    
    Sheets("Res").Cells(max + 3, maxTabs + 5) = "Есть только в " + oldTab + ":" + Str(sn)
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
            If Sheets("Res").Cells(i, 1) = Sheets(oldTab).Cells(j, 1) Then
                Find = True
                Exit For
            End If
        Next
        If Not Find Then
            
            'Рисование "разницы"
            Sheets("Res").Cells(i, 6).Cells = Sheets("Res").Cells(i, maxTabs).Cells
            Sheets("Res").Cells(i, 3).Interior.Color = RGB(255, 128, 128)
            Sheets("Res").Cells(i, 5).Interior.Color = RGB(255, 128, 128)
            Sheets("Res").Cells(i, 6).Interior.Color = RGB(255, 196, 196)
            
            Sheets("Res").Cells(i, fCom) = "Есть в " + newTab + ", но нет в " + oldTab
            sumDif = sumDif + Sheets("Res").Cells(i, maxTabs).Cells
            
            s = s + 1
            
        End If
    Next
    
    Sheets("Res").Cells(max + 2, 3) = sumNew
    Sheets("Res").Cells(max + 2, 4) = sumOld
    Sheets("Res").Cells(max + 2, 5) = sumDif
    
    Sheets("Res").Cells(max + 2, maxTabs + 5) = "Есть только в " + newTab + ":" + Str(s)
    
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


