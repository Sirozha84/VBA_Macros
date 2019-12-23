Attribute VB_Name = "NewAndOldFind"
Const fNew = 1          'Колонка в "Новой" таблице
Const fOld = 2          'Колонка в "старых" таблицах
Const method = 2        'Метод сравнения
Const newTab = "УФА"    '"Новая" таблица
'Const newTab = "УФА тест"    '"Новая" таблица
Const maxTabs = 4       'Максимум полей

Global sn As Integer        'Счётчик новых строк
Global max As Integer       'Счётчик строк всего
Global oldTabs As Integer   'Счётчик старых таблиц

Sub NewAndOldFind()

    MakeCopy
    
    'AddNew("ХВСиВО")
    'AddNew("Тепло")
    AddNew ("Access")
    'AddNew ("Assecc тест")
    
    'FindDead("ХВСиВО")
    'FindDead("Тепло")
    'FindDead ("Assecc тест")
    FindDead ("Access")
    
    DeadSum
    
End Sub

'Подготовка итоговой таблицы
Private Sub MakeCopy()
    
    Application.ScreenUpdating = True
    Application.StatusBar = "Подготовка..."
    Application.ScreenUpdating = False
    
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

'Поиск и добавления новых строк
Private Sub AddNew(sheet)
    
    'Находим максимум в старой таблице
    Application.ScreenUpdating = True
    Application.StatusBar = "Подсчёт строк..."
    Application.ScreenUpdating = False
    maxOld = 0
    Do Until Sheets(sheet).Cells(maxOld + 1, 1) = ""
        maxOld = maxOld + 1
    Loop
    
    For i = 1 To maxOld
    
        If i Mod 50 = 0 Then
            Application.ScreenUpdating = True
            Application.StatusBar = "Поиск новых:" + Persent(i, maxOld)
            Application.ScreenUpdating = False
        End If
    
        Find = False
        
        For j = 1 To max
            If Sheets("Res").Cells(j, 2) <> "" Then
                If Compare(Sheets("Res").Cells(j, fNew), Sheets(sheet).Cells(i, fOld)) Then
                    Find = True
                End If
            Else
                Exit For
            End If
        Next
        If Not Find Then
            For c = 1 To maxTabs
                Sheets("Res").Cells(max, c) = Sheets(sheet).Cells(i, c)
            Next
            Sheets("Res").Cells(max, maxTabs + 1) = "Новый из " + sheet
            max = max + 1
            sn = sn + 1
        End If
    Next
    
    Sheets("Res").Cells(max, maxTabs + 1) = "Новых:" + Str(sn)
    
End Sub

'Поиск удалённых строк
Private Sub FindDead(sheet)
    
    oldTabs = oldTabs + 1
    
    'Находим максимум в старой таблице
    Application.ScreenUpdating = True
    Application.StatusBar = "Подсчёт строк..."
    Application.ScreenUpdating = False
    maxOld = 0
    Do Until Sheets(sheet).Cells(maxOld + 1, 1) = ""
        maxOld = maxOld + 1
    Loop
    
    
    For i = 1 To max - sn - 1
        
        If i Mod 50 = 0 Then
            Application.ScreenUpdating = True
            Application.StatusBar = "Поиск удалённых:" + Persent(i, max - sn - 1)
            Application.ScreenUpdating = False
        End If
        
        Find = False
        For j = 1 To maxOld
            If Compare(Sheets("Res").Cells(i, fNew), Sheets(sheet).Cells(j, fOld)) Then
                Find = True
                Exit For
            End If
        Next
        If Not Find Then
            If Sheets("Res").Cells(i, maxTabs + 2) = "" Then
                Sheets("Res").Cells(i, maxTabs + 2) = 1
            Else
                Sheets("Res").Cells(i, maxTabs + 2) = Val(Sheets("Res").Cells(i, maxTabs + 2)) + 1
            End If
        End If
    Next
    
    
End Sub

'Итог по удалённым
Private Sub DeadSum()
        
    Application.StatusBar = "Завершение..."
    
    s = 0
    For i = 1 To max
        If Sheets("Res").Cells(i, maxTabs + 2) = oldTabs Then
            Sheets("Res").Cells(i, maxTabs + 2) = "Удалён!"
            s = s + 1
        Else
            'Sheets("Res").Cells(i, maxTabs + 2) = ""
        End If
    Next
    Sheets("Res").Cells(max, maxTabs + 2) = "Удалено:" + Str(s)
    
    Application.StatusBar = "Готово!"
End Sub

'Процедура сравнения
Private Function Compare(newCell As String, oldCell As String) As Boolean
    
    Compare = False
    
    'Метод сравнения 0 - тупо равенство
    If method = 0 Then
        Compare = (newCell = oldCell)
    End If
    
    'Метод сравнения 1 - сравнение последних 5-и символов
    If method = 1 Then
        Compare = (Left(newCell, 5) = Left(oldCell, 5))
    End If
    
    'Метод сравнения 2 - сравнение по коду в строке типа "aaaaa код: х хх" и "xxxx"
    If method = 2 Then
        s = Split(newCell, "код: ")
        If UBound(s) > 0 Then k = Replace(s(1), " ", "")
        Compare = (k = oldCell)
    End If
    
End Function

Private Function Persent(ByVal cur As Integer, ByVal all As Integer) As String
    Persent = Str(cur) + " из" + Str(all) + " (" + Str(Int(cur / all * 100)) + "% )"
End Function


