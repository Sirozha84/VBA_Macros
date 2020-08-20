Attribute VB_Name = "CalcCompare"
'Версия 1.0 (19.12.2019)
'Версия 1.2 (20.12.2019)
'Версия 1.3 (23.12.2019) - Выделение кода из строки
'Версия 1.4 (24.12.2019) - Оптимизация сообщений
'Версия 1.5 (24.12.2019) - Поиск изменений
'Версия 1.6 (26.12.2019) - Рефакторинг
'Версия 1.7 (10.01.2020) - Вычисление разницы
'Версия 1.8 (17.01.2020) - Правка итоговой таблицы
'Версия 1.9 (27.01.2020) - Подготовка таблиц и создание результирующей таблицы программно
'Версия 1.10 (04.02.2020) - Исправлен счётчик совпавших
'Версия 1.11 (07.02.2020) - Убирание лишних пробелов
'Версия 2.0 (21.07.2020) - Упрощение код после форка, теперь решаем только одну задачу
'Версия 2.1 (22.07.2020) - Исправлен небольшой "упс" с именем таблицы и сделан трим номеров (внезапно появились пробелы)
'Версия 2.2 (23.07.2020) - Ещё одна страница для сравнения "УК"
'Версия 2.3 (20.08.2020) - Значение люфта вынесено как константа

Const Luft = 5              'Люфт, максимальное число которого будет считатся как "почти совпало"
Const fCom = 7              'Поле для комментария
Global Tab1C As String      'Первая таблица
Global TabAccess As String  'Вторая таблица
Global TabResult As String  'Таблица с результатом
Global max As Integer       'Счётчик строк всего


Sub Start()
    
    Tab1C = "TDSheet"
    
    TabAccess = "Тепло"
    TabResult = "Тепло R"
    MakeCopy
    Compare
    
    TabAccess = "Вода"
    TabResult = "Вода R"
    MakeCopy
    Compare
    
    TabAccess = "УК"
    TabResult = "УК R"
    MakeCopy
    Compare
    
    Message "Готово!"
    
End Sub

'Подготовка итоговой таблицы
Private Sub MakeCopy()
    
    Message TabAccess + ": Копирование..."
    
    If Not SheetExist(TabResult) Then
        Sheets.Add(Sheets(Sheets.Count)).name = TabResult
    End If
    
    Sheets(TabResult).Select
    Cells.Clear
    maxStrings = 0
    
    'Копируем данные из Access
    max = 1
    Do While Sheets(TabAccess).Cells(max + 1, 1) <> ""
        max = max + 1
        For i = 1 To 3
            Sheets(TabResult).Cells(max, i) = Sheets(TabAccess).Cells(max, i)
        Next
    Loop
    
    'Дорисовываем шапку
    Cells(1, 1) = "Договор"
    Cells(1, 2) = "Access"
    Cells(1, 3) = "Сумма"
    Cells(1, 4) = "УТВК (1С)"
    Cells(1, 5) = "Сумма"
    Cells(1, 6) = "Разница"
    Cells(1, 7) = "Итог"
    Columns(2).ColumnWidth = 10
    Columns(2).ColumnWidth = 30
    Columns(4).ColumnWidth = 30
    Columns(7).ColumnWidth = 18
    
End Sub

'Сравнение
Private Sub Compare()
    
    Message TabAccess + ": Сравнение с УТВК..."
    Sheets(TabResult).Select
    
    i = 2
    Dim Mached As Long
    Dim Pochti As Long
    Dim Changed As Long
    
    Do While Sheets(Tab1C).Cells(i, 1) <> ""
        
        Find = False
        For j = 2 To max
            Find = False
            If Trim(Cells(j, 1)) = Trim(Sheets(Tab1C).Cells(i, 1)) Then
                Find = True
                Exit For
            End If
        Next
        If Find Then
            'Если нашли, тогда сравниваем
            Cells(j, 4) = Sheets(Tab1C).Cells(i, 2)
            Cells(j, 5) = Sheets(Tab1C).Cells(i, 3)
            rz = Round(Sheets(TabResult).Cells(j, 5) - Cells(j, 3), 2)
            Cells(j, 6) = rz
            If Abs(rz) = 0 Then
                Cells(j, 3).Interior.Color = RGB(196, 255, 196)
                Cells(j, 5).Interior.Color = RGB(196, 255, 196)
                Cells(j, 6).Interior.Color = RGB(128, 255, 128)
                Cells(j, fCom) = "Совпал"
                Mached = Mached + 1
            End If
            If Abs(rz) > 0 And Abs(rz) <= Luft Then
                Cells(j, 3).Interior.Color = RGB(255, 255, 196)
                Cells(j, 5).Interior.Color = RGB(255, 255, 196)
                Cells(j, 6).Interior.Color = RGB(255, 255, 128)
                Cells(j, fCom) = "Почти"
                Pochti = Pochti + 1
            End If
            If Abs(rz) > Luft Then
                Cells(j, 3).Interior.Color = RGB(255, 196, 196)
                Cells(j, 5).Interior.Color = RGB(255, 196, 196)
                Cells(j, 6).Interior.Color = RGB(255, 128, 128)
                Changed = Changed + 1
                Cells(j, fCom) = "Не совпал"
            End If
            'sumDif = sumDif + rz
        End If
        i = i + 1
    Loop
    
    Message TabAccess + ": Завершение..."
    
    'Поиск не попавших
    Dim NonFind As Long
    For i = 2 To max
        If Cells(i, 6) = "" Then
            Cells(i, fCom) = "Не найдено"
            NonFind = NonFind + 1
        End If
    Next
    
    'Итоги
    Dim all As Long
    all = Mached + Pochti + Changed + NonFind
    Cells(max + 2, fCom) = "Всего: " + CStr(all)
    Cells(max + 3, fCom) = "Совпало: " + CStr(Mached) + " + " + CStr(Pochti) + Prc(Mached + Pochti, all)
    Cells(max + 4, fCom) = "Не совпало: " + CStr(Changed) + Prc(Changed, all)
    Cells(max + 5, fCom) = "Не найдено: " + CStr(NonFind) + Prc(NonFind, all)
    
End Sub

'Подсчёт процентов
Private Function Prc(val As Long, all As Long) As String
    Prc = " (" + CStr(Round(val / all * 100, 1)) + "%)"
End Function

'Вывод сообщение в статусбар
Private Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    Application.ScreenUpdating = False
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
