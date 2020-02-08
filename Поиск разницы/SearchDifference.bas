Attribute VB_Name = "SearchDifference"
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

Const firstTab = "Access"   'Первая таблица
Const secondTab = "УФА"     'Вторая таблица
Const resultTab = "Result"  'Таблица с результатом
Const maxTabs = 3           'Максимум полей
Const machTabs = 1          'Количество полей для поиска строк
Const compareTab = 3        'Колонка для сравнения значения
Const calculateDif = True   'Считать разницу значений (для чисел true, для строк false)
Const fCom = 8              'Поле для комментария

Global sn As Integer        'Счётчик новых строк
Global max As Integer       'Счётчик строк всего


Sub Start()

    PrepareTabFromAccess (firstTab)
    PrepareTabFromAccess (secondTab)

    CreateSheet (resultTab)

    MakeCopy
    AddDead
    FindNew
    
    Message "Готово!"
    
End Sub

'Создание результирующей таблицы
Private Sub CreateSheet(name As String)
    If Not SheetExist(name) Then
        Sheets.Add(, Sheets(Sheets.Count)).name = name
    End If
    Sheets(name).Cells.Clear
End Sub

'Проверка на существование листа
Function SheetExist(name As String) As Boolean
    Dim objSheet As Object
    
    On Error GoTo HandleError
    ThisWorkbook.Worksheets(name).Activate
    SheetExist = True
    Exit Function
    
HandleError:
    SheetExist = False
End Function

'Подготовка данных, выгруженных из Access
Sub PrepareTabFromAccess(shit As String)
    
    Sheets(shit).Select
    
    Message "Подсчёт строк..."
    max = 1
    
    Do Until Cells(max + 1, 1) = ""
        max = max + 1
    Loop
    
    For i = 2 To max
        
        Call ProgressBar("Обработка", i, max)
        
        'Перевод в текст
        Cells(i, 1).NumberFormat = "@"
        
        'Убирание из номера договора лишнего текста
        Cells(i, 1) = Replace(Cells(i, 1), "государственный контракт ВК № ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "договор  ВК № ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "договор № ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "договор ВК №  ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "договор ВК № ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "договор КС № ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "договор КС №", "")
        Cells(i, 1) = Replace(Cells(i, 1), "контракт ВК №  ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "контракт ВК № ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "муниципальный контракт ВК № ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "муниципальный ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "мцниципальный ", "")
        
        'Убирание нолей
        s = Cells(i, 1)
        If s <> "" Then
            Do While Left(s, 1) = "0"
                s = Right(s, Len(s) - 1)
            Loop
            Cells(i, 1) = s
        End If
        
        'Убирание лишних пробелов
        Do While Right(Cells(i, 1), 1) = " "
            Cells(i, 1) = Left(Cells(i, 1), Len(Cells(i, 1)) - 1)
        Loop
        
        'Убирание "рублей"
        s = Replace(Cells(i, 3), "р.", "")
        If s <> "" Then Cells(i, 3) = CDbl(s)
        
        'Ссумирование НДС (если есть)
        If Cells(1, 4) = -1 Then
            Cells(i, 3) = Cells(i, 3) + Cells(i, 4)
            Cells(i, 4) = ""
        End If

    Next

    For i = 2 To max + 5

        'Схлопывание
        If Cells(i, 1) <> "" And Cells(i, 1) <> "." And Cells(i, 1) = Cells(i + 1, 1) Then
           Cells(i, 3) = Cells(i, 3) + Cells(i + 1, 3)
           Rows(i + 1).EntireRow.Delete
           i = i - 1
        End If
        
        'Чистка хвостов
        If Cells(i, 1) = "" Then Rows(i).EntireRow.Delete
        
    Next
    
    Message "Готово!"
    
End Sub

'Подготовка итоговой таблицы
Private Sub MakeCopy()
    
    Message "Подготовка..."
    
    sn = 0
    mx = 0
    
    Sheets(resultTab).Cells.Clear
    maxStrings = 0
    max = 1
    Do While Sheets(secondTab).Cells(max + 1, 1) <> ""
        max = max + 1
        For i = 1 To 3
            Sheets(resultTab).Cells(max, i) = Sheets(secondTab).Cells(max, i)
        Next
    Loop
    
    'Дорисовываем шапку
    Sheets(resultTab).Cells(1, 1) = "Договор"
    Sheets(resultTab).Cells(1, 3) = secondTab
    Sheets(resultTab).Cells(1, 5) = firstTab
    Sheets(resultTab).Cells(1, 6) = "Разница"
    
End Sub

'Поиск и добавления удалённых строк
Private Sub AddDead()
    
    'Находим максимум в старой таблице
    Message "Подсчёт строк..."
    maxOld = 0
    Do Until Sheets(firstTab).Cells(maxOld + 1, 1) = ""
        maxOld = maxOld + 1
    Loop
    
    mached = 0
    changed = 0
    
    Find = False
    For i = 2 To maxOld
        Call ProgressBar("Поиск удалённых", i, maxOld)
        For j = 2 To max
            Find = False
            If Sheets(resultTab).Cells(j, 1) = Sheets(firstTab).Cells(i, 1) Then
                Find = True
                Exit For
            End If
        Next
        If Find Then
            'Если нашли, тогда сравниваем
            Sheets(resultTab).Cells(j, 4) = Sheets(firstTab).Cells(i, 2)
            Sheets(resultTab).Cells(j, 5) = Sheets(firstTab).Cells(i, 3)
            rz = Round(Sheets(resultTab).Cells(j, 3) - Sheets(resultTab).Cells(j, 5), 2)
            Sheets(resultTab).Cells(j, 6) = rz
            If Abs(rz) = 0 Then
                Sheets(resultTab).Cells(j, 3).Interior.Color = RGB(128, 255, 128)
                Sheets(resultTab).Cells(j, 5).Interior.Color = RGB(128, 255, 128)
                Sheets(resultTab).Cells(j, 6).Interior.Color = RGB(196, 255, 196)
                Sheets(resultTab).Cells(j, fCom) = "Совпал"
            End If
            If Abs(rz) > 0 Then
                Sheets(resultTab).Cells(j, 3).Interior.Color = RGB(255, 255, 128)
                Sheets(resultTab).Cells(j, 5).Interior.Color = RGB(255, 255, 128)
                Sheets(resultTab).Cells(j, 6).Interior.Color = RGB(255, 255, 196)
                Sheets(resultTab).Cells(j, fCom) = "Совпал (почти)"
            End If
            If Abs(rz) > 10 Then
                Sheets(resultTab).Cells(j, 3).Interior.Color = RGB(255, 128, 128)
                Sheets(resultTab).Cells(j, 5).Interior.Color = RGB(255, 128, 128)
                Sheets(resultTab).Cells(j, 6).Interior.Color = RGB(255, 196, 196)
                changed = changed + 1
                Sheets(resultTab).Cells(j, fCom) = "Изменён"
            End If
            If Abs(rz) < 10 Then mached = mached + 1
            sumDif = sumDif + rz
        Else
            'А если не нашли, копируем
            max = max + 1
            Sheets(resultTab).Cells(max, 1) = Sheets(firstTab).Cells(i, 1)
            Sheets(resultTab).Cells(max, 4) = Sheets(firstTab).Cells(i, 2)
            Sheets(resultTab).Cells(max, 5) = Sheets(firstTab).Cells(i, 3)
            Sheets(resultTab).Cells(max, fCom) = "Есть в " + firstTab + ", но нет в " + secondTab
            
            'Рисование разницы
            'Sheets(resultTab).Cells(max, 3).Cells = 0
            Sheets(resultTab).Cells(max, 6).Cells = -Sheets(resultTab).Cells(max, 5).Cells
            Sheets(resultTab).Cells(max, 3).Interior.Color = RGB(255, 128, 128)
            Sheets(resultTab).Cells(max, 5).Interior.Color = RGB(255, 128, 128)
            Sheets(resultTab).Cells(max, 6).Interior.Color = RGB(255, 196, 196)
            
            sumDif = sumDif - Sheets(resultTab).Cells(max, 3).Cells
            
            sn = sn + 1
        End If
    Next
    
    Sheets(resultTab).Cells(max + 3, maxTabs + 5) = "Есть только в " + firstTab + ":" + Str(sn)
    Sheets(resultTab).Cells(max + 4, maxTabs + 5) = "Совпало:" + Str(mached)
    Sheets(resultTab).Cells(max + 5, maxTabs + 5) = "Изменено:" + Str(changed)
    
End Sub

'Поиск новых строк
Private Sub FindNew()
    
    'Находим максимум в старой таблице
    Message "Подсчёт строк..."
    maxOld = 0
    Do Until Sheets(firstTab).Cells(maxOld + 1, 1) = ""
        maxOld = maxOld + 1
    Loop
    
    s = 0
    For i = 2 To max - sn
        
        Call ProgressBar("Поиск новых", i, max - sn - 1)
        
        Find = False
        For j = 2 To maxOld
            If Sheets(resultTab).Cells(i, 1) = Sheets(firstTab).Cells(j, 1) Then
                Find = True
                Exit For
            End If
        Next
        If Not Find Then
            
            'Рисование "разницы"
            Sheets(resultTab).Cells(i, 6).Cells = Sheets(resultTab).Cells(i, maxTabs).Cells
            Sheets(resultTab).Cells(i, 3).Interior.Color = RGB(255, 128, 128)
            Sheets(resultTab).Cells(i, 5).Interior.Color = RGB(255, 128, 128)
            Sheets(resultTab).Cells(i, 6).Interior.Color = RGB(255, 196, 196)
            
            Sheets(resultTab).Cells(i, fCom) = "Есть в " + secondTab + ", но нет в " + firstTab
            sumDif = sumDif + Sheets(resultTab).Cells(i, maxTabs).Cells
            
            s = s + 1
            
        End If
    Next
    
    Sheets(resultTab).Cells(max + 2, maxTabs + 5) = "Есть только в " + secondTab + ":" + Str(s)
    
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


