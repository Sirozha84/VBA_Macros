Attribute VB_Name = "Misc"
'Подготовка данных, выгруженных из Access
Sub PrepareTabFromAccess()
    
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

'Выделение кода из строки и помещение его в отдельную ячейку
Sub CodeEject()
    
    codeFrom = 1    'Поле где брать код
    codeTo = 5      'Поле куда помещать обработанный

    Message "Подсчёт строк..."
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

'Вывод статуса обработки
Private Sub ProgressBar(text As String, ByVal cur As Integer, ByVal all As Integer)
    If cur Mod 100 = 0 Then
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


