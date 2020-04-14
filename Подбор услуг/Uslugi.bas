Attribute VB_Name = "Uslugi"
'Выборка услуг и схлопывание строк

Const Col = 3
Const Usl = 4

Sub Uslugi1()
    Message "Начало..."
    Max = 194311
    sh = 0
    last = 0
    For i = 2 To Max
        
        sh = sh + 1
        If sh > 100 Then sh = 0: Call ProgressBar("Обработка", i, Max)
        
        If last = Cells(i, Col) Then
            Cells(i - 1, Usl) = Cells(i, Usl)
            
            Rows(i).EntireRow.Delete
            i = i - 1
            Max = Max - 1
        Else
            last = Cells(i, Col)
        End If
        
        If Cells(i, Usl) = "ХВС" Then Cells(i, 5) = "+"
        If Cells(i, Usl) = "ГВС ТН" Then Cells(i, 6) = "+"
        If Cells(i, Usl) = "ВО" Then Cells(i, 7) = "+"
        If Cells(i, Usl) = "Отопление" Then Cells(i, 8) = "+"
        If Cells(i, 1) = "" Then Exit For
        
    Next
    Message "Готово!"
End Sub

Sub Uslugi2()

    ALlMoney = 12 'Колонка со всеми деньгами
    Money = 13    'Первая колонка с деньгами
    Volumes = 17  'Первая колонка с объёмами
    
    Message "Начало..."
    Max = 3938
    sh = 0
    last = 0
    For i = 2 To Max
        sh = sh + 1
        If sh > 100 Then sh = 0: Call ProgressBar("Обработка", i, Max)
        
        'Раздвигашки выставленного
        
        'ТН-ТЭ
        If Cells(i, Volumes) <> "" Then Cells(i, Money) = Cells(i, ALlMoney)
        'Отопление
        If Cells(i, Volumes + 2) <> "" Then Cells(i, Money + 2) = Cells(i, ALlMoney)
        'ХВС
        If Cells(i, Volumes + 3) <> "" Then Cells(i, Money + 3) = Cells(i, ALlMoney)
        
        If last = Cells(i, 1) Then
            
            'Схлопывашки
            
            'ТН-ТЭ
            If Cells(i, Volumes) <> "" Then
                If Cells(i - 1, Volumes) = "" Then
                    'сверху пусто, делаем как обычно
                    Cells(i - 1, Money) = Cells(i, ALlMoney)
                    Cells(i - 1, Volumes) = Cells(i, Volumes)
                Else
                    'Сверху не пусто...
                    If Cells(i - 1, Volumes) > Cells(i, Volumes) Then
                        Cells(i - 1, Money + 1) = Cells(i, ALlMoney)
                        Cells(i - 1, Volumes + 1) = Cells(i, Volumes)
                    Else
                    
                        Cells(i - 1, Money) = Cells(i, ALlMoney)
                        Cells(i - 1, Volumes + 1) = Cells(i - 1, Volumes)
                        Cells(i - 1, Volumes) = Cells(i, Volumes)
                    End If
                End If
            End If
            'Отопление
            If Cells(i, Volumes + 2) <> "" Then
                Cells(i - 1, Money + 2) = Cells(i, ALlMoney)
                Cells(i - 1, Volumes + 2) = Cells(i, Volumes + 2)
            End If
            'ХВС
            If Cells(i, Volumes + 3) <> "" Then
                Cells(i - 1, Money + 3) = Cells(i, ALlMoney)
                Cells(i - 1, Volumes + 3) = Cells(i, Volumes + 3)
            End If
            
            Rows(i).EntireRow.Delete
            i = i - 1
            Max = Max - 1
        Else
            last = Cells(i, 1)
        End If
        
        If Cells(i, 1) = "" Then Exit For
        
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


Private Function Adress(i As Long) As String
    Adress = CStr(Cells(i, 1)) + CStr(Cells(i, 2)) + CStr(Cells(i, 3)) + CStr(Cells(i, 4))
    Adress = LCase(Adress)
    Adress = Replace(Adress, "ё", "е")
End Function

Private Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    DoEvents
    Application.ScreenUpdating = False
End Sub





