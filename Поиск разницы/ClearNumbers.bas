Attribute VB_Name = "ClearNumbers"
'Чистка номеров договоров от надписей

Sub Start()
    
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
        Cells(i, 1) = Replace(Cells(i, 1), "муниипальный ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "муниципальны ", "")
        Cells(i, 1) = Replace(Cells(i, 1), "клнтракт ВК № ", "")

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


