Attribute VB_Name = "Find"
Sub Find()
    
    rang = 200000   'Максимальная строчка в исходнике
    src = "10.2019" 'Страница с исходником
    res = "Result"  'Страница с результатом
    c_adr = 8       'Поле с адресом
    c_usl = 17      'Поле с услугой
    c_vip = 22      'Выпадающий доход
    c_fld = 24      'Количество полей в таблице
    
    ReDim tmp(rang) As String
    
    Sheets(res).Cells.Clear
    
    'Делаем фильтрованный временный список
    Sheets(res).Cells(1, 1) = "Подготовка..."
    rec = 1
    Max = 1
    For i = 2 To rang
        If Sheets(src).Cells(i, 1) = "" Then
            rec = i
            Exit For
        End If
        If Sheets(src).Cells(i, c_usl) = "отопление" And Sheets(src).Cells(i, c_vip) <> 0 Then
            tmp(Max) = Sheets(src).Cells(i, c_adr)
            Max = Max + 1
        End If
    Next
   
    
    'Ищем Данные, которые подходят под фильтр
    Sheets(res).Cells(1, 1) = "Подготовка..."
    Application.ScreenUpdating = False
    f = 1
    For i = 2 To rang
        If Sheets(src).Cells(i, 1) = "" Then Exit For
        If (i Mod 100) = 0 Then
            Application.ScreenUpdating = True
            Sheets(res).Cells(1, 1) = "Обработано:" + Str(i) + " из" + Str(rec) + _
                " (" + Str(i / rec * 100) + " % )    Найдено:" + Str(f)
            Application.ScreenUpdating = False
        End If
        
        'fnd = False 'IsInArray(Sheets(src).Cells(i, 8), tmp)
        fnd = False
        adr = Sheets(src).Cells(i, 8)
        If adr = last Then
            fnd = True
        Else
            For j = 1 To Max
                If adr = tmp(j) Then
                    fnd = True
                    last = adr
                    Exit For
                End If
            Next
        End If
        If fnd Then
            For c = 1 To c_fld
                Sheets(res).Cells(f + 1, c) = Sheets(src).Cells(i, c)
            Next
            f = f + 1
        End If
    Next
    
    'Копируем шапку
    For i = 1 To c_fld
        Sheets(res).Cells(1, i) = Sheets(src).Cells(1, i)
    Next
    
End Sub

