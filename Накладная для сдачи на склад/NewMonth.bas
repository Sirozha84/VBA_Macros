Attribute VB_Name = "NewMonth"
'Новый месяц
'Автор: Сергей гордеев
'Дата изменения: 31.01.2017

Sub NewMonth()
    'В случае покупки убрать
    If Not WorkIt Then Exit Sub
    
    Dim data As String
    data = InputBox("Введите дату (в формате месяц.год, например январь 2017 будет выглядеть как 1.17)." + Chr(13))
    If data = "" Then Exit Sub
    On Error GoTo error1
    dt = Split(data, ".")
    mnth = dt(0)
    yer = dt(1)
    'Делаем разметку дат
    
    On Error GoTo 0
    If MsgBox("Сделать сдвиг месяца?" + Chr(13) + "Внимание!!! При этом все существующие данные удалятся, " + _
        "а последние 5 дней сдвинутся влево.", vbYesNo) = 6 Then
        Application.ScreenUpdating = False
        'Делаем сдвиг
        For d = 27 To 31
            CopyPage Trim(Str(d)) + "д", "-" + Trim(Str(d)) + "д"
            CopyPage Trim(Str(d)) + "н", "-" + Trim(Str(d)) + "н"
        Next
        'Очищаем новый месяц
        For d = 1 To 31
            CLearPage Trim(Str(d)) + "д"
            CLearPage Trim(Str(d)) + "н"
        Next
        Application.ScreenUpdating = True
    End If
    'Заполняем даты
    FillDates mnth, yer
    
    Exit Sub
error1:
    MsgBox ("Ошибка. Проверьте введёное значение.")
End Sub

Sub FillDates(m, y)
    nn = 1
    'Dim dd As Date 'Правильней сделать так, но не знаю как :( VBA - гавно сраное
    For d = 1 To 31
        'dd = System.DateTime(1993, 5, 31, 12, 14, 0)
        'dd = 3 / 17 / 1984
        Sheets(Trim(Str(d)) + "д").Cells(1, 6) = "Накладная №" + Str(nn)
        Sheets(Trim(Str(d)) + "д").Cells(2, 6) = Trim(Str(d)) + "." + m + "." + y
        'Sheets(Trim(Str(d)) + "д").Cells(2, 6) = dd
        nn = nn + 1
        Sheets(Trim(Str(d)) + "н").Cells(1, 6) = "Накладная №" + Str(nn)
        Sheets(Trim(Str(d)) + "н").Cells(2, 6) = Trim(Str(d)) + "." + m + "." + y
        'Sheets(Trim(Str(d)) + "д").Cells(2, 6) = dd
        nn = nn + 1
    Next
End Sub

'Копирование страницы
Sub CopyPage(sorce, dist)
    'MsgBox """" + sorce + """ - """ + dist + """"
    For s = 6 To 25
        For C = 2 To 17
            Sheets(dist).Cells(s, C) = Sheets(sorce).Cells(s, C)
        Next
    Next
    Sheets(dist).Cells(1, 6) = Sheets(sorce).Cells(1, 6)
    Sheets(dist).Cells(2, 6) = Sheets(sorce).Cells(2, 6)
End Sub

'Очистка страницы
Sub CLearPage(page)
    For s = 6 To 25
        For C = 2 To 17
            Sheets(page).Cells(s, C) = ""
        Next
    Next
End Sub
