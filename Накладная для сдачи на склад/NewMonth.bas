Attribute VB_Name = "NewMonth"
Sub NewMonth()
    Dim data As String
    data = InputBox("Введите дату (в формате месяц.год, например январь 2017 будет выглядеть как 1.17)." + Chr(13) + _
    "Внимание!!! Данные файла удалятся, останутся только последние 5 дней, которые сдвинутся на месяц влево.")
    If data = "" Then Exit Sub
    On Error GoTo error1
    dt = Split(data, ".")
    mnth = dt(0)
    yer = dt(1)
    'Делаем разметку дат
    
    
    If MsgBox("Сделать сдвиг месяца?", vbYesNo, "Что-то") = 6 Then
        'Делаем сдвиг и очищаем новый месяц
    End If
    Exit Sub
error1:
    MsgBox ("Ошибка. Проверьте введёное значение.")
End Sub
