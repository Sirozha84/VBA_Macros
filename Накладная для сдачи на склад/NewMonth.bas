Attribute VB_Name = "NewMonth"
'Новый месяц
'Автор: Сергей гордеев
'Дата изменения: 20.01.2017

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
    
    
    If MsgBox("Сделать сдвиг месяца?" + Chr(13) + "Внимание!!! При этом все существующие данные удалятся, " + _
        "а последние 5 дней сдвинутся влево.", vbYesNo) = 6 Then
        'Делаем сдвиг и очищаем новый месяц
    End If
    Exit Sub
error1:
    MsgBox ("Ошибка. Проверьте введёное значение.")
End Sub
