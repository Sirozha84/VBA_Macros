Attribute VB_Name = "Trial"
'Проверка на право использования программы
'Автор: Сергей гордеев
'Дата изменения: 20.01.2017

Function WorkIt() As Boolean
    Days = DateDiff("d", Now, "1.4.2017")
    If (Days <= 0) Then
        MsgBox ("Период триального использования программы истёк." + _
            Chr(13) + "Необходимо приобрести лицензию.")
        WorkIt = False
        Exit Function
    End If
    If (Days < 30) Then MsgBox ("Внимание!" + Chr(13) + "Период триального использования программы истекает." + _
        Chr(13) + "Осталось дней: " + Str(Days))
    WorkIt = True
End Function
