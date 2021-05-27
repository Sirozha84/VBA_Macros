Attribute VB_Name = "Starter"
'Last change: 19.04.2021 13:20

Type Adress
    UK As String    'Управляющая компания
    ul As String
    dom As String
    korp As String
    kv As String
    ls As String    'Номер лицевого счёта
    index As Long
    t1 As Long
    t2 As Long
End Type

Public Sub Start()
    FormReport.Show
End Sub

Public Sub SendEmails()
    FormEmails.Show
End Sub

Public Sub Instruction()
    MsgBox "Вставляем данные по счётчикам в порядке: ""Горячая вода"", ""Холодная вода"" и нажимаем " + _
    """Запуск"". По каждой управляющей компании строится отдельный отчёт на отдельной вкладке." + Chr(13) + _
    "Внимание! Таблица с управляющими компаниями должна быть отсортирована по колонке ""Наименование лицензиата"", " + _
    "и все поля адреса (Населённый пункт, поселение, улица, номер дома) должны быть заполнены.", _
    vbInformation, "Инстракция"
    'Может, в последствии сортировку и проверку данных сделаю в автомате...
End Sub

'******************** End of File ********************