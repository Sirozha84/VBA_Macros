Attribute VB_Name = "Starter"
'Last change: 01.06.2021 09:17

Type Adress
    UK As String    'Управляющая компания
    ul As String
    dom As String
    korp As String
    kv As String
    potr As String  'Потребитель
    prop As Long    'Прописано
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
    MsgBox "Вставляем данные по счётчикам в порядке: ""Горячая вода"", ""Холодная вода"" и нажимаем ""Запуск""." + Chr(13) + _
    "По каждой управляющей компании строится отдельный отчёт на отдельной вкладке.", vbInformation
End Sub

'******************** End of File ********************