Attribute VB_Name = "MKD"
Const ResTab = "итог"

Sub Start()
    
    Message "Начало..."
    
    Call Podbor("янв", 4)
    Call Podbor("фев", 5)
    Call Podbor("март", 6)
    Call Podbor("апр", 7)
    Call Podbor("май", 8)
    Call Podbor("июн", 9)
    Call Podbor("июл", 10)
    Call Podbor("авг", 11)
    Call Podbor("сен", 12)
    Call Podbor("окт", 13)
    Call Podbor("ноя", 14)
    Call Podbor("дек", 15)
    
    Message "Готово"
    
End Sub

Private Sub Podbor(SrcTab As String, k As Integer)
        
    max = GetMax(SrcTab)
    maxx = GetMax(ResTab)
    For i = 2 To max
        'If Sheets(SrcTab).Cells(i, 1) = 41 Then
        '    MsgBox ("Нашёл")
        'End If
    
        If i Mod 100 = 0 Then Call ProgressBar("Обработка", i, max)
        s = Search(ResTab, Sheets(SrcTab).Cells(i, 1), maxx, False)
        If s = 0 Then maxx = maxx + 1: s = maxx
        
        Sheets(ResTab).Cells(s, 1) = Sheets(SrcTab).Cells(i, 1)
        Sheets(ResTab).Cells(s, 2) = Sheets(SrcTab).Cells(i, 2)
        Sheets(ResTab).Cells(s, 3) = Sheets(SrcTab).Cells(i, 3)
        If Sheets(ResTab).Cells(s, k) = "" Then Sheets(ResTab).Cells(s, k) = Sheets(SrcTab).Cells(i, 4)
        
        Sheets(ResTab).Cells(s, 16) = Sheets(SrcTab).Cells(i, 1)
        Sheets(ResTab).Cells(s, 17) = Sheets(SrcTab).Cells(i, 2)
        Sheets(ResTab).Cells(s, 18) = Sheets(SrcTab).Cells(i, 3)
        If Sheets(ResTab).Cells(s, k + 15) = "" Then Sheets(ResTab).Cells(s, k + 15) = Sheets(SrcTab).Cells(i, 5)
    Next
        
End Sub

'###################################
'########## Общие функции ##########
'###################################

'Поиск максимальной строки
'inTab - имя таблицы
Private Function GetMax(inTab As String)
    i = 0
    Do
        i = i + 1000
    Loop Until Sheets(inTab).Cells(i, 1) = ""
    Do
        i = i - 1
    Loop Until Sheets(inTab).Cells(i, 1) <> ""
    GetMax = i
End Function

'Поиск ячейки
'inTab - имя таблицы
'str - искомая строка
'max - максимальная строка, если неизвестно ставим 0
'sort - указывает на то что этот список сортированный (для быстрого поиска)
Private Function Search(inTab As String, str As String, max, sort As Boolean)
    Search = 0
    If max = 0 Then max = GetMax(inTab)
    If sort Then
        'Когда-нибудь потом :-)
    Else
        Find = False
        For i = 1 To max
            If Sheets(inTab).Cells(i, 1) = str Then
                Find = True
                Exit For
            End If
        Next
        If Find Then Search = i
    End If
End Function

'Рисование прогресса
'text - имя
'cur - текущее значение
'all - всего, отображать каждые over штук
Private Sub ProgressBar(text As String, ByVal cur As Long, ByVal all As Long)
    Application.ScreenUpdating = True
    Application.StatusBar = text + ":" + str(cur) + " из" + str(all) + _
        " (" + str(Int(cur / all * 100)) + "% )"
        DoEvents
    Application.ScreenUpdating = False
End Sub

'Рисование сообщения
'test - сообщение
Private Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    DoEvents
    Application.ScreenUpdating = False
End Sub



