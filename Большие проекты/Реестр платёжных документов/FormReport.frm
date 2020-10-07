VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormReport 
   Caption         =   "Настройка отчёта"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5520
   OleObjectBlob   =   "FormReport.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "FormReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Загрузка программы
Private Sub UserForm_Activate()
    LabelVersion = "Версия: 2.2 (07.10.2020)"
    TextBoxHeat = Sheets(1).name
    TextBoxHW = Sheets(2).name
End Sub

'Поменять местами месяцы
Private Sub ButtonSwap_Click()
    temp = TextBoxHeat
    TextBoxHeat = TextBoxHW
    TextBoxHW = temp
End Sub

'Выход
Private Sub ButtonClose_Click()
    End
End Sub

'Запуск
Private Sub ButtonOK_Click()
    ReDim records(0) As Record
    Dim tabRes As String
    Dim tabHeat As String
    Dim tabHW As String
    Dim i As Long
    tabHeat = TextBoxHeat
    tabHW = TextBoxHW
    tabRes = TextBoxReport
    Mnth = TextBoxMonth
    Max = 0
    
    'Добавление данных "Тепловая энергия"
    Call Misc.Message("Добавление данных из ""Тепловая энергия""...")
    i = 2
    lastad = MakeAdress(tabHeat, i)
    pp = 0
    volume = 0
    Price = 0
    Tag = ""
    Do
        ad = MakeAdress(tabHeat, i)
        'Текущая запись отличается от предыдущей или вовсе отсутствует
        If StrComp(ad, lastad) <> 0 Or Sheets(tabHeat).Cells(i, 1) = "" Then
            Max = Max + 1
            ReDim Preserve records(Max) As Record
            records(Max).Adress = lastad
            records(Max).PPHeat = pp
            records(Max).VolumeHeat = volume
            records(Max).PriceHeat = Price
            records(Max).Tag = Tag
            'Сбрасываем счётчики
            pp = 0
            volume = 0
            Price = 0
            lastad = ad
        End If
        'Адрес как в предыдущей строке, "схлапываем" данные
        If StrComp(ad, lastad) = 0 Then
            pp = pp + 1
            volume = volume + Sheets(tabHeat).Cells(i, 8)
            Price = Price + Sheets(tabHeat).Cells(i, 9)
            Tag = Sheets(tabHeat).Cells(i, 11)
        End If
        If Sheets(tabHeat).Cells(i, 1) = "" Then Exit Do
        lastad = ad
        i = i + 1
    Loop
    MaxG = Max
    
    'Добавление данных "Горячая вода"
    Call Misc.Message("Добавление данных из ""Горячая вода""...")
    i = 2
    lastad = MakeAdress(tabHW, i)
    pp = 0
    volume = 0
    Price = 0
    Tag = ""
    Do
        ad = MakeAdress(tabHW, i)
        'Текущая запись отличается от предыдущей или вовсе отсутствует
        If StrComp(ad, lastad, 1) <> 0 Or Sheets(tabHW).Cells(i, 1) = "" Then
            'Данные схлопнули, ищем теперь что получилось в "Тепловой энергии"
            Find = 0
            For j = 1 To MaxG
                If StrComp(lastad, records(j).Adress, 1) = 0 Then
                    Find = j
                    Exit For
                End If
            Next
            If Find > 0 Then
                'Запись есть, дополняем её
                If records(Find).VolumeHW <> 0 Then
                    MsgBox "Вкладка """ + tabHW + """, Строка " + CStr(i) + Chr(13) + "Адрес """ + lastad + """ был ранее, но предыдущий отличается. Адреса должны располагаться по порядку!"
                    End
                End If
                records(Find).PPHW = pp
                records(Find).VolumeHW = volume
                records(Find).PriceHW = Price
            Else
                'Если такой записи ещё нет, добавляем новую
                Max = Max + 1
                ReDim Preserve records(Max) As Record
                records(Max).Adress = lastad
                records(Max).PPHW = pp
                records(Max).VolumeHW = volume
                records(Max).PriceHW = Price
                records(Max).Tag = Tag
            End If
            'Сбрасываем счётчики
            pp = 0
            volume = 0
            Price = 0
            lastad = ad
        End If
        'Адрес как в предыдущей строке, "схлапываем" данные
        If StrComp(ad, lastad, 1) = 0 Then
            pp = pp + 1
            volume = volume + Sheets(tabHW).Cells(i, 8)
            Price = Price + Sheets(tabHW).Cells(i, 9)
            Tag = Sheets(tabHW).Cells(i, 11)
        End If
        If Sheets(tabHW).Cells(i, 1) = "" Then Exit Do
        lastad = ad
        i = i + 1
    Loop
    
    'Создаём/очищаем результирующую страницу
    Call Misc.Message("Формирование отчёта...")
    Call Misc.NewTab(tabRes, True)
    Sheets(tabRes).Select
    
    'Устанавливаем ширину колонок и форматы
    Columns(1).ColumnWidth = 5.86
    Columns(2).ColumnWidth = 46.71
    Columns(3).ColumnWidth = 18
    Columns(4).ColumnWidth = 16
    Columns(5).ColumnWidth = 17.56
    Columns(6).ColumnWidth = 11.71
    Columns(7).ColumnWidth = 16.71
    Columns(8).ColumnWidth = 13.28
    Columns(9).ColumnWidth = 13.29
    Columns(4).NumberFormat = "### ### ##0.00"
    Columns(6).NumberFormat = "### ### ##0.00"
    Columns(7).NumberFormat = "### ### ##0.00"
    Columns(8).NumberFormat = "### ### ##0.00"
 
    'Заголовок
    Rows(1).RowHeight = 56.25
    Call MergeAndCenter(1, 1, 1, 9, "Реестр платежных документов для внесения платы за коммунальные услуги, " + _
        "предъявленной собственникам и пользователям помещений в многоквартирных домах или жилых домах, " + _
        "проживающих на территории края при наличии прямых договоров с организациями осуществляющими горячее " + _
        "водоснабжение, и теплоснабжающими организациями, поставляющими тепловую энергию, вырабатываемую " + _
        "электрокотельными (электробойлерными) или котельными, работающими на мазутном топливе " + _
        "(далее – ресурсоснабжающие организации)")
    Call MergeAndCenter(2, 1, 1, 9, "за " + Mnth + " года")
    Call MergeAndCenter(3, 1, 1, 9, "(наименование месяца)")
    Call MergeAndCenter(4, 1, 1, 9, "ООО ""КРАСЭКО-ЭЛЕКТРО""")
    Call MergeAndCenter(5, 1, 1, 9, "(наименование ресурсоснабжающей организации)")
    
    'Шапка
    Call MergeAndCenter(7, 1, 2, 1, "№ п/п")
    Call MergeAndCenter(7, 2, 2, 1, "Адрес многоквартирного или  жилого дома")
    Call MergeAndCenter(7, 3, 2, 1, "Наименование коммунального ресурса (тепловая энергия, горячая вода)")
    Call MergeAndCenter(7, 4, 2, 1, "Объем потребления коммунального ресурса по платежным документам")
    Call MergeAndCenter(7, 5, 2, 1, "Количество  платежных документов для внесения платы за  коммунальные услуги (платежные документы)")
    Call MergeAndCenter(7, 6, 1, 3, "Сумма по платежным документам" + Chr(10) + "(тыс. рублей)")
    Rows(8).RowHeight = 90
    Call MergeAndCenter(8, 6, 1, 1, "за отопление")
    Call MergeAndCenter(8, 7, 1, 1, "за компонент «тепловая энергия» при оказании услуги по горячему водоснабжению")
    Call MergeAndCenter(8, 8, 1, 1, "итого")
    Call MergeAndCenter(7, 9, 2, 1, "Период, за который предъявлены платежные документы")
    For i = 1 To 9
        Cells(9, i) = i
        Cells(9, i).HorizontalAlignment = xlCenter
    Next
    Range(Cells(7, 1), Cells(9, 9)).Borders.Weight = 3
    
    'Выводим таблицы
    Dim s As Integer
    s = 10
    allVolumeHeat = 0
    allVolumeHW = 0
    allDocsHeat = 0
    allDocsHW = 0
    allPriceHeat = 0
    allPriceHW = 0
    For t = 1 To 2
        num = 1
        sumVolumeHeat = 0
        sumVolumeHW = 0
        sumDocsHeat = 0
        sumDocsHW = 0
        sumPriceHeat = 0
        sumPriceHW = 0
        
        'Шапка
        Call MergeAndCenter(s, 1, 1, 1, CStr(t))
        If t = 1 Then
            Cells(s, 2) = "Многоквартирные дома"
        Else
            Cells(s, 2) = "Жилые дома"
        End If
        Sheets(tabRes).Range(Cells(s, 2), Cells(s, 9)).Merge
        Range(Cells(s, 1), Cells(s, 9)).Borders.Weight = 3
        s = s + 1
        
        For i = 1 To Max
            If ((t = 1 And StrComp(records(i).Tag, "мкд", 1) = 0) Or _
                (t = 2 And StrComp(records(i).Tag, "ижд", 1) = 0)) And _
                (records(i).VolumeHeat <> 0 Or records(i).VolumeHW <> 0) Then
                
                'тепловая энергия
                Call MergeAndCenter(s, 1, 3, 1, "'" + CStr(t) + "." + CStr(num))
                Cells(s, 2) = records(i).Adress
                Cells(s, 3) = "тепловая энергия"
                Cells(s, 4) = records(i).VolumeHeat
                Cells(s, 5) = records(i).PPHeat
                Cells(s, 6) = Round(records(i).PriceHeat / 1000, 3)
                Cells(s, 8) = Round(records(i).PriceHeat / 1000, 3)
                Cells(s, 9) = Mnth
                VolumeSum = records(i).VolumeHeat
                PriceSum = records(i).PriceHeat
                Range(Cells(s, 1), Cells(s, 9)).Borders.Weight = 2
                s = s + 1
                
                'горячая вода
                Cells(s, 2) = records(i).Adress
                Cells(s, 3) = "горячая вода"
                Cells(s, 4) = records(i).VolumeHW
                Cells(s, 5) = records(i).PPHW
                Cells(s, 7) = Round(records(i).PriceHW / 1000, 3)
                Cells(s, 8) = Round(records(i).PriceHW / 1000, 3)
                Cells(s, 9) = Mnth
                VolumeSum = VolumeSum + records(i).VolumeHW
                PriceSum = PriceSum + records(i).PriceHW
                Range(Cells(s, 1), Cells(s, 9)).Borders.Weight = 2
                s = s + 1
                
                'итого
                Cells(s, 2) = records(i).Adress
                Cells(s, 3) = "итого"
                Cells(s, 4) = VolumeSum
                Cells(s, 8) = Round(PriceSum / 1000, 3)
                Cells(s, 9) = Mnth
                Range(Cells(s, 2), Cells(s, 9)).Interior.Color = RGB(221, 235, 247)
                Range(Cells(s, 1), Cells(s, 9)).Borders.Weight = 2
                Range(Cells(s - 2, 1), Cells(s, 1)).Borders.Weight = 3
                Range(Cells(s - 2, 9), Cells(s, 9)).Borders(xlEdgeRight).Weight = 3
                s = s + 1
                
                'Счётчики
                sumVolumeHeat = sumVolumeHeat + records(i).VolumeHeat
                sumVolumeHW = sumVolumeHW + records(i).VolumeHW
                sumDocsHeat = sumDocsHeat + records(i).PPHeat
                sumDocsHW = sumDocsHW + records(i).PPHW
                sumPriceHeat = sumPriceHeat + records(i).PriceHeat
                sumPriceHW = sumPriceHW + records(i).PriceHW
                num = num + 1
                
            End If
        Next
        
        'Итоги
        If t = 1 Then
            Sheets(tabRes).Cells(s, 2) = "По группе пногоквартирные дома"
        Else
            Sheets(tabRes).Cells(s, 2) = "По группе жилые дома"
        End If
        
        'тепловая энергия
        Range(Cells(s, 2), Cells(s + 2, 2)).Merge
        Range(Cells(s, 2), Cells(s + 2, 2)).VerticalAlignment = xlCenter
        Cells(s, 3) = "тепловая энергия"
        Cells(s, 4) = sumVolumeHeat
        Cells(s, 5) = sumDocsHeat
        Cells(s, 6) = Round(sumPriceHeat / 1000, 3)
        Cells(s, 8) = Round(sumPriceHeat / 1000, 3)
        Range(Cells(s, 1), Cells(s, 9)).Borders.Weight = 2
        s = s + 1
        
        'горячая вода
        Cells(s, 3) = "горячая вода"
        Cells(s, 4) = sumVolumeHW
        Cells(s, 5) = sumDocsHW
        Cells(s, 7) = Round(sumPriceHW / 1000, 3)
        Cells(s, 8) = Round(sumPriceHW / 1000, 3)
        Range(Cells(s, 1), Cells(s, 9)).Borders.Weight = 2
        s = s + 1
        
        'итого
        Cells(s, 3) = "итого"
        Cells(s, 4) = sumVolumeHeat + sumVolumeHW
        Cells(s, 8) = Round((sumPriceHeat + sumPriceHW) / 1000, 3)
        Range(Cells(s - 2, 1), Cells(s, 1)).Merge
        Range(Cells(s, 1), Cells(s, 9)).Borders.Weight = 2
        Range(Cells(s - 2, 1), Cells(s, 1)).Borders.Weight = 3
        Range(Cells(s - 2, 2), Cells(s, 2)).Borders.Weight = 3
        Range(Cells(s - 2, 3), Cells(s - 2, 9)).Borders(xlEdgeTop).Weight = 3
        Range(Cells(s - 2, 9), Cells(s, 9)).Borders(xlEdgeRight).Weight = 3
        s = s + 1
        
        'Счётчики
        allVolumeHeat = allVolumeHeat + sumVolumeHeat
        allVolumeHW = allVolumeHW + sumVolumeHW
        allDocsHeat = allDocsHeat + sumDocsHeat
        allDocsHW = allDocsHW + sumDocsHW
        allPriceHeat = allPriceHeat + sumPriceHeat
        allPriceHW = allPriceHW + sumPriceHW
    Next
    
    'Самые итоговые итоги
    Sheets(tabRes).Cells(s, 1) = "По ресурсоснабжающей организации"
    Range(Cells(s, 1), Cells(s + 2, 2)).Merge
    Range(Cells(s, 1), Cells(s + 2, 2)).VerticalAlignment = xlCenter
    
    'тепловая энергия
    Cells(s, 3) = "тепловая энергия"
    Cells(s, 4) = allVolumeHeat
    Cells(s, 5) = allDocsHeat
    Cells(s, 6) = Round(allPriceHeat / 1000, 3)
    Cells(s, 8) = Round(allPriceHeat / 1000, 3)
    Range(Cells(s, 1), Cells(s, 9)).Borders.Weight = 2
    s = s + 1
    
    'горячая вода
    Cells(s, 3) = "горячая вода"
    Cells(s, 4) = allVolumeHW
    Cells(s, 5) = allDocsHW
    Cells(s, 7) = Round(allPriceHW / 1000, 3)
    Cells(s, 8) = Round(allPriceHW / 1000, 3)
    Range(Cells(s, 1), Cells(s, 9)).Borders.Weight = 2
    s = s + 1
    
    'итого
    Cells(s, 3) = "итого"
    Cells(s, 4) = allVolumeHeat + allVolumeHW
    Cells(s, 8) = Round((allPriceHeat + allPriceHW) / 1000, 3)
    Range(Cells(s, 1), Cells(s, 9)).Borders.Weight = 2
    Range(Cells(s - 2, 1), Cells(s, 2)).Borders.Weight = 3
    Range(Cells(s - 2, 3), Cells(s - 2, 9)).Borders(xlEdgeTop).Weight = 3
    Range(Cells(s - 2, 9), Cells(s, 9)).Borders(xlEdgeRight).Weight = 3
    Range(Cells(s, 3), Cells(s, 9)).Borders(xlEdgeBottom).Weight = 3
    
    s = s + 4
    
    'Подписи
    Cells(s, 1) = "Исполнительный директор"
    Cells(s, 8) = "И.П. Михайленко"
    Cells(s, 8).HorizontalAlignment = xlCenter
    s = s + 1
    Call MergeAndCenter(s, 4, 1, 2, "(подпись)")
    Range(Cells(s, 4), Cells(s, 5)).Borders(xlEdgeTop).LineStyle = True
    Cells(s, 8) = "(ФИО)"
    Cells(s, 8).HorizontalAlignment = xlCenter
    Cells(s, 8).Borders(xlEdgeTop).LineStyle = True
    s = s + 1
    Cells(s, 3) = "М. П."
    s = s + 1
    Cells(s, 3) = "(в случае наличия)"
    s = s + 3
    Cells(s, 1) = "Дорошенко Н.Н. (3919)757719"
    
    Call Misc.Message("Готово!")
    End
    
    Application.ScreenUpdating = True
End Sub

Sub MergeAndCenter(R As Integer, C As Integer, height As Integer, width As Integer, text As String)
    Cells(R, C).HorizontalAlignment = xlCenter
    Cells(R, C).VerticalAlignment = xlCenter
    Range(Cells(R, C), Cells(R + height - 1, C + width - 1)).Merge
    Cells(R, C).WrapText = True
    Cells(R, C) = text
End Sub

Function MakeAdress(sheet As String, i As Long) As String
On Error GoTo er:
    MakeAdress = Sheets(sheet).Cells(i, 1) + ", " + _
                 Sheets(sheet).Cells(i, 2) + ", " + _
                 CStr(Sheets(sheet).Cells(i, 3)) + _
                 Sheets(sheet).Cells(i, 4)
    If Sheets(sheet).Cells(i, 4) <> "" Then
        MakeAdress = MakeAdress + ", Корпус" + CStr(Sheets(sheet).Cells(i, 5))
    End If
    Exit Function
er:
    MsgBox "Ошибка во входных данных, проверьте правильность расположения колонок!"
    End
End Function