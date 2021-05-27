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
'Last change: 19.04.2021 13:26

'Колонки входных данных
Const cPotr = 1     'Потребитель
Const cNP = 2       'Населённый пункт
Const cUl = 3       'Улица
Const cDom = 4      'Дом
Const cKv = 7       'Квартира
Const cLs = 8       'Номер лицевого счёта
Const cPr = 10      'Прописано

Const v1 = 13       'Колонка объём потребления ИПУ
Const v2 = 14       'Колонка объём потребления Норматив
Const v3 = 15       'Колонка объём потребления Перерасчёт

Dim res() As Adress 'Таблица с результатами
Dim curDom As String 'Адрес текущего дома для поиска

'Загрузка программы
Private Sub UserForm_Activate()
    LabelVersion = "Версия: 1.2 (19.04.2021)"
    On Error GoTo er
    TextBoxTN = Sheets(1).name
    TextBoxHVS = Sheets(2).name
    TextBoxUK = Sheets("УК").name
    Exit Sub
er:
    MsgBox ("Не хватает данных")
    End
End Sub

'Выход
Private Sub ButtonClose_Click()
    End
End Sub

'Запуск
Private Sub ButtonOK_Click()
    
    ReDim houses(0) As Adress
    ReDim tn(0) As Adress
    ReDim HVS(0) As Adress
    Dim tab1 As Variant
    Dim tab2 As Variant
    Dim max1 As Long
    Dim max2 As Long
    Dim tabA As Variant
    Dim UK As String
    Dim LaskUK As String
    
    Set tab1 = Sheets(TextBoxTN.text)
    Set tab2 = Sheets(TextBoxHVS.text)
    Set tabA = Sheets(TextBoxUK.text)
    
    Call Misc.Message("Подготовка...")
    
    'Чтение адресов
    Mnth = TextBoxMonth
    Dim aMax As Long
    aMax = 0
    i = 2
    Do While tabA.Cells(i, 3) <> ""
        If LCase(tabA.Cells(i, 14).text) = "да" Then
            aMax = aMax + 1
            ReDim Preserve houses(aMax) As Adress
            houses(aMax).UK = Trim(tabA.Cells(i, 3))
            houses(aMax).ul = Trim(tabA.Cells(i, 8)) + " " + Trim(tabA.Cells(i, 9))
            houses(aMax).dom = LCase(Trim(tabA.Cells(i, 10)) + Trim(tabA.Cells(i, 11)))
            houses(aMax).korp = Trim(tabA.Cells(i, 13))
        End If
        i = i + 1
    Loop
    
    'Узнаём размер входных таблиц
    max1 = FindMax(tab1)
    max2 = FindMax(tab2)
    
    'Построение отчёта
    StartProcess
    lastUK = ""
    For a = 1 To aMax
        UK = houses(a).UK
        If lastUK <> UK Then
            lastUK = UK
            'Создаём/очищаем результирующую страницу
            Call Misc.NewTab(UK, True)
            Sheets(UK).Select
            r = 3
    
            'Шапка
            Cells(1, 1) = "Номер": Range(Cells(1, 1), Cells(2, 1)).Merge
            Cells(1, 2) = "Населённый пункт, улица": Range(Cells(1, 2), Cells(2, 2)).Merge
            Cells(1, 3) = "Дом": Range(Cells(1, 3), Cells(2, 3)).Merge
            Cells(1, 4) = "Корпус": Range(Cells(1, 4), Cells(2, 4)).Merge
            Cells(1, 5) = "Квартира": Range(Cells(1, 5), Cells(2, 5)).Merge
            For i = 6 To 11
                Cells(2, i).WrapText = True
            Next
            For i = 0 To 2
                Cells(2, 6 + i) = tab1.Cells(1, 14 + i)
                Cells(2, 9 + i) = tab2.Cells(1, 14 + i)
            Next
            Cells(1, 6) = tab1.name
            Range(Cells(1, 6), Cells(1, 8)).Merge
            Cells(1, 9) = tab2.name
            Range(Cells(1, 9), Cells(1, 11)).Merge
            Call Borders(1, 1, 2, 11)
            Call Borders(1, 6, 2, 8)
            
            Columns(1).ColumnWidth = 7
            Columns(2).ColumnWidth = 30
            For i = 3 To 5
                Columns(i).ColumnWidth = 10
            Next
            For i = 6 To 11
                Columns(i).ColumnWidth = 15
            Next
            nd = 1
        End If
    
        'Построение отчёта
        curDom = houses(a).ul + ", " + LCase(houses(a).dom)
        Message "Построение отчёта... " + CStr(a) + " из " + CStr(aMax) + _
                " (" + curDom + ")" + TimePredict(a, aMax)
        ReDim res(0) As Adress
        Call FindHome(tab1, max1, 1)
        Call FindHome(tab2, max2, 2)
        
        'Выводим в отчёт данные по текущему дому (если нашлась хоть одна квартира)
        firstr = r
        kvcount = UBound(res)
        If kvcount > 0 Then
            ReDim v(5) As Double 'Итоговые суммы
            For i = 1 To kvcount
                If i = 1 Then Cells(r, 1) = nd
                Cells(r, 2) = res(i).ul
                Cells(r, 3) = res(i).dom
                Cells(r, 4) = res(i).korp
                Cells(r, 5) = res(i).kv
                
                If res(i).t1 <> 0 Then
                    For j = 0 To 2
                        Cells(r, 6 + j) = tab1.Cells(res(i).t1, 14 + j)
                    Next
                End If
                If res(i).t2 <> 0 Then
                    For j = 0 To 2
                        Cells(r, 9 + j) = tab2.Cells(res(i).t2, 14 + j)
                    Next
                End If
                For j = 0 To 5
                    v(j) = v(j) + Cells(r, 6 + j)
                Next
                r = r + 1
            Next
            Cells(r, 2) = "Итого:"
            For i = 0 To 5
                Cells(r, 6 + i) = v(i)
            Next
            
            r = r + 1
            Call Borders(firstr, 1, r - 1, 11)
            Call Borders(firstr, 6, r - 1, 8)
            nd = nd + 1
        End If
    Next
    
    Call Misc.Message("Готово!")
    End
    
End Sub

'Поиск дома по таблице. n - имя, t - номер таблицы (в зависимости от этого в ту или другую часть будут собираться объёмы)
Sub FindHome(tb As Variant, max As Long, t As Integer)
    For i = 2 To max
    
        ul = Trim(tb.Cells(i, cNP)) + " " + Trim(tb.Cells(i, cUl))
        dom = CStr(tb.Cells(i, cDom)) + LCase(Trim(tb.Cells(i, cDom + 1)))
        kv = tb.Cells(i, cKv)
        ls = tb.Cells(i, cLs)
        
        If curDom = ul + ", " + dom Then
            
            'Ищем, есть ли адрес в результатах
            Find = 0
            For j = 1 To UBound(res)
                If res(j).ul = ul And res(j).dom = dom And res(j).kv = kv And res(j).ls = ls Then
                    Find = j
                    Exit For
                Else
                End If
            Next
            
            'Если нет - добавляем новую запись
            If Find = 0 Then
                Find = UBound(res) + 1
                ReDim Preserve res(Find)
                res(Find).ul = ul
                res(Find).dom = dom
            End If
            
            'Размещаем остальные данные в запись
            res(Find).kv = tb.Cells(i, cKv)
            res(Find).ls = tb.Cells(i, cLs)
            If t = 1 Then
                res(Find).t1 = i
            Else
                res(Find).t2 = i
            End If
            
        End If
    Next
End Sub

Sub Borders(ByVal x1 As Long, y1 As Long, x2 As Long, y2 As Long)
    Range(Cells(x1, y1), Cells(x2, y2)).Borders.Weight = 2
    Range(Cells(x1, y1), Cells(x2, y2)).Borders(xlEdgeBottom).Weight = 4
    Range(Cells(x1, y1), Cells(x2, y2)).Borders(xlEdgeLeft).Weight = 4
    Range(Cells(x1, y1), Cells(x2, y2)).Borders(xlEdgeRight).Weight = 4
    Range(Cells(x1, y1), Cells(x2, y2)).Borders(xlEdgeTop).Weight = 4
End Sub

'******************** End of File ********************