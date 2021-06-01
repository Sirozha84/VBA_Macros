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
'Last change: 01.06.2021 09:17

Const cPotr = 1     'Колонка потребителя
Const cUl = 4       'Колонка улицы
Const cDom = 5      'Колонка дом
Const cKv = 9       'Колонка квартира
Const cPr = 10      'Колонка прописано
Const v1 = 13       'Колонка объём потребления ИПУ
Const v2 = 14       'Колонка объём потребления Норматив
Const v3 = 15       'Колонка объём потребления Перерасчёт

Dim res() As Adress 'Таблица с результатами
Dim curDom As String 'Адрес текущего дома для поиска

'Загрузка программы
Private Sub UserForm_Activate()
    LabelVersion = "Версия: 1.1 (01.06.2021)"
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
    Dim tab1 As String
    Dim tab2 As String
    Dim max1 As Long
    Dim max2 As Long
    Dim tabA As String
    Dim UK As String
    Dim LaskUK As String
    
    tab1 = TextBoxTN
    tab2 = TextBoxHVS
    tabA = TextBoxUK
    
    Call Misc.Message("Подготовка...")
    
    'Чтение адресов
    Mnth = TextBoxMonth
    Dim aMax As Long
    aMax = 0
    Do While Sheets(tabA).Cells(aMax + 2, 1) <> ""
        aMax = aMax + 1
        ReDim Preserve houses(aMax) As Adress
        houses(aMax).UK = Trim(Sheets(tabA).Cells(aMax + 1, 1))
        houses(aMax).ul = Trim(Sheets(tabA).Cells(aMax + 1, 6)) + " " + Trim(Sheets(tabA).Cells(aMax + 1, 7))
        houses(aMax).dom = LCase(Trim(Sheets(tabA).Cells(aMax + 1, 8)) + Trim(Sheets(tabA).Cells(aMax + 1, 9)))
        houses(aMax).korp = Trim(Sheets(tabA).Cells(aMax + 1, 10))
    Loop
    
    'Узнаём размер входных таблиц
    max1 = FindMax(tab1)
    max2 = FindMax(tab2)
    
    'Попёрли клепать отчёт...
    
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
            Cells(1, 2) = "Потребитель": Range(Cells(1, 2), Cells(2, 2)).Merge
            Cells(1, 3) = "Населённый пункт, улица": Range(Cells(1, 3), Cells(2, 3)).Merge
            Cells(1, 4) = "Дом": Range(Cells(1, 4), Cells(2, 4)).Merge
            Cells(1, 5) = "Корпус": Range(Cells(1, 5), Cells(2, 5)).Merge
            Cells(1, 6) = "Квартира": Range(Cells(1, 6), Cells(2, 6)).Merge
            Cells(1, 7) = "Прописано": Range(Cells(1, 7), Cells(2, 7)).Merge
            
            For i = 1 To 19
                Cells(2, 7 + i) = Sheets(tab1).Cells(1, 10 + i)
                Cells(2, 26 + i) = Sheets(tab2).Cells(1, 10 + i)
            Next
            Cells(1, 8) = tab1
            Range(Cells(1, 8), Cells(1, 26)).Merge
            Cells(1, 27) = tab2
            Range(Cells(1, 27), Cells(1, 45)).Merge
            Call Borders(1, 1, 2, 45)
            Call Borders(1, 8, 2, 26)
            
            Columns(1).ColumnWidth = 5
            Columns(2).ColumnWidth = 13
            Columns(3).ColumnWidth = 13
            For i = 4 To 7: Columns(i).ColumnWidth = 5: Next
            Columns(26).ColumnWidth = 28
            Columns(45).ColumnWidth = 28
            nd = 1
        End If
    
        'Построение отчёта
        curDom = houses(a).ul + ", " + LCase(houses(a).dom)
        Call Misc.Message("Построение отчёта... " + CStr(a) + " из " + CStr(aMax) + " (" + curDom + ")")
        ReDim res(0) As Adress
        Call FindHome(tab1, max1, 1)
        Call FindHome(tab2, max2, 2)
        
        'Выводим в отчёт данные по текущему дому (если нашлась хоть одна квартира)
        firstr = r
        kvcount = UBound(res)
        If kvcount > 0 Then
            For i = 1 To kvcount
                If i = 1 Then Cells(r, 1) = nd
                Cells(r, 2) = res(i).potr
                Cells(r, 3) = res(i).ul
                Cells(r, 4) = res(i).dom
                Cells(r, 5) = res(i).korp
                Cells(r, 6) = res(i).kv
                Cells(r, 7) = res(i).prop
                
                If res(i).t1 <> 0 Then
                    For j = 1 To 19
                        Cells(r, 7 + j) = Sheets(tab1).Cells(res(i).t1, 10 + j)
                    Next
                End If
                If res(i).t2 <> 0 Then
                    For j = 1 To 19
                        Cells(r, 26 + j) = Sheets(tab2).Cells(res(i).t2, 10 + j)
                    Next
                End If
                r = r + 1
            Next
            Call Borders(firstr, 1, r - 1, 45)
            Call Borders(firstr, 8, r - 1, 26)
            nd = nd + 1
        End If
    Next
    
    Call Misc.Message("Готово!")
    End
    
End Sub

'Поиск дома по таблице. n - имя, t - номер таблицы (в зависимости от этого в ту или другую часть будут собираться объёмы)
Sub FindHome(name As String, max As Long, t As Integer)
    For i = 2 To max
    
        ul = Trim(Sheets(name).Cells(i, cUl))
        dom = CStr(Sheets(name).Cells(i, cDom)) + LCase(Trim(Sheets(name).Cells(i, cDom + 1)))
        kv = Sheets(name).Cells(i, cKv)
        
        If curDom = ul + ", " + dom Then
            
            'Ищем, есть ли адрес в результатах
            Find = 0
            For j = 1 To UBound(res)
                If res(j).ul = ul And res(j).dom = dom And res(j).kv = kv Then
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
            res(Find).potr = Sheets(name).Cells(i, cPotr)
            res(Find).kv = Sheets(name).Cells(i, cKv)
            res(Find).prop = Sheets(name).Cells(i, cPr)
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