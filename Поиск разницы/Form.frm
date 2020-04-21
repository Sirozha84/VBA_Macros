VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "Поиск разницы"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7200
   OleObjectBlob   =   "Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Version = "0.2 (21.04.2020)"

Private Tab1, Tab2, TabRes As String
Private Max1, Max2, MaxRes As Long
Private Sum() As Integer
Private FPS As Integer

Private Sub UserForm_Activate()
    LabelVersion = "Версия: " + Version
    TextBoxTab1 = Sheets(1).name
    TextBoxTab2 = Sheets(2).name
End Sub

Private Sub CheckBoxCompare_Click()
    TextBoxCompare.Enabled = CheckBoxCompare.Value
End Sub

Private Sub CommandButtonRun_Click()
    LabelStatus = "Подготовка..."
    CommandButtonRun.Enabled = False
    DoEvents
    
    ReDim Sum(2)
    Sum(0) = 8
    Sum(1) = 9
    Sum(2) = 10
    
    'On Error GoTo Error
    
    Tab1 = TextBoxTab1.Value
    Tab2 = TextBoxTab2.Value
    TabRes = TextBoxTabRes.Value
    Call Misc.NewTab(TabRes, CheckBoxCreate.Value)
    
    Max1 = Misc.FindMax(Tab1)
    Max2 = Misc.FindMax(Tab2)
    MaxRes = Max1
    FPS = 100 ' MaxRes / 1000
    
    Search
    
    LabelStatus = "Завершено..."
    'LabelStatus = Max1
Error:
    CommandButtonRun.Enabled = True
End Sub

Private Sub CommandButtonExit_Click()
    End
End Sub

Private Sub Search()
    
    'Рисуем шапку
    c = 1
    For i = 1 To Columns
        For j = 1 To Head
            Sheets(TabRes).Cells(j, c) = Sheets(Tab1).Cells(j, i)
        Next
        If WhatToDo(i) > 0 Then
            Sheets(TabRes).Cells(1, c) = Sheets(Tab1).Cells(1, i) + " (" + Tab1 + ")"
            c = c + 1
            Sheets(TabRes).Cells(1, c) = Sheets(Tab2).Cells(1, i) + " (" + Tab2 + ")"
            c = c + 1
            Sheets(TabRes).Cells(1, c) = Sheets(Tab1).Cells(1, i) + " (Сумма)"
        End If
        c = c + 1
    Next
    
    'Копируем данные из первой таблицы 1
    For i = Head + 1 To Max1
        c = 1
        For j = 1 To Columns
            Sheets(TabRes).Cells(i, c) = Sheets(Tab1).Cells(i, j)
            If WhatToDo(j) > 0 Then c = c + 2
            c = c + 1
        Next
    Next
    
    'Подставляем данные из таблицы 2
    LabelStatus = "Сопостваление таблиц..."
    CommandButtonRun.Enabled = False
    For ii = Head + 1 To Max2
        If ii Mod FPS = 0 Then LabelStatus = "Сопоставление таблиц" + Misc.Progress(ii, Max2): DoEvents
        'Поиск
        Find = False
        
        'Обычный поиск
        'For i = Head + 1 To MaxRes
        '    If Sheets(TabRes).Cells(i, 1) = Sheets(Tab2).Cells(ii, 1) Then
        '        Find = True
        '        Exit For
        '    End If
        'Next
        
        'Поиск в первоначальном диапазоне
        Find = Misc.Search(TabRes, Sheets(Tab2).Cells(ii, 1), Head + 1, Max1)
        'и если там ничего не нашлось - ищем в новых значениях
        If Not Find Then Find = Misc.Search(TabRes, Sheets(Tab2).Cells(ii, 1), Max1, MaxRes)
        
        If Find Then
            'Найдена строка в обоих таблицах
            c = 1
            For j = 1 To Columns
                If WhatToDo(j) = 1 Then
                    c = c + 1
                    Sheets(TabRes).Cells(ii, c) = Sheets(Tab2).Cells(ii, j)
                    c = c + 1
                End If
                c = c + 1
            Next
        Else
            'Найдена строка в таблице 2, которой нет в таблице 1
            MaxRes = MaxRes + 1
            Sheets(TabRes).Cells(MaxRes, 1) = Sheets(Tab2).Cells(ii, 1)
            c = 1
            For j = 1 To Columns
                If WhatToDo(j) = 0 Then
                    Sheets(TabRes).Cells(MaxRes, c) = Sheets(Tab2).Cells(ii, j)
                Else
                    c = c + 1
                    Sheets(TabRes).Cells(MaxRes, c) = Sheets(Tab2).Cells(ii, j)
                    c = c + 1
                End If
                c = c + 1
            Next
        End If
        
    Next
    
    'Действия
    LabelStatus = "Обработка полученных данных..."
    DoEvents
    For i = Head + 1 To MaxRes
        c = 1
        For j = 1 To Columns
            If WhatToDo(j) = 1 Then
                c = c + 2
                Sheets(TabRes).Cells(i, c) = Sheets(TabRes).Cells(i, c - 2) + Sheets(TabRes).Cells(i, c - 1)
                'Помечаем цветом (далее надо будет сделать это опционально)
                Sheets(TabRes).Cells(i, c - 2).Interior.Color = RGB(255, 255, 196)
                Sheets(TabRes).Cells(i, c - 1).Interior.Color = RGB(255, 255, 196)
                Sheets(TabRes).Cells(i, c).Interior.Color = RGB(196, 255, 196)
            End If
            c = c + 1
        Next
    Next
    
End Sub

'Проверка "Что делать в этой ячейке?": 0 - ничего, 1 - сумма, 2 - сравнение
Function WhatToDo(ByVal n As Integer)
    Find = False
    For i = 0 To UBound(Sum)
        If Sum(i) = n Then
            Find = True
            Exit For
        End If
    Next
    If Find Then WhatToDo = 1
    
    'И то же самое надо сделать для сравнения
    
End Function