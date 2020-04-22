VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "Поиск разницы"
   ClientHeight    =   6360
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
Const Version = "0.3 (22.04.2020)"

Private Tab1, Tab2, TabRes As String
Private Max1, Max2, MaxRes As Long
Private ACop() As Integer
Private ASum() As Integer
Private ACom() As Integer
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
    CommandButtonRun.Enabled = False
    Application.ScreenUpdating = False
    
    'Пока хардкодим это, потом надо будет сделать читку с формы
    ReDim ACop(0)
    ACop(0) = 2
    ReDim ASum(0)
    ASum(0) = 3
    ReDim ACom(0)
    ACom(0) = 4
    
    'On Error GoTo Error
    
    FPS = 100 ' MaxRes / 1000
    
    Prepare
    Search
    
    LabelStatus = "Завершено..."
    Application.ScreenUpdating = True
Error:
    CommandButtonRun.Enabled = True
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButtonExit_Click()
    End
End Sub

'Подготовка
Private Sub Prepare()
    Tab1 = TextBoxTab1.Value
    Tab2 = TextBoxTab2.Value
    TabRes = TextBoxTabRes.Value
    
    Call Misc.NewTab(TabRes, CheckBoxCreate.Value)
    
    LabelStatus = "Проверка таблиц...": DoEvents
    Max1 = Misc.FindMax(Tab1)
    Max2 = Misc.FindMax(Tab2)
    Call MethodSelect(Tab1, Head + 1, Max1)
    Call MethodSelect(Tab2, Head + 1, Max2)
    
    MaxRes = Max1

End Sub

Private Sub Search()
    
    'Рисуем шапку
    c = 1
    For i = 1 To Columns
        For j = 1 To Head
            Sheets(TabRes).Cells(j, c) = Sheets(Tab1).Cells(j, i)
        Next
        Action = WhatToDo(i)
        If Action > 0 Then
            Sheets(TabRes).Cells(1, c) = Sheets(Tab1).Cells(1, i) + " (" + Tab1 + ")"
            c = c + 1
            Sheets(TabRes).Cells(1, c) = Sheets(Tab2).Cells(1, i) + " (" + Tab2 + ")"
            If Action = 2 Then
                c = c + 1
                Sheets(TabRes).Cells(1, c) = Sheets(Tab1).Cells(1, i) + " (Сумма)"
            End If
            If Action = 3 Then
                c = c + 1
                Sheets(TabRes).Cells(1, c) = Sheets(Tab1).Cells(1, i) + " (Разница)"
            End If
        End If
        c = c + 1
    Next
    
    'Копируем данные из первой таблицы 1
    For i = Head + 1 To Max1
        c = 1
        For j = 1 To Columns
            Sheets(TabRes).Cells(i, c) = Sheets(Tab1).Cells(i, j)
            Action = WhatToDo(j)
            If WhatToDo(j) = 1 Then c = c + 1
            If WhatToDo(j) > 1 Then c = c + 2
            c = c + 1
        Next
    Next
    
    'Подставляем данные из таблицы 2
    LabelStatus = "Сопостваление таблиц..."
    CommandButtonRun.Enabled = False
    For ii = Head + 1 To Max2
        If ii Mod FPS = 0 Then LabelStatus = "Сопоставление таблиц" + Misc.Progress(ii, Max2): DoEvents
        
        'Поиск
        Find = 0
        'Поиск в первоначальном диапазоне
        Find = Misc.Search(TabRes, Sheets(Tab2).Cells(ii, 1), Head + 1, Max1)
        'и если там ничего не нашлось - ищем в новых значениях
        If Find = 0 Then Find = Misc.Search(TabRes, Sheets(Tab2).Cells(ii, 1), Max1, MaxRes)
            
            
        
        str2 = 0
        If Find > 0 Then
            'Найдена строка в обоих таблицах
            str2 = Find
        Else
            'Найдена строка в таблице 2, которой нет в таблице 1
            MaxRes = MaxRes + 1
            str2 = MaxRes
            Sheets(TabRes).Cells(MaxRes, 1) = Sheets(Tab2).Cells(ii, 1)
        End If
        'Копирование данных из второй таблицы
        c = 1
        For j = 1 To Columns
            Action = WhatToDo(j)
            If Action > 0 Then c = c + 1
            Sheets(TabRes).Cells(str2, c) = Sheets(Tab2).Cells(ii, j)
            If Action > 1 Then c = c + 1
            c = c + 1
        Next
    Next
    
    'Действия
    LabelStatus = "Обработка полученных данных..."
    DoEvents
    For i = Head + 1 To MaxRes
        c = 1
        For j = 1 To Columns
            Action = WhatToDo(j)
            If Action = 1 Then
                'Помечаем цветом (далее надо будет сделать это опционально)
                Sheets(TabRes).Cells(i, c + 0).Interior.Color = RGB(196, 196, 196)
                Sheets(TabRes).Cells(i, c + 1).Interior.Color = RGB(196, 196, 196)
                c = c + 1
            End If
            If Action = 2 Then
                Sheets(TabRes).Cells(i, c + 2) = Sheets(TabRes).Cells(i, c + 0) + Sheets(TabRes).Cells(i, c + 1)
                'Помечаем цветом (далее надо будет сделать это опционально)
                Sheets(TabRes).Cells(i, c + 0).Interior.Color = RGB(196, 196, 255)
                Sheets(TabRes).Cells(i, c + 1).Interior.Color = RGB(196, 196, 255)
                Sheets(TabRes).Cells(i, c + 2).Interior.Color = RGB(128, 128, 255)
                c = c + 2
            End If
            If Action = 3 Then
                Min = 5 'Потом надо будет это засунуть в опции
                r = Sheets(TabRes).Cells(i, c + 1) - Sheets(TabRes).Cells(i, c + 0)
                Sheets(TabRes).Cells(i, c + 2) = r
                r = Abs(r)
                'Помечаем цветом (далее надо будет сделать это опционально)
                If r = 0 Then
                    Sheets(TabRes).Cells(i, c + 0).Interior.Color = RGB(196, 255, 196)
                    Sheets(TabRes).Cells(i, c + 1).Interior.Color = RGB(196, 255, 196)
                    Sheets(TabRes).Cells(i, c + 2).Interior.Color = RGB(128, 255, 128)
                End If
                If r > 0 And r <= Min Then
                    Sheets(TabRes).Cells(i, c + 0).Interior.Color = RGB(255, 255, 196)
                    Sheets(TabRes).Cells(i, c + 1).Interior.Color = RGB(255, 255, 196)
                    Sheets(TabRes).Cells(i, c + 2).Interior.Color = RGB(255, 255, 128)
                End If
                If r > Min Then
                    Sheets(TabRes).Cells(i, c + 0).Interior.Color = RGB(255, 196, 196)
                    Sheets(TabRes).Cells(i, c + 1).Interior.Color = RGB(255, 196, 196)
                    Sheets(TabRes).Cells(i, c + 2).Interior.Color = RGB(255, 128, 128)
                End If
                c = c + 2
            End If
            c = c + 1
        Next
    Next
    
End Sub

'Проверка "Что делать в этой ячейке?": 0 - ничего, 1 - копия, 2 - сумма, 3 - сравнение
Function WhatToDo(ByVal n As Integer)
    Find = False
    'Копируем?
    For i = 0 To UBound(ACop)
        If ACop(i) = n Then
            Find = True
            Exit For
        End If
    Next
    If Find Then WhatToDo = 1: Exit Function
    'Суммируем
    For i = 0 To UBound(ASum)
        If ASum(i) = n Then
            Find = True
            Exit For
        End If
    Next
    If Find Then WhatToDo = 2: Exit Function
    'Сравниваем
    For i = 0 To UBound(ACom)
        If ACom(i) = n Then
            Find = True
            Exit For
        End If
    Next
    If Find Then WhatToDo = 3: Exit Function
End Function