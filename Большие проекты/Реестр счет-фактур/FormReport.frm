VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormReport 
   Caption         =   "��������� ������"
   ClientHeight    =   3600
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
Private Sub Frame1_Click()

End Sub

'�������� ���������
Private Sub UserForm_Activate()
    LabelVersion = "������: 0.1 (09.07.2020)"
    TextBoxSrc = Sheets(1).Name
End Sub

'�����
Private Sub ButtonClose_Click()
    End
End Sub

'������
Private Sub ButtonOK_Click()
    ReDim records(0) As Record
    Dim tabSrc As String
    Dim tabReport As String
    Dim i As Long
    Dim Name As String
    Dim Dat As String
    tabSrc = TextBoxSrc
    tabReport = TextBoxReport
    Mnth = TextBoxMonth
    Max = 0
    
    '������ �������� ������
    Call Misc.Message("������ �������� ������...")
    i = 2
    last = Sheets(tabSrc).Cells(i, 3)

    Name = ""
    Dat = ""
    VolumeHeat = 0
    VolumeHW = 0
    PriceHeat = 0
    PriceHW = 0
    VolumeInfo = 0
    Do
        num = Sheets(tabSrc).Cells(i, 3)
        '������� ������ ���������� �� ���������� ��� ����� �����������
        If num <> last Or Sheets(tabSrc).Cells(i, 1) = "" Then
            Max = Max + 1
            ReDim Preserve records(Max) As Record
            records(Max).Name = Name
            records(Max).Number = Number
            records(Max).Date = Dat
            records(Max).VolumeHeat = VolumeHeat
            records(Max).VolumeHW = VolumeHW
            records(Max).PriceHeat = PriceHeat
            records(Max).PriceHW = PriceHW
            records(Max).VolumeInfo = VolumeInfo
'            '���������� ��������
            VolumeHeat = 0
            VolumeHW = 0
            PriceHeat = 0
            PriceHW = 0
            VolumeInfo = 0
            last = num
        End If
        '����� ��� � ���������� ������, ������������ ������
        If last = num Then
            Name = Sheets(tabSrc).Cells(i, 1)
            Number = Sheets(tabSrc).Cells(i, 3)
            Dat = Sheets(tabSrc).Cells(i, 4)
            If Sheets(tabSrc).Cells(i, 2) = "�������� �������" Then
                VolumeHeat = VolumeHeat + Sheets(tabSrc).Cells(i, 5)
                PriceHeat = PriceHeat + Sheets(tabSrc).Cells(i, 8)
            End If
            If Sheets(tabSrc).Cells(i, 2) = "������� ����" Then
                VolumeHW = VolumeHW + Sheets(tabSrc).Cells(i, 5)
                PriceHW = PriceHW + Sheets(tabSrc).Cells(i, 8)
            End If
            'MsgBox (Sheets(tabSrc).Cells(i, 2))
            If Left(Sheets(tabSrc).Cells(i, 2), 10) = "���������:" Then
                VolumeInfo = Sheets(tabSrc).Cells(i, 8)
            End If
        
        End If
        If Sheets(tabSrc).Cells(i, 1) = "" Then Exit Do
        last = num
        i = i + 1
    Loop
    
    '������/������� �������������� ��������
    Call Misc.Message("������������ ������...")
    Call Misc.NewTab(tabReport, True)
    Sheets(tabReport).Select
    
    '������������� ������ ������� � �������
    Columns(1).ColumnWidth = 4.86
    Columns(2).ColumnWidth = 37.14
    Columns(3).ColumnWidth = 20.14
    Columns(4).ColumnWidth = 14.71
    Columns(5).ColumnWidth = 15.86
    Columns(6).ColumnWidth = 19.29
    Columns(7).ColumnWidth = 16.86
    Columns(8).ColumnWidth = 14.71
    Columns(9).ColumnWidth = 14.71
    Columns(10).ColumnWidth = 14.71
    Columns(11).ColumnWidth = 14.71
    For i = 4 To 7
        Columns(i).NumberFormat = "### ### ##0.00"
    Next
    
    '���������
    Dim s As Integer
    Rows(1).RowHeight = 56.25
    Call MergeAndCenter(1, 1, 1, 9, "������ ��������� ���������� ��� �������� ����� �� ������������ ������, " + _
        "������������� ������������� � ������������� ��������� � ��������������� ����� ��� ����� �����, " + _
        "����������� �� ���������� ���� ��� ������� ������ ��������� � ������������� ��������������� ������� " + _
        "�������������, � ���������������� �������������, ������������� �������� �������, �������������� " + _
        "����������������� (�����������������) ��� ����������, ����������� �� �������� ������� " + _
        "(����� � ����������������� �����������)")
    Call MergeAndCenter(2, 1, 1, 9, "�� " + Mnth + " ����")
    Call MergeAndCenter(3, 1, 1, 9, "(������������ ������)")
    Call MergeAndCenter(4, 1, 1, 9, "��� ""�������-�������""")
    Call MergeAndCenter(5, 1, 1, 9, "(������������ ����������������� �����������)")
    s = 15
    
    '�����
    Call MergeAndCenter(s, 1, 2, 1, "� �/�")
    Call MergeAndCenter(s, 2, 2, 1, "������������ �����������")
    Call MergeAndCenter(s, 3, 2, 1, "������������ ������������� ������� (�������� �������, ������� ����)")
    Call MergeAndCenter(s, 4, 2, 1, "����� ����������� ������������� ������� �� ��������� ����������")
    Call MergeAndCenter(s, 5, 1, 3, "����� �� ����-�������� (���. ������)")
    Call MergeAndCenter(s, 8, 2, 2, "� � ���� ����-�������")
    Call MergeAndCenter(s, 10, 2, 1, "������, �� ������� ����������� ����-�������")
    Call MergeAndCenter(s, 11, 2, 1, "���������: ����� �������������," + Chr(10) + "���.�.")
    Rows(8).RowHeight = 90
    s = s + 1
    Call MergeAndCenter(s, 5, 1, 1, "�� ���������")
    Call MergeAndCenter(s, 6, 1, 1, "�� ��������� ��������� �������� ��� �������� ������ �� �������� �������������")
    Call MergeAndCenter(s, 7, 1, 1, "�����")
    
    s = 17
    Range(Cells(s, 8), Cells(s, 9)).Merge
    For i = 1 To 8
        Cells(s, i).NumberFormat = "#"
        Cells(s, i) = i
        Cells(s, i).HorizontalAlignment = xlCenter
    Next
    For i = 9 To 10
        Cells(s, i).NumberFormat = "#"
        Cells(s, i + 1) = i
        Cells(s, i + 1).HorizontalAlignment = xlCenter
    Next
    Range(Cells(s - 2, 1), Cells(s, 11)).Borders.Weight = 3
    s = s + 1
    
    '������� �������
    allVolumeHeat = 0
    allVolumeHW = 0
    allDocsHeat = 0
    allDocsHW = 0
    allPriceHeat = 0
    allPriceHW = 0
    
    'For t = 1 To 2 (��� ��� ���� �� �����, �������� � ����� �����-������ �����������
    num = 1
    sumVolumeHeat = 0
    sumVolumeHW = 0
    sumPriceHeat = 0
    sumPriceHW = 0
    sumVolumeInfo = 0
    
    '�����
    Cells(s, 1) = "1"
    Call MergeAndCenter(s, 1, 1, 1, CStr(t))
    Cells(s, 2) = "���������"
    Range(Cells(s, 2), Cells(s, 11)).Merge
    Range(Cells(s, 1), Cells(s, 11)).Borders.Weight = 3
    s = s + 1
    
    '��������� �����
    For i = 1 To Max
            
        '�������� �������
        Call MergeAndCenter(s, 1, 3, 1, "'1." + CStr(num))
        Call MergeAndCenter(s, 2, 3, 1, records(i).Name)
        'Cells(s, 2) = records(i).Name
        Cells(s, 3) = "�������� �������"
        Cells(s, 4) = records(i).VolumeHeat
        Cells(s, 5) = Round(records(i).PriceHeat / 1000, 3)
        Cells(s, 7) = Round(records(i).PriceHeat / 1000, 3)
        Cells(s, 8) = records(i).Number
        Cells(s, 9).HorizontalAlignment = xlRight
        Cells(s, 9) = records(i).Date
        Cells(s, 10) = Mnth
        Cells(s, 11) = records(i).VolumeInfo
        VolumeSum = records(i).VolumeHeat
        PriceSum = records(i).PriceHeat
        
        Range(Cells(s, 1), Cells(s, 7)).Borders.Weight = 2
        Range(Cells(s, 10), Cells(s, 11)).Borders.Weight = 2
        Range(Cells(s, 8), Cells(s, 9)).Borders(xlEdgeBottom).Weight = 2
        s = s + 1
        
        '������� ����
        Cells(s, 3) = "������� ����"
        Cells(s, 4) = records(i).VolumeHW
        Cells(s, 6) = Round(records(i).PriceHW / 1000, 3)
        Cells(s, 7) = Round(records(i).PriceHW / 1000, 3)
        Cells(s, 8) = records(i).Number
        Cells(s, 9).HorizontalAlignment = xlRight
        Cells(s, 9) = records(i).Date
        Cells(s, 10) = Mnth
        VolumeSum = VolumeSum + records(i).VolumeHW
        PriceSum = PriceSum + records(i).PriceHW
        Range(Cells(s, 1), Cells(s, 7)).Borders.Weight = 2
        Range(Cells(s, 10), Cells(s, 11)).Borders.Weight = 2
        Range(Cells(s, 8), Cells(s, 9)).Borders(xlEdgeBottom).Weight = 2
        s = s + 1
            
        '�����
        Cells(s, 3) = "����� �� �����������"
        Cells(s, 4) = VolumeSum
        Cells(s, 7) = Round(PriceSum / 1000, 3)
        Cells(s, 10) = Mnth
        Range(Cells(s, 2), Cells(s, 11)).Interior.Color = RGB(221, 235, 247)
        Range(Cells(s, 1), Cells(s, 7)).Borders.Weight = 2
        Range(Cells(s, 10), Cells(s, 11)).Borders.Weight = 2
        Range(Cells(s, 8), Cells(s, 9)).Borders(xlEdgeBottom).Weight = 2
        Range(Cells(s - 2, 1), Cells(s, 1)).Borders.Weight = 3
        Range(Cells(s - 2, 11), Cells(s, 11)).Borders(xlEdgeRight).Weight = 3
        s = s + 1
            
        '��������
        sumVolumeHeat = sumVolumeHeat + records(i).VolumeHeat
        sumVolumeHW = sumVolumeHW + records(i).VolumeHW
        sumPriceHeat = sumPriceHeat + records(i).PriceHeat
        sumPriceHW = sumPriceHW + records(i).PriceHW
        sumVolumeInfo = sumVolumeInfo + records(i).VolumeInfo
        num = num + 1
            
    Next
    
    '�����
    
    '�������� �������
    Range(Cells(s, 2), Cells(s + 2, 2)).Merge
    Range(Cells(s, 2), Cells(s + 2, 2)).VerticalAlignment = xlCenter
    Cells(s, 2) = "�� ������  ������������ ����������"
    Cells(s, 3) = "�������� �������"
    Cells(s, 4) = sumVolumeHeat
    Cells(s, 5) = Round(sumPriceHeat / 1000, 3)
    Cells(s, 7) = Round(sumPriceHeat / 1000, 3)
    Cells(s, 11) = sumVolumeInfo
    Range(Cells(s, 1), Cells(s, 11)).Borders.Weight = 2
    s = s + 1
    
    '������� ����
    Cells(s, 3) = "������� ����"
    Cells(s, 4) = sumVolumeHW
    Cells(s, 6) = Round(sumPriceHW / 1000, 3)
    Cells(s, 7) = Round(sumPriceHW / 1000, 3)
    Range(Cells(s, 1), Cells(s, 11)).Borders.Weight = 2
    s = s + 1
    
    '�����
    Cells(s, 3) = "����� �� �����������"
    Cells(s, 4) = sumVolumeHeat + sumVolumeHW
    Cells(s, 7) = Round((sumPriceHeat + sumPriceHW) / 1000, 3)
    Cells(s, 11) = sumVolumeInfo
    Range(Cells(s - 2, 1), Cells(s, 1)).Merge
    Range(Cells(s, 1), Cells(s, 11)).Borders.Weight = 2
    Range(Cells(s - 2, 1), Cells(s, 1)).Borders.Weight = 3
    Range(Cells(s - 2, 2), Cells(s, 2)).Borders.Weight = 3
    Range(Cells(s - 2, 3), Cells(s - 2, 11)).Borders(xlEdgeTop).Weight = 3
    Range(Cells(s - 2, 11), Cells(s, 11)).Borders(xlEdgeRight).Weight = 3
    Range(Cells(s, 1), Cells(s, 11)).Borders(xlEdgeBottom).Weight = 3
    s = s + 1
    
    '��������
    allVolumeHeat = allVolumeHeat + sumVolumeHeat
    allVolumeHW = allVolumeHW + sumVolumeHW
    allDocsHeat = allDocsHeat + sumDocsHeat
    allDocsHW = allDocsHW + sumDocsHW
    allPriceHeat = allPriceHeat + sumPriceHeat
    allPriceHW = allPriceHW + sumPriceHW
    'Next
    
    
    s = s + 4
    
    '�������
    Cells(s, 1) = "�������������� ��������"
    Cells(s, 8) = "�.�. ����������"
    Cells(s, 8).HorizontalAlignment = xlCenter
    s = s + 1
    Call MergeAndCenter(s, 4, 1, 2, "(�������)")
    Range(Cells(s, 4), Cells(s, 5)).Borders(xlEdgeTop).LineStyle = True
    Cells(s, 8) = "(���)"
    Cells(s, 8).HorizontalAlignment = xlCenter
    Cells(s, 8).Borders(xlEdgeTop).LineStyle = True
    s = s + 1
    Cells(s, 3) = "�. �."
    s = s + 1
    Cells(s, 3) = "(� ������ �������)"
    s = s + 3
    Cells(s, 1) = "��������� �.�. (3919)757719"
    
    '������!
    Call Misc.Message("������!")
    Application.ScreenUpdating = True
    
End Sub

Sub MergeAndCenter(R As Integer, C As Integer, height As Integer, width As Integer, text As String)
    Cells(R, C).HorizontalAlignment = xlCenter
    Cells(R, C).VerticalAlignment = xlCenter
    Range(Cells(R, C), Cells(R + height - 1, C + width - 1)).Merge
    Cells(R, C).WrapText = True
    Cells(R, C) = text
End Sub