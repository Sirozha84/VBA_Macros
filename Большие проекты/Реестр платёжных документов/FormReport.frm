VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormReport 
   Caption         =   "��������� ������"
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
'�������� ���������
Private Sub UserForm_Activate()
    LabelVersion = "������: 0.2 (02.07.2020)"
    TextBoxHeat = Sheets(1).name
    TextBoxHW = Sheets(2).name
End Sub

'������� �� ������� ������
Private Sub ButtonHelp_Click()
    MsgBox ("�� �������� �������� ��� �������: �������� ������� � ������� ����. ������� ������� ���������:" + Chr(10) + _
        "������� 1: �����" + Chr(10) + _
        "������� 2: ���������� �������� ����������" + Chr(10) + _
        "������� 3: ����� �����������. ��������" + Chr(10) + _
        "������� 4: ���������� �� ������������ ������" + Chr(10) + _
        "������� 5: ������� ������� (��� ��� ���)" + Chr(10) + _
        "������ ������� �� ������ ������, �������� ������ ��� �����.")
End Sub

'�������� ������� ������
Private Sub ButtonSwap_Click()
    temp = TextBoxHeat
    TextBoxHeat = TextBoxHW
    TextBoxHW = temp
End Sub

'�����
Private Sub ButtonClose_Click()
    End
End Sub

'������
Private Sub ButtonOK_Click()
    ReDim records(0) As Record
    Dim tabRes As String
    tabHeat = TextBoxHeat
    tabHW = TextBoxHW
    tabRes = TextBoxReport
    Mnth = TextBoxMonth
    Max = 0
    
    '���������� ������ "�������� �������"
    Call Misc.Message("���������� ������ �� ""�������� �������""...")
    i = 2
    Do While Sheets(tabHeat).Cells(i, 1) <> ""
        Max = Max + 1
        ReDim Preserve records(Max) As Record
        records(Max).Adress = Sheets(tabHeat).Cells(i, 1)
        records(Max).PPHeat = Sheets(tabHeat).Cells(i, 2)
        records(Max).VolumeHeat = Sheets(tabHeat).Cells(i, 3)
        records(Max).PriceHeat = Sheets(tabHeat).Cells(i, 4)
        records(Max).tag = Sheets(tabHeat).Cells(i, 5)
        i = i + 1
    Loop
    MaxG = Max
    
    '���������� ������ "������� ����"
    Call Misc.Message("���������� ������ �� ""������� ����""...")
    i = 2
    Do While Sheets(tabHW).Cells(i, 1) <> ""
        '���������, ��� �� �� ����� ������, ������ �� "�������� �������"
        Find = 0
        For j = 1 To MaxG
            If Sheets(tabHW).Cells(i, 1) = records(j).Adress Then
                Find = j
                Exit For
            End If
        Next
        If Find > 0 Then
            '������ ����, ��������� �
            records(Find).PPHW = Sheets(tabHW).Cells(i, 2)
            records(Find).VolumeHW = Sheets(tabHW).Cells(i, 3)
            records(Find).PriceHW = Sheets(tabHW).Cells(i, 4)
        Else
            '���� ����� ������ ��� ���, ��������� �����
            Max = Max + 1
            ReDim Preserve records(Max) As Record
            records(Max).Adress = Sheets(tabHW).Cells(i, 1)
            records(Max).PPHW = Sheets(tabHW).Cells(i, 2)
            records(Max).VolumeHW = Sheets(tabHW).Cells(i, 3)
            records(Max).PriceHW = Sheets(tabHW).Cells(i, 4)
            records(Max).tag = Sheets(tabHW).Cells(i, 5)
        End If
        
        i = i + 1
    Loop
    
    '������/������� �������������� ��������
    Call Misc.Message("������������ ������...")
    Call Misc.NewTab(tabRes, True)
    Sheets(tabRes).Select
    
    '������������� ������ �������
    Columns(1).ColumnWidth = 5.86
    Columns(2).ColumnWidth = 46.71
    Columns(3).ColumnWidth = 18
    Columns(4).ColumnWidth = 16
    Columns(5).ColumnWidth = 17.56
    Columns(6).ColumnWidth = 11.71
    Columns(7).ColumnWidth = 16.71
    Columns(8).ColumnWidth = 13.28
    Columns(9).ColumnWidth = 13.29
    
    '�����
    Rows(1).RowHeight = 56.25
    Call MergeAndCenter(1, 1, 1, 9, "������ ��������� ���������� ��� �������� ����� �� ������������ ������, " + _
        "������������� ������������� � ������������� ��������� � ��������������� ����� ��� ����� �����, " + _
        "����������� �� ���������� ���� ��� ������� ������ ��������� � ������������� ��������������� ������� " + _
        "�������������, � ���������������� �������������, ������������� �������� �������, �������������� " + _
        "����������������� (�����������������) ��� ����������, ����������� �� �������� ������� " + _
        "(����� � ����������������� �����������)")
    Call MergeAndCenter(2, 1, 1, 9, Cells(2, 1) = "�� " + Mnth + " ����")
    Call MergeAndCenter(3, 1, 1, 9, Cells(3, 1) = "(������������ ������)")
    Call MergeAndCenter(4, 1, 1, 9, "��� ""�������-�������""")
    Call MergeAndCenter(5, 1, 1, 9, "(������������ ����������������� �����������)")
    Call MergeAndCenter(7, 1, 3, 1, "� �/�")
    Call MergeAndCenter(7, 2, 3, 1, "����� ���������������� ���  ������ ����")
    Call MergeAndCenter(7, 3, 3, 1, "������������ ������������� ������� (�������� �������, ������� ����)")
    Call MergeAndCenter(7, 4, 3, 1, "����� ����������� ������������� ������� �� ��������� ����������")
    Call MergeAndCenter(7, 5, 3, 1, "����������  ��������� ���������� ��� �������� ����� ��  ������������ ������ (��������� ���������)")
    Call MergeAndCenter(7, 6, 1, 3, "����� �� ��������� ����������")
    Call MergeAndCenter(8, 6, 1, 3, "(���. ������)")
    Call MergeAndCenter(9, 6, 1, 1, "�� ���������")
    Call MergeAndCenter(9, 7, 1, 1, "�� ��������� ��������� �������� ��� �������� ������ �� �������� �������������")
    Call MergeAndCenter(9, 8, 1, 1, "�����")
    Call MergeAndCenter(7, 9, 3, 1, "������, �� ������� ����������� ��������� ���������")
    
    For i = 1 To 9
        Cells(10, i) = i
        Cells(10, i).HorizontalAlignment = xlCenter
    Next
    
    '��������� �����
    
    '������� �������
    Dim s As Integer
    s = 11
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
        
        '�����
        Call MergeAndCenter(s, 1, 1, 1, t)
        If t = 1 Then
            Cells(s, 2) = "��������������� ����"
        Else
            Cells(s, 2) = "����� ����"
        End If
        Sheets(tabRes).Range(Cells(s, 2), Cells(s, 9)).Merge
        s = s + 1
        
        For i = 1 To Max
            If ((t = 1 And StrComp(records(i).tag, "���", 1) = 0) Or _
                (t = 2 And StrComp(records(i).tag, "���", 1) = 0)) And _
                (records(i).VolumeHeat <> 0 Or records(i).VolumeHW <> 0) Then
                
                '�������� �������
                Call MergeAndCenter(s, 3, 1, 1, "'" + CStr(t) + "." + CStr(num))
                Cells(s, 2) = records(i).Adress
                Cells(s, 3) = "�������� �������"
                Cells(s, 4) = records(i).VolumeHeat
                Cells(s, 5) = records(i).PPHeat
                Cells(s, 6) = Round(records(i).PriceHeat / 1000, 3)
                Cells(s, 8) = Round(records(i).PriceHeat / 1000, 3)
                Cells(s, 9) = Mnth
                VolumeSum = records(i).VolumeHeat
                PriceSum = records(i).PriceHeat
                s = s + 1
                
                '������� ����
                Cells(s, 2) = records(i).Adress
                Cells(s, 3) = "������� ����"
                Cells(s, 4) = records(i).VolumeHW
                Cells(s, 5) = records(i).PPHW
                Cells(s, 7) = Round(records(i).PriceHW / 1000, 3)
                Cells(s, 8) = Round(records(i).PriceHW / 1000, 3)
                Cells(s, 9) = Mnth
                VolumeSum = VolumeSum + records(i).VolumeHW
                PriceSum = PriceSum + records(i).PriceHW
                s = s + 1
                
                '�����
                Cells(s, 2) = records(i).Adress
                Cells(s, 3) = "�����"
                Cells(s, 4) = VolumeSum
                Cells(s, 8) = Round(PriceSum / 1000, 3)
                Cells(s, 9) = Mnth
                Range(Cells(s, 2), Cells(s, 9)).Interior.Color = RGB(221, 235, 247)
                s = s + 1
                
                '��������
                sumVolumeHeat = sumVolumeHeat + records(i).VolumeHeat
                sumVolumeHW = sumVolumeHW + records(i).VolumeHW
                sumDocsHeat = sumDocsHeat + records(i).PPHeat
                sumDocsHW = sumDocsHW + records(i).PPHW
                sumPriceHeat = sumPriceHeat + records(i).PriceHeat
                sumPriceHW = sumPriceHW + records(i).PriceHW
                num = num + 1
                
            End If
        Next
        
        '�����
        If t = 1 Then
            Sheets(tabRes).Cells(s, 2) = "�� ������ ��������������� ����"
        Else
            Sheets(tabRes).Cells(s, 2) = "�� ������ ����� ����"
        End If
        
        '�������� �������
        Range(Cells(s, 2), Cells(s + 2, 2)).Merge
        Range(Cells(s, 2), Cells(s + 2, 2)).VerticalAlignment = xlCenter
        Cells(s, 3) = "�������� �������"
        Cells(s, 4) = sumVolumeHeat
        Cells(s, 5) = sumDocsHeat
        Cells(s, 6) = Round(sumPriceHeat / 1000, 3)
        Cells(s, 8) = Round(sumPriceHeat / 1000, 3)
        s = s + 1
        
        '������� ����
        Cells(s, 3) = "������� ����"
        Cells(s, 4) = sumVolumeHW
        Cells(s, 5) = sumDocsHW
        Cells(s, 7) = Round(sumPriceHW / 1000, 3)
        Cells(s, 8) = Round(sumPriceHW / 1000, 3)
        s = s + 1
        
        '�����
        Cells(s, 3) = "�����"
        Cells(s, 4) = sumVolumeHeat + sumVolumeHW
        Cells(s, 8) = Round((sumPriceHeat + sumPriceHW) / 1000, 3)
        s = s + 1
        
        '��������
        allVolumeHeat = allVolumeHeat + sumVolumeHeat
        allVolumeHW = allVolumeHW + sumVolumeHW
        allDocsHeat = allDocsHeat + sumDocsHeat
        allDocsHW = allDocsHW + sumDocsHW
        allPriceHeat = allPriceHeat + sumPriceHeat
        allPriceHW = allPriceHW + sumPriceHW
    Next
    
    '����� �������� �����
    Sheets(tabRes).Cells(s, 1) = "�� ����������������� �����������"
    Range(Cells(s, 1), Cells(s + 2, 2)).Merge
    Range(Cells(s, 1), Cells(s + 2, 2)).VerticalAlignment = xlCenter
    
    '�������� �������
    Cells(s, 3) = "�������� �������"
    Cells(s, 4) = allVolumeHeat
    Cells(s, 5) = allDocsHeat
    Cells(s, 6) = Round(allPriceHeat / 1000, 3)
    Cells(s, 8) = Round(allPriceHeat / 1000, 3)
    s = s + 1
    
    '������� ����
    Cells(s, 3) = "������� ����"
    Cells(s, 4) = allVolumeHW
    Cells(s, 5) = allDocsHW
    Cells(s, 7) = Round(allPriceHW / 1000, 3)
    Cells(s, 8) = Round(allPriceHW / 1000, 3)
    s = s + 1
    
    '�����
    Cells(s, 3) = "�����"
    Cells(s, 4) = allVolumeHeat + allVolumeHW
    Cells(s, 8) = Round((allPriceHeat + allPriceHW) / 1000, 3)
    s = s + 1
    
    
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