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
    LabelVersion = "������: 0.1 (30.06.2020)"
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
        records(Max).Tag = Sheets(tabHeat).Cells(i, 5)
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
            records(Max).Tag = Sheets(tabHW).Cells(i, 5)
        End If
        
        i = i + 1
    Loop
    
    '������/������� �������������� ��������
    Call Misc.Message("������������ ������...")
    Call Misc.NewTab(tabRes, True)
    
    For i = 1 To Max
        Sheets(tabRes).Cells(i * 3, 1) = records(i).Adress
        Sheets(tabRes).Cells(i * 3, 2) = "�������� �������"
        Sheets(tabRes).Cells(i * 3, 3) = records(i).PPHeat
        Sheets(tabRes).Cells(i * 3, 4) = records(i).VolumeHeat
        Sheets(tabRes).Cells(i * 3, 5) = records(i).PriceHeat
        Sheets(tabRes).Cells(i * 3, 6) = records(i).Tag
        Sheets(tabRes).Cells(i * 3 + 1, 1) = records(i).Adress
        Sheets(tabRes).Cells(i * 3 + 1, 2) = "������� ���"
        Sheets(tabRes).Cells(i * 3 + 1, 3) = records(i).PPHW
        Sheets(tabRes).Cells(i * 3 + 1, 4) = records(i).VolumeHW
        Sheets(tabRes).Cells(i * 3 + 1, 5) = records(i).PriceHW
        Sheets(tabRes).Cells(i * 3 + 1, 6) = records(i).Tag
    Next
    
    
    
    Call Misc.Message("������!")
    Application.ScreenUpdating = True
End Sub