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
Const cUl = 3       '������� �����
Const cDom = 4      '������� ���
Const cKv = 7       '������� ��������
Const v1 = 13       '������� ����� ����������� ���
Const v2 = 14       '������� ����� ����������� ��������
Const v3 = 15       '������� ����� ����������� ����������

Dim res() As Adress '������� � ������������
Dim curDom As String '����� �������� ���� ��� ������

'�������� ���������
Private Sub UserForm_Activate()
    LabelVersion = "������: 1.0 (02.10.2020)"
    On Error GoTo er
    TextBoxTN = Sheets(1).name
    TextBoxHVS = Sheets(2).name
    TextBoxUK = Sheets(3).name
    Exit Sub
er:
    MsgBox ("�� ������� ������")
    End
End Sub

'�����
Private Sub ButtonClose_Click()
    End
End Sub

'������
Private Sub ButtonOK_Click()
    
    ReDim uk(0) As Adress
    ReDim tn(0) As Adress
    ReDim HVS(0) As Adress
    Dim tab1 As String
    Dim tab2 As String
    Dim max1 As Long
    Dim max2 As Long
    Dim tabA As String
    Dim tabRes As String
    
    tab1 = TextBoxTN
    tab2 = TextBoxHVS
    tabA = TextBoxUK
    tabRes = TextBoxReport
    
    Call Misc.Message("����������...")
    
    '������ �������
    Mnth = TextBoxMonth
    Dim aMax As Long
    aMax = 0
    Do While Sheets(tabA).Cells(aMax + 2, 1) <> ""
        aMax = aMax + 1
        ReDim Preserve uk(aMax) As Adress
        uk(aMax).ul = Sheets(tabA).Cells(aMax + 1, 6) '6-7 - ������� �� ���������?
        uk(aMax).dom = Sheets(tabA).Cells(aMax + 1, 7)
    Loop
    
    '����� ������ ������� ������
    max1 = FindMax(tab1)
    max2 = FindMax(tab2)
    
    '������/������� �������������� ��������
    Call Misc.NewTab(tabRes, True)
    Sheets(tabRes).Select
    
    '�����
    Cells(1, 1) = "����� ����������� ������������ �������� �� ������� � �������� ���� �� " + TextBoxMonth
    Range(Cells(1, 1), Cells(1, 10)).Merge
    
    Cells(2, 1) = "�����"
    Range(Cells(2, 1), Cells(2, 4)).Merge
    
    Cells(2, 5) = "������� ����"
    Range(Cells(2, 5), Cells(2, 7)).Merge
    
    Cells(2, 8) = "�������� ����"
    Range(Cells(2, 8), Cells(2, 10)).Merge
    
    Cells(3, 1) = "�����"
    Cells(3, 2) = "� ����"
    Cells(3, 3) = "����� ����"
    Cells(3, 4) = "��������"
    Cells(3, 5) = "����� �����������. ��� (�� � ��)"
    Cells(3, 6) = "����� �����������. ��������"
    Cells(3, 7) = "����� �����������. ����������"
    Cells(3, 8) = "����� �����������. ��� (�� � ��)"
    Cells(3, 9) = "����� �����������. ��������"
    Cells(3, 10) = "����� �����������. ����������"
    
    Range(Cells(1, 1), Cells(3, 10)).HorizontalAlignment = xlCenter
    Range(Cells(3, 1), Cells(3, 10)).WrapText = True
    
    Columns(1).ColumnWidth = 13
    Columns(2).ColumnWidth = 6
    Columns(3).ColumnWidth = 6
    Columns(4).ColumnWidth = 9
    Columns(5).ColumnWidth = 10
    Columns(6).ColumnWidth = 10
    Columns(7).ColumnWidth = 10
    Columns(8).ColumnWidth = 10
    Columns(9).ColumnWidth = 10
    Columns(10).ColumnWidth = 10
    
    '���������� ������
    R = 4
    For a = 1 To aMax
        curDom = uk(a).ul + ", " + LCase(uk(a).dom)
        'If a Mod 50 = 0 Then
        Call Misc.Message("���������� ������... " + CStr(a) + " �� " + CStr(aMax) + " (" + curDom + ")")
        ReDim res(0) As Adress
        Call FindHome(tab1, max1, 1)
        Call FindHome(tab2, max2, 2)
        
        '������� � ����� ������ �� �������� ����
        TN1 = 0
        TN2 = 0
        TN3 = 0
        HVS1 = 0
        HVS2 = 0
        HVS3 = 0
        For i = 1 To UBound(res)
            Cells(R, 1) = res(i).ul
            Cells(R, 2) = res(i).dom
            
            Cells(R, 4) = res(i).kv
            Cells(R, 5) = res(i).VolumeTN1
            Cells(R, 6) = res(i).VolumeTN2
            Cells(R, 7) = res(i).VolumeTN3
            Cells(R, 8) = res(i).VolumeHVS1
            Cells(R, 9) = res(i).VolumeHVS2
            Cells(R, 10) = res(i).VolumeHVS3
            TN1 = TN1 + res(i).VolumeTN1
            TN2 = TN2 + res(i).VolumeTN2
            TN3 = TN3 + res(i).VolumeTN3
            HVS1 = HVS1 + res(i).VolumeHVS1
            HVS2 = HVS2 + res(i).VolumeHVS2
            HVS3 = HVS3 + res(i).VolumeHVS3
            R = R + 1
        Next
        
        '�����, ���� ���-�� ����
        If UBound(res) > 0 Then
            Cells(R, 5) = TN1
            Cells(R, 6) = TN2
            Cells(R, 7) = TN3
            Cells(R, 8) = HVS1
            Cells(R, 9) = HVS2
            Cells(R, 10) = HVS3
            Range(Cells(R, 1), Cells(R, 12)).Font.Bold = True
            Range(Cells(R, 1), Cells(R, 12)).Font.Underline = True
            R = R + 2
        End If
    Next
    
    '�������
    Range(Cells(2, 1), Cells(R - 2, 10)).Borders.Weight = 2
    
    Call Misc.Message("������!")
    End
    
End Sub

'����� ���� �� �������. n - ���, t - ����� ������� (� ����������� �� ����� � �� ��� ������ ����� ����� ���������� ������)
Sub FindHome(name As String, max As Long, t As Integer)
    
    For i = 2 To max
    
        ul = Sheets(name).Cells(i, cUl)
        dom = CStr(Sheets(name).Cells(i, cDom)) + LCase(Sheets(name).Cells(i, cDom + 1))
        kv = Sheets(name).Cells(i, cKv)
        
        If curDom = ul + ", " + dom Then
            
            '����, ���� �� ����� � �����������
            Find = 0
            For j = 1 To UBound(res)
                If res(j).ul = ul And res(j).dom = dom And res(j).kv = kv Then
                    Find = j
                    Exit For
                Else
                End If
            Next
            
            '���� ��� - ��������� ����� ������
            If Find = 0 Then
                Find = UBound(res) + 1
                ReDim Preserve res(Find)
                res(Find).ul = ul
                res(Find).dom = dom
            End If
            
            '��������� ��������� ������ � ������
            res(Find).kv = Sheets(name).Cells(i, cKv)
            If t = 1 Then
                res(Find).VolumeTN1 = Sheets(name).Cells(i, v1)
                res(Find).VolumeTN2 = Sheets(name).Cells(i, v2)
                res(Find).VolumeTN3 = Sheets(name).Cells(i, v3)
            Else
                res(Find).VolumeHVS1 = Sheets(name).Cells(i, v1)
                res(Find).VolumeHVS2 = Sheets(name).Cells(i, v2)
                res(Find).VolumeHVS3 = Sheets(name).Cells(i, v3)
            End If
        
        End If
    Next
End Sub