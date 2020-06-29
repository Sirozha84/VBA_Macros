Attribute VB_Name = "MKD"
Const ResTab = "����"

Sub Start()
    
    Message "������..."
    
    Call Podbor("���", 4)
    Call Podbor("���", 5)
    Call Podbor("����", 6)
    Call Podbor("���", 7)
    Call Podbor("���", 8)
    Call Podbor("���", 9)
    Call Podbor("���", 10)
    Call Podbor("���", 11)
    Call Podbor("���", 12)
    Call Podbor("���", 13)
    Call Podbor("���", 14)
    Call Podbor("���", 15)
    
    Message "������"
    
End Sub

Private Sub Podbor(SrcTab As String, k As Integer)
        
    max = GetMax(SrcTab)
    maxx = GetMax(ResTab)
    For i = 2 To max
        'If Sheets(SrcTab).Cells(i, 1) = 41 Then
        '    MsgBox ("�����")
        'End If
    
        If i Mod 100 = 0 Then Call ProgressBar("���������", i, max)
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
'########## ����� ������� ##########
'###################################

'����� ������������ ������
'inTab - ��� �������
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

'����� ������
'inTab - ��� �������
'str - ������� ������
'max - ������������ ������, ���� ���������� ������ 0
'sort - ��������� �� �� ��� ���� ������ ������������� (��� �������� ������)
Private Function Search(inTab As String, str As String, max, sort As Boolean)
    Search = 0
    If max = 0 Then max = GetMax(inTab)
    If sort Then
        '�����-������ ����� :-)
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

'��������� ���������
'text - ���
'cur - ������� ��������
'all - �����, ���������� ������ over ����
Private Sub ProgressBar(text As String, ByVal cur As Long, ByVal all As Long)
    Application.ScreenUpdating = True
    Application.StatusBar = text + ":" + str(cur) + " ��" + str(all) + _
        " (" + str(Int(cur / all * 100)) + "% )"
        DoEvents
    Application.ScreenUpdating = False
End Sub

'��������� ���������
'test - ���������
Private Sub Message(text As String)
    Application.ScreenUpdating = True
    Application.StatusBar = text
    DoEvents
    Application.ScreenUpdating = False
End Sub



