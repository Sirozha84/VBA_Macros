VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Form 
   Caption         =   "����� �������"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6750
   OleObjectBlob   =   "Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
    TextBoxTab1 = Sheets(1).name
    TextBoxTab2 = Sheets(2).name
End Sub

Private Sub CheckBoxCompare_Click()
    TextBoxCompare.Enabled = CheckBoxCompare.Value
End Sub

Private Sub CommandButtonRun_Click()
    CommandButtonRun.Enabled = False
    DoEvents
    On Error GoTo Error
    Call PrepareResult(TextBoxTabRes, CheckBoxCreate.Value)
    LabelStatus = "���������..."
    
Error:
    CommandButtonRun.Enabled = True
End Sub

Private Sub CommandButtonExit_Click()
    End
End Sub

'�������� �������������� �������
Private Sub PrepareResult(name As String, create As Boolean)
    If create Then
        If Not SheetExist(name) Then
            Sheets.Add(, Sheets(Sheets.Count)).name = name
        End If
    End If
    If Not SheetExist(name) Then
        LabelStatus = "������: �������������� ������� �� ����������"
        Err.Raise (0)
    End If
    Sheets(name).Cells.Clear
End Sub

'�������� �� ������������� �����
Function SheetExist(name As String) As Boolean
    Dim objSheet As Object
    
    On Error GoTo HandleError
    ThisWorkbook.Worksheets(name).Activate
    SheetExist = True
    Exit Function
    
HandleError:
    SheetExist = False
End Function
�Lp/    P�<N��5m�s'%  ����\!��i/    ���i(G���l�   �����h1��b1    ���l�lK��/H�'�   ����