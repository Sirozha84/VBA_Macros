Attribute VB_Name = "Calculate"
'�����
Const Street = 3
Const House = 4
Const Letter = 5

Const Area = 9      '�������
Const HReg = 10     '�����������������
Const HOwner = 11   '������������
Const PObject = 30  '������� �������

'�������� �������
Const MKD = "���"
Const IZH = "���"
Dim MKDHouse As Variant
Dim Kvart As Long
Dim MKDArea As Double
Dim MKDHumans As Long
Dim IZHHouse As Variant
Dim IZHArea As Double
Dim IZHHumans As Long

Sub Calculate()
    Set MKDHouse = CreateObject("Scripting.Dictionary")
    Set IZHHouse = CreateObject("Scripting.Dictionary")
    Kvart = 0
    MKDArea = 0
    MKDHumans = 0
    IZHArea = 0
    IZHHumans = 0
    i = 2
    Do While Cells(i, Street) <> "" ' And i < 10
        hous = Cells(i, Street).Text + Cells(i, House).Text + Cells(i, Letter).Text
        If Cells(i, PObject) = MKD Then
            MKDHouse(hous) = 1
            Kvart = Kvart + 1
            MKDArea = MKDArea + Cells(i, Area)
            MKDHumans = MKDHumans + Cells(i, HReg)
            If Cells(i, HReg) = 0 Then MKDHumans = MKDHumans + Cells(i, HOwner)
        End If
        If Cells(i, PObject) = IZH Then
            IZHHouse(hous) = 1
            IZHArea = IZHArea + Cells(i, Area)
            IZHHumans = IZHHumans + Cells(i, HReg)
            If Cells(i, HReg) = 0 Then IZHHumans = IZHHumans + Cells(i, HOwner)
        End If
        i = i + 1
    Loop
    e = Chr(10)
    MsgBox "������ ��������!" + e + _
            e + "���, ���������� ����� = " + CStr(MKDHouse.Count) + _
            e + "���, ���������� ������� = " + CStr(Kvart) + _
            e + "���, ������� = " + CStr(MKDArea) + _
            e + "���, ������� = " + CStr(MKDHumans) + e + _
            e + "���, ���������� ����� = " + CStr(IZHHouse.Count) + _
            e + "���, ������� = " + CStr(IZHArea) + _
            e + "���, ������� = " + CStr(IZHHumans)
End Sub