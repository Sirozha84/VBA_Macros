Attribute VB_Name = "Starter"
Type Record
    Name As String
    Res As String
    Number As Double
    Date As String
    VolumeHeat As Double
    VolumeHW As Double
    PriceHeat As Double
    PriceHW As Double
    VolumeInfo As Double
End Type

Public Sub Start()
    FormReport.Show
End Sub