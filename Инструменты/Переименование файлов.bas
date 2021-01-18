Attribute VB_Name = "Module1"
Sub rename()
    On Error Resume Next
    patch = "s:\бухгалтерия\Бортникова А. С\_____СчетаПДФ\123\"
    For i = 1 To 256
        oldname = patch + Cells(i, 1).Text
        newname = patch + Cells(i, 2).Text
        Name oldname As newname
    Next
End Sub