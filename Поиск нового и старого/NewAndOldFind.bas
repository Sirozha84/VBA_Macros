Attribute VB_Name = "NewAndOldFind"
Sub NewAndOldFind()

    Call MakeCopy
    
    'Call AddNew ("������")
    'Call AddNew ("�����")
    
    Call FindDead("������")
    Call FindDead("�����")
    
    MsgBox ("������")
    
End Sub

Private Sub MakeCopy()
    Application.ScreenUpdating = False
    Sheets("Res").Cells.Clear
    For i = 2 To 99999
        If Sheets("���").Cells(i, 1) <> "" Then
            For j = 1 To 3
                Sheets("Res").Cells(i, j) = Sheets("���").Cells(i, j)
            Next
        Else
            Exit For
        End If
    Next
End Sub

Private Sub AddNew(sheet)
    For i = 2 To 99999
        If Sheets(sheet).Cells(i, 1) <> "" Then
            Find = False
            For j = 2 To 99999
                If Sheets("Res").Cells(j, 2) <> "" Then
                    If Sheets("Res").Cells(j, 1) = Sheets(sheet).Cells(i, 1) Then
                        Find = True
                    End If
                Else
                    Exit For
                End If
            Next
            If Not Find Then
                Sheets("Res").Cells(j, 1) = Sheets(sheet).Cells(i, 1)
                Sheets("Res").Cells(j, 2) = Sheets(sheet).Cells(i, 2)
                Sheets("Res").Cells(j, 4) = "����� �� " + sheet
            End If
        Else
            Exit For
        End If
    Next
End Sub

Private Sub FindDead(sheet)
    
    For i = 2 To 99999
        t = Sheets("Res").Cells(i, 1)
        If t = "" Then Exit For
        Find = False
        For j = 1 To 99999
            If Sheets(sheet).Cells(j, 1) <> "" Then
                If Sheets(sheet).Cells(j, 1) = t Then
                    Find = True
                    Exit For
                End If
            Else
                Exit For
            End If
        Next
        If Not Find Then
            If Sheets("Res").Cells(i, 4) = "-" Then
                Sheets("Res").Cells(i, 4) = "�����!"
            Else
                Sheets("Res").Cells(i, 4) = "-"
            End If
        End If
    Next

End Sub
