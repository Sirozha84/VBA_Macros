Attribute VB_Name = "NewAndOldFind"
Global sn As Integer
Global max As Integer


Sub NewAndOldFind()

    Call MakeCopy
    
    Call AddNew("������")
    Call AddNew("�����")
    
    Call FindDead("������")
    Call FindDead("�����")
    
    MsgBox ("������")
    
End Sub

Private Sub MakeCopy()
    Application.ScreenUpdating = False
    Sheets("Res").Cells.Clear
    For i = 1 To 99999
        If Sheets("���").Cells(i, 1) <> "" Then
            For j = 1 To 3
                Sheets("Res").Cells(i, j) = Sheets("���").Cells(i, j)
            Next
        Else
            Exit For
        End If
    Next
    max = i
End Sub

Private Sub AddNew(sheet)
    For i = 1 To 99999
        If Sheets(sheet).Cells(i, 1) <> "" Then
            Find = False
            For j = 2 To 99999
                If Sheets("Res").Cells(j, 2) <> "" Then
                    If Right(Sheets("Res").Cells(j, 1), 5) = Right(Sheets(sheet).Cells(i, 1), 5) Then
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
                max = max + 1
                sn = sn + 1
            End If
        Else
            Exit For
        End If
    Next
    Sheets("Res").Cells(max, 4) = "�����:" + Str(sn)
End Sub

Private Sub FindDead(sheet)
    s = 0
    For i = 1 To 99999
        t = Sheets("Res").Cells(i, 1)
        If t = "" Then Exit For
        Find = False
        For j = 1 To 99999
            If Sheets(sheet).Cells(j, 1) <> "" Then
                If Right(Sheets(sheet).Cells(j, 1), 5) = Right(t, 5) Then
                    Find = True
                    Exit For
                End If
            Else
                Exit For
            End If
        Next
        If Not Find And Sheets("Res").Cells(i, 4) = "" Then
            If Sheets("Res").Cells(i, 5) = "-" Then
                Sheets("Res").Cells(i, 5) = "�����!"
                s = s + 1
            Else
                Sheets("Res").Cells(i, 5) = "-"
            End If
        End If
    Next
    Sheets("Res").Cells(i, 5) = "�������:" + Str(s)
End Sub
