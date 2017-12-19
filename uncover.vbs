Sub uncover()
    Dim sht As Worksheet
    For Each sht In Worksheets
           sht.Visible = xlSheetVisible
    Next
End Sub
