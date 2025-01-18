Sub CleanUpNRICsInColumnA()
    Dim wsCurrent As Worksheet
    Dim lastRow As Long
    
    ' Set the current worksheet
    Set wsCurrent = ThisWorkbook.ActiveSheet
    
    ' Confirm with the user before proceeding
    If Not ConfirmCleanup() Then Exit Sub
    
    ' Find the last row in column A
    lastRow = wsCurrent.Cells(wsCurrent.Rows.Count, "A").End(xlUp).Row
    
    ' Clear the contents of column A from row 2 to the last row with data (leaving the header intact)
    If lastRow > 1 Then
        wsCurrent.Range("A2:A" & lastRow).ClearContents
    End If
    
    ' Notify the user that the NRICs have been cleaned up
    MsgBox "NRICs in column A have been cleaned up.", vbInformation
End Sub

