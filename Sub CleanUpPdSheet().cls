Sub CleanUpPdSheet()
    Dim wsPD As Worksheet
    Dim lastRowPD As Long
     
    ' Set the pd sheet
    Set wsPD = ThisWorkbook.Sheets("pd")
    
    ' Confirm with the user before proceeding
    If Not ConfirmCleanup() Then Exit Sub
    
    ' Find the last row in column J of the pd sheet
    lastRowPD = wsPD.Cells(wsPD.Rows.Count, "J").End(xlUp).Row
    
    ' Clear the data from row 2 onwards (preserving the header)
    If lastRowPD > 1 Then
        wsPD.Range("A2:AA" & lastRowPD).ClearContents
    End If
    
    ' Notify user of completion
    MsgBox "The sheet has been cleaned up.", vbInformation
End Sub



