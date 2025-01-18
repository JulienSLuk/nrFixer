Sub CleanUpOutputSheet()
    Dim wsOutput As Worksheet
    Dim lastRowOutput As Long
     
    ' Set the output sheet
    Set wsOutput = ThisWorkbook.Sheets("output")
    
    ' Confirm with the user before proceeding
    If Not ConfirmCleanup() Then Exit Sub
    
    ' Find the last row in column A of the output sheet
    lastRowOutput = wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).Row
    
    ' Clear the data from row 2 onwards (preserving the header)
    If lastRowOutput > 1 Then
        wsOutput.Range("A2:AA" & lastRowOutput).ClearContents
    End If
    
    ' Notify user of completion
    MsgBox "The sheet has been cleaned up.", vbInformation
End Sub

