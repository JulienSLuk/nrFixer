Sub CleanUpRawNRSheet()
    Dim wsRawNR As Worksheet
    Dim lastRowRawNR As Long
     
    ' Set the rawNR sheet
    Set wsRawNR = ThisWorkbook.Sheets("rawNR")
    
    ' Confirm with the user before proceeding
    If Not ConfirmCleanup() Then Exit Sub
    
    ' Find the last row in column B of the rawNR sheet
    lastRowRawNR = wsRawNR.Cells(wsRawNR.Rows.Count, "B").End(xlUp).Row
    
    ' Clear the data from row 2 onwards (preserving the header)
    If lastRowRawNR > 1 Then
        wsRawNR.Range("A2:AA" & lastRowRawNR).ClearContents
    End If
    
    ' Notify user of completion
    MsgBox "The sheet has been cleaned up.", vbInformation
End Sub


