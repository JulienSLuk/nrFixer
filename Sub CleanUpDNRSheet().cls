Sub CleanUpDNRSheet()
    Dim wsDNR As Worksheet
    Dim lastRowDNR As Long
     
    ' Set the D.NR sheet
    Set wsDNR = ThisWorkbook.Sheets("D.NR")
    
    ' Confirm with the user before proceeding
    If Not ConfirmCleanup() Then Exit Sub
    
    ' Find the last row in column B of the DNR sheet
    lastRowDNR = wsDNR.Cells(wsDNR.Rows.Count, "B").End(xlUp).Row
    
    ' Clear the data from row 2 onwards (preserving the header)
    If lastRowDNR > 1 Then
        wsDNR.Range("A2:AA" & lastRowDNR).ClearContents
    End If
    
    ' Notify user of completion
    MsgBox "The sheet has been cleaned up.", vbInformation
End Sub



