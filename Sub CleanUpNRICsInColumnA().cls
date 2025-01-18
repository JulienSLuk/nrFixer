Sub CleanUpNRICsInColumnA()
    Dim wsCurrent As Worksheet
    Dim lastRow As Long
    
    ' Set the current worksheet
    Set wsCurrent = ThisWorkbook.ActiveSheet
    
    ' Find the last row in column A
    lastRow = wsCurrent.Cells(wsCurrent.Rows.Count, "A").End(xlUp).Row
    
    ' Clear the contents of column A from row 2 to the last row with data
    wsCurrent.Range("A2:A" & lastRow).ClearContents
    
    ' Notify the user that the NRICs have been cleared
    MsgBox "NRICs in column A have been cleaned up.", vbInformation
End Sub

