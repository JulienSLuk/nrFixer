Sub FormatBuildingNumbers()
    Dim cell As Range
    Dim lastRow As Long
    Dim value As Variant
    Dim ws As Worksheet

    ' Set reference to the "latestNR" sheet
    Set ws = ThisWorkbook.Sheets("latestNR")

    ' Get the last row with data in column E of the "latestNR" sheet
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    ' Loop through all the cells in column E of the "latestNR" sheet
    For Each cell In ws.Range("E2:E" & lastRow)
        ' Store the value of the cell in a variable
        value = cell.value
        
        ' Check if the value is a date or can be converted to a date
        If IsDate(value) Then
            ' Force the output as text in "d-m" format (e.g., 9-2) by converting to string
            cell.value = "'" & CStr(Day(value)) & "-" & CStr(Month(value))
        End If
    Next cell

    ' Display message box to notify user that the process is complete
    MsgBox "Date formatting complete!", vbInformation
End Sub

