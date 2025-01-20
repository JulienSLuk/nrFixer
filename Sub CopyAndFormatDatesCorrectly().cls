Sub CopyAndFormatDatesCorrectly()
    Dim wsOutput As Worksheet
    Dim wsLatestNR As Worksheet
    Dim lastRow As Long
    Dim columnsToFormat As Variant
    Dim col As Variant
    Dim rng As Range
    Dim cell As Range
    Dim dateParts() As String
    Dim formattedDate As Date
    
    ' Set worksheets
    Set wsOutput = ThisWorkbook.Sheets("output")
    Set wsLatestNR = ThisWorkbook.Sheets("latestNR")
    
    ' Clear latestNR sheet and copy headers from output
    wsLatestNR.Cells.Clear
    wsOutput.Rows(1).EntireRow.Copy wsLatestNR.Rows(1)
    
    ' Find the last row in the output sheet
    lastRow = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row
    
    ' Copy rows from row 2 onwards from output to latestNR
    If lastRow > 1 Then
        wsOutput.Rows("2:" & lastRow).Copy wsLatestNR.Rows(2)
    End If
    
    ' Define columns with dates to be formatted
    columnsToFormat = Array("K", "L", "U", "V", "X", "AA", "AB", "AF", "AR", "BC")
    
    ' Format dates in latestNR
    For Each col In columnsToFormat
        ' Get range for the column starting from row 2
        Set rng = wsLatestNR.Range(col & "2:" & col & wsLatestNR.Cells(wsLatestNR.Rows.Count, col).End(xlUp).Row)
        
        ' Check if range is valid before formatting
        If Not rng Is Nothing Then
            For Each cell In rng
                If cell.Value <> "" Then
                    ' Split the text by "."
                    dateParts = Split(cell.Value, ".")
                    
                    ' Ensure there are 3 parts to form a valid date (DD.MM.YYYY)
                    If UBound(dateParts) = 2 Then
                        ' Convert to date (Excel-compatible format)
                        On Error Resume Next
                        formattedDate = DateSerial(dateParts(2), dateParts(1), dateParts(0))
                        On Error GoTo 0
                        
                        ' If valid date, update the cell
                        If IsDate(formattedDate) Then
                            cell.Value = formattedDate
                            cell.NumberFormat = "DD/MM/YYYY"
                        Else
                            cell.ClearContents ' Clear invalid date
                        End If
                    Else
                        cell.ClearContents ' Clear invalid data
                    End If
                End If
            Next cell
        End If
    Next col
    
    MsgBox "Data copied to 'latestNR' and dates formatted.", vbInformation
End Sub

