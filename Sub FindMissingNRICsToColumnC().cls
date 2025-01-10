Sub FindMissingNRICsToColumnC()
    Dim wsCurrent As Worksheet
    Dim wsRawNR As Worksheet
    Dim currentNRICs As Object
    Dim cell As Range
    Dim lastRowCurrent As Long
    Dim lastRowRawNR As Long
    Dim outputRow As Long
    
    ' Set worksheets
    Set wsCurrent = ThisWorkbook.ActiveSheet
    Set wsRawNR = ThisWorkbook.Sheets("rawNR")
    
    ' Create a dictionary to store NRICs from column A of the current sheet
    Set currentNRICs = CreateObject("Scripting.Dictionary")
    
    ' Clear previous data in Column C
    wsCurrent.Columns("C").ClearContents
    
    ' Find the last row in column A of the current sheet
    lastRowCurrent = wsCurrent.Cells(wsCurrent.Rows.Count, "A").End(xlUp).Row
    
    ' Store NRICs from column A of the current sheet into the dictionary
    For Each cell In wsCurrent.Range("A2:A" & lastRowCurrent)
        If Not IsEmpty(cell.Value) Then
            currentNRICs(cell.Value) = True
        End If
    Next cell
    
    ' Find the last row in column A of the rawNR sheet
    lastRowRawNR = wsRawNR.Cells(wsRawNR.Rows.Count, "A").End(xlUp).Row
    
    ' Start outputting missing NRICs to column C, starting at row 2
    outputRow = 2
    
    ' Check NRICs in rawNR sheet but not in the dictionary
    For Each cell In wsRawNR.Range("A2:A" & lastRowRawNR)
        If Not currentNRICs.exists(cell.Value) Then
            wsCurrent.Cells(outputRow, "C").Value = cell.Value
            outputRow = outputRow + 1
        End If
    Next cell
    
    ' Notify user of completion
    MsgBox "Missing NRICs have been output to column C.", vbInformation
End Sub

