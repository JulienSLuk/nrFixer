Sub OutputMissingNRICsWithData()
    Dim wsCurrent As Worksheet
    Dim wsRawNR As Worksheet
    Dim wsOutput As Worksheet
    Dim currentNRICs As Object
    Dim lastRowCurrent As Long
    Dim lastRowRawNR As Long
    Dim lastRowOutput As Long
    Dim cell As Range
    Dim outputRow As Long
    
    ' Set worksheets
    Set wsCurrent = ThisWorkbook.ActiveSheet
    Set wsRawNR = ThisWorkbook.Sheets("rawNR")
    Set wsOutput = ThisWorkbook.Sheets("output")
    
    ' Create a dictionary to store NRICs from column A of the current sheet (missing NRICs)
    Set currentNRICs = CreateObject("Scripting.Dictionary")
    
    ' Find the last row in column C of the current sheet (missing NRICs)
    lastRowCurrent = wsCurrent.Cells(wsCurrent.Rows.Count, "C").End(xlUp).Row
    
    ' Store the missing NRICs from Column C of the current sheet into the dictionary
    For Each cell In wsCurrent.Range("C2:C" & lastRowCurrent)
        If Not IsEmpty(cell.Value) Then
            currentNRICs(cell.Value) = True
        End If
    Next cell
    
    ' Find the last row in column B (NRIC) of the rawNR sheet
    lastRowRawNR = wsRawNR.Cells(wsRawNR.Rows.Count, "B").End(xlUp).Row
    
    ' Find the last row in the output sheet to append data
    lastRowOutput = wsOutput.Cells(wsOutput.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Loop through each NRIC in the rawNR sheet and check if it's missing (exists in currentNRICs)
    For Each cell In wsRawNR.Range("B2:B" & lastRowRawNR)
        If currentNRICs.exists(cell.Value) Then
            ' Copy the data to the output sheet
            wsOutput.Cells(lastRowOutput, "A").Value = cell.Offset(0, -1).Value   ' PERNR
            wsOutput.Cells(lastRowOutput, "B").Value = cell.Value                ' NRIC
            wsOutput.Cells(lastRowOutput, "C").Value = cell.Offset(0, 3).Value   ' RANK
            wsOutput.Cells(lastRowOutput, "D").Value = cell.Offset(0, 1).Value   ' NAME
            wsOutput.Cells(lastRowOutput, "K").Value = cell.Offset(0, 4).Value   ' DOB
            wsOutput.Cells(lastRowOutput, "L").Value = cell.Offset(0, 5).Value   ' DOI
            lastRowOutput = lastRowOutput + 1
        End If
    Next cell
    
    ' Notify user of completion
    MsgBox "Missing NRICs and their data have been output to the 'output' sheet.", vbInformation
End Sub

