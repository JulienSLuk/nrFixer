Sub AppendPd()
    Dim wsOutput As Worksheet
    Dim wsPD As Worksheet
    Dim outputNRICs As Object
    Dim lastRowOutput As Long
    Dim lastRowPD As Long
    Dim cell As Range
    Dim pdValueS As String
    Dim pdValueQ As String
    
    ' Set worksheets
    Set wsOutput = ThisWorkbook.Sheets("output")
    Set wsPD = ThisWorkbook.Sheets("pd")
    
    ' Create a dictionary to store NRICs from the output sheet
    Set outputNRICs = CreateObject("Scripting.Dictionary")
    
    ' Find the last row in the output sheet and the pd sheet
    lastRowOutput = wsOutput.Cells(wsOutput.Rows.Count, "B").End(xlUp).Row
    lastRowPD = wsPD.Cells(wsPD.Rows.Count, "J").End(xlUp).Row
    
    ' Store NRICs from the output sheet into the dictionary with row numbers
    For Each cell In wsOutput.Range("B2:B" & lastRowOutput) ' NRIC in Column B of output
        If Not IsEmpty(cell.Value) Then
            outputNRICs(cell.Value) = cell.Row
        End If
    Next cell
    
    ' Loop through each NRIC in the pd sheet and match with the output sheet
    For Each cell In wsPD.Range("J2:J" & lastRowPD) ' NRIC in Column J of pd
        If outputNRICs.exists(cell.Value) Then
            ' Get the value from Column S of pd
            pdValueS = cell.Offset(0, 9).Value ' Column S is 9 columns away from Column J
            
            ' Get the value from Column Q of pd
            pdValueQ = cell.Offset(0, 7).Value ' Column Q is 7 columns away from Column J
            
            ' Write the values to Columns M and N of the output sheet
            wsOutput.Cells(outputNRICs(cell.Value), "M").Value = pdValueS ' Column S to Column M
            wsOutput.Cells(outputNRICs(cell.Value), "N").Value = pdValueQ ' Column Q to Column N
        End If
    Next cell
    
    ' Notify user of completion
    MsgBox "Data from 'pd' has been added to the 'output' sheet.", vbInformation
End Sub

