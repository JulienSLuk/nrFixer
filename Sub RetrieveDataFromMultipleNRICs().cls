Sub RetrieveDataFromMultipleNRICs()
    Dim wsDNR As Worksheet
    Dim wsOutput As Worksheet
    Dim nricInput As String
    Dim persRow As Long
    Dim outputRow As Long
    Dim lastOutputRow As Long
    Dim outputRng As Range
    Dim field As String
    Dim column As Long
    Dim nricRange As Range
    Dim nricCell As Range
    Dim lastColumn As Long
    Dim outputColumn As Variant ' Changed to Variant to handle errors from Match
    
    ' Set the current worksheet (D.NR) and output sheet
    Set wsDNR = ThisWorkbook.Sheets("D.NR")
    Set wsOutput = ThisWorkbook.Sheets("output")
    
    ' Find the last row with NRIC data in column B of D.NR
    lastRow = wsDNR.Cells(wsDNR.Rows.Count, "B").End(xlUp).Row
    Set nricRange = wsDNR.Range("B2:B" & lastRow) ' Range of NRICs
    
    ' Loop through each NRIC in column B of D.NR
    For Each nricCell In nricRange
        nricInput = nricCell.Value ' NRIC input is in column B
        
        ' If no NRIC is entered, skip this row
        If nricInput = "" Then
            GoTo NextNRIC
        End If
        
        ' Find the row in the output sheet that matches the NRIC (Column B in output)
        lastOutputRow = wsOutput.Cells(wsOutput.Rows.Count, "B").End(xlUp).Row
        Set outputRng = wsOutput.Range("B2:B" & lastOutputRow)
        
        ' Look for the NRIC in the output sheet
        outputRow = 0 ' Default value if NRIC is not found
        For i = 2 To lastOutputRow
            If wsOutput.Cells(i, 2).Value = nricInput Then ' Check column B for NRIC
                outputRow = i
                Exit For
            End If
        Next i
        
        ' If NRIC is not found, skip to next NRIC
        If outputRow = 0 Then
            GoTo NextNRIC
        End If
        
        ' Find the next available row in D.NR (using the current row in nricCell)
        persRow = nricCell.Row
        
        ' Retrieve PERS, RANK, and NAME from output sheet to D.NR
        wsDNR.Cells(persRow, 1).Value = wsOutput.Cells(outputRow, 1).Value ' PERNR (Column A in output to A in D.NR)
        wsDNR.Cells(persRow, 3).Value = wsOutput.Cells(outputRow, 3).Value ' RANK (Column C in output to C in D.NR)
        wsDNR.Cells(persRow, 4).Value = wsOutput.Cells(outputRow, 4).Value ' NAME (Column D in output to D in D.NR)
        
        ' Now handle additional fields dynamically (columns after NAME in D.NR)
        lastColumn = wsDNR.Cells(1, wsDNR.Columns.Count).End(xlToLeft).column ' Find last column in row 1 (headers)
        
        ' Loop through the additional columns (E onwards)
        For column = 5 To lastColumn
            field = wsDNR.Cells(1, column).Value ' Read field name from row 1
            
            ' If the header is not empty, retrieve the corresponding data
            If field <> "" Then
                ' Check the field name and map to the correct output column
                Select Case LCase(field)
                    Case "eligibility"
                        outputColumn = 13 ' Column M in output for Eligibility
                    Case "performance"
                        outputColumn = 14 ' Column N in output for Performance
                    Case Else
                        ' If it's a new field, dynamically match it to the corresponding output column
                        outputColumn = Application.Match(field, wsOutput.Rows(1), 0)
                        
                        ' If the Match returns an error, set outputColumn to a default invalid value
                        If IsError(outputColumn) Then
                            outputColumn = -1
                        End If
                End Select
                
                ' If we have a valid output column (not -1), get the value, else clear the contents
                If outputColumn <> -1 Then
                    wsDNR.Cells(persRow, column).Value = wsOutput.Cells(outputRow, outputColumn).Value
                Else
                    wsDNR.Cells(persRow, column).ClearContents ' Clear if header is removed
                End If
            End If
        Next column
        
NextNRIC:
    Next nricCell
End Sub

