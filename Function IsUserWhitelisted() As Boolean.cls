Function IsUserWhitelisted() As Boolean
    Dim wsWhitelist As Worksheet
    Dim userName As String
    Dim lastRowWhitelist As Long
    Dim cell As Range
    
    ' Set the whitelist sheet
    Set wsWhitelist = ThisWorkbook.Sheets("whitelist")
    
    ' Prompt user to enter their name
    userName = InputBox("Please enter your name to access the macros:", "Whitelister")
    
    ' Exit if the user cancels or leaves the input blank
    If userName = "" Then
        MsgBox "You must enter a name.", vbExclamation
        IsUserWhitelisted = False
        Exit Function
    End If
    
    ' Find the last row in column A of the whitelist sheet
    lastRowWhitelist = wsWhitelist.Cells(wsWhitelist.Rows.Count, "A").End(xlUp).Row
    
    ' Check if the entered name exists in the whitelist
    For Each cell In wsWhitelist.Range("A2:A" & lastRowWhitelist) ' Assuming names start from A2
        If StrComp(cell.Value, userName, vbTextCompare) = 0 Then
            IsUserWhitelisted = True
            Exit Function
        End If
    Next cell
    
    ' If name is not found in the whitelist
    MsgBox "Your name is not in the whitelist. You are not authorized to run the macros.", vbCritical
    IsUserWhitelisted = False
End Function
