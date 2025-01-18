Function ConfirmCleanup() As Boolean
    Dim userResponse As VbMsgBoxResult
    
    ' Prompt the user for confirmation
    userResponse = MsgBox("Do you want to clean up? This will clear all data except headers.", _
                          vbYesNo + vbQuestion, "Confirm Cleanup")
    
    ' Return True if the user clicks Yes, otherwise return False
    ConfirmCleanup = (userResponse = vbYes)
End Function
