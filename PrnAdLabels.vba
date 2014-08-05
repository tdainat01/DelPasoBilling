Option Compare Database

Private Sub cmdExit_Click()
   On Error GoTo cmdExit_Click_Error

    DoCmd.OpenForm "DPM Main Menu", acNormal
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdExit_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdExit_Click of VBA Document Form_PrnAdLabels")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdExit_Click of VBA Document Form_PrnAdLabels"
End Sub

Private Sub cmdGatherAddresses_Click()
   On Error GoTo cmdGatherAddresses_Click_Error

    MsgBox "Addresses have been gathered", vbOKOnly + vbInformation, "Done"

   On Error GoTo 0
   Exit Sub

cmdGatherAddresses_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdGatherAddresses_Click of VBA Document Form_PrnAdLabels")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdGatherAddresses_Click of VBA Document Form_PrnAdLabels"
End Sub

Private Sub cmdPrint_Click()
    'DoCmd.OpenForm "frmPrintLabelsAnswers", acNormal
   On Error GoTo cmdPrint_Click_Error

    DoCmd.OpenForm "PrintMailing", acNormal
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdPrint_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrint_Click of VBA Document Form_PrnAdLabels")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrint_Click of VBA Document Form_PrnAdLabels"
End Sub
