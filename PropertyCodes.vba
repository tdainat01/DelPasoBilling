Option Compare Database

Private Sub cmdPrint_Click()
   On Error GoTo cmdPrint_Click_Error

    DoCmd.OpenReport "rptPropertyCodes", acViewReport
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrint_Click of VBA Document Form_PropertyCodes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrint_Click of VBA Document Form_PropertyCodes"
End Sub

Private Sub cmdQuit_Click()
   On Error GoTo cmdQuit_Click_Error

    DoCmd.OpenForm "DPM Main Menu", acNormal
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdQuit_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdQuit_Click of VBA Document Form_PropertyCodes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdQuit_Click of VBA Document Form_PropertyCodes"
End Sub
