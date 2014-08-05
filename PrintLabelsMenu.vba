Option Compare Database

Private Sub cmdExit_Click()
   On Error GoTo cmdExit_Click_Error

    DoCmd.OpenForm "ReportMenu", acNormal
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdExit_Click of VBA Document Form_PrintLabelsMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdExit_Click of VBA Document Form_PrintLabelsMenu"
End Sub

Private Sub cmdPrintLabels_Click()
   On Error GoTo cmdPrintLabels_Click_Error

    'DoCmd.OpenForm "PrnAdLabels", acNormal
    DoCmd.OpenForm "PrintMailing", acNormal, , , , , Me.Name
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdPrintLabels_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrintLabels_Click of VBA Document Form_PrintLabelsMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrintLabels_Click of VBA Document Form_PrintLabelsMenu"
End Sub
