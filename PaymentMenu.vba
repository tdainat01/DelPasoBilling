Option Compare Database

Private Sub cmdPayEdit_Click()
   On Error GoTo cmdPayEdit_Click_Error

    DoCmd.OpenForm "PrintMenu", acNormal, , , , , 1
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdPayEdit_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPayEdit_Click of VBA Document Form_PaymentMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPayEdit_Click of VBA Document Form_PaymentMenu"
End Sub

Private Sub cmdPayEditLine_Click()
   On Error GoTo cmdPayEditLine_Click_Error

    DoCmd.OpenForm "PrintMenu", acNormal, , , , , 2
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdPayEditLine_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPayEditLine_Click of VBA Document Form_PaymentMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPayEditLine_Click of VBA Document Form_PaymentMenu"
End Sub

Private Sub cmdQuit_Click()
   On Error GoTo cmdQuit_Click_Error

    DoCmd.Close acForm, Me.Form.Name, acSaveYes
    DoCmd.OpenForm "DPM Main Menu"

   On Error GoTo 0
   Exit Sub

cmdQuit_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdQuit_Click of VBA Document Form_PaymentMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdQuit_Click of VBA Document Form_PaymentMenu"
End Sub
