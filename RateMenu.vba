Option Compare Database

Private Sub chkDev_Click()
   On Error GoTo chkDev_Click_Error

If chkDev.value Then
    cmdReset.Visible = True
    cmdSimPayment.Visible = True
    cmdFixPayment.Visible = True
Else
    cmdReset.Visible = False
    cmdSimPayment.Visible = False
    cmdFixPayment.Visible = False
End If

   On Error GoTo 0
   Exit Sub

chkDev_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure chkDev_Click of VBA Document Form_RateMenu")
    'MsgBox "Error " & errNum & " (" & errMsg & ") in procedure chkDev_Click of VBA Document Form_RateMenu"

End Sub

Private Sub cmdFixPayment_Click()

   On Error GoTo cmdFixPayment_Click_Error

    DoCmd.SetWarnings False
    CurrentDb.Execute "UPDATE [Money] SET [Money].code = 'MON' WHERE (((Money.code)='PMT'));", dbFailOnError
    CurrentDb.Execute "UPDATE [Money] SET [Money].code = 'MON' WHERE (((Money.code)='MOM'));", dbFailOnError

    MsgBox "Completed with cmdFixPayment. Codes updated", vbOKOnly, "Done"

    DoCmd.SetWarnings True
    
   On Error GoTo 0
   Exit Sub

cmdFixPayment_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdFixPayment_Click of VBA Document Form_RateMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdFixPayment_Click of VBA Document Form_RateMenu"

End Sub

Private Sub cmdLoadLot_Click()
   On Error GoTo cmdLoadLot_Click_Error

    DoCmd.OpenForm "PropertyCodes", acNormal
    DoCmd.Close acForm, Me.Name, acSaveYes
    'Call MsgBox("This procedure still needs to be done.", vbInformation, "Do Be Done")

   On Error GoTo 0
   Exit Sub

cmdLoadLot_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdLoadLot_Click of VBA Document Form_RateMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdLoadLot_Click of VBA Document Form_RateMenu"
End Sub

Private Sub cmdMaintRate_Click()
   On Error GoTo cmdMaintRate_Click_Error

    DoCmd.OpenForm "frmServiceConnections", acNormal
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdMaintRate_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdMaintRate_Click of VBA Document Form_RateMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdMaintRate_Click of VBA Document Form_RateMenu"
End Sub

Private Sub cmdPrintRate_Click()
   On Error GoTo cmdPrintRate_Click_Error

    DoCmd.OpenReport "rptServiceConnections", acViewReport
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdPrintRate_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrintRate_Click of VBA Document Form_RateMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrintRate_Click of VBA Document Form_RateMenu"
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdQuit_Click of VBA Document Form_RateMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdQuit_Click of VBA Document Form_RateMenu"
End Sub

Private Sub cmdRecurringCharges_Click()
   On Error GoTo cmdRecurringCharges_Click_Error

    DoCmd.OpenForm "frmMaintRecurringCharges", acNormal

   On Error GoTo 0
   Exit Sub

cmdRecurringCharges_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdRecurringCharges_Click of VBA Document Form_RateMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdRecurringCharges_Click of VBA Document Form_RateMenu"
    
End Sub

Private Sub cmdReset_Click()

Dim result As Variant

   On Error GoTo cmdReset_Click_Error

Select Case MsgBox("Are you sure you want to do that?", vbYesNoCancel Or vbExclamation Or vbDefaultButton2, "Verify")

    Case vbYes
        Call ResetAccounts
    Case vbNo
        Exit Sub
    Case vbCancel
        Exit Sub
End Select

    Call MsgBox("All accounts have been reset to zero. Everything in the Money Table has been deleted as well.", _
        vbInformation Or vbDefaultButton1, "Done")
    
   On Error GoTo 0
   Exit Sub

cmdReset_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdReset_Click of VBA Document Form_RateMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdReset_Click of VBA Document Form_RateMenu"

End Sub
Private Sub ResetAccounts()
    
    Dim sQuery As String
    
   On Error GoTo ResetAccounts_Error

    DoCmd.SetWarnings False
    
    sQuery = "UPDATE customer SET customer.deposit = 0, customer.use_charge = 0, customer.past_due = 0, " & _
            " customer.prev_balance = 0, customer.current_due = 0, customer.special_credit = 0, " & _
            " customer.total_due = 0, customer.special_charge = 0 " '& _ " WHERE (((customer.rate_code) Not In ('FP','AA')));"

    DoCmd.RunSQL sQuery

    sQuery = "DELETE * FROM [Money]"

    DoCmd.RunSQL sQuery

    DoCmd.SetWarnings True
   On Error GoTo 0
   Exit Sub

ResetAccounts_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure ResetAccounts of VBA Document Form_RateMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure ResetAccounts of VBA Document Form_RateMenu"
    
End Sub

Private Sub cmdSimPayment_Click()
Call MsgBox("I'm sorry - This feature has not been implemented yet", vbInformation, "Not implemented")

End Sub

Private Sub Form_Load()
Me.chkDev.value = False
Me.cmdReset.Visible = False
Me.cmdSimPayment.Visible = False
End Sub
