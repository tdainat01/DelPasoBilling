Option Compare Database

Private Sub cmdCreditAccts_Click()

   On Error GoTo cmdCreditAccts_Click_Error
    DoCmd.OpenReport "rptCreditAccounts", acViewReport
    DoCmd.Close acForm, Me.Name, acSaveYes
    

   On Error GoTo 0
   Exit Sub

cmdCreditAccts_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdCreditAccts_Click of VBA Document Form_ReportMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdCreditAccts_Click of VBA Document Form_ReportMenu"
End Sub

Private Sub cmdDelinquentReport_Click()
   On Error GoTo cmdDelinquentReport_Click_Error

    DoCmd.OpenReport "rptDelinquent", acViewReport
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdDelinquentReport_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdDelinquentReport_Click of VBA Document Form_ReportMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdDelinquentReport_Click of VBA Document Form_ReportMenu"
End Sub

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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdExit_Click of VBA Document Form_ReportMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdExit_Click of VBA Document Form_ReportMenu"
End Sub

Private Sub cmdFinancialReport_Click()
   On Error GoTo cmdFinancialReport_Click_Error

    DoCmd.OpenForm "frmFinReportWiz", acNormal
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdFinancialReport_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdFinancialReport_Click of VBA Document Form_ReportMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdFinancialReport_Click of VBA Document Form_ReportMenu"
End Sub

Private Sub cmdLabelsMenu_Click()
   On Error GoTo cmdLabelsMenu_Click_Error

    DoCmd.OpenForm "PrintLabelsMenu", acNormal, , , , , Me.Name
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdLabelsMenu_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdLabelsMenu_Click of VBA Document Form_ReportMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdLabelsMenu_Click of VBA Document Form_ReportMenu"
End Sub

Private Sub cmdOutAcctRpt_Click()
   On Error GoTo cmdOutAcctRpt_Click_Error

    DoCmd.OpenReport "rptOutstanding", acViewReport
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdOutAcctRpt_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdOutAcctRpt_Click of VBA Document Form_ReportMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdOutAcctRpt_Click of VBA Document Form_ReportMenu"
End Sub

Private Sub cmdOutstandingList_Click()

   On Error GoTo cmdOutstandingList_Click_Error
    DoCmd.OpenReport "rptOutstandingBal", acViewReport
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdOutstandingList_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdOutstandingList_Click of VBA Document Form_ReportMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdOutstandingList_Click of VBA Document Form_ReportMenu"
End Sub

Private Sub cmdPrintRoll_Click()

   On Error GoTo cmdPrintRoll_Click_Error

    DoCmd.OpenReport "rptPrintRoll", acViewReport

   On Error GoTo 0
   Exit Sub

cmdPrintRoll_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrintRoll_Click of VBA Document Form_ReportMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrintRoll_Click of VBA Document Form_ReportMenu"
End Sub

Private Sub cmdPropertyUse_Click()

   On Error GoTo cmdPropertyUse_Click_Error
    DoCmd.OpenReport "rptCountPropertyUse", acViewReport
    DoCmd.Close acForm, Me.Name, acSaveYes
    

   On Error GoTo 0
   Exit Sub

cmdPropertyUse_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPropertyUse_Click of VBA Document Form_ReportMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPropertyUse_Click of VBA Document Form_ReportMenu"
End Sub

Private Sub cmdQryTrans_Click()
   On Error GoTo cmdQryTrans_Click_Error

    DoCmd.OpenQuery "qryAcctTrans", acViewNormal, acReadOnly

   On Error GoTo 0
   Exit Sub

cmdQryTrans_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdQryTrans_Click of VBA Document Form_ReportMenu")
    'MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdQryTrans_Click of VBA Document Form_ReportMenu"

End Sub

Private Sub cmdRateCodeCounts_Click()

   On Error GoTo cmdRateCodeCounts_Click_Error
    DoCmd.OpenReport "rptRateCodeCounts", acViewReport
    DoCmd.Close acForm, Me.Name, acSaveYes
    
   On Error GoTo 0
   Exit Sub

cmdRateCodeCounts_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdRateCodeCounts_Click of VBA Document Form_ReportMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdRateCodeCounts_Click of VBA Document Form_ReportMenu"
End Sub

Private Sub cmdSideBySide_Click()

   On Error GoTo cmdSideBySide_Click_Error

    

   On Error GoTo 0
   Exit Sub

cmdSideBySide_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    'Call LogError(errNum, errSource, errMsg & " in procedure cmdSideBySide_Click of VBA Document Form_ReportMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdSideBySide_Click of VBA Document Form_ReportMenu"
End Sub

Private Sub cmdTotalCharged_Click()
   On Error GoTo cmdTotalCharged_Click_Error

    'DoCmd.OpenForm "frmSelectAccounts", acNormal
    DoCmd.OpenForm "frmSelDate", acNormal
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdTotalCharged_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdTotalCharged_Click of VBA Document Form_ReportMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdTotalCharged_Click of VBA Document Form_ReportMenu"
End Sub

Private Sub cmdUnknownOwners_Click()

   On Error GoTo cmdUnknownOwners_Click_Error
    DoCmd.OpenReport "rptUnknownOwners", acViewPreview, , , acWindowNormal
    
    

   On Error GoTo 0
   Exit Sub

cmdUnknownOwners_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdUnknownOwners_Click of VBA Document Form_ReportMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdUnknownOwners_Click of VBA Document Form_ReportMenu"
End Sub

Private Sub cmdWaterOffRpt_Click()

   On Error GoTo cmdWaterOffRpt_Click_Error
    DoCmd.OpenReport "rptWaterOff", acViewPreview, , , acWindowNormal
    

   On Error GoTo 0
   Exit Sub

cmdWaterOffRpt_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdWaterOffRpt_Click of VBA Document Form_ReportMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdWaterOffRpt_Click of VBA Document Form_ReportMenu"
End Sub
