Option Compare Database
Option Explicit

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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdExit_Click of VBA Document Form_frmChgReportWiz")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdExit_Click of VBA Document Form_frmChgReportWiz"
End Sub

Private Sub cmdGo_Click()
    
    Dim strQuery As String
    Dim rpt As Report
    Dim strPeriod As String

   On Error GoTo cmdGo_Click_Error

    strQuery = "SELECT Sum(Money.amount) AS SumOfamount" & _
            " FROM [Money] WHERE (((Money.code)='CHG') AND ((Money.trans_date) Between " & _
            txtStartDate & " And " & txtEndDate & "));"
    strPeriod = txtStartDate & " to " & txtEndDate
    DoCmd.OpenReport "rptCharging", acViewDesign
    Set rpt = Reports![rptCharging]
    rpt.RecordSource = strQuery
    DoCmd.Close acReport, rpt.Name, acSaveYes
    DoCmd.OpenReport "rptCharging", acViewReport, , , , strPeriod
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdGo_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdGo_Click of VBA Document Form_frmChgReportWiz")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdGo_Click of VBA Document Form_frmChgReportWiz"
    
End Sub

Private Sub Form_Load()

   On Error GoTo Form_Load_Error

Me.txtStartDate = Format(Now, "mm/dd/yyyy")
Me.txtEndDate = Format(Now, "mm/dd/yyyy")

   On Error GoTo 0
   Exit Sub

Form_Load_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Load of VBA Document Form_frmChgReportWiz")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Load of VBA Document Form_frmChgReportWiz"

End Sub
