Option Compare Database
Option Explicit

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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdExit_Click of VBA Document Form_frmFinReportWiz")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdExit_Click of VBA Document Form_frmFinReportWiz"
End Sub

Private Sub cmdGo_Click()
   On Error GoTo cmdGo_Click_Error

    'GoTo LastLine
    Dim strQuery As String
    Dim rpt As Report
    Dim strPeriod As String

    strQuery = "SELECT Sum(Money.amount) AS SumOfamount" & _
            " FROM [Money] WHERE (((Money.code)='SCH') AND ((Money.trans_date) Between #" & _
            txtStartDate & "# And #" & txtEndDate & "#));"
    strPeriod = txtStartDate & " to " & txtEndDate
    DoCmd.OpenReport "rptCharging", acViewDesign
    Set rpt = Reports![rptCharging]
    rpt.RecordSource = strQuery
    DoCmd.Close acReport, rpt.Name, acSaveYes
    DoCmd.OpenReport "rptFinancial", acViewPreview, strPeriod

'LastLine:
'    DoCmd.Close acForm, Me.name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdGo_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdGo_Click of VBA Document Form_frmFinReportWiz")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdGo_Click of VBA Document Form_frmFinReportWiz"
    
End Sub

Private Sub Form_Load()

Me.txtStartDate = Format(Now, "mm/dd/yyyy")
Me.txtEndDate = Format(Now, "mm/dd/yyyy")

End Sub
