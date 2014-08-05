Option Compare Database
Option Explicit

Private Sub cmdGo_Click()
Dim sReport As String
Dim ix As Long
Dim iy As Long
Dim lRecs As Long
Dim rpt As Report
Dim qry As QueryDef
Dim query As String


   On Error GoTo cmdGo_Click_Error

sReport = "rptChargeReport"

'Parameter Checking

If IsNull(txtDateFrom) Or IsEmpty(txtDateFrom) Or Trim(txtDateFrom) = "" Then
    Call MsgBox("The From Date is empty.", vbExclamation, "Nothing to do")
    Exit Sub
End If

If IsNull(txtDateTo) Or IsEmpty(txtDateTo) Or Trim(txtDateFrom) = "" Then
    Call MsgBox("The To Date is empty.", vbExclamation, "Nothing to do")
    Exit Sub
End If

If Not IsDate(txtDateFrom) Or Not IsDate(txtDateTo) Then
    Call MsgBox("The From Date or the To Date is not valid.", vbExclamation, "Nothing to do")
    Exit Sub
End If

If CDate(txtDateFrom) > CDate(txtDateTo) Then
    Call MsgBox("The From Date cannot be greater than the To Date.", vbExclamation, "Nothing to do")
    Exit Sub
End If

query = "SELECT customer.account, customer.name, Money.amount, customer.total_due" & _
        " FROM customer INNER JOIN [Money] ON customer.account = Money.account_number" & _
        " WHERE (((Money.code)='CHG' Or (Money.code)='SCH') AND ((Money.posted)='Y') AND" & _
        " ((Money.trans_date) Between " & "#" & txtDateFrom & " 00:00:00# and #" & txtDateTo & " 23:59:59#));"

DoCmd.OpenReport sReport, acViewDesign
Set rpt = Reports.Item(sReport)
rpt.RecordSource = query
DoCmd.Close acReport, rpt.Name, acSaveYes
DoCmd.OpenReport sReport, acViewPreview, , , , txtDateFrom & "-" & txtDateTo

   On Error GoTo 0
   Exit Sub

cmdGo_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdGo_Click of VBA Document Form_frmSelDate")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdGo_Click of VBA Document Form_frmSelDate"

End Sub

Private Sub cmdQuit_Click()

   On Error GoTo cmdQuit_Click_Error
    DoCmd.Close acForm, Me.Name, acSaveNo
    DoCmd.OpenForm "ReportMenu"

   On Error GoTo 0
   Exit Sub

cmdQuit_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdQuit_Click of VBA Document Form_frmSelDate")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdQuit_Click of VBA Document Form_frmSelDate"
End Sub
