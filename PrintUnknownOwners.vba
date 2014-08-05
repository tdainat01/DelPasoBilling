Option Compare Database

Private Sub cmdPrint_Click()
Dim crit As String

'If opt = "fOpt1.Selected" Then
   On Error GoTo cmdPrint_Click_Error

    crit = "account>=" & Me.cmbMin.value & " AND account <= " & Me.cmbMax.value & " AND name like 'UNKNOWN'"
    DoCmd.OpenReport "FullRollReport", acViewLayout, , crit
'End If

   On Error GoTo 0
   Exit Sub

cmdPrint_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrint_Click of VBA Document Form_PrintUnknownOwners")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrint_Click of VBA Document Form_PrintUnknownOwners"
End Sub

Private Sub Form_Load()
Dim rst As New ADODB.Recordset
Dim min As String
Dim max As String

   On Error GoTo Form_Load_Error

rst.Open "SELECT Min(customer.account) AS MinOfaccount FROM customer;", CurrentProject.Connection
min = rst.Fields(0).value
rst.Close

rst.Open "SELECT Max(customer.account) AS MinOfaccount FROM customer;", CurrentProject.Connection
max = rst.Fields(0).value
rst.Close

Me.cmbMin.value = min
Me.cmbMax.value = max
'Me.txtYear = Format(Now(), "yyyy")

   On Error GoTo 0
   Exit Sub

Form_Load_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Load of VBA Document Form_PrintUnknownOwners")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Load of VBA Document Form_PrintUnknownOwners"
End Sub
