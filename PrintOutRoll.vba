Option Compare Database
Option Explicit

Private Sub cmbMax_Enter()
   On Error GoTo cmbMax_Enter_Error

    Call cmdPrint_Click

   On Error GoTo 0
   Exit Sub

cmbMax_Enter_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmbMax_Enter of VBA Document Form_PrintOutRoll")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmbMax_Enter of VBA Document Form_PrintOutRoll"
End Sub

Private Sub cmbMax_LostFocus()
   On Error GoTo cmbMax_LostFocus_Error

    Call cmdPrint_Click

   On Error GoTo 0
   Exit Sub

cmbMax_LostFocus_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmbMax_LostFocus of VBA Document Form_PrintOutRoll")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmbMax_LostFocus of VBA Document Form_PrintOutRoll"
End Sub

Private Sub Form_Open(Cancel As Integer)
Dim rst As New ADODB.Recordset
Dim min As String
Dim max As String
   On Error GoTo Form_Open_Error

vOpenArgs = Me.OpenArgs

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

Form_Open_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Open of VBA Document Form_PrintOutRoll")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Open of VBA Document Form_PrintOutRoll"
End Sub
'------------------------------------------------------------
' cmdGoBack_Click
'
'------------------------------------------------------------
Private Sub cmdGoBack_Click()
   On Error GoTo cmdGoBack_Click_Error

    DoCmd.OpenForm "Opt 6 Form", acNormal, "", "", , acNormal
    DoCmd.Close acForm, "PrintOutRoll"

   On Error GoTo 0
   Exit Sub

cmdGoBack_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdGoBack_Click of VBA Document Form_PrintOutRoll")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdGoBack_Click of VBA Document Form_PrintOutRoll"

End Sub
'------------------------------------------------------------
' cmdPrint_Click
'
'------------------------------------------------------------
Private Sub cmdPrint_Click()
Dim crit As String

   On Error GoTo cmdPrint_Click_Error

If vOpenArgs = "fOpt1.Selected" Then
    crit = "account>=" & Me.cmbMin.value & " AND account <= " & Me.cmbMax.value
    DoCmd.OpenReport "FullRollReport", acViewLayout, , crit
End If

   On Error GoTo 0
   Exit Sub

cmdPrint_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrint_Click of VBA Document Form_PrintOutRoll")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrint_Click of VBA Document Form_PrintOutRoll"

End Sub


