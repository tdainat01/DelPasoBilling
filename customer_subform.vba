Option Compare Database

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    Me.Requery

   On Error GoTo 0
   Exit Sub

Form_Load_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Load of VBA Document Form_customer_subform")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Load of VBA Document Form_customer_subform"
End Sub

Private Sub txtAccount_LostFocus()
Dim qry As String
Dim rst As New ADODB.Recordset

   On Error GoTo txtAccount_LostFocus_Error

If IsNull(Me.txtAccount) Or IsEmpty(Me.txtAccount) Or Me.txtAccount = 0 Or Me.txtAccount = "" Then Exit Sub

qry = "select name from customer where account = " & Me.txtAccount

rst.Open qry, CurrentProject.Connection

If rst.BOF And rst.EOF Then
    MsgBox "Account not on file", vbOKOnly + vbInformation, "Account not found"
    Exit Sub
End If

Do While Not rst.EOF
    Me.txtName = rst.Fields("name").value
    rst.MoveNext
Loop

DoCmd.GoToRecord , , acNewRec

   On Error GoTo 0
   Exit Sub

txtAccount_LostFocus_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure txtAccount_LostFocus of VBA Document Form_customer_subform")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure txtAccount_LostFocus of VBA Document Form_customer_subform"

End Sub
