Option Compare Database

Private Sub transaction_DblClick(Cancel As Integer)
   On Error GoTo transaction_DblClick_Error

    If Me.txtAccount = "" Or Me.txtAccount = 0 Then
        MsgBox "No account number has been assigned yet", vbOKOnly + vbInformation, "No account number"
        Exit Sub
    End If
    DoCmd.OpenForm "frmNotes", acNormal, , , , acDialog, Me.txtAccount
    
    If CurrentProject.AllForms("frmNotes").IsLoaded Then
        Me.transaction = Forms("frmNotes")!txtNote
        DoCmd.Close acForm, "frmNotes", acSaveYes
    End If

   On Error GoTo 0
   Exit Sub

transaction_DblClick_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure transaction_DblClick of VBA Document Form_Temp_Money subform1")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure transaction_DblClick of VBA Document Form_Temp_Money subform1"
End Sub

Private Sub txtAccount_LostFocus()
Dim rst As New ADODB.Recordset
Dim qry As String

   On Error GoTo txtAccount_LostFocus_Error

qry = "Select * from customer where account = " & Me.txtAccount

rst.Open qry, CurrentProject.Connection

Do While Not rst.EOF
    Me.Parent.txtAccount = Me.txtAccount
    Me.Parent.txtName = rst.Fields("name").value
    rst.MoveNext
Loop

rst.Close

If Me.txtAccount = "" Then
    MsgBox "Account not found", vbOKOnly + vbInformation, "No account"
End If

   On Error GoTo 0
   Exit Sub

txtAccount_LostFocus_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure txtAccount_LostFocus of VBA Document Form_Temp_Money subform1")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure txtAccount_LostFocus of VBA Document Form_Temp_Money subform1"

End Sub
