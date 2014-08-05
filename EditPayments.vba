Option Compare Database

Private Sub cmdQuit_Click()
    DoCmd.OpenForm "DPM Main Menu", acNormal
    DoCmd.Close acForm, Me.Name, acSaveYes
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo 0
Dim I As Integer
Dim strCharacter As String
    ' Convert ANSI value to character string.
    strCharacter = Chr(KeyAscii)
    ' Convert character to upper case, then to ANSI value.
    KeyAscii = Asc(UCase(strCharacter))
    If KeyAscii = 81 Or KeyAscii = 27 Then
        DoCmd.Close acForm, Me.Name, acSaveYes
    End If
    
    If KeyAscii = 89 Then
        'Perform some action
        If Me.txtAccountNumber <> "" And IsNumeric(Me.txtAccountNumber) Then
            'validate all other fields
            If Not IsDate(Me.txtDate) Then
                MsgBox Me.txtDate & " is not a valid date. Please enter a valid date and then try again. ", _
                    vbOKOnly + vbExclamation, "Error"
                Exit Sub
            End If
            'update changes
            Dim qry As String
            qry = "UPDATE [Money] SET [m_month]='" & Format(Me.txtDate, "mm") & "', [m_day]='" & Format(Me.txtDate, "dd") & _
                "', [m_year]='" & Format(Me.txtDate, "yyyy") & "', [account_number]=" & Me.txtAccountNumber & "," & _
                " [amount]=" & Me.txtAmount & ", [transaction]='" & Me.txtTransNote & "', [code]='" & Me.txtTransCode & "'" & _
                " WHERE ID=" & Me.txtLineNum
            CurrentProject.Connection.Execute qry
            If CurrentProject.Connection.Errors.Count > 0 Then
            Dim sErrors As String
                For I = 0 To CurrentProject.Connection.Errors.Count
                    sErrors = sErrors & CurrentProject.Connection.Errors(I).Description & vbCrLf
                Next I
                MsgBox "There was an error trying to update the transaction. the error returned by the system was: " & _
                    sErrors, vbOKOnly + vbInformation, "Error"
            End If
            MsgBox "Payment was successfully changed.", vbOKOnly + vbInformation, "Success"
            'RESET all the fields
            Me.txtLineNum = ""
            Me.txtAccountNumber = ""
            Me.txtAmount = ""
            Me.txtDate = ""
            Me.txtTransNote = ""
            Me.txtTransCode = ""
        End If
    End If
    
End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.txtDate = Format(Now, "mm/dd/yyyy")
    Me.txtLineNum.SetFocus
End Sub

Private Sub txtLineNum_LostFocus()

If Me.txtLineNum = "" Or Not IsNumeric(Me.txtLineNum) Then
    Exit Sub
End If

'Look up a transaction by the line number and then if found load it.
Dim qry As String
Dim rst As New ADODB.Recordset
qry = "SELECT * FROM [MONEY] WHERE ID = " & Me.txtLineNum
rst.Open qry, CurrentProject.Connection

If rst.BOF And rst.EOF Then
    'exit
End If

Do While Not rst.EOF
    Me.txtDate = rst.Fields("trans_date").value
    Me.txtAccountNumber = rst.Fields("account_number").value
    Me.txtAmount = rst.Fields("amount").value
    Me.txtTransNote = rst.Fields("transaction").value
    Me.txtTransCode = rst.Fields("code").value
    rst.MoveNext
Loop

End Sub
