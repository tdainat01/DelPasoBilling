Option Explicit
Dim acct As Long
Dim fTimeStamp As Boolean

Private Sub cmdQuit_Click()
    'DoCmd.Close acForm, Me.name, acSaveYes

   On Error GoTo cmdQuit_Click_Error

    Me.Visible = False


   On Error GoTo 0
   Exit Sub

cmdQuit_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdQuit_Click of VBA Document Form_frmNotes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdQuit_Click of VBA Document Form_frmNotes"

End Sub

Private Sub cmdSave_Click()
'Save the data to a temp table
Dim qry As String
Dim rst As New ADODB.Recordset
  
  
   On Error GoTo cmdSave_Click_Error

If IsNull(Me.txtAccount) Or Me.txtAccount = "" Then
    'some error occurred here as we don't have a valid account number
    Call MsgBox("Unable to save the note(s) at this time as there is no valid account number. If you know the account number, please enter it into the account field and then try and save your note(s) again.", vbExclamation, "No Account")
    Exit Sub
End If

qry = "select note from notes where cust_acct = " & Me.txtAccount
rst.Open qry, CurrentProject.Connection

If rst.BOF And rst.EOF Then
    'no notes exist, carry on
Else
    qry = "delete * from notes where cust_acct = " & Me.txtAccount
    CurrentProject.Connection.Execute qry
End If
Dim sNote As String
sNote = Replace(Me.txtNote, "'", "''")
qry = "insert into notes(cust_acct,[note]) VALUES(" & CLng(Me.txtAccount) & ",'" & sNote & "')"
CurrentProject.Connection.Execute qry

   On Error GoTo 0
   Exit Sub

cmdSave_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdSave_Click of VBA Document Form_frmNotes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdSave_Click of VBA Document Form_frmNotes"

End Sub

Private Sub Form_Activate()
Dim qry As String
Dim rst As New ADODB.Recordset

'check to see if the account field matches the value being passed in
   On Error GoTo Form_Activate_Error

If IsNull(Me.OpenArgs) And IsEmpty(vOpenArgs) Then
    Exit Sub
End If

If CLng(Me.txtAccount) = CLng(vOpenArgs) Then
    'do nothing
Else
    Me.txtAccount = vOpenArgs
    Me.txtNote = ""
    qry = "select note from notes where cust_acct = " & vOpenArgs
    rst.Open qry, CurrentProject.Connection
    If rst.BOF And rst.EOF Then
        'no notes exist, carry on
        If Me.txtNote <> "" Then
            Me.txtNote = ""
        End If
    Else
        Do While Not rst.EOF
            Me.txtNote = Me.txtNote & rst.Fields(0).value
            rst.MoveNext
        Loop
        Me.txtNote = Me.txtNote & vbCrLf
    End If
End If

   On Error GoTo 0
   Exit Sub

Form_Activate_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Activate of VBA Document Form_frmNotes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Activate of VBA Document Form_frmNotes"

End Sub

Private Sub Form_Open(Cancel As Integer)
   
Dim qry As String
Dim rst As New ADODB.Recordset
Dim acct As Long
  
   On Error GoTo Form_Open_Error

If IsNull(Me.OpenArgs) And IsEmpty(vOpenArgs) Then
    acct = InputBox("The note you have opended is not connected to an account number. Please enter an account number here", "No Acct Number")
    If acct = 0 Then
        Call cmdQuit_Click
    End If
Else
    If IsNull(Me.OpenArgs) Then
        acct = vOpenArgs
    Else
        acct = Me.OpenArgs
        vOpenArgs = Me.OpenArgs
    End If
End If

'Get any values for this account in the notes table
qry = "select note from notes where cust_acct = " & acct
rst.Open qry, CurrentProject.Connection
If rst.BOF And rst.EOF Then
    'no notes exist, carry on
    If Me.txtNote <> "" Then
        Me.txtNote = ""
    End If
Else
    Do While Not rst.EOF
        Me.txtNote = Me.txtNote & rst.Fields(0).value
        rst.MoveNext
    Loop
    Me.txtNote = Me.txtNote & vbCrLf
End If

Me.txtAccount = acct


   On Error GoTo 0
   Exit Sub

Form_Open_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Open of VBA Document Form_frmNotes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Open of VBA Document Form_frmNotes"
 
End Sub

Private Sub txtNote_AfterUpdate()
    Me.txtNote = Me.txtNote & vbCrLf & "*********" & Now
End Sub

