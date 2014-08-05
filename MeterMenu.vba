Option Compare Database

Private Sub cmdEnterBatch_Click()
   On Error GoTo cmdEnterBatch_Click_Error

    sCallingForm = Me.Name
    DoCmd.OpenForm "MeterReads", acNormal
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdEnterBatch_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdEnterBatch_Click of VBA Document Form_MeterMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdEnterBatch_Click of VBA Document Form_MeterMenu"
End Sub

Private Sub cmdPostBatch_Click()
'this routine has been deprecated.
Exit Sub
Dim rst As New ADODB.Recordset
Dim wtr As New ADODB.Recordset
Dim query As String
Dim sQry As String
   
   On Error GoTo cmdPostBatch_Click_Error

query = "SELECT * FROM [MeterReads] WHERE [posted] = 'N'"
rst.Open query, CurrentProject.Connection

If rst.BOF And rst.EOF Then
    'nothing to do
    Call MsgBox("Nothing to post.", vbOKOnly + vbInformation, "No Records")
    Exit Sub
End If

'This only picks up records where an operator entered meter reads.
'What about any account that was missed?
Do While Not rst.EOF
    sQry = "SELECT [current_read], [current_date], [previous_read], [previous_date] " & _
        "from [customer] where [account] = " & rst.Fields("account").value
    wtr.Open sQry, CurrentProject.Connection, adOpenDynamic, adLockPessimistic
    If wtr.BOF And wtr.EOF Then
        'somehow that customer doesn't exist
        Call MsgBox("There is a problem accessing account " & rst.Fields("account").value & _
            ". This account does not exist. Please check the Meter Reads table, make any necessary changes and then try again.", _
            vbOKOnly + vbInformation, "Missing Customer")
        Exit Sub
    Else
    wtr.Fields("previous_date").value = wtr.Fields("current_date").value
    wtr.Fields("previous_read").value = wtr.Fields("current_read").value
        If IsNull(rst.Fields("low_read").value) Then
            wtr.Fields("current_read").value = rst.Fields("normal_read").value
        Else
            wtr.Fields("current_read").value = rst.Fields("normal_read").value + rst.Fields("low_read").value
        End If
    wtr.Fields("current_date").value = rst.Fields("batch_date").value
    wtr.Update
    wtr.Close
    rst.MoveNext
    End If
Loop

   'empty the meter reads table
   query = "UPDATE [MeterReads] SET [posted] = 'Y' WHERE [posted] = 'N'"
   CurrentProject.Connection.Execute query
   On Error GoTo 0
   Call MsgBox("All meter reads have been posted", vbInformation, "Posted")
    
   On Error GoTo 0
   Exit Sub

cmdPostBatch_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description
Dim msg As String
Dim F As Boolean
    F = LogError(Err.Number, Err.source, Err.Description)
    If F Then
        msg = "This error was logged"
    Else
        msg = "This error was NOT logged"
    End If
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPostBatch_Click of VBA Document Form_MeterMenu. " & msg

End Sub


Private Sub cmdQuit_Click()
   On Error GoTo cmdQuit_Click_Error

    DoCmd.OpenForm "DPM Main Menu", acNormal
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdQuit_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdQuit_Click of VBA Document Form_MeterMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdQuit_Click of VBA Document Form_MeterMenu"
End Sub
