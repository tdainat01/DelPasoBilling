Option Compare Database
Option Explicit

'Global variables
Dim m As Integer
Dim d As Integer
Dim y As Integer
Dim acct As Long
Dim amnt As Single
Dim trans As String
Dim Code As String
Dim acct_name As String

Private Sub cmdOpenNotes_Click()

   On Error GoTo cmdOpenNotes_Click_Error

    If Me.txtAccount = "" Or Me.txtAccount = 0 Then
        MsgBox "No account number has been assigned yet", vbOKOnly + vbInformation, "No account number"
        Exit Sub
    End If
    DoCmd.OpenForm "frmNotes", acNormal, , , , acDialog, Me.txtAccount
    
    If CurrentProject.AllForms("frmNotes").IsLoaded Then
        trans = Forms("frmNotes")!txtNote
        DoCmd.Close acForm, "frmNotes", acSaveYes
    End If

   On Error GoTo 0
   Exit Sub

cmdOpenNotes_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdOpenNotes_Click of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdOpenNotes_Click of VBA Document Form_LoadPayments2"
        
End Sub


Private Sub cmdApply_Click()
'Add the fields to the list box.
Dim query As String
Dim lRecs As Long
Dim idx As Integer
'Dim dtTime As Date
Dim dtRecvd As Date
Dim iAcct As Long
Dim sAmnt As Currency
Dim sNotes As String
Dim sName As String
Dim sBillName As String
Dim sCareOf As String
Dim sTotal As Currency
'Dim rst As New ADODB.Recordset

   On Error GoTo cmdApply_Click_Error

idx = lstTransactions.ListCount

 'query = "SELECT top 1 System.system_date AS FirstOfDate" & _
 '           " FROM System order by System.auto_id DESC;"

    'rst.Open query, CurrentProject.Connection
    'If rst.BOF And rst.EOF Then
        'Me.txtDateReceived = Forms![DPM Main Menu]!txtSystemDate
        'Me.[DPM Main Menu].txtSystemDate
    'Else
    '    Me.txtDateReceived = rst.fields(0).value
    'End If

'dtTime = Me.txtPostDate

If Not IsNull(Me.txtAccount) Or Me.txtAccount <> "" Then
    iAcct = Me.txtAccount
Else
    MsgBox "An account number must be specified. No account number was found", vbOKOnly + vbExclamation, "Error"
    Exit Sub
End If

'If Not IsNull(Me.txtDateReceived) Then
'    dtRecvd = Me.txtDateReceived
'Else
'    dtRecvd = Now
'End If

If Not IsNull(Me.txtName) Then
    sName = Me.txtName
Else
    sName = ""
End If

If Not IsNull(Me.txtBillToName) Then
    sBillName = Me.txtBillToName
Else
    sBillName = ""
End If

If Not IsNull(Me.txtCareOfName) Then
    sCareOf = Me.txtCareOfName
Else
    sCareOf = ""
End If

If Not IsNull(Me.txtNotes) Then
    sNotes = Me.txtNotes
Else
    sNotes = ""
End If

If Not IsNull(Me.txtAmount) Then
    sAmnt = Me.txtAmount
Else
    sAmnt = 0
End If

Dim ix As Integer
Dim res As VbMsgBoxResult

If lstTransactions.ListCount > 1 Then
    For ix = 1 To lstTransactions.ListCount - 1
        If lstTransactions.Column(1, ix) = iAcct And CSng(lstTransactions.Column(5, ix)) = sAmnt Then
            res = MsgBox("You are about to add an account with the same amount already in the list. Continue?", _
                vbYesNo + vbInformation, "Duplicate Account")
            If res = vbYes Then
                'we know its a duplicate - add it anyway
                Exit For
            Else
                Exit Sub
            End If
        End If
    Next
End If

'add it
'With Me.lstTransactions
'    .AddItem iAcct & ";" & sName & ";" & sNotes & ";" & dtRecvd & ";" & Format(sAmnt, "$#,###.00")
'End With
'insert into temp_loadpayments

query = "insert into temp_loadpayments ([acct],[acct_name],[ref],[received],[amount]) " & _
        "Values('" & iAcct & "','" & Replace(sName, "'", "''") & "','" & Replace(sNotes, "'", "''") & "',#" & Forms![DPM Main Menu]!txtSystemDate & "#,'" & Round(sAmnt, 2) & "')"

CurrentProject.Connection.Execute query, lRecs

'clear the listbox, then requery
Me.lstTransactions.Requery

'Increment the count and values
Me.txtCount = lstTransactions.ListCount - 1 'remove the header from the count

'### these two lines force the last record of the listbox to be selected
Me.lstTransactions.SetFocus
Me.lstTransactions.Selected(CInt(Me.txtCount)) = True

If IsNull(Me.txtTotal) Then
    Me.txtTotal = 0
End If

sTotal = 0
For ix = 1 To lstTransactions.ListCount - 1
    If IsNull(lstTransactions.Column(5, ix)) Or lstTransactions.Column(5, ix) = "" Then
        sTotal = sTotal + CSng(0)
    Else
        sTotal = sTotal + CSng(lstTransactions.Column(5, ix))
    End If
Next

Me.txtTotal = sTotal

'clear out the fields
Me.txtAccount = ""
Me.txtName = ""
Me.txtBillToName = ""
Me.txtCareOfName = ""
'Me.txtDateReceived = Date
Me.txtNotes = ""
Me.txtAmount = ""

   On Error GoTo 0
   ' send cursor to account number field
   Me.txtAccount.SetFocus
   Exit Sub

cmdApply_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description
Debug.Print "help"
Me.txtAccount = ""
Me.txtName = ""
Me.txtCareOfName = ""
Me.txtBillToName = ""
'Me.txtDateReceived = ""
Me.txtNotes = ""
Me.txtAmount = ""

    Call LogError(errNum, errSource, errMsg & " in procedure cmdApply_Click of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdApply_Click of VBA Document Form_LoadPayments2"

End Sub

Private Sub cmdClear_Click()
   On Error GoTo cmdClear_Click_Error

    Me.txtAccount.SetFocus
    Me.txtAccount.text = ""
    Me.txtName.SetFocus
    Me.txtName.text = ""
    Me.txtBillToName = ""
    Me.txtCareOfName = ""
    'Me.txtDateReceived.SetFocus
    Me.txtDateReceived.text = ""
    'Me.txtTotal.SetFocus
    'Me.txtTotal.text = ""

   On Error GoTo 0
   Exit Sub

cmdClear_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdClear_Click of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdClear_Click of VBA Document Form_LoadPayments2"
End Sub

Private Sub cmdClearAll_Click()
    'Clear the listbox
    Dim query As String
    Dim lRecs As Long
   On Error GoTo cmdClearAll_Click_Error

    query = "DELETE FROM [temp_LoadPayments]"
    
    CurrentProject.Connection.Execute query, lRecs
    
    'clear the listbox, then requery
    Me.lstTransactions.Requery

    Me.txtCount = 0
    Me.txtTotal = 0

   On Error GoTo 0
   Exit Sub

cmdClearAll_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdClearAll_Click of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdClearAll_Click of VBA Document Form_LoadPayments2"

End Sub

Private Sub cmdDelete_Click()
    'lstTransactions.RemoveItem (lstTransactions.ListIndex + 1)
    'Get the row that needs to be deleted - we will have a problem if multiple values exist in the table that are identical
    Dim iTransId As Long
   On Error GoTo cmdDelete_Click_Error

    iTransId = lstTransactions.Column(0)
    
    Dim sQuery As String
    Dim lRecs As Long
    sQuery = "DELETE FROM [temp_LoadPayments] WHERE [trans_id] = " & iTransId
    
    
    CurrentProject.Connection.Execute sQuery, lRecs
    
    'clear the listbox, then requery
    Me.lstTransactions.Requery
    
    Me.txtCount = 0
    Me.txtTotal = 0
    
    'retotal
    Dim ix As Integer
    
    For ix = 1 To lstTransactions.ListCount - 1
    If IsNull(lstTransactions.Column(5, ix)) Or lstTransactions.Column(5, ix) = "" Then
        Me.txtTotal = CSng(Me.txtTotal) + CSng(0)
    Else
        Me.txtTotal = CSng(Me.txtTotal) + CSng(lstTransactions.Column(5, ix))
    End If
Next
    
    Me.txtCount = lstTransactions.ListCount - 1

   On Error GoTo 0
   Exit Sub

cmdDelete_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdDelete_Click of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdDelete_Click of VBA Document Form_LoadPayments2"
    
End Sub

Private Sub cmdDoLoad_Click()

   On Error GoTo cmdDoLoad_Click_Error

    Call ExecutingLoad
    Call CountBatches
    Me.lstTransactions.SetFocus
    
   On Error GoTo 0
   Exit Sub

cmdDoLoad_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdDoLoad_Click of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdDoLoad_Click of VBA Document Form_LoadPayments2"
    
End Sub

Private Sub cmdDoLoad_Exit(Cancel As Integer)
   On Error GoTo cmdDoLoad_Exit_Error

Me.TimerInterval = 1

   On Error GoTo 0
   Exit Sub

cmdDoLoad_Exit_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdDoLoad_Exit of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdDoLoad_Exit of VBA Document Form_LoadPayments2"

End Sub

Private Sub cmdDoLoad_LostFocus()
   On Error GoTo cmdDoLoad_LostFocus_Error

Me.TimerInterval = 1

   On Error GoTo 0
   Exit Sub

cmdDoLoad_LostFocus_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdDoLoad_LostFocus of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdDoLoad_LostFocus of VBA Document Form_LoadPayments2"
End Sub

Private Sub cmdLoadBatch_Click()

   On Error GoTo cmdLoadBatch_Click_Error

Me!frmBatch.Visible = True
Me!frmBatch.Requery
cmdDoLoad.Visible = True

   On Error GoTo 0
   Exit Sub

cmdLoadBatch_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdLoadBatch_Click of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdLoadBatch_Click of VBA Document Form_LoadPayments2"

End Sub

Private Sub cmdPostBatch_Click()

Dim ix As Integer
Dim query As String
Dim lRecs As Long
   On Error GoTo cmdPostBatch_Click_Error

For ix = 1 To lstTransactions.ListCount - 1
    
    m = Format(Me.txtPostDate, "mm")
    d = Format(Me.txtPostDate, "dd")
    y = Format(Me.txtPostDate, "yyyy")
    acct = lstTransactions.Column(1, ix)
    amnt = lstTransactions.Column(5, ix)
    Code = "MON"
    
    trans = lstTransactions.Column(3, ix)
    Call WriteMoney
Next

'clear out the ListBox
query = "delete from temp_loadpayments"
CurrentProject.Connection.Execute query, lRecs
Me.lstTransactions.Requery

Me.txtCount = 0
Me.txtTotal = 0

    Call Forms("DPM Main Menu").ExecutePost
   On Error GoTo 0
   Exit Sub

cmdPostBatch_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPostBatch_Click of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPostBatch_Click of VBA Document Form_LoadPayments2"

End Sub

Private Sub ExecutingLoad()
Dim rst As New ADODB.Recordset
Dim qry As String
Dim query As String
Dim lRecs As Long
Dim dt As String
Dim acct As String
Dim acctName As String
Dim ref As String
Dim amnt As Single
Dim lBatch As Long
Dim sTotal As Single
Dim ix As Integer

'Get which batch number to load
   On Error GoTo ExecutingLoad_Error

lBatch = Me!frmBatch.Controls("Batch ID").value


qry = "SELECT Temp_Money.ID, customer.name, Temp_Money.account_number, Temp_Money.amount, Temp_Money.transaction, " & _
    "Temp_Money.code, Temp_Money.posted, Temp_Money.behind_me, Temp_Money.trans_date " & _
    "FROM customer INNER JOIN Temp_Money ON customer.account = Temp_Money.account_number where behind_me = " & lBatch & ";"

'qry = "SELECT * from [temp_LoadPayments] where behind_me = " & lBatch

rst.Open qry, CurrentProject.Connection

If rst.BOF And rst.EOF Then
    'nothing to do
    Exit Sub
End If

Do While Not rst.EOF
    acct = rst.Fields("account_number").value
    acctName = rst.Fields("name").value
    dt = rst.Fields("trans_date").value
    ref = rst.Fields("transaction").value
    amnt = rst.Fields("amount").value
    query = "insert into temp_loadpayments ([acct],[acct_name],[ref],[received],[amount]) " & _
        "Values('" & acct & "','" & Replace(acctName, "'", "''") & "','" & ref & "',#" & dt & "#,'" & amnt & "')"
    CurrentProject.Connection.Execute query, lRecs
    'if lrecs = 0 then we have a problem
    If lRecs = 0 Then
        'log the error
    End If
    rst.MoveNext
Loop

'now delete all values the temp table
qry = "delete from temp_money where behind_me = " & lBatch
Dim iRec As Long

CurrentProject.Connection.Execute qry, iRec
'hide the load batch form
Me!frmBatch.Visible = False
lstTransactions.Requery

    If Me.lstTransactions.ListCount > 0 Then
        'calculate the value
        sTotal = 0
        For ix = 1 To lstTransactions.ListCount - 1
            If IsNull(lstTransactions.Column(5, ix)) Or lstTransactions.Column(5, ix) = "" Then
                sTotal = sTotal + CSng(0)
            Else
                sTotal = sTotal + CSng(lstTransactions.Column(5, ix))
            End If
        Next
        Me.txtTotal = sTotal
        Me.txtCount = lstTransactions.ListCount - 1
    Else
        Me.txtCount = 0
        Me.txtTotal = 0
    End If

   On Error GoTo 0
   Exit Sub

ExecutingLoad_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure ExecutingLoad of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure ExecutingLoad of VBA Document Form_LoadPayments2"

End Sub

Private Function WriteMoney() As Long
    Dim ResultString As String
    Dim myMatches As MatchCollection
    Dim myRegExp As RegExp
   On Error GoTo WriteMoney_Error

    Set myRegExp = New RegExp
    myRegExp.Pattern = "[+-]?[0-9]{1,3}(?:,?[0-9]{3})*(?:\.[0-9]{2})?"
    Set myMatches = myRegExp.Execute(amnt)
    If myMatches.Count >= 1 Then
        ResultString = myMatches(0).value
    Else
        ResultString = ""
    End If
        
    Dim qry As String
    Dim lRecs As Long
    
    
    trans = Replace(trans, "'", "''")
    
    'TODO dehind_me never gets set to anything meaningful
    qry = "INSERT Into [Money] (m_month,m_day,m_year,account_number,amount," & _
          "[transaction],code,posted,behind_me,trans_date)" & _
        " VALUES ('" & Format(m, "0#") & "','" & Format(d, "0#") & "','" & y & "','" & _
        acct & "'," & amnt & ",'" & trans & "','" & Code & _
        "','N',0,#" & Now & "#);"
    CurrentProject.Connection.Execute qry, lRecs
    WriteMoney = lRecs

   On Error GoTo 0
   Exit Function

WriteMoney_Error:
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
        msg = ". This error has been logged."
    Else
        msg = ". This error has NOT been logged."
    End If
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure WriteMoney of VBA Document Form_LoadPayments2" & msg
    WriteMoney = -1

End Function


Private Sub cmdPrint_Click()
    Dim answer As VbMsgBoxResult
    Dim ix As Integer
    Dim qry As String
    
   On Error GoTo cmdPrint_Click_Error

    answer = MsgBox("This will send the results of the batch list to the default printer. " & vbCrLf & _
        "Selecting NO will send the information to the screen. Do you want to continue?", _
        vbYesNoCancel + vbInformation, "Send to Printer")
    
    'save contents of listbox to table
    For ix = 1 To lstTransactions.ListCount - 1
        m = Format(Me.txtPostDate, "mm")
        d = Format(Me.txtPostDate, "dd")
        y = Format(Me.txtPostDate, "yyyy")
        acct = lstTransactions.Column(1, ix)
        acct_name = lstTransactions.Column(2, ix)
        amnt = lstTransactions.Column(5, ix)
        Code = "MON"
        trans = Replace(lstTransactions.Column(3, ix), "'", "''")
        Call Write_Print
    Next
    
    If answer = vbYes Then
        'Use table as recordsource and open print form
        DoCmd.OpenReport "rptPrintBatch", acViewPreview
        'prtdefault
    ElseIf answer = vbNo Then
        DoCmd.OpenReport "rptPrintBatch", acViewReport
    Else
        'cancel was selected - do nothing
    End If
    
    'delete the contents of the temp table
    qry = "delete from temp_print"
    CurrentProject.Connection.Execute qry

   On Error GoTo 0
   Exit Sub

cmdPrint_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrint_Click of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrint_Click of VBA Document Form_LoadPayments2"
    
End Sub

Private Sub cmdQuit_Click()
   On Error GoTo cmdQuit_Click_Error

    DoCmd.Close acForm, Me.Name, acSaveYes
    'docmd.OpenForm

   On Error GoTo 0
   Exit Sub

cmdQuit_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdQuit_Click of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdQuit_Click of VBA Document Form_LoadPayments2"
End Sub

Private Sub cmdSaveBatch_Click()
'write all the data to temp_money
Dim ix As Integer
Dim lBatch As Long
Dim query As String
Dim lRecs As Long
Dim rst As New ADODB.Recordset

    'Get a batch number
   On Error GoTo cmdSaveBatch_Click_Error

    query = "SELECT Max(Temp_Money.behind_me) AS MaxOfbehind_me FROM [Temp_Money];"
    rst.Open query, CurrentProject.Connection, adOpenForwardOnly, adLockReadOnly
    
    If rst.BOF And rst.EOF Then
        lBatch = 1
    Else
        'should only be one row
        Do While Not rst.EOF
            If IsNull(rst.Fields(0).value) Then
                lBatch = 1
            Else
                lBatch = rst.Fields(0).value + 1    'increment the batch number to the next highest
            End If
            rst.MoveNext
        Loop
    End If
    
    If lBatch = 0 Then
        lBatch = 1
    End If

For ix = 1 To lstTransactions.ListCount - 1
    
    m = Format(Me.txtPostDate, "mm")
    d = Format(Me.txtPostDate, "dd")
    y = Format(Me.txtPostDate, "yyyy")
    acct = lstTransactions.Column(1, ix)
    amnt = lstTransactions.Column(5, ix)
    Code = "MON"
    trans = lstTransactions.Column(3, ix) 'notes or refs
    Call WriteMoney_Temp(lBatch)
Next

'clear out the ListBox
query = "delete from temp_loadpayments"
CurrentProject.Connection.Execute query, lRecs
Me.lstTransactions.Requery

Me.txtCount = 0
Me.txtTotal = 0

    Call CountBatches
   On Error GoTo 0
   Exit Sub

cmdSaveBatch_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdSaveBatch_Click of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdSaveBatch_Click of VBA Document Form_LoadPayments2"

End Sub

Private Sub Form_Open(Cancel As Integer)
   On Error GoTo Form_Open_Error

    Me.txtPostDate = Now
    Dim sTotal As Single
    Dim ix As Integer
    
    If Me.lstTransactions.ListCount > 0 Then
        'calculate the value
        sTotal = 0
        For ix = 1 To lstTransactions.ListCount - 1
            If IsNull(lstTransactions.Column(5, ix)) Or lstTransactions.Column(5, ix) = "" Then
                sTotal = sTotal + CSng(0)
            Else
                sTotal = sTotal + CSng(lstTransactions.Column(5, ix))
            End If
        Next
        Me.txtTotal = sTotal
        Me.txtCount = lstTransactions.ListCount - 1
    Else
        Me.txtCount = 0
        Me.txtTotal = 0
    End If
    
    Call CountBatches

   On Error GoTo 0
   Exit Sub

Form_Open_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Open of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Open of VBA Document Form_LoadPayments2"
    
End Sub


Private Sub Form_Timer()

If Me.cmdDoLoad.Visible = True Then
    Me.TimerInterval = 0
    Me.cmdDoLoad.Visible = False
End If

End Sub

Private Sub txtAccount_LostFocus()
   On Error GoTo txtAccount_LostFocus_Error
'=FormatNumber([txtCurrentRead]-[txtPreviousRead],0)
'Dim lCurRead As Long
'Dim lPrevRead As Long
'Dim lUsage As Long
'Dim sym As String
'Dim lSize As Single

If IsNull(Me.txtAccount.text) Or Me.txtAccount.text = "" Then
    Exit Sub
End If

'Look up the account name
Dim qry As String
Dim rst As New ADODB.Recordset
'qry = "SELECT IIf([bill_name] Is Null Or [bill_name]='' Or [bill_name]='UNKNOWN',IIf([name] Is Null Or [name]='' " & _
'      " Or [name]='UNKNOWN','OCCUPANT',[name]),[bill_name]) AS username, total_due FROM customer " & _
'      " where account = " & Me.txtAccount
qry = "SELECT [name], [bill_name], [care_of], [total_due], [current_read], [meter_size], [previous_read], [unit_measure] FROM customer where account = " & Me.txtAccount

rst.Open qry, CurrentProject.Connection

If rst.BOF And rst.EOF Then
    MsgBox "Account not on file", vbOKOnly + vbInformation, "Account not found"
    Exit Sub
End If

Do While Not rst.EOF
    'Me.txtName = rst.Fields("username").value
'    lCurRead = IIf(IsNull(rst.Fields("current_read").value), 0, rst.Fields("current_read").value)
'    lPrevRead = IIf(IsNull(rst.Fields("previous_read").value), 0, rst.Fields("previous_read").value)
'    lUsage = lCurRead - lPrevRead
'    sym = IIf(IsNull(rst.Fields("unit_measure").value), 0, rst.Fields("unit_measure").value)
'    lSize = IIf(IsNull(rst.Fields("meter_size").value), 0, rst.Fields("meter_size").value)
'
'    If Len(sym) >= 1 Then
'        Me.txtNotes = FormatNumber(lSize, 2) & " " & CStr(lUsage) & sym
'    End If
    
    Me.txtName = rst.Fields("name").value
    Me.txtBillToName = rst.Fields("bill_name").value
    Me.txtCareOfName = rst.Fields("care_of").value
    Me.txtAmount = rst.Fields("total_due").value
    rst.MoveNext
Loop

   On Error GoTo 0
   Exit Sub

txtAccount_LostFocus_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure txtAccount_LostFocus of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure txtAccount_LostFocus of VBA Document Form_LoadPayments2"
End Sub

Private Function ListBox2Array(ListBox1 As ListBox) As String()
Dim ar() As String
Dim Count As Integer, I As Integer, j As Integer
   On Error GoTo ListBox2Array_Error

Count = 0

For I = 0 To ListBox1.ListCount - 1
    'check if the row is selected and add to count
    If ListBox1.Selected(I) Then Count = Count + 1
Next I

'based on the above count declare the array
ReDim ar(Count)

j = 0
For I = 0 To ListBox1.ListCount - 1
    If ListBox1.Selected(I) Then
        'if selected then store the item from the
        'first column in the array. change 1 to the
        'respective column number
        ar(j) = ListBox1.Column(0, I) & ";" & ListBox1.Column(1, I) & ";" & ListBox1.Column(2, I) & ";" & ListBox1.Column(3, I) & ";" & ListBox1.Column(4, I)
        j = j + 1
    End If
Next I

'Check values stored in array
'For i = 0 To Count - 1
'
'    MsgBox ar(i)
'
'Next i

ListBox2Array = ar

   On Error GoTo 0
   Exit Function

ListBox2Array_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure ListBox2Array of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure ListBox2Array of VBA Document Form_LoadPayments2"

End Function

Private Function WriteMoney_Temp(lBatch As Long) As Long

    Dim ResultString As String
    Dim myMatches As MatchCollection
    Dim myRegExp As RegExp
   On Error GoTo WriteMoney_Temp_Error

    Set myRegExp = New RegExp
    myRegExp.Pattern = "[+-]?[0-9]{1,3}(?:,?[0-9]{3})*(?:\.[0-9]{2})?"
    Set myMatches = myRegExp.Execute(amnt)
    If myMatches.Count >= 1 Then
        ResultString = myMatches(0).value
    Else
        ResultString = ""
    End If
'
'    If ResultString <> "" Then
'        amnt = ResultString
'    End If
        
    Dim qry As String
    Dim lRecs As Long
    
'    If Code = "PMT" Then
'        Amnt = Amnt * -1
'    End If
    trans = Replace(trans, "'", "''")
    qry = "INSERT Into [Temp_Money] (m_month,m_day,m_year,account_number,amount," & _
          "[transaction],code,posted,behind_me,trans_date)" & _
        " VALUES ('" & m & "','" & d & "','" & y & "','" & _
        acct & "'," & amnt & ",'" & trans & "','" & Code & _
        "','N'," & lBatch & ",#" & Now & "#);"
    CurrentProject.Connection.Execute qry, lRecs
    WriteMoney_Temp = lRecs
   
   On Error GoTo 0
   Exit Function

WriteMoney_Temp_Error:
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
        msg = ". This error has been logged."
    Else
        msg = ". This error has NOT been logged."
    End If
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure WriteMoney_Temp of VBA Document Form_LoadPayments2" & msg
    WriteMoney_Temp = -1

End Function

Private Function Write_Print() As Long
    Dim ResultString As String
    Dim myMatches As MatchCollection
    Dim myRegExp As RegExp
   On Error GoTo Write_Print_Error

    Set myRegExp = New RegExp
    myRegExp.Pattern = "[+-]?[0-9]{1,3}(?:,?[0-9]{3})*(?:\.[0-9]{2})?"
    Set myMatches = myRegExp.Execute(amnt)
    If myMatches.Count >= 1 Then
        ResultString = myMatches(0).value
    Else
        ResultString = ""
    End If
'
'    If ResultString <> "" Then
'        amnt = ResultString
'    End If
    
    Dim qry As String
    Dim lRecs As Long
    qry = "INSERT Into [Temp_Print] (m_month,m_day,m_year,account_number,account_name,amount," & _
          "[transaction],code,posted,behind_me,trans_date)" & _
        " VALUES ('" & Format(m, "mm") & "','" & Format(d, "dd") & "','" & Format(y, "yyyy") & "','" & _
        acct & "','" & Replace(acct_name, "'", "''") & "'," & amnt & ",'" & Replace(trans, "'", "''") & "','" & Code & _
        "','N',0,#" & Now & "#);"
    CurrentProject.Connection.Execute qry, lRecs
    Write_Print = lRecs
      
   On Error GoTo 0
   Exit Function

Write_Print_Error:
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
        msg = ". This error has been logged."
    Else
        msg = ". This error has NOT been logged."
    End If
    Write_Print = -1
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Write_Print of VBA Document Form_LoadPayments2" & msg
End Function

Private Sub txtAmount_Click()
   
   On Error GoTo txtAmount_Click_Error
    If IsNumeric(Me.txtAmount) Then
        Me.txtAmount = Format(Me.txtAmount, "standard")
    End If

   On Error GoTo 0
   Exit Sub

txtAmount_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure txtAmount_Click of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure txtAmount_Click of VBA Document Form_LoadPayments2"
End Sub

Private Sub CountBatches()
    'Now count up how many batches have been saved
    Dim qry As String
    Dim rst As New ADODB.Recordset
    Dim bCount As Integer
   On Error GoTo CountBatches_Error

    qry = "SELECT Temp_Money.behind_me, Count(Temp_Money.behind_me) AS CountOfbehind_me" & _
            " FROM Temp_Money GROUP BY Temp_Money.behind_me;"
    
    rst.Open qry, CurrentProject.Connection
    
    If rst.BOF And rst.EOF Then
        bCount = 0
    End If
    
    Do While Not rst.EOF
        bCount = bCount + 1
        rst.MoveNext
    Loop
    
    If bCount = 1 Then
        Me.lblBatchCount.Caption = bCount & " batch"
    Else
        Me.lblBatchCount.Caption = bCount & " batches"
    End If

   On Error GoTo 0
   Exit Sub

CountBatches_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure CountBatches of VBA Document Form_LoadPayments2")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure CountBatches of VBA Document Form_LoadPayments2"

End Sub
