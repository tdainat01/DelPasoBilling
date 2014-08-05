Option Compare Database
Option Explicit
Dim frmSender As String
Private Sub cmdClear_Click()
Dim qry As String
    On Error GoTo cmdClear_Click_Error
    'delete all records in the temp_results table
    qry = "delete * from temp_results"
    CurrentProject.Connection.Execute qry
    Me.temp_results_subform.Requery
    On Error GoTo 0
    Exit Sub

cmdClear_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
    errNum = Err.Number
    errSource = Err.source
    errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdClear_Click of VBA Document Form_frmSearch")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdClear_Click of VBA Document Form_frmSearch"
End Sub

Private Sub cmdQuit_Click()

   On Error GoTo cmdQuit_Click_Error

DoCmd.Close acForm, Me.Form.Name, acSaveYes
DoCmd.OpenForm "DPM Main Menu"

   On Error GoTo 0
   Exit Sub

cmdQuit_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdQuit_Click of VBA Document Form_frmSearch")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdQuit_Click of VBA Document Form_frmSearch"

End Sub

'Form searches based on criteria and then inserts the results in a search table.
Private Sub cmdSearch_Click()
Dim qry As String
Dim field As String
Dim criteria As String
Dim sortby As String

   On Error GoTo cmdSearch_Click_Error
Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset

'select the field to search in
Select Case Me.fraSearchIn.value
    Case Is = 1 'customer name
        field = "[name]"
    Case Is = 2 'account number
        field = "[account]"
    Case Is = 3 'physical address
        field = "[phy_address]"
    Case Is = 4 'mail to address
        field = "[addr1]"
    Case Is = 5 'home phone
        'TODO - needs a seperate query as this is from a different table
        'Call MsgBox("Searching for an Home Phone is not yet implemented. Please try a different search criteria.", vbExclamation, "Not yet implemented")
        'Exit Sub
        field = "phones"
    Case Is = 6 'APN number (not implemented)
        Call MsgBox("Searching for an APN is not yet implemented. Please try a different search criteria.", vbExclamation, "Not yet implemented")
        Exit Sub
End Select

'select how the search should be conducted
Select Case Me.fraSearchFor.value
    Case Is = 1 'start with
    If field = "account" Then
        Select Case Len(Me.txtSearch)
            Case Is = 1
                criteria = field & " BETWEEN " & Me.txtSearch & "000 AND " & Me.txtSearch & "999"
            Case Is = 2
                criteria = field & " BETWEEN " & Me.txtSearch & "00 AND " & Me.txtSearch & "99"
            Case Is = 3
                criteria = field & " BETWEEN " & Me.txtSearch & "0 AND " & Me.txtSearch & "9"
            Case Is = 4
                'same as equal given the assumption that the account is a 4 digit number
                criteria = field & " = " & Me.txtSearch
        End Select
    Else
        If field = "phones" Then
            criteria = "WHERE (((Phones.Phone1) Like [instance])) OR (((Phones.Phone2) Like [instance])) OR (((Phones.Phone3) Like [instance]))"
            criteria = Replace(criteria, "[instance]", Chr(34) & Me.txtSearch & "%" & Chr(34))
        Else
            criteria = field & " like " & Chr(34) & Me.txtSearch & "%" & Chr(34)
        End If
    End If
    Case Is = 2 'contains
    If field = "account" Then
        Call MsgBox("Using the Contains criteria for an account is invalid. Please use either Starts With or Equals, or modify your other search criteria to get you the results you need.", vbExclamation, "Invalid Criteria")
        Exit Sub
    Else
        If field = "phones" Then
            criteria = "WHERE (((Phones.Phone1) Like [instance])) OR (((Phones.Phone2) Like [instance])) OR (((Phones.Phone3) Like [instance]))"
            criteria = Replace(criteria, "[instance]", Chr(34) & "%" & Me.txtSearch & "%" & Chr(34))
        Else
            criteria = field & " like " & Chr(34) & "%" & Me.txtSearch & "%" & Chr(34)
        End If
    End If
    Case Is = 3 'equals
    If field = "[account]" Then
        criteria = field & " = " & Me.txtSearch
    Else
        If field = "phones" Then
            criteria = "WHERE (((Phones.Phone1) like [instance])) OR (((Phones.Phone2) Like [instance])) OR (((Phones.Phone3) Like [instance]))"
            criteria = Replace(criteria, "Like [instance]", "=" & Chr(34) & Me.txtSearch & Chr(34))
        Else
            criteria = field & " = '" & Me.txtSearch & "'"
        End If
    End If
End Select

'select how the results should be returned.
Select Case Me.fraSortBy.value
    Case Is = 1 'sort by name ASC
    sortby = " order by name"
    Case Is = 2 'sort by account number ASC
    sortby = " order by account"
End Select

Dim lRecs As Long

If field = "phones" Then
qry = "INSERT INTO temp_results ( account, name )" & _
      " SELECT customer.account, customer.name" & _
      " FROM customer INNER JOIN Phones ON customer.account = Phones.CustomerID " & _
      criteria & sortby
Else
    qry = "insert into temp_results select account, name from customer where " & criteria & sortby
End If

qry = Replace(qry, "'", "''")

With cmd
    .ActiveConnection = CurrentProject.Connection
    .CommandType = adCmdText
    .CommandText = qry
    .Execute lRecs
    'Set rst = .Execute
End With

If lRecs = 0 Then
    Call MsgBox("Your current search has not produced any results. Please adjust your criteria or search string and then try again.", vbExclamation, "No Records Found")
    Exit Sub
End If


Me.temp_results_subform.Requery

   On Error GoTo 0
   Exit Sub

cmdSearch_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdSearch_Click of VBA Document Form_frmSearch")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdSearch_Click of VBA Document Form_frmSearch"
End Sub

Private Sub cmdSelect_Click()

   On Error GoTo cmdSelect_Click_Error
    If IsNull(frmSender) And IsEmpty(frmSender) Then
        'do nothing
    Else
        Dim acct As String
        acct = Me.temp_results_subform.Form.Controls.Item(0).value
        'DoCmd.OpenForm "Opt 4 Form", acNormal, , , , , "existing," & acct
        DoCmd.OpenForm frmSender, acNormal, , , , , acct
    End If
    
    DoCmd.Close acForm, Me.Name, acSaveNo
    
   On Error GoTo 0
   Exit Sub

cmdSelect_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdSelect_Click of VBA Document Form_frmSearch")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdSelect_Click of VBA Document Form_frmSearch"
End Sub

Private Sub Form_Load()
Dim qry As String
If IsNull(Me.OpenArgs) Or IsEmpty(Me.OpenArgs) Then
    frmSender = ""
Else
    frmSender = Me.OpenArgs
End If
'set defaults
   On Error GoTo Form_Load_Error

Me.fraSearchIn.DefaultValue = 2
Me.fraSearchFor.DefaultValue = 1
Me.fraSortBy.DefaultValue = 1

'delete all records in the temp_results table
qry = "delete * from temp_results"
CurrentProject.Connection.Execute qry
Me.temp_results_subform.Requery

   On Error GoTo 0
   Exit Sub

Form_Load_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Load of VBA Document Form_frmSearch")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Load of VBA Document Form_frmSearch"

End Sub

