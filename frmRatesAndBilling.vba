Option Compare Database
Option Explicit

Dim sWaterService As String
Dim cState As clsState

Private Sub cmdInquiry_Click()
Dim query As String
Dim rst As New ADODB.Recordset

   On Error GoTo cmdInquiry_Click_Error
   'was an account specified
    If IsEmpty(txtAccount) Or IsNull(txtAccount) Then
        Call MsgBox("Please specify a valid account number.", vbExclamation, "No Account")
        Exit Sub
    End If
    
    'is the account valid
    query = "SELECT account from customer where account = " & CLng(txtAccount)
    rst.Open query, CurrentProject.Connection
    If rst.BOF And rst.EOF Then
        Call MsgBox("Please specify a valid account number.", vbExclamation, "Invalid Account")
        Exit Sub
    End If
    
    'if we made it here, open the account inquiry form
    'passing in the account number
    DoCmd.OpenForm "Opt 4 Form", acNormal, , , , acWindowNormal, txtAccount
   
   On Error GoTo 0
   Exit Sub

cmdInquiry_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdInquiry_Click of VBA Document Form_frmRatesAndBilling")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdInquiry_Click of VBA Document Form_frmRatesAndBilling"
End Sub

Private Sub cmdNext_Click()
Dim account As Long
Dim qry As String
Dim rst As New ADODB.Recordset

    'was an account specified
   On Error GoTo cmdNext_Click_Error

    If IsEmpty(txtAccount) Or IsNull(txtAccount) Then
        'Call MsgBox("Please specify a valid account number.", vbExclamation, "No Account")
        Exit Sub
    End If
    
    account = GetNextRecord(Me.txtAccount)
    Call FillForm(account)

   On Error GoTo 0
   Exit Sub

cmdNext_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdNext_Click of VBA Document Form_frmRatesAndBilling")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdNext_Click of VBA Document Form_frmRatesAndBilling"

End Sub

Private Sub cmdNotes_Click()
   On Error GoTo cmdNotes_Click_Error

    If IsNull(Me.txtAccount) Or IsEmpty(Me.txtAccount) Then
        Call MsgBox("To use the notes feature, please select a valid account.", vbExclamation, "No Account")
        Exit Sub
    End If

    'test to see if it is open but not visible
    Dim bool As Boolean
    bool = CheckFormStatus("frmNotes")
    If bool Then
        DoCmd.Close acForm, "frmNotes", acSaveYes
    End If
    vOpenArgs = Me.txtAccount
    
    DoCmd.OpenForm "frmNotes", acNormal, , , acFormPropertySettings, acWindowNormal, vOpenArgs

   On Error GoTo 0
   Exit Sub

cmdNotes_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdNotes_Click of VBA Document Form_frmRatesAndBilling")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdNotes_Click of VBA Document Form_frmRatesAndBilling"

End Sub

Private Sub cmdPrev_Click()
Dim account As Long
Dim qry As String
Dim rst As New ADODB.Recordset

    'was an account specified
   On Error GoTo cmdPrev_Click_Error

    If IsEmpty(txtAccount) Or IsNull(txtAccount) Then
        'Call MsgBox("Please specify a valid account number.", vbExclamation, "No Account")
        Exit Sub
    End If
    
    account = GetPrevRecord(Me.txtAccount)
    Call FillForm(account)

   On Error GoTo 0
   Exit Sub

cmdPrev_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrev_Click of VBA Document Form_frmRatesAndBilling")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrev_Click of VBA Document Form_frmRatesAndBilling"

End Sub

Private Sub cmdQuit_Click()

   On Error GoTo cmdQuit_Click_Error
    DoCmd.Close , ""
   
   On Error GoTo 0
   Exit Sub

cmdQuit_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdQuit_Click of VBA Document Form_frmRatesAndBilling")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdQuit_Click of VBA Document Form_frmRatesAndBilling"
End Sub

Private Sub cmdSave_Click()
'write values back into their respective buckets
'this only applies to cboclass, cbowaterservice and the recurring values
Dim sClass As String
Dim service_id As Integer
Dim query As String
Dim lRecs As Long
Dim ix As Integer
Dim charge_id As Long
Dim rst As New ADODB.Recordset
Dim objListItem As ListItem

    'was an account specified
   On Error GoTo cmdSave_Click_Error

    If IsEmpty(txtAccount) Or IsNull(txtAccount) Then
        Call MsgBox("Please specify a valid account number.", vbExclamation, "No Account")
        Exit Sub
    End If

    If cState.account = Me.txtAccount Then
    
        cboClass.SetFocus
        
        'did the propert use change?
        If cState.PropertyUse <> cboClass.text Then
            Select Case cboClass.text
                Case "Residential"
                    sClass = "RE"
                Case "Commercial"
                    sClass = "CO"
            End Select
            cboWaterService.SetFocus
            service_id = GetServiceID(cboWaterService.text)
            query = "UPDATE customer set property_use = '" & sClass & "' WHERE account = " & Me.txtAccount
            CurrentProject.Connection.Execute query, lRecs
            If lRecs = 0 Then
                'could be an error - but only if a change was made that did not update
            End If
        End If
        
        'did the service change?
        cboWaterService.SetFocus
        If cState.Service <> cboWaterService.text Then
            query = "UPDATE CustomerServiceConnection set service_id = " & service_id & " WHERE account = " & Me.txtAccount
            CurrentProject.Connection.Execute query, lRecs
        
            If lRecs = 0 Then
                'could be an error - but only if a change was made that did not update
            End If
        End If
        
        'now find out which lstCharges listitems are checked and then add these values to the RatesAndCharges table
        'TODO - what if an item was unchecked - we now have to check for a missing item.
            For ix = 1 To Me.lstCharges.ListItems.Count - 1
                Set objListItem = Me.lstCharges.ListItems(ix)
                If objListItem.checked <> cState.Item(ix) Then
                    charge_id = GetChargeID(Me.lstCharges.ListItems(ix).text)
                    'does this exist in the RatesAndCharges table?
                    query = "SELECT id from RatesAndCharges WHERE recurring_charge_id = " & charge_id & " AND account = " & Me.txtAccount
                    rst.Open query, CurrentProject.Connection
                    If rst.BOF And rst.EOF And objListItem.checked = True Then
                        'it doesn't exist and the item is checked add it
                        lRecs = 0
                        query = "INSERT INTO RatesAndCharges(account,recurring_charge_id) VALUES(" & Me.txtAccount & "," & charge_id & ")"
                        CurrentProject.Connection.Execute query, lRecs
                        If lRecs = 0 Then
                            'there was an error
                            Err.Raise vbObjectError + 1001, "cmdSave_Click of frmRatesBilling", _
                                "There was an error adding a values to the RatesAndCharges Table: " & Me.txtAccount & " charge_id: " & charge_id
                        End If
                    Else
                        If cState.Item(ix) = True Then
                            'The item was checked but is not checked now so remove it
                            query = "DELETE FROM RatesAndCharges WHERE recurring_charge_id = " & charge_id & " AND account = " & Me.txtAccount
                        CurrentProject.Connection.Execute query, lRecs
                            If lRecs = 0 Then
                                'there was an error
                                Err.Raise vbObjectError + 1001, "cmdSave_Click of frmRatesBilling", _
                                    "There was an error deleting a values from the RatesAndCharges Table: " & Me.txtAccount & " charge_id: " & charge_id
                            End If
                        End If
                    End If
                End If
                If rst.state = adStateOpen Then rst.Close
            Next
    End If
    
   On Error GoTo 0
   Exit Sub

cmdSave_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdSave_Click of VBA Document Form_frmRatesAndBilling")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdSave_Click of VBA Document Form_frmRatesAndBilling"

End Sub

Private Sub cmdSearch_Click()

   On Error GoTo cmdSearch_Click_Error
    DoCmd.OpenForm "frmSearch", acNormal, , , , , "frmRatesAndBilling"
    DoCmd.Close acForm, Me.Name, acSaveNo
   On Error GoTo 0
   Exit Sub

cmdSearch_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdSearch_Click of VBA Document Form_frmRatesAndBilling")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdSearch_Click of VBA Document Form_frmRatesAndBilling"
End Sub

Private Sub Form_Load()
Dim ix As Integer
Dim query As String
Dim rst As New ADODB.Recordset
Dim listArray As Variant
Dim objListItem As ListItem
Dim ColHeaders As ColumnHeaders
Dim colH1 As ColumnHeader
Dim colH2 As ColumnHeader
Dim colH3 As ColumnHeader
'we should be getting an account
   On Error GoTo Form_Load_Error

    If IsNull(Me.OpenArgs) Or IsEmpty(Me.OpenArgs) Then
        'do nothing
    Else
        Me.txtAccount = Me.OpenArgs
    End If

    'load all values for the cboWaterService control
    If cboWaterService.ListCount > 0 Then
        Do While cboWaterService.ListCount > 0
            cboWaterService.RemoveItem 0
        Loop
    End If
    
    query = "SELECT ServiceConnections.service_id, ServiceConnections.Description FROM ServiceConnections;"
    rst.Open query, CurrentProject.Connection
    If rst.BOF And rst.EOF Then
        'there is a problem
        Exit Sub
    Else
        Do While Not rst.EOF
            cboWaterService.AddItem rst.Fields(1).value
            rst.MoveNext
        Loop
    End If
    rst.Close
    
    query = "SELECT RecurringCharges.charge_description, RecurringCharges.charge_code, RecurringCharges.charge_amount" & _
            " FROM RecurringCharges;"

    rst.Open query, CurrentProject.Connection

    If rst.BOF And rst.EOF Then
        'do nothing
    Else
    'One method is to iterate through the recordset and add an instance of a ListItem object to the
    'ListItems collection of the ListView for each record. The value for the first column would be
    'assigned to the Text argument when adding the ListItem and the values for the other columns
    'would be added to the SubItems(n) collection. You will need to make sure that you have the
    'ColumnHeaders defined in order to have each column available, and use the Report-View setting
    'for the display.
        'add column headers
        Set colH1 = Me.lstCharges.ColumnHeaders.Add(1, "Description", "Description")
        With colH1
            .Width = Me.lstCharges.Width * 0.66
        End With
        
        Set colH2 = Me.lstCharges.ColumnHeaders.Add(2, "Amount", "Amount")
        With colH2
            .Width = Me.lstCharges.Width * 0.3
        End With
        
'        Set colH3 = Me.lstCharges.ColumnHeaders.Add(3, "Selected", "Selected")
'        With colH3
'            .Width = Me.lstCharges.Width * 0.1
'        End With
        
        'set the report-view mode
        Me.lstCharges.View = 3
        Me.lstCharges.Checkboxes = True
        
        Do While Not rst.EOF
            Set objListItem = Me.lstCharges.ListItems.Add(text:=rst.Fields(0).value & " (" & rst.Fields(1).value & ")")
                objListItem.ListSubItems.Add , , Format(rst.Fields(2).value, "Currency")
                objListItem.checked = False
                
            rst.MoveNext
        Loop
    End If
    
   On Error GoTo 0
   Exit Sub

Form_Load_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Load of VBA Document Form_frmRatesAndBilling")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Load of VBA Document Form_frmRatesAndBilling"
End Sub

Private Sub txtAccount_LostFocus()
   
   On Error GoTo txtAccount_LostFocus_Error
   
    'was an account specified
    If IsEmpty(txtAccount) Or IsNull(txtAccount) Then
        'Call MsgBox("Please specify a valid account number.", vbExclamation, "No Account")
        Exit Sub
    End If
    Call FillForm(Me.txtAccount)
    txtAccountName.SetFocus
    
   On Error GoTo 0
   Exit Sub

txtAccount_LostFocus_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure txtAccount_LostFocus of VBA Document Form_frmRatesAndBilling")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure txtAccount_LostFocus of VBA Document Form_frmRatesAndBilling"
End Sub

Private Sub cboWaterService_Change()
    If cboWaterService.text = sWaterService Then
        'do nothing
    Else
        'the user has made a change to the DropDown box
        'do something
    End If
End Sub

Private Sub FillForm(acct As Long)

Dim query As String
Dim rst As New ADODB.Recordset
Dim listArray As Variant
Dim ix As Integer
Dim objListItem As ListItem

    'is the account valid
   On Error GoTo FillForm_Error

    query = "SELECT * from customer where account = " & acct
    rst.Open query, CurrentProject.Connection
    If rst.BOF And rst.EOF Then
        Call MsgBox("Please specify a valid account number.", vbExclamation, "Invalid Account")
        Exit Sub
    End If
    
    'set the account name value
    txtAccount = acct
    txtAccountName = rst.Fields("name").value
    
    cboClass.SetFocus
    If rst.Fields("property_use").value = "RE" Then
        cboClass.ListIndex = 1
    Else
        cboClass.ListIndex = 0
    End If

    rst.Close
    
    'now select their water service
    query = "SELECT ServiceConnections.Description" & _
            " FROM CustomerServiceConnection INNER JOIN ServiceConnections ON " & _
            " CustomerServiceConnection.service_id = ServiceConnections.service_id" & _
            " WHERE (((CustomerServiceConnection.account)=" & acct & "));"
    
    rst.Open query, CurrentProject.Connection
    If rst.BOF And rst.EOF Then
        'cboWaterService.ListIndex = -1
    Else
        'we have to find the corrosponding listindex
        cboWaterService.SetFocus
        cboWaterService.SelText = rst.Fields(0).value
    End If
    
    rst.Close
    
    'now get the account status, balance and reading
    query = "SELECT customer.total_due, customer.service_discon, (select top 1 MeterReads.posted " & _
            " from MeterReads where account = " & acct & " order by batch_date) AS Posted" & _
            " FROM customer WHERE (((customer.account)=" & acct & "));"

    rst.Open query, CurrentProject.Connection
    
    If rst.BOF And rst.EOF Then
        Me.txtBalance = 0
        Me.txtStatus = "Normal"
        Me.txtReading = "None"
    Else
        Me.txtBalance = IIf(IsNull(rst.Fields(0).value), 0, Round(rst.Fields(0).value, 2))
        Me.txtStatus = IIf(IsNull(rst.Fields(1).value), "Normal", IIf(rst.Fields(1).value = True, "Disconnected", "Normal"))
        Me.txtReading = IIf(IsNull(rst.Fields(2).value), "None", IIf(rst.Fields(2).value = "Y", "Posted", "Unposted"))
    End If
    
    rst.Close
    'now load up the recurring charges -- loop through the charges that apply and select these
    'first reset all items to unchecked
    For ix = 1 To Me.lstCharges.ListItems.Count - 1
        Set objListItem = Me.lstCharges.ListItems(ix)
        objListItem.checked = False
    Next
    
    query = "SELECT RecurringCharges.charge_description, RecurringCharges.charge_code" & _
            " FROM RecurringCharges INNER JOIN RatesAndCharges ON RecurringCharges.charge_id = RatesAndCharges.recurring_charge_id" & _
            " WHERE (((RatesAndCharges.account)=" & Me.txtAccount & "));"

    rst.Open query, CurrentProject.Connection
    
    If rst.BOF And rst.EOF Then
        'do nothing
    Else
        Do While Not rst.EOF
            For ix = 1 To Me.lstCharges.ListItems.Count - 1
                If Me.lstCharges.ListItems(ix).text = rst.Fields(0).value & " (" & rst.Fields(1).value & ")" Then
                    Set objListItem = Me.lstCharges.ListItems(ix)
                    objListItem.checked = True
                    Exit For
                End If
            Next
            rst.MoveNext
        Loop
    End If
    
    'now save the state of all items into a class
    Set cState = New clsState
    For ix = 1 To Me.lstCharges.ListItems.Count - 1
        Set objListItem = Me.lstCharges.ListItems(ix)
        cState.Add CLng(ix), objListItem.checked
    Next
    cState.account = Me.txtAccount
    
    cboWaterService.SetFocus
    sWaterService = cboWaterService.text
    Me.txtAccountName.SetFocus
    
   On Error GoTo 0
   Exit Sub

FillForm_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure FillForm of VBA Document Form_frmRatesAndBilling")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure FillForm of VBA Document Form_frmRatesAndBilling"

End Sub

Private Function GetServiceID(Service As String) As Integer
Dim query As String
Dim rst As New ADODB.Recordset

   On Error GoTo GetServiceID_Error
    'service = FixQuotes(service)
    
    query = "select service_id from ServiceConnections WHERE [Description] = '" & Service & "'"
    rst.Open query, CurrentProject.Connection
    
    If rst.BOF And rst.EOF Then
        GetServiceID = -1
    Else
        GetServiceID = IIf(IsNull(rst.Fields(0).value), 0, rst.Fields(0).value)
    End If

   On Error GoTo 0
   Exit Function

GetServiceID_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure GetServiceID of VBA Document Form_frmRatesAndBilling")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure GetServiceID of VBA Document Form_frmRatesAndBilling"

End Function

Private Function GetChargeID(Charge As String) As Integer
Dim query As String
Dim rst As New ADODB.Recordset
Dim sTempChg As String
Dim iPos1 As Integer
Dim iPos2 As Integer

   On Error GoTo GetChargeID_Error

    iPos1 = InStr(Charge, " (")
    'iPos2 =
    sTempChg = Left(Charge, iPos1 - 1)

    query = "select charge_id from RecurringCharges WHERE [charge_description] = '" & sTempChg & "'"
    rst.Open query, CurrentProject.Connection
    
    If rst.BOF And rst.EOF Then
        GetChargeID = -1
    Else
        GetChargeID = IIf(IsNull(rst.Fields(0).value), 0, rst.Fields(0).value)
    End If

   On Error GoTo 0
   Exit Function

GetChargeID_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure GetChargeID of VBA Document Form_frmRatesAndBilling")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure GetChargeID of VBA Document Form_frmRatesAndBilling"

End Function
