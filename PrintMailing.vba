Option Explicit
Dim strSender As String

Private Enum PrintTab
    ByAccount
    ByRoute
    BySelected
End Enum

Public x As Integer

Private Sub cboRateTable_Change()

   On Error GoTo cboRateTable_Change_Error
    Call MsgBox("This feature is not currently available.", vbExclamation, "Not Available")
   On Error GoTo 0
   Exit Sub

cboRateTable_Change_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cboRateTable_Change of VBA Document Form_PrintMailing")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cboRateTable_Change of VBA Document Form_PrintMailing"
End Sub

Private Sub cboService_Change()

   On Error GoTo cboService_Change_Error
        Call MsgBox("This feature is not currently available.", vbExclamation, "Not Available")

   On Error GoTo 0
   Exit Sub

cboService_Change_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cboService_Change of VBA Document Form_PrintMailing")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cboService_Change of VBA Document Form_PrintMailing"
End Sub

Private Sub cmdExit_Click()

   On Error GoTo cmdExit_Click_Error

    If strSender = "" Then
        DoCmd.OpenForm "PrintLabelsMenu", acNormal
    Else
        DoCmd.OpenForm strSender, acNormal
    End If
    
    DoCmd.Close acForm, Me.Name, acSaveYes
    
   On Error GoTo 0
   Exit Sub

cmdExit_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdExit_Click of VBA Document Form_PrintMailing")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdExit_Click of VBA Document Form_PrintMailing"
End Sub

Private Sub cmdPrint_Click()

Dim query As String
Dim options As Integer
Dim reportType As Integer
Dim rpt As Report
Dim strWhere As String
Dim strOrderBy As String

' chkAccountNum.value = checked         A
' chkAccountNum.value = unchecked       B
' chkPhoneNum.value = checked           C
' chkPhoneNum.value = unchecked         D
'A & C OR A & D OR B & C OR B & D

   On Error GoTo cmdPrint_Click_Error

If chkAccountNum.value = -1 And chkPhoneNum.value = -1 Then
    'both phone and account number must appear on the label
    options = 1
ElseIf chkAccountNum.value = -1 And (IsNull(chkPhoneNum.value) Or chkPhoneNum = 0) Then
    'only the account number must appear on the label
    options = 2
ElseIf (IsNull(chkAccountNum.value) Or chkAccountNum = 0) And chkPhoneNum.value = -1 Then
    'only the phone number must appear on the label
    options = 3
Else
    'chkPhoneNum.value = 0 And chkPhoneNum.value = 0
    'print the label with no account or phone number
    options = 4
End If

'then Get which tab is active
Select Case Me.TabCtl0.value
    Case 0  'This prints all accounts
        'while this feature is being built it will not be available
        'Call MsgBox("This feature is not currently available.", vbExclamation, "Not Available")
        'Exit Sub
        'first we need to figure out are we printing by account number, last name or by zip code
        'then we need to know are we printing all accounts or only by a range (start to finish)
        Select Case Me.cboSelectBy.value
            Case "Account Number"
                Select Case fraSelectBy
                    Case 1  'print all accounts
                        strWhere = "WHERE customer.account is not null"
                    Case 2  'select only a range of account number
                        If IsNull(Me.txtRangeFrom) Or IsEmpty(Me.txtRangeFrom) Or IsNull(Me.txtRangeTo) Or IsEmpty(Me.txtRangeTo) Then
                            Call MsgBox("Either the Range from or the Range to is empty. Please select a valid range and then try again.", vbExclamation, "Invalid Range")
                            Exit Sub
                        End If
                        If IsNumeric(Me.txtRangeFrom) And IsNumeric(Me.txtRangeTo) Then
                            If Me.txtRangeFrom > Me.txtRangeTo Then
                                Call MsgBox("The Range from cannot be larger than the Range to. Please select a valid range and then try again.", vbExclamation, "Invalid Range")
                                Exit Sub
                            End If
                        Else
                            'not a valid number
                            Call MsgBox("The Range from or Range to contains an invalid account type number (must be a numeric number for both from and to). " & _
                            " Please select a valid account range and then try again.", vbExclamation, "Invalid Range")
                            Exit Sub
                        End If
                        
                        strWhere = "WHERE (customer.account >= " & Me.txtRangeFrom & " AND customer.account <= " & Me.txtRangeTo & ")"
                        strOrderBy = " ORDER BY customer.account"
                End Select
            Case "Last Name"
                Select Case fraSelectBy
                    Case 1  'print all accounts
                        strWhere = "WHERE [account] is not null"
                    Case 2  'print only a range of names
                        If IsNumeric(Me.txtRangeFrom) And IsNumeric(Me.txtRangeTo) Then
                            'not a valid number
                             Call MsgBox("You are searching by last name. Please select a valid range type for this selection and then try again.", vbExclamation, "Invalid Range Type")
                                Exit Sub
                            End If
                            
                        'not a numeric value Take the first character of from and to,
                        'convert to an ascii numeric value and then compare
                        Dim a As Integer
                        Dim b As Integer
                        a = Asc(Left(Me.txtRangeFrom, 1))
                        b = Asc(Left(Me.txtRangeTo, 1))
                        
                        If a > b Then
                            Call MsgBox("The Range from cannot contain a higher character value than the Range to (e.g. the letter B comes after the letter A)" & _
                                ". Please select a valid range and then try again.", vbExclamation, "Invalid Range")
                            Exit Sub
                        End If
                        
                        strWhere = "WHERE (left(name,1) between '" & Left(Me.txtRangeFrom, 1) & "' and '" & Left(Me.txtRangeTo, 1) & "')"
                        strOrderBy = " ORDER BY name"
                        
                End Select
            Case "Zip Code"
                Select Case fraSelectBy
                    Case 1  'print all zip codes
                        strWhere = ""
                    Case 2  'select only a range of zip codes
                        If IsNull(Me.txtRangeFrom) Or IsEmpty(Me.txtRangeFrom) Or IsNull(Me.txtRangeTo) Or IsEmpty(Me.txtRangeTo) Then
                            Call MsgBox("Either the Range from or the Range to is empty. Please select a valid range and then try again.", vbExclamation, "Invalid Range")
                            Exit Sub
                        End If
                        If IsNumeric(Me.txtRangeFrom) And IsNumeric(Me.txtRangeTo) Then
                            If Me.txtRangeFrom > Me.txtRangeTo Then
                                Call MsgBox("The Range from cannot be larger than the Range to. Please select a valid range and then try again.", vbExclamation, "Invalid Range")
                                Exit Sub
                            End If
                        Else
                            'not a valid number
                            Call MsgBox("The Range from or Range to contains an invalid zip code value (must be a numeric value, without the plus four) for both from and to). " & _
                            " Please select a valid zip code range and then try again.", vbExclamation, "Invalid Range")
                            Exit Sub
                        End If
                        
                        strWhere = "WHERE (left(zip,5) between '" & Left(Me.txtRangeFrom, 5) & "' and '" & Left(Me.txtRangeTo, 5) & "')"
                        strOrderBy = " ORDER BY zip"
                        
                End Select
        End Select
        
        'then we need to know are we printing all balances, only positive bals or negative bals or zero bals
            If InStr(strWhere, "WHERE") > 0 Then
            
            Else
                strWhere = "WHERE " & strWhere
            End If
        
        Select Case cboBalances
            Case "All Balances"
                strWhere = strWhere & ""
            Case "Positive only"
                strWhere = strWhere & " AND ([current_due] > 0 or [total_due] > 0)"
            Case "Negative only"
                strWhere = strWhere & " AND ([current_due] < 0 or [total_due] < 0)"
            Case "Zero"
                strWhere = strWhere & " AND ([current_due] = 0 or [total_due] = 0)"
        End Select
        
        'finally we need to know if we're including only active accounts, in active accounts or both
        If Me.chkActive = True And Me.chkInactive = True Then
            strWhere = strWhere & " AND ([status] = 'A' or [status] = 'I')"
        ElseIf Me.chkActive = True And (Me.chkInactive = False Or IsNull(Me.chkInactive)) Then
            strWhere = strWhere & " AND ([status] = 'A')"
        ElseIf (Me.chkActive = False Or IsNull(Me.chkActive)) And Me.chkInactive = True Then
            strWhere = strWhere & " AND ([status] = 'I')"
        Else
            'both are false - should get us nothing
            strWhere = strWhere & " AND ([status] <> 'A' or [status] <> 'I')"
        End If
        
        Select Case Me.fraWhereToSend.value
        Case 1  'Billing address
            Select Case options 'there are 4 possible options
            Case 2  'only the account number must appear on the label
                query = " SELECT customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'','OCCUPANT'," & _
                        " [bill_name] Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]='')," & _
                        " 'OCCUPANT',[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or " & _
                        " [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT'," & _
                        " [name] Is Not Null Or [name]<>'',[name]) AS CustName," & _
                        " iif(customer.addr1 is null or customer.addr1 =''," & _
                        " customer.phy_address,customer.addr1) as address, customer.city, customer.state, " & _
                        " customer.zip from customer " & strWhere & strOrderBy
            Case 4  'print the label with no account or phone number
                query = " SELECT customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'','OCCUPANT'," & _
                        " [bill_name] Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]='')," & _
                        " 'OCCUPANT',[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or " & _
                        " [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT'," & _
                        " [name] Is Not Null Or [name]<>'',[name]) AS CustName," & _
                        " iif(customer.addr1 is null or customer.addr1 =''," & _
                        " customer.phy_address,customer.addr1) as address, customer.city, customer.state, " & _
                        " customer.zip from customer " & strWhere & strOrderBy
            Case Else
                'default to with account (OPTION = 2)
                query = " SELECT customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'','OCCUPANT'," & _
                        " [bill_name] Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]='')," & _
                        " 'OCCUPANT',[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or " & _
                        " [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT'," & _
                        " [name] Is Not Null Or [name]<>'',[name]) AS CustName," & _
                        " iif(customer.addr1 is null or customer.addr1 =''," & _
                        " customer.phy_address,customer.addr1) as address, customer.city, customer.state," & _
                        " customer.zip from customer " & strWhere & strOrderBy
            End Select
        Case 2  'Service Location is selected
            Select Case options
                Case 1
                    If Len(strWhere) > 0 Then strWhere = Replace(strWhere, "WHERE", "AND")
                    If Len(strOrderBy) > 0 Then strOrderBy = Replace(strOrderBy, "name", "IIf([name]=" & Chr(34) & _
                        "UNKNOWN" & Chr(34) & "," & Chr(34) & "OCCUPANT" & Chr(34) & ",[name])")
                    query = "SELECT customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'','OCCUPANT'," & _
                        " [bill_name] Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]='')," & _
                        " 'OCCUPANT',[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or " & _
                        " [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT'," & _
                        " [name] Is Not Null Or [name]<>'',[name]) AS CustName," & _
                        " customer.phy_address as address, 'Sacramento' as city, 'CA' as state, " & _
                        " IIf(IsNull([phy_zipcode]) Or IsEmpty([phy_zipcode]),95821,[phy_zipcode]) AS ZIP" & _
                        " FROM physicalzip INNER JOIN customer ON physicalzip.account = customer.account" & _
                        " where ((customer.phy_address is not null) or (customer.phy_address <> '')) " & strWhere & strOrderBy
                Case 2
                    If Len(strWhere) > 0 Then strWhere = Replace(strWhere, "WHERE", "AND")
                    If Len(strOrderBy) > 0 Then strOrderBy = Replace(strOrderBy, "name", "IIf([name]=" & Chr(34) & _
                        "UNKNOWN" & Chr(34) & "," & Chr(34) & "OCCUPANT" & Chr(34) & ",[name])")
                query = "SELECT customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'','OCCUPANT'," & _
                        " [bill_name] Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]='')," & _
                        " 'OCCUPANT',[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or " & _
                        " [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT'," & _
                        " [name] Is Not Null Or [name]<>'',[name]) AS CustName," & _
                        " customer.phy_address as address, 'Sacramento' as city, 'CA' as state, " & _
                        " IIf(IsNull([phy_zipcode]) Or IsEmpty([phy_zipcode]),95821,[phy_zipcode]) AS ZIP" & _
                        " FROM physicalzip INNER JOIN customer ON physicalzip.account = customer.account" & _
                        " where ((customer.phy_address is not null) or (customer.phy_address <> '')) " & strWhere & strOrderBy
                Case Else
                    If Len(strWhere) > 0 Then strWhere = Replace(strWhere, "WHERE", "AND")
                    If Len(strOrderBy) > 0 Then strOrderBy = Replace(strOrderBy, "name", "IIf([name]=" & Chr(34) & "UNKNOWN" & Chr(34) & "," & Chr(34) & "OCCUPANT" & Chr(34) & ",[name])")
                    query = "SELECT customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'','OCCUPANT'," & _
                        " [bill_name] Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]='')," & _
                        " 'OCCUPANT',[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or " & _
                        " [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT'," & _
                        " [name] Is Not Null Or [name]<>'',[name]) AS CustName," & _
                        " customer.phy_address as address, 'Sacramento' as city, 'CA' as state, " & _
                        " IIf(IsNull([phy_zipcode]) Or IsEmpty([phy_zipcode]),95821,[phy_zipcode]) AS ZIP" & _
                        " FROM physicalzip INNER JOIN customer ON physicalzip.account = customer.account" & _
                        " where ((customer.phy_address is not null) or (customer.phy_address <> '')) " & strWhere & strOrderBy
            End Select
        Case Else
            'default to service location
            Select Case options
                Case 2
                    If Len(strWhere) > 0 Then strWhere = Replace(strWhere, "where", "and")
                    If Len(strOrderBy) > 0 Then strOrderBy = Replace(strOrderBy, "name", "IIf([name]=" & Chr(34) & _
                        "UNKNOWN" & Chr(34) & "," & Chr(34) & "OCCUPANT" & Chr(34) & ",[name])")
                    query = "SELECT customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'','OCCUPANT'," & _
                        " [bill_name] Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]='')," & _
                        " 'OCCUPANT',[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or " & _
                        " [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT'," & _
                        " [name] Is Not Null Or [name]<>'',[name]) AS CustName," & _
                        " customer.phy_address as address, 'Sacramento' as city, 'CA' as state, " & _
                        " IIf(IsNull([phy_zipcode]) Or IsEmpty([phy_zipcode]),95821,[phy_zipcode]) AS ZIP" & _
                        " FROM physicalzip INNER JOIN customer ON physicalzip.account = customer.account" & _
                        " where ((customer.phy_address is not null) or (customer.phy_address <> '')) " & strWhere & strOrderBy
                Case 4
                If Len(strWhere) > 0 Then strWhere = Replace(strWhere, "where", "and")
                If Len(strOrderBy) > 0 Then strOrderBy = Replace(strOrderBy, "name", "IIf([name]=" & Chr(34) & "UNKNOWN" & Chr(34) & "," & Chr(34) & "OCCUPANT" & Chr(34) & ",[name])")
                query = "SELECT Switch([bill_name] Is Not Null And [bill_name]<>'','OCCUPANT'," & _
                        " [bill_name] Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]='')," & _
                        " 'OCCUPANT',[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or " & _
                        " [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT'," & _
                        " [name] Is Not Null Or [name]<>'',[name]) AS CustName," & _
                        " customer.phy_address as address, 'Sacramento' as city, 'CA' as state, " & _
                        " IIf(IsNull([phy_zipcode]) Or IsEmpty([phy_zipcode]),95821,[phy_zipcode]) AS ZIP" & _
                        " FROM physicalzip INNER JOIN customer ON physicalzip.account = customer.account" & _
                        " where ((customer.phy_address is not null) or (customer.phy_address <> ''))  " & strWhere & strOrderBy
                Case Else   'option = 2
                    If Len(strWhere) > 0 Then strWhere = Replace(strWhere, "where", "and")
                    If Len(strOrderBy) > 0 Then strOrderBy = Replace(strOrderBy, "name", "IIf([name]=" & Chr(34) & "UNKNOWN" & Chr(34) & "," & Chr(34) & "OCCUPANT" & Chr(34) & ",[name])")
                    query = "SELECT customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'','OCCUPANT'," & _
                        " [bill_name] Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]='')," & _
                        " 'OCCUPANT',[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or " & _
                        " [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT'," & _
                        " [name] Is Not Null Or [name]<>'',[name]) AS CustName," & _
                        " customer.phy_address as address, 'Sacramento' as city, 'CA' as state, " & _
                        " IIf(IsNull([phy_zipcode]) Or IsEmpty([phy_zipcode]),95821,[phy_zipcode]) AS ZIP" & _
                        " FROM physicalzip INNER JOIN customer ON physicalzip.account = customer.account" & _
                        " where ((customer.phy_address is not null) or (customer.phy_address <> '')) " & strWhere & strOrderBy
            End Select
        End Select
    Case 1 'prints by route
        'MsgBox "Not yet implemented", vbInformation + vbOKOnly, "Not Implemented"
        'Exit Sub
        Select Case Me.fraWhereToSend.value
        Case 1  'billing addresses
            Select Case options
                Case 2
                    query = "SELECT CustomerRoutes.sequence, customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'' And" & _
                        " [care_of] Is Not Null And [care_of]<>'',[bill_name] & ' ' & [care_of],[bill_name]" & _
                        " Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]=''),[bill_name]," & _
                        " [care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or [bill_name]='')," & _
                        " [care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT',[name] Is Not Null" & _
                        " Or [name]<>'',[name]) AS CustName," & _
                        " iif(customer.addr1 is null or customer.addr1 = '', customer.phy_address, customer.addr1) as address, " & _
                        " customer.city, customer.state, customer.zip" & _
                        " FROM customer INNER JOIN CustomerRoutes ON customer.account = CustomerRoutes.account_num" & _
                        " WHERE (CustomerRoutes.route_id=" & Me.cboRoutes.Column(1) & ") ORDER BY CustomerRoutes.sequence;"
                Case 4
                    query = "SELECT CustomerRoutes.sequence, customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'' And" & _
                        " [care_of] Is Not Null And [care_of]<>'',[bill_name] & ' ' & [care_of],[bill_name]" & _
                        " Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]=''),[bill_name]," & _
                        " [care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or [bill_name]='')," & _
                        " [care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT',[name] Is Not Null" & _
                        " Or [name]<>'',[name]) AS CustName," & _
                        " iif(customer.addr1 is null or customer.addr1 = '', customer.phy_address, customer.addr1) as address, " & _
                        " customer.city, customer.state, customer.zip" & _
                        " FROM customer INNER JOIN CustomerRoutes ON customer.account = CustomerRoutes.account_num" & _
                        " WHERE (CustomerRoutes.route_id=" & Me.cboRoutes.Column(1) & ") ORDER BY CustomerRoutes.sequence;"
            End Select
        Case 2  'service locations
            Select Case options
                Case 2
                    query = "SELECT customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'','OCCUPANT'," & _
                        " [bill_name] Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]='')," & _
                        " 'OCCUPANT',[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or " & _
                        " [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT'," & _
                        " [name] Is Not Null Or [name]<>'',[name]) AS CustName," & _
                        " customer.phy_address as address, 'Sacramento' as city, 'CA' as state, " & _
                        " IIf(IsNull([phy_zipcode]) Or IsEmpty([phy_zipcode]),95821,[phy_zipcode]) AS ZIP" & _
                        " FROM physicalzip INNER JOIN (customer INNER JOIN CustomerRoutes ON customer.account = " & _
                        " CustomerRoutes.account_num) ON physicalzip.account = customer.account" & _
                        " WHERE (((customer.phy_address) Is Not Null Or (customer.phy_address)<>'') AND ((CustomerRoutes.route_id)=" & _
                        Me.cboRoutes.Column(1) & ")) ORDER BY CustomerRoutes.sequence;"
                Case 4
                query = "SELECT Switch([bill_name] Is Not Null And [bill_name]<>'','OCCUPANT'," & _
                        " [bill_name] Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]='')," & _
                        " 'OCCUPANT',[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or " & _
                        " [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT'," & _
                        " [name] Is Not Null Or [name]<>'',[name]) AS CustName," & _
                        " customer.phy_address as address, 'Sacramento' as city, 'CA' as state, " & _
                        " IIf(IsNull([phy_zipcode]) Or IsEmpty([phy_zipcode]),95821,[phy_zipcode]) AS ZIP" & _
                        " FROM physicalzip INNER JOIN (customer INNER JOIN CustomerRoutes ON customer.account = " & _
                        " CustomerRoutes.account_num) ON physicalzip.account = customer.account" & _
                        " WHERE (((customer.phy_address) Is Not Null Or (customer.phy_address)<>'') AND ((CustomerRoutes.route_id)=" & _
                        Me.cboRoutes.Column(1) & ")) ORDER BY CustomerRoutes.sequence;"
                Case Else   'option = 2
                    query = "SELECT customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'','OCCUPANT'," & _
                        " [bill_name] Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]='')," & _
                        " 'OCCUPANT',[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or " & _
                        " [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT'," & _
                        " [name] Is Not Null Or [name]<>'',[name]) AS CustName," & _
                        " customer.phy_address as address, 'Sacramento' as city, 'CA' as state, " & _
                        " IIf(IsNull([phy_zipcode]) Or IsEmpty([phy_zipcode]),95821,[phy_zipcode]) AS ZIP" & _
                        " FROM physicalzip INNER JOIN (customer INNER JOIN CustomerRoutes ON customer.account = " & _
                        " CustomerRoutes.account_num) ON physicalzip.account = customer.account" & _
                        " WHERE (((customer.phy_address) Is Not Null Or (customer.phy_address)<>'') AND ((CustomerRoutes.route_id)=" & _
                        Me.cboRoutes.Column(1) & ")) ORDER BY CustomerRoutes.sequence;"
            End Select
        End Select
    Case 2  'prints selected accounts
    'find out which accounts were selected
        Select Case Me.fraWhereToSend.value
        Case 1
            Select Case options
                Case 2
                    query = "SELECT customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'' And" & _
                        " [care_of] Is Not Null And [care_of]<>'',[bill_name] & ' ' & [care_of],[bill_name]" & _
                        " Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]=''),[bill_name]," & _
                        " [care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or [bill_name]='')," & _
                        " [care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT',[name] Is Not Null" & _
                        " Or [name]<>'',[name]) AS CustName, iif(customer.addr1 is null or customer.addr1 = ''," & _
                        " customer.phy_address, customer.addr1) as address," & _
                        " customer.city, customer.state, customer.zip " & _
                        " FROM customer WHERE (customer.account in (SELECT DISTINCTROW [cust_acct] from [temp_Customers]))"
                Case 4
                    query = "SELECT customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'' And" & _
                        " [care_of] Is Not Null And [care_of]<>'',[bill_name] & ' ' & [care_of],[bill_name]" & _
                        " Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]=''),[bill_name]," & _
                        " [care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or [bill_name]='')," & _
                        " [care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT',[name] Is Not Null" & _
                        " Or [name]<>'',[name]) AS CustName, iif(customer.addr1 is null or customer.addr1 = ''," & _
                        " customer.phy_address, customer.addr1) as address," & _
                        " customer.city, customer.state, customer.zip " & _
                        " FROM customer WHERE (customer.account in (SELECT DISTINCTROW [cust_acct] from [temp_Customers]))"
                Case Else 'option = 2
                    query = "SELECT customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'' And" & _
                        " [care_of] Is Not Null And [care_of]<>'',[bill_name] & ' ' & [care_of],[bill_name]" & _
                        " Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]=''),[bill_name]," & _
                        " [care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or [bill_name]='')," & _
                        " [care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT',[name] Is Not Null" & _
                        " Or [name]<>'',[name]) AS CustName, iif(customer.addr1 is null or customer.addr1 = ''," & _
                        " customer.phy_address, customer.addr1) as address," & _
                        " customer.city, customer.state, customer.zip " & _
                        " FROM customer WHERE (customer.account in (SELECT DISTINCTROW [cust_acct] from [temp_Customers]))"
            End Select
        Case 2
        'optServLoc is selected
            Select Case options
                Case 2
                    query = "SELECT customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'','OCCUPANT'," & _
                        " [bill_name] Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]='')," & _
                        " 'OCCUPANT',[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or " & _
                        " [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT'," & _
                        " [name] Is Not Null Or [name]<>'',[name]) AS CustName," & _
                        " customer.phy_address as address, 'Sacramento' as city, 'CA' as state, " & _
                        " IIf(IsNull([phy_zipcode]) Or IsEmpty([phy_zipcode]),95821,[phy_zipcode]) AS ZIP" & _
                        " FROM physicalzip INNER JOIN customer ON physicalzip.account = customer.account" & _
                        " WHERE customer.account in (SELECT DISTINCTROW [cust_acct] from [temp_Customers])" & _
                        " AND ((customer.phy_address is not null) or (customer.phy_address <> ''))"
                Case 4
                    query = "SELECT Switch([bill_name] Is Not Null And [bill_name]<>'','OCCUPANT'," & _
                        " [bill_name] Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]='')," & _
                        " 'OCCUPANT',[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or " & _
                        " [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT'," & _
                        " [name] Is Not Null Or [name]<>'',[name]) AS CustName," & _
                        " customer.phy_address as address, 'Sacramento' as city, 'CA' as state, " & _
                        " customer.account, IIf(IsNull([phy_zipcode]) Or IsEmpty([phy_zipcode]),95821,[phy_zipcode]) AS ZIP" & _
                        " FROM physicalzip INNER JOIN customer ON physicalzip.account = customer.account" & _
                        " WHERE customer.account in (SELECT DISTINCTROW [cust_acct] from [temp_Customers])" & _
                        " AND ((customer.phy_address is not null) or (customer.phy_address <> ''))"
                Case Else 'option 2
                    query = "SELECT customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'','OCCUPANT'," & _
                        " [bill_name] Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]='')," & _
                        " 'OCCUPANT',[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or " & _
                        " [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT'," & _
                        " [name] Is Not Null Or [name]<>'',[name]) AS CustName," & _
                        " customer.phy_address as address, 'Sacramento' as city, 'CA' as state, " & _
                        " IIf(IsNull([phy_zipcode]) Or IsEmpty([phy_zipcode]),95821,[phy_zipcode]) AS ZIP" & _
                        " FROM physicalzip INNER JOIN customer ON physicalzip.account = customer.account" & _
                        " WHERE customer.account in (SELECT DISTINCTROW [cust_acct] from [temp_Customers])" & _
                        " AND ((customer.phy_address is not null) or (customer.phy_address <> ''))"
            End Select
        Case Else
            'default to service location
            Select Case options
                Case 2
                    query = "SELECT customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'','OCCUPANT'," & _
                        " [bill_name] Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]='')," & _
                        " 'OCCUPANT',[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or " & _
                        " [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT'," & _
                        " [name] Is Not Null Or [name]<>'',[name]) AS CustName," & _
                        " customer.phy_address as address, 'Sacramento' as city, 'CA' as state, " & _
                        " IIf(IsNull([phy_zipcode]) Or IsEmpty([phy_zipcode]),95821,[phy_zipcode]) AS ZIP" & _
                        " FROM physicalzip INNER JOIN customer ON physicalzip.account = customer.account" & _
                        " WHERE customer.account in (SELECT DISTINCTROW [cust_acct] from [temp_Customers])" & _
                        " AND ((customer.phy_address is not null) or (customer.phy_address <> ''))"
                Case 4
                    query = "SELECT Switch([bill_name] Is Not Null And [bill_name]<>'' And" & _
                        " [care_of] Is Not Null And [care_of]<>'',[bill_name] & ' ' & [care_of],[bill_name]" & _
                        " Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]=''),[bill_name]," & _
                        " [care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or [bill_name]='')," & _
                        " [care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT',[name] Is Not Null" & _
                        " Or [name]<>'',[name]) AS CustName, customer.phy_address as address, 'Sacramento' as city, 'CA' as state, " & _
                        " IIf(IsNull([phy_zipcode]) Or IsEmpty([phy_zipcode]),95821,[phy_zipcode]) AS ZIP" & _
                        " FROM physicalzip INNER JOIN customer ON physicalzip.account = customer.account" & _
                        " WHERE customer.account in (SELECT DISTINCTROW [cust_acct] from [temp_Customers])" & _
                        " AND ((customer.phy_address is not null) or (customer.phy_address <> ''))"
                Case Else 'option 2
                    query = "SELECT customer.account, Switch([bill_name] Is Not Null And [bill_name]<>'','OCCUPANT'," & _
                        " [bill_name] Is Not Null And [bill_name]<>'' And (IsNull([care_of]) Or [care_of]='')," & _
                        " 'OCCUPANT',[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or " & _
                        " [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or [name]='UNKNOWN','OCCUPANT'," & _
                        " [name] Is Not Null Or [name]<>'',[name]) AS CustName," & _
                        " customer.phy_address as address, 'Sacramento' as city, 'CA' as state, " & _
                        " IIf(IsNull([phy_zipcode]) Or IsEmpty([phy_zipcode]),95821,[phy_zipcode]) AS ZIP" & _
                        " FROM physicalzip INNER JOIN customer ON physicalzip.account = customer.account" & _
                        " WHERE customer.account in (SELECT DISTINCTROW [cust_acct] from [temp_Customers])" & _
                        " AND ((customer.phy_address is not null) or (customer.phy_address <> ''))"
            End Select
        End Select
End Select

    Debug.Print query

Dim sPath As String
'then output
Select Case Me.fraDestination.value
    Case 1  'Laser printer
        'open a report and pass in the query as it's record source
         Select Case options
            Case 1  'both phone and account number must appear on the label
                Call MsgBox("Printing the phone number has not yet been implemented. Please uncheck this option.", vbExclamation, "Not Implemented")
                Exit Sub
            Case 2  'only the account number must appear on the label
                'This could be to a physical address or mailing address
                Dim fAddr As Boolean
                Dim sReport As String
                If InStr(query, "phy_address") > 0 Then
                    fAddr = True
                    sReport = "prnLabelswAcctPA"
                Else
                    fAddr = False
                    sReport = "prnLabelswAcctMA"
                End If
                
                DoCmd.OpenReport sReport, acViewDesign
                Set rpt = Reports.Item(sReport)
                rpt.RecordSource = query
                DoCmd.Close acReport, rpt.Name, acSaveYes
                DoCmd.OpenReport sReport, acViewPreview
                
            Case 3  'only the phone number must appear on the label
                Call MsgBox("Printing the phone number has not yet been implemented. Please uncheck this option.", vbExclamation, "Not Implemented")
                Exit Sub
            Case 4  'with no account number or phone
                If InStr(query, "phy_address") > 0 Then
                    fAddr = True
                    sReport = "prnLabelsPA"
                Else
                    fAddr = False
                    sReport = "prnLabelsMA"
                End If
                
                DoCmd.OpenReport sReport, acViewDesign
                Set rpt = Reports.Item(sReport)
                rpt.RecordSource = query
                'rpt.Requery
                DoCmd.Close acReport, rpt.Name, acSaveYes
                DoCmd.OpenReport sReport, acViewPreview
        
        End Select
    Case 3  'Dot matrix
        MsgBox "Not yet implemented. Use the laser printer option instead.", vbInformation + vbOKOnly, "Not Implemented"
        Exit Sub
    Case 4 'File
        If options = 3 Or options = 1 Then
            Call MsgBox("Printing the phone number has not yet been implemented. Please uncheck this option.", vbExclamation, "Not Implemented")
            Exit Sub
        End If
        
        Dim F As Object
        Set F = Application.FileDialog(2)   'msoFileDialogSaveAs
        F.InitialFileName = "Accounts.txt"
        F.Show
        'turn off error handling
        On Error Resume Next
        sPath = IIf(F.SelectedItems.Count > 0, F.SelectedItems.Item(1), "")
        'turn error handling back on
        On Error GoTo cmdPrint_Click_Error
        
        If sPath = "" Then
            Call MsgBox("Invalid Path encountered. You either cancled the operation or you selected an invalid path or file name. " & _
                "Please select a valid path or file name before trying to continue with the save.", vbCritical, "Invalid Path/FileName")
            Exit Sub
        End If
        
        Dim fnum As Integer
        fnum = FreeFile
        Dim rst As New ADODB.Recordset
        Dim ix As Integer
        rst.Open query, CurrentProject.Connection
        
        If rst.BOF And rst.EOF Then
            'housten we have a problem, abort
                MsgBox "An error occurred trying to get a list of customers to print for. Process is aborting", vbCritical + vbOKOnly, "Error"
            Exit Sub
        End If
        
        Open sPath For Output As fnum
        Do While Not rst.EOF
            For ix = 0 To rst.Fields.Count - 1
                If (options = 3 Or options = 4) And rst.Fields(ix).Name = "account" Then
                    'don't print anything
                Else
                    If ix = rst.Fields.Count - 1 Then
                        Print #fnum, rst.Fields(ix).value;
                    Else
                        Print #fnum, rst.Fields(ix).value & ", ";
                    End If
                End If
            Next
            Print #fnum, vbCrLf
            rst.MoveNext
        Loop
        Close #fnum
        MsgBox "Done saving the file. at " & vbCrLf & sPath, vbInformation + vbOKOnly, "Completed"
End Select

    'clean up
    If Me.TabCtl0.value = 2 Then
        query = "DELETE FROM temp_Customers"
        CurrentProject.Connection.Execute query
        Form_customer_subform.Requery
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrint_Click of VBA Document Form_PrintMailing")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrint_Click of VBA Document Form_PrintMailing"
    
End Sub

Private Sub Form_Load()
'get who the sender was
'Dim strSender As String
If Not IsNull(Me.OpenArgs) Then
    strSender = Me.OpenArgs
End If
    Me.chkActive = True
    
End Sub

Private Sub TabCtl0_Change()

   On Error GoTo TabCtl0_Change_Error

Select Case Me.TabCtl0.value
    Case 0  'By Account
    
    Case 1  'By Route
    
    Case 2  'By Selected
        
End Select

   On Error GoTo 0
   Exit Sub

TabCtl0_Change_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure TabCtl0_Change of VBA Document Form_PrintMailing")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure TabCtl0_Change of VBA Document Form_PrintMailing"

End Sub

Private Sub CreateReport(qry As String)
Dim rpt As Report
Dim ctl As Control
Dim newCtl As Control
Dim str() As String
Dim ix As Integer
Dim ctr As Integer
ctr = 0

   On Error GoTo CreateReport_Error
   'first we parse the query to figure out how many controls we need
    str = parseQuery(qry)
    ix = UBound(str)
    
    'now we have to create the controls, add these to the report
    DoCmd.OpenReport "prnLabelswAcct", acViewDesign
    
    Set rpt = Reports![prnLabelswAcct]
    'loop through all the controls in the report and delete them
    For Each ctl In rpt.Controls
        If ctr <= ix Then
            If ctl.ControlType = acTextBox Then
                'newCtl = CreateControl(rpt.name, acTextBox, acDetail, ctl.Parent, str(ctr))
                'ctl.Column(0) = str(ctr)
                ctr = ctr + 1
            End If
        Else
            'hide the control
            ctl = Null
            ctr = ctr + 1
        End If
    Next
    
    
    DoCmd.Close acReport, rpt.Name, acSaveYes
    DoCmd.OpenReport "prnLabelswAcct", acViewPreview

   On Error GoTo 0
   Exit Sub

CreateReport_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure CreateReport of VBA Document Form_PrintMailing")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure CreateReport of VBA Document Form_PrintMailing"
End Sub

'Private Sub txtRangeFrom_Change()
' On Error Resume Next
'    txtRangeFrom.text = UCase(txtRangeFrom.text)
'    On Error GoTo 0
'End Sub
'
'Private Sub txtRangeTo_Change()
'On Error Resume Next
'    txtRangeTo.text = UCase(txtRangeTo.text)
'On Error GoTo 0
'End Sub
Private Sub txtRangeFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtRangeTo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
