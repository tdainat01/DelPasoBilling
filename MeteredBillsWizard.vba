Option Compare Database
Option Explicit

Private Sub cmdGo_Click()
'Collect up the data, create a query and then launch the report
Dim query As String
Dim subquery As String
Dim subSubQuery As String
Dim sPath As String
Dim sFileContents As String
Dim sPort As String
Dim min As Long
Dim max As Long
Dim lRecs As Long
Dim dtNow As Date
Dim lLastMonth As Long
Dim lThisYear As Long
Dim sChg As Single
Dim sPmt As Single
Dim sPrevBal As Single
Dim fMissing As Boolean
Dim strFile As String

   On Error GoTo cmdGo_Click_Error

Dim cmd As New ADODB.Command
Dim rst As New ADODB.Recordset
Dim rstSub As New ADODB.Recordset
fMissing = False

dtNow = Now
lLastMonth = Month(dtNow) - 1
lThisYear = Year(dtNow)

If Not IsNull(Me.cmbMinAcctNum) Or Len(Me.cmbMinAcctNum) > 0 Then
    If IsNumeric(Me.cmbMinAcctNum) Then
        min = CLng(cmbMinAcctNum)
    Else
        Exit Sub
    End If
Else
    Exit Sub
End If

If Not IsNull(Me.cmdMaxAcctNum) Or Len(Me.cmdMaxAcctNum) > 0 Then
    If IsNumeric(Me.cmdMaxAcctNum) Then
        max = CLng(cmdMaxAcctNum)
    Else
        Exit Sub
    End If
Else
    Exit Sub
End If



If SetMeteredQuery Then

    CurrentProject.Connection.Execute "DELETE * FROM tmpPrintMeteredBill"
query = "INSERT INTO tmpPrintMeteredBill ( account, [group], mastpar, cycle, mfg_code, start_date, status, meter_number, term_date, out_town, meter_size, property_use, backflow, fire_size," & _
    " unit_measure, current_read, [current_date], rate_code, previous_read, previous_date, gal_cub_used, meter_site, deposit, use_charge, past_due, prev_balance, current_due, " & _
    " special_credit, total_due, special_charge, special_description, phy_address, lien, CustName, bill_name, address, care_of, city, state, zip, comment, Extra_Charges )" & _
    " SELECT customer.account, customer.group, customer.mastpar, customer.cycle, customer.mfg_code, customer.start_date, customer.status, customer.meter_number, customer.term_date," & _
    " customer.out_town, customer.meter_size, customer.property_use, customer.backflow, customer.fire_size, customer.unit_measure, customer.current_read, customer.current_date," & _
    " customer.rate_code, customer.previous_read, customer.previous_date, customer.gal_cub_used, customer.meter_site, customer.deposit, customer.use_charge, customer.past_due," & _
    " customer.prev_balance, customer.current_due, customer.special_credit, customer.total_due, customer.special_charge, customer.special_description, customer.phy_address, customer.lien," & _
    " Switch([bill_name] Is Not Null And [bill_name]<>'' And [care_of] Is Not Null And [care_of]<>'',[bill_name] & ' ' & [care_of],[bill_name] Is Not Null And [bill_name]<>'' And" & _
    " (IsNull([care_of]) Or [care_of]=''),[bill_name],[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or" & _
    " [name]='UNKNOWN','OCCUPANT',[name] Is Not Null Or [name]<>'',[name]) AS CustName, customer.bill_name, IIf(customer.addr1 Is Null Or customer.addr1='',customer.phy_address," & _
    " customer.addr1) AS address, customer.care_of, customer.city, customer.state, customer.zip, customer.comment, Sum(RecurringCharges.charge_amount) AS Extra_Charges" & _
    " FROM RecurringCharges INNER JOIN (customer INNER JOIN RatesAndCharges ON customer.account = RatesAndCharges.account) ON RecurringCharges.charge_id = RatesAndCharges.recurring_charge_id" & _
    " GROUP BY customer.account, customer.group, customer.mastpar, customer.cycle, customer.mfg_code, customer.start_date, customer.status, customer.meter_number, customer.term_date," & _
    " customer.out_town, customer.meter_size, customer.property_use, customer.backflow, customer.fire_size, customer.unit_measure, customer.current_read, customer.current_date, " & _
    " customer.rate_code, customer.previous_read, customer.previous_date, customer.gal_cub_used, customer.meter_site, customer.deposit, customer.use_charge, customer.past_due," & _
    " customer.prev_balance, customer.current_due, customer.special_credit, customer.total_due, customer.special_charge, customer.special_description, customer.phy_address, customer.lien," & _
    " Switch([bill_name] Is Not Null And [bill_name]<>'' And [care_of] Is Not Null And [care_of]<>'',[bill_name] & ' ' & [care_of],[bill_name] Is Not Null And [bill_name]<>'' And" & _
    " (IsNull([care_of]) Or [care_of]=''),[bill_name],[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or" & _
    " [name]='UNKNOWN','OCCUPANT',[name] Is Not Null Or [name]<>'',[name]), customer.bill_name, IIf(customer.addr1 Is Null Or customer.addr1='',customer.phy_address,customer.addr1)," & _
    " customer.care_of, customer.city, customer.state, customer.zip, customer.comment" & _
    " HAVING (((customer.account)>=" & min & " And (customer.account)<=" & max & ") AND ((customer.cycle)=" & Me.txtCycle & ") AND ((customer.status)<>'I') AND ((customer.term_date)" & _
    " Is Null Or (customer.term_date)=#1/1/1900#) AND ((customer.current_read)>0) AND ((customer.previous_read)>0));"

'query = Replace(query, "'", "''")
CurrentProject.Connection.Execute query, lRecs

'Now update the tmpTable and insert the fire-protection values
query = "SELECT * FROM qryFireProtection"
rst.Open query, CurrentProject.Connection
If rst.BOF And rst.EOF Then
    'nothing to do
Else
    Do While Not rst.EOF
        CurrentProject.Connection.Execute "UPDATE tmpPrintMeteredBill set fire_charge = " & _
            rst.Fields("charge_amount").value & " where account = " & rst.Fields("account").value
        rst.MoveNext
    Loop
End If

rst.Close
query = ""

'now calculate smc charges only
query = "SELECT RatesAndCharges.account, Sum(RecurringCharges.charge_amount) AS charge_amount," & _
        " RatesAndCharges.recurring_charge_id FROM RecurringCharges INNER JOIN " & _
        " (tmpPrintMeteredBill INNER JOIN RatesAndCharges ON tmpPrintMeteredBill.account = " & _
        " RatesAndCharges.account) ON RecurringCharges.charge_id = RatesAndCharges.recurring_charge_id" & _
        " WHERE (((RecurringCharges.charge_code) Like 'M%'))" & _
        " GROUP BY RatesAndCharges.account, RatesAndCharges.recurring_charge_id" & _
        " ORDER BY RatesAndCharges.account;"

rst.Open query, CurrentProject.Connection
If rst.BOF And rst.EOF Then
    'nothing to do
Else
    Do While Not rst.EOF
        subquery = "UPDATE tmpPrintMeteredBill set smc_charge = " & _
           IIf(IsNull(rst.Fields("charge_amount").value) Or IsEmpty(rst.Fields("charge_amount").value) _
           Or rst.Fields("charge_amount").value = 0, 0, rst.Fields("charge_amount").value) & " where account = " & rst.Fields("account").value
        CurrentProject.Connection.Execute subquery, lRecs
        If lRecs <= 0 Then
            'an error occurred
            Err.Raise vbObjectError + 8901, "Form_MeteredBillsWizard:: cmdGo_Click", "No update occurred for the tmpPrintMeteredBill table"
        Else
            'Debug.Print subquery & " with records affected = " & CStr(lRecs)
        End If
        rst.MoveNext
    Loop
End If

rst.Close

'Now update the previous balance
'query = "SELECT account, prev_balance from tmpPrintMeteredBill"
'rst.Open query, CurrentProject.Connection
'If rst.BOF And rst.EOF Then
'    'nothing to do
'Else
'    Do While Not rst.EOF
'        subquery = "SELECT Money.account_number, Money.amount, Money.code, Money.trans_date" & _
'                " FROM [Money] WHERE (((Money.account_number)=" & rst.Fields("account").value & ")" & _
'                " AND ((Month([Money].[trans_date]))=" & lLastMonth & _
'                ") AND ((Year([Money].[trans_date]))=" & lThisYear & "));"
'
'        rstSub.Open subquery, CurrentProject.Connection
'        If rstSub.BOF And rstSub.EOF Then
'            sChg = 0
'            sPmt = 0
'        Else
'            Do While Not rstSub.EOF
'                If rstSub.Fields("code").value = "CHG" Then
'                    sChg = sChg + CSng(IIf(IsNull(rstSub.Fields("amount").value), 0, rstSub.Fields("amount").value))
'                ElseIf rstSub.Fields("code").value = "PMT" Then
'                    sPmt = sPmt + CSng(IIf(IsNull(rstSub.Fields("amount").value), 0, rstSub.Fields("amount").value))
'                Else
'                    'do nothing
'                End If
'                rstSub.MoveNext
'            Loop
'        End If
'
'        sPrevBal = sChg - Abs(sPmt)
'
'        subSubQuery = "UPDATE tmpprintmeteredbill set prev_balance = " & sPrevBal & " WHERE account = " & rst.Fields("account").value
'        CurrentProject.Connection.Execute subSubQuery, lRecs
'        sChg = 0
'        sPmt = 0
'        rstSub.Close
'
'        rst.MoveNext
'    Loop
'End If
'
'rst.Close

arguments = "SELECT tmpPrintMeteredBill.* FROM tmpPrintMeteredBill WHERE (((tmpPrintMeteredBill.total_due)>0));" & _
    "|" & Me.txtStartDate & "|" & Me.txtEndDate & "|" & Me.txtYear

'Call routine to output values to text file
strFile = CreateFileName

query = "SELECT SettingsName, SettingsValue from Settings WHERE SettingsName = 'Path'"
rst.Open query, CurrentProject.Connection
If rst.BOF And rst.EOF Then
    'Path variable not set. Get this from the user.
    sPath = InputBox("No output path has been set for the report. Please enter one now", "Set Path", "C:\Temp")
    If Len(sPath) < 1 Then
        Call MsgBox("No path has been specified. Aborting Process", vbCritical + vbOKOnly, "Aborting ...")
        Exit Sub
    End If
    'validate folder
    If Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If
    
    If Dir(sPath) <> "" Then
        'Path exists.
        'Insert the value into the Settimgs Table for future use.
        CurrentProject.Connection.Execute "INSERT INTO Settings(SettingsName,SettingsValue) VALUES('Path'," & sPath & ")"
        Call OutPutTextBill(arguments, sPath & "MeterBill" & strFile)
    Else
        'Invalid path. Abort process
        Call MsgBox("An invalid path has been specified. Aborting Process", vbCritical + vbOKOnly, "Aborting ...")
        Exit Sub
    End If
Else
    'take the first value
    rst.MoveFirst
    sPath = IIf(IsNull(rst.Fields("SettingsValue").value), "", rst.Fields("SettingsValue").value)
    If Len(sPath) < 1 Then
        fMissing = True
        sPath = InputBox("No output path has been set for the report. Please enter one now", "Set Path", "C:\Temp\")
    End If
    
    If Dir(sPath, vbDirectory) <> vbNullString Then
        'Path exists.
        If Right(sPath, 1) <> "\" Then
            sPath = sPath & "\"
        End If
        'Insert the value into the Settings Table for future use.
        If fMissing Then
            CurrentProject.Connection.Execute "INSERT INTO Settings(SettingsName,SettingsValue) VALUES('Path'," & sPath & ")"
        End If
        Call OutPutTextBill(arguments, sPath & "MeterBill" & strFile)
    Else
        'Invalid path. Abort process
        Call MsgBox("An invalid path has been specified. Aborting Process", vbCritical + vbOKOnly, "Aborting ...")
        Exit Sub
    End If
End If
rst.Close
'Call MsgBox("Bill Output has been completed to " & sPath & "MeterBill" & strFile, vbOKOnly, "Output Dome")
'reset
fMissing = False

'Select Case MsgBox("Send File to Printer?", vbYesNo Or vbQuestion Or vbSystemModal Or vbDefaultButton1, "Send to Printer")

    'Case vbYes
        sFileContents = ReadAsciiFile(sPath & "MeterBill" & strFile)
        If Len(sFileContents) < 1 Then
            'we got nothing???
            Exit Sub
        End If
        rst.Open "SELECT * FROM Settings WHERE SettingsName = 'PrinterPort'", CurrentProject.Connection
        If rst.BOF And rst.EOF Then
            'we have no printer port set
            Call MsgBox("You will be given a chance to enter a printer port. " & _
                "Enter a value such as 'LPT1:' or 'USB1:'. Note the Colon as the last character." & _
                " Entering an incorrect value will result in an error.", vbOKOnly, "Message")
            fMissing = True
            sPort = InputBox("Enter Port value", "Port?", "LPT1:")
        Else
            rst.MoveFirst
            sPort = IIf(IsNull(rst.Fields("SettingsValue").value), "", rst.Fields("SettingsValue").value)
            If Len(sPort) < 1 Then
                Call MsgBox("You will be given a chance to enter a printer port. " & _
                    "Enter a value such as 'LPT1:' or 'USB1:'. Note the Colon as the last character." & _
                    " Entering an incorrect value will result in an error.", vbOKOnly, "Message")
                sPort = InputBox("Enter Port value", "Port?", "LPT1:")
                CurrentProject.Connection.Execute "UPDATE Settings SET SettingsValue = '" & sPort & "'" & _
                                                  " WHERE (((Settings.SettingsName)='Port'));"
            End If
            
            If fMissing Then
                CurrentProject.Connection.Execute "INSERT INTO Settings(SettingsName,SettingsValue) VALUES('Port'," & sPort & ")"
            End If
        
            Call PrintFile(sPath & "MeterBill" & strFile, sPort)
        End If
'    Case vbNo
'        'do nothing
'        Exit Sub
'End Select
'DoCmd.OpenReport "MeteredBills", acViewPreview, "", "", acWindowNormal

Else
'Flat Rate Customer
    CurrentProject.Connection.Execute "DELETE * FROM tmpPrintFlatBill"
query = "INSERT INTO tmpPrintFlatBill ( account, [group], mastpar, cycle, mfg_code, start_date, status, meter_number, term_date, out_town, meter_size, property_use, backflow, fire_size," & _
    " unit_measure, current_read, [current_date], rate_code, previous_read, previous_date, gal_cub_used, meter_site, deposit, use_charge, past_due, prev_balance, current_due, special_credit," & _
    " total_due, special_charge, special_description, phy_address, lien, CustName, bill_name, address, care_of, city, state, zip, comment, Extra_Charges )" & _
    " SELECT customer.account, customer.group, customer.mastpar, customer.cycle, customer.mfg_code, customer.start_date, customer.status, customer.meter_number, customer.term_date, " & _
    " customer.out_town, customer.meter_size, customer.property_use, customer.backflow, customer.fire_size, customer.unit_measure, customer.current_read, customer.current_date, " & _
    " customer.rate_code, customer.previous_read, customer.previous_date, customer.gal_cub_used, customer.meter_site, customer.deposit, customer.use_charge, customer.past_due, " & _
    " customer.prev_balance, customer.current_due, customer.special_credit, customer.total_due, customer.special_charge, customer.special_description, customer.phy_address, customer.lien," & _
    " Switch([bill_name] Is Not Null And [bill_name]<>'' And [care_of] Is Not Null And [care_of]<>'',[bill_name] & ' ' & [care_of],[bill_name] Is Not Null And [bill_name]<>'' And " & _
    " (IsNull([care_of]) Or [care_of]=''),[bill_name],[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or" & _
    " [name]='UNKNOWN','OCCUPANT',[name] Is Not Null Or [name]<>'',[name]) AS CustName, customer.bill_name, IIf(customer.addr1 Is Null Or customer.addr1='',customer.phy_address,customer.addr1)" & _
    " AS address, customer.care_of, customer.city, customer.state, customer.zip, customer.comment, Sum(RecurringCharges.charge_amount) AS Extra_Charges" & _
    " FROM RecurringCharges INNER JOIN (customer INNER JOIN RatesAndCharges ON customer.account = RatesAndCharges.account) ON RecurringCharges.charge_id = RatesAndCharges.recurring_charge_id" & _
    " GROUP BY customer.account, customer.group, customer.mastpar, customer.cycle, customer.mfg_code, customer.start_date, customer.status, customer.meter_number, customer.term_date," & _
    " customer.out_town, customer.meter_size, customer.property_use, customer.backflow, customer.fire_size, customer.unit_measure, customer.current_read, customer.current_date, " & _
    " customer.rate_code, customer.previous_read, customer.previous_date, customer.gal_cub_used, customer.meter_site, customer.deposit, customer.use_charge, customer.past_due, " & _
    " customer.prev_balance, customer.current_due, customer.special_credit, customer.total_due, customer.special_charge, customer.special_description, customer.phy_address, customer.lien," & _
    " customer.bill_name, IIf(customer.addr1 Is Null Or customer.addr1='',customer.phy_address,customer.addr1), customer.care_of, customer.city, customer.state, customer.zip, customer.comment," & _
    " Switch([bill_name] Is Not Null And [bill_name]<>'' And [care_of] Is Not Null And [care_of]<>'',[bill_name] & ' ' & [care_of],[bill_name] Is Not Null And [bill_name]<>'' And " & _
    " (IsNull([care_of]) Or [care_of]=''),[bill_name],[care_of] Is Not Null And [care_of]<>'' And (IsNull([bill_name]) Or [bill_name]=''),[care_of],IsNull([name]) Or [name]='' Or" & _
    " [name]='UNKNOWN','OCCUPANT',[name] Is Not Null Or [name]<>'',[name]) HAVING (((customer.account)>=" & min & " And (customer.account)<=" & max & ") AND ((customer.cycle)=" & Me.txtCycle & _
    ") AND ((customer.status)<>'I') AND ((customer.term_date) Is Null Or (customer.term_date)=#1/1/1900#) AND ((customer.current_read)=0 Or (customer.current_read) Is Null) AND " & _
    " ((customer.previous_read)=0 Or (customer.previous_read) Is Null));"

    'query = Replace(query, "'", "''")
    CurrentProject.Connection.Execute query

'Now update the tmpTable and insert the fire-protection values
query = "SELECT * FROM qryFireProtection"
rst.Open query, CurrentProject.Connection
If rst.BOF And rst.EOF Then
    'nothing to do
Else
    Do While Not rst.EOF
        CurrentProject.Connection.Execute "UPDATE tmpPrintFlatBill set fire_charge = " & _
            rst.Fields("charge_amount").value & " where account = " & rst.Fields("account").value
        rst.MoveNext
    Loop
End If

rst.Close

'now calculate smc charges only
query = "SELECT RatesAndCharges.account, Sum(RecurringCharges.charge_amount) AS charge_amount," & _
        " RatesAndCharges.recurring_charge_id FROM RecurringCharges INNER JOIN " & _
        " (tmpPrintFlatBill INNER JOIN RatesAndCharges ON tmpPrintFlatBill.account = " & _
        " RatesAndCharges.account) ON RecurringCharges.charge_id = RatesAndCharges.recurring_charge_id" & _
        " WHERE (((RecurringCharges.charge_code) Like 'M%'))" & _
        " GROUP BY RatesAndCharges.account, RatesAndCharges.recurring_charge_id" & _
        " ORDER BY RatesAndCharges.account;"

rst.Open query, CurrentProject.Connection
If rst.BOF And rst.EOF Then
    'nothing to do
Else
    Do While Not rst.EOF
        subquery = "UPDATE tmpPrintFlatBill set smc_charge = " & _
           IIf(IsNull(rst.Fields("charge_amount").value) Or IsEmpty(rst.Fields("charge_amount").value) _
           Or rst.Fields("charge_amount").value = 0, 0, rst.Fields("charge_amount").value) & " where account = " & rst.Fields("account").value
        'Debug.Print subquery
        CurrentProject.Connection.Execute subquery
        rst.MoveNext
    Loop
End If

rst.Close

'Now update the previous balance
'query = "SELECT account, prev_balance from tmpFlatRateBill"
'rst.Open query, CurrentProject.Connection
'If rst.BOF And rst.EOF Then
'    'nothing to do
'Else
'    Do While Not rst.EOF
'        subquery = "SELECT Money.account_number, Money.amount, Money.code, Money.trans_date" & _
'                " FROM [Money] WHERE (((Money.account_number)=" & rst.Fields("account").value & ")" & _
'                " AND ((Month([Money].[trans_date]))=" & lLastMonth & _
'                ") AND ((Year([Money].[trans_date]))=" & lThisYear & "));"
'
'        rstSub.Open subquery, CurrentProject.Connection
'        If rstSub.BOF And rstSub.EOF Then
'            sChg = 0
'            sPmt = 0
'        Else
'            Do While Not rstSub.EOF
'                If rstSub.Fields("code").value = "CHG" Then
'                    sChg = sChg + CSng(IIf(IsNull(rstSub.Fields("amount").value), 0, rstSub.Fields("amount").value))
'                ElseIf rstSub.Fields("code").value = "PMT" Then
'                    sPmt = sPmt + CSng(IIf(IsNull(rstSub.Fields("amount").value), 0, rstSub.Fields("amount").value))
'                Else
'                    'do nothing
'                End If
'                rstSub.MoveNext
'            Loop
'        End If
'
'        sPrevBal = sChg - Abs(sPmt)
'
'        subSubQuery = "UPDATE tmpFlatRateBill set prev_balance = " & sPrevBal & " WHERE account = " & rst.Fields("account").value
'        CurrentProject.Connection.Execute subSubQuery, lRecs
'        rstSub.Close
'
'        rst.MoveNext
'    Loop
'End If
'
'rst.Close

'arguments are global
arguments = "SELECT tmpPrintFlatBill.* FROM tmpPrintFlatBill WHERE (((tmpPrintFlatBill.total_due)>0));" & _
    "|" & Me.txtStartDate & "|" & Me.txtEndDate & "|" & Me.txtYear
'Call routine to output values to text file
strFile = CreateFileName

query = "SELECT SettingsName, SettingsValue from Settings WHERE SettingsName = 'Path'"
rst.Open query, CurrentProject.Connection
If rst.BOF And rst.EOF Then
    'Path variable not set. Get this from the user.
    sPath = InputBox("No output path has been set for the report. Please enter one now", "Set Path", "C:\Temp")
    If Len(sPath) < 1 Then
        Call MsgBox("No path has been specified. Aborting Process", vbCritical + vbOKOnly, "Aborting ...")
        Exit Sub
    End If
    'validate folder
        If Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If

    If Dir(sPath, vbDirectory) <> vbNullString Then
        'Path exists.
        'Insert the value into the Settimgs Table for future use.
        CurrentProject.Connection.Execute "INSERT INTO Settings(SettingsName,SettingsValue) VALUES('Path'," & sPath & ")"
        Call OutPutTextBill(arguments, sPath & "FlatBill" & strFile)
    Else
        'Invalid path. Abort process
        Call MsgBox("An invalid path has been specified. Aborting Process", vbCritical + vbOKOnly, "Aborting ...")
        Exit Sub
    End If
Else
    'take the first value
    rst.MoveFirst
    sPath = IIf(IsNull(rst.Fields("SettingsValue").value), "", rst.Fields("SettingsValue").value)
    If Len(sPath) < 1 Then
        fMissing = True
        sPath = InputBox("No output path has been set for the report. Please enter one now", "Set Path", "C:\Temp\")
    End If
    
    If Right(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If

    If Dir(sPath, vbDirectory) <> vbNullString Then
        'Path exists.
        'Insert the value into the Settimgs Table for future use.
        If fMissing Then
            CurrentProject.Connection.Execute "INSERT INTO Settings(SettingsName,SettingsValue) VALUES('Path'," & sPath & ")"
        End If
        Call OutPutTextBill(arguments, sPath & "FlatBill" & strFile)
    Else
        'Invalid path. Abort process
        Call MsgBox("An invalid path has been specified. Aborting Process", vbCritical + vbOKOnly, "Aborting ...")
        Exit Sub
    End If
    'Exit Sub
End If
rst.Close
'Call MsgBox("Bill Output has been completed to " & sPath & "FlatBill" & strFile, vbOKOnly, "Output Dome")

'reset
fMissing = False

'Select Case MsgBox("Send File to Printer?", vbYesNo Or vbQuestion Or vbSystemModal Or vbDefaultButton1, "Send to Printer")

    'Case vbYes
        sFileContents = ReadAsciiFile(sPath & "FlatBill" & strFile)
        If Len(sFileContents) < 1 Then
            'we got nothing???
            Exit Sub
        End If
        rst.Open "SELECT * FROM Settings WHERE SettingsName = 'PrinterPort'", CurrentProject.Connection
        If rst.BOF And rst.EOF Then
            'we have no printer port set
            Call MsgBox("You will be given a chance to enter a printer port. " & _
                "Enter a value such as 'LPT1:' or 'USB1:'. Note the Colon as the last character." & _
                " Entering an incorrect value will result in an error.", vbOKOnly, "Message")
            fMissing = True
            sPort = InputBox("Enter Port value", "Port?", "LPT1:")
        Else
            rst.MoveFirst
            sPort = IIf(IsNull(rst.Fields("SettingsValue").value), "", rst.Fields("SettingsValue").value)
            If Len(sPort) < 1 Then
                Call MsgBox("You will be given a chance to enter a printer port. " & _
                    "Enter a value such as 'LPT1:' or 'USB1:'. Note the Colon as the last character." & _
                    " Entering an incorrect value will result in an error.", vbOKOnly, "Message")
                sPort = InputBox("Enter Port value", "Port?", "LPT1:")
                CurrentProject.Connection.Execute "UPDATE Settings SET SettingsValue = '" & sPort & "'" & _
                                                  " WHERE (((Settings.SettingsName)='Port'));"
            End If
            
            If fMissing Then
                CurrentProject.Connection.Execute "INSERT INTO Settings(SettingsName,SettingsValue) VALUES('Port'," & sPort & ")"
            End If
        
            Call PrintFile(sPath & "FlatBill" & strFile, sPort)
        End If
    'Case vbNo
        'do nothing
        'Exit Sub
'End Select

'DoCmd.OpenReport "FlatRateBills", acViewPreview, "", "", acWindowNormal

End If
                
   On Error GoTo 0
   Exit Sub

cmdGo_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdGo_Click of VBA Document Form_MeteredBillsWizard")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdGo_Click of VBA Document Form_MeteredBillsWizard"

End Sub

Private Sub Form_Load()

Dim rst As New ADODB.Recordset
Dim min As String
Dim max As String
Dim vOpenArgs As Variant

   On Error GoTo Form_Load_Error

'Find out who called this. Set Form caption as applicable
vOpenArgs = Me.OpenArgs
If vOpenArgs = "FlatRate" Then
    Me.Caption = "Flat Rate Billing Wizard Step 1"
ElseIf vOpenArgs = "MeterRate" Then
    Me.Caption = "Metered Rate Billing Wizard Step 1"
Else
    Me.Caption = "Billing Wizard Step 1"
End If

rst.Open "SELECT Min(customer.account) AS MinOfaccount FROM customer;", CurrentProject.Connection
min = rst.Fields(0).value
rst.Close

rst.Open "SELECT Max(customer.account) AS MinOfaccount FROM customer;", CurrentProject.Connection
max = rst.Fields(0).value
rst.Close

Me.cmbMinAcctNum.value = min
Me.cmdMaxAcctNum.value = max
Me.txtYear = Format(Now(), "yyyy")

   On Error GoTo 0
   Exit Sub

Form_Load_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Load of VBA Document Form_MeteredBillsWizard")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Load of VBA Document Form_MeteredBillsWizard"

End Sub

