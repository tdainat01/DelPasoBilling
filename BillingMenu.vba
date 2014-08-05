Option Compare Database
Option Explicit

Private Sub cmdChangeDate_Click()
   On Error GoTo cmdChangeDate_Click_Error

    DoCmd.OpenForm "frmChangeDate"

   On Error GoTo 0
   Exit Sub

cmdChangeDate_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdChangeDate_Click of VBA Document Form_BillingMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdChangeDate_Click of VBA Document Form_BillingMenu"
    
End Sub

Private Sub cmdChargeFlat_Click()
Dim Cycle As Integer
Dim inputResult As Variant
Dim strQuery As String
Dim strAccounts As String
Dim strFireSize As String
Dim strRollOverAccounts As String
Dim rst As New ADODB.Recordset
Dim rstRecur As New ADODB.Recordset
Dim rstServ As New ADODB.Recordset
Dim rstDate As New ADODB.Recordset
Dim ctr As Integer
Dim account As Long
Dim charged As Single
Dim prevBal As Single
Dim TotalDue As Single
Dim used As Single
Dim current As Single
Dim Last As Single
Dim strTransaction As String
Dim process As Long
Dim lRecs As Long

   On Error GoTo cmdChargeFlat_Click_Error

inputResult = InputBox("Cycle #", "Cycle", 1)

If inputResult = False Or inputResult = "" Then
    Exit Sub
Else
    Cycle = CInt(inputResult)
End If

If Cycle <= 0 Then
    Exit Sub
End If

Me.recOuter.Visible = True
Me.recInner.Visible = True

'@TODO - Remove this when launching for testing or production.
'MsgBox "Flat Rate Charges would now be calculated.", _
'    vbOKOnly + vbInformation, "This is a placeholder message"
''Exit Sub


Me.lstRollOver.Visible = False
Me.lblRollOver.Visible = False

'Check to see if there are unmatched records in the customer table on the rate codes
'This query gets both flat rate and metered accounts
strQuery = "SELECT customer.account FROM customer LEFT JOIN RecurringCharges ON customer.[rate_code] = RecurringCharges.[charge_code]" & _
            " WHERE (((RecurringCharges.charge_code) Is Null) AND ((customer.status)='A'));"

rst.Open strQuery, CurrentProject.Connection
Do While Not rst.EOF
    ctr = ctr + 1
    strAccounts = strAccounts & rst.Fields(0).value & vbCrLf
    rst.MoveNext
Loop

rst.Close

If ctr > 0 Then
    MsgBox "There are unmatched records in the customer table on the following records:" _
           & strAccounts, vbCritical + vbOKOnly, _
        "Unmatched Rate Code"
    Me.recOuter.Visible = False
    Me.recInner.Visible = False
    Exit Sub
End If

'Check to see if there are any unmatched records on line size and meter size
'strQuery = "SELECT customer.account FROM customer LEFT JOIN Rates ON customer.fire_size = Rates.Line_Size" & _
'            " WHERE (((Rates.Line_Size) Is Null)AND ((Rates.Meter_Resident) Like 'F*'));"
            
strQuery = " SELECT customer.account FROM customer LEFT JOIN ServiceConnections ON customer.[meter_size] = ServiceConnections.[LineSize]" & _
            " WHERE ServiceConnections.LineSize Is Null AND ServiceConnections.MR_Code like 'F%';"

rst.Open strQuery, CurrentProject.Connection
ctr = 0

Do While Not rst.EOF
    ctr = ctr + 1
    strFireSize = strFireSize & rst.Fields(0).value & vbCrLf
    rst.MoveNext
Loop

rst.Close

If ctr > 0 Then
    MsgBox "There are unmatched records for meter_size and line_size in the customer table on the following records:" _
           & strAccounts, vbCritical + vbOKOnly, _
        "Unmatched Line Size"
    Exit Sub
End If

'Filter out inactive accounts and Metered Accounts and check to see if last read is greater than the current read
'strQuery = "SELECT customer.*, Rates.Code, Rates.Oper1, Rates.Amount FROM customer " & _
'            " INNER JOIN Rates ON customer.rate_code = Rates.Code " & _
'            " WHERE ((customer.status)<>'I') AND " & _
'            " (((customer.term_date) Is Null OR (customer.term_date) = #1/1/1900#) AND ((Rates.Meter_Resident)<>'M') AND " & _
'            " ((customer.cycle))=" & Cycle & ");"

'strQuery = "SELECT customer.*, CustomerServiceConnection.service_id" & _
'           " FROM customer INNER JOIN CustomerServiceConnection ON customer.account = CustomerServiceConnection.account" & _
'           " WHERE (((customer.cycle)=" & Cycle & ") AND ((customer.status)<>'I') AND ((customer.term_date) Is Null Or " & _
'           " (customer.term_date)=#1/1/1900#) AND ((CustomerServiceConnection.service_id) In (9,10,11,12,15)));"

strQuery = "SELECT customer.* FROM customer" & _
            " WHERE (((customer.cycle)=" & Cycle & ") AND ((customer.status)<>'I') AND ((customer.term_date) Is Null Or " & _
            " (customer.term_date)=#1/1/1900#) AND ((customer.current_read = 0 " & _
            " or customer.current_read is null)) AND ((customer.previous_read) = 0 or customer.previous_read is null));"


rst.Open strQuery, CurrentProject.Connection, adOpenDynamic, adLockPessimistic
'Set rst = CurrentProject.Connection.Execute(strQuery, lRecs)

Do While Not rst.EOF
    lRecs = lRecs + 1
    rst.MoveNext
Loop

rst.MoveFirst

ctr = 0
process = (ProgressBarWidth - ProgressIndicator) / lRecs

'Calculate the used values
Do While Not rst.EOF    'outer loop
    'reset the charged variable
    charged = 0 'rst.Fields("total_due").value
    account = IIf(IsNull(rst.Fields("account").value), 0, rst.Fields("account").value)
    If account = 0 Then
        Err.Raise vbObjectError + 5001, "cmdChargeMeter_Click of Form_BillingMenu", "account cannot be zero"
    End If
    
    'Calculate the service line charge
    query = "SELECT CustomerServiceConnection.account, ServiceConnections.Amount" & _
            " FROM CustomerServiceConnection INNER JOIN ServiceConnections ON CustomerServiceConnection.service_id = " & _
            " ServiceConnections.service_id WHERE (((CustomerServiceConnection.account)=" & account & "));"

    rstServ.Open query, CurrentProject.Connection ', adOpenDynamic, adLockOptimistic
    If rstServ.BOF And rstServ.EOF Then
        'no service charge
    Else
        'take only the first service charge as techinically there can only be one
        'add the usage charge to the service charge
        charged = charged + rstServ.Fields(1).value
    End If
    
    rstServ.Close
        
    'next calculate any recurring charges such as fire service, maint charges etc.
    query = "SELECT RatesAndCharges.account, RecurringCharges.charge_amount, RecurringCharges.charge_description" & _
            " FROM RecurringCharges INNER JOIN RatesAndCharges ON RecurringCharges.charge_id = RatesAndCharges.recurring_charge_id" & _
            " WHERE (((RatesAndCharges.account)=" & account & "));"

    rstRecur.Open query, CurrentProject.Connection
    strTransaction = ""
    
    If rstRecur.BOF And rstRecur.EOF Then
        'nothing to do - no recurring service charge
    Else
        'could be many recurring charges - take them all
        Do While Not rstRecur.EOF
            charged = charged + rstRecur.Fields(1).value
            strTransaction = strTransaction & " " & rstRecur.Fields(2).value
            rstRecur.MoveNext
        Loop
    End If
    
    rstRecur.Close
    
    'Now move the values into their respective buckets
    If rst.Fields("total_due").value <> 0 Then
        rst.Fields("past_due").value = rst.Fields("total_due").value
    End If
    
    current = charged
    rst.Fields("use_charge").value = charged
    
    current = current + rst.Fields("special_charge").value
    current = current - rst.Fields("special_credit").value
      
    
    If rst.Fields("prev_balance").value < 0 Then
        rst.Fields("prev_balance").value = Abs(rst.Fields("prev_balance").value)    'set to absolute value same as multiplying by -1
    End If
    
    rst.Fields("current_due").value = current
    
    If rst.Fields("prev_balance").value <> 0 Then
        prevBal = rst.Fields("prev_balance").value
        
        If prevBal - current > 0 Then
            prevBal = prevBal - current
            rst.Fields("prev_balance").value = prevBal
            rst.Fields("current_due").value = 0
            current = 0
        ElseIf prevBal - current < 0 Then
            prevBal = prevBal - current
            rst.Fields("prev_balance").value = 0
            prevBal = Abs(prevBal)
            rst.Fields("current_due").value = prevBal 'this will overwrite the current due with the prev bal
            current = prevBal
        Else
            rst.Fields("prev_balance").value = 0#
            rst.Fields("current_due").value = 0#
            current = 0#
        End If
    End If
        
    current = current + rst.Fields("past_due").value
    rst.Fields("total_due").value = current
    
    'append everything to the transaction file (money)
    Dim strMoney As String
    Dim recsAffected As Long
    Dim tmpQuery As String
    Dim dt As Date
    
    'look up the date and use this for posting the transactions
    tmpQuery = "SELECT TOP 1 Max(System.auto_id) AS MaxOfauto_id, System.system_date, System.use_system_date" & _
               " FROM System GROUP BY System.system_date, System.use_system_date ORDER BY Max(System.auto_id) DESC;"
    Dim tmpBool As Boolean
    
    rstDate.Open tmpQuery, CurrentProject.Connection
    
    '###
    If rstDate.BOF And rstDate.EOF Then
        dt = Now
    Else
        tmpBool = IIf(IsNull(rstDate.Fields(2).value), False, CBool(rstDate.Fields(2).value))
        
        If tmpBool Then
            dt = CDate(rstDate.Fields(1).value)
        Else
            dt = Now
        End If
    End If
    
    rstDate.Close
    
    strTransaction = Replace(strTransaction, "'", "''")
    
    strMoney = "INSERT INTO [Money] (m_month, m_day, m_year, account_number, [amount], [transaction], [code], [posted], behind_me, trans_date) " & _
                " VALUES ('" & Format(dt, "mm") & "','" & Format(dt, "dd") & "','" & Format(dt, "yyyy") & "','" & rst.Fields("account").value & "'," & _
                charged & ",'" & strTransaction & "'," & "'CHG'" & "," & "'Y'," & "0,#" & dt & "#)"
    
    CurrentProject.Connection.Execute strMoney, recsAffected
    
    If recsAffected < 1 Then
        'an error occurred do something to handle this
        Err.Raise vbObjectError + 5010, "cmdChargeFlat_Click of Form BillingMenu", "No records were appended to the Money Table for query: " & strMoney
    End If
    
    'Now update the customer table
    'we have to move the values in the buckets around.
    rst.Update
    rst.MoveNext
    ctr = ctr + 1
    DoEvents
    Me.recInner.Width = (Me.recInner.Width + process)
    'Debug.Print "Width = " & Me.recInner.Width
    Me.Repaint
Loop

rst.Close

If Me.lstRollOver.ListCount > 0 Then
    Me.lstRollOver.Visible = True
    Me.lblRollOver.Visible = True
Else
    Call MsgBox("Flat Rate Account charging has completed. " & ctr & " accounts were charged.", vbInformation, "Complete")
End If

    Me.recOuter.Visible = False
    Me.recInner.Width = ProgressIndicator
    Me.recInner.Visible = False
   On Error GoTo 0
   Exit Sub

cmdChargeFlat_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdChargeFlat_Click of VBA Document Form_BillingMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdChargeFlat_Click of VBA Document Form_BillingMenu"

End Sub

Private Sub cmdChargeMetered_Click()

Dim Cycle As Integer
Dim inputResult As Variant
Dim rstRecur As New ADODB.Recordset
Dim rstServ As New ADODB.Recordset
Dim rst As New ADODB.Recordset
Dim rstUC As New ADODB.Recordset
Dim strQuery As String
Dim account As Long
Dim strAccounts As String
Dim strFireSize As String
Dim query As String
Dim strRollOverAccounts As String
Dim ctr As Integer
Dim Service As Single
Dim charged As Currency
Dim prevBal As Currency
Dim current As Currency
Dim used As Single
Dim Last As Single
Dim strTransaction As String
Dim lRecs As Long
Dim process As Long
Dim iCalc As Long

   On Error GoTo cmdChargeMetered_Click_Error

'Get the cycle
inputResult = InputBox("Cycle #", "Cycle", 1)

If inputResult = False Or inputResult = "" Then
    Exit Sub
Else
    Cycle = CInt(inputResult)
End If

If Cycle <= 0 Then
    Exit Sub
End If

Me.recOuter.Visible = True
Me.recInner.Visible = True

'@TODO - Remove this when launcing for testing or production.
'MsgBox "Metered Charges would now be calculated.", _
'    vbOKOnly + vbInformation, "This is a placeholder message"
'Exit Sub

Me.lstRollOver.Visible = False
Me.lblRollOver.Visible = False

'Check to see if there are unmatched records in the customer table on the rate codes
'This query gets both flat rate and metered accounts
strQuery = "SELECT customer.account FROM customer LEFT JOIN RecurringCharges ON customer.[rate_code] = RecurringCharges.[charge_code]" & _
            " WHERE (((RecurringCharges.charge_code) Is Null) AND ((customer.status)='A'));"

rst.Open strQuery, CurrentProject.Connection

Do While Not rst.EOF
    ctr = ctr + 1
    strAccounts = strAccounts & rst.Fields(0).value & vbCrLf
    rst.MoveNext
Loop

rst.Close

If ctr > 0 Then
    MsgBox "There are unmatched records in the customer table on the following records:" _
           & strAccounts, vbCritical + vbOKOnly, _
        "Unmatched Rate Code"
        Call WriteToFile("C:\DelPaso\unmatched.txt", strAccounts)
    Exit Sub
End If

'Now select all the metered customer accounts Filter out inactive accounts and check to see if last read is greater than the current read
'TODO Verify that using previous and current reads is a correct way to determine a metered customer.
strQuery = "SELECT customer.* FROM customer" & _
            " WHERE (((customer.cycle)=" & Cycle & ") AND ((customer.status)<>'I') AND ((customer.term_date) Is Null Or " & _
            " (customer.term_date)=#1/1/1900#) AND ((customer.current_read)>0) AND ((customer.previous_read)>0));"

rst.Open strQuery, CurrentProject.Connection, adOpenDynamic, adLockPessimistic
'Set rst = CurrentProject.Connection.Execute(strQuery, lRecs)

If Not rst.BOF And Not rst.EOF Then
    Do While Not rst.EOF
        lRecs = lRecs + 1
        rst.MoveNext
    Loop
    rst.MoveFirst
Else
    'there is no work do to
    Call MsgBox("No records were found for the criteria you selected. " & _
                " Please adjust the cycle or customer records (to fit the cycle) and try again.", _
                vbExclamation, "No Records")
    Exit Sub
End If

ctr = 0
process = (ProgressBarWidth - ProgressIndicator) / lRecs

'TODO check for meter roll over

'Calculate the used values
Do While Not rst.EOF    'outer loop
    charged = 0 'reset the charged value
    account = IIf(IsNull(rst.Fields("account").value), 0, rst.Fields("account").value)
    If account = 0 Then
        Err.Raise vbObjectError + 5001, "cmdChargeMeter_Click of Form_BillingMenu", "account cannot be zero"
    End If
    If rst.Fields("current_read").value < rst.Fields("previous_read").value Then
    'roll over
    Me.lstRollOver.AddItem (rst.Fields("account").value)
    'Need an update query to handle this
    'strRollOverAccounts = strRollOverAccounts & "'" & rst.Fields(0).Value & "'"
    Else
        'Calculate gallons uses
        used = rst.Fields("current_read").value - rst.Fields("previous_read").value
        
        'If the fields unit of measure is a G then do a converstion
        If Left(rst.Fields("unit_measure").value, 1) = "G" Then
            rst.Fields("gal_cub_used").value = used * GALSTOCUFEET
        Else
            rst.Fields("gal_cub_used").value = used
        End If
        'usage charge is calculated per 100 cu feet
        
        strTransaction = FormatNumber(rst.Fields("meter_size").value, 2) & Space(7) & CStr(Format(used, "0000000000")) 'TD 5/9/14 - Added
        
        lRecs = 0
        rstUC.Open "SELECT [Amount] from MeterRates", CurrentProject.Connection
        If rstUC.BOF And rstUC.EOF Then
            'throw an error
            Exit Sub
        End If
        'iCalc = CLng(rst.Fields("gal_cub_used").value)
        charged = Round((rst.Fields("gal_cub_used").value / 100 * rstUC.Fields(0).value), 2) 'multiply by usage charge
        rstUC.Close
    End If

    'Calculate the service line charge
    query = "SELECT CustomerServiceConnection.account, ServiceConnections.Amount, ServiceConnections.Description" & _
            " FROM CustomerServiceConnection INNER JOIN ServiceConnections ON CustomerServiceConnection.service_id = " & _
            " ServiceConnections.service_id WHERE (((CustomerServiceConnection.account)=" & account & "));"

    rstServ.Open query, CurrentProject.Connection
    If rstServ.BOF And rstServ.EOF Then
        'no service charge
    Else
        'take only the first service charge as techinically there can only be one
        'add the usage charge to the service charge
        charged = charged + rstServ.Fields(1).value
    End If
    
    rstServ.Close '###

    'now add other recurring charges such as Readiness to Serve Charge and any other fees (e.g. fire protection)
    query = "SELECT RatesAndCharges.account, RecurringCharges.charge_amount, RecurringCharges.charge_description" & _
            " FROM RecurringCharges INNER JOIN RatesAndCharges ON RecurringCharges.charge_id = RatesAndCharges.recurring_charge_id" & _
            " WHERE (((RatesAndCharges.account)=" & account & "));"

    rstRecur.Open query, CurrentProject.Connection
    'strTransaction = ""
    
    If rstRecur.BOF And rstRecur.EOF Then
        'nothing to do - no recurring service charge
    Else
        'could be many recurring charges - take them all
        Do While Not rstRecur.EOF
            charged = charged + rstRecur.Fields(1).value
            'TD 5/9/14 - this is wrong. strTransaction needs to be water usage (meter size & consumption)
            'strTransaction = strTransaction & " " & rstRecur.Fields(2).value
            rstRecur.MoveNext
        Loop
    End If
    
    rstRecur.Close '###
    
    'Now move the values into their respective buckets
    If rst.Fields("total_due").value <> 0 Then
        rst.Fields("past_due").value = rst.Fields("total_due").value
    End If
    
    current = charged
    rst.Fields("use_charge").value = charged
    
    current = current + rst.Fields("special_charge").value
    current = current - rst.Fields("special_credit").value
      
    
    If rst.Fields("prev_balance").value < 0 Then
        rst.Fields("prev_balance").value = Abs(rst.Fields("prev_balance").value)    'set to absolute value same as multiplying by -1
    End If
    
    rst.Fields("current_due").value = current
    
    If rst.Fields("prev_balance").value <> 0 Then
        prevBal = rst.Fields("prev_balance").value
        
        If prevBal - current > 0 Then
            prevBal = prevBal - current
            rst.Fields("prev_balance").value = prevBal
            rst.Fields("current_due").value = 0
            current = 0
        ElseIf prevBal - current < 0 Then
            prevBal = prevBal - current
            rst.Fields("prev_balance").value = 0
            prevBal = Abs(prevBal)
            rst.Fields("current_due").value = prevBal 'this will overwrite the current due with the prev bal
            current = prevBal
        Else
            rst.Fields("prev_balance").value = 0#
            rst.Fields("current_due").value = 0#
            current = 0#
        End If
    End If
        
    current = current + rst.Fields("past_due").value
    rst.Fields("total_due").value = current
    
    'append everything to the transaction file (money)
    Dim strMoney As String
    Dim recsAffected As Long
    Dim tmpQuery As String
    Dim dt As Date
    Dim rstDate As New ADODB.Recordset
    
    'look up the date and use this for posting the transactions
    tmpQuery = "SELECT TOP 1 Max(System.auto_id) AS MaxOfauto_id, System.system_date, System.use_system_date" & _
               " FROM System GROUP BY System.system_date, System.use_system_date ORDER BY Max(System.auto_id) DESC;"
         
    '###
    Dim tmpBool As Boolean
    
    rstDate.Open tmpQuery, CurrentProject.Connection
    
    '###
    If rstDate.BOF And rstDate.EOF Then
        dt = Now
    Else
        tmpBool = IIf(IsNull(rstDate.Fields(2).value), False, CBool(rstDate.Fields(2).value))
        
        If tmpBool Then
            dt = CDate(rstDate.Fields(1).value)
        Else
            dt = Now
        End If
    End If
    
    rstDate.Close
    strTransaction = Replace(strTransaction, "'", "''")
    strMoney = "INSERT INTO [Money] (m_month, m_day, m_year, account_number, [amount], [transaction], [code], [posted], behind_me, trans_date) " & _
                " VALUES ('" & Format(dt, "mm") & "','" & Format(dt, "dd") & "','" & Format(dt, "yyyy") & "','" & rst.Fields("account").value & "'," & _
                charged & ",'" & strTransaction & "'," & "'CHG'" & "," & "'Y'," & "0,#" & dt & "#)"
    
    CurrentProject.Connection.Execute strMoney, recsAffected
    
    If recsAffected < 1 Then
        'an error occurred do something to handle this
        Err.Raise vbObjectError + 5002, "cmdChargeMeter_Click of Form_BillingMenu", "No records were inserted for the following query: " & strMoney
    End If
    rst.Update
    rst.MoveNext
    ctr = ctr + 1
    DoEvents
    Me.recInner.Width = (Me.recInner.Width + process)
    'Debug.Print "Width = " & Me.recInner.Width
    'Me.Repaint
Loop

rst.Close

If Me.lstRollOver.ListCount > 0 Then
    Me.lstRollOver.Visible = True
    Me.lblRollOver.Visible = True
Else
    Call MsgBox("Metered Rate Account charging has completed. " & ctr & " accounts were charged.", vbInformation, "Complete")
End If

    Me.recOuter.Visible = False
    Me.recInner.Width = ProgressIndicator
    Me.recInner.Visible = False
   On Error GoTo 0
   Exit Sub

cmdChargeMetered_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdChargeMetered_Click of VBA Document Form_BillingMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdChargeMetered_Click of VBA Document Form_BillingMenu"

End Sub

Private Sub cmdPrintFlatRate_Click()
   On Error GoTo cmdPrintFlatRate_Click_Error

    SetMeteredQuery = False
    DoCmd.OpenForm "MeteredBillsWizard", , , , , , "FlatRate"

   On Error GoTo 0
   Exit Sub

cmdPrintFlatRate_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrintFlatRate_Click of VBA Document Form_BillingMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrintFlatRate_Click of VBA Document Form_BillingMenu"
End Sub

Private Sub cmdPrintMeter_Click()
   On Error GoTo cmdPrintMeter_Click_Error
    SetMeteredQuery = True
    DoCmd.OpenForm "MeteredBillsWizard", , , , , , "MeterRate"

   On Error GoTo 0
   Exit Sub

cmdPrintMeter_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrintMeter_Click of VBA Document Form_BillingMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrintMeter_Click of VBA Document Form_BillingMenu"
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdQuit_Click of VBA Document Form_BillingMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdQuit_Click of VBA Document Form_BillingMenu"

End Sub

Private Sub Form_Load()

Me.recOuter.Visible = False
Me.recInner.Visible = False
Me.recOuter.Width = ProgressBarWidth
Me.recInner.Width = ProgressIndicator
End Sub
