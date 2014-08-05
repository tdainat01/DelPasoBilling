Option Compare Database
Option Explicit

Sub Test()
Dim pad1 As String * 34         'Space between left margine and Prev_Balance
Dim Label1_Pad As String * 16   'Top labels for Prev_Balance Current Charge, Fire and Maint
Dim TextVal_Pad As String * 8  'Fixed field for Value fields
Dim Label2_Pad As String * 13   'Fixed field for tear-off labels
Dim Left_Pad As String * 9     'Fixed field for left padding
Dim AddrField As String * 52    'Fixed field for customer data
Dim ShortAddrField As String * 19 'Fixed field for customer data on tear-off sheet
Const Space1 As Integer = 6     'space between Prev_Balance and value
Const Space2 As Integer = 10     'space between field value and tear-off label
Const TearOffPad As Integer = 3
Dim tmpStr As String
Const strSpace As String = " "

pad1 = ""
Label1_Pad = "PREV BALANCE"
Label2_Pad = "PREV BAL"
TextVal_Pad = PadLeft("1234.56", Len(TextVal_Pad), strSpace)
Debug.Print "         1         2         3         4         5         6         7         8         9"
Debug.Print "123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890"
Debug.Print pad1 & Label1_Pad & TextVal_Pad & Space(TearOffPad) & Label2_Pad & RTrim(TextVal_Pad)

pad1 = ""
Label1_Pad = "CURRENT CHARGE"
Label2_Pad = "CUR. CHG"
TextVal_Pad = PadLeft("123.45", Len(TextVal_Pad), strSpace)
Debug.Print pad1 & Label1_Pad & TextVal_Pad & Space(TearOffPad) & Label2_Pad & RTrim(TextVal_Pad)

pad1 = ""
Label1_Pad = "SYST MAINT CHG"
Label2_Pad = "SMC"
TextVal_Pad = PadLeft("12.34", Len(TextVal_Pad), strSpace)
Debug.Print pad1 & Label1_Pad & TextVal_Pad & Space(TearOffPad) & Label2_Pad & RTrim(TextVal_Pad)

'pad1 = ""
'Label1_Pad = "PREV BALANCE"
'Label2_Pad = "PREV BAL"
'TextVal_Pad = PadLeft("1.23", Len(TextVal_Pad), " ")
'Debug.Print pad1 & Label1_Pad & TextVal_Pad & Space(3) & Label2_Pad & TextVal_Pad


End Sub

Sub TestPrintTextBills()

Dim arguments As String
Dim Path As String

Path = "c:\temp\metered_bills.txt"

arguments = "SELECT tmpPrintMeteredBill.* FROM tmpPrintMeteredBill WHERE (((tmpPrintMeteredBill.total_due)>0));" & _
    "|" & "AUG01" & "|" & "AUG31" & "|" & "2013"

Call OutPutTextBill(arguments, Path)
Call PrintFile(Path, "LPT2:")


End Sub

Public Sub GetTableInfo()

    Call TableInfo("money")

End Sub

Public Sub TestPostPayment()

Call PostPaymentTest(375, 94.5)

End Sub

Public Sub TestGetGUID()
    Debug.Print GetGUID ', vbInformation, "GUID Generated"
End Sub

Private Sub testCharging()
Dim lstAccount() As Long
    Call ChargeFlat(375)
    'Call ChargeMetered(60007)
    'lstAccount = Split("60001,60002,60003", ",")
    
End Sub

Private Sub ChargeFlat(account As Long)
Dim Cycle As Integer
Dim inputResult As Variant
Dim strQuery As String
Dim query As String
Dim strAccounts As String
Dim strFireSize As String
Dim strRollOverAccounts As String
Dim rst As New ADODB.Recordset
Dim rstRecur As New ADODB.Recordset
Dim rstServ As New ADODB.Recordset
Dim rstDate As New ADODB.Recordset
Dim ctr As Integer
Dim charged As Single
Dim prevBal As Single
Dim TotalDue As Single
Dim used As Single
Dim current As Single
Dim Last As Single
Dim strTransaction As String
Dim process As Long
Dim lRecs As Long

inputResult = InputBox("Cycle #", "Cycle", 1)

If inputResult = False Or inputResult = "" Then
    Exit Sub
Else
    Cycle = CInt(inputResult)
End If

If Cycle <= 0 Then
    Exit Sub
End If

'Me.recOuter.Visible = True
'Me.recInner.Visible = True

'@TODO - Remove this when launching for testing or production.
'MsgBox "Flat Rate Charges would now be calculated.", _
'    vbOKOnly + vbInformation, "This is a placeholder message"
''Exit Sub


'Me.lstRollOver.Visible = False
'Me.lblRollOver.Visible = False

'Check to see if there are unmatched records in the customer table on the rate codes
'This query gets both flat rate and metered accounts
strQuery = "SELECT customer.account FROM customer LEFT JOIN RecurringCharges ON customer.[rate_code] = RecurringCharges.[charge_code]" & _
            " WHERE (((RecurringCharges.charge_code) Is Null));"

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
    'Me.recOuter.Visible = False
    'Me.recInner.Visible = False
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
            " (customer.term_date)=#1/1/1900#) AND (customer.account = " & account & ") AND ((customer.current_read = 0 " & _
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
    'account = IIf(IsNull(rst.Fields("account").value), 0, rst.Fields("account").value)
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
    
'    If rst.fields("Oper1").value = "A" Then
'        'charged = charged + CSng(rst.Fields("Amount").Value)
'        charged = charged + CSng(rst.fields("Amount").value)
'        current = CSng(rst.fields("Amount").value)
'        rst.fields("use_charge").value = CSng(rst.fields("Amount").value)
'    Else
'        'charged = charged + (charged * CSng(rst.Fields("Amount").Value))
'        charged = charged * CSng(rst.fields("Amount").value)
'        current = CSng(rst.fields("Amount").value)
'        rst.fields("use_charge").value = CSng(rst.fields("Amount").value)
'    End If
    
    'add fire charge
'    If rst.fields("fire_size").value > 0 Then
'        'look up the fire charge from the rates table
'        Dim strTemp As String
'        Dim rstTemp As New ADODB.Recordset
'
'        strTemp = "SELECT * from Rates where line_size = " & rst.fields("fire_size").value & _
'            " and meter_resident = 'F'"
'        rstTemp.Open strTemp, CurrentProject.Connection, adOpenForwardOnly, adLockReadOnly
'        'should only return 1 row
'        Dim counter As Integer
'        Do While Not rstTemp.EOF
'            If counter > 1 Then
'                Exit Do
'            End If
'
'            'charged = charged + CSng(rstTemp.Fields("Amount").value)
'            current = current + CSng(rstTemp.fields("Amount").value)
'            counter = counter + 1
'            rstTemp.MoveNext
'        Loop
'        rstTemp.Close
'    End If
    
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
                current & ",'" & strTransaction & "'," & "'CHG'" & "," & "'Y'," & "0,#" & dt & "#)"
    
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
    'Me.recInner.Width = (Me.recInner.Width + process)
    'Debug.Print "Width = " & Me.recInner.Width
    'Me.Repaint
Loop

rst.Close

'If Me.lstRollOver.ListCount > 0 Then
'    Me.lstRollOver.Visible = True
'    Me.lblRollOver.Visible = True
'Else
    Call MsgBox("Flat Rate Account charging has completed. " & ctr & " accounts were charged.", vbInformation, "Complete")
'End If

'    Me.recOuter.Visible = False
'    Me.recInner.Width = ProgressIndicator
'    Me.recInner.Visible = False
   On Error GoTo 0
    Exit Sub

End Sub

Private Sub ChargeMetered(account As Long)

Dim Cycle As Integer
Dim inputResult As Variant
Dim rstRecur As New ADODB.Recordset
Dim rstServ As New ADODB.Recordset
Dim rst As New ADODB.Recordset
Dim rstUC As New ADODB.Recordset
Dim strQuery As String
'Dim account As Long
Dim strAccounts As String
Dim strFireSize As String
Dim query As String
Dim strRollOverAccounts As String
Dim ctr As Integer
Dim Service As Single
Dim charged As Single
Dim prevBal As Single
Dim current As Single
Dim used As Single
Dim Last As Single
Dim strTransaction As String
Dim lRecs As Long
Dim process As Long

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

'Me.recOuter.Visible = True
'Me.recInner.Visible = True

'@TODO - Remove this when launcing for testing or production.
'MsgBox "Metered Charges would now be calculated.", _
'    vbOKOnly + vbInformation, "This is a placeholder message"
'Exit Sub

'Me.lstRollOver.Visible = False
'Me.lblRollOver.Visible = False

'Check to see if there are unmatched records in the customer table on the rate codes
'This query gets both flat rate and metered accounts
strQuery = "SELECT customer.account FROM customer LEFT JOIN RecurringCharges ON customer.[rate_code] = RecurringCharges.[charge_code]" & _
            " WHERE (((RecurringCharges.charge_code) Is Null));"

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
    Exit Sub
End If

'Check to see if there are any unmatched records on line size and meter size
'#### This check can no longer be done. This is because the ServiceConnection Table only has a linesize of 1 for the Metered accounts
'#### meaning all the metered accounts will be unmatched. This is part of the logic in the old program which no longer works here.

'strQuery = " SELECT customer.account FROM customer LEFT JOIN ServiceConnections ON customer.[meter_size] = ServiceConnections.[LineSize]" & _
'            " WHERE (((ServiceConnections.LineSize) Is Null));"
'
'rst.Open strQuery, CurrentProject.Connection
'ctr = 0
'
'Do While Not rst.EOF
'    ctr = ctr + 1
'    strFireSize = strFireSize & rst.Fields(0).value & vbCrLf
'    rst.MoveNext
'Loop
'rst.Close

'If ctr > 0 Then
'    MsgBox "There are unmatched records for meter_size and line_size in the customer table on the following records:" _
'           & strFireSize, vbCritical + vbOKOnly, _
'        "Unmatched Line Size"
'    Exit Sub
'End If

'Now select all the metered customer accounts Filter out inactive accounts and check to see if last read is greater than the current read
'TODO Verify that using previous and current reads is a correct way to determine a metered customer.
strQuery = "SELECT customer.* FROM customer" & _
            " WHERE (((customer.cycle)=" & Cycle & ") AND ((customer.status)<>'I') AND ((customer.term_date) Is Null Or " & _
            " (customer.term_date)=#1/1/1900#) AND ((customer.current_read)>0) AND ((customer.previous_read)>0)) AND account = " & account & ";"

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
    'account = IIf(IsNull(rst.Fields("account").value), 0, rst.Fields("account").value)
    If account = 0 Then
        Err.Raise vbObjectError + 5001, "cmdChargeMeter_Click of Form_BillingMenu", "account cannot be zero"
    End If
    If rst.Fields("current_read").value < rst.Fields("previous_read").value Then
    'roll over
    'Me.lstRollOver.AddItem (rst.Fields("account").value)
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
        lRecs = 0
        rstUC.Open "SELECT [Amount] from MeterRates", CurrentProject.Connection
        If rstUC.BOF And rstUC.EOF Then
            'throw an error
            Exit Sub
        End If
        charged = (rst.Fields("gal_cub_used").value / 100) * CSng(rstUC.Fields(0).value)   'multiply by usage charge
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

'    If rst.fields("Oper1").value = "A" Then
'        charged = charged + CSng(rst.fields("Amount").value)
'        If CSng(rst.fields("Amount").value) = 0 Then
'            current = (used / 100)
'        Else
'            current = CSng(rst.fields("Amount").value)
'        End If
'
'        rst.fields("use_charge").value = CSng(rst.fields("Amount").value)
'    Else
'        'charged = charged + (charged * CSng(rst.Fields("Amount").value))
'        charged = charged * CSng(rst.fields("Amount").value)
'        current = CSng(rst.fields("Amount").value)
'        rst.fields("use_charge").value = CSng(rst.fields("Amount").value)
'    End If
    
    'add fire charge
'    If rst.fields("fire_size").value > 0 Then
'        'look up the fire charge from the rates table
'        Dim strTemp As String
'        Dim rstTemp As New ADODB.Recordset
'
'        strTemp = "SELECT * from Rates where line_size = " & rst.fields("fire_size").value & _
'            " and meter_resident = 'F'"
'        rstTemp.Open strTemp, CurrentProject.Connection, adOpenForwardOnly, adLockReadOnly
'        'should only return 1 row
'        Dim counter As Integer
'        Do While Not rstTemp.EOF
'            If counter > 1 Then
'                Exit Do
'            End If
'
'            charged = charged + CSng(rstTemp.fields("Amount").value)
'            counter = counter + 1
'            rstTemp.MoveNext
'        Loop
'        rstTemp.Close
'    End If

    'now add other recurring charges such as Readiness to Serve Charge and any other fees (e.g. fire protection)
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
    'Me.recInner.Width = (Me.recInner.Width + process)
    'Debug.Print "Width = " & Me.recInner.Width
    'Me.Repaint
Loop

rst.Close

'If Me.lstRollOver.ListCount > 0 Then
'    Me.lstRollOver.Visible = True
'    Me.lblRollOver.Visible = True
'Else
'    Call MsgBox("Metered Rate Account charging has completed. " & ctr & " accounts were charged.", vbInformation, "Complete")
'End If
'
'    Me.recOuter.Visible = False
'    Me.recInner.Width = ProgressIndicator
'    Me.recInner.Visible = False
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

Sub IncrementSequence()

Dim rst As New ADODB.Recordset
Dim query As String
Dim maxnum As Long
Dim ctr As Long
Dim rte As Long

rte = 2

query = "SELECT Max(CustomerRoutes.sequence) AS MaxOfsequence" & _
        " FROM CustomerRoutes WHERE route_id = " & rte
rst.Open query, CurrentProject.Connection
maxnum = IIf(IsNull(rst.Fields(0).value), 1, rst.Fields(0).value)
ctr = maxnum + 1
rst.Close

query = "select * from customerroutes where sequence is null"
rst.Open query, CurrentProject.Connection, adOpenDynamic, adLockOptimistic

If rst.BOF And rst.EOF Then
    Exit Sub
Else
    Do While Not rst.EOF
        rst.Fields("sequence").value = ctr
        rst.MoveNext
        ctr = ctr + 1
    Loop
End If

End Sub

Sub PostPaymentTest(acct As Long, amnt As Currency)
    'Validate all records payment codes are ok
    'query the money table for all transactions where posted is N. Then take those records,
    'and update the transactions table and the Customer table
    Dim query As String
    Dim tmpQuery As String
    Dim rst As New ADODB.Recordset
    Dim rstMon As New ADODB.Recordset
    Dim wtr As New ADODB.Recordset
    Dim rstRecur As New ADODB.Recordset
    Dim counter As Integer
    Dim tmpAccount As Long
    Dim strTransaction As String
    Dim inf As clsCustomer
    Dim oin As clsMoney

   On Error GoTo cmdPostPayment_Click_Error

'### THIS SECTION IS BEING SKIPPED
    'Modified this query to only select accounts older than 2 months ago
    query = "SELECT customer.*, NOW() as trans_date, 'MON' as code, " & _
            "" & amnt & " as amount, 'N' FROM customer " & _
            " WHERE customer.account = " & acct & " AND ((NOW())>DateAdd('m',-2,Now()));"

    '### TODO BECAUSE of the joins on the below table the query needs to be reqorked so only the first instance of each account is pulled up.
'    query = "SELECT customer.account, Money.trans_date, RecurringCharges.charge_code AS [transaction], Money.code, " & _
'            " Money.amount, Money.posted FROM RecurringCharges INNER JOIN ((customer INNER JOIN [Money] ON " & _
'            " customer.account = Money.account_number) INNER JOIN RatesAndCharges ON customer.account = " & _
'            " RatesAndCharges.account) ON RecurringCharges.charge_id = RatesAndCharges.recurring_charge_id" & _
'            " WHERE (((Money.posted)='N'));"

    rst.Open query, CurrentProject.Connection, adOpenDynamic, adLockPessimistic
    If rst.BOF And rst.EOF Then
        'nothing to do
        counter = 0
        GoTo exit_message
    End If
    tmpAccount = acct
    Set inf = FillCustomer(rst)
    'Do While Not rst.EOF
        'If tmpAccount = rst.Fields("account").value Then        'bug here. if account was entered twice in  load payment. only one gets processed
        '    rst.MoveNext
        'Else
            'tmpAccount = rst.Fields("account").value
            rstMon.Open "SELECT * FROM Money WHERE POSTED = 'N' and account_number = " & rst.Fields("account").value, CurrentProject.Connection
            
            If rstMon.BOF And rstMon.EOF Then
                'apparently there is no work to be done
                Exit Sub
            End If
            
                tmpQuery = "SELECT RatesAndCharges.account, RecurringCharges.charge_amount, RecurringCharges.charge_description" & _
                        " FROM RecurringCharges INNER JOIN RatesAndCharges ON RecurringCharges.charge_id = RatesAndCharges.recurring_charge_id" & _
                        " WHERE (((RatesAndCharges.account)=" & tmpAccount & "));"
            
                rstRecur.Open tmpQuery, CurrentProject.Connection
                strTransaction = ""
                
                If rstRecur.BOF And rstRecur.EOF Then
                    'nothing to do - no recurring service charge
                    strTransaction = "NONE"
                Else
                    'could be many recurring charges - take them all
                    Do While Not rstRecur.EOF
                        strTransaction = strTransaction & " " & rstRecur.Fields(2).value
                        strTransaction = Trim(strTransaction)
                        rstRecur.MoveNext
                    Loop
                End If
                
                rstRecur.Close
                strTransaction = Replace(strTransaction, "'", "''")
                
            query = "INSERT INTO [TRANSACTION] ([account],[trans_date],[transaction],[trans_note],[code],[amount]) VALUES(" & _
            acct & ",#" & Now() & "#,'" & _
            strTransaction & "','','MON'," & CCur(amnt) & ")"
            query = Replace(query, Chr(34), "in")
            'Insert into the Transaction Table
            'CurrentProject.Connection.Execute query
            Call WriteToFile("C:\DelPaso\UpdateTransaction.txt", query)
            
            'Now update the customer Table
            query = "SELECT deposit, past_due, prev_balance, current_due, special_credit, total_due, special_charge " & _
                    " FROM customer WHERE account = " & acct
            wtr.Open query, CurrentProject.Connection
            
            If wtr.BOF And wtr.EOF Then
                'An error will be raised, logged and the processing will halt.
                Err.Raise vbObjectError + 1024, "PostPaymentTest" & "cmdPostPayment", "No records were found for " & rst.Fields("account").value & _
                    ". Unable to continue with the Post Payment method"
            End If
            
            'declare variables for calculations
            Dim sTmp As String
            Dim lngRec As Long
            Dim oreo As Currency
            Dim preo As Currency
            Dim rreo As Currency
            Dim Amount As Currency
            
            Amount = amnt 'Round(rst.Fields("amount").value, 2)
            
            'Take the first row returned. Only one row should be returned.
            Select Case rst.Fields("code").value
                        
                Case Is = "SCH" '/* "special charges " */
                'sprintf (inf.special_charge, "%7.2f", atof (in.amount));
                'sprintf (inf.special_description, "%-25.25s", in.transaction)
                    inf.SpecialCharge = Round(rst.Fields("amount").value, 2)
                    inf.SpecialDescription = rst.Fields("trans_note").value
                Case Is = "CRE" '/* "special credits " */
                'sprintf (inf.special_credit, "%7.2f", atof (in.amount));
                'sprintf (inf.special_description, "%-25.25s", in.transaction);
                
                    inf.SpecialCredit = Round(rst.Fields("amount").value, 2)
                    inf.SpecialDescription = rst.Fields("trans_note").value
                    
                Case Is = "DEA" '/* "deposit applieds" */
                
'                oreo = atof (inf.deposit);
'                oreo -= atof (in.amount);
                oreo = inf.Deposit
                oreo = oreo - Amount
'                sprintf (inf.deposit, "%7.2f", oreo);
                inf.Deposit = oreo
'                oreo = atof (in.amount);
                oreo = Amount
'                preo = atof (inf.total_due);
                preo = inf.TotalDue
                If preo = oreo Then
                    inf.PastDue = 0
                    inf.CurrentDue = 0
                    inf.TotalDue = 0
'                    sprintf (inf.past_due, "%7.2f", 0.00);
'                    sprintf (inf.current_due, "%7.2f",0.00);
'                    sprintf (inf.total_due, "%7.2f", 0.00);
                ElseIf preo < oreo Then
                    inf.PastDue = 0
                    inf.CurrentDue = 0
                    inf.TotalDue = 0
'                    sprintf (inf.past_due, "%7.2f", 0.00);
'                    sprintf (inf.current_due, "%7.2f",0.00);
'                    sprintf (inf.total_due, "%7.2f", 0.00);
                    rreo = oreo - preo
'                    rreo += atof (inf.prev_balance);
                    rreo = rreo + inf.PrevBalance
'                    sprintf (inf.prev_balance,"%7.2f",rreo);
                    inf.PrevBalance = rreo
                
                Else
'                    {
'                    rreo = atof (inf.past_due);
                    rreo = inf.PastDue
                    If oreo > rreo Then
'                    if (oreo > rreo)
'                        {
                        rreo = oreo - rreo
'                        rreo = oreo - rreo;
                        inf.PastDue = 0
'                        sprintf(inf.past_due,"%7.2f",0.00);
                        rreo = inf.CurrentDue - rreo
'                        rreo=atof(inf.current_due)-rreo;
'                     sprintf(inf.current_due,"%7.2f", rreo);
                        inf.CurrentDue = rreo
'                        }
                    ElseIf oreo = rreo Then
'                    else if (oreo == rreo)
'                        {
                        inf.PastDue = 0
'                        sprintf(inf.past_due,"%7.2f",0.00);
'                        }
'                    Else
                    Else
'                        {
                        rreo = rreo - oreo
'                        rreo = rreo - oreo;
'                        sprintf(inf.past_due,"%7.2f",rreo);
                        inf.PastDue = rreo
'                        }
                    End If
                    
'                    rreo = atof (inf.past_due);
                    rreo = inf.PastDue
'                    rreo += atof (inf.current_due);
                    rreo = rreo + inf.CurrentDue
'                    sprintf (inf.total_due, "%7.2f", rreo);
                    inf.TotalDue = rreo
'                    }
            End If
'
'                }
                
                Case Is = "MON"
'                /* "water payments  " */
'                oreo = atof (in.amount);
                oreo = Amount 'oin.Amount
'                preo = atof (inf.total_due);
                preo = inf.TotalDue
'                if (preo == oreo)
                If preo = oreo Then
'                    {
'                    strcpy (inf.past_due, "0.00");
                    inf.PastDue = 0
'                    strcpy (inf.current_due, "0.00");
                    inf.CurrentDue = 0
'                    strcpy (inf.total_due, "0.00");
                    inf.TotalDue = 0
'                    }
                ElseIf preo < oreo Then
'                else if (preo < oreo)
'                    {
'                    strcpy (inf.past_due, "0.00");
                    inf.PastDue = 0
'                    strcpy (inf.current_due, "0.00");
                    inf.CurrentDue = 0
'                    strcpy (inf.total_due, "0.00");
                    inf.TotalDue = 0
'                    rreo = oreo - preo;
                    rreo = oreo - preo
'                    rreo += atof (inf.prev_balance);
                    rreo = rreo + inf.PrevBalance
'                    sprintf (inf.prev_balance,"%7.2f",rreo);
                    inf.PrevBalance = rreo
'                    }
'                Else
                Else
'                    {
'                    rreo = atof (inf.past_due);
                    rreo = inf.PastDue
                    If oreo > rreo Then
'                    if (oreo > rreo)
'                        {
'                        rreo = oreo - rreo;
                        rreo = oreo - rreo
'                        strcpy (inf.past_due,"0.00");
                        inf.PastDue = 0
'                        rreo=atof(inf.current_due)-rreo;
                        rreo = inf.CurrentDue - rreo
'                     sprintf(inf.current_due,"%7.2f", rreo);
                        rreo = inf.CurrentDue
'                        }
                    ElseIf oreo = rreo Then
'                    else if (oreo == rreo)
'                        {
'                        strcpy (inf.past_due,"0.00");
                        inf.PastDue = 0
'                        }
                    Else
'                    Else
'                        {
'                        rreo = rreo - oreo;
                        rreo = rreo - oreo
'                        sprintf(inf.past_due,"%7.2f",rreo);
                        inf.PastDue = rreo
'                        }
                    End If
'                    rreo = atof (inf.past_due);
                    rreo = inf.PastDue
'                    rreo += atof (inf.current_due);
                    rreo = rreo + inf.CurrentDue
'                    sprintf (inf.total_due, "%7.2f", rreo);
                    inf.TotalDue = rreo
'                    }
                End If
                
                Case Is = "DEP" '/* "customer deposits  " */
'                /* "water deposits  " */
'                rreo = atof (inf.deposit);
                rreo = inf.Deposit
                If rreo <> 0 Then
                    Call MsgBox("Warning: account " & inf.account & " has a deposit already.", vbOKOnly + vbExclamation, "Warning")
'                if (rreo != 0)
'                    {
'                    clear ();
'                    locate (12, 1);
'                     printw ("Warning account #%s", in.account_number);
'                    printw (" has deposit already.");
'                    }
                End If
'                rreo += atof (in.amount);
                rreo = rreo + Amount 'oin.Amount
'                sprintf (inf.deposit, "%7.2f", rreo);
                inf.Deposit = rreo
                
                Case Is = "TRD" '/* "transfer deposit" */
'                oreo = atof (inf.deposit);
                oreo = inf.Deposit
'                oreo -= atof (in.amount);
                oreo = oreo - Amount 'in.amount
'                sprintf (inf.deposit, "%7.2f", oreo);
                inf.Deposit = oreo
                
                Case Is = "RED" '/* "refund deposits " */
'                oreo = atof (inf.deposit);
                oreo = inf.Deposit
'                oreo -= atof (in.amount);
                oreo = oreo - Amount 'oin.Amount
'                sprintf (inf.deposit, "%7.2f", oreo);
                inf.Deposit = oreo
                
                Case Is = "REB" '/* "refund balances " */
'                oreo = atof (inf.prev_balance);
                oreo = inf.PrevBalance
'                oreo -= atof (in.amount);
                oreo = oreo - Amount 'oin.Amount
'                sprintf (inf.prev_balance, "%7.2f", oreo);
                inf.PrevBalance = oreo
                
                Case Is = "TRB" '/* "transfer balance" */
'                oreo = atof (inf.prev_balance);
                oreo = inf.PrevBalance
'                oreo -= atof (in.amount);
                oreo = oreo - Amount 'oin.Amount
'                sprintf (inf.prev_balance, "%7.2f", oreo);
                inf.PrevBalance = oreo
                
                Case Is = "CHG" '/* "computer charges" */
'                /* "computer charges" */
'                /* do nothing */
               
                Case Is = "VOI" '/* "Voided entrys   " */
                    'do nothing
                Case Else
                    'TODO
            End Select
                
                'Use_Charge = Current_Due
                
                'Do the update
                sTmp = " UPDATE customer SET deposit = " & inf.Deposit & ", past_due = " & inf.PastDue & _
                    ", prev_balance = " & inf.PrevBalance & ", current_due = " & inf.CurrentDue & ", special_credit = " & _
                    inf.SpecialCredit & ", total_due = " & inf.TotalDue & ", special_charge = " & inf.SpecialCharge & _
                    ", special_description = '" & inf.SpecialDescription & _
                    "' WHERE account = " & inf.account
                'CurrentProject.Connection.Execute sTmp, lngRec
                '## TEST - write the data to a file for review.
                Call WriteToFile("C:\DelPaso\UpdateCustomer.txt", sTmp)
                lngRec = 1
                
                If lngRec <= 0 Then
                    Err.Raise vbObjectError + 1025, "mTestFunctions" & ".cmdPostPayment", "An error occurred in trying to update the customer table " & _
                        " for " & rst.Fields("account").value & ". No update occurred. The query that ran was: " & sTmp
                End If
                
                wtr.Close
                    
            'Update the money table to reflect the field has been posted
            'rst.Fields("posted").value = "Y"
    
            rst.MoveNext
            counter = counter + 1
        'End If
    'Loop
    
exit_message:
    'end if with a message that all payments have been posted
    If counter = 1 Then
        Call MsgBox(counter & " payment was posted.", vbInformation, "Posted")
    Else
        Call MsgBox(counter & " payments were posted.", vbInformation, "Posted")
    End If
    
    On Error GoTo 0
   Exit Sub

cmdPostPayment_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description
Dim F As Boolean
Dim msg As String
    F = LogError(Err.Number, Err.source, Err.Description)
    If F Then
        msg = "This error has been logged"
    Else
        msg = "This error has NOT been logged"
    End If
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPostPayment_Click of VBA Document Form_DPM Main Menu." & msg

End Sub

Private Sub TestInputbox()
Dim ans As String

        ans = InputBox("The account you're attempting to add already exists in the meter reads table. Are you sure you want to add it again?(Y/N)" _
            , "Existing Account?", "N")
        If ans = "" Then
            Call MsgBox("Please enter either Y or N or y or n (Y = Yes, N = No). Any other value is invalid.", vbExclamation, "Invalid Entry")
            Exit Sub
        ElseIf LCase(ans) = "n" Then
            Exit Sub
        ElseIf LCase(ans) = "y" Then
            'do nothing
        Else
            Call MsgBox("Please enter either Y or N or y or n (Y = Yes, N = No). Any other value is invalid.", vbExclamation, "Invalid Entry")
            Exit Sub
        End If
End Sub

Public Function PadRight(text As Variant, totalLength As Integer, padCharacter As String) As String
    'PadRight = Left(CStr(text) & String(totalLength - Len(CStr(text)), padCharacter), totalLength)
    Dim PadLength As Integer
    Dim x As Integer
    
    PadLength = totalLength - Len(text)
      Dim PadString As String
      For x = 1 To PadLength
         PadString = padCharacter & PadString
      Next
      PadRight = text + PadString
End Function

Public Function PadLeft(text As Variant, totalLength As Integer, padCharacter As String) As String
    'PadLeft = String(totalLength - Len(CStr(Trim(text))), padCharacter) & CStr(text)
    Dim PadLength As Integer
    Dim x As Integer
    
    PadLength = totalLength - Len(Trim(text))
      Dim PadString As String
      For x = 1 To PadLength
         PadString = PadString & padCharacter
      Next
      PadLeft = RTrim(PadString + text)
End Function
