Option Compare Database
Option Explicit

Public Sub testPosting()
    Call cmdPostPayment_Click
End Sub

Public Sub ExecutePost()
   On Error GoTo ExecutePost_Error

    Call cmdPostPayment_Click

   On Error GoTo 0
   Exit Sub

ExecutePost_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure ExecutePost of VBA Document Form_DPM Main Menu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure ExecutePost of VBA Document Form_DPM Main Menu"
End Sub
Private Sub cmdBillingMenu_Click()
   On Error GoTo cmdBillingMenu_Click_Error

    DoCmd.OpenForm "BillingMenu"

   On Error GoTo 0
   Exit Sub

cmdBillingMenu_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdBillingMenu_Click of VBA Document Form_DPM Main Menu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdBillingMenu_Click of VBA Document Form_DPM Main Menu"
End Sub

Private Sub cmdEnterMeterReads_Click()
   On Error GoTo cmdEnterMeterReads_Click_Error

    DoCmd.OpenForm "MeterMenu", acNormal

   On Error GoTo 0
   Exit Sub

cmdEnterMeterReads_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdEnterMeterReads_Click of VBA Document Form_DPM Main Menu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdEnterMeterReads_Click of VBA Document Form_DPM Main Menu"
End Sub

Private Sub cmdLoadPayments_Click()
   On Error GoTo cmdLoadPayments_Click_Error

    DoCmd.OpenForm "LoadPayments2", acNormal

   On Error GoTo 0
   Exit Sub

cmdLoadPayments_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdLoadPayments_Click of VBA Document Form_DPM Main Menu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdLoadPayments_Click of VBA Document Form_DPM Main Menu"
End Sub

Private Sub cmdMasterMaintAcct_Click()
   On Error GoTo cmdMasterMaintAcct_Click_Error

    DoCmd.OpenForm "Opt1Form", acNormal, , , , , "account"

   On Error GoTo 0
   Exit Sub

cmdMasterMaintAcct_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdMasterMaintAcct_Click of VBA Document Form_DPM Main Menu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdMasterMaintAcct_Click of VBA Document Form_DPM Main Menu"
End Sub

Private Sub cmdMasterMaintParc_Click()
   On Error GoTo cmdMasterMaintParc_Click_Error

    DoCmd.OpenForm "Opt1Form", acNormal, , , , , "mastpar"

   On Error GoTo 0
   Exit Sub

cmdMasterMaintParc_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdMasterMaintParc_Click of VBA Document Form_DPM Main Menu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdMasterMaintParc_Click of VBA Document Form_DPM Main Menu"
End Sub

Private Sub cmdPaymentEditListing_Click()
   On Error GoTo cmdPaymentEditListing_Click_Error

    DoCmd.OpenForm "PaymentMenu", acNormal

   On Error GoTo 0
   Exit Sub

cmdPaymentEditListing_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPaymentEditListing_Click of VBA Document Form_DPM Main Menu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPaymentEditListing_Click of VBA Document Form_DPM Main Menu"
End Sub

Private Sub cmdPaymentMaint_Click()
    DoCmd.OpenForm "EditPayments", acNormal
End Sub

Private Sub cmdPostPayment_Click()
    
    'Validate all records payment codes are ok
    'query the money table for all transactions where posted is N. Then take those records,
    'and update the transactions table and the customer table
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

    'Modified this query to only select accounts older than 2 months ago
    query = "SELECT customer.*, Money.trans_date, Money.code, " & _
            " Money.amount, Money.posted FROM customer " & _
            " INNER JOIN [Money] ON customer.account = Money.account_number" & _
            " WHERE (((Money.posted)='N') AND ((Money.trans_date)>DateAdd('m',-2,Now())));"

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
    tmpAccount = 0
    Set inf = FillCustomer(rst)
    
    Do While Not rst.EOF
        Set inf = FillCustomer(rst)

        'If tmpAccount = rst.Fields("account").value Then        'bug here. if account was entered twice in  load payment. only one gets processed
        '    rst.MoveNext
        'Else
            tmpAccount = rst.Fields("account").value
            rstMon.Open "SELECT * FROM [Money] WHERE POSTED = 'N' and account_number = " & tmpAccount, CurrentProject.Connection
            
            If rstMon.BOF And rstMon.EOF Then
                'apparently there is no work to be done
                Exit Sub
            Else
                Set oin = FillMoney(rstMon)
            End If
            
            rstMon.Close
            
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
            rst.Fields("account").value & ",#" & rst.Fields("trans_date").value & "#,'" & _
            strTransaction & "','','" & rst.Fields("code").value & "'," & CCur(rst.Fields("amount").value) & ")"
            query = Replace(query, Chr(34), "in")
            'Insert into the Transaction Table
            CurrentProject.Connection.Execute query
                    
            'Now update the customer Table
            query = "SELECT deposit, past_due, prev_balance, current_due, special_credit, total_due, special_charge " & _
                    " FROM customer WHERE account = " & rst.Fields("account").value
            wtr.Open query, CurrentProject.Connection
            
            If wtr.BOF And wtr.EOF Then
                'An error will be raised, logged and the processing will halt.
                Err.Raise vbObjectError + 1024, Me.Name & " " & "cmdPostPayment", "No records were found for " & rst.Fields("account").value & _
                    ". Unable to continue with the Post Payment method"
            End If
            
            'declare variables for calculations
            Dim sTmp As String
            Dim lngRec As Long
            Dim oreo As Currency
            Dim preo As Currency
            Dim rreo As Currency
            Dim Amount As Currency
            
            Amount = oin.Amount
            
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
'                       fix 8/8/2013
                        inf.CurrentDue = rreo ' fix chuck
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
                CurrentProject.Connection.Execute sTmp, lngRec
                
                If lngRec <= 0 Then
                    Err.Raise vbObjectError + 1025, Me.Name & ".cmdPostPayment", "An error occurred in trying to update the customer table " & _
                        " for " & rst.Fields("account").value & ". No update occurred. The query that ran was: " & sTmp
                End If
                wtr.Close
                    
            'Update the money table to reflect the field has been posted
            rst.Fields("posted").value = "Y"
    
            rst.MoveNext
            counter = counter + 1
        'End If
    Loop
    
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

Private Sub cmdQuit_Click()

'Dim query As String
'Dim lRecs As Long
'   On Error GoTo cmdQuit_Click_Error
'
'    'Save system Date in maintenance table
'    query = "INSERT INTO SYSTEM(system_date) VALUES(#" & Format(Now, "Short Date") & "#)"
'    CurrentProject.Connection.Execute query, lRecs

    Application.Quit acQuitPrompt

   On Error GoTo 0
   Exit Sub

cmdQuit_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdQuit_Click of VBA Document Form_DPM Main Menu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdQuit_Click of VBA Document Form_DPM Main Menu"
End Sub

Private Sub cmdRatesAndCharges_Click()
   On Error GoTo cmdRatesAndCharges_Click_Error

    DoCmd.OpenForm "frmRatesAndBilling", acNormal

   On Error GoTo 0
   Exit Sub

cmdRatesAndCharges_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdRatesAndCharges_Click of VBA Document Form_DPM Main Menu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdRatesAndCharges_Click of VBA Document Form_DPM Main Menu"
End Sub

Private Sub cmdRateWindowCodes_Click()
   On Error GoTo cmdRateWindowCodes_Click_Error

    DoCmd.OpenForm "RateMenu", acNormal

   On Error GoTo 0
   Exit Sub

cmdRateWindowCodes_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdRateWindowCodes_Click of VBA Document Form_DPM Main Menu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdRateWindowCodes_Click of VBA Document Form_DPM Main Menu"
End Sub

Private Sub cmdReportMenu_Click()
   On Error GoTo cmdReportMenu_Click_Error

    DoCmd.OpenForm "ReportMenu", acNormal

   On Error GoTo 0
   Exit Sub

cmdReportMenu_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdReportMenu_Click of VBA Document Form_DPM Main Menu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdReportMenu_Click of VBA Document Form_DPM Main Menu"
End Sub

Private Sub cmdRoutes_Click()
   On Error GoTo cmdRoutes_Click_Error

    'Call MsgBox("This feature has not been coded yet. Please try again later.", vbExclamation, "Not Implemented Yet")
    'Exit Sub
    'this will open the routes form.
    DoCmd.OpenForm "frmRoutes", acNormal
   On Error GoTo 0
   Exit Sub

cmdRoutes_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdRoutes_Click of VBA Document Form_DPM Main Menu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdRoutes_Click of VBA Document Form_DPM Main Menu"
End Sub

Private Sub cmdSearch_Click()
   On Error GoTo cmdSearch_Click_Error

    'default to the Opt 4 Form as being the form sender
    DoCmd.OpenForm "frmSearch", acNormal, , , , , "Opt 4 Form"

   On Error GoTo 0
   Exit Sub

cmdSearch_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdSearch_Click of VBA Document Form_DPM Main Menu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdSearch_Click of VBA Document Form_DPM Main Menu"
End Sub

Private Sub cmdMaint_Click()

   On Error GoTo cmdMaint_Click_Error
    DoCmd.OpenForm "RateMenu", acNormal

   On Error GoTo 0
   Exit Sub

cmdMaint_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdMaint_Click of VBA Document Form_DPM Main Menu")
    'MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdMaint_Click of VBA Document Form_DPM Main Menu"
End Sub

Private Sub cmdSettings_Click()
   On Error GoTo cmdSettings_Click_Error

    DoCmd.OpenForm "frmSettings"

   On Error GoTo 0
   Exit Sub

cmdSettings_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    'Call LogError(errNum, errSource, errMsg & " in procedure cmdSettings_Click of VBA Document Form_DPM Main Menu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdSettings_Click of VBA Document Form_DPM Main Menu"
End Sub

Private Sub Form_Load()
'Delete everything from the system date table.
' set the system date to today
'set the use_system_date flag to false
Dim rst As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim sQuery As String
Dim F As Boolean

   On Error GoTo Form_Load_Error

DoCmd.Maximize

sQuery = "DELETE FROM System"
CurrentProject.Connection.Execute sQuery

F = False
sQuery = "INSERT INTO System (system_date, use_system_date) VALUES ('" & Now & "'," & F & ")"
CurrentProject.Connection.Execute sQuery


   On Error GoTo 0
   Exit Sub

Form_Load_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Load of VBA Document Form_DPM Main Menu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Load of VBA Document Form_DPM Main Menu"

End Sub

Private Sub txtSystemDate_LostFocus()

   On Error GoTo txtSystemDate_LostFocus_Error

Dim rst As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim sQuery As String
Dim F As Boolean

If Not IsDate(txtSystemDate.text) Then
    Call MsgBox("I'm sorry - The date value enetered into the system date field is not a valid date. Please fix this before going on.", _
    vbCritical, "Invalid Date")
    Exit Sub
End If

sQuery = "DELETE FROM System"
CurrentProject.Connection.Execute sQuery

F = True
sQuery = "INSERT INTO System (system_date, use_system_date) VALUES ('" & Me.txtSystemDate.text & "'," & F & ")"
CurrentProject.Connection.Execute sQuery
    

   On Error GoTo 0
   Exit Sub

txtSystemDate_LostFocus_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure txtSystemDate_LostFocus of VBA Document Form_DPM Main Menu")
    'MsgBox "Error " & errNum & " (" & errMsg & ") in procedure txtSystemDate_LostFocus of VBA Document Form_DPM Main Menu"
End Sub
