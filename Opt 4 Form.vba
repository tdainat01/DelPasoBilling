Option Compare Database
Option Explicit

Dim vOpenArgs4 As Variant

Private Sub chkService_Click()
   On Error GoTo chkService_Click_Error

    If Me.chkService = True Then
        imgRedFlag.Visible = True
    ElseIf Me.chkService = False Then
        imgRedFlag.Visible = False
    Else
        imgRedFlag.Visible = False
    End If

   On Error GoTo 0
   Exit Sub

chkService_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure chkService_Click of VBA Document Form_Opt 4 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure chkService_Click of VBA Document Form_Opt 4 Form"
End Sub

Private Sub cmdNotes_Click()

   On Error GoTo cmdNotes_Click_Error
    If Me.txtAccount = "" Then
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdNotes_Click of VBA Document Form_Opt 4 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdNotes_Click of VBA Document Form_Opt 4 Form"
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdQuit_Click of VBA Document Form_Opt 4 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdQuit_Click of VBA Document Form_Opt 4 Form"
End Sub

Private Sub cmdUsage_Click()
Dim rst As New ADODB.Recordset
Dim query As String

   On Error GoTo cmdUsage_Click_Error

If IsNull(Me.txtAccount) Or Me.txtAccount = "" Or IsEmpty(Me.txtAccount) Or Me.txtAccount <= 0 Then
    Call MsgBox("No valid account number has been specified. Please enter a valid account number first and then try again.", _
        vbExclamation Or vbSystemModal, "No Account")
    Exit Sub
End If
    
    query = "select top 2 batch_date, normal_read from MeterReads where posted = 'Y' and account = " & Me.txtAccount
    rst.Open query, CurrentProject.Connection
    If rst.BOF And rst.EOF Then
        Call MsgBox("No meter read information has been returned for " & Me.txtAccount & ".", _
        vbExclamation Or vbSystemModal, "No Meter Reads")
        Exit Sub
    Else
        DoCmd.OpenForm "frmUsage", acNormal, , , acFormReadOnly, acDialog, Me.txtAccount
    End If

   On Error GoTo 0
   Exit Sub

cmdUsage_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdUsage_Click of VBA Document Form_Opt 4 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdUsage_Click of VBA Document Form_Opt 4 Form"
End Sub

Private Sub Form_Load()

   On Error GoTo Form_Load_Error

If IsNull(Me.OpenArgs) Then
    'do nothing
Else
    vOpenArgs4 = Me.OpenArgs
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

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Load of VBA Document Form_Opt 4 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Load of VBA Document Form_Opt 4 Form"

End Sub

Private Sub txtAccount_Enter()
Dim acct As Long
   
   On Error GoTo txtAccount_Enter_Error

    acct = IIf(IsNull(Me.txtAccount) Or Me.txtAccount = "", 0, Me.txtAccount)
    'is there work to do?
    If acct = 0 Then
        'did we get anything from the form being opened?
        If IsNull(vOpenArgs4) Or vOpenArgs4 = "" Or vOpenArgs4 = 0 Or IsEmpty(vOpenArgs4) Then
            'nothing to do
            Exit Sub
        Else
            If vOpenArgs4 = acct Then
                'do nothing
            Else
                If acct = 0 Then
                    Me.txtAccount = vOpenArgs4
                    acct = vOpenArgs4
                End If
            End If
        End If
    End If
    
    Call CustomerLookup(acct)
    Me.txtAccount.SetFocus

   On Error GoTo 0
   Exit Sub

txtAccount_Enter_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure txtAccount_Enter of VBA Document Form_Opt 4 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure txtAccount_Enter of VBA Document Form_Opt 4 Form"

End Sub

Private Sub txtAccount_LostFocus()
Dim acct As Long

   On Error GoTo txtAccount_LostFocus_Error
    
    acct = IIf(IsNull(Me.txtAccount) Or Me.txtAccount = "", 0, Me.txtAccount)
    'is there work to do?
    If acct = 0 Then
        'did we get anything from the form being opened?
        If IsNull(vOpenArgs4) Or vOpenArgs4 = "" Or vOpenArgs4 = 0 Or IsEmpty(vOpenArgs4) Then
            'nothing to do
            Exit Sub
        Else
            If vOpenArgs4 = acct Then
                'do nothing
            Else
                If acct = 0 Then
                    Me.txtAccount = vOpenArgs4
                    acct = vOpenArgs4
                End If
            End If
        End If
    End If
    
    Call CustomerLookup(acct)
    Me.txtAccount.SetFocus

   On Error GoTo 0
   Exit Sub

txtAccount_LostFocus_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure txtAccount_LostFocus of VBA Document Form_Opt 4 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure txtAccount_LostFocus of VBA Document Form_Opt 4 Form"

End Sub

Private Sub CustomerLookup(acct As Long)
Dim rst As New ADODB.Recordset
Dim query As String
Dim pNum1 As String
Dim pNum2 As String
Dim pNum3 As String
   
   On Error GoTo CustomerLookup_Error
imgRedFlag.Visible = False

If acct <= 0 Then
    Exit Sub
End If

    query = "SELECT * FROM customer where ACCOUNT = " & Me.txtAccount
    rst.Open query, CurrentProject.Connection
    
    If rst.BOF And rst.EOF Then
        Call ResetValues
        Call MsgBox("The account you entered does not exist. Please enter a different account number and try again.", _
            vbExclamation, "Account Not Found")
        Me.txtAccount = ""
        Exit Sub
        'record not found. Create one?
    Else
        Do While Not rst.EOF
            Me.txtAccount = acct
            Me.txtGroup = rst.Fields("group").value
            Me.txtMastPar = rst.Fields("mastpar").value
            Me.txtCycle = rst.Fields("cycle").value
            Me.txtMfgCode = rst.Fields("mfg_code").value
            Me.txtStartDate = rst.Fields("start_date").value
            Me.txtStatus = rst.Fields("status").value
            Me.txtMeterNum = rst.Fields("meter_number").value
            Me.txtTermDate = rst.Fields("term_date").value
            
            If rst.Fields("out_town").value = False Then
                Me.chkOutTown.value = False
            Else
                Me.chkOutTown = True
            End If
            
            Me.txtMeterSize = rst.Fields("meter_size").value
            Me.txtPropertyUse = rst.Fields("property_use").value
            
            If rst.Fields("backflow").value = False Then
                Me.chkBackflow.value = False
            Else
                Me.chkBackflow.value = True
            End If
            
            If rst.Fields("service_discon").value = False Then
                Me.chkService = False
                imgRedFlag.Visible = False
            Else
                Me.chkService = True
                imgRedFlag.Visible = True
            End If
            
            Me.txtFireSize = rst.Fields("fire_size").value
            Me.txtUnitofMeasure = rst.Fields("unit_measure").value
            Me.txtCurrentRead = rst.Fields("current_read").value
            Me.txtCurrentDate = rst.Fields("current_date").value
            Me.txtRateCode = rst.Fields("rate_code").value
            Me.txtPreviousRead = rst.Fields("previous_read").value
            Me.txtGalsCubUsed = rst.Fields("gal_cub_used").value
            Me.txtMeterSite = rst.Fields("meter_site").value
            Me.txtDeposit = Round(rst.Fields("deposit").value, 2)
            Me.txtUseCharge = Round(rst.Fields("use_charge").value, 2)
            Me.txtPastDue = Round(rst.Fields("past_due").value, 2)
            Me.txtPrevBalance = Round(rst.Fields("prev_balance").value, 2)
            Me.txtCurrentDue = Round(rst.Fields("current_due").value, 2)
            Me.txtSpecialCredit = Round(rst.Fields("special_credit").value, 2)
            Me.txtTotalDue = Round(rst.Fields("total_due").value, 2)
            Me.txtSpecialCharge = Round(rst.Fields("special_charge").value, 2)
            Me.txtSpecialDescr = rst.Fields("special_description").value
            Me.txtPhysicalAddress = rst.Fields("phy_address").value
            'No lien field?
            Me.txtCareOfName = rst.Fields("care_of").value
            Me.txtName = rst.Fields("name").value
            Me.txtBillName = rst.Fields("bill_name").value
            Me.txtAddress = rst.Fields("addr1").value
            Me.txtCity = rst.Fields("city").value
            Me.txtState = rst.Fields("state").value
            Me.txtZip = rst.Fields("zip").value
            'no comment field?
            rst.MoveNext
        Loop
        rst.Close
        
        'Now open the phones
        query = "SELECT Phones.Phone1, Phones.Phone2, Phones.Phone3" & _
                " FROM Phones WHERE (((Phones.CustomerID)=" & CLng(Me.txtAccount) & "));"
        
        rst.Open query, CurrentProject.Connection
        
        If rst.BOF And rst.EOF Then
            Me.txtWorkPhone = ""
            Me.txtHomePhone = ""
            Me.txtMobilePhone = ""
        Else
            Do While Not rst.EOF
                pNum1 = IIf(IsNull(rst.Fields("phone1").value), "", rst.Fields("phone1").value)
                pNum2 = IIf(IsNull(rst.Fields("phone2").value), "", rst.Fields("phone2").value)
                pNum3 = IIf(IsNull(rst.Fields("phone3").value), "", rst.Fields("phone3").value)
            rst.MoveNext
            Loop
            
            Me.txtWorkPhone = pNum1
            Me.txtHomePhone = pNum2
            Me.txtMobilePhone = pNum3
        End If
        
        rst.Close
    
        'get the route number if assigned
        query = "SELECT Routes.route_name" & _
                " FROM Routes INNER JOIN (customer INNER JOIN CustomerRoutes ON customer.account = " & _
                " CustomerRoutes.account_num) ON Routes.route_id = CustomerRoutes.route_id" & _
                " WHERE (((customer.account)=" & CLng(Me.txtAccount) & "));"
        
        rst.Open query, CurrentProject.Connection
        
        If rst.BOF And rst.EOF Then
            Me.txtRoute = ""
        Else
            Me.txtRoute = IIf(IsNull(rst.Fields(0).value), "", rst.Fields(0).value)
        End If
    End If
    
   On Error GoTo 0
   Exit Sub

CustomerLookup_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure CustomerLookup of VBA Document Form_Opt 4 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure CustomerLookup of VBA Document Form_Opt 4 Form"

End Sub

Private Sub cmdLiens_Click()

   On Error GoTo cmdLiens_Click_Error

If Me.txtAccount <> "" And IsNumeric(Me.txtAccount) Then
    Dim strQuery As String
    Dim rst As New ADODB.Recordset
    Dim ctr As Integer
    strQuery = "select * from customer where account = " & Me.txtAccount & " and lien <> 'A'"
    rst.Open strQuery, CurrentProject.Connection, adOpenForwardOnly
    Do While Not rst.EOF
        ctr = ctr + 1
        rst.MoveNext
    Loop
    
    If ctr > 0 Then
        'do something
    Else
        MsgBox "No Lien", vbInformation + vbOKOnly, "No Lien Found"
    End If
Else
    'do nothing - no account record selected
End If


   On Error GoTo 0
   Exit Sub

cmdLiens_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdLiens_Click of VBA Document Form_Opt 4 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdLiens_Click of VBA Document Form_Opt 4 Form"

End Sub

Private Sub cmdNext_Click()
Dim account As Long
Dim qry As String
Dim rst As New ADODB.Recordset
Dim rstPhones As New ADODB.Recordset

   On Error GoTo cmdNext_Click_Error
imgRedFlag.Visible = False
account = GetNextRecord(CLng(Me.txtAccount))
qry = "select * from customer where account = " & account '& " order by account"
rst.Open qry, CurrentProject.Connection

If rst.BOF And rst.EOF Then
    'otherwise do nothing
    Call MsgBox("The end of the file was encountered. No further records to display.", vbExclamation, "End of File")
    Exit Sub
End If

Do While Not rst.EOF
        Me.txtAccount = rst.Fields("account").value
        Me.txtGroup = rst.Fields("group").value
        Me.txtMastPar = rst.Fields("mastpar").value
        Me.txtCycle = rst.Fields("cycle").value
        Me.txtMfgCode = rst.Fields("mfg_code").value
        Me.txtStartDate = rst.Fields("start_date").value
        Me.txtStatus = rst.Fields("status").value
        Me.txtMeterNum = rst.Fields("meter_number").value
        Me.txtTermDate = rst.Fields("term_date").value
        Me.chkOutTown = rst.Fields("out_town").value
        Me.txtMeterSize = rst.Fields("meter_size").value
        Me.txtPropertyUse = rst.Fields("property_use").value
        Me.chkBackflow = rst.Fields("backflow").value
        Me.chkService = rst.Fields("service_discon").value
        If Not IsNull(Me.chkService) Then
            If Me.chkService = True Then
                imgRedFlag.Visible = True
            End If
        End If
        Me.txtFireSize = rst.Fields("fire_size").value
        Me.txtUnitofMeasure = rst.Fields("unit_measure").value
        Me.txtCurrentRead = rst.Fields("current_read").value
        Me.txtCurrentDate = rst.Fields("current_date").value
        Me.txtRateCode = rst.Fields("rate_code").value
        Me.txtPreviousRead = rst.Fields("previous_read").value
        Me.txtPreviousDate = rst.Fields("previous_date").value
        Me.txtGalsCubUsed = rst.Fields("gal_cub_used").value
        Me.txtMeterSite = rst.Fields("meter_site").value
        Me.txtDeposit = Round(rst.Fields("deposit").value, 2)
        Me.txtUseCharge = Round(rst.Fields("use_charge").value, 2)
        Me.txtPastDue = Round(rst.Fields("past_due").value, 2)
        Me.txtPrevBalance = Round(rst.Fields("prev_balance").value, 2)
        Me.txtCurrentDue = Round(rst.Fields("current_due").value, 2)
        Me.txtSpecialCredit = Round(rst.Fields("special_credit").value, 2)
        Me.txtTotalDue = Round(rst.Fields("total_due").value, 2)
        Me.txtSpecialCharge = Round(rst.Fields("special_charge").value, 2)
        Me.txtSpecialDescr = rst.Fields("special_description").value
        Me.txtPhysicalAddress = rst.Fields("phy_address").value
        Me.txtName = rst.Fields("name").value
        Me.txtBillName = rst.Fields("bill_name").value
        Me.txtCareOfName = rst.Fields("care_of").value
        Me.txtAddress = rst.Fields("addr1").value
        Me.txtCity = rst.Fields("city").value
        Me.txtState = rst.Fields("state").value
        Me.txtZip = rst.Fields("zip").value
        
        'get these two values from the phones table
        qry = "SELECT * from phones WHERE CustomerID = " & account
        rstPhones.Open qry, CurrentProject.Connection
        
        If rstPhones.BOF And rstPhones.EOF Then
            'no records exist
            Me.txtWorkPhone = ""
            Me.txtHomePhone = ""
            Me.txtMobilePhone = ""
        Else
            Me.txtWorkPhone = IIf(IsNull(rstPhones.Fields("phone1").value), "", rstPhones.Fields("phone1").value)
            Me.txtHomePhone = IIf(IsNull(rstPhones.Fields("phone2").value), "", rstPhones.Fields("phone2").value)
            Me.txtMobilePhone = IIf(IsNull(rstPhones.Fields("phone3").value), "", rstPhones.Fields("phone3").value)
        End If
        rstPhones.Close
    rst.MoveNext
Loop

    'Call SetValues
    rst.Close
    
    'get the route number if assigned
    qry = "SELECT Routes.route_name" & _
            " FROM Routes INNER JOIN (customer INNER JOIN CustomerRoutes ON customer.account = " & _
            " CustomerRoutes.account_num) ON Routes.route_id = CustomerRoutes.route_id" & _
            " WHERE (((customer.account)=" & CLng(Me.txtAccount) & "));"
    
    rst.Open qry, CurrentProject.Connection
    
    If rst.BOF And rst.EOF Then
        Me.txtRoute = ""
    Else
        Me.txtRoute = IIf(IsNull(rst.Fields(0).value), "", rst.Fields(0).value)
    End If

   On Error GoTo 0
   Exit Sub

cmdNext_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdNext_Click of VBA Document Form_Opt 4 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdNext_Click of VBA Document Form_Opt 4 Form"
   
End Sub

Private Sub cmdPayments_Click()
   On Error GoTo cmdPayments_Click_Error

If Me.txtAccount <> "" And IsNumeric(Me.txtAccount) Then
    sCallingForm = Me.Name
    DoCmd.OpenReport "rptAccountPayment", acViewReport, , , , Me.txtAccount
    DoCmd.Close acForm, Me.Name, acSaveYes
Else
    Call MsgBox("The account number field must contain a numeric value.", vbInformation + vbOKOnly, "Alert")
    'DoCmd.OpenForm "DPM Main Menu", acNormal
    'DoCmd.Close acForm, Me.name, acSaveYes
End If

   On Error GoTo 0
   Exit Sub

cmdPayments_Click_Error:
    Call LogError(Err.Number, Err.source, Err.Description & " in procedure cmdPayments_Click of VBA Document Form_Opt 4 Form")
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdPayments_Click of VBA Document Form_Opt 4 Form"
End Sub

Private Sub cmdPrev_Click()
Dim account As Long
Dim qry As String
Dim rst As New ADODB.Recordset
Dim rstPhones As New ADODB.Recordset

   On Error GoTo cmdPrev_Click_Error
imgRedFlag.Visible = False
account = GetPrevRecord(CLng(Me.txtAccount))
qry = "select * from customer where account = " & account
rst.Open qry, CurrentProject.Connection

If rst.BOF And rst.EOF Then
    'otherwise do nothing
    Call MsgBox("The beginning  of the file was encountered. No further records to display.", vbExclamation, "Beginning of File")
    Exit Sub
End If

Do While Not rst.EOF
        Me.txtAccount = rst.Fields("account").value
        Me.txtGroup = rst.Fields("group").value
        Me.txtMastPar = rst.Fields("mastpar").value
        Me.txtCycle = rst.Fields("cycle").value
        Me.txtMfgCode = rst.Fields("mfg_code").value
        Me.txtStartDate = rst.Fields("start_date").value
        Me.txtStatus = rst.Fields("status").value
        Me.txtMeterNum = rst.Fields("meter_number").value
        Me.txtTermDate = rst.Fields("term_date").value
        Me.chkOutTown = rst.Fields("out_town").value
        Me.txtMeterSize = rst.Fields("meter_size").value
        Me.txtPropertyUse = rst.Fields("property_use").value
        Me.chkBackflow = rst.Fields("backflow").value
        Me.chkService = rst.Fields("service_discon").value
        If Not IsNull(Me.chkService) Then
            If Me.chkService = True Then
                imgRedFlag.Visible = True
            End If
        End If
        Me.txtFireSize = rst.Fields("fire_size").value
        Me.txtUnitofMeasure = rst.Fields("unit_measure").value
        Me.txtCurrentRead = rst.Fields("current_read").value
        Me.txtCurrentDate = rst.Fields("current_date").value
        Me.txtRateCode = rst.Fields("rate_code").value
        Me.txtPreviousRead = rst.Fields("previous_read").value
        Me.txtPreviousDate = rst.Fields("previous_date").value
        Me.txtGalsCubUsed = rst.Fields("gal_cub_used").value
        Me.txtMeterSite = rst.Fields("meter_site").value
        Me.txtDeposit = Round(rst.Fields("deposit").value, 2)
        Me.txtUseCharge = Round(rst.Fields("use_charge").value, 2)
        Me.txtPastDue = Round(rst.Fields("past_due").value, 2)
        Me.txtPrevBalance = Round(rst.Fields("prev_balance").value, 2)
        Me.txtCurrentDue = Round(rst.Fields("current_due").value, 2)
        Me.txtSpecialCredit = Round(rst.Fields("special_credit").value, 2)
        Me.txtTotalDue = Round(rst.Fields("total_due").value, 2)
        Me.txtSpecialCharge = Round(rst.Fields("special_charge").value, 2)
        Me.txtSpecialDescr = rst.Fields("special_description").value
        Me.txtPhysicalAddress = rst.Fields("phy_address").value
        Me.txtName = rst.Fields("name").value
        Me.txtBillName = rst.Fields("bill_name").value
        Me.txtCareOfName = rst.Fields("care_of").value
        Me.txtAddress = rst.Fields("addr1").value
        Me.txtCity = rst.Fields("city").value
        Me.txtState = rst.Fields("state").value
        Me.txtZip = rst.Fields("zip").value
        
        'get these two values from the phones table
        qry = "SELECT * from phones WHERE CustomerID = " & account
        rstPhones.Open qry, CurrentProject.Connection
        
        If rstPhones.BOF And rstPhones.EOF Then
            'no records exist
            Me.txtWorkPhone = ""
            Me.txtHomePhone = ""
            Me.txtMobilePhone = ""
        Else
            Me.txtWorkPhone = IIf(IsNull(rstPhones.Fields("phone1").value), "", rstPhones.Fields("phone1").value)
            Me.txtHomePhone = IIf(IsNull(rstPhones.Fields("phone2").value), "", rstPhones.Fields("phone2").value)
            Me.txtMobilePhone = IIf(IsNull(rstPhones.Fields("phone3").value), "", rstPhones.Fields("phone3").value)
        End If
        rstPhones.Close
    rst.MoveNext
Loop

    'Call SetValues
    rst.Close
    
    'get the route number if assigned
    qry = "SELECT Routes.route_name" & _
            " FROM Routes INNER JOIN (customer INNER JOIN CustomerRoutes ON customer.account = " & _
            " CustomerRoutes.account_num) ON Routes.route_id = CustomerRoutes.route_id" & _
            " WHERE (((customer.account)=" & CLng(Me.txtAccount) & "));"
    
    rst.Open qry, CurrentProject.Connection
    
    If rst.BOF And rst.EOF Then
        Me.txtRoute = ""
    Else
        Me.txtRoute = IIf(IsNull(rst.Fields(0).value), "", rst.Fields(0).value)
    End If

   On Error GoTo 0
   Exit Sub

cmdPrev_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrev_Click of VBA Document Form_Opt 4 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrev_Click of VBA Document Form_Opt 4 Form"
 
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim strCharacter As String

    ' Convert ANSI value to character string.
   On Error GoTo Form_KeyPress_Error

    strCharacter = Chr(KeyAscii)
    ' Convert character to upper case, then to ANSI value.
    KeyAscii = Asc(UCase(strCharacter))
    If KeyAscii = 27 Then
        Call ResetValues
        Me.txtAccount = ""
        Me.txtAccount.SetFocus
    End If

   On Error GoTo 0
   Exit Sub

Form_KeyPress_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_KeyPress of VBA Document Form_Opt 4 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_KeyPress of VBA Document Form_Opt 4 Form"

End Sub

Private Sub ResetValues()
        'Reset all the fields
  
   On Error GoTo ResetValues_Error

        Me.txtGroup = ""
            Me.txtMastPar = ""
            Me.txtCycle = 0
            Me.txtMfgCode = ""
            Me.txtStartDate = ""
            Me.txtStatus = ""
            Me.txtMeterNum = ""
            Me.txtTermDate = ""
            Me.txtAccount = ""
            Me.chkOutTown = False
            Me.txtMeterSize = ""
            Me.txtPropertyUse = ""
            Me.chkBackflow = False
            Me.chkService = False
            Me.txtFireSize = ""
            Me.txtUnitofMeasure = ""
            Me.txtCurrentRead = 0
            Me.txtCurrentDate = ""
            Me.txtRateCode = ""
            Me.txtPreviousRead = 0
            Me.txtGalsCubUsed = ""
            Me.txtMeterSize = ""
            Me.txtDeposit = ""
            Me.txtUseCharge = ""
            Me.txtPastDue = ""
            Me.txtPrevBalance = ""
            Me.txtCurrentDue = ""
            Me.txtSpecialCredit = ""
            Me.txtTotalDue = ""
            Me.txtSpecialCharge = ""
            Me.txtSpecialDescr = ""
            Me.txtPhysicalAddress = ""
            'No lien field?
            Me.txtName = ""
            Me.txtCareOfName = ""
            Me.txtAddress = ""
            Me.txtBillName = ""
            Me.txtCareOfName = ""
            Me.txtCity = ""
            Me.txtState = ""
            Me.txtZip = ""
            'no comment field?
            Me.txtRoute = ""
            Me.txtWorkPhone = ""
            Me.txtHomePhone = ""
            Me.txtMobilePhone = ""
            
   On Error GoTo 0
   Exit Sub

ResetValues_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure ResetValues of VBA Document Form_Opt 4 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure ResetValues of VBA Document Form_Opt 4 Form"
  
End Sub

Private Sub txtZip_AfterUpdate()
    Call FormatZip(Me.txtZip)
End Sub
