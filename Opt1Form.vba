Option Compare Database
Option Explicit
Dim account As Variant
Dim group As Variant
Dim mastpar As Variant
Dim iCycle As Variant
Dim Mfg_Code As Variant
Dim start_date As Variant
Dim Status As Variant
Dim meter_number As Variant
Dim term_date As Variant
Dim Out_Town As Variant
Dim Meter_Size As Variant
Dim Property_Use As Variant
Dim Backflow As Variant
Dim Service As Variant
Dim Fire_Size As Variant
Dim unit_measure As Variant
Dim Current_Read As Variant
Dim current_date As Variant
Dim Rate_code As Variant
Dim Previous_Read As Variant
Dim previous_date As Variant
Dim gal_cub_used As Variant
Dim meter_site As Variant
Dim Deposit As Variant
Dim Use_Charge As Variant
Dim Past_Due As Variant
Dim Prev_Balance As Variant
Dim Current_Due As Variant
Dim Special_Credit As Variant
Dim Total_Due As Variant
Dim Special_Charge As Variant
Dim special_description As Variant
Dim phy_address As Variant
Dim lien As Variant
Dim sName As Variant
Dim bill_name As Variant
Dim addr1 As Variant
Dim care_of As Variant
Dim city As Variant
Dim state As Variant
Dim Zip As Variant
Dim comment As Variant
Dim trans_loc As Variant
Dim phone_work As Variant
Dim phone_home As Variant
Dim phone_cell As Variant
Dim fDirty As Boolean

Private Enum DataType
    Intg
    Dble
    Sngl
    Strn
End Enum

Private Sub chkService_Click()
    If Me.chkService = True Then
        imgRedFlag.Visible = True
    ElseIf Me.chkService = False Then
        imgRedFlag.Visible = False
    Else
        imgRedFlag.Visible = False
    End If
End Sub

Private Sub cmdAdd_Click()
    'take all the elements and insert them into the database
    Dim curDate As Date
    Dim prevDate As Date
    Dim TermDate As Date
    Dim StartDate As Date

    'Validate certain fields
    Dim F As Boolean
   On Error GoTo cmdAdd_Click_Error

    F = ValidateNumber(Me.txtAccount)
        If Not F Then
            MsgBox Me.txtAccount & " is not a number.", vbOKOnly + vbExclamation, "Error"
            Exit Sub
        End If
    

    F = ValidateNumber(Me.txtCycle)
        If Not F Then
            MsgBox "Cycle " & Me.txtCycle & " is not a number.", vbOKOnly + vbExclamation, "Error"
            Exit Sub
        End If
    
    If Me.txtPhysicalAddress = "" Then
        MsgBox "The Physical Address cannot be blank", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    Else
        Me.txtPhysicalAddress = CheckString(Me.txtPhysicalAddress)
    End If

        If IsNull(Me.txtCurrentDate) Or Me.txtCurrentDate = "" Then
            curDate = CDate("1/1/1900")
        Else
            curDate = CDate(Me.txtCurrentDate)
        End If
        If IsNull(Me.txtPreviousDate) Or Me.txtPreviousDate = "" Then
            prevDate = CDate("1/1/1900")
        Else
            prevDate = CDate(Me.txtPreviousDate)
        End If
        If IsNull(Me.txtStartDate) Or Me.txtStartDate = "" Then
            StartDate = CDate("1/1/1900")
        Else
            StartDate = CDate(Me.txtStartDate)
        End If
        If IsNull(Me.txtTermDate) Or Me.txtTermDate = "" Then
            TermDate = CDate("1/1/1900")
        Else
            TermDate = CDate(Me.txtTermDate)
        End If

    If IsNull(Me.txtGroup) Then
        Me.txtGroup = ""
    End If
    
    If IsNull(Me.txtMastPar) Then
        Me.txtMastPar = ""
    End If
        
    If IsNull(Me.txtMfgCode) Then
        Me.txtMfgCode = ""
        'MsgBox "The MFG Code field cannot be blank", vbOKOnly + vbExclamation, "Error"
        'Exit Sub
    End If
    
    If IsNull(Me.txtStatus) Then
        Me.txtStatus = ""
    End If
    
    If IsNull(Me.txtMeterNum) Then
        Me.txtMeterNum = 0
    End If
    
    If IsNull(Me.chkOutTown) Then
        Me.chkOutTown = False
    End If
    
    If IsNull(Me.txtMeterSize) Then
        Me.txtMeterSize = 0
    End If
    
    If IsNull(Property_Use = Me.txtPropertyUse) Then
        Me.txtPropertyUse = ""
    End If
    
    If IsNull(Me.chkBackflow) Then
        Me.chkBackflow = False
    End If
    
    If IsNull(Me.chkService) Then
        Me.chkService = False
    End If
    
    If IsNull(Me.txtFireSize) Then
        Me.txtFireSize = 0
    End If
    
    If IsNull(Me.txtUnitofMeasure) Then
        Me.txtUnitofMeasure = ""
    End If
    
    If IsNull(Me.txtCurrentRead) Then
        Me.txtCurrentRead = 0
    End If
    
    If IsNull(Me.txtRateCode) Then
        MsgBox "The Rate Code cannot be blank", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
    
    If IsNull(Me.txtPreviousRead) Then
        Me.txtPreviousRead = 0
    End If
    
    If IsNull(Me.txtGalsCubUsed) Then
        Me.txtGalsCubUsed = 0
    End If
    
    If IsNull(Me.txtMeterSite) Then
        Me.txtMeterSite = ""
    End If
    
    If IsNull(Me.txtDeposit) Then
        Me.txtDeposit = 0
    End If
    
    If IsNull(Me.txtUseCharge) Then
        Me.txtUseCharge = 0
    End If
    
    If IsNull(Me.txtPastDue) Then
        Me.txtPastDue = 0
    End If
    
    If IsNull(Me.txtPrevBalance) Then
        Me.txtPrevBalance = 0
    End If
    
    If IsNull(Me.txtCurrentDue) Then
        Me.txtCurrentDue = 0
    End If
    
    If IsNull(Me.txtSpecialCredit) Then
        Me.txtSpecialCredit = 0
    End If
    
    If IsNull(Me.txtTotalDue) Then
        Me.txtTotalDue = 0
    End If
    
    If IsNull(Me.txtSpecialCharge) Then
        Me.txtSpecialCharge = 0
    End If
    
    If IsNull(Me.txtSpecialDescr) Then
        Me.txtSpecialDescr = ""
    Else
        Me.txtSpecialDescr = CheckString(Me.txtSpecialDescr)
    End If
        
    If IsNull(Me.txtName) Then
        MsgBox "The Name field cannot be blank", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    Else
        Me.txtName = CheckString(Me.txtName)
    End If
    
    If IsNull(Me.txtBillName) Then
        Me.txtBillName = Me.txtName
    Else
        Me.txtBillName = CheckString(Me.txtBillName)
    End If
    
    If IsNull(Me.txtAddress) Then
        Me.txtAddress = Me.txtPhysicalAddress
    Else
        Me.txtAddress = CheckString(Me.txtAddress)
    End If
    
    If IsNull(Me.txtCareOfName) Then
        Me.txtCareOfName = Me.txtName
    Else
        Me.txtCareOfName = CheckString(Me.txtCareOfName)
    End If
    
    If IsNull(Me.txtCity) Then
        MsgBox "The City field cannot be blank", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    Else
        Me.txtCity = CheckString(Me.txtCity)
    End If
    
    If IsNull(Me.txtState) Then
        Me.txtState = "CA"
    End If
    
    If IsNull(Me.txtZip) Then
        MsgBox "The ZIP Code field cannot be blank", vbOKOnly + vbExclamation, "Error"
        Exit Sub
    End If
    
    If IsNull(Me.txtWorkPhone) Then
        Me.txtWorkPhone = ""
    End If
    
    If IsNull(Me.txtHomePhone) Then
        Me.txtHomePhone = ""
    End If
    
    If IsNull(Me.txtMobilePhone) Then
        Me.txtMobilePhone = ""
    End If

    Dim qry As String
    Dim sName As String
    Dim sCofName As String
    
    sName = Replace(txtName, "'", "''")
    sCofName = Replace(txtCareOfName, "'", "''")
    
    qry = "insert into customer([account], [group], [mastpar], [cycle], [mfg_code], [start_date], [status], [meter_number], [term_date], " & _
        "[out_town], [meter_size], [property_use], [backflow], [service_discon], [fire_size], [unit_measure], [current_read]," & _
        "[current_date], [rate_code], [previous_read], [previous_date], [gal_cub_used], [meter_site], [deposit]," & _
        "[use_charge], [past_due], [prev_balance], [current_due], [special_credit], [total_due], [special_charge]," & _
        "[special_description], [phy_address], [lien], [name], [bill_name], [addr1], [care_of], [city], [state]," & _
        "[zip], [comment], [trans_loc]) VALUES(" & _
        CLng(Me.txtAccount) & ",'" & Me.txtGroup & "','" & Me.txtMastPar & "'," & CInt(Me.txtCycle) & ",'" & Me.txtMfgCode & "',#" & _
        StartDate & "#,'" & Me.txtStatus & "','" & Me.txtMeterNum & "'" & ",#" & TermDate & "#," & Me.chkOutTown & "," & _
        CDbl(Me.txtMeterSize) & ",'" & Me.txtPropertyUse & "'," & Me.chkBackflow & "," & Me.chkService & "," & CSng(Me.txtFireSize) & ",'" & _
        Me.txtUnitofMeasure & "'," & CDbl(Me.txtCurrentRead) & ",#" & curDate & "#,'" & Me.txtRateCode & "'," & _
        CDbl(Me.txtPreviousRead) & ",#" & prevDate & "#," & CDbl(Me.txtGalsCubUsed) & ",'" & Me.txtMeterSite & "'," & _
        CCur(Me.txtDeposit) & "," & CCur(Me.txtUseCharge) & "," & CCur(Me.txtPastDue) & "," & CCur(Me.txtPrevBalance) & "," & _
        CCur(Me.txtCurrentDue) & "," & CCur(Me.txtSpecialCredit) & "," & CCur(Me.txtTotalDue) & "," & CCur(Me.txtSpecialCharge) & ",'" & _
        Me.txtSpecialDescr & "','" & Me.txtPhysicalAddress & "','','" & sName & "','" & Me.txtBillName & "','" & _
        Me.txtAddress & "','" & sCofName & "','" & Me.txtCity & "','" & Me.txtState & "','" & Me.txtZip & "','',0)"
        
        CurrentProject.Connection.Execute qry
        
        Call UpdatePhones
        Call SetValues
        Dim answer As VbMsgBoxResult
        
        answer = MsgBox("Record Successfully added. Add another one?", vbYesNo + vbExclamation, "Error")
        If answer = vbYes Then
            Dim intRec As String
            intRec = InputBox("Please enter an Account number to maintain", "Add Account")
            If intRec = "" Then
                Exit Sub
            End If
            Call ResetFields
            Me.txtAccount = intRec
        End If
        
    On Error GoTo 0
    Exit Sub
cmdAdd_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description
    'log the error
    Dim bool As Boolean
    Dim msg As String
    If InStr(Err.Description, "would create duplicate values") Then
        MsgBox "Unable to create the account as it would create duplicate values."
        Exit Sub
    End If
    bool = LogError(Err.Number, "procedure cmdAdd of Form Opt1Form", Err.Description)
    If bool Then
        msg = "was logged."
    Else
        msg = "was NOT logged."
    End If
    
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdAdd_Click of VBA Document Form_Opt1Form" & msg

End Sub

Private Sub cmdFind_Click()
    Call Form_Load
End Sub

Private Sub cmdHistory_Click()

   On Error GoTo cmdHistory_Click_Error
    If IsNull(Me.txtAccount) Then
        Exit Sub
    End If
   DoCmd.OpenForm "frmHistory", acNormal, , "account = " & Me.txtAccount
   On Error GoTo 0
   Exit Sub

cmdHistory_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdHistory_Click of VBA Document Form_Opt1Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdHistory_Click of VBA Document Form_Opt1Form"
End Sub

Private Sub cmdNext_Click()

Dim account As Long
Dim qry As String
Dim rst As New ADODB.Recordset
Dim rstPhones As New ADODB.Recordset
   On Error GoTo cmdNext_Click_Error
If IsNull(txtAccount) Then
    Exit Sub
End If

imgRedFlag.Visible = False

account = GetNextRecord(CLng(Me.txtAccount))
qry = "select * from customer where account = " & account
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
        Me.txtAddress = rst.Fields("addr1").value
        Me.txtCareOfName = rst.Fields("care_of").value
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

    Call SetValues
    rst.Close

   On Error GoTo 0
   Exit Sub

cmdNext_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdNext_Click of VBA Document Form_Opt1Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdNext_Click of VBA Document Form_Opt1Form"

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
    
    DoCmd.OpenForm "frmNotes", acNormal, , , acFormPropertySettings, acWindowNormal, Me.txtAccount
    On Error GoTo 0
    Exit Sub

   On Error GoTo 0
   Exit Sub

cmdNotes_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdNotes_Click of VBA Document Form_Opt1Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdNotes_Click of VBA Document Form_Opt1Form"

End Sub

Private Sub cmdPayments_Click()
   
   On Error GoTo cmdPayments_Click_Error

If Me.txtAccount <> "" And IsNumeric(Me.txtAccount) Then
    sCallingForm = Me.Name
    DoCmd.OpenReport "rptAccountPayment", acViewReport, , , , Me.txtAccount
    DoCmd.Close acForm, Me.Name, acSaveYes
Else
    Call MsgBox("The account field must contain a numeric value.", vbInformation + vbOKOnly, "Alert")
    DoCmd.OpenForm "DPM Main Menu", acNormal
    'DoCmd.Close acForm, Me.name, acSaveYes
End If

   On Error GoTo 0
   Exit Sub

cmdPayments_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPayments_Click of VBA Document Form_Opt1Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPayments_Click of VBA Document Form_Opt1Form"
End Sub

Private Sub cmdPrev_Click()
Dim account As Long
Dim qry As String
Dim rst As New ADODB.Recordset
Dim rstPhones As New ADODB.Recordset
   On Error GoTo cmdPrev_Click_Error
imgRedFlag.Visible = False
If IsNull(Me.txtAccount) Then
    Exit Sub
End If
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
        Me.txtAddress = rst.Fields("addr1").value
        Me.txtCareOfName = rst.Fields("care_of").value
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

    Call SetValues
    rst.Close

   On Error GoTo 0
   Exit Sub

cmdPrev_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrev_Click of VBA Document Form_Opt1Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrev_Click of VBA Document Form_Opt1Form"

End Sub

Private Sub Form_Load()
'Get the record to go to
Dim strRecord As String
Dim strPrompt As String
Dim strFilter As String
Dim rst As New ADODB.Recordset
Dim rstPhones As New ADODB.Recordset
Dim qry As String
    
   On Error GoTo Form_Load_Error

vOpenArgs = Me.OpenArgs
imgRedFlag.Visible = False
If IsNumeric(vOpenArgs) Then
    'open in add mode
    Me.txtAccount.SetFocus
    If Me.Name = sCallingForm Then
        Me.lblHeader.Caption = "Account Maintenance"
    Else
        Me.lblHeader.Caption = "Add new Account"
    End If
    Me.txtAccount = vOpenArgs
    strRecord = vOpenArgs
    strFilter = "account"
    qry = "select * from customer where " & strFilter & " = " & strRecord
        
End If

If vOpenArgs = "mastpar" Then
    strPrompt = "Parcel"
    strFilter = "mastpar"
    strRecord = InputBox("Please enter a " & strPrompt & " number to Maintain", strPrompt & "Maintenance")
    If strRecord = "" Then
        DoCmd.Close acForm, Me.Name, acSaveYes
        Exit Sub
    End If
    Me.txtMastPar.SetFocus
    Me.lblHeader.Caption = "Master Parcel Maintenance"
    qry = "select * from customer where " & strFilter & " = '" & strRecord & "'"
End If

If vOpenArgs = "account" Then
    strPrompt = "Account"
    strFilter = "account"
    strRecord = InputBox("Please enter a " & strPrompt & " number to Maintain", strPrompt & "Maintenance")
    If strRecord = "" Then
        DoCmd.Close acForm, Me.Name, acSaveYes
        Exit Sub
    End If
    Me.txtAccount.SetFocus
    Me.lblHeader.Caption = "Account Maintenance"
    qry = "select * from customer where " & strFilter & " = " & strRecord
End If

If InStr(vOpenArgs, "existing") Then
    'split the string at the comm
    Dim arr() As String
    arr = Split(vOpenArgs, ",")
    strRecord = arr(1)
    strFilter = "account"
    Me.txtAccount.SetFocus
    Me.lblHeader.Caption = "Account Maintenance"
    qry = "select * from customer where " & strFilter & " = " & strRecord
End If

If Me.lblHeader.Caption = "Text" Then
    Me.lblHeader.Caption = "Maintenance Form"
End If

If Not IsNumeric(strRecord) And strFilter = "account" Then
    MsgBox "The " & strPrompt & " you entered was not a valid number. Please try again", vbCritical + vbOKOnly, "Error"
    DoCmd.OpenForm "DPM Main Menu", acNormal
    DoCmd.Close acForm, Me.Name, acSaveYes
    Exit Sub
End If

'We don't have a set query so exit
If qry = "" Then
    Exit Sub
End If

rst.Open qry, CurrentProject.Connection

If rst.BOF And rst.EOF Then
    If fHaveAsked Then
        'reset the have asked boolean
        If Me.cmdAdd.Visible = False Then
            Me.cmdAdd.Visible = True
        End If
        fHaveAsked = False
    Else
    'alert user that no record was found and ask him he they want to add it
    Dim result As VbMsgBoxResult
    result = MsgBox(strFilter & " " & strRecord & " could not be found. " & _
        "Do you want to create this now", vbYesNo + vbExclamation, "Add Record?")
        If result = vbYes Then
            'call add routine
            Call AddNewRecord
            If vOpenArgs = "account" Then
                Me.txtAccount = strRecord
            ElseIf vOpenArgs = "mastpar" Then
                Me.txtMastPar = strRecord
            Else
                'not sure where to put strRecord
                Stop
            End If
            Exit Sub
        Else
            'otherwise do not load the form.
            'DoCmd.OpenForm "DPM Main Menu", acNormal
            DoCmd.Close acForm, Me.Name, acSaveYes
            Exit Sub
        End If
    End If
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
        Me.txtAddress = rst.Fields("addr1").value
        Me.txtCareOfName = rst.Fields("care_of").value
        Me.txtCity = rst.Fields("city").value
        Me.txtState = rst.Fields("state").value
        Me.txtZip = rst.Fields("zip").value
        
        'get these two values from the phones table
        qry = "SELECT * from phones WHERE CustomerID = " & Me.txtAccount
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

    Call SetValues
    rst.Close
   On Error GoTo 0
   Exit Sub

Form_Load_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description
    'log the error
    Dim F As Boolean
    Dim msg As String
    F = LogError(Err.Number, "procedure Load of Form Opt1Form", Err.Description)
    If F Then
        msg = "was logged."
    Else
        msg = "was NOT logged."
    End If
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Load of VBA Document Form_Opt1Form" & msg

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim strCharacter As String

    ' Convert ANSI value to character string.
   On Error GoTo Form_KeyPress_Error

    strCharacter = Chr(KeyAscii)
    ' Convert character to upper case, then to ANSI value.
    KeyAscii = Asc(UCase(strCharacter))
    If KeyAscii = 27 Then
        'close the form
        DoCmd.Close acForm, Me.Name, acSaveYes
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

    Call LogError(errNum, errSource, errMsg & " in procedure Form_KeyPress of VBA Document Form_Opt1Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_KeyPress of VBA Document Form_Opt1Form"

End Sub

Private Sub cmdSave_Click()
    Dim lRec As Long
   On Error GoTo cmdSave_Click_Error

    If IsNull(Me.txtAccount) Then
        Exit Sub
    End If
    'find out if the owner, case of name, bill to name or any of the address fields have changed.
    'if so, the save the old data into a history table
    If Me.txtName <> sName Or Me.txtBillName <> bill_name Or Me.txtCareOfName <> care_of _
      Or Me.txtAddress <> addr1 Or Me.txtCity <> city Or Me.txtState <> state Or Me.txtZip <> Zip Then
      RecordHistory
    End If
    
    lRec = UpdateRecord
    Call SetValues
   
   On Error GoTo 0
   Exit Sub

cmdSave_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description
    Dim F As Boolean
    Dim msg As String
    F = LogError(Err.Number, "procedure UpdateRecord of Form Opt1Form", Err.Description)
    If F Then
        msg = "Error was logged."
    Else
        msg = "Error was NOT logged."
    End If
    
    'alert user
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdSave_Click of VBA Document Form_Opt1Form. " & msg
End Sub


'------------------------------------------------------------
' cmdPrint_Click
'
'------------------------------------------------------------
Private Sub cmdPrint_Click()
    
   On Error GoTo cmdPrint_Click_Error
    If IsNull(Me.txtAccount) Then
        Exit Sub
    End If
    
    DoCmd.RunCommand acCmdPrint


   On Error GoTo 0
   Exit Sub

cmdPrint_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrint_Click of VBA Document Form_Opt1Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrint_Click of VBA Document Form_Opt1Form"

End Sub


'------------------------------------------------------------
' cmdExit_Click
'
'------------------------------------------------------------
Private Sub cmdExit_Click()
   On Error GoTo cmdExit_Click_Error

    DoCmd.Close , ""

   On Error GoTo 0
   Exit Sub

cmdExit_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdExit_Click of VBA Document Form_Opt1Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdExit_Click of VBA Document Form_Opt1Form"

End Sub


Private Sub Form_Unload(Cancel As Integer)
   On Error GoTo Form_Unload_Error

fDirty = False
'check all the fields against the loaded values. if any changed, warn the user
If account <> Me.txtAccount Then
    fDirty = True
End If
If group <> Me.txtGroup Then
    fDirty = True
End If
If mastpar <> Me.txtMastPar Then
    fDirty = True
End If
If iCycle <> Me.txtCycle Then
    fDirty = True
End If
If Mfg_Code <> Me.txtMfgCode Then
    fDirty = True
End If
If start_date <> Me.txtStartDate Then
    fDirty = True
End If
If Status <> Me.txtStatus Then
    fDirty = True
End If
If meter_number <> Me.txtMeterNum Then
    fDirty = True
End If
If term_date <> Me.txtTermDate Then
    fDirty = True
End If
If Out_Town <> Me.chkOutTown Then
    fDirty = True
End If
If Meter_Size <> Me.txtMeterSize Then
    fDirty = True
End If
If Property_Use <> Me.txtPropertyUse Then
    fDirty = True
End If
If Backflow <> Me.chkBackflow Then
    fDirty = True
End If
If Service <> Me.chkService Then
    fDirty = True
End If
If Fire_Size <> Me.txtFireSize Then
    fDirty = True
End If
If unit_measure <> Me.txtUnitofMeasure Then
    fDirty = True
End If
If Current_Read <> Me.txtCurrentRead Then
    fDirty = True
End If
If current_date <> Me.txtCurrentDate Then
    fDirty = True
End If
If Rate_code <> Me.txtRateCode Then
    fDirty = True
End If
If Previous_Read <> Me.txtPreviousRead Then
    fDirty = True
End If
If previous_date <> Me.txtPreviousDate Then
    fDirty = True
End If
If gal_cub_used <> Me.txtGalsCubUsed Then
    fDirty = True
End If
If meter_site <> Me.txtMeterSite Then
    fDirty = True
End If
If Deposit <> Me.txtDeposit Then
    fDirty = True
End If
If Use_Charge <> Me.txtUseCharge Then
    fDirty = True
End If
If Past_Due <> Me.txtPastDue Then
    fDirty = True
End If
If Prev_Balance <> Me.txtPrevBalance Then
    fDirty = True
End If
If Current_Due <> Me.txtCurrentDue Then
    fDirty = True
End If
If Special_Credit <> Me.txtSpecialCredit Then
    fDirty = True
End If
If Total_Due <> Me.txtTotalDue Then
    fDirty = True
End If
If Special_Charge <> Me.txtSpecialCharge Then
    fDirty = True
End If
If special_description <> Me.txtSpecialDescr Then
    fDirty = True
End If
If phy_address <> Me.txtPhysicalAddress Then
    fDirty = True
End If
If sName <> Me.txtName Then
    fDirty = True
End If
If bill_name <> Me.txtBillName Then
    fDirty = True
End If
If addr1 <> Me.txtAddress Then
    fDirty = True
End If
If care_of <> Me.txtCareOfName Then
    fDirty = True
End If
If city <> Me.txtCity Then
    fDirty = True
End If
If state <> Me.txtState Then
    fDirty = True
End If
If Zip <> Me.txtZip Then
    fDirty = True
End If
If phone_work <> Me.txtWorkPhone Then
    fDirty = True
End If
If phone_home <> Me.txtHomePhone Then
    fDirty = True
End If
If phone_cell <> Me.txtMobilePhone Then
    fDirty = True
End If

If fDirty Then
    Dim result As VbMsgBoxResult
    result = MsgBox("The current record has not been saved. Save Now?", vbYesNo + vbExclamation, "Save?")
    
    If result = vbYes Then
        'call the update function
        Dim I As Long
        I = UpdateRecord
        If I = 0 Then
            Err.Raise vbObjectError ', "procedure Unload of Form Opt1Form", "Record was not updated"
        End If
    End If
End If

    On Error GoTo 0
    Exit Sub
    
Form_Unload_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description
    Dim F As Boolean
    Dim msg As String
    F = LogError(Err.Number, "procedure Unload of Form Opt1Form", Err.Description)
    If F Then
        msg = ". Error was logged."
    Else
        msg = ". Error was NOT logged."
    End If
    
    'alert user
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Unload of VBA Document Form_Opt1Form" & msg

End Sub

Private Sub AddNewRecord()

   On Error GoTo AddNewRecord_Error

Me.cmdAdd.Visible = True
Me.lblHeader.Caption = "Add New Record"
Me.cmdSave.Visible = False
Me.cmdPayments.Visible = False
Me.cmdPrint.Visible = False

   On Error GoTo 0
   Exit Sub

AddNewRecord_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure AddNewRecord of VBA Document Form_Opt1Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure AddNewRecord of VBA Document Form_Opt1Form"

End Sub

Private Function UpdateRecord() As Integer
    Dim curDate As Date
    Dim prevDate As Date
    Dim TermDate As Date
    Dim StartDate As Date
    Dim sHoldPostraphe As String
    Dim tmpOwner As String
    Dim tmpPhyAddress As String
    Dim tmpAddress As String
    Dim tmpBillToName As String
    Dim tmpCareofName As String
    
    
    
    'ensure all fields are correct, no nulls, valid dates, etc
   On Error GoTo UpdateRecord_Error

        If IsNull(Me.txtAccount) Then
            Me.txtAccount = 0
            Exit Function
        ElseIf Me.txtAccount = 0 Then
            'nothing to do
            Exit Function
        End If
        If IsNull(Me.txtGroup) Then
                    Me.txtGroup = ""
        End If
        If IsNull(Me.txtMastPar) Then
                    Me.txtMastPar = ""
        End If
        If IsNull(Me.txtCycle) Then
                    Me.txtCycle = 0
        End If
        If IsNull(Me.txtMfgCode) Then
                    Me.txtMfgCode = ""
        End If
        If IsNull(Me.txtStartDate) Or Me.txtStartDate = "" Then
            StartDate = CDate("1/1/1900")
        Else
            StartDate = CDate(Me.txtStartDate)
        End If
        If IsNull(Me.txtStatus) Then
                    Me.txtStatus = ""
        End If
        If IsNull(Me.txtMeterNum) Then
                    Me.txtMeterNum = 0
        End If
        If IsNull(Me.txtTermDate) Or Me.txtTermDate = "" Then
            TermDate = CDate("1/1/1900")
        Else
            TermDate = CDate(Me.txtTermDate)
        End If
        If IsNull(Me.chkOutTown) Then
                    Me.chkOutTown = False
        End If
        If IsNull(Me.txtMeterSize) Then
                    Me.txtMeterSize = 0
        End If
        If IsNull(Me.txtPropertyUse) Then
                    Me.txtPropertyUse = ""
        End If
        If IsNull(Me.chkBackflow) Then
                    Me.chkBackflow = False
        End If
        If IsNull(Me.chkService) Then
            Me.chkService = False
        End If
        If IsNull(Me.txtFireSize) Then
                    Me.txtFireSize = 0
        End If
        If IsNull(Me.txtUnitofMeasure) Then
                    Me.txtUnitofMeasure = ""
        End If
        If IsNull(Me.txtCurrentRead) Then
                    Me.txtCurrentRead = 0
        End If
        If IsNull(Me.txtCurrentDate) Or Me.txtCurrentDate = "" Then
            curDate = CDate("1/1/1900")
        Else
            curDate = CDate(Me.txtCurrentDate)
        End If
        If IsNull(Me.txtRateCode) Then
                    Me.txtRateCode = ""
        End If
        If IsNull(Me.txtPreviousRead) Then
                    Me.txtPreviousRead = 0
        End If
        If IsNull(Me.txtPreviousDate) Or Me.txtPreviousDate = "" Then
            prevDate = CDate("1/1/1900")
        Else
            prevDate = CDate(Me.txtPreviousDate)
        End If
        If IsNull(Me.txtGalsCubUsed) Then
                    Me.txtGalsCubUsed = 0
        End If
        If IsNull(Me.txtMeterSite) Then
                    Me.txtMeterSite = ""
        End If
        
        If InStr(Me.txtMeterSite, "'") > 0 Then
            sHoldPostraphe = Replace(Me.txtMeterSite, "'", "''")
        Else
            sHoldPostraphe = Me.txtMeterSite
        End If
                
        If IsNull(Me.txtDeposit) Then
                    Me.txtDeposit = 0
        End If
        If IsNull(Me.txtUseCharge) Then
                    Me.txtUseCharge = 0
        End If
        If IsNull(Me.txtPastDue) Then
                    Me.txtPastDue = 0
        End If
        If IsNull(Me.txtPrevBalance) Then
                    Me.txtPrevBalance = 0
        End If
        If IsNull(Me.txtCurrentDue) Then
                    Me.txtCurrentDue = 0
        End If
        If IsNull(Me.txtSpecialCredit) Then
                    Me.txtSpecialCredit = 0
        End If
        If IsNull(Me.txtTotalDue) Then
                    Me.txtTotalDue = 0
        End If
        If IsNull(Me.txtSpecialCharge) Then
                    Me.txtSpecialCharge = 0
        End If
        If IsNull(Me.txtSpecialDescr) Then
                    Me.txtSpecialDescr = ""
        End If
        If IsNull(Me.txtPhysicalAddress) Then
            tmpPhyAddress = ""
        Else
            If InStr(Me.txtPhysicalAddress, "'") > 0 Then
                tmpPhyAddress = Replace(Me.txtPhysicalAddress, "'", "''")
            Else
                tmpPhyAddress = Me.txtPhysicalAddress
            End If
        End If
        
        If IsNull(Me.txtName) Then
            tmpOwner = ""
        Else
            If InStr(Me.txtName, "'") > 0 Then
                tmpOwner = Replace(Me.txtName, "'", "''")
            Else
                tmpOwner = Me.txtName
            End If
        End If
        
        
        If IsNull(Me.txtBillName) Then
            tmpBillToName = ""
        Else
            If InStr(Me.txtBillName, "'") > 0 Then
                tmpBillToName = Replace(Me.txtBillName, "'", "''")
            Else
                tmpBillToName = Me.txtBillName
            End If
        End If
        
        If IsNull(Me.txtAddress) Then
            tmpAddress = ""
        Else
            If InStr(Me.txtAddress, "'") > 0 Then
                tmpAddress = Replace(Me.txtAddress, "'", "''")
            Else
                tmpAddress = Me.txtAddress
            End If
        End If
        
        If IsNull(Me.txtCareOfName) Then
            tmpCareofName = ""
        Else
            If InStr(Me.txtCareOfName, "'") > 0 Then
                tmpCareofName = Replace(Me.txtCareOfName, "'", "''")
            Else
                tmpCareofName = Me.txtCareOfName
            End If
        End If
        
        If IsNull(Me.txtCity) Then
            Me.txtCity = ""
        End If
        If IsNull(Me.txtState) Then
            Me.txtState = ""
        End If
        If IsNull(Me.txtZip) Then
            Me.txtZip = ""
        End If
        
        If IsNull(Me.txtWorkPhone) Then
            Me.txtWorkPhone = ""
        End If
        
        If IsNull(Me.txtHomePhone) Then
            Me.txtHomePhone = ""
        End If
        
        If IsNull(Me.txtMobilePhone) Then
            Me.txtMobilePhone = ""
        End If
                
    Dim l As Long
    Dim qry As String
    qry = ""
    
    qry = "UPDATE customer " & _
        "set [group] = '" & Me.txtGroup & "',[mastpar] = '" & _
        Me.txtMastPar.value & "',[cycle] = " & CInt(Me.txtCycle.value) & ",[mfg_code] = '" & Me.txtMfgCode.value & "'," & _
        "[start_date] = #" & StartDate & "#,[status] = '" & Me.txtStatus.value & "'," & _
        "[meter_number] = '" & Me.txtMeterNum.value & "',[term_date] = #" & TermDate & "#,[out_town] = " & _
        Me.chkOutTown.value & ",[meter_size] = " & CDbl(Me.txtMeterSize.value) & _
        ",[property_use] = '" & Me.txtPropertyUse.value & "',[backflow] = " & Me.chkBackflow.value & _
        ",[service_discon] = " & Me.chkService.value & _
        ",[fire_size] = " & CSng(Me.txtFireSize.value) & ",[unit_measure] = '" & _
        Me.txtUnitofMeasure.value & "',[current_read] = " & CDbl(Me.txtCurrentRead.value) & ",[current_date] = #" & curDate & _
        "#,[rate_code] = '" & Me.txtRateCode.value & "',[previous_read] = " & CDbl(Me.txtPreviousRead.value) & ",[previous_date] = #" & _
        prevDate & "#,[gal_cub_used] = " & CDbl(Me.txtGalsCubUsed.value) & ",[meter_site] = '" & sHoldPostraphe & _
        "',[deposit] = " & CCur(Me.txtDeposit.value) & ",[use_charge] = " & CCur(Me.txtUseCharge.value) & ",[past_due] = " & _
        CCur(Me.txtPastDue.value) & ",[prev_balance] = " & CCur(Me.txtPrevBalance.value) & ",[current_due] = " & _
        CCur(Me.txtCurrentDue.value) & ",[special_credit] = " & CCur(Me.txtSpecialCredit.value) & ",[total_due] = " & _
        CCur(Me.txtTotalDue.value) & ", [special_charge] = " & CCur(Me.txtSpecialCharge.value) & _
        ",[special_description] = '" & Me.txtSpecialDescr.value & "',[phy_address] = '" & tmpPhyAddress & _
        "',[name] = '" & tmpOwner & "',[bill_name] = '" & tmpBillToName & "', [addr1] = '" & tmpAddress & _
        "',[care_of] = '" & tmpCareofName & "',[city] = '" & Me.txtCity.value & "',[state] = '" & Me.txtState.value & _
        "',[zip] = '" & Me.txtZip.value & "' WHERE [account] = " & CLng(Me.txtAccount.value)
        
    CurrentProject.Connection.Execute qry, l
    'Now update the Phones Table
    Call UpdatePhones
    UpdateRecord = l
   
   On Error GoTo 0
   Exit Function

UpdateRecord_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description
    Dim F As Boolean
    Dim msg As String
    F = LogError(Err.Number, "procedure UpdateRecord of Form Opt1Form", Err.Description)
    If F Then
        msg = ". Error was logged."
    Else
        msg = ". Error was NOT logged."
    End If

    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure UpdateRecord of VBA Document Form_Opt1Form" & msg
    
        UpdateRecord = 0

End Function

Private Sub SetValues()

   On Error GoTo SetValues_Error

    account = Me.txtAccount
    group = Me.txtGroup
    mastpar = Me.txtMastPar
    iCycle = Me.txtCycle
    Mfg_Code = Me.txtMfgCode
    start_date = Me.txtStartDate
    Status = Me.txtStatus
    meter_number = Me.txtMeterNum
    term_date = Me.txtTermDate
    Out_Town = Me.chkOutTown
    Meter_Size = Me.txtMeterSize
    Property_Use = Me.txtPropertyUse
    Backflow = Me.chkBackflow
    Service = Me.chkService
    Fire_Size = Me.txtFireSize
    unit_measure = Me.txtUnitofMeasure
    Current_Read = Me.txtCurrentRead
    current_date = Me.txtCurrentDate
    Rate_code = Me.txtRateCode
    Previous_Read = Me.txtPreviousRead
    previous_date = Me.txtPreviousDate
    gal_cub_used = Me.txtGalsCubUsed
    meter_site = Me.txtMeterSite
    Deposit = Round(Me.txtDeposit, 2)
    Use_Charge = Round(Me.txtUseCharge, 2)
    Past_Due = Round(Me.txtPastDue, 2)
    Prev_Balance = Round(Me.txtPrevBalance, 2)
    Current_Due = Round(Me.txtCurrentDue, 2)
    Special_Credit = Round(Me.txtSpecialCredit, 2)
    Total_Due = Round(Me.txtTotalDue, 2)
    Special_Charge = Round(Me.txtSpecialCharge, 2)
    special_description = Me.txtSpecialDescr
    phy_address = Me.txtPhysicalAddress
    sName = Me.txtName
    bill_name = Me.txtBillName
    addr1 = Me.txtAddress
    care_of = Me.txtCareOfName
    city = Me.txtCity
    state = Me.txtState
    Zip = Me.txtZip
    phone_work = Me.txtWorkPhone
    phone_home = Me.txtHomePhone
    phone_cell = Me.txtMobilePhone

   On Error GoTo 0
   Exit Sub

SetValues_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure SetValues of VBA Document Form_Opt1Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure SetValues of VBA Document Form_Opt1Form"

End Sub

Private Sub ResetFields()
    
   On Error GoTo ResetFields_Error

    Me.txtAccount = ""
    Me.txtGroup = ""
    Me.txtMastPar = ""
    Me.txtCycle = ""
    Me.txtMfgCode = ""
    Me.txtStartDate = "1/1/1900"
    Me.txtStatus = "A"
    Me.txtMeterNum = 0
    Me.txtTermDate = "1/1/1900"
    Me.chkOutTown = False
    Me.txtMeterSize = 0
    Me.txtPropertyUse = ""
    Me.chkBackflow = False
    Me.chkService = False
    Me.txtFireSize = 0
    Me.txtUnitofMeasure = ""
    Me.txtCurrentRead = 0
    Me.txtCurrentDate = "1/1/1900"
    Me.txtRateCode = ""
    Me.txtPreviousRead = 0
    Me.txtPreviousDate = "1/1/1900"
    Me.txtGalsCubUsed = 0
    Me.txtMeterSite = ""
    Me.txtDeposit = 0
    Me.txtUseCharge = 0
    Me.txtPastDue = 0
    Me.txtPrevBalance = 0
    Me.txtCurrentDue = 0
    Me.txtSpecialCredit = 0
    Me.txtTotalDue = 0
    Me.txtSpecialCharge = 0
    Me.txtSpecialDescr = ""
    Me.txtPhysicalAddress = ""
    Me.txtName = ""
    Me.txtBillName = ""
    Me.txtAddress = ""
    Me.txtCareOfName = ""
    Me.txtCity = ""
    Me.txtState = "CA"
    Me.txtZip = ""
    Me.txtWorkPhone = ""
    Me.txtHomePhone = ""
    Me.txtMobilePhone = ""

   On Error GoTo 0
   Exit Sub

ResetFields_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure ResetFields of VBA Document Form_Opt1Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure ResetFields of VBA Document Form_Opt1Form"

End Sub

Private Function CheckString(value As String) As String

    Dim replacement As String
   On Error GoTo CheckString_Error

    If InStr(value, "'") Then
        replacement = Replace(value, "'", "''")
    Else
        replacement = value
    End If
    
    CheckString = replacement

   On Error GoTo 0
   Exit Function

CheckString_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure CheckString of VBA Document Form_Opt1Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure CheckString of VBA Document Form_Opt1Form"
    
End Function

'this routine will only save the last selected value for Me.cboPhoneType.value on exit.
'In otherwords if the user had selected work, then only the work number will be saved.
' It is up to the user to save the other values before exiting
Private Sub UpdatePhones()
    'Find out if we're adding a phone or updating an existing phone.
Dim pNum1 As String
Dim pNum2 As String
Dim pNum3 As String
Dim rst As New ADODB.Recordset
Dim query As String
Dim lRecs As Long

   On Error GoTo UpdatePhones_Error

If IsNull(Me.txtAccount) Then
Call MsgBox("No account number has been selected or enetered. Please enter a valid account number and then try again.", vbExclamation, "No Account Number")
    Exit Sub
End If

query = "Select phone1, phone2, phone3 from phones where customerid = " & Me.txtAccount
rst.Open query, CurrentProject.Connection

If rst.BOF And rst.EOF Then
    'no records could be found
    Call AddNewPhone
    Exit Sub
End If
    
'Find out if a number was changed, or if a new phone type and number were added

        pNum1 = IIf(IsNull(Me.txtWorkPhone) Or Me.txtWorkPhone = "", "", Me.txtWorkPhone)
        pNum2 = IIf(IsNull(Me.txtHomePhone.value) Or Me.txtHomePhone = "", "", Me.txtHomePhone)
        pNum3 = IIf(IsNull(Me.txtMobilePhone.value) Or Me.txtMobilePhone = "", "", Me.txtMobilePhone)
        query = "UPDATE Phones set phone1 = '" & pNum1 & "', phone2 = '" & pNum2 & "', phone3 = '" & pNum3 & _
                "' WHERE customerid = " & Me.txtAccount.value
        
        CurrentProject.Connection.Execute query, lRecs

If lRecs <> 1 Then
    'log the error
    Err.Raise vbObjectError + 3113, "UpdatePhones of VBA Document Form_Opt1Form", _
    "The following update query failed to update the records, but no error was returned by the database: " & query
End If


   On Error GoTo 0
   Exit Sub

UpdatePhones_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure UpdatePhones of VBA Document Form_Opt1Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure UpdatePhones of VBA Document Form_Opt1Form"

End Sub

Private Sub AddNewPhone()
   'Find out if any of the phone fields have been taken
   Dim qry As String
   Dim upDateqry As String
   Dim rst As New ADODB.Recordset
   Dim pNum1 As String
   Dim pNum2 As String
   Dim pNum3 As String
   Dim lRecs As Long
   
   On Error GoTo AddNewPhone_Error

   qry = "select phone1, phone2, phone3 from phones where customerid = " & Me.txtAccount.value
   rst.Open qry, CurrentProject.Connection
   
    If rst.BOF And rst.EOF Then
        'were phone numbers added
        If Not IsNull(Me.txtWorkPhone) Or Not IsNull(Me.txtHomePhone) Or Not IsNull(Me.txtMobilePhone) Or Me.txtWorkPhone <> "" Or Me.txtHomePhone <> "" Or Me.txtMobilePhone <> "" Then
            pNum1 = IIf(IsNull(Me.txtWorkPhone) Or Me.txtWorkPhone = "", "", Me.txtWorkPhone)
            pNum2 = IIf(IsNull(Me.txtHomePhone.value) Or Me.txtHomePhone = "", "", Me.txtHomePhone)
            pNum3 = IIf(IsNull(Me.txtMobilePhone.value) Or Me.txtMobilePhone = "", "", Me.txtMobilePhone)
            upDateqry = "INSERT INTO Phones (CustomerID, phone1, phone2, phone3) VALUES(" & Me.txtAccount & ",'" & pNum1 & "','" & pNum2 & "','" & pNum3 & _
                    "')"
            CurrentProject.Connection.Execute upDateqry, lRecs
        End If
    Else
        'simply update what is there
        pNum1 = IIf(IsNull(Me.txtWorkPhone) Or Me.txtWorkPhone = "", "", Me.txtWorkPhone)
        pNum2 = IIf(IsNull(Me.txtHomePhone.value) Or Me.txtHomePhone = "", "", Me.txtHomePhone)
        pNum3 = IIf(IsNull(Me.txtMobilePhone.value) Or Me.txtMobilePhone = "", "", Me.txtMobilePhone)
        upDateqry = "UPDATE Phones set Phone1 = '" & pNum1 & "', Phone2 = '" & pNum2 & " Phone3 = " & pNum3 & _
            "' WHERE CustomerID = " & Me.txtAccount.value
        CurrentProject.Connection.Execute upDateqry
    End If
    
    If lRecs <> 1 Then
        'log the error
        Err.Raise vbObjectError + 3113, "UpdatePhones of VBA Document Form_Opt1Form", _
        "The following update query failed to update the records, but no error was returned by the database: " & query
    End If

   On Error GoTo 0
   Exit Sub

AddNewPhone_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure AddNewPhone of VBA Document Form_Opt1Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure AddNewPhone of VBA Document Form_Opt1Form"

End Sub

Private Sub RecordHistory()

   On Error GoTo RecordHistory_Error

    account = Me.txtAccount
    sName = Replace(Me.txtName, "'", "''")
    bill_name = Replace(Me.txtBillName, "'", "''")
    addr1 = Replace(Me.txtAddress, "'", "''")
    care_of = Replace(Me.txtCareOfName, "'", "''")
    city = Me.txtCity
    state = Me.txtState
    Zip = Me.txtZip
    'txtMeterSite = Replace(txtMeterSite, "'", "''")

Dim qry As String
Dim lRecs As Long
qry = "insert into History (ChangeDateTime, Account, Owner, CareOfName, BillToName, Addr1, City, State, ZIP) " & _
      " VALUES(#" & Now & "#," & account & ",'" & sName & "','" & care_of & "','" & bill_name & "','" & addr1 & "','" & _
      city & "','" & state & "','" & Zip & "')"
CurrentProject.Connection.Execute qry, lRecs

If lRecs = 0 Then
    'an error occurred
    Err.Raise vbObjectError + 3003, "Opt1Form.RecordHistory", "No records were inserted into the History Table for " & qry
End If

   On Error GoTo 0
   Exit Sub

RecordHistory_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure RecordHistory of VBA Document Form_Opt1Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure RecordHistory of VBA Document Form_Opt1Form"

End Sub

Private Sub txtZip_AfterUpdate()
    Call FormatZip(Me.txtZip)
End Sub

Private Sub LoadRateValues()
Dim rst As New ADODB.Recordset
Dim query As String
Dim SysMaint As String
Dim rCode As String
   On Error GoTo LoadRateValues_Error

query = "SELECT customer.account, CustomerServiceConnection.service_id, RatesAndCharges.recurring_charge_id, " & _
        "RecurringCharges.charge_code, ServiceConnections.service_id FROM (RecurringCharges INNER JOIN " & _
        "((customer INNER JOIN CustomerServiceConnection ON customer.account = CustomerServiceConnection.account) " & _
        "INNER JOIN RatesAndCharges ON customer.account = RatesAndCharges.account) ON RecurringCharges.charge_id = " & _
        "RatesAndCharges.recurring_charge_id) INNER JOIN ServiceConnections ON CustomerServiceConnection.service_id = " & _
        "ServiceConnections.service_id WHERE (((customer.account)=" & Me.txtAccount & "));"

rst.Open query, CurrentProject.Connection

If rst.BOF And rst.EOF Then
    'nothing to do
    Exit Sub
End If

Do While Not rst.EOF

    Select Case rst.Fields(2).value
        Case 58 To 70   'SYSTEM MAINT CHARGE
            SysMaint = SysMaint & rst.Fields(3).value & " "
        Case 72 To 78    'METER SIZE - not used
        Case 79 To 82   'FIRE PROTECTION - not used
        Case Is < 58    'RATE
            rCode = rCode & rst.Fields(3).value & " "
        Case Is > 82    'RATE
            rCode = rCode & rst.Fields(3).value & " "
    End Select
    rst.MoveNext
Loop

    SysMaint = Trim(SysMaint)
    rCode = Trim(rCode)
    Me.txtMfgCode = SysMaint
    Me.txtRateCode = rCode

   On Error GoTo 0
   Exit Sub

LoadRateValues_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure LoadRateValues of VBA Document Form_Opt1Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure LoadRateValues of VBA Document Form_Opt1Form"

End Sub

Private Sub SaveRateValues()

End Sub
