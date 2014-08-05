Option Compare Database
Option Explicit

Private Sub cmdLiens_Click()
   On Error GoTo cmdLiens_Click_Error

If Me.Account_Number <> "" And IsNumeric(Me.Account_Number) Then
    Dim strQuery As String
    Dim rst As New ADODB.Recordset
    Dim ctr As Integer
    strQuery = "select * from customer where account = " & Me.Account_Number & " and lien <> 'A'"
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdLiens_Click of VBA Document Form_Opt 5 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdLiens_Click of VBA Document Form_Opt 5 Form"
End Sub

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

    Call LogError(errNum, errSource, errMsg & " in procedure chkService_Click of VBA Document Form_Opt 5 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure chkService_Click of VBA Document Form_Opt 5 Form"
End Sub

Private Sub cmdNext_Click()
Dim account As Long
Dim qry As String
Dim rst As New ADODB.Recordset
Dim rstPhones As New ADODB.Recordset


   On Error GoTo cmdNext_Click_Error

imgRedFlag.Visible = False
account = GetNextRecord(CLng(Me.Account_Number))
qry = "select * from customer where account = " & account '& " order by account"
rst.Open qry, CurrentProject.Connection

If rst.BOF And rst.EOF Then
    'otherwise do nothing
    Call MsgBox("The end of the file was encountered. No further records to display.", vbExclamation, "End of File")
    Exit Sub
End If

Do While Not rst.EOF
        Me.Account_Number = rst.Fields("account").value
        Me.Group_Number = rst.Fields("group").value
        Me.Master_Parcel_Number = rst.Fields("mastpar").value
        Me.Cycle = rst.Fields("cycle").value
        Me.Mfg_Code = rst.Fields("mfg_code").value
        Me.Account_Start_Date = rst.Fields("start_date").value
        Me.Status = rst.Fields("status").value
        Me.Meter__ = rst.Fields("meter_number").value
        Me.Account_Term_Date = rst.Fields("term_date").value
        Me.Out_Town = rst.Fields("out_town").value
        Me.Meter_Size = rst.Fields("meter_size").value
        Me.Property_Use = rst.Fields("property_use").value
        Me.Backflow = rst.Fields("backflow").value
        Me.chkService = rst.Fields("service_discon").value
        If Not IsNull(Me.chkService) Then
            If Me.chkService = True Then
                imgRedFlag.Visible = True
            End If
        End If
        Me.Fire_Size = rst.Fields("fire_size").value
        Me.Unit_of_Measure = rst.Fields("unit_measure").value
        Me.Current_Read = rst.Fields("current_read").value
        Me.Current_read_date = rst.Fields("current_date").value
        Me.Rate_code = rst.Fields("rate_code").value
        Me.Previous_Read = rst.Fields("previous_read").value
        Me.Previous_read_date = rst.Fields("previous_date").value
        Me.Gallons_Cubic = rst.Fields("gal_cub_used").value
        Me.Meter_Valve_Site = rst.Fields("meter_site").value
        Me.Deposit = Round(rst.Fields("deposit").value, 2)
        Me.Use_Charge = Round(rst.Fields("use_charge").value, 2)
        Me.Past_Due = Round(rst.Fields("past_due").value, 2)
        Me.Prev_Balance = Round(rst.Fields("prev_balance").value, 2)
        Me.Current_Due = Round(rst.Fields("current_due").value, 2)
        Me.Special_Credit = Round(rst.Fields("special_credit").value, 2)
        Me.Total_Due = Round(rst.Fields("total_due").value, 2)
        Me.Special_Charge = Round(rst.Fields("special_charge").value, 2)
        'Me.txtSpecialDescr = rst.Fields("special_description").value   'not included on this form
        Me.Physical_Address = rst.Fields("phy_address").value
        Me.Owner = rst.Fields("name").value
        Me.BILL_TO__Name = rst.Fields("bill_name").value
        Me.CARE_OF__Name = rst.Fields("care_of").value
        Me.Address = rst.Fields("addr1").value
        Me.City__State = rst.Fields("city").value & " " & rst.Fields("state").value
        Me.Zip = rst.Fields("zip").value
        
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
            " WHERE (((customer.account)=" & CLng(Me.Account_Number) & "));"
    
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdNext_Click of VBA Document Form_Opt 5 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdNext_Click of VBA Document Form_Opt 5 Form"

End Sub

Private Sub cmdNotes_Click()
   On Error GoTo cmdNotes_Click_Error
    If Me.Account_Number = "" Then
        Call MsgBox("To use the notes feature, please select a valid account.", vbExclamation, "No Account")
        Exit Sub
    End If

    'test to see if it is open but not visible
    Dim bool As Boolean
    bool = CheckFormStatus("frmNotes")
    If bool Then
        DoCmd.Close acForm, "frmNotes", acSaveYes
    End If
    vOpenArgs = Me.Account_Number
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

Private Sub cmdPayments_Click()
   On Error GoTo cmdPayments_Click_Error

If Me.Account_Number <> "" And IsNumeric(Me.Account_Number) Then
    sCallingForm = Me.Name
    DoCmd.OpenReport "rptAccountPayment", acViewReport, , , , Me.Account_Number
    DoCmd.Close acForm, Me.Name, acSaveYes
Else
    Call MsgBox("The account number field must contain a numeric value. Please enter an address to search for and try again", _
    vbInformation + vbOKOnly, "Alert")
    'DoCmd.OpenForm "DPM Main Menu", acNormal
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPayments_Click of VBA Document Form_Opt 5 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPayments_Click of VBA Document Form_Opt 5 Form"
End Sub

Private Sub cmdPrev_Click()
Dim account As Long
Dim qry As String
Dim rst As New ADODB.Recordset
Dim rstPhones As New ADODB.Recordset

   On Error GoTo cmdPrev_Click_Error

imgRedFlag.Visible = False
account = GetPrevRecord(CLng(Me.Account_Number))
qry = "select * from customer where account = " & account
rst.Open qry, CurrentProject.Connection

If rst.BOF And rst.EOF Then
    'otherwise do nothing
    Call MsgBox("The beginning  of the file was encountered. No further records to display.", vbExclamation, "Beginning of File")
    Exit Sub
End If

Do While Not rst.EOF
        Me.Account_Number = rst.Fields("account").value
        Me.Group_Number = rst.Fields("group").value
        Me.Master_Parcel_Number = rst.Fields("mastpar").value
        Me.Cycle = rst.Fields("cycle").value
        Me.Mfg_Code = rst.Fields("mfg_code").value
        Me.Account_Start_Date = rst.Fields("start_date").value
        Me.Status = rst.Fields("status").value
        Me.Meter__ = rst.Fields("meter_number").value
        Me.Account_Term_Date = rst.Fields("term_date").value
        Me.Out_Town = rst.Fields("out_town").value
        Me.Meter_Size = rst.Fields("meter_size").value
        Me.Property_Use = rst.Fields("property_use").value
        Me.Backflow = rst.Fields("backflow").value
        Me.chkService = rst.Fields("service_discon").value
        If Not IsNull(Me.chkService) Then
            If Me.chkService = True Then
                imgRedFlag.Visible = True
            End If
        End If
        Me.Fire_Size = rst.Fields("fire_size").value
        Me.Unit_of_Measure = rst.Fields("unit_measure").value
        Me.Current_Read = rst.Fields("current_read").value
        Me.Current_read_date = rst.Fields("current_date").value
        Me.Rate_code = rst.Fields("rate_code").value
        Me.Previous_Read = rst.Fields("previous_read").value
        Me.Previous_read_date = rst.Fields("previous_date").value
        Me.Gallons_Cubic = rst.Fields("gal_cub_used").value
        Me.Meter_Valve_Site = rst.Fields("meter_site").value
        Me.Deposit = Round(rst.Fields("deposit").value, 2)
        Me.Use_Charge = Round(rst.Fields("use_charge").value, 2)
        Me.Past_Due = Round(rst.Fields("past_due").value, 2)
        Me.Prev_Balance = Round(rst.Fields("prev_balance").value, 2)
        Me.Current_Due = Round(rst.Fields("current_due").value, 2)
        Me.Special_Credit = Round(rst.Fields("special_credit").value, 2)
        Me.Total_Due = Round(rst.Fields("total_due").value, 2)
        Me.Special_Charge = Round(rst.Fields("special_charge").value, 2)
        'Me.txtSpecialDescr = rst.Fields("special_description").value '-- NOT USED
        Me.Physical_Address = rst.Fields("phy_address").value
        Me.Owner = rst.Fields("name").value
        Me.BILL_TO__Name = rst.Fields("bill_name").value
        Me.CARE_OF__Name = rst.Fields("care_of").value
        Me.Address = rst.Fields("addr1").value
        Me.City__State = rst.Fields("city").value & " " & rst.Fields("state").value
        Me.Zip = rst.Fields("zip").value
        
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
            " WHERE (((customer.account)=" & CLng(Me.Account_Number) & "));"
    
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrev_Click of VBA Document Form_Opt 5 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrev_Click of VBA Document Form_Opt 5 Form"

End Sub

Private Sub cmdUsage_Click()
Dim rst As New ADODB.Recordset
Dim query As String

   On Error GoTo cmdUsage_Click_Error

If IsNull(Me.Account_Number) Or Me.Account_Number = "" Or IsEmpty(Me.Account_Number) Or Me.Account_Number <= 0 Then
    Call MsgBox("No valid account number has been specified. Please enter a valid address that locates a valid " & _
     " account number and then try again.", _
        vbExclamation Or vbSystemModal, "No Account")
    Exit Sub
End If
    
    query = "select top 2 batch_date, normal_read from MeterReads where posted = 'Y' and account = " & Me.Account_Number
    rst.Open query, CurrentProject.Connection
    If rst.BOF And rst.EOF Then
        Call MsgBox("No meter read information has been returned for " & Me.Account_Number & ".", _
        vbExclamation Or vbSystemModal, "No Meter Reads")
        Exit Sub
    Else
        DoCmd.OpenForm "frmUsage", acNormal, , , acFormReadOnly, acDialog, Me.Account_Number
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdUsage_Click of VBA Document Form_Opt 5 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdUsage_Click of VBA Document Form_Opt 5 Form"

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim strCharacter As String

    ' Convert ANSI value to character string.
    strCharacter = Chr(KeyAscii)
    ' Convert character to upper case, then to ANSI value.
    KeyAscii = Asc(UCase(strCharacter))
    If KeyAscii = 27 Then
        Call ResetValues
        Me.Account_Number = ""
        Me.Physical_Address.SetFocus
    End If

End Sub

Private Sub CustomerLookup()
'On Error GoTo 0 'ignore errors
Dim rst As New ADODB.Recordset
Dim query As String
Dim lRecs As Long
Dim pNum1 As String
Dim pNum2 As String
Dim pNum3 As String
Dim pType1 As String
Dim pType2 As String
Dim pType3 As String
Dim vOpenArgs As Variant

   On Error GoTo CustomerLookup_Error
    vOpenArgs = Me.OpenArgs
    Me.imgRedFlag.Visible = False
        If IsNumeric(vOpenArgs) Then
            query = "SELECT phy_address from customer where account = " & vOpenArgs
            rst.Open query, CurrentProject.Connection, adOpenDynamic, adLockPessimistic
            If rst.BOF And rst.EOF Then
                'nothing was found
            Else
                Me.Physical_Address = rst.Fields(0).value
            End If
            rst.Close
        End If
            
    If Me.Physical_Address <> "" Then
        query = "SELECT * FROM customer WHERE (((customer.phy_address) Like '%" & Me.Physical_Address & "%'));"
        'rst.Open query, CurrentProject.Connection, adOpenForwardOnly, adLockPessimistic
        rst.Open query, CurrentProject.Connection, adOpenDynamic, adLockPessimistic
        
        If rst.BOF And rst.EOF Then
            'seems like no records were pulled
            If CurrentProject.Connection.Errors.Count > 0 Then
                Dim x As Integer
                For x = 0 To CurrentProject.Connection.Errors.Count
                    Debug.Print CurrentProject.Connection.Errors.Item(x)
                Next x
            End If
            Call ResetValues
            Exit Sub
        End If
        'Take only the first Match
        Dim ctr As Integer
        ctr = 0
        Do While Not rst.EOF
            If ctr >= 1 Then
                Exit Do
            End If
            
            Me.Account_Number = rst.Fields("account").value
            Me.Group_Number = rst.Fields("group").value
            Me.Master_Parcel_Number = rst.Fields("mastpar").value
            Me.Cycle = rst.Fields("cycle").value
            Me.Mfg_Code = rst.Fields("mfg_code").value
            Me.Account_Start_Date = rst.Fields("start_date").value
            Me.Status = rst.Fields("status").value
            Me.Meter__ = rst.Fields("meter_number").value
            Me.Account_Term_Date = rst.Fields("term_date").value
            
            If rst.Fields("out_town").value = False Then
                Me.Out_Town.value = False
            Else
                Me.Out_Town = True
            End If
            
            Me.Meter_Size = rst.Fields("meter_size").value
            Me.Property_Use = rst.Fields("property_use").value
            
            If rst.Fields("backflow").value = False Then
                Me.Backflow.value = False
            Else
                Me.Backflow.value = True
            End If
            
            If rst.Fields("service_discon").value = False Then
                Me.chkService = False
                imgRedFlag.Visible = False
            Else
                Me.chkService = True
                imgRedFlag.Visible = True
            End If
            
            Me.Fire_Size = rst.Fields("fire_size").value
            Me.Unit_of_Measure = rst.Fields("unit_measure").value
            Me.Current_Read = rst.Fields("current_read").value
            Me.Current_read_date = rst.Fields("current_date").value
            Me.Rate_code = rst.Fields("rate_code").value
            Me.Previous_Read = rst.Fields("previous_read").value
            Me.Gallons_Cubic = rst.Fields("gal_cub_used").value
            Me.Meter_Valve_Site = rst.Fields("meter_site").value
            Me.Deposit = Round(rst.Fields("deposit").value, 2)
            Me.Use_Charge = Round(rst.Fields("use_charge").value, 2)
            Me.Past_Due = Round(rst.Fields("past_due").value, 2)
            Me.Prev_Balance = Round(rst.Fields("prev_balance").value, 2)
            Me.Current_Due = Round(rst.Fields("current_due").value, 2)
            Me.Special_Credit = Round(rst.Fields("special_credit").value, 2)
            Me.Total_Due = Round(rst.Fields("total_due").value, 2)
            Me.Special_Charge = Round(rst.Fields("special_charge").value, 2)
            'There is no special description field ??
            Me.Physical_Address = rst.Fields("phy_address").value
            Me.Address = rst.Fields("addr1").value
            Me.BILL_TO__Name = rst.Fields("bill_name").value
            Me.CARE_OF__Name = rst.Fields("care_of").value
            
            'No lien field?
            Me.Owner = rst.Fields("name").value
            Me.CARE_OF__Name = rst.Fields("bill_name").value
            Me.City__State = rst.Fields("city").value & " " & rst.Fields("state").value
            Me.Zip = rst.Fields("zip").value
            'no comment field?
            rst.MoveNext
            ctr = ctr + 1
        Loop
        
        rst.Close
        
        'Now open the phones
        query = "SELECT Phones.Phone1, Phones.Phone2, Phones.Phone3" & _
                " FROM Phones WHERE (((Phones.CustomerID)=" & CLng(Me.Account_Number) & "));"
        
        rst.Open query, CurrentProject.Connection
        
        If rst.BOF And rst.EOF Then
            Me.txtWorkPhone = ""
            Me.txtHomePhone = ""
            Me.txtMobilePhone = ""
            Exit Sub
        End If
        
        Do While Not rst.EOF
            pNum1 = IIf(IsNull(rst.Fields("phone1").value), "", rst.Fields("phone1").value)
            pNum2 = IIf(IsNull(rst.Fields("phone2").value), "", rst.Fields("phone2").value)
            pNum3 = IIf(IsNull(rst.Fields("phone3").value), "", rst.Fields("phone3").value)
        rst.MoveNext
        Loop
    
            Me.txtWorkPhone = pNum1
            Me.txtHomePhone = pNum2
            Me.txtMobilePhone = pNum3
        
        Else
            Call ResetValues
        End If
    'rst.Close

   On Error GoTo 0
   Exit Sub

CustomerLookup_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure CustomerLookup of VBA Document Form_Opt 5 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure CustomerLookup of VBA Document Form_Opt 5 Form"
End Sub

Private Sub ResetValues()
        'Reset all the fields
   On Error GoTo ResetValues_Error

        Me.Group_Number = ""
            Me.Master_Parcel_Number = ""
            Me.Cycle = 0
            Me.Mfg_Code = ""
            Me.Account_Start_Date = ""
            Me.Status = ""
            Me.Meter__ = ""
            Me.Account_Term_Date = ""
            Me.Out_Town = ""
            Me.Meter_Size = ""
            Me.Property_Use = ""
            Me.Backflow = ""
            Me.chkService = ""
            Me.Fire_Size = ""
            Me.Unit_of_Measure = ""
            Me.Current_Read = ""
            Me.Current_read_date = ""
            Me.Rate_code = ""
            Me.Previous_Read = ""
            Me.Gallons_Cubic = ""
            Me.Meter_Valve_Site = ""
            Me.Deposit = ""
            Me.Use_Charge = ""
            Me.Past_Due = ""
            Me.Prev_Balance = ""
            Me.Current_Due = ""
            Me.Special_Credit = ""
            Me.Total_Due = ""
            Me.Special_Charge = ""
            'There is no special description field ??
            Me.Physical_Address = ""
            Me.Address = ""
            Me.BILL_TO__Name = ""
            Me.CARE_OF__Name = ""
            
            'No lien field?
            Me.Owner = ""
            Me.CARE_OF__Name = ""
            Me.City__State = ""
            Me.Zip = ""
            'no comment field?

   On Error GoTo 0
   Exit Sub

ResetValues_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure ResetValues of VBA Document Form_Opt 5 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure ResetValues of VBA Document Form_Opt 5 Form"
End Sub

Private Sub Form_Load()

    'for each control on the page
    Dim x As Integer
   On Error GoTo Form_Load_Error

    For x = 0 To Me.Controls.Count - 1
    'if the control is disabled, set its backcolor to #FFFFFF
        If Me.Controls.Item(x).ControlType = 109 Then
            If Me.Controls.Item(x).Enabled = False Then
                Me.Controls.Item(x).BackColor = 16775416
            End If
        End If
    Next
    
    Me.Physical_Address.SetFocus
    
   On Error GoTo 0
   Exit Sub

Form_Load_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Load of VBA Document Form_Opt 5 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Load of VBA Document Form_Opt 5 Form"

End Sub

Private Sub Physical_Address_Enter()
   On Error GoTo Physical_Address_Enter_Error

    Call CustomerLookup

   On Error GoTo 0
   Exit Sub

Physical_Address_Enter_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Physical_Address_Enter of VBA Document Form_Opt 5 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Physical_Address_Enter of VBA Document Form_Opt 5 Form"
End Sub

Private Sub Physical_Address_LostFocus()
   On Error GoTo Physical_Address_LostFocus_Error

    Call CustomerLookup

   On Error GoTo 0
   Exit Sub

Physical_Address_LostFocus_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Physical_Address_LostFocus of VBA Document Form_Opt 5 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Physical_Address_LostFocus of VBA Document Form_Opt 5 Form"
End Sub

Private Sub Zip_AfterUpdate()
    Call FormatZip(Me.Zip)
End Sub
