Option Compare Database
Option Explicit

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

    Call LogError(errNum, errSource, errMsg & " in procedure chkService_Click of VBA Document Form_Opt 3 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure chkService_Click of VBA Document Form_Opt 3 Form"
End Sub

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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdExit_Click of VBA Document Form_Opt 3 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdExit_Click of VBA Document Form_Opt 3 Form"
End Sub

Private Sub cmdUsage_Click()
Dim rst As New ADODB.Recordset
Dim query As String

   On Error GoTo cmdUsage_Click_Error

If IsNull(Me.Account_Number) Or Me.Account_Number = "" Or IsEmpty(Me.Account_Number) Or Me.Account_Number <= 0 Then
    Call MsgBox("No valid account number has been specified. Please enter a valid parcel number that locates a valid " & _
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdUsage_Click of VBA Document Form_Opt 3 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdUsage_Click of VBA Document Form_Opt 3 Form"

End Sub

Private Sub Form_Load()
    Call CustomerLookup
End Sub

Private Sub Master_Parcel_Number_Enter()
    Call CustomerLookup
End Sub

Private Sub Master_Parcel_Number_LostFocus()
    Call CustomerLookup
End Sub

Private Sub CustomerLookup()
   On Error GoTo CustomerLookup_Error

Dim rst As New ADODB.Recordset
Dim query As String
Dim pNum1 As String
Dim pNum2 As String
Dim pNum3 As String
Dim vOpenArgs As Variant

vOpenArgs = Me.OpenArgs
Me.imgRedFlag.Visible = False

If Not IsNull(vOpenArgs) And IsNumeric(vOpenArgs) Then
    query = "SELECT mastpar from customer where account = " & vOpenArgs
    rst.Open query, CurrentProject.Connection
    If rst.BOF And rst.EOF Then
        'do nothing
    Else
        Me.Master_Parcel_Number = rst.Fields(0).value
    End If
    rst.Close
End If

    If Me.Master_Parcel_Number <> "" Then
        query = "SELECT * FROM customer where MASTPAR = '" & Me.Master_Parcel_Number & "'"
        rst.Open query, CurrentProject.Connection
        
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
            'No lien field?
            Me.Owner = rst.Fields("name").value
            Me.CARE_OF__Name = rst.Fields("bill_name").value
            Me.City__State = rst.Fields("city").value & " " & rst.Fields("state").value
            Me.Zip = rst.Fields("zip").value
            'no comment field?
            rst.MoveNext
        Loop
            rst.Close
            
        'Now open the phones
        query = "SELECT Phones.Phone1, Phones.Phone2, Phones.Phone3" & _
                " FROM Phones WHERE (((Phones.CustomerID)=" & CLng(Me.Account_Number) & "));"
        
        rst.Open query, CurrentProject.Connection
        
        If rst.BOF And rst.EOF Then
            rst.Close
            Exit Sub
        End If
        
        Do While Not rst.EOF
            pNum1 = IIf(IsNull(rst.Fields("phone1").value), "", rst.Fields("phone1").value)
            pNum2 = IIf(IsNull(rst.Fields("phone2").value), "", rst.Fields("phone2").value)
            pNum3 = IIf(IsNull(rst.Fields("phone3").value), "", rst.Fields("phone3").value)
        rst.MoveNext
        Loop
        rst.Close
        
            Me.txtWorkPhone = pNum1
            Me.txtHomePhone = pNum2
            Me.txtMobilePhone = pNum3
    Else
        Call ResetValues
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

    Call LogError(errNum, errSource, errMsg & " in procedure CustomerLookup of VBA Document Form_Opt 3 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure CustomerLookup of VBA Document Form_Opt 3 Form"

End Sub

Private Sub cmdLiens_Click()
If Me.Master_Parcel_Number <> "" And IsNumeric(Me.Master_Parcel_Number) Then
    Dim strQuery As String
    Dim rst As New ADODB.Recordset
    Dim ctr As Integer
    strQuery = "select * from customer where mastpar = '" & Me.Master_Parcel_Number & "' and lien <> 'A'"
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
End Sub

Private Sub cmdPayments_Click()
If Me.Master_Parcel_Number <> "" Then
    sCallingForm = Me.Name
    DoCmd.OpenReport "rptAccountPayment", acViewReport, , , , Me.Account_Number
    DoCmd.Close acForm, Me.Name, acSaveYes
Else
    Call MsgBox("The parcel number field must contain a valid parcel value.", vbInformation + vbOKOnly, "Alert")
    'DoCmd.OpenForm "DPM Main Menu", acNormal
    'DoCmd.Close acForm, Me.name, acSaveYes
End If
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
        Me.Master_Parcel_Number.SetFocus
    End If

End Sub

Private Sub ResetValues()
        'Reset all the fields
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
            'No lien field?
            Me.Owner = ""
            Me.CARE_OF__Name = ""
            Me.City__State = ""
            Me.Zip = ""
            Me.txtWorkPhone = ""
            Me.txtHomePhone = ""
            Me.txtMobilePhone = ""
            'no comment field?
End Sub

