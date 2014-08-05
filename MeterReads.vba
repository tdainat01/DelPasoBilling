Option Compare Database

Private Sub cmdAdd_Click()
Dim query As String
Dim rst As New ADODB.Recordset
Dim lRecs As Long
Dim ans As String
Dim fErr As Boolean

   On Error GoTo cmdAdd_Click_Error
    fErr = False
    
    'validate all fields
    If IsNull(Me.batch_date) Or IsEmpty(Me.batch_date) Or Not IsDate(Me.batch_date) Then
        Call MsgBox("Either the batch date is missing or is not a valid date. " & _
            " Please fix this error condition first and then try again.", vbExclamation, "Not a Date")
        Exit Sub
    End If
    
    If IsNull(Me.account) Or IsEmpty(Me.account) Or Me.account <= 0 Then
        Call MsgBox("Either the account number is missing or is not a valid account number. " & _
            " Please fix this error condition first and then try again.", vbExclamation, "Not an Account")
        Exit Sub
    Else
        query = "select account from customer where account = " & Me.account & " AND status = 'A'"
        rst.Open query, CurrentProject.Connection
        If rst.BOF And rst.EOF Then
            Call MsgBox("The account number specified is not a valid account number. " & _
                " Please fix this error condition first and then try again.", vbExclamation, "Not an Account")
            Exit Sub
        End If
        rst.Close
    End If

    If IsNull(Me.normal_read) Or IsEmpty(Me.normal_read) Or Me.normal_read <= 0 Then
        Call MsgBox("Either the normal read is missing or is not a number. " & _
            " Please fix this error condition first and then try agaiin.", vbExclamation, "Not a Number")
        Exit Sub
    End If
    
    If IsNull(Me.low_read) Or IsEmpty(Me.low_read) Or Me.low_read < 0 Or Len(Me.low_read) < 1 Then
        Me.low_read = 0
    End If
    
    If IsNull(Me.txtUsage) Or IsEmpty(Me.txtUsage) Or Me.txtUsage = "" Or Len(Me.txtUsage) < 1 Then
        fErr = CalcUsage
         If fErr Then
         'an error occurred
             MsgBox "An error occurred trying to calculate usage. Please ensure that valid values were entered for" & _
                " [Normal Read], [Low Read] and/or [Previous Read] and then try again.", vbOKOnly + vbCritical, "Error"
         End If
    End If
    
    'Has the account been added into the MeterReads table already?
    query = "SELECT [account] FROM MeterReads WHERE [account] = " & Me.account
    rst.Open query, CurrentProject.Connection
    If rst.BOF And rst.EOF Then
        'do nothing as nothing was found
    Else
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
    End If
    rst.Close
    
    'now insert the value into the meter reads table and refresh the subform
    query = "INSERT INTO MeterReads([batch_date],[account],[normal_read],[low_read],[usage],[previous_read],[previous_date],[posted]) " & _
            " VALUES(#" & Me.batch_date & "#," & Me.account & "," & Me.normal_read & "," & Me.low_read & "," & Me.txtUsage & "," & _
            Me.txtPrevRead & ",#" & Me.txtPrevReadDate & "#," & Chr(34) & "N" & Chr(34) & ")"

    CurrentProject.Connection.Execute query, lRecs
    If lRecs < 1 Then
        'an error occurred
        Exit Sub
    End If
        
    Me.meterreads_subform.Requery
    Me.meterreads_subform.SetFocus

    
    Me.batch_date = Now
    Me.account = ""
    Me.normal_read = ""
    Me.low_read = ""
    Me.txtUsage = ""
    Me.txtPrevRead = ""
    Me.batch_date.SetFocus
    
   On Error GoTo 0
   Exit Sub

cmdAdd_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdAdd_Click of VBA Document Form_MeterReads")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdAdd_Click of VBA Document Form_MeterReads"
End Sub

Private Sub cmdDelete_Click()
   
   On Error GoTo cmdDelete_Click_Error

    Call MsgBox("Not yet implemented.", vbExclamation, "Uh")
    Exit Sub

   On Error GoTo 0
   Exit Sub

cmdDelete_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdDelete_Click of VBA Document Form_MeterReads")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdDelete_Click of VBA Document Form_MeterReads"

End Sub

Private Sub cmdPrint_Click()

   On Error GoTo cmdPrint_Click_Error

    'Call MsgBox("Not yet implemented.", vbExclamation, "Uh")
    'Exit Sub

    Dim query As String
    Dim dt As Date
    Dim rpt As Access.Report
    'dt = InputBox("Select a Date to View", "Date", Date)

    query = "SELECT [batch_id], [account], [normal_read] as current_read, [low_read], [batch_date] as current_date, [previous_read], " & _
            "[previous_date], ([normal_read]+[low_read])-[previous_read] AS [Usage] " & _
            "FROM MeterReads ORDER BY [batch_id]"
            '"FROM customer WHERE customer.[current_date] between #" & Format(dt, "mm/dd/yyyy") & " 00:00:00# and #" & Format(dt, "mm/dd/yyyy") & " 23:59:59#"

    Dim sReport As String
    sReport = "rptMeterReadsReport"

    DoCmd.OpenReport sReport, acViewDesign
    Set rpt = Reports.Item(sReport)
    rpt.RecordSource = query
    'rpt.OrderBy = ""
    rpt.OrderByOnLoad = True
    rpt.Width = 11232 'twips = 8.5 inches
    DoCmd.Close acReport, rpt.Name, acSaveYes
    DoCmd.OpenReport sReport, acViewPreview

   On Error GoTo 0
   Exit Sub

cmdPrint_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrint_Click of VBA Document Form_MeterReads")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrint_Click of VBA Document Form_MeterReads"

End Sub

Private Sub cmdQuit_Click()
   On Error GoTo cmdQuit_Click_Error

    'If Me.Dirty Then
        DoCmd.RunCommand acCmdSaveRecord
    'End If
    
    DoCmd.OpenForm "MeterMenu", acNormal
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdQuit_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdQuit_Click of VBA Document Form_MeterReads")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdQuit_Click of VBA Document Form_MeterReads"
End Sub

Private Sub cmdSave_Click()
Dim query As String
Dim msg As String
Dim rst As New ADODB.Recordset
Dim counter As Long
Dim nRead As Long
Dim lRead As Long
Dim lReadData As Long

   On Error GoTo cmdSave_Click_Error
   
    msg = ""
    counter = 0
    query = "select * from MeterReads"
    rst.Open query, CurrentProject.Connection
    
    If rst.BOF And rst.EOF Then
        'nothing to do
        Exit Sub
    End If

    query = ""
    
        
    
    Do While Not rst.EOF
        nRead = CLng(IIf(IsNull(rst.Fields("normal_read").value), 0, rst.Fields("normal_read").value))
        lRead = CLng(IIf(IsNull(rst.Fields("low_read").value), 0, rst.Fields("low_read").value))
        lReadData = nRead + lRead
        
        'insert the normal_read into the current_read field
        query = "UPDATE customer set customer.current_read = " & lReadData & ", customer.previous_read = " & _
                rst.Fields("previous_read").value & ", customer.current_date = #" & rst.Fields("batch_date").value & _
                "#, customer.previous_date = #" & rst.Fields("previous_date").value & _
                "# WHERE customer.account = " & rst.Fields("account").value
        
        CurrentProject.Connection.Execute query, lRecs
        If lRecs < 1 Then
            'an error occurred log the error
            msg = msg & rst.Fields("account").value & "," & rst.Fields("normal_read").value & "," & rst.Fields("batch_date").value & _
            "," & rst.Fields("previous_read").value & "," & rst.Fields("previous_date").value & vbCrLf
        End If
        rst.MoveNext
        counter = counter + 1
    Loop

    rst.Close
    
    If Len(msg) > 0 Then
        Call WriteError(msg)
    End If
    
    query = "DELETE From MeterReads"
        CurrentProject.Connection.Execute query, lRecs
        If lRecs < 1 Then
            'an error occurred
            Call MsgBox("The system was unable to purge the MeterReads table.", vbCritical, "Critical Error")
            Exit Sub
        End If
    
    Me.meterreads_subform.Requery
    
    If counter = 1 Then
        Call MsgBox("Successfully updated 1 Meter Read", vbInformation, "Done")
    Else
        Call MsgBox("Successfully updated " & counter & " Meter Reads", vbInformation, "Done")
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdSave_Click of VBA Document Form_MeterReads")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdSave_Click of VBA Document Form_MeterReads"
End Sub

Private Sub account_LostFocus()
Dim query As String
Dim rst As New ADODB.Recordset

   On Error GoTo account_LostFocus_Error

    If IsNull(Me.account) Or IsEmpty(Me.account) Or Me.account <= 0 Or Len(Me.account) < 1 Then
        'do nothing
        Exit Sub
    End If

    'query = "SELECT TOP 1 [normal_read] from MeterReads where account = " & Me.account & " ORDER BY [batch_date] DESC"
    query = "SELECT customer.[current_read], customer.[current_date] from customer where customer.[account] = " & Me.account

    rst.Open query, CurrentProject.Connection
    
    If rst.BOF And rst.EOF Then
        'there is no usage history. Usage is 0
        Me.txtPrevRead = 0
    Else
        Me.txtPrevRead = IIf(IsNull(rst.Fields(0).value), 0, rst.Fields(0).value)
        Me.txtPrevReadDate = IIf(IsNull(rst.Fields(1).value), "1/1/1900", rst.Fields(1).value)
    End If
    
    rst.Close
    

   On Error GoTo 0
   Exit Sub

account_LostFocus_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure account_LostFocus of VBA Document Form_MeterReads")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure account_LostFocus of VBA Document Form_MeterReads"
End Sub

Private Sub low_read_LostFocus()

   On Error GoTo low_read_LostFocus_Error
   Dim fErr As Boolean
   
   fErr = CalcUsage
    If fErr Then
    'an error occurred
        MsgBox "An error occurred trying to calculate usage. Please ensure that valid values were entered for" & _
           " [Normal Read], [Low Read] and/or [Previous Read] and then try again.", vbOKOnly + vbCritical, "Error"
    End If

   On Error GoTo 0
   Exit Sub

low_read_LostFocus_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure low_read_LostFocus of VBA Document Form_MeterReads")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure low_read_LostFocus of VBA Document Form_MeterReads"

End Sub

Private Sub WriteError(msg As String)
Dim sFileName As String

   On Error GoTo WriteError_Error

sFileName = CurrentProject.Path & "\MeterReadUpdateErrors.txt"

' does the file exist?  simpleminded test:
'If Len(Dir$(sFileName)) = 0 Then
'    Exit Sub
'End If
    
iFile = FreeFile
Open sFileName For Output As iFile
Print #iFile, msg
Close #iFile

    Call MsgBox("Wrote Error Data to" & sFileName, vbInformation, "Error Message")
    
   On Error GoTo 0
   Exit Sub

WriteError_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure WriteError of VBA Document Form_MeterReads")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure WriteError of VBA Document Form_MeterReads"

End Sub

Private Sub meterreads_subform_Enter()
    DoCmd.GoToRecord acActiveDataObject, , acLast
End Sub

Private Function CalcUsage() As Boolean
Dim fErr As Boolean

   On Error GoTo CalcUsage_Error
    fErr = False 'assume everything is OK
    
    If IsNull(Me.normal_read) Or IsEmpty(Me.normal_read) Or Me.normal_read <= 0 Or Len(Me.normal_read) < 1 Then
        'do nothing
        fErr = True
    End If

    If IsNull(Me.low_read) Or IsEmpty(Me.low_read) Or Me.low_read < 0 Or Len(Me.low_read) < 1 Then
        Me.low_read = 0
    End If

    'are the values in normal_read and prevread numbers
    If Not IsNumeric(Me.normal_read) And Not IsNumeric(Me.txtPrevRead) Then
        'alert the user
        Call MsgBox("Either the normal read or the previous read is not a number. " & _
            " Please fix this error condition first and then try agaiin.", vbExclamation, "Not a Number")
        fErr = True
    Else
            Me.txtUsage = (CLng(Me.normal_read) + CLng(Me.low_read)) - CLng(Me.txtPrevRead)
    End If

   On Error GoTo 0
   CalcUsage = fErr
   Exit Function

CalcUsage_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    'Call LogError(errNum, errSource, errMsg & " in procedure CalcUsage of VBA Document Form_MeterReads")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure CalcUsage of VBA Document Form_MeterReads"
    CalcUsage = True    'an error occurred
End Function
