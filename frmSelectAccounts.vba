Option Compare Database
Option Explicit

Private Sub cboFilter_Change()

   On Error GoTo cboFilter_Change_Error

If cboFilter.text = "Select Accounts" Then
    lstAccounts.Visible = True
Else
    lstAccounts.Visible = False
End If

   On Error GoTo 0
   Exit Sub

cboFilter_Change_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cboFilter_Change of VBA Document Form_frmSelectAccounts")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cboFilter_Change of VBA Document Form_frmSelectAccounts"

End Sub

Private Sub cmdExit_Click()

   On Error GoTo cmdExit_Click_Error
    DoCmd.OpenForm "ReportMenu", acNormal
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdExit_Click of VBA Document Form_frmSelectAccounts")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdExit_Click of VBA Document Form_frmSelectAccounts"
End Sub

Private Sub cmdSelect_Click()

   On Error GoTo cmdSelect_Click_Error
    Dim sAccounts As String
    Dim ix As Long
    Dim iy As Long
    Dim lRecs As Long
    Dim selItem As Variant
    Dim query As String
    Dim rst As New ADODB.Recordset
    Dim arrAccount() As String
    Dim clsAcct() As clsState

    If lstAccounts.Visible = True Then
         If lstAccounts.ItemsSelected.Count > 0 Then
             For Each selItem In lstAccounts.ItemsSelected
                 sAccounts = sAccounts & lstAccounts.Column(0, selItem) & ","
             Next
             If Right(sAccounts, 1) = "," Then
                  sAccounts = Left(sAccounts, Len(sAccounts) - 1)
             End If
         End If
        
         If Len(sAccounts) > 0 Then
            'do nothing
         Else
             Call MsgBox("No accounts were selected. Please go back, select the accounts you wish to view and then try again.", vbExclamation, "Nothing to do")
             Exit Sub
         End If
    Else 'all items were selected
        'alert that it may take a while to build this report
        Select Case MsgBox("Selecting all records can take a long time to process. It may cause your system to appear unresponsive." & _
            "If you choose to continue, please be patient and let the system complete. Do you want to continue (Yes to go on)?", _
                vbYesNo Or vbQuestion Or vbDefaultButton1, "Are you sure?")
        
            Case vbYes
                query = "SELECT customer.account FROM customer" & _
                        " WHERE (((customer.status)<>'I') AND ((customer.term_date) Is Null Or (customer.term_date)=#1/1/1900#));"
                rst.Open query, CurrentProject.Connection
                
                If rst.BOF And rst.EOF Then
                    Call MsgBox("No accounts were selected. Please go back, select the accounts you wish to view and then try again.", vbExclamation, "Nothing to do")
                    Exit Sub
                End If
                
                Do While Not rst.EOF
                    sAccounts = sAccounts & rst.Fields(0).value & ","
                    rst.MoveNext
                Loop
                
                If Right(sAccounts, 1) = "," Then
                     sAccounts = Left(sAccounts, Len(sAccounts) - 1)
                End If
            Case vbNo
                Exit Sub
        End Select
    End If
   
    'calculate all charges
    arrAccount = Split(sAccounts, ",")
    ReDim clsAcct(UBound(arrAccount))
   
    clsAcct = CalculateCharges(sAccounts)
    
    'empty both temp tables
    CurrentProject.Connection.Execute "DELETE * FROM temp_ServiceCharges"
    CurrentProject.Connection.Execute "DELETE * FROM temp_RecurringCharges"
    
    'inset all values into a table
    For ix = 0 To UBound(clsAcct)
        If clsAcct(ix) Is Nothing Then
            'do nothing
            Debug.Print "clsAcct(" & ix & ") is nothing"
        Else
            query = "insert into temp_ServiceCharges([account],[Description],[Amount]) VALUES (" & _
                clsAcct(ix).account & ",'" & Replace(clsAcct(ix).ServiceDescription, "'", "''") & "','" & clsAcct(ix).Service & "')"
            CurrentProject.Connection.Execute query, lRecs
            If lRecs < 1 Then
                'an error occurred. deal with it
            End If
        End If
    Next
   
    For ix = 0 To UBound(clsAcct)
        If clsAcct(ix) Is Nothing Then
            'do nothing
            Debug.Print "clsAcct(" & ix & ") is nothing"
        Else
            For iy = 1 To clsAcct(ix).CountRecurring
                If clsAcct(ix).UsageCharge > 0 And clsAcct(ix).Flag = False Then
                    query = "insert into temp_RecurringCharges([account],[charge_description],[charge_amount]) VALUES(" & _
                            clsAcct(ix).ItemRecurring(iy).account & ",'" & "Usage Charge" & "'," & _
                            clsAcct(ix).UsageCharge & ")"
                    clsAcct(ix).Flag = True
                Else
                    query = "insert into temp_RecurringCharges([account],[charge_description],[charge_amount]) VALUES(" & _
                            clsAcct(ix).ItemRecurring(iy).account & ",'" & Replace(clsAcct(ix).ItemRecurring(iy).Description, "'", "''") & "'," & _
                            clsAcct(ix).ItemRecurring(iy).Charge & ")"
                End If
                CurrentProject.Connection.Execute query, lRecs
                If lRecs < 1 Then
                    'an error occurred. deal with it
                End If
            Next
        End If
    Next
    DoCmd.OpenReport "rptTotalCharged", acViewReport
    
   On Error GoTo 0
   Exit Sub

cmdSelect_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdSelect_Click of VBA Document Form_frmSelectAccounts")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdSelect_Click of VBA Document Form_frmSelectAccounts"
End Sub
