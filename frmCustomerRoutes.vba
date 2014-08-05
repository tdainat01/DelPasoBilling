Option Compare Database
Dim fBeingDeleted As Boolean
Dim fCanceled As Boolean

Private Sub account_num_LostFocus()
Dim query As String
Dim seq As Long
Dim rst As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim topSeq As Long

   On Error GoTo account_num_LostFocus_Error
   
    If fBeingDeleted Then
        fBeingDeleted = False
        Exit Sub
    End If
    If IsNull(Me.account_num.text) Or Me.account_num.text = "" Then
        query = "select * from temp_CustomerRoutes"
        rst.Open query, CurrentProject.Connection
        If rst.BOF And rst.EOF Then
            'nothing to do
            Exit Sub
        End If
        rst.Close
    End If
    'first see if the account number entered even exists
    If Me.account_num.text = "" Then
        Exit Sub
    End If
    
    query = "SELECT account from customer where account = " & Me.account_num
    rst.Open query, CurrentProject.Connection
    If rst.BOF And rst.EOF Then
        'the account number doesn't exist
        Call MsgBox("The account number you entered " & Me.account_num & _
            " was not found in the customer table. Please check this number then try again.", _
            vbExclamation, "Account Not Found")
        Me.account_num = ""
        Me.account_num.SetFocus
        Exit Sub
    End If
    rst.Close
    
    Me.route_id = Forms("frmRoutes")!cboRoutes.Column(1)
    
    If IsNull(Me.sequence) Or Me.sequence = "" Then
        seq = GetNextNumber(Me.route_id)
    End If
    
    'next check to see if either the account number or the route_id/sequence combo already exists
    query = "select * from temp_CustomerRoutes where account_num = " & Me.account_num & _
            " or (route_id = " & Me.route_id & " and sequence = " & seq & ")"

    rst.Open query, CurrentProject.Connection
    
    If rst.BOF And rst.EOF Then
        'do nothing, all is good
    Else
        ' a record already exists
        Call MsgBox("A record for the account number already exists", vbCritical, "Record Exists")
        Me.account_num = ""
        fCanceled = True
        Exit Sub
    End If
    rst.Close
    
    query = "select * from CustomerRoutes where account_num = " & Me.account_num & _
            " or (route_id = " & Me.route_id & " and sequence = " & seq & ")"

    rst.Open query, CurrentProject.Connection
    
    If rst.BOF And rst.EOF Then
        'do nothing, all is good
    Else
        ' a record already exists
        Call MsgBox("A record for the account number already exists", vbCritical, "Record Exists")
        Me.account_num = ""
        fCanceled = True
        Exit Sub
    End If
    rst.Close
    
        Me.sequence = seq
        fCanceled = False
        
'    'Get the next sequence number
'    query = "select top 1 sequence from CustomerRoutes order by sequence desc"
'    rst.Open query, CurrentProject.Connection
'    If rst.BOF And rst.EOF Then
'        seq = 1
'    Else
'        Do While Not rst.EOF
'            seq = IIf(IsNull(rst.fields(0).value), 1, rst.fields(0).value)
'            rst.MoveNext
'        Loop
'    End If
'    rst.Close
'
'    'check to see if that number is in use
'    query = "select top 1 sequence" & _
'            " from temp_CustomerRoutes" & _
'            " order by sequence desc"
'
'    rst.Open query, CurrentProject.Connection
'    If rst.BOF And rst.EOF Then
'        seq = seq + 1
'    Else
'        Do While Not rst.EOF
'            topSeq = IIf(IsNull(rst.fields(0).value), 1, rst.fields(0).value)
'            If Not IsNull(Me.sequence) Or Me.sequence <> "" Then
'                If topSeq = Me.sequence Then
'                    'do nothing
'                Else
'                'now we need to check if topSeq is in use in the tem_CustomerRoutes table
'                    query = "select sequence from temp_CustomerRoutes" & _
'                            " order by sequence desc"
'                    rst2.Open query, CurrentProject.Connection
'                    If rst2.BOF And rst2.EOF Then
'                        'there is nothing in the temp_CustomerRoutes table.
'                        'technically we should never get here
'                        If seq < topSeq Then
'                           seq = topSeq + 1
'                        Else
'                            seq = seq + 1
'                        End If
'                    Else
'                    'loop through the temp_CustomerRoutes table and find out if topSeq is in use.
'                        Do While Not rst2.EOF
'                            If Not IsNull(rst2.fields(0).value) Then
'                                If topSeq = rst2.fields(0).value Then
'                                    'well, the number is in use.
'                                    Exit Sub
'                                End If
'                            End If
'                            rst2.MoveNext
'                        Loop
'                    End If
'                    If seq < topSeq Then
'                       seq = topSeq + 1
'                    Else
'                        seq = seq + 1
'                    End If
'                End If
'            End If
'            rst.MoveNext
'        Loop
'    End If
'
'    Me.sequence = seq
    
   On Error GoTo 0
   Exit Sub

account_num_LostFocus_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    'Call LogError(errNum, errSource, errMsg & " in procedure account_num_LostFocus of VBA Document Form_frmCustomerRoutes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure account_num_LostFocus of VBA Document Form_frmCustomerRoutes"
End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
    If fCanceled Then
        Cancel = True
    End If
End Sub

Private Sub Form_Delete(Cancel As Integer)
    
   On Error GoTo Form_Delete_Error

    fBeingDeleted = True

   On Error GoTo 0
   Exit Sub

Form_Delete_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Delete of VBA Document Form_frmCustomerRoutes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Delete of VBA Document Form_frmCustomerRoutes"
End Sub

