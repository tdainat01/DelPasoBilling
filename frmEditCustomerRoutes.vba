Option Compare Database

Private Sub cboRoutesFilter_AfterUpdate()
   On Error GoTo cboRoutesFilter_AfterUpdate_Error

 If IsNull(Me.cboRoutesFilter) Then
        Me.FilterOn = False
    Else
        Me.filter = "route_id = " & Me.cboRoutesFilter.Column(0)
        Me.FilterOn = True
    End If

   On Error GoTo 0
   Exit Sub

cboRoutesFilter_AfterUpdate_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cboRoutesFilter_AfterUpdate of VBA Document Form_frmEditCustomerRoutes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cboRoutesFilter_AfterUpdate of VBA Document Form_frmEditCustomerRoutes"
End Sub

Private Sub cboRoutesFilter_Change()

   On Error GoTo cboRoutesFilter_Change_Error
    Me.Requery
    
   On Error GoTo 0
   Exit Sub

cboRoutesFilter_Change_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cboRoutesFilter_Change of VBA Document Form_frmEditCustomerRoutes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cboRoutesFilter_Change of VBA Document Form_frmEditCustomerRoutes"
End Sub

Private Sub cmdResync_Click()
Dim rst As New ADODB.Recordset
Dim query As String
Dim vArray As Variant
Dim RouteID As Long
Dim lRecs As Long
Dim lRecordsAffected As Long
Dim myDict As New colMissingNumClass

   On Error GoTo cmdResync_Click_Error

RouteID = Me.route_id

    Set myDict = FindMissingNumbers(RouteID)
    
    For lRecs = 0 To myDict.Count - 1
        'query = "DELETE FROM CustomerRoutes WHERE account_num = " & myDict.Item(CStr(lRecs)).account
        'CurrentProject.Connection.Execute query, lRecordsAffected
        query = "UPDATE CustomerRoutes set sequence = " & myDict.Item(CStr(lRecs)).MissingNum & " where account_num = " & myDict.Item(CStr(lRecs)).account
        CurrentProject.Connection.Execute query, lRecordsAffected
    Next lRecs

    Me.Requery
    MsgBox "Resequencing has completed.", vbOKOnly, "Done"
   On Error GoTo 0
   Exit Sub

cmdResync_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdResync_Click of VBA Document Form_frmEditCustomerRoutes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdResync_Click of VBA Document Form_frmEditCustomerRoutes"

End Sub

Private Sub Form_AfterDelConfirm(Status As Integer)
    Dim rst As New ADODB.Recordset
    Dim query As String
    Dim ctr As Integer
    Dim firstNum As Integer
    Dim lastNum As Integer
    Dim numCount As Integer
    
    'now we need to resequence all the existing routes
   On Error GoTo Form_AfterDelConfirm_Error

    query = "select * from CustomerRoutes where route_id = " & Me.route_id & " order by route_id, sequence"
    'query = "SELECT Max(CustomerRoutes.sequence) AS MaxOfsequence, Min(CustomerRoutes.sequence) AS MinOfsequence, Count(CustomerRoutes.sequence) AS CountOfsequence" & _
    '        " FROM CustomerRoutes GROUP BY CustomerRoutes.route_id HAVING (((CustomerRoutes.route_id)=" & Me.route_id & "));"

    rst.Open query, CurrentProject.Connection
    If rst.BOF And rst.EOF Then
        Exit Sub
    End If
    
    'fill in any number's missing in the sequence
    ctr = 0
    Do While Not rst.EOF
        If rst.Fields("sequence").value - ctr <> 0 Then
            'we have a number missing in the sequence
            
            
        End If
        rst.MoveNext
    Loop

   On Error GoTo 0
   Exit Sub

Form_AfterDelConfirm_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_AfterDelConfirm of VBA Document Form_frmEditCustomerRoutes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_AfterDelConfirm of VBA Document Form_frmEditCustomerRoutes"

End Sub

Private Sub Form_Delete(Cancel As Integer)
   On Error GoTo Form_Delete_Error


   On Error GoTo 0
   Exit Sub

Form_Delete_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Delete of VBA Document Form_frmEditCustomerRoutes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Delete of VBA Document Form_frmEditCustomerRoutes"

End Sub

Private Sub Form_Load()
    'select only those items according to the cboRoutesFilter
   On Error GoTo Form_Load_Error

    Me.cboRoutesFilter = Me.cboRoutesFilter.ItemData(0)
    Me.filter = "route_id = " & Me.cboRoutesFilter.ItemData(0)
    Me.Requery
   On Error GoTo 0
   Exit Sub

Form_Load_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Load of VBA Document Form_frmEditCustomerRoutes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Load of VBA Document Form_frmEditCustomerRoutes"
End Sub
