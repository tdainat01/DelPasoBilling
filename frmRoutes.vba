Option Compare Database
Option Explicit

Dim strSender As String

Private Sub cboRoutes_Change()
Dim query As String
Dim rst As New ADODB.Recordset
Dim I As Integer

   On Error GoTo cboRoutes_Change_Error
   I = Me.cboRoutes.ListIndex + 1
    query = "SELECT * FROM CUSTOMERROUTES WHERE route_id = " & I
   rst.Open query, CurrentProject.Connection
   
    If rst.BOF And rst.EOF Then
        Exit Sub
    End If

   On Error GoTo 0
   Exit Sub

cboRoutes_Change_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cboRoutes_Change of VBA Document Form_frmRoutes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cboRoutes_Change of VBA Document Form_frmRoutes"
End Sub

Private Sub cmdExit_Click()

   On Error GoTo cmdExit_Click_Error

    If strSender = "" Then
        DoCmd.OpenForm "DPM Main Menu", acNormal
    Else
        DoCmd.OpenForm strSender, acNormal
    End If
    
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdExit_Click of VBA Document Form_frmRoutes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdExit_Click of VBA Document Form_frmRoutes"

End Sub

Private Sub cmdHelp_Click()
    If Me.txtHelp.Visible = True Then
        Me.txtHelp.Visible = False
    Else
        Me.txtHelp.Visible = True
    End If
    
    If Me.TabCtl1.value = 1 Then
        Me.txtHelp = "To make changes in the existing routes, simply edit them on this form. As soon as you leave the row you made changes in, the changes become permamnent"
    End If
End Sub

Private Sub cmdNew_Click()

   On Error GoTo cmdNew_Click_Error
    DoCmd.OpenForm "frmAddEditRoutes", acNormal
        

   On Error GoTo 0
   Exit Sub

cmdNew_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdNew_Click of VBA Document Form_frmRoutes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdNew_Click of VBA Document Form_frmRoutes"
End Sub

Private Sub cmdSubmit_Click()
'take whatever is in the temp_CustomerRoutes and Commit it to the CustomerRoutes table,
'then delete whatever is in the temp table.
Dim query As String
Dim rst As New ADODB.Recordset
Dim lRecs As Long

   On Error GoTo cmdSubmit_Click_Error
   
    If Me.TabCtl1.value = 1 Then
        Exit Sub
    End If
    
    If IsNull(Me.frmNewAccountToRoute.Controls!account_num) Or Me.frmNewAccountToRoute.Controls!account_num = "" Then
        'first check to see how many records there are in the temp_CustomerRoutes Table
        query = "SELECT * FROM temp_CustomerRoutes"
        rst.Open query, CurrentProject.Connection
        
        If rst.BOF And rst.EOF Then
        Call MsgBox("An account number if missing from the entry table. Please check this to make " & _
        " sure all account numbers are present and valid Then try your submission again.", _
        vbExclamation, "No Account Number")
        Exit Sub
        Else
            'do nothing
        End If
        rst.Close
    End If
    query = "INSERT INTO CustomerRoutes ( account_num, lng, lat, route_id , sequence)" & _
            " SELECT temp_CustomerRoutes.account_num, temp_CustomerRoutes.lng, temp_CustomerRoutes.lat, " & _
            " temp_CustomerRoutes.route_id, temp_CustomerRoutes.sequence" & _
            " FROM temp_CustomerRoutes;"
    CurrentProject.Connection.Execute query, lRecs
    If lRecs = 0 Then
        'could be an error
        Exit Sub
    End If
    
    query = "DELETE FROM temp_CustomerRoutes"
    CurrentProject.Connection.Execute query, lRecs
    'me.TabCtl1.form frmCustomerRoutes.Form.Refresh
    Select Case Me.TabCtl1.value
        Case 0  'New
            Me.frmNewAccountToRoute.Requery
        Case 1  'Edit
        
        Case 2  'Special
            
    End Select
    
    
   On Error GoTo 0
   Exit Sub

cmdSubmit_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdSubmit_Click of VBA Document Form_frmRoutes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdSubmit_Click of VBA Document Form_frmRoutes"
End Sub

Private Sub Form_GotFocus()

   On Error GoTo Form_GotFocus_Error
    Me.cboRoutes.Requery

   On Error GoTo 0
   Exit Sub

Form_GotFocus_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_GotFocus of VBA Document Form_frmRoutes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_GotFocus of VBA Document Form_frmRoutes"
End Sub

Private Sub Form_Load()
Dim query As String
Dim rst As New ADODB.Recordset
Dim ix As Integer

   On Error GoTo Form_Load_Error
'get who the sender was
Me.txtHelp.Visible = False
If Not IsNull(Me.OpenArgs) Then
    strSender = Me.OpenArgs
End If

'If Me.cboRoutes.ListCount > 0 Then
'    For ix = 0 To Me.cboRoutes.ListCount - 1
'        Me.cboRoutes.RemoveItem (0)
'    Next ix
'    'Me.cboRoutes.SelText = ""
'End If
Me.cboRoutes.Requery

'query = "SELECT * FROM Routes"
'rst.Open query, CurrentProject.Connection
'
'If rst.BOF And rst.EOF Then
'    Call MsgBox("No routes have been set up. Use the ""Create New Route"" button to create a new route.", vbExclamation, "No Routes")
'    Exit Sub
'End If
'
'Do While Not rst.EOF
'    Me.cboRoutes.AddItem rst.fields(1).value
'    rst.MoveNext
'Loop

    Me.cboRoutes = Me.cboRoutes.ItemData(0)
   On Error GoTo 0
   Exit Sub

Form_Load_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Load of VBA Document Form_frmRoutes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Load of VBA Document Form_frmRoutes"
End Sub

Private Sub TabCtl1_Change()
   On Error GoTo TabCtl1_Change_Error

    If Me.TabCtl1.value = 1 Then
        Me.cmdHelp.Visible = True
        Me.EditCustomerRoutesSubForm.Form.filter = "route_id=" & Me.EditCustomerRoutesSubForm.Controls!cboRoutesFilter.Column(0)
        Me.EditCustomerRoutesSubForm.Form.FilterOn = True
        Me.EditCustomerRoutesSubForm.Requery
        Me.lblChooseRoute.Visible = False
        Me.cboRoutes.Visible = False
        Me.cmdNew.Visible = False
        Me.cmdSubmit.Visible = False
    Else
        Me.cmdHelp.Visible = False
        Me.txtHelp.Visible = False
        Me.lblChooseRoute.Visible = True
        Me.cboRoutes.Visible = True
        Me.cmdNew.Visible = True
        Me.cmdSubmit.Visible = True
    End If

   On Error GoTo 0
   Exit Sub

TabCtl1_Change_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure TabCtl1_Change of VBA Document Form_frmRoutes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure TabCtl1_Change of VBA Document Form_frmRoutes"
End Sub
