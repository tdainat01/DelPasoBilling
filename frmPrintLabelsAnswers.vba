Option Compare Database
Option Explicit

Private Sub cmdExit_Click()
   On Error GoTo cmdExit_Click_Error

    DoCmd.OpenForm "DPM Main Menu", acNormal
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdExit_Click of VBA Document Form_frmPrintLabelsAnswers")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdExit_Click of VBA Document Form_frmPrintLabelsAnswers"
End Sub

Private Sub cmdGo_Click()
    'Call cmdGo_Enter
    'Determine how many labels across were selected.
    'call that report and set its data source property
    
    'check to make sure that txtLabels is numeric
   On Error GoTo cmdGo_Click_Error

    If Not IsNumeric(Me.txtLabels) Then
        MsgBox "The number of labels across must be a numeric value between 1 and 5", vbCritical + vbOKOnly, "Warning"
        Exit Sub
    End If

    If CInt(Me.txtLabels) < 1 Or CInt(Me.txtLabels) > 5 Then
        MsgBox "The number of labels across must be a numeric value between 1 and 5", vbCritical + vbOKOnly, "Warning"
        Exit Sub
    End If
    
    Dim iLabelsAcross As Integer
    iLabelsAcross = CInt(Me.txtLabels)
    Dim reportQuery As String
    Dim rpt As Report
    
    reportQuery = "SELECT customer.account, IIf([name]='UNKNOWN','OCCUPANT',[name]) AS CustName, customer.phy_address, customer.city, " & _
        " customer.state, customer.zip FROM customer " & _
        " WHERE (((customer.account)>= " & Me.cmbMin.value & " And (customer.account)<= " & Me.cmdMax.value & "));"
    
    Select Case iLabelsAcross
        Case Is = 1
            DoCmd.OpenReport "rptAdLabels1", acViewDesign
            Set rpt = Reports![rptAdLabels1]
            rpt.RecordSource = reportQuery
            DoCmd.Close acReport, rpt.Name, acSaveYes
            DoCmd.OpenReport "rptAdLabels1", acViewPreview
        Case Is = 2
            DoCmd.OpenReport "rptAdLabels2", acViewDesign
            Set rpt = Reports![rptAdLabels2]
            rpt.RecordSource = reportQuery
            DoCmd.Close acReport, rpt.Name, acSaveYes
            DoCmd.OpenReport "rptAdLabels2", acViewPreview
        Case Is = 3
            DoCmd.OpenReport "rptAdLabels3", acViewDesign
            Set rpt = Reports![rptAdLabels3]
            rpt.RecordSource = reportQuery
            DoCmd.Close acReport, rpt.Name, acSaveYes
            DoCmd.OpenReport "rptAdLabels3", acViewPreview
        Case Is = 4
            DoCmd.OpenReport "rptAdLabels4", acViewDesign
            Set rpt = Reports![rptAdLabels4]
            rpt.RecordSource = reportQuery
            DoCmd.Close acReport, rpt.Name, acSaveYes
            DoCmd.OpenReport "rptAdLabels4", acViewPreview
        Case Is = 5
            DoCmd.OpenReport "rptAdLabels5", acViewDesign
            Set rpt = Reports![rptAdLabels5]
            rpt.RecordSource = reportQuery
            DoCmd.Close acReport, rpt.Name, acSaveYes
            DoCmd.OpenReport "rptAdLabels5", acViewPreview
    End Select

   On Error GoTo 0
   Exit Sub

cmdGo_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdGo_Click of VBA Document Form_frmPrintLabelsAnswers")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdGo_Click of VBA Document Form_frmPrintLabelsAnswers"
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Dim strCharacter As String
    ' Convert ANSI value to character string.
   On Error GoTo Form_KeyPress_Error

    strCharacter = Chr(KeyAscii)
    ' Convert character to upper case, then to ANSI value.
    KeyAscii = Asc(UCase(strCharacter))
    If KeyAscii = 27 Then
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

    Call LogError(errNum, errSource, errMsg & " in procedure Form_KeyPress of VBA Document Form_frmPrintLabelsAnswers")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_KeyPress of VBA Document Form_frmPrintLabelsAnswers"
End Sub

Private Sub Form_Load()
    Dim rst As New ADODB.Recordset
    Dim min As String
    Dim max As String
    
   On Error GoTo Form_Load_Error

    Me.txtLabels = "1"
    
    rst.Open "SELECT Min(customer.account) AS MinOfaccount FROM customer;", CurrentProject.Connection
    min = rst.Fields(0).value
    rst.Close
    
    rst.Open "SELECT Max(customer.account) AS MinOfaccount FROM customer;", CurrentProject.Connection
    max = rst.Fields(0).value
    rst.Close

    Me.cmbMin.value = min
    Me.cmdMax.value = max

   On Error GoTo 0
   Exit Sub

Form_Load_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Load of VBA Document Form_frmPrintLabelsAnswers")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Load of VBA Document Form_frmPrintLabelsAnswers"

End Sub

Private Sub ClearAll()
    Dim rpt As Report
    Dim ctrl As Access.Control
   On Error GoTo ClearAll_Error

    DoCmd.OpenReport "rptAdLabels", acViewDesign
    Set rpt = Reports![rptAdLabels]
    Dim x As Integer
    
    'loop through all controls and delete them
    For x = 0 To Me.Controls.Count - 1
        Set ctrl = rpt.Controls.Item(x)
        
        rpt.Remove (ctrl)
        rpt.Controls(x).Remove
    Next

    'To save the modification to the report,  uncomment the following line of code:
    DoCmd.Close acReport, rpt.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

ClearAll_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure ClearAll of VBA Document Form_frmPrintLabelsAnswers")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure ClearAll of VBA Document Form_frmPrintLabelsAnswers"

End Sub
