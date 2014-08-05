Option Compare Database
Option Explicit

Dim opt As String

Private Sub cmbMax_Enter()
   On Error GoTo cmbMax_Enter_Error

    Call cmdPrint_Click

   On Error GoTo 0
   Exit Sub

cmbMax_Enter_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmbMax_Enter of VBA Document Form_PrintMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmbMax_Enter of VBA Document Form_PrintMenu"
End Sub

Private Sub cmbMax_LostFocus()
   On Error GoTo cmbMax_LostFocus_Error

    Call cmdPrint_Click

   On Error GoTo 0
   Exit Sub

cmbMax_LostFocus_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmbMax_LostFocus of VBA Document Form_PrintMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmbMax_LostFocus of VBA Document Form_PrintMenu"
End Sub

Private Sub Form_Open(Cancel As Integer)
Dim rst As New ADODB.Recordset
Dim min As String
Dim max As String

Me.cmbMin = DateAdd("d", -1, Now)
Me.cmbMax = Now
'Me.txtYear = Format(Now(), "yyyy")
End Sub
'------------------------------------------------------------
' cmdGoBack_Click
'
'------------------------------------------------------------
Private Sub cmdGoBack_Click()
   On Error GoTo cmdGoBack_Click_Error

    DoCmd.OpenForm "PaymentMenu", acNormal, "", "", , acNormal
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdGoBack_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdGoBack_Click of VBA Document Form_PrintMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdGoBack_Click of VBA Document Form_PrintMenu"

End Sub


'------------------------------------------------------------
' cmdPrint_Click
'
'------------------------------------------------------------
Private Sub cmdPrint_Click()
Dim crit As String

   On Error GoTo cmdPrint_Click_Error

vOpenArgs = Me.OpenArgs

    Dim sStartDate As String
    Dim sEndDate As String
    
    'validate start and end dates
    If IsDate(Me.cmbMin.value) Then
        sStartDate = Me.cmbMin.value
    Else
        'alert user and exit
        MsgBox Me.cmbMin.value & " is not a valid date. Please enter a valid date and try again", vbExclamation + vbOKOnly, "Error"
        Exit Sub
    End If
    
    If IsDate(Me.cmbMax.value) Then
        sEndDate = Me.cmbMax.value
    Else
        MsgBox Me.cmbMax.value & " is not a valid date. Please enter a valid date and try again", vbExclamation + vbOKOnly, "Error"
        Exit Sub
    End If

    Select Case vOpenArgs
    Case Is = 1
        DoCmd.OpenReport "rptPaymentEdit", acViewReport, , "trans_date BETWEEN #" & sStartDate & "# AND #" & sEndDate & "#"
        DoCmd.Close acForm, Me.Name, acSaveYes
    Case Is = 2
        DoCmd.OpenReport "rptPaymentEdit2", acViewReport, , "trans_date BETWEEN #" & sStartDate & "# AND #" & sEndDate & "#"
        DoCmd.Close acForm, Me.Name, acSaveYes
    End Select

   On Error GoTo 0
   Exit Sub

cmdPrint_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrint_Click of VBA Document Form_PrintMenu")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrint_Click of VBA Document Form_PrintMenu"

End Sub


