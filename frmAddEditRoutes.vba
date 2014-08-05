Option Compare Database
Option Explicit

Private Sub cmdExit_Click()

   On Error GoTo cmdExit_Click_Error
    If Me.Dirty Then Me.Dirty = False
    DoCmd.Close acForm, Me.Name, acSaveYes
    If CurrentProject.AllForms("frmRoutes").IsLoaded Then
        Forms("frmRoutes").SetFocus
        'Forms("frmRoutes").Controls("cboRoutes").Requery
    End If
   On Error GoTo 0
   Exit Sub

cmdExit_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdExit_Click of VBA Document Form_frmAddEditRoutes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdExit_Click of VBA Document Form_frmAddEditRoutes"
End Sub

Private Sub cmdSave_Click()

   On Error GoTo cmdSave_Click_Error
    If Me.Dirty Then Me.Dirty = False

   On Error GoTo 0
   Exit Sub

cmdSave_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdSave_Click of VBA Document Form_frmAddEditRoutes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdSave_Click of VBA Document Form_frmAddEditRoutes"
End Sub

