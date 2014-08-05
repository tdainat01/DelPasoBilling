Option Compare Database
Option Explicit

Private Sub cmdChange_Click()
Dim query As String
Dim lRecs As Long

   On Error GoTo cmdChange_Click_Error
'    If IsNull(Me.txtOldDate) Or IsEmpty(Me.txtOldDate) Or IsNull(Me.txtNewDate) _
'        Or IsEmpty(Me.txtNewDate) Then
'        Exit Sub
'    End If
'    query = "INSERT INTO System (system_date) VALUES (#" & Me.txtNewDate & "#)"
'    CurrentProject.Connection.Execute query, lRecs
'    If lRecs <= 0 Then
'        Err.Raise vbObjectError + 5005, "cmdChange_Click of Form frmChangeDate", _
'            "No records inserted for " & query
'    End If
'
'    Call MsgBox("The date has been changed.", vbInformation, "Date Changed")
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdChange_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdChange_Click of VBA Document Form_frmChangeDate")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdChange_Click of VBA Document Form_frmChangeDate"
End Sub

Private Sub cmdExit_Click()

   On Error GoTo cmdExit_Click_Error
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdExit_Click of VBA Document Form_frmChangeDate")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdExit_Click of VBA Document Form_frmChangeDate"
End Sub

Private Sub Form_Load()
Dim rst As New ADODB.Recordset
Dim query As String

   On Error GoTo Form_Load_Error
    'look into the systems table. If there is a value there
    'then use it to populate the old date. Otherwise use Now
'    query = "SELECT top 1 System.system_date AS FirstOfDate" & _
'            " FROM System order by System.auto_id DESC;"
'
'    rst.Open query, CurrentProject.Connection
'    If rst.BOF And rst.EOF Then
'        Me.txtOldDate = Now
'    Else
'        Me.txtOldDate = rst.Fields(0).value
'    End If

   On Error GoTo 0
   Exit Sub

Form_Load_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Load of VBA Document Form_frmChangeDate")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Load of VBA Document Form_frmChangeDate"
End Sub
