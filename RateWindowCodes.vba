Option Compare Database

Private Sub cmdPrint_Click()
   On Error GoTo cmdPrint_Click_Error

    DoCmd.OpenReport "rptRates", acViewReport

   On Error GoTo 0
   Exit Sub

cmdPrint_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrint_Click of VBA Document Form_RateWindowCodes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrint_Click of VBA Document Form_RateWindowCodes"
End Sub

Private Sub cmdQuit_Click()
   On Error GoTo cmdQuit_Click_Error

    DoCmd.OpenForm "DPM Main Menu", acNormal
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdQuit_Click of VBA Document Form_RateWindowCodes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdQuit_Click of VBA Document Form_RateWindowCodes"
End Sub

Private Sub Form_Close()
    'check to see if there is any unsaved data
   On Error GoTo Form_Close_Error

    If Me.Dirty Then
        Dim result As VbMsgBoxResult
        result = MsgBox("You have unsaved changed. Click OK to save these now, or click Cancel to exit without saving", _
        vbOKCancel + vbQuestion, "Save?")
        If result = 1 Then
            'save
            RunCommand acCmdSaveRecord
        End If
    End If

   On Error GoTo 0
   Exit Sub

Form_Close_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Close of VBA Document Form_RateWindowCodes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Close of VBA Document Form_RateWindowCodes"
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

    Call LogError(errNum, errSource, errMsg & " in procedure Form_KeyPress of VBA Document Form_RateWindowCodes")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_KeyPress of VBA Document Form_RateWindowCodes"
End Sub
