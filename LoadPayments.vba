Option Compare Database

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

    Call LogError(errNum, errSource, errMsg & " in procedure Form_KeyPress of VBA Document Form_LoadPayments")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_KeyPress of VBA Document Form_LoadPayments"

End Sub

Private Sub Form_Open(Cancel As Integer)
   On Error GoTo Form_Open_Error

    Me.txtDate = Format(Now, "mm/dd/yyyy")
    Me.txtAccountNumber.SetFocus

   On Error GoTo 0
   Exit Sub

Form_Open_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Open of VBA Document Form_LoadPayments")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Open of VBA Document Form_LoadPayments"
End Sub

Private Sub txtCorrect_Change()
   On Error GoTo txtCorrect_Change_Error

    If txtCorrect.text = "" Then
        Exit Sub
    End If
    Call txtCorrect_KeyPress(Asc(txtCorrect.text))

   On Error GoTo 0
   Exit Sub

txtCorrect_Change_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure txtCorrect_Change of VBA Document Form_LoadPayments")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure txtCorrect_Change of VBA Document Form_LoadPayments"
End Sub

Private Sub txtCorrect_KeyPress(KeyAscii As Integer)
Dim strCharacter As String
    ' Convert ANSI value to character string.
   On Error GoTo txtCorrect_KeyPress_Error

    strCharacter = Chr(KeyAscii)
    ' Convert character to upper case, then to ANSI value.
    KeyAscii = Asc(UCase(strCharacter))
    If KeyAscii = 81 Or KeyAscii = 27 Then
        DoCmd.Close acForm, Me.Name, acSaveYes
    End If
    If KeyAscii = 89 Then
        'Perform some action
        If Me.txtAccountNumber <> "" And IsNumeric(Me.txtAccountNumber) Then
            'MsgBox "Loading Payment now ... ", vbOKOnly + vbInformation, "Loading Payment"
            'Insert into the money table
            Dim I As Integer
            I = WriteMoney
            If I = 0 Then
            'no records were written
            ElseIf I = -1 Then
                'an error occurred which has already neeb logged
            End If
                Me.txtAccountNumber = ""
                Me.txtAmount = ""
                Me.txtTransCode = ""
                Me.txtCorrect = ""
                Me.txtAccountNumber.SetFocus
        End If
    End If

   On Error GoTo 0
   Exit Sub

txtCorrect_KeyPress_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure txtCorrect_KeyPress of VBA Document Form_LoadPayments")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure txtCorrect_KeyPress of VBA Document Form_LoadPayments"
End Sub

Private Function WriteMoney() As Long
    Dim ResultString As String
    Dim myMatches As MatchCollection
    Dim myRegExp As RegExp
   On Error GoTo WriteMoney_Error

    Set myRegExp = New RegExp
    myRegExp.Pattern = "[+-]?[0-9]{1,3}(?:,?[0-9]{3})*(?:\.[0-9]{2})?"
    Set myMatches = myRegExp.Execute(Me.txtAmount)
    If myMatches.Count >= 1 Then
        ResultString = myMatches(0).value
    Else
        ResultString = ""
    End If
    
    If ResultString <> "" Then
        Me.txtAmount = ResultString
    End If
    
    Dim qry As String
    Dim lRecs As Long
    qry = "INSERT Into [Money] (m_month,m_day,m_year,account_number,amount," & _
          "[transaction],code,posted,behind_me,trans_date)" & _
        " VALUES ('" & Format(Now, "mm") & "','" & Format(Now, "dd") & "','" & Format(Now, "yyyy") & "','" & _
        Me.txtAccountNumber & "'," & Me.txtAmount & ",'" & Replace(Me.txtTransNote, "'", "''") & "','" & Me.txtTransCode & _
        "','N',0,#" & Now & "#);"
    CurrentProject.Connection.Execute qry, lRecs
    WriteMoney = lRecs
   On Error GoTo 0
   Exit Function

WriteMoney_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description
Dim msg As String
Dim F As Boolean
    F = LogError(Err.Number, Err.source, Err.Description)
    If F Then
        msg = ". This error has been logged."
    Else
        msg = ". This error has NOT been logged."
    End If
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure WriteMoney of VBA Document Form_LoadPayments" & msg
    WriteMoney = -1
End Function
