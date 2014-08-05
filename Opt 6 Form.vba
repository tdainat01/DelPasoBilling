Option Compare Database
Option Explicit
Dim fOpt1 As Boolean

'------------------------------------------------------------
' cmdExit_Click
'
'------------------------------------------------------------
Private Sub cmdExit_Click()
   On Error GoTo cmdExit_Click_Error

    DoCmd.OpenForm "DPM Main Menu", acNormal, "", "", , acNormal
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

    Call LogError(errNum, errSource, errMsg & " in procedure cmdExit_Click of VBA Document Form_Opt 6 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdExit_Click of VBA Document Form_Opt 6 Form"

End Sub


Private Sub cmdPrintRoll_Click()
   On Error GoTo cmdPrintRoll_Click_Error

    fOpt1 = True
    DoCmd.OpenForm "PrintOutRoll", acNormal, "", "", , acNormal, "fOpt1.Selected"

   On Error GoTo 0
   Exit Sub

cmdPrintRoll_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrintRoll_Click of VBA Document Form_Opt 6 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrintRoll_Click of VBA Document Form_Opt 6 Form"

End Sub

Private Sub cmdPrintTotals_Click()
   On Error GoTo cmdPrintTotals_Click_Error

    DoCmd.OpenForm "DPM Main Menu"
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdPrintTotals_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrintTotals_Click of VBA Document Form_Opt 6 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrintTotals_Click of VBA Document Form_Opt 6 Form"
End Sub

Private Sub cmdPrintTotalsPrinter_Click()
   On Error GoTo cmdPrintTotalsPrinter_Click_Error

    DoCmd.OpenForm "DPM Main Menu"
    DoCmd.Close acForm, Me.Name, acSaveYes

   On Error GoTo 0
   Exit Sub

cmdPrintTotalsPrinter_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrintTotalsPrinter_Click of VBA Document Form_Opt 6 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrintTotalsPrinter_Click of VBA Document Form_Opt 6 Form"
End Sub

Private Sub cmdPrintUnknown_Click()
   On Error GoTo cmdPrintUnknown_Click_Error

    DoCmd.OpenForm "PrintUnknownOwners", acNormal, "", "", , acNormal

   On Error GoTo 0
   Exit Sub

cmdPrintUnknown_Click_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure cmdPrintUnknown_Click of VBA Document Form_Opt 6 Form")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure cmdPrintUnknown_Click of VBA Document Form_Opt 6 Form"

End Sub
