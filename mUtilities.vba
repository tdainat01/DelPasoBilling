Option Compare Database
Option Explicit
Private Declare Function a2Ku_apigettime Lib "winmm.dll" _
Alias "timeGetTime" () As Long
Dim lngstartingtime As Long

Sub ResetCurrent()

Dim rst As New ADODB.Recordset
Dim strQuery As String

   On Error GoTo ResetCurrent_Error

strQuery = "SELECT current_read from customer"
rst.Open strQuery, CurrentProject.Connection, adOpenDynamic, adLockPessimistic

Do While Not rst.EOF
    rst.Fields(0).value = Int((10000 * Rnd()) + 1)
    rst.MoveNext
Loop

rst.Close

   On Error GoTo 0
   Exit Sub

ResetCurrent_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure ResetCurrent of Module mUtilities")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure ResetCurrent of Module mUtilities"

End Sub

Sub ResetPrevious()

Dim rst As New ADODB.Recordset
Dim strQuery As String

   On Error GoTo ResetPrevious_Error

strQuery = "SELECT previous_read from customer"
rst.Open strQuery, CurrentProject.Connection, adOpenDynamic, adLockPessimistic

Do While Not rst.EOF
    rst.Fields(0).value = Int((3000 * Rnd()) + 1)
    rst.MoveNext
Loop

rst.Close

   On Error GoTo 0
   Exit Sub

ResetPrevious_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure ResetPrevious of Module mUtilities")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure ResetPrevious of Module mUtilities"

End Sub

Sub ResetPreviousDate()

Dim rst As New ADODB.Recordset
Dim strQuery As String

strQuery = "SELECT previous_date from customer"
rst.Open strQuery, CurrentProject.Connection, adOpenDynamic, adLockPessimistic

Do While Not rst.EOF
    rst.Fields(0).value = DateAdd("m", -1, Now)
    rst.MoveNext
Loop

rst.Close

End Sub

Sub ResetCurrentDate()

Dim rst As New ADODB.Recordset
Dim strQuery As String
Dim lRecsAffected As Long

   On Error GoTo ResetCurrentDate_Error

strQuery = "update customer set current_date =#" & Now & "# where current_date is null or current_date = ''"
CurrentProject.Connection.Execute strQuery, lRecsAffected

MsgBox lRecsAffected & " records were updated.", vbInformation + vbOKOnly, "Update completed"

   On Error GoTo 0
   Exit Sub

ResetCurrentDate_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure ResetCurrentDate of Module mUtilities")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure ResetCurrentDate of Module mUtilities"

End Sub

Public Function parseQuery(qry As String) As String()
Dim sTemp1 As String
Dim sTemp2 As String
Dim pos1 As Integer
Dim pos2 As Integer
Dim arr() As String
Dim retArr() As String
Dim ctr As Integer
Dim ix As Integer

   On Error GoTo parseQuery_Error

pos1 = InStr(qry, "FROM")
sTemp1 = Trim(Mid(qry, 7, pos1 - 8))
arr = Split(sTemp1, " ")

For ix = 0 To UBound(arr)
    If arr(ix) = "" Or arr(ix) = "AS" Then
        If arr(ix) = "AS" Then
            arr(ix + 1) = "." + arr(ix + 1)
        End If
    Else
        If InStr(arr(ix), ".") > 0 Then
            pos1 = InStr(arr(ix), ".")
            sTemp2 = sTemp2 & Mid(arr(ix), pos1 + 1)
        End If
    End If
Next

retArr = Split(sTemp2, ",")
parseQuery = retArr

   On Error GoTo 0
   Exit Function

parseQuery_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure parseQuery of Module mUtilities")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure parseQuery of Module mUtilities"

End Function

Private Sub TestSequence()
Dim rst As New ADODB.Recordset
Dim query As String
Dim vArray As Variant
Dim RouteID As Long
Dim lRecs As Long
Dim myDict As New colMissingNumClass
   
   On Error GoTo TestSequence_Error

RouteID = 1
lRecs = GetNextNumber(RouteID)
'    Set myDict = FindMissingNumbers(RouteID)
'
'    For lRecs = 0 To myDict.Count - 1
'        query = "update CustomerRoutes set sequence = " & myDict.Item(CStr(lRecs)).MissingNum & " where account_num = " & myDict.Item(CStr(lRecs)).account
'        CurrentProject.Connection.Execute query
'    Next lRecs
    
   On Error GoTo 0
   Exit Sub

TestSequence_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure TestSequence of Module mUtilities")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure TestSequence of Module mUtilities"
    
End Sub

Function SequenceArray(vArray As Variant)
    Dim ctr As Integer
    Dim ix As Integer
    Dim firstNum As Integer
    Dim lastNum As Integer
    Dim numCount As Integer
    
    'now we need to resequence all the existing routes
   On Error GoTo SequenceArray_Error
    'fill in any number's missing in the sequence
    For ctr = 0 To UBound(vArray, 2)
        For ix = 1 To UBound(vArray, 2) - 1
            If vArray(1, ix) - vArray(1, ctr) <> 1 Then
                'we have our missing number
            End If
        Next ix
    Next ctr

   On Error GoTo 0
   Exit Function

SequenceArray_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure SequenceArray of Module mUtilities")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure SequenceArray of Module mUtilities"

End Function

Sub a2kuStartClock()
    lngstartingtime = a2Ku_apigettime()
End Sub
Function a2kuEndClock()
    a2kuEndClock = a2Ku_apigettime() - lngstartingtime
End Function

Function FindMissingNumbers(RouteID As Long) As colMissingNumClass
Dim rst As New ADODB.Recordset
Dim strSQL As String
Dim loend As Long, hiend As Long
Dim reccount As Long, recnum As Long
Dim rechold As Long
Dim iMissingAccount As String
Dim iMissingNum As String
Dim ctr As Long
Dim myCol As New colMissingNumClass
'Start the clock
'a2kuStartClock

   On Error GoTo FindMissingNumbers_Error

strSQL = "SELECT Min(CustomerRoutes.sequence) AS MinOfOrderID, Max(CustomerRoutes.sequence) AS MaxOfOrderID, Count(CustomerRoutes.sequence) AS CountOfOrderID, [maxoforderid]-[minoforderid]+1 AS Expr1" _
    & " FROM CustomerRoutes where CustomerRoutes.route_id =" & RouteID
    rst.Open strSQL, CurrentProject.Connection
loend = IIf(IsNull(rst!MinOfOrderID), 0, rst!MinOfOrderID)
hiend = IIf(IsNull(rst!MaxOfOrderID), 0, rst!MaxOfOrderID)
reccount = IIf(IsNull(rst!CountOfOrderID), 0, rst!CountOfOrderID)
recnum = IIf(IsNull(rst!Expr1), 0, rst!Expr1)
rechold = loend
strSQL = "SELECT account_num, sequence FROM CustomerRoutes where CustomerRoutes.route_id = " & RouteID & " ORDER BY sequence;"
rst.Close

    rst.Open strSQL, CurrentProject.Connection
    rechold = 1
    ctr = 0
Do While rechold <= reccount
    If rst!sequence <> rechold Then
        myCol.Add ctr, rst.Fields("account_num").value, rechold, CStr(ctr)
        ctr = ctr + 1
        'rechold = rst.fields("sequence").value
   End If
        rechold = rechold + 1
        rst.MoveNext
Loop
rst.Close
'Stop the clock and display the results
'msg = msg & vbCrLf & "This procedure executed in: " & a2kuEndClock & " milliseconds"
'MsgBox msg, vbInformation + vbOKOnly, "Find Missing OrderID's"
Set FindMissingNumbers = myCol
Set myCol = Nothing

   On Error GoTo 0
   Exit Function

FindMissingNumbers_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure FindMissingNumbers of Module mUtilities")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure FindMissingNumbers of Module mUtilities"
End Function

Function GetNextNumber(RouteID As Long) As Long
Dim rst As New ADODB.Recordset
Dim strSQL As String
Dim loend As Long
Dim hiend As Long
Dim reccount As Long
Dim recnum As Long
Dim rechold As Long
Dim loend2 As Long
Dim hiend2 As Long
Dim reccount2 As Long
Dim recnum2 As Long
Dim rechold2 As Long
Dim iMissingAccount As String
Dim iMissingNum As String
Dim ctr As Long
Dim myCol As New colMissingNumClass
Dim lRecs As Long
Dim query As String

   On Error GoTo GetNextNumber_Error
    'first get if there are any missing numbers. if one is found
    Set myCol = FindMissingNumbers(RouteID)
    
    For lRecs = 0 To myCol.Count - 1
        GetNextNumber = myCol.Item(CStr(lRecs)).account
        Exit Function
    Next
    'otherwise get the next number for that route
    query = "SELECT Min(CustomerRoutes.sequence) AS MinOfOrderID, Max(CustomerRoutes.sequence) AS MaxOfOrderID, Count(CustomerRoutes.sequence) AS CountOfOrderID, [maxoforderid]-[minoforderid]+1 AS Expr1" _
    & " FROM CustomerRoutes where CustomerRoutes.route_id =" & RouteID
    
    rst.Open query, CurrentProject.Connection
    If rst.BOF And rst.EOF Then
        GetNextNumber = 1
        Exit Function
    End If
    
    loend = rst!MinOfOrderID
    hiend = rst!MaxOfOrderID
    reccount = rst!CountOfOrderID
    recnum = rst!Expr1
    rechold = hiend + 1
    rst.Close

    query = "SELECT Min(sequence) AS MinOfOrderID, Max(sequence) AS MaxOfOrderID, Count(sequence) AS CountOfOrderID, [maxoforderid]-[minoforderid]+1 AS Expr1" _
            & " FROM temp_CustomerRoutes where route_id =" & RouteID
    rst.Open query, CurrentProject.Connection
    
    If rst.BOF And rst.EOF Then
        GetNextNumber = rechold
        Exit Function
    Else
        loend2 = IIf(IsNull(rst!MinOfOrderID), 0, rst!MinOfOrderID)
        hiend2 = IIf(IsNull(rst!MaxOfOrderID), 0, rst!MaxOfOrderID)
        reccount2 = IIf(IsNull(rst!CountOfOrderID), 0, rst!CountOfOrderID)
        recnum2 = IIf(IsNull(rst!Expr1), 0, rst!Expr1)
        rechold2 = hiend2 + 1
        rst.Close
    End If
    
    If rechold2 > rechold Then
        GetNextNumber = rechold2
    Else
        GetNextNumber = rechold
    End If
        
   On Error GoTo 0
   Exit Function

GetNextNumber_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure GetNextNumber of Module mUtilities")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure GetNextNumber of Module mUtilities"
End Function

Private Sub ClassBuilder()

Dim MyString As String
Dim MyNumber As Long
Dim ff As Long
Dim sGetProp As String
Dim sGet As String
Dim sLet As String
Dim sPropName As String
Dim sEndProp As String
Dim sLetProp As String
Dim VariableName As String
Dim VariableType As String
Dim myMatches As MatchCollection
Dim myRegExp As RegExp
Dim ctr As Long
Dim sGetProperty As String
Dim sLetProperty As String
Dim sMethods As String

   On Error GoTo ClassBuilder_Error

ctr = 0

ff = FreeFile

Open "c:\temp\moneyvars.txt" For Input As ff ' Open file for input.

Do While Not EOF(ff) ' Loop until end of file.
    Input #ff, MyString
        
    sGetProp = "Public Property Get {methodname}() As [methodtype]"
    sGet = "{account} = [mvarAccount]"
    sEndProp = "End Property"
    sLetProp = "Public Property Let {methodname}(ByVal vData As [methodtype])"
    sLet = "{methodname} = vData"
        
        
    Set myRegExp = New RegExp
    myRegExp.IgnoreCase = True
    myRegExp.Multiline = True
    myRegExp.Pattern = "m[a-zA-Z]*"
    Set myMatches = myRegExp.Execute(MyString)
    If myMatches.Count >= 1 Then
        VariableName = myMatches(0).value
    Else
        VariableName = "UnKnownVariable" & ctr
        ctr = ctr + 1
    End If
    
    myRegExp.Pattern = "As (Long|String|Date|Boolean|Double|Single|Currency)"
    Set myMatches = myRegExp.Execute(MyString)
    If myMatches.Count >= 1 Then
        VariableType = myMatches(0).value
    Else
        VariableType = "Variant"
    End If

    sPropName = Mid(VariableName, 2)
    sGetProp = Replace(sGetProp, "{methodname}() As [methodtype]", sPropName & "() " & VariableType)
    sGet = Replace(sGet, "{account} = [mvarAccount]", sPropName & " = " & VariableName)
    sLetProp = Replace(sLetProp, "{methodname}(ByVal vData As [methodtype]", sPropName & "(ByVal vData " & VariableType)
    sLet = Replace(sLet, "{methodname}", VariableName)
    
    'putting it all together
    sGetProperty = sGetProp & vbCrLf & vbTab & sGet & vbCrLf & sEndProp & vbCrLf
    sLetProperty = sLetProp & vbCrLf & vbTab & sLet & vbCrLf & sEndProp & vbCrLf
    sMethods = sMethods & sGetProperty & vbCrLf & sLetProperty
    
Loop

Close ff ' Close file.

Call WriteToFile("c:\temp\Class.txt", sMethods)
Debug.Print "Output written"

   On Error GoTo 0
   Exit Sub

ClassBuilder_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure ClassBuilder of Module mUtilities")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure ClassBuilder of Module mUtilities"

End Sub

'Save the code for all modules to files in currentDatabaseDir\Code
Public Function ExportAllClassesandModules()

Dim Name As String
Dim WasOpen As Boolean
Dim Last As Integer
Dim I As Integer
Dim TopDir As String, Path As String, FileName As String
Dim F As Long                          'File for saving code
Dim LineCount As Long                  'Line count of current module

I = InStrRev(CurrentDb.Name, "\")
TopDir = VBA.Left(CurrentDb.Name, I - 1)
Path = TopDir & "\" & "Code"           'Path where the files will be written

If (Dir(Path, vbDirectory) = "") Then
  MkDir Path                           'Ensure this exists
End If

'--- SAVE THE STANDARD MODULES CODE ---

Last = Application.CurrentProject.AllModules.Count - 1

For I = 0 To Last
  Name = CurrentProject.AllModules(I).Name
  WasOpen = True                       'Assume already open

  If Not CurrentProject.AllModules(I).IsLoaded Then
    WasOpen = False                    'Not currently open
    DoCmd.OpenModule Name              'So open it
  End If

  LineCount = Access.Modules(Name).CountOfLines
  FileName = Path & "\" & Name & ".vba"

  If (Dir(FileName) <> "") Then
    Kill FileName                      'Delete previous version
  End If

  'Save current version
  F = FreeFile
  Open FileName For Output Access Write As #F
  Print #F, Access.Modules(Name).Lines(1, LineCount)
  Close #F

  If Not WasOpen Then
    DoCmd.Close acModule, Name         'It wasn't open, so close it again
  End If
Next

'--- SAVE FORMS MODULES CODE ---

Last = Application.CurrentProject.AllForms.Count - 1

For I = 0 To Last
  Name = CurrentProject.AllForms(I).Name
  WasOpen = True

  If Not CurrentProject.AllForms(I).IsLoaded Then
    WasOpen = False
    DoCmd.OpenForm Name, acDesign
  End If

  LineCount = Access.Forms(Name).Module.CountOfLines
  FileName = Path & "\" & Name & ".vba"

  If (Dir(FileName) <> "") Then
    Kill FileName
  End If

  F = FreeFile
  Open FileName For Output Access Write As #F
  Print #F, Access.Forms(Name).Module.Lines(1, LineCount)
  Close #F

  If Not WasOpen Then
    DoCmd.Close acForm, Name
  End If
Next
MsgBox "Created source files in " & Path
End Function
