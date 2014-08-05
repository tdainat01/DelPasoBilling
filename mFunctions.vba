Option Compare Database
Option Explicit

Public arguments As String
Public sCallingForm As String
Public vOpenArgs As Variant
Private GMAPI_Key As String  '//Fix here (change to Private Const) if you want wkb level Google API control
Private appIE As Object    'Hold late bound IE object
Private Const CST_NOGMAPI_KEY As String = "Sorry, could not Load Google Maps API Key"

Sub test1()
    Call TableInfo("tmpPrintMeteredBill")
End Sub
Sub main()
'Call TableInfo("customer")
'Call GetNextRecord(2200)
'Dim l As Long
'l = GetPrevRecord(2200)
Dim rst As New ADODB.Recordset
Dim query As String
Dim results() As String
Dim sAddress As String
Dim sAddresses() As String
Dim sScore As String
Dim googResultA As String
Dim GeoPoint As String

    query = "SELECT customer.account, IIf([bill_name]=" & Chr(34) & "" & Chr(34) & "," & Chr(34) & "OCCUPANT" & Chr(34) & ",[bill_name]) " & _
            " AS CustName, customer.addr1, customer.city, customer.state, customer.zip " & _
            " FROM customer WHERE customer.account in (SELECT DISTINCTROW [cust_acct] from [temp_Customers])"

'results = parseQuery(query)
rst.Open query, CurrentProject.Connection

If rst.BOF And rst.EOF Then
    Exit Sub
End If

Do While Not rst.EOF
    sAddress = sAddress & rst.Fields(2).value & "," & rst.Fields(3).value & "," & rst.Fields(4).value & "," & rst.Fields(5).value & "|"
    rst.MoveNext
Loop

    If Right(sAddress, 1) = "|" Then
        sAddress = Left(sAddress, Len(sAddress) - 1)
    End If
    
    sAddresses = Split(sAddress, "|")

'sAddress = rst.fields(2).value & "," & rst.fields(3).value & "," & rst.fields(4).value & "," & rst.fields(5).value

'sScore = GeoCode(sAddress, GeoPoint)
'googResultA = GeoCodeA(sScore)
results = GetDirections(sAddresses)

End Sub

Function GetNextRecord(curRecord As Long) As Long

Dim qry As String
Dim rst As New ADODB.Recordset

   On Error GoTo GetNextRecord_Error

    qry = "SELECT account from customer where account > " & curRecord & " order by account"
    
    rst.Open qry, CurrentProject.Connection
    
    If rst.BOF And rst.EOF Then
        GetNextRecord = -1
    End If
    
    rst.MoveFirst
    GetNextRecord = rst.Fields(0).value


   On Error GoTo 0
   Exit Function

GetNextRecord_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure GetNextRecord of Module mFunctions")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure GetNextRecord of Module mFunctions"

End Function

Function GetPrevRecord(curRecord As Long) As Long

Dim qry As String
Dim rst As New ADODB.Recordset
    
   On Error GoTo GetPrevRecord_Error

    qry = "SELECT top 1 account from customer where account < " & curRecord & " order by account desc;"
    
    rst.Open qry, CurrentProject.Connection
    
    If rst.BOF And rst.EOF Then
        GetPrevRecord = -1
        Exit Function
    End If
    
        rst.MoveFirst
    
    GetPrevRecord = rst.Fields(0).value


   On Error GoTo 0
   Exit Function

GetPrevRecord_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure GetPrevRecord of Module mFunctions")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure GetPrevRecord of Module mFunctions"

End Function

Function TableInfo(strTableName As String)
On Error GoTo TableInfoErr
   ' Purpose:   Display the field names, types, sizes and descriptions for a table.
   ' Argument:  Name of a table in the current database.
   Dim db As DAO.Database
   Dim tdf As DAO.TableDef
   Dim fld As DAO.field
   
   Set db = CurrentDb()
   Set tdf = db.TableDefs(strTableName)
   Debug.Print "FIELD NAME", "FIELD TYPE", "SIZE", "DESCRIPTION"
   Debug.Print "==========", "==========", "====", "==========="

   For Each fld In tdf.Fields
      Debug.Print fld.Name,
      Debug.Print FieldTypeName(fld),
      Debug.Print fld.Size,
      Debug.Print GetDescrip(fld)
   Next
   Debug.Print "==========", "==========", "====", "==========="

TableInfoExit:
   Set db = Nothing
   Exit Function

TableInfoErr:
   Select Case Err
   Case 3265&  'Table name invalid
      MsgBox strTableName & " table doesn't exist"
   Case Else
      Debug.Print "TableInfo() Error " & Err & ": " & Error
   End Select
   Resume TableInfoExit
End Function

Function GetDescrip(obj As Object) As String
    On Error Resume Next
    GetDescrip = obj.Properties("Description")
End Function

Function FieldTypeName(fld As DAO.field) As String
    'Purpose: Converts the numeric results of DAO Field.Type to text.
    Dim strReturn As String    'Name to return

    Select Case CLng(fld.Type) 'fld.Type is Integer, but constants are Long.
        Case dbBoolean: strReturn = "Yes/No"            ' 1
        Case dbByte: strReturn = "Byte"                 ' 2
        Case dbInteger: strReturn = "Integer"           ' 3
        Case dbLong                                     ' 4
            If (fld.Attributes And dbAutoIncrField) = 0& Then
                strReturn = "Long Integer"
            Else
                strReturn = "AutoNumber"
            End If
        Case dbCurrency: strReturn = "Currency"         ' 5
        Case dbSingle: strReturn = "Single"             ' 6
        Case dbDouble: strReturn = "Double"             ' 7
        Case dbDate: strReturn = "Date/Time"            ' 8
        Case dbBinary: strReturn = "Binary"             ' 9 (no interface)
        Case dbText                                     '10
            If (fld.Attributes And dbFixedField) = 0& Then
                strReturn = "Text"
            Else
                strReturn = "Text (fixed width)"        '(no interface)
            End If
        Case dbLongBinary: strReturn = "OLE Object"     '11
        Case dbMemo                                     '12
            If (fld.Attributes And dbHyperlinkField) = 0& Then
                strReturn = "Memo"
            Else
                strReturn = "Hyperlink"
            End If
        Case dbGUID: strReturn = "GUID"                 '15

        'Attached tables only: cannot create these in JET.
        Case dbBigInt: strReturn = "Big Integer"        '16
        Case dbVarBinary: strReturn = "VarBinary"       '17
        Case dbChar: strReturn = "Char"                 '18
        Case dbNumeric: strReturn = "Numeric"           '19
        Case dbDecimal: strReturn = "Decimal"           '20
        Case dbFloat: strReturn = "Float"               '21
        Case dbTime: strReturn = "Time"                 '22
        Case dbTimeStamp: strReturn = "Time Stamp"      '23

        'Constants for complex types don't work prior to Access 2007.
        Case 101&: strReturn = "Attachment"         'dbAttachment
        Case 102&: strReturn = "Complex Byte"       'dbComplexByte
        Case 103&: strReturn = "Complex Integer"    'dbComplexInteger
        Case 104&: strReturn = "Complex Long"       'dbComplexLong
        Case 105&: strReturn = "Complex Single"     'dbComplexSingle
        Case 106&: strReturn = "Complex Double"     'dbComplexDouble
        Case 107&: strReturn = "Complex GUID"       'dbComplexGUID
        Case 108&: strReturn = "Complex Decimal"    'dbComplexDecimal
        Case 109&: strReturn = "Complex Text"       'dbComplexText
        Case Else: strReturn = "Field type " & fld.Type & " unknown"
    End Select

    FieldTypeName = strReturn
End Function

Public Function LogError(num As Long, source As String, desc As String) As Boolean
On Error GoTo err_handler
    Dim query As String
    desc = FixQuotes(desc)
    desc = Replace(desc, "'", "''")
    query = "INSERT INTO ErrorLog(error_num,error_source,error_message,trans_date) " & _
            " VALUES(" & num & ",'" & source & "','" & desc & "',#" & Now & "#)"
    query = Replace(query, "'", "''")
    CurrentProject.Connection.Execute query
    LogError = True
    Exit Function
err_handler:
    LogError = False
End Function

Public Function FixQuotes(text As String) As String
    Dim str() As String
    Dim x As Integer
    Dim strReturn As String
   On Error GoTo FixQuotes_Error

    If InStr(text, Chr(34)) > 0 Then
        'split the string into parts, delimiting on the double quotes
        str = Split(text, Chr(34))
        If UBound(str) = 0 Then
            'raise error condition
        End If
        For x = 0 To UBound(str) - 1
            str(x) = str(x) & Chr(34) & Chr(34) & Chr(34)
        Next x
        
        'now reassemble the string
        For x = 0 To UBound(str)
            strReturn = strReturn + str(x)
        Next
        FixQuotes = strReturn
    Else
        'nothing to do
        FixQuotes = text
    End If

   On Error GoTo 0
   Exit Function

FixQuotes_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure FixQuotes of Module mFunctions")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure FixQuotes of Module mFunctions"

End Function

Public Function ValidateNumber(num As Object) As Boolean

   On Error GoTo ValidateNumber_Error

    If num Is Nothing Then
        ValidateNumber = False
        Exit Function
    End If
    
    If IsNumeric(num) Then
        ValidateNumber = True
    Else
        ValidateNumber = False
    End If

   On Error GoTo 0
   Exit Function

ValidateNumber_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure ValidateNumber of Module mFunctions")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure ValidateNumber of Module mFunctions"

End Function

Private Function ResetTotalDue() As Boolean
'dim

End Function
Public Function CheckFormStatus(Myform As String) As Boolean
    Dim objForm As AccessObject
    Dim fForm As Form
    Dim FlgLoaded As Boolean
    Dim FlgShown As Boolean
    Dim ix As Integer
    
   On Error GoTo CheckFormStatus_Error

    FlgLoaded = False
    FlgShown = False
    CheckFormStatus = False 'default to false
    'For ix = 0 To CurrentProject.AllForms.Count - 1
    For Each objForm In CurrentProject.AllForms
        'Set objForm = Forms(CurrentProject.AllForms.Item(ix).name)
        If (Trim(objForm.Name) = Trim(Myform)) Then
            If objForm.IsLoaded Then
                FlgLoaded = True
                Set fForm = Forms(objForm.Name)
                If fForm.Visible Then
                    FlgShown = True
                End If
            End If
            Exit For
        End If
    Next

'    If CurrentProject.AllForms(Myform).IsLoaded Then
'        FlgLoaded = True
'    End If
'
'    Set objForm = CurrentProject.AllForms(Myform)
'    If objForm.Visible Then
'         FlgShown = True
'    End If
    
    If FlgLoaded And FlgShown Then
        CheckFormStatus = True
        'MsgBox "Load Status: " & FlgLoaded & vbCrLf & "Show Status:" & FlgShown
    End If

   On Error GoTo 0
   Exit Function

CheckFormStatus_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure CheckFormStatus of Module mFunctions")
    'MsgBox "Error " & errNum & " (" & errMsg & ") in procedure CheckFormStatus of Module mFunctions"

End Function

Function GeoCode(sLocationData As String, ByRef GeoCodeString As String) As String

'//Dont want to open and close all day long - make once use many

   On Error GoTo GeoCode_Error

    If appIE Is Nothing Then
        CreateIEApp
        '// if = nothing now then there was an error!
        If appIE Is Nothing Then
            GeoCode = "Sorry could not launch IE"
            Exit Function
            Else
            '//do nothing
        End If
    Else
        '//do nothing!
    End If

    
    If GMAPI_Key = "" Then
    '//Get Google API key
    GMAPI_Key = GetGMAPIKey
    End If
    

    '// check we got API key OK
    If GMAPI_Key = CST_NOGMAPI_KEY Then
        GeoCode = CST_NOGMAPI_KEY
        Exit Function
    Else
        '//do nothing
    End If

    '//clearing up input data
    'sLocationData = Replace(sLocationData, ",", " ")
    sLocationData = Replace(sLocationData, " ", "+")
    sLocationData = Trim(sLocationData)


    '//Build URL for Query
    sLocationData = "http://maps.google.com/maps/geo?q=%20_" & sLocationData
    sLocationData = sLocationData & "&output=csv&key=%20"
    sLocationData = sLocationData & GMAPI_Key
Debug.Print sLocationData


    '// go to the google web service and get the raw CSV data!
    appIE.Navigate sLocationData

    Do While appIE.Busy
        'Application.StatusBar = "Contacting Google Maps API..."
        Call StatusBar("Contacting Google Maps API...")
    Loop
    
    'Application.StatusBar = False
    Call StatusBar
    On Error Resume Next
    '//we have to do a bit of prasing, luckily the formate is constant
    
    GeoCodeString = appIE.Document.Body.innerHTML
    GeoCode = Mid(GeoCodeString, InStr(GeoCodeString, ",") + 1, InStr(GeoCodeString, "/") - InStr(GeoCodeString, ",") - 2)
    

   On Error GoTo 0
   Exit Function

GeoCode_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure GeoCode of Module mFunctions")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure GeoCode of Module mFunctions"

End Function

Private Function GetGMAPIKey() As String

    On Error GoTo errHand

    If GMAPI_Key = "" Or GMAPI_Key = CST_NOGMAPI_KEY Then
        '//Load Google API Key form text file
        '// not that GMAPI_Key is public so should nly need to load file once!
        Dim iChars As Integer
        Dim iFile As Integer

        iFile = FreeFile
        Open CurrentProject.Path & "\GM_Key.txt" For Input As iFile
        iChars = LOF(iFile)
        GetGMAPIKey = Input(iChars, iFile)
        Exit Function
    Else
        GetGMAPIKey = GMAPI_Key
    End If
    Exit Function

errHand:
    GetGMAPIKey = CST_NOGMAPI_KEY
End Function

Private Function CreateIEApp()
'//Create Internet Explorer application Object
    On Error GoTo errHand
    Set appIE = CreateObject("InternetExplorer.Application")
    Exit Function

errHand:
    appIE = Nothing
End Function

Public Function CloseIEApp() As Byte
'// Made public to use if user likes
    On Error GoTo errHand
    appIE.Quit
    CloseIEApp = 1
    Exit Function
errHand:
    CloseIEApp = 0
End Function

Sub Auto_close()
'// I keep the appIE open for the life of the work book.
'//If you have a auto close sub already, add this call to it!
   On Error GoTo Auto_close_Error

CloseIEApp

   On Error GoTo 0
   Exit Sub

Auto_close_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Auto_close of Module mFunctions")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Auto_close of Module mFunctions"
End Sub


Public Function GeoCodeA(sScore As String) As String

   On Error GoTo GeoCodeA_Error

sScore = Left(sScore, 1)

Select Case sScore
Case 0
GeoCodeA = "Unknown location"
Case 1
GeoCodeA = "Country level"
Case 2
GeoCodeA = "Region level"
Case 3
GeoCodeA = "Sub-region level"
Case 4
GeoCodeA = "Town/Village level"
Case 5
GeoCodeA = "Post Code level"
Case 6
GeoCodeA = "Street level"
Case 7
GeoCodeA = "Intersection level"
Case 8
GeoCodeA = "Address level"
Case Else
GeoCodeA = "Not Found"
End Select

   On Error GoTo 0
   Exit Function

GeoCodeA_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure GeoCodeA of Module mFunctions")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure GeoCodeA of Module mFunctions"

End Function

Public Function GetDirections(sAddresses() As String) As String()
Dim WayPoint As String
Dim query As String
Dim rst As New ADODB.Recordset
Dim sCompany As String
Dim sURL As String
Dim xmlResult As String
Dim x As Integer

'//Dont want to open and close all day long - make once use many

    If appIE Is Nothing Then
        CreateIEApp
        '// if = nothing now then there was an error!
        If appIE Is Nothing Then
            'GeoCode = "Sorry could not launch IE"
            Exit Function
            Else
            '//do nothing
        End If
    Else
        '//do nothing!
    End If

    
    If GMAPI_Key = "" Then
    '//Get Google API key
    GMAPI_Key = GetGMAPIKey
    End If
    

    '// check we got API key OK
    If GMAPI_Key = CST_NOGMAPI_KEY Then
        'GeoCode = CST_NOGMAPI_KEY
        Exit Function
    Else
        '//do nothing
    End If
    
    WayPoint = ""
    
    For x = 0 To UBound(sAddresses)
        WayPoint = WayPoint & Replace(sAddresses(x), " ", "+")
        WayPoint = Trim(WayPoint)
        WayPoint = WayPoint & "|"
    Next x
    
    If Right(WayPoint, 1) = "|" Then
        WayPoint = Left(WayPoint, Len(WayPoint) - 1)
    End If

    '//Build URL for Query
    query = "SELECT * FROM Company" 'we're assuming only one company exists on file
    rst.Open query, CurrentProject.Connection
    
    If rst.BOF And rst.EOF Then
        Exit Function
    End If
    
    sCompany = IIf(IsNull(rst.Fields(2).value), " ", rst.Fields(2).value) & "," & _
                IIf(IsNull(rst.Fields(3).value), " ", rst.Fields(3).value) & "," & _
                IIf(IsNull(rst.Fields(4).value), " ", rst.Fields(4).value) & "," & _
                IIf(IsNull(rst.Fields(5).value), " ", rst.Fields(5).value) & "," & _
                IIf(IsNull(rst.Fields(6).value), " ", rst.Fields(6).value) & ","
    
    sCompany = Replace(sCompany, " ", "+")
    sCompany = Trim(sCompany)
    
    
    sURL = "http://maps.googleapis.com/maps/api/directions/xml?origin=" & sCompany & "&destination=" & sCompany & _
        "&waypoints=" & WayPoint & "&sensor=false"

Debug.Print sURL

    '// go to the google web service and get the raw CSV data!
    appIE.Navigate sURL

    Do While appIE.Busy
        'Application.StatusBar = "Contacting Google Maps API..."
        Call StatusBar("Contacting Google Maps API...")
    Loop
    
    'Application.StatusBar = False
    Call StatusBar
    On Error Resume Next

    xmlResult = appIE.Document.Body.innerText

End Function


Sub StatusBar(Optional msg As Variant)
Dim Temp As Variant

' if the Msg variable is omitted or is empty, return the control of the status bar to Access

If Not IsMissing(msg) Then
 If msg <> "" Then
  Temp = SysCmd(acSysCmdSetStatus, msg)
 Else
  Temp = SysCmd(acSysCmdClearStatus)
 End If
Else
  Temp = SysCmd(acSysCmdClearStatus)
End If
End Sub

'''The Directions service limits are:
'''http://code.google.com/apis/maps/documentation/directions/#Limits
'''Use of the Google Directions API is subject to a query limit of 2,500
'''directions requests per day. Individual directions requests may
'''contain up to 8 intermediate waypoints in the request. Google Maps
'''Premier customers may query up to 100,000 directions requests per day,
'''with up to 23 waypoints allowed in each request.
Private Function ParseXML(xmlString) As String
Dim xmlDoc As MSXML2.DOMDocument
Dim fSuccess As Boolean
Dim xmlStatus As MSXML2.IXMLDOMElement
Dim xmlRoute As MSXML2.IXMLDOMNodeList
Dim xmlChild As MSXML2.IXMLDOMNode
Dim ix As Integer

   On Error GoTo ParseXML_Error

Set xmlDoc = New MSXML2.DOMDocument
' Load the  XML from string, without validating it. Wait
' for the load to finish before proceeding.
xmlDoc.async = False
xmlDoc.validateOnParse = False    'we trust google will supply us with valid xml

fSuccess = xmlDoc.loadXML(xmlString)

If Not fSuccess Then
    MsgBox xmlDoc.parseError.reason, vbOKOnly, "Error Loading XML"
    Exit Function
End If

Set xmlRoute = xmlDoc.getElementsByTagName("route")

For Each xmlChild In xmlRoute.Item(0).childNodes
    If xmlChild.baseName = "leg" Then
        For ix = 0 To xmlChild.childNodes.Length - 1
            If xmlChild.childNodes.Item(ix).baseName = "start_address" Then
                Debug.Print "Start At: " & xmlChild.childNodes.Item(ix).text
            ElseIf xmlChild.childNodes.Item(ix).baseName = "end_address" Then
                Debug.Print "End At: " & xmlChild.childNodes.Item(ix).text
            End If
        Next
        'Set xmlLeg = xmlChild.selectNodes("leg")
    End If
Next

   On Error GoTo 0
   Exit Function

ParseXML_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure ParseXML of Module mFunctions")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure ParseXML of Module mFunctions"

End Function

Public Function FormatZip(zipcode As Control)

  'Exit if a null is passed to the function.
  If IsNull(zipcode) Then
    Exit Function
  End If
  If IsThereAlpha(zipcode) Then
    MsgBox "Your ZIP Code Contains Letters"
    Exit Function
  Else
    'Strip unwanted characters, leaving only numbers.
    zipcode = ZStrip(zipcode, "-")
    zipcode = ZStrip(zipcode, " ")
    zipcode = ZStrip(zipcode, ")")
    zipcode = ZStrip(zipcode, "(")

    'Reformat the character string.
    Select Case Len(zipcode)
      Case 5
        Screen.ActiveControl = Format(zipcode, "@@@@@")
      Case 9
        Screen.ActiveControl = Format(zipcode, "@@@@@-@@@@")
      Case Else
        MsgBox "This is not a valid ZIP Code."
        Screen.ActiveControl = zipcode
    End Select
  End If
End Function

Function ZStrip(InZip, StripZip)
  Do While InStr(InZip, StripZip)
    InZip = Mid(InZip, 1, InStr(InZip, StripZip) - 1) & Mid _
    (InZip, InStr(InZip, StripZip) + 1)
  Loop
  ZStrip = InZip
End Function

Function IsThereAlpha(zipcode) As Integer
   Dim Pos, a_char$, Clean
   Pos = 1
   Clean = 0
   While (Pos <= Len(zipcode) And (Clean = 0))
      a_char$ = Mid(zipcode, Pos, 1)
      If a_char$ >= "0" And a_char$ <= "9" Then
         Clean = 0
      Else
         If a_char$ <> "-" Then Clean = 1
      End If
      Pos = Pos + 1
   Wend
   IsThereAlpha = Clean
End Function

Public Function GetGUID() As String
    GetGUID = Mid$(CreateObject("Scriptlet.TypeLib").Guid, 2, 36)
End Function

Public Function CalculateCharges(lstAccount As String) As clsState()
Dim rstRecur As New ADODB.Recordset
Dim rstServ As New ADODB.Recordset
Dim rst As New ADODB.Recordset
Dim rstUC As New ADODB.Recordset
Dim strQuery As String
Dim account As Long
Dim query As String
Dim charged As Single
Dim used As Single
Dim strTransaction As String
Dim lRecs As Long
Dim arrAccount() As String
Dim clsAcct() As clsState
Dim ctr As Long
Dim ix As Long

'## Provide input to get range of accounts

'This query can get any account, metered or otherwise
   On Error GoTo CalculateCharges_Error

ctr = 0
ix = 0
'lstAccount = "60001,60002,60003"
arrAccount = Split(lstAccount, ",")
ReDim clsAcct(UBound(arrAccount))

strQuery = "SELECT customer.* FROM customer" & _
            " WHERE (((customer.account) in (" & lstAccount & ")) AND ((customer.status)<>'I') AND ((customer.term_date) Is Null Or " & _
            " (customer.term_date)=#1/1/1900#));"

rst.Open strQuery, CurrentProject.Connection, adOpenDynamic, adLockPessimistic

Do While Not rst.EOF    'outer loop
    Set clsAcct(ctr) = New clsState
    charged = 0 'reset the charged value
    account = IIf(IsNull(rst.Fields("account").value), 0, rst.Fields("account").value)
    If account = 0 Then
        Err.Raise vbObjectError + 5001, "cmdChargeMeter_Click of Form_BillingMenu", "account cannot be zero"
    End If
    clsAcct(ctr).account = account
    'Is this a metered account?
    If rst.Fields("current_read").value > 0 And rst.Fields("previous_read").value > 0 Then
        If rst.Fields("current_read").value < rst.Fields("previous_read").value Then
        'roll over
        'Need an update query to handle this
        'strRollOverAccounts = strRollOverAccounts & "'" & rst.Fields(0).Value & "'"
        Else
            'Calculate gallons uses
            used = rst.Fields("current_read").value - rst.Fields("previous_read").value
            
            'If the fields unit of measure is a G then do a converstion
            If Left(rst.Fields("unit_measure").value, 1) = "G" Then
                used = used * GALSTOCUFEET
            Else
                'used = used
            End If
            'usage charge is calculated per 100 cu feet
            lRecs = 0
            rstUC.Open "SELECT [Amount] from MeterRates", CurrentProject.Connection
            If rstUC.BOF And rstUC.EOF Then
                'throw an error
                Exit Function
            End If
            clsAcct(ctr).UsageCharge = (rst.Fields("gal_cub_used").value / 100) * CSng(rstUC.Fields(0).value)     'multiply by usage charge
            rstUC.Close
        End If

        'Calculate the service line charge
        query = "SELECT CustomerServiceConnection.account, ServiceConnections.Amount, ServiceConnections.Description" & _
                " FROM CustomerServiceConnection INNER JOIN ServiceConnections ON CustomerServiceConnection.service_id = " & _
                " ServiceConnections.service_id WHERE (((CustomerServiceConnection.account)=" & account & "));"

        rstServ.Open query, CurrentProject.Connection
        If rstServ.BOF And rstServ.EOF Then
            'no service charge
        Else
            'take only the first service charge as techinically there can only be one
            'add the usage charge to the service charge
            clsAcct(ctr).Service = rstServ.Fields(1).value
            clsAcct(ctr).ServiceDescription = rstServ.Fields(2).value
        End If
        
        rstServ.Close '###

        'now add other recurring charges such as Readiness to Serve Charge and any other fees (e.g. fire protection)
        query = "SELECT RatesAndCharges.account, RecurringCharges.charge_amount, RecurringCharges.charge_description" & _
                " FROM RecurringCharges INNER JOIN RatesAndCharges ON RecurringCharges.charge_id = RatesAndCharges.recurring_charge_id" & _
                " WHERE (((RatesAndCharges.account)=" & account & "));"

        rstRecur.Open query, CurrentProject.Connection
        
        If rstRecur.BOF And rstRecur.EOF Then
            'nothing to do - no recurring service charge
        Else
            'could be many recurring charges - take them all
            Do While Not rstRecur.EOF
                clsAcct(ctr).AddRecurring rstRecur.Fields(0).value, rstRecur.Fields(1).value, rstRecur.Fields(2).value
                rstRecur.MoveNext
            Loop
        End If
        
        rstRecur.Close '###
    Else
    'Calculate the service line charge
    query = "SELECT CustomerServiceConnection.account, ServiceConnections.Amount, ServiceConnections.Description" & _
            " FROM CustomerServiceConnection INNER JOIN ServiceConnections ON CustomerServiceConnection.service_id = " & _
            " ServiceConnections.service_id WHERE (((CustomerServiceConnection.account)=" & account & "));"

    rstServ.Open query, CurrentProject.Connection ', adOpenDynamic, adLockOptimistic
    If rstServ.BOF And rstServ.EOF Then
        'no service charge
    Else
        'take only the first service charge as techinically there can only be one
        'add the usage charge to the service charge
        clsAcct(ctr).Service = rstServ.Fields(1).value
        clsAcct(ctr).ServiceDescription = rstServ.Fields(2).value
    End If
    
    rstServ.Close
        
    'next calculate any recurring charges such as fire service, maint charges etc.
    query = "SELECT RatesAndCharges.account, RecurringCharges.charge_amount, RecurringCharges.charge_description" & _
            " FROM RecurringCharges INNER JOIN RatesAndCharges ON RecurringCharges.charge_id = RatesAndCharges.recurring_charge_id" & _
            " WHERE (((RatesAndCharges.account)=" & account & "));"

    rstRecur.Open query, CurrentProject.Connection
    strTransaction = ""
    
    If rstRecur.BOF And rstRecur.EOF Then
        'nothing to do - no recurring service charge
    Else
        'could be many recurring charges - take them all
        Do While Not rstRecur.EOF
            clsAcct(ctr).AddRecurring rstRecur.Fields(0).value, rstRecur.Fields(1).value, rstRecur.Fields(2).value
            rstRecur.MoveNext
        Loop
    End If
    rstRecur.Close
    
    End If
    rst.MoveNext
    ctr = ctr + 1
Loop

    'do something with the account data and the charged amount
    'Debug.Print ""
    CalculateCharges = clsAcct
    
   On Error GoTo 0
   Exit Function

CalculateCharges_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure CalculateCharges of VBA Document Report_rptTotalCharged")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure CalculateCharges of VBA Document Report_rptTotalCharged"
End Function

Public Sub WriteToFile(fPathAndName As String, sWhat As String)
' Declare a FileSystemObject.
Dim fso As FileSystemObject

' Create a FileSystemObject.
   On Error GoTo WriteToFile_Error

Set fso = New FileSystemObject

' Declare a TextStream.
Dim stream As TextStream

' Create a TextStream.
If FileExists(fPathAndName) Then
    Set stream = fso.OpenTextFile(fPathAndName, ForAppending)
Else
    Set stream = fso.CreateTextFile(fPathAndName, True)
End If

stream.WriteLine (sWhat)
stream.Close

   On Error GoTo 0
   Exit Sub

WriteToFile_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    'Call LogError(errNum, errSource, errMsg & " in procedure WriteToFile of Module mFunctions")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure WriteToFile of Module mFunctions"

End Sub

Function FileExists(ByVal strFile As String, Optional bFindFolders As Boolean) As Boolean
    'Purpose:   Return True if the file exists, even if it is hidden.
    'Arguments: strFile: File name to look for. Current directory searched if no path included.
    '           bFindFolders. If strFile is a folder, FileExists() returns False unless this argument is True.
    'Note:      Does not look inside subdirectories for the file.
    'Author:    Allen Browne. http://allenbrowne.com June, 2006.
    Dim lngAttributes As Long

    'Include read-only files, hidden files, system files.
    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)

    If bFindFolders Then
        lngAttributes = (lngAttributes Or vbDirectory) 'Include folders as well.
    Else
        'Strip any trailing slash, so Dir does not look inside the folder.
        Do While Right$(strFile, 1) = "\"
            strFile = Left$(strFile, Len(strFile) - 1)
        Loop
    End If

    'If Dir() returns something, the file exists.
    On Error Resume Next
    FileExists = (Len(Dir(strFile, lngAttributes)) > 0)
End Function

Public Function FillCustomer(rst As ADODB.Recordset) As clsCustomer

Dim Cust As New clsCustomer

   On Error GoTo FillCustomer_Error

        If rst.BOF And rst.EOF Then
            'nothing to do
            FillCustomer = Nothing
            Exit Function
        End If
        
        'hopefully there is only 1 row. But regardless - we're only taking the first row anyway
        
        Cust.account = IIf(IsNull(rst.Fields("account").value), 0, rst.Fields("account").value)
        Cust.addr1 = IIf(IsNull(rst.Fields("addr1").value), "", rst.Fields("addr1").value)
        Cust.Backflow = IIf(IsNull(rst.Fields("backflow").value), False, rst.Fields("backflow").value)
        Cust.BillName = IIf(IsNull(rst.Fields("bill_name").value), "", rst.Fields("bill_name").value)
        Cust.CareOf = IIf(IsNull(rst.Fields("care_of").value), "", rst.Fields("care_of").value)
        Cust.city = IIf(IsNull(rst.Fields("city").value), "", rst.Fields("city").value)
        Cust.comment = IIf(IsNull(rst.Fields("comment").value), "", rst.Fields("comment").value)
        Cust.CurrentDate = IIf(IsNull(rst.Fields("current_date").value), Now, rst.Fields("current_date").value)
        Cust.CurrentDue = IIf(IsNull(rst.Fields("current_due").value), 0, rst.Fields("current_due").value)
        Cust.CurrentRead = IIf(IsNull(rst.Fields("current_read").value), 0, rst.Fields("current_read").value)
        Cust.Cycle = IIf(IsNull(rst.Fields("cycle").value), 0, rst.Fields("cycle").value)
        Cust.Deposit = IIf(IsNull(rst.Fields("deposit").value), 0, rst.Fields("deposit").value)
        Cust.FireSize = IIf(IsNull(rst.Fields("fire_size").value), 0, rst.Fields("fire_size").value)
        Cust.GalCubUsed = IIf(IsNull(rst.Fields("gal_cub_used").value), 0, rst.Fields("gal_cub_used").value)
        Cust.group = IIf(IsNull(rst.Fields("group").value), 0, rst.Fields("group").value)
        Cust.lien = IIf(IsNull(rst.Fields("lien").value), 0, rst.Fields("lien").value)
        Cust.mastpar = IIf(IsNull(rst.Fields("mastpar").value), 0, rst.Fields("mastpar").value)
        Cust.MeterNumber = IIf(IsNull(rst.Fields("meter_number").value), 0, rst.Fields("meter_number").value)
        Cust.MeterSite = IIf(IsNull(rst.Fields("meter_site").value), 0, rst.Fields("meter_site").value)
        Cust.MeterSize = IIf(IsNull(rst.Fields("meter_size").value), 0, rst.Fields("meter_size").value)
        Cust.MfgCode = IIf(IsNull(rst.Fields("mfg_code").value), 0, rst.Fields("mfg_code").value)
        Cust.Name = IIf(IsNull(rst.Fields("name").value), "Unknown", rst.Fields("name").value)
        Cust.OutTown = IIf(IsNull(rst.Fields("out_town").value), False, rst.Fields("out_town").value)
        Cust.PastDue = IIf(IsNull(rst.Fields("past_due").value), 0, rst.Fields("past_due").value)
        Cust.PhyAddress = IIf(IsNull(rst.Fields("phy_address").value), "", rst.Fields("phy_address").value)
        Cust.PrevBalance = IIf(IsNull(rst.Fields("prev_balance").value), 0, rst.Fields("prev_balance").value)
        Cust.PreviousDate = IIf(IsNull(rst.Fields("previous_date").value), Now, rst.Fields("previous_date").value)
        Cust.PreviousRead = IIf(IsNull(rst.Fields("previous_read").value), 0, rst.Fields("previous_read").value)
        Cust.PropertyUse = IIf(IsNull(rst.Fields("property_use").value), "", rst.Fields("property_use").value)
        Cust.RateCode = IIf(IsNull(rst.Fields("rate_code").value), "", rst.Fields("rate_code").value)
        Cust.ServiceDiscon = IIf(IsNull(rst.Fields("service_discon").value), False, rst.Fields("service_discon").value)
        Cust.SpecialCharge = IIf(IsNull(rst.Fields("special_charge").value), 0, rst.Fields("special_charge").value)
        Cust.SpecialCredit = IIf(IsNull(rst.Fields("special_credit").value), 0, rst.Fields("special_credit").value)
        Cust.SpecialDescription = IIf(IsNull(rst.Fields("special_description").value), "", rst.Fields("special_description").value)
        Cust.StartDate = IIf(IsNull(rst.Fields("start_date").value), Now, rst.Fields("start_date").value)
        Cust.state = IIf(IsNull(rst.Fields("state").value), "", rst.Fields("state").value)
        Cust.Status = IIf(IsNull(rst.Fields("status").value), "", rst.Fields("status").value)
        Cust.TermDate = IIf(IsNull(rst.Fields("term_date").value), Now, rst.Fields("term_date").value)
        Cust.TotalDue = IIf(IsNull(rst.Fields("total_due").value), 0, rst.Fields("total_due").value)
        Cust.Transloc = IIf(IsNull(rst.Fields("trans_loc").value), 0, rst.Fields("trans_loc").value)
        Cust.UnitMeasure = IIf(IsNull(rst.Fields("unit_measure").value), "", rst.Fields("unit_measure").value)
        Cust.UseCharge = IIf(IsNull(rst.Fields("use_charge").value), 0, rst.Fields("use_charge").value)
        Cust.Zip = IIf(IsNull(rst.Fields("zip").value), "", rst.Fields("zip").value)
        
        Set FillCustomer = Cust

   On Error GoTo 0
   Exit Function

FillCustomer_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure FillCustomer of Module mFunctions")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure FillCustomer of Module mFunctions"

End Function

Public Function FillMoney(rst As ADODB.Recordset) As clsMoney

Dim Mon As New clsMoney

   On Error GoTo FillMoney_Error

    If rst.BOF And rst.EOF Then
        'nothing to do
        Set FillMoney = Nothing
        Exit Function
    End If
    
    Mon.AccountNumber = IIf(IsNull(rst.Fields("account_number").value), 0, rst.Fields("account_number").value)
    Mon.Amount = IIf(IsNull(rst.Fields("amount").value), 0, rst.Fields("amount").value)
    Mon.BehindMe = IIf(IsNull(rst.Fields("behind_me").value), 0, rst.Fields("behind_me").value)
    Mon.Code = IIf(IsNull(rst.Fields("code").value), "", rst.Fields("code").value)
    Mon.Day = IIf(IsNull(rst.Fields("m_day").value), 0, rst.Fields("m_day").value)
    Mon.Month = IIf(IsNull(rst.Fields("m_month").value), 0, rst.Fields("m_month").value)
    Mon.posted = IIf(IsNull(rst.Fields("posted").value), "", rst.Fields("posted").value)
    Mon.transaction = IIf(IsNull(rst.Fields("transaction").value), "", rst.Fields("transaction").value)
    Mon.TransDate = IIf(IsNull(rst.Fields("trans_date").value), Now, rst.Fields("trans_date").value)
    Mon.Year = IIf(IsNull(rst.Fields("m_year").value), 0, rst.Fields("m_year").value)

    Set FillMoney = Mon

   On Error GoTo 0
   Exit Function

FillMoney_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure FillMoney of Module mFunctions")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure FillMoney of Module mFunctions"

End Function

Public Sub OutPutTextBill(args As String, fPathAndName As String)
Dim rst As New ADODB.Recordset
Dim fso As FileSystemObject
Dim stream As TextStream
Dim calc_due As Double
Dim strArgs As String
Dim s() As String
Dim lblDateTime As String
Dim query As String
Dim strTemp As String
Dim sModCustName As String
Dim cn() As String

Dim pad1 As String * 34         'Space between left margine and Prev_Balance
Dim Label1_Pad As String * 16   'Top labels for Prev_Balance Current Charge, Fire and Maint
Dim TextVal_Pad As String * 9  'Fixed field for Value fields
Dim Label2_Pad As String * 9   'Fixed field for tear-off labels
Dim Left_Pad As String * 9     'Fixed field for left padding
Dim AddrField As String * 52    'Fixed field for customer data
Dim ShortAddrField As String * 19 'Fixed field for customer data on tear-off sheet
Const Space1 As Integer = 6     'space between Prev_Balance and value
Const Space2 As Integer = 10     'space between field value and tear-off label
Const TearOffPad As Integer = 3
Dim tmpStr As String
Const strSpace As String = " "

   On Error GoTo OutPutTextBill_Error

If args = "" Or Len(args) < 1 Then
    Call MsgBox("There was a problem executing the OutPutTextBill method. No arguments specified. Please contact the developer for a fix", _
        vbOKOnly + vbCritical, "Critical Error")
    Exit Sub
End If

strArgs = args
If InStr(strArgs, "|") > 0 Then
    s = Split(strArgs, "|")
    's(0) is the query, s(1) = start date, s(2) = end date s(3) = year
    query = s(0)
    lblDateTime = s(1) & "-" & s(2) & " " & s(3)
Else
    'we have a problem with a malformed argument
End If

Set fso = New FileSystemObject

If FileExists(fPathAndName) Then
    Set stream = fso.OpenTextFile(fPathAndName, ForAppending)
Else
    Set stream = fso.CreateTextFile(fPathAndName, True)
End If

rst.Open query, CurrentProject.Connection

If rst.BOF And rst.EOF Then
    'no work to do
    'alert user
    
    Exit Sub
Else
    Do While Not rst.EOF
        'write two lines
        stream.WriteBlankLines (3)
        
        'add padding before Prev_Balance
        pad1 = ""
        Label1_Pad = "PREV BALANCE"
        Label2_Pad = "PREV BAL"
        tmpStr = FormatNumber(rst.Fields("calc_prevbal").value, 2)
        TextVal_Pad = PadLeft(tmpStr, Len(TextVal_Pad), strSpace)
        stream.WriteLine (pad1 & Label1_Pad & TextVal_Pad & Space(TearOffPad) & Label2_Pad & RTrim(TextVal_Pad))
        
        calc_due = CSng(IIf(IsNull(rst.Fields("calculated_due").value) Or IsEmpty(rst.Fields("calculated_due").value), 0, rst.Fields("calculated_due").value))
        Label1_Pad = "CURRENT CHARGE"
        Label2_Pad = "CUR. CHG"
        tmpStr = CStr(FormatNumber(calc_due, 2))
        TextVal_Pad = PadLeft(tmpStr, Len(TextVal_Pad), strSpace)
        stream.WriteLine (pad1 & Label1_Pad & TextVal_Pad & Space(TearOffPad) & Label2_Pad & RTrim(TextVal_Pad))
        
        pad1 = "  " & CStr(rst.Fields("phy_address").value)
        Label1_Pad = IIf(rst.Fields("fire_charge").value <= 0 Or IsEmpty(rst.Fields("fire_charge").value) Or IsNull(rst.Fields("fire_charge").value), " ", "FIRE PROTECT.")
        Label2_Pad = IIf(rst.Fields("fire_charge").value <= 0 Or IsEmpty(rst.Fields("fire_charge").value) Or IsNull(rst.Fields("fire_charge").value), " ", "FIRE")
        tmpStr = IIf(rst.Fields("fire_charge").value <= 0 Or IsEmpty(rst.Fields("fire_charge").value) Or IsNull(rst.Fields("fire_charge").value), "", FormatNumber(rst.Fields("fire_charge").value, 2))
        TextVal_Pad = PadLeft(tmpStr, Len(TextVal_Pad), strSpace)
        stream.WriteLine (pad1 & Label1_Pad & TextVal_Pad & Space(TearOffPad) & Label2_Pad & RTrim(TextVal_Pad))
        
        On Error Resume Next
        Err.Clear
        pad1 = "  CUBIC FEET USED:      " & CStr(FormatNumber(rst.Fields("calc_gals_used").value / 7.48052, 0))
        If Err.Number <> 0 Then
            pad1 = ""
        End If
        On Error GoTo OutPutTextBill_Error
        
        Label1_Pad = "SYST MAINT CHG"
        Label2_Pad = "SMC"
        tmpStr = CStr(IIf(IsNull(rst.Fields("smc_charge").value), 0#, FormatNumber(rst.Fields("smc_charge").value, 2)))
        TextVal_Pad = PadLeft(tmpStr, Len(TextVal_Pad), strSpace)
        stream.WriteLine (pad1 & Label1_Pad & TextVal_Pad & Space(TearOffPad) & Label2_Pad & RTrim(TextVal_Pad))
        
        pad1 = "  " & lblDateTime
        stream.WriteLine (RTrim(pad1))
        
        pad1 = "  ACCT# " & CStr(rst.Fields("account").value)
        stream.WriteLine (RTrim(pad1))
        
        '2 blank lines, then the total
        'stream.WriteBlankLines (2)
        
        'Print the total amount on it's own line
        pad1 = ""
        Label1_Pad = ""
        Label2_Pad = ""
        tmpStr = FormatNumber(rst.Fields("total_due").value, 2)
        TextVal_Pad = PadLeft(tmpStr, Len(TextVal_Pad), strSpace)
        stream.WriteLine (pad1 & Label1_Pad & TextVal_Pad & Space(9 + TearOffPad) & RTrim(TextVal_Pad))
                
        'Customer Data.
        strTemp = rst.Fields("CustName").value
        If InStr(strTemp, "C/O") > 0 Then
            'we have to write the customer data over two lines.
            sModCustName = Replace(strTemp, "C/O", "|C/O")
        ElseIf InStr(strTemp, "ATTN") > 0 Then
            sModCustName = Replace(strTemp, "ATTN", "|ATTN")
        Else
            sModCustName = "|" & strTemp
        End If

        AddrField = ""
        ShortAddrField = "ACCT# " & CStr(rst.Fields("account").value)
        stream.WriteLine (AddrField & Space(Space2) & RTrim(ShortAddrField))
        
        cn = Split(sModCustName, "|")
        AddrField = cn(0)
        ShortAddrField = cn(0)
        Left_Pad = ""
        stream.WriteLine (Left_Pad & AddrField & Trim(ShortAddrField))
        
        AddrField = cn(1)
        ShortAddrField = cn(1)
        Left_Pad = ""
        stream.WriteLine (Left_Pad & AddrField & Trim(ShortAddrField))
        
        AddrField = rst.Fields("address").value
        ShortAddrField = rst.Fields("address").value
        Left_Pad = ""
        stream.WriteLine (Left_Pad & AddrField & Trim(ShortAddrField))
        
        AddrField = rst.Fields("city").value & " " & rst.Fields("state").value & " " & rst.Fields("zip").value
        ShortAddrField = rst.Fields("city").value & " " & rst.Fields("state").value
        Left_Pad = ""
        stream.WriteLine (Left_Pad & AddrField & Trim(ShortAddrField))
        
        'Last Line
        stream.WriteBlankLines (1)
        AddrField = ""
        Left_Pad = ""
        stream.WriteLine (Left_Pad & AddrField & Trim(lblDateTime))
        stream.WriteBlankLines (4)
    rst.MoveNext
    DoEvents
    Loop
End If

stream.Close

   On Error GoTo 0
   Exit Sub

OutPutTextBill_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure OutPutTextBill of Module mFunctions")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure OutPutTextBill of Module mFunctions"

End Sub
Public Sub PrintFile(sFile As String, sPort As String)
Dim cmd As String
Dim rst As New ADODB.Recordset
Dim query As String
Dim strCommandCodes As String
Dim escOff As String
Dim escOn As String
Dim strArray() As String
Dim x As Integer

   On Error GoTo PrintFile_Error

    'Open the settings and get the escape character sequence to send.
    query = "SELECT SettingsName, SettingsValue from Settings WHERE SettingsName like 'EscapeCodeOff'"
    rst.Open query, CurrentProject.Connection
    If rst.BOF And rst.EOF Then
        'nothing to do
        strCommandCodes = ""
    Else
        If IsNull(rst.Fields("SettingsValue").value) Or Len(rst.Fields("SettingsValue").value) < 1 Then
            strCommandCodes = ""
        Else
            strCommandCodes = rst.Fields("SettingsValue").value
        End If
    End If
    
    If Len(strCommandCodes) < 1 Then
        'nothing to do
    Else
        strArray = Split(strCommandCodes, ",")
        For x = 0 To UBound(strArray)
            escOff = escOff & "chr(" & strArray(x) & ") & "
        Next x
        
        escOff = Trim(escOff)
        
        If Right(escOff, 1) = "&" Then
            escOff = Trim(Left(escOff, Len(escOff) - 1))
        End If
    End If

    rst.Close
    
    query = "SELECT SettingsName, SettingsValue from Settings WHERE SettingsName like 'EscapeCodeOn'"
    rst.Open query, CurrentProject.Connection
    If rst.BOF And rst.EOF Then
        'nothing to do
        strCommandCodes = ""
    Else
        If IsNull(rst.Fields("SettingsValue").value) Or Len(rst.Fields("SettingsValue").value) < 1 Then
            strCommandCodes = ""
        Else
            strCommandCodes = rst.Fields("SettingsValue").value
        End If
    End If
    
        If Len(strCommandCodes) < 1 Then
        'nothing to do
    Else
        strArray = Split(strCommandCodes, ",")
        For x = 0 To UBound(strArray)
            escOn = escOn & "chr(" & strArray(x) & ") & "
        Next x
        
        escOn = Trim(escOn)
        
        If Right(escOn, 1) = "&" Then
            escOn = Trim(Left(escOn, Len(escOn) - 1))
        End If
    End If
    
    rst.Close
    
    If Len(escOff) > 5 And Len(escOn) > 5 Then
        cmd = "cmd /C type " & escOff & " >" & sPort
        Call Shell(cmd)
    End If
    
    cmd = "cmd /C type " & sFile & " >" & sPort
    Call Shell(cmd)
    
    If Len(escOff) > 5 And Len(escOn) > 5 Then
        cmd = "cmd /C type " & escOn & " >" & sPort
        Call Shell(cmd)
    End If
   On Error GoTo 0
   Exit Sub

PrintFile_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure PrintFile of Module mFunctions")
    'MsgBox "Error " & errNum & " (" & errMsg & ") in procedure PrintFile of Module mFunctions"

End Sub

'Public Sub PrintFile(sWhat As String, sPort As String)
'Dim ff As Integer
'   On Error GoTo PrintFile_Error
'ff = FreeFile
'
'    Open sPort For Output As ff
'    'Open Application.Printer.Port For Output As ff
'    Print #ff, sWhat
'    Close #ff
'
'   On Error GoTo 0
'   Exit Sub
'
'PrintFile_Error:
'Dim errNum As Long
'Dim errMsg As String
'Dim errSource As String
'errNum = Err.Number
'errSource = Err.source
'errMsg = Err.Description
'
'    'Call LogError(errNum, errSource, errMsg & " in procedure PrintFile of Module mFunctions")
'    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure PrintFile of Module mFunctions"
'End Sub

Public Function ReadAsciiFile(sFileName As String) As String

    Dim iFileNum As Integer
    Dim sBuf As String
    Dim msg As String
    
    ' does the file exist?  simpleminded test:
'    If Len(Dir$(sFileName)) = 0 Then
'        Exit Function
'    End If
    If Not FileExists(sFileName) Then
        'alert user and exit
        Call MsgBox("The file: " & sFileName & " does not exist. Aborting printing.", vbOKOnly + vbCritical, "Aborting")
        Exit Function
    End If
    
    iFileNum = FreeFile()
    Open sFileName For Input As iFileNum

    Do While Not EOF(iFileNum)
        Line Input #iFileNum, sBuf
        msg = msg & sBuf
    Loop

    ' close the file
    Close iFileNum
    ReadAsciiFile = msg
End Function

Function GenGuid() As String
Dim TypeLib As Object
Dim Guid As String
   On Error GoTo GenGuid_Error

    Set TypeLib = CreateObject("Scriptlet.TypeLib")
    Guid = TypeLib.Guid
    ' format is {24DD18D4-C902-497F-A64B-28B2FA741661}
    Guid = Replace(Guid, "{", "")
    Guid = Replace(Guid, "}", "")
    Guid = Replace(Guid, "-", "")
    GenGuid = Guid

   On Error GoTo 0
   Exit Function

GenGuid_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    'Call LogError(errNum, errSource, errMsg & " in procedure GenGuid of Module mFunctions")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure GenGuid of Module mFunctions"
End Function

Function CreateFileName() As String
Dim strFile As String

   On Error GoTo CreateFileName_Error

    strFile = Format(Now(), "yyyy-MM-dd hh:mm:ss")
    strFile = Replace(strFile, "-", "")
    strFile = Replace(strFile, " ", "")
    strFile = Replace(strFile, ":", "")
    CreateFileName = strFile

   On Error GoTo 0
   Exit Function

CreateFileName_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    'Call LogError(errNum, errSource, errMsg & " in procedure CreateFileName of Module mFunctions")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure CreateFileName of Module mFunctions"

End Function
