Option Compare Database

'local variable(s) to hold property value(s)
Private mvarKey As Long 'local copy
Private mvarAccount As Long 'local copy
Private mvarMissingNum As Long 'local copy

'Public Property Get colMissingNum() As colMissingNumClass
'    If mvarcolMissingNum Is Nothing Then
'        Set mvarcolMissingNum = New colMissingNumClass
'    End If
'
'
'    Set colMissingNum = mvarcolMissingNum
'End Property

'Public Property Set colMissingNum(vData As colMissingNumClass)
'    Set mvarcolMissingNum = vData
'End Property

Private Sub Class_Terminate()
    Set mvarcolMissingNum = Nothing
End Sub

'Public Sub Add(Acct As Long, Key As Long, MNum As Long)
'End Sub

Public Property Let MissingNum(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MissingNum = 5
    mvarMissingNum = vData
End Property


Public Property Get MissingNum() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MissingNum
    MissingNum = mvarMissingNum
End Property


Public Property Let account(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Account = 5
    mvarAccount = vData
End Property


Public Property Get account() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Account
    account = mvarAccount
End Property

Public Property Let Key(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vData
End Property


Public Property Get Key() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    Key = mvarKey
End Property

