Option Compare Database
Option Explicit

'set the private variables
'local variable to hold collection
Private mCol As Collection  'to hold the Rates and Charges values
Private mColRecurring As clsRecurring

Private mvarAccount As Long
Private mvarPropertyUse As String
Private mvarService As String
Private mServiceDescription As String
Private mUsageCharge As Single
Private mFlag As Boolean

Public Sub Add(Key As Long, checked As Boolean)
    mCol.Add checked, CStr(Key)
End Sub

Public Property Get Item(vntIndexKey As Variant) As Boolean
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Item = mCol.Item(vntIndexKey)
End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub AddRecurring(Key As Long, value As Single, Description As String)
    mColRecurring.Add account, value, Description
End Sub

Public Property Get CountRecurring() As Long
    CountRecurring = mColRecurring.Count
End Property

Public Property Get ItemRecurring(vntIndexKey As Variant) As clsRecurring
    Set ItemRecurring = mColRecurring.Item(vntIndexKey)
End Property

Public Sub RemoveRecurring(vntIndexKey As Variant)
    mColRecurring.Remove vntIndexKey
End Sub

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Property Get account() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Account
    account = mvarAccount
End Property

Public Property Let account(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Account = 5
    mvarAccount = vData
End Property

Public Property Get UsageCharge() As Single
    UsageCharge = mUsageCharge
End Property

Public Property Let UsageCharge(ByVal vData As Single)
    mUsageCharge = vData
End Property

Public Property Get PropertyUse() As String
    PropertyUse = mvarPropertyUse
End Property

Public Property Let PropertyUse(ByVal vData As String)
    mvarPropertyUse = vData
End Property

Public Property Get Service() As String
    Service = mvarService
End Property

Public Property Let Service(ByVal vData As String)
    mvarService = vData
End Property

Public Property Get ServiceDescription() As String
    ServiceDescription = mServiceDescription
End Property

Public Property Let ServiceDescription(ByVal vData As String)
    mServiceDescription = vData
End Property

Public Property Get Flag() As Boolean
    Flag = mFlag
End Property

Public Property Let Flag(ByVal vData As Boolean)
    mFlag = vData
End Property

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    Set mColRecurring = New clsRecurring
    Flag = False
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub


