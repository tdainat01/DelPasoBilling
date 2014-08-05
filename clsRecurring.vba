Option Compare Database
Option Explicit

Private mCol As Collection
Private mAccount As Long
Private mCharge As Single
Private mTransaction As String

Public Property Get account() As Long
    account = mAccount
End Property

Public Property Let account(ByVal vData As Long)
    mAccount = vData
End Property

Public Property Get Charge() As Single
    Charge = mCharge
End Property

Public Property Let Charge(ByVal vData As Single)
    mCharge = vData
End Property

Public Property Get Description() As String
    Description = mTransaction
End Property

Public Property Let Description(ByVal vData As String)
    mTransaction = vData
End Property

Public Function Add(account As Long, Charge As Single, Description As String) As clsRecurring
Dim objRecurring As New clsRecurring
Dim Key As Variant

    With objRecurring
        .account = account
        .Charge = Charge
        .Description = Description
        Key = GetGUID
    End With

    mCol.Add objRecurring, Key
    Set Add = objRecurring
Exit_Function:
    Set objRecurring = Nothing
    Exit Function
End Function

Public Sub Remove(Index As Variant)
    mCol.Remove Index
End Sub
 
Function Item(Index As Variant) As clsRecurring
    Set Item = mCol.Item(Index)
End Function
 
Property Get Count() As Long
    Count = mCol.Count
End Property
 
Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub
 
Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

