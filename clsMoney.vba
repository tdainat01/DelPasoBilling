Option Compare Database
Option Explicit

Private mMonth As String
Private mDay As String
Private mYear As String
Private mAccountNumber As Long
Private mAmount As Currency
Private mTransaction As String
Private mCode As String
Private mPosted As String
Private mBehindMe As Long
Private mTransDate As Date


Public Property Get Month() As String
    Month = mMonth
End Property

Public Property Let Month(ByVal vData As String)
    mMonth = vData
End Property
Public Property Get Day() As String
    Day = mDay
End Property

Public Property Let Day(ByVal vData As String)
    mDay = vData
End Property
Public Property Get Year() As String
    Year = mYear
End Property

Public Property Let Year(ByVal vData As String)
    mYear = vData
End Property
Public Property Get AccountNumber() As Long
    AccountNumber = mAccountNumber
End Property

Public Property Let AccountNumber(ByVal vData As Long)
    mAccountNumber = vData
End Property
Public Property Get Amount() As Currency
    Amount = mAmount
End Property

Public Property Let Amount(ByVal vData As Currency)
    mAmount = vData
End Property
Public Property Get transaction() As String
    transaction = mTransaction
End Property

Public Property Let transaction(ByVal vData As String)
    mTransaction = vData
End Property
Public Property Get Code() As String
    Code = mCode
End Property

Public Property Let Code(ByVal vData As String)
    mCode = vData
End Property
Public Property Get posted() As String
    posted = mPosted
End Property

Public Property Let posted(ByVal vData As String)
    mPosted = vData
End Property
Public Property Get BehindMe() As Long
    BehindMe = mBehindMe
End Property

Public Property Let BehindMe(ByVal vData As Long)
    mBehindMe = vData
End Property
Public Property Get TransDate() As Date
    TransDate = mTransDate
End Property

Public Property Let TransDate(ByVal vData As Date)
    mTransDate = vData
End Property


