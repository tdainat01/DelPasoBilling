Option Compare Database
Option Explicit

Private mAccount As Long
Private mGroup As String
Private mMastpar As String
Private mCycle As Long
Private mMfgCode As String
Private mStartDate As Date
Private mStatus As String
Private mMeterNumber As String
Private mTermDate As Date
Private mOutTown As Boolean
Private mMeterSize As Double
Private mPropertyUse As String
Private mBackFlow As Boolean
Private mServiceDiscon As Boolean
Private mFireSize As Single
Private mUnitMeasure As String
Private mCurrentRead As Double
Private mCurrentDate As Date
Private mRateCode As String
Private mPreviousRead As Double
Private mPreviousDate As Date
Private mGalcubUsed As Double
Private mMeterSite As String
Private mDeposit As Currency
Private mUseCharge As Currency
Private mPastDue As Currency
Private mPrevBalance As Currency
Private mCurrentDue As Currency
Private mSpecialCredit As Currency
Private mTotalDue As Currency
Private mSpecialCharge As Currency
Private mSpecialDescription As String
Private mPhyAddress As String
Private mLien As String
Private mName As String
Private mBillName As String
Private mAddr1 As String
Private mCareOf As String
Private mCity As String
Private mState As String
Private mZip As String
Private mComment As String
Private mTransloc As Long

Public Property Get account() As Long
    account = mAccount
End Property

Public Property Let account(ByVal vData As Long)
    mAccount = vData
End Property
Public Property Get group() As String
    group = mGroup
End Property

Public Property Let group(ByVal vData As String)
    mGroup = vData
End Property
Public Property Get mastpar() As String
    mastpar = mMastpar
End Property

Public Property Let mastpar(ByVal vData As String)
    mMastpar = vData
End Property
Public Property Get Cycle() As Long
    Cycle = mCycle
End Property

Public Property Let Cycle(ByVal vData As Long)
    mCycle = vData
End Property
Public Property Get MfgCode() As String
    MfgCode = mMfgCode
End Property

Public Property Let MfgCode(ByVal vData As String)
    mMfgCode = vData
End Property
Public Property Get StartDate() As Date
    StartDate = mStartDate
End Property

Public Property Let StartDate(ByVal vData As Date)
    mStartDate = vData
End Property
Public Property Get Status() As String
    Status = mStatus
End Property

Public Property Let Status(ByVal vData As String)
    mStatus = vData
End Property
Public Property Get MeterNumber() As String
    MeterNumber = mMeterNumber
End Property

Public Property Let MeterNumber(ByVal vData As String)
    mMeterNumber = vData
End Property

Public Property Get TermDate() As Date
    TermDate = mTermDate
End Property

Public Property Let TermDate(ByVal vData As Date)
    mTermDate = vData
End Property
Public Property Get OutTown() As Boolean
    OutTown = mOutTown
End Property

Public Property Let OutTown(ByVal vData As Boolean)
    mOutTown = vData
End Property

Public Property Get MeterSize() As Double
    MeterSize = mMeterSize
End Property

Public Property Let MeterSize(ByVal vData As Double)
    mMeterSize = vData
End Property
Public Property Get PropertyUse() As String
    PropertyUse = mPropertyUse
End Property

Public Property Let PropertyUse(ByVal vData As String)
    mPropertyUse = vData
End Property
Public Property Get Backflow() As Boolean
    Backflow = mBackFlow
End Property

Public Property Let Backflow(ByVal vData As Boolean)
    mBackFlow = vData
End Property
Public Property Get ServiceDiscon() As Boolean
    ServiceDiscon = mServiceDiscon
End Property

Public Property Let ServiceDiscon(ByVal vData As Boolean)
    mServiceDiscon = vData
End Property
Public Property Get FireSize() As Single
    FireSize = mFireSize
End Property

Public Property Let FireSize(ByVal vData As Single)
    mFireSize = vData
End Property
Public Property Get UnitMeasure() As String
    UnitMeasure = mUnitMeasure
End Property

Public Property Let UnitMeasure(ByVal vData As String)
    mUnitMeasure = vData
End Property
Public Property Get CurrentRead() As Double
    CurrentRead = mCurrentRead
End Property

Public Property Let CurrentRead(ByVal vData As Double)
    mCurrentRead = vData
End Property
Public Property Get CurrentDate() As Date
    CurrentDate = mCurrentDate
End Property

Public Property Let CurrentDate(ByVal vData As Date)
    mCurrentDate = vData
End Property
Public Property Get RateCode() As String
    RateCode = mRateCode
End Property

Public Property Let RateCode(ByVal vData As String)
    mRateCode = vData
End Property
Public Property Get PreviousRead() As Double
    PreviousRead = mPreviousRead
End Property

Public Property Let PreviousRead(ByVal vData As Double)
    mPreviousRead = vData
End Property
Public Property Get PreviousDate() As Date
    PreviousDate = mPreviousDate
End Property

Public Property Let PreviousDate(ByVal vData As Date)
    mPreviousDate = vData
End Property
Public Property Get GalCubUsed() As Double
    GalCubUsed = mGalcubUsed
End Property

Public Property Let GalCubUsed(ByVal vData As Double)
    mGalcubUsed = vData
End Property
Public Property Get MeterSite() As String
    MeterSite = mMeterSite
End Property

Public Property Let MeterSite(ByVal vData As String)
    mMeterSite = vData
End Property
Public Property Get Deposit() As Double
    Deposit = mDeposit
End Property

Public Property Let Deposit(ByVal vData As Double)
    mDeposit = vData
End Property
Public Property Get UseCharge() As Double
    UseCharge = mUseCharge
End Property

Public Property Let UseCharge(ByVal vData As Double)
    mUseCharge = vData
End Property
Public Property Get PastDue() As Double
    PastDue = mPastDue
End Property

Public Property Let PastDue(ByVal vData As Double)
    mPastDue = vData
End Property
Public Property Get PrevBalance() As Double
    PrevBalance = mPrevBalance
End Property

Public Property Let PrevBalance(ByVal vData As Double)
    mPrevBalance = vData
End Property
Public Property Get CurrentDue() As Double
    CurrentDue = mCurrentDue
End Property

Public Property Let CurrentDue(ByVal vData As Double)
    mCurrentDue = vData
End Property
Public Property Get SpecialCredit() As Double
    SpecialCredit = mSpecialCredit
End Property

Public Property Let SpecialCredit(ByVal vData As Double)
    mSpecialCredit = vData
End Property
Public Property Get TotalDue() As Double
    TotalDue = mTotalDue
End Property

Public Property Let TotalDue(ByVal vData As Double)
    mTotalDue = vData
End Property
Public Property Get SpecialCharge() As Double
    SpecialCharge = mSpecialCharge
End Property

Public Property Let SpecialCharge(ByVal vData As Double)
    mSpecialCharge = vData
End Property
Public Property Get SpecialDescription() As String
    SpecialDescription = mSpecialDescription
End Property

Public Property Let SpecialDescription(ByVal vData As String)
    mSpecialDescription = vData
End Property
Public Property Get PhyAddress() As String
    PhyAddress = mPhyAddress
End Property

Public Property Let PhyAddress(ByVal vData As String)
    mPhyAddress = vData
End Property
Public Property Get lien() As String
    lien = mLien
End Property

Public Property Let lien(ByVal vData As String)
    mLien = vData
End Property
Public Property Get Name() As String
    Name = mName
End Property

Public Property Let Name(ByVal vData As String)
    mName = vData
End Property
Public Property Get BillName() As String
    BillName = mBillName
End Property

Public Property Let BillName(ByVal vData As String)
    mBillName = vData
End Property
Public Property Get addr1() As String
    addr1 = mAddr1
End Property

Public Property Let addr1(ByVal vData As String)
    mAddr1 = vData
End Property
Public Property Get CareOf() As String
    CareOf = mCareOf
End Property

Public Property Let CareOf(ByVal vData As String)
    mCareOf = vData
End Property
Public Property Get city() As String
    city = mCity
End Property

Public Property Let city(ByVal vData As String)
    mCity = vData
End Property
Public Property Get state() As String
    state = mState
End Property

Public Property Let state(ByVal vData As String)
    mState = vData
End Property
Public Property Get Zip() As String
    Zip = mZip
End Property

Public Property Let Zip(ByVal vData As String)
    mZip = vData
End Property
Public Property Get comment() As String
    comment = mComment
End Property

Public Property Let comment(ByVal vData As String)
    mComment = vData
End Property
Public Property Get Transloc() As Long
    Transloc = mTransloc
End Property

Public Property Let Transloc(ByVal vData As Long)
    mTransloc = vData
End Property
