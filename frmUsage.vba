Option Compare Database

Private Sub Form_Load()
Dim vOpenArgs As Variant
Dim rst As New ADODB.Recordset
Dim query As String
Dim iFirstRead As Long
Dim iSecondRead As Long
Dim iUsage As Long
Dim iCtr As Long

   On Error GoTo Form_Load_Error

    If IsNull(Me.OpenArgs) Or IsEmpty(Me.OpenArgs) Then Exit Sub
    
    vOpenArgs = Me.OpenArgs
    
    query = " SELECT top 2 MeterReads.account, MeterReads.batch_date, customer.meter_number, MeterReads.normal_read, MeterReads.low_read" & _
            " FROM customer INNER JOIN MeterReads ON customer.account = MeterReads.account" & _
            " where customer.account = " & vOpenArgs & " and posted = 'Y'" & _
            " order by batch_date DESC"

    rst.Open query, CurrentProject.Connection
    iCtr = 1
    If rst.BOF And rst.EOF Then
        'nothing returned
        Me.txtAccount = vOpenArgs
        Me.txtReadDate = Now
        Me.txtMeterNum = 0
        Me.txtCurRead = 0
        Me.txtPrevRead = 0
        Me.txtUsage = 0
    Else
        Do While Not rst.EOF
        If iCtr < 2 Then
            Me.txtAccount = vOpenArgs
            Me.txtReadDate = IIf(IsNull(rst.Fields(1).value), Now, rst.Fields(1).value)
            Me.txtMeterNum = IIf(IsNull(rst.Fields(2).value), 0, rst.Fields(2).value)
            iSecondRead = IIf(IsNull(rst.Fields(3).value), 0, rst.Fields(3).value)
            Me.txtCurRead = iSecondRead
        Else
            iFirstRead = IIf(IsNull(rst.Fields(3).value), 0, rst.Fields(3).value)
            Me.txtPrevRead = iFirstRead
        End If
            iCtr = iCtr + 1
            rst.MoveNext
        Loop
    End If

    iUsage = Abs(iSecondRead - iFirstRead)
    Me.txtUsage = iUsage
    
    rst.Close
    
   On Error GoTo 0
   Exit Sub

Form_Load_Error:
Dim errNum As Long
Dim errMsg As String
Dim errSource As String
errNum = Err.Number
errSource = Err.source
errMsg = Err.Description

    Call LogError(errNum, errSource, errMsg & " in procedure Form_Load of VBA Document Form_frmUsage")
    MsgBox "Error " & errNum & " (" & errMsg & ") in procedure Form_Load of VBA Document Form_frmUsage"
End Sub
