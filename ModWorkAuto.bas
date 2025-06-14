Attribute VB_Name = "ModWorkAuto"
Public DbDataDB As New ADODB.Connection
Public RsUser As New ADODB.Recordset
Public RsComp As New ADODB.Recordset
Public RsGroup As New ADODB.Recordset
Public RsMsr As New ADODB.Recordset
Public RsItem As New ADODB.Recordset
Public hPath As String
Public sPath As String
Public mUser As String
Public mPassword As String
Public mCmpName As String
Public mUName As String
Public mUType As String
Public mBranch As Double
Public mBYear As Double
Public MSCONNECT As String
Public mDsnName As String
'Public MyDsnConnect As New DsnConn
Public mYear As String
Public mTYear As String
Public mFinYear As String
Public mSPlace As String
Public mLVer As String
Public mSVer As String
Public Function OpenRecordSet()
Set RsUser = Nothing
RsUser.Open "Select * From [UsrMst]", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
'Set RsGroup = Nothing
'RsGroup.Open "Select GName,GType,GCode From GrpMst Order By GName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
End Function
'Public Function OpDtlBf()
'Dim DBLocalDB As New ADODB.Connection
'Dim DbDataDBL As New ADODB.Connection
'Dim RSSysUser As New ADODB.Recordset
'Dim mQ As String
'Dim mAcCode As Double
'Dim mDr As Double
'Dim mCr As Double
'DBLocalDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source='" & App.Path + "\LocalDB.Mdb" & "'"
'DBLocalDB.Open
'RSSysUser.Open "Select * From TranBr Where F_Code=" & mBYear - 1, DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
'If RSSysUser.EOF = False Then
'    Set DbDataDBL = Nothing
'    DbDataDBL.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & hPath + "1\" & RSSysUser.Fields("F_Year") & "\DataDB.Mdb"
'    DbDataDBL.Open
'    mQ = "Select * From QLedger Where AcCode>0 Order By AcCode"
'    Set RSSysUser = Nothing
'    RSSysUser.Open mQ, DbDataDBL, adOpenDynamic, adLockReadOnly, adCmdText
'    Do While RSSysUser.EOF = False
'        mAcCode = RSSysUser.Fields("AcCode")
'        mDr = 0
'        mCr = 0
'        Do While mAcCode = RSSysUser.Fields("AcCode")
'            mDr = mDr + RSSysUser.Fields("Dr")
'            mCr = mCr + RSSysUser.Fields("Cr")
'            RSSysUser.MoveNext
'            If RSSysUser.EOF = True Then Exit Do
'        Loop
'        mCr = Round(mCr, 2)
'        mDr = Round(mDr, 2)
'        If mDr - mCr > 0 Then
'            DbDataDB.Execute "Update AcMst Set OpBal=" & Round(mDr - mCr, 2) & ",AcSide='Dr' Where AcCode=" & mAcCode
'        ElseIf mDr - mCr < 0 Then
'            DbDataDB.Execute "Update AcMst Set OpBal=" & Round(mCr - mDr, 2) & ",AcSide='Cr' Where AcCode=" & mAcCode
'        Else
'            DbDataDB.Execute "Update AcMst Set OpBal=0 Where AcCode=" & mAcCode
'        End If
'    Loop
'    mQ = "Select * From QVatLedger Where AcCode=0 Order By AcCode"
'    Set RSSysUser = Nothing
'    RSSysUser.Open mQ, DbDataDBL, adOpenDynamic, adLockReadOnly, adCmdText
'    DbDataDB.BeginTrans
'    If RSSysUser.EOF = False Then
'        mAcCode = RSSysUser.Fields("AcCode")
'        mDr = 0
'        mCr = 0
'        Do While RSSysUser.EOF = False
'            mDr = mDr + RSSysUser.Fields("Dr")
'            mCr = mCr + RSSysUser.Fields("Cr")
'            RSSysUser.MoveNext
'        Loop
'        mCr = Round(mCr, 2)
'        mDr = Round(mDr, 2)
'        If mDr - mCr > 0 Then
'            DbDataDB.Execute "Update AcMst Set OpBal=" & Round(mDr - mCr, 2) & ",AcSide='Dr' Where AcCode=" & mAcCode
'        ElseIf mDr - mCr < 0 Then
'            DbDataDB.Execute "Update AcMst Set OpBal=" & Round(mCr - mDr, 2) & ",AcSide='Cr' Where AcCode=" & mAcCode
'        Else
'            DbDataDB.Execute "Update AcMst Set OpBal=0 Where AcCode=" & mAcCode
'        End If
'    End If
'    DbDataDB.CommitTrans
'    MsgBox "Successfully A/c. Balance Transfered.", vbInformation, "Alert"
'End If
'End Function

Public Function SetProfit(ByVal mAcCode As Double) As Double
Dim RsQ As New ADODB.Recordset
Dim mAmount As Double
Set RsQ = Nothing
mAmount = 0
RsQ.Open "Select IIF(IsNull(Sum(CBal-DBal))=False,Sum(CBal-DBal),0) As RTotal From QTrialBal Where AcCode=" & mAcCode & " And HType=0", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsQ.EOF = False Then mAmount = RsQ.Fields("RTotal")
SetProfit = mAmount
End Function

Public Function SetProfitAll(ByVal mAcList As String) As Double
Dim RsQ As New ADODB.Recordset
Dim mAmount As Double
mAmount = 0
RsQ.Open "Select IIF(IsNull(Sum(CBal-DBal))=False,Sum(CBal-DBal),0) As RTotal From QTrialBal Where AcCode In (" & mAcList & ") And HType=0 And LCode Not In (107,108)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsQ.EOF = False Then mAmount = RsQ.Fields("RTotal")
SetProfitAll = Round(mAmount, 2)
End Function

