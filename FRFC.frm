VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Begin VB.Form FrmFCReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Foreign Contribution Report"
   ClientHeight    =   6585
   ClientLeft      =   540
   ClientTop       =   435
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRFC.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   9855
   Begin VB.Frame FraMst 
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   15732
      Begin VB.TextBox TxtDTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   16560
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   0
         Width           =   1776
      End
      Begin VB.TextBox TxtCTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   18720
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   0
         Width           =   1776
      End
      Begin VB.CommandButton TlbSav 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   3
         Left            =   6996
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Print"
         Top             =   6000
         Width           =   732
      End
      Begin VB.CommandButton TlbSav 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   1
         Left            =   7764
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancel"
         Top             =   6000
         Width           =   852
      End
      Begin VB.CommandButton TlbSav 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   2
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Exit"
         Top             =   6000
         Width           =   732
      End
      Begin VSFlex7Ctl.VSFlexGrid LsvClient 
         Height          =   5652
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   9336
         _cx             =   16468
         _cy             =   9970
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   15335136
         ForeColor       =   -2147483640
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   128
         BackColorBkg    =   15335136
         BackColorAlternate=   15335136
         GridColor       =   192
         GridColorFixed  =   128
         TreeColor       =   -2147483632
         FloodColor      =   255
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FRFC.frx":0442
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   2
         AutoSearchDelay =   20
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   128
         ForeColorFrozen =   128
         WallPaperAlignment=   9
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfHelp 
         Height          =   3372
         Left            =   9720
         TabIndex        =   5
         Top             =   3000
         Width           =   5892
         _cx             =   10393
         _cy             =   5948
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   8438015
         ForeColor       =   -2147483640
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   128
         BackColorBkg    =   12648447
         BackColorAlternate=   16777215
         GridColor       =   192
         GridColorFixed  =   128
         TreeColor       =   -2147483632
         FloodColor      =   255
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FRFC.frx":048B
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   2
         AutoSearchDelay =   20
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   128
         ForeColorFrozen =   128
         WallPaperAlignment=   9
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfHelp1 
         Height          =   3372
         Left            =   9720
         TabIndex        =   8
         Top             =   240
         Width           =   5892
         _cx             =   10393
         _cy             =   5948
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   8438015
         ForeColor       =   -2147483640
         BackColorFixed  =   12632256
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16777215
         ForeColorSel    =   128
         BackColorBkg    =   12648447
         BackColorAlternate=   16777215
         GridColor       =   192
         GridColorFixed  =   128
         TreeColor       =   -2147483632
         FloodColor      =   255
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FRFC.frx":04D4
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   2
         AutoSearchDelay =   20
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   7
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   128
         ForeColorFrozen =   128
         WallPaperAlignment=   9
      End
      Begin VB.Shape ShpMst 
         BorderWidth     =   2
         Height          =   6372
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   9612
      End
   End
   Begin Crystal.CrystalReport RepPrint 
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
End
Attribute VB_Name = "FrmFCReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DBWorkTmp As New ADODB.Connection
Dim RsRep As New ADODB.Recordset
Dim RsClient As New ADODB.Recordset
Dim mAcCode As Double
Dim mAcList As String
Dim mProfit As Double
Private Sub Form_Load()
    Me.Left = 50
    Me.Top = 50
    DBWorkTmp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source='" & App.Path + "\LocalDB.Mdb'"
    DBWorkTmp.Open
    SetCombo
    SetTool (True)
End Sub
Private Function SetCombo()
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select '' As ParentNm,AcMst.FileNo,AcMst.City,AcMst.AcName,GrpMst.GName,AcMst.AcCode,'' As RName,0 As MPaCode From AcMst,GrpMst Where IsNull(AcMst.PaCode)=True " & _
"And AcMst.AcType=GrpMst.GCode And GrpMst.GCode=2 Union All Select AcMst1.AcName+IIF(Len(AcMst1.City)>0,', '+AcMst1.City,''),AcMst.FileNo,AcMst.City,AcMst.AcName," & _
"GrpMst.GName,AcMst.AcCode,'',AcMst1.AcCode From AcMst,GrpMst,AcMst As AcMst1 Where AcMst.PaCode=AcMst1.AcCode And AcMst.AcType=GrpMst.GCode And GrpMst.GCode=2 Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Set LsvClient.DataSource = RsQry
With LsvClient
    .TextMatrix(0, 0) = "NAME"
    .ColWidth(0) = 6000
    .TextMatrix(0, 1) = "FILE NO."
    .ColWidth(1) = 1500
    .TextMatrix(0, 2) = "" 'City
    .ColWidth(2) = 0
    .TextMatrix(0, 3) = "" 'AcName
    .ColWidth(3) = 0
    .TextMatrix(0, 4) = "TYPE"
    .ColWidth(4) = 500
    .TextMatrix(0, 5) = "ACCODE"
    .ColWidth(5) = 0
    .TextMatrix(0, 6) = ""
    .ColWidth(6) = 0
    .ColWidth(7) = 0    'Main PaCode
    .Col = 1
    .ExtendLastCol = True
    If .Rows > 1 Then .Row = 1
    .FontBold = False
    .Refresh
    Set RsQry = Nothing
    RsQry.Open "Select AcMst.AcCode,AcMst1.AcName+', '+AcMst1.City As RName,AcMst1.AcCode As PaCode From AcMst,AcMst" & _
    " As AcMst1 Where AcMst.PaCode=AcMst1.AcCode Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        .Row = 1
        .Row = .FindRow(RsQry.Fields("AcCode"), , 5)
        If .Row > 1 Then
            .TextMatrix(.Row, 6) = RsQry.Fields("RName")
            .TextMatrix(.Row, 7) = RsQry.Fields("PaCode")
        End If
        RsQry.MoveNext
    Loop
End With
Set RsClient = Nothing
RsClient.Open "Select * From AcMst Order By AcCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MDIWork.PctMdi.Visible = True
    Unload Me
End Sub

Private Sub Tlbsav_Click(Index As Integer)
Select Case TlbSav(Index).ToolTipText
    Case "Cancel"
        SetCombo
        mAcCode = 0
        LsvClient.SetFocus
    Case "Print"
        SetParent
        PrintRec
        DbDataDB.BeginTrans
'            DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'FC_REP','PRINT','" & Date & "','" & Time & "')"
        DbDataDB.CommitTrans
        mAcCode = 0
    Case "Exit"
        If MsgBox("Close All Reports? ", vbInformation + vbYesNo, "Confirmation") = vbYes Then
        DBWorkTmp.Close
        Unload Me
        End If
End Select
End Sub
Private Function SetTool(ByVal mVal As Boolean)
    TlbSav(1).Enabled = True
    TlbSav(2).Enabled = True
End Function
Private Sub PrintRec()
Dim i As Integer
Dim mTotal As Double
Dim RsQry As New ADODB.Recordset
mAcCode = Val(LsvClient.TextMatrix(LsvClient.Row, 5))
Set RsRep = Nothing
RsRep.Open "Select RepDtl.*,AdtMst.AdtName From RepDtl,AdtMst Where AcCode=" & mAcCode & " And RepDtl.AdtNo=AdtMst.AdtNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsRep.EOF = False Then
    If Len(RsRep.Fields("RUDIN")) = 0 Then
        MsgBox "UDIN not issued. Non-UDIN Report will be printed.", vbInformation, "Information"
    End If
Else
    MsgBox "Signing Information not available. Report cannot be printed.", vbCritical, "Alert"
    Exit Sub
End If
Set RsClient = Nothing
RsClient.Open "Select * From AcMst Where AcCode=" & Val(LsvClient.TextMatrix(LsvClient.Row, 7)), DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsClient.EOF = False Then
    If RsRep.EOF = False Then
        If Len(RsRep.Fields("RUDIN")) = 0 Then
            MsgBox "UDIN not issued. Report cannot be printed.", vbInformation, "Information"
            Exit Sub
        End If
        Dim RsQ As New ADODB.Recordset
        DBWorkTmp.BeginTrans
        DBWorkTmp.Execute "Delete From TmpTrialBal"
        DBWorkTmp.Execute "Delete From TmpCtDtl"
        DBWorkTmp.Execute "Delete From TmpRpDtl"
        DBWorkTmp.CommitTrans
        DBWorkTmp.BeginTrans
        RsQ.Open "Select * From QTrialBal Where AcCode In (" & mAcList & ")", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
        Do While RsQ.EOF = False
            DBWorkTmp.Execute "Insert InTo TmpTrialBal (HType,HSide,AcCode,HCode,LCode,OpDr,OpCr,ADr,ACr,DBal,CBal) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HSide") & _
            "'," & RsQ.Fields("AcCode") & "," & RsQ.Fields("HCode") & "," & RsQ.Fields("LCode") & "," & RsQ.Fields("OpDr") & "," & RsQ.Fields("OpCr") & "," & RsQ.Fields("ADr") & _
            "," & RsQ.Fields("ACr") & "," & RsQ.Fields("DBal") & "," & RsQ.Fields("CBal") & ")"
            RsQ.MoveNext
        Loop
        Set RsQ = Nothing
        RsQ.Open "Select * From QCtDtl Where AcCode In (" & mAcList & ")", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
        Do While RsQ.EOF = False
            DBWorkTmp.Execute "Insert InTo TmpCtDtl (HType,HSide,AcCode,HCode,LCode,ECode,EName,Side,Amt) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HSide") & "'," & _
            RsQ.Fields("AcCode") & "," & RsQ.Fields("HCode") & "," & RsQ.Fields("LCode") & "," & RsQ.Fields("ECode") & ",'" & RsQ.Fields("EName") & "','" & RsQ.Fields("Side") & _
            "'," & RsQ.Fields("Amt") & ")"
            RsQ.MoveNext
        Loop
        Set RsQ = Nothing
        RsQ.Open "Select HType,HSide,AcCode,HCode,LCode,ECode,Side,Sum(Amt) As Amt From QRpDtl Where AcCode In (" & mAcList & ") Group By AcCode,HType,HSide,Side,HCode,LCode,ECode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
        Do While RsQ.EOF = False
            DBWorkTmp.Execute "Insert InTo TmpRpDtl (HType,HSide,AcCode,HCode,LCode,ECode,Side,Amt) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HSide") & "'," & _
            RsQ.Fields("AcCode") & "," & RsQ.Fields("HCode") & "," & RsQ.Fields("LCode") & "," & RsQ.Fields("ECode") & ",'" & RsQ.Fields("Side") & _
            "'," & RsQ.Fields("Amt") & ")"
            RsQ.MoveNext
        Loop
        DBWorkTmp.CommitTrans
        With RepPrint
            .Connect = MSCONNECT
            .ReportFileName = App.Path + "\Report\AdtRpt.Rpt"
            .SelectionFormula = "{TranBr.RepName}='VPDUW\'"
            .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
            .Formulas(1) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
            .Formulas(2) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
            .Formulas(3) = "mTitle1='Certificate to be given by a Chartered Accountant under Rule 17(5) of the Foreign Contribution (Regulation) Rules, 2011'"
            .Formulas(4) = "mPlace='Place: " & RsRep.Fields("RPlace") & "'"
            .Formulas(5) = "mDate='Date: " & RsRep.Fields("RpDt") & "'"
            .Formulas(6) = "mAuditNm='" & RsRep.Fields("AdtName") & "'"
            .Formulas(7) = "mAuditNo='" & RsRep.Fields("AdtNo") & "'"
            .Formulas(8) = "mUDIN='" & RsRep.Fields("RUDIN") & "'"
            .Formulas(9) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
            .Formulas(10) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
            .Formulas(11) = "mTName='" & RsClient.Fields("AcName") & ", " & RsClient.Fields("City") & "'"
            .Formulas(12) = "mAdd1='" & RsComp.Fields("Add1") & "'"
            .Formulas(13) = "mAdd2='" & RsComp.Fields("Add2") & "'"
            .Formulas(14) = "mCity='" & RsComp.Fields("City") & ", " & RsComp.Fields("State") & ", " & RsComp.Fields("PinCode") & "'"
            .Formulas(15) = "mPhone='" & RsComp.Fields("Phone") & "'"
            .Formulas(16) = "mCYear='" & CStr(Year(RsRep.Fields("RtDt"))) & "'"
            .Formulas(17) = "mCType='F'"
            .Formulas(18) = "mFY='" & CStr(Year(RsRep.Fields("RfDt"))) & "-" & CStr(Year(RsRep.Fields("RtDt"))) & "'"
            .Formulas(19) = "mTFCNo='" & RsClient.Fields("FCRegNo") & "'"
            .Formulas(20) = "mTFCDate='" & RsClient.Fields("FCRegDt") & "'"
            .Formulas(21) = "mTAddress='" & RsClient.Fields("Address") & IIf(Len(RsClient.Fields("Address")) > 0, ", ", "") & RsClient.Fields("City") & IIf(RsClient.Fields("Taluka") <> RsClient.Fields("City"), ", " + RsClient.Fields("Taluka"), "") & IIf((RsClient.Fields("District") <> RsClient.Fields("Taluka") And RsClient.Fields("District") <> RsClient.Fields("City")), ", " + RsClient.Fields("District"), "") & ", " & RsClient.Fields("State") & ", " & RsClient.Fields("PinCode") & "'"
            Set RsQry = Nothing
            RsQry.Open "Select IIF(IsNull(Sum(DBal))=True,0,Sum(Dbal)) As BalRs From TmpTrialBal Where AcCode In (" & mAcList & ") And HCode=12", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            .Formulas(22) = "mCFFNh='Rs. " & Format(CStr(RsQry.Fields("BalRs")), "0.00") & "'"
            Set RsQry = Nothing
            RsQry.Open "Select IIF(IsNull(Sum(OpDr))=True,0,Sum(OpDr)) As BalRs From TmpTrialBal Where AcCode In (" & mAcList & ") And HCode=12", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            .Formulas(23) = "mOFFNh='Rs. " & Format(CStr(RsQry.Fields("BalRs")), "0.00") & "'"
            Set RsQry = Nothing
            RsQry.Open "Select IIF(IsNull(Sum(ACr))=True,0,Sum(ACr)) As BalRs From TmpTrialBal Where AcCode In (" & mAcList & ") And (HCode In (13,51) or LCode In (75,76))", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            .Formulas(24) = "mFDon='Rs. " & Format(CStr(RsQry.Fields("BalRs")), "0.00") & "'"
            Set RsQry = Nothing
            RsQry.Open "Select IIF(IsNull(Sum(ACr))=True,0,Sum(ACr)) As BalRs From TmpTrialBal Where AcCode In (" & mAcList & ") And HType=0 And HCode Not In (13,51,54) And LCode Not In (75,76,107,108)", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            .Formulas(25) = "mFInt='Rs. " & Format(CStr(RsQry.Fields("BalRs")), "0.00") & "'"
            .Action = 1
            For i = 0 To 25
                .Formulas(i) = ""
            Next
            .SelectionFormula = ""
        End With
    Else
        MsgBox "Signing Information not available. Report cannot be printed.", vbCritical, "Alert"
        Exit Sub
    End If
End If
SetParent
PrintBS
SetParent
PrintIE
SetParent
PrintRecPay
End Sub
Private Sub SetParent()
Dim RsQ As New ADODB.Recordset
mAcCode = Val(LsvClient.TextMatrix(LsvClient.Row, 5))
RsQ.Open "Select * From QGroup Where PaCode=" & mAcCode & " Order By AcCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
mAcList = CStr(mAcCode)
Do While RsQ.EOF = False
    mAcList = mAcList + "," + CStr(RsQ.Fields("AcCode"))
    RsQ.MoveNext
Loop
End Sub
Private Sub PrintBS()
    mProfit = 0
    SetData
    BSheetPrint
    NotePrint
End Sub
Private Sub PrintIE()
    SetData1
    PlAcPrint
End Sub
Private Sub SetData()
Dim RsQry As New ADODB.Recordset
If mFinYear = "19-20" Then
    RsQry.Open "Select HedMst.HName,SeqMst.HCode From SeqMst,HedMst Where SeqMst.GCode=6 And SeqMst.HCode=HedMst.HCode And HedMst.HType=1 And HedMst.HCode<>14 And HedMst.HSide='C' Order By SeqMst.GSrN", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Else
    RsQry.Open "Select HedMst.HName,SeqMst.HCode From SeqMst,HedMst Where SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & _
    mAcCode & ") And SeqMst.HCode=HedMst.HCode And HedMst.HType=1 And HedMst.HCode<>14 And HedMst.HSide='C' Order By SeqMst.GSrN", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
End If
With VsfHelp
    .Cols = 8
    .Rows = 1
    .TextMatrix(0, 0) = "FUNDS AND LIABILITIES"
    .ColWidth(0) = 5000
    .TextMatrix(0, 1) = "NOTE"
    .ColWidth(1) = 700
    .TextMatrix(0, 2) = "AMOUNT RS"
    .ColWidth(2) = 1700
    .ColAlignment(2) = flexAlignRightCenter
    .TextMatrix(0, 3) = "PROPERTY AND ASSETS"
    .ColWidth(3) = 5000
    .TextMatrix(0, 4) = "NOTE"
    .ColWidth(4) = 700
    .TextMatrix(0, 5) = "AMOUNT RS"
    .ColWidth(5) = 1700
    .ColAlignment(5) = flexAlignRightCenter
    .TextMatrix(0, 6) = "DSIDE"
    .ColWidth(6) = 0
    .TextMatrix(0, 7) = "CSIDE"
    .ColWidth(7) = 0
    .Refresh
    .Rows = 2
    .Row = 1
    Do While RsQry.EOF = False
        .TextMatrix(.Row, 0) = RsQry.Fields("HName")
        .TextMatrix(.Row, 6) = RsQry.Fields("HCode")
        RsQry.MoveNext
        .Rows = .Rows + 1
        .Row = .Rows - 1
    Loop
    .Rows = .Rows + 1
    .Row = 1
    Set RsQry = Nothing
    RsQry.Open "Select HedMst.HName,SeqMst.HCode From SeqMst,HedMst Where SeqMst.GCode In (Select AcType From AcMst Where AcCode=" & mAcCode & ") And SeqMst.HCode=" & _
    "HedMst.HCode And HedMst.HType=1 And HedMst.HSide='D' Order By SeqMst.GSrN", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        .TextMatrix(.Row, 3) = RsQry.Fields("HName")
        .TextMatrix(.Row, 7) = RsQry.Fields("HCode")
        RsQry.MoveNext
        If RsQry.EOF = False Then
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            Else
                .Rows = .Rows + 1
                .Row = .Rows - 1
            End If
        End If
    Loop
    Dim i As Double
    Set RsQry = Nothing
    RsQry.Open "Select HSide As AcSide,HCode As HeadCode,IIF(IsNull(Sum(CBal))=True,0,Sum(CBal)) As TotRs From TmpTrialBal Where AcCode In (" & mAcList & ") And HType=1 And HCode<>14 Group By HSide,HCode Union All " & _
    "Select HSide,HCode,IIF(IsNull(Sum(DBal))=True,0,Sum(DBal)) From TmpTrialBal Where AcCode In (" & mAcList & ") And HType=1 And HCode<>14 Group By HSide,HCode Order By HeadCode", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        If RsQry.Fields("AcSide") = "C" Then
            If mFinYear = "19-20" Then
                If RsQry.Fields("HeadCode") = 5 Or RsQry.Fields("HeadCode") = 13 Or RsQry.Fields("HeadCode") = 58 Then
                    i = .FindRow(58, 1, 6)
                Else
                    i = .FindRow(RsQry.Fields("HeadCode"), 1, 6)
                End If
            Else
                i = .FindRow(RsQry.Fields("HeadCode"), 1, 6)
            End If
0            If i > 0 Then
                If RsQry.Fields("HeadCode") = 58 Or RsQry.Fields("HeadCode") = 13 Or RsQry.Fields("HeadCode") = 5 Then    '   Income And Expenditure
                    If mProfit = 0 Then
                        mProfit = SetProfitAll(mAcList)
                        .TextMatrix(i, 2) = Format(CStr(Val(.TextMatrix(i, 2)) + RsQry.Fields("TotRs")) + mProfit, "0.00")
                    Else
                        .TextMatrix(i, 2) = Format(CStr(Val(.TextMatrix(i, 2)) + RsQry.Fields("TotRs")), "0.00")
                    End If
                Else
                    .TextMatrix(i, 2) = Format(CStr(Val(.TextMatrix(i, 2)) + RsQry.Fields("TotRs")), "0.00")
                End If
            End If
        Else
            i = .FindRow(RsQry.Fields("HeadCode"), 1, 7)
            If i > 0 Then
                .TextMatrix(i, 5) = Format(CStr(Val(.TextMatrix(i, 5)) + RsQry.Fields("TotRs")), "0.00")
            End If
        End If
        RsQry.MoveNext
    Loop
    i = .FindRow(9, 1, 6)
    If i > 0 Then
        If Val(.TextMatrix(i, 2)) = 0 Then
            mProfit = SetProfitAll(mAcList)
            .TextMatrix(i, 2) = mProfit
        End If
    End If
    Set RsQry = Nothing
    RsQry.Open "Select HCode As HeadCode,IIF(IsNull(Sum(CBal))=True,0,Sum(CBal)) As TotRs From TmpTrialBal Where AcCode In (" & mAcList & _
    ") And LCode In(87,88,86,1) Group By HCode Order By HCode", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        If RsQry.Fields("HeadCode") = 58 Or RsQry.Fields("HeadCode") = 57 Then
            i = .FindRow(5, 1, 6)
        Else
            i = .FindRow(RsQry.Fields("HeadCode"), 1, 6)
        End If
        If i > 0 Then
            If Val(.TextMatrix(i, 2)) = 0 Then
                If RsQry.Fields("HeadCode") = 58 Or RsQry.Fields("HeadCode") = 57 Or RsQry.Fields("HeadCode") = 5 Then     '   Income And Expenditure
                    mProfit = SetProfitAll(mAcList)
                    .TextMatrix(i, 2) = Format(CStr(Val(.TextMatrix(i, 2)) + RsQry.Fields("TotRs")) + mProfit, "0.00")
                End If
            End If
        End If
        RsQry.MoveNext
    Loop
    Set RsQry = Nothing
    RsQry.Open "Select * From NtDtl Where AcCode=" & mAcCode & " And RType=1 Order By HCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        If RsQry.Fields("RSide") = "C" Then
            i = .FindRow(RsQry.Fields("HCode"), 1, 6)
            If i > 0 Then .TextMatrix(i, 1) = RsQry.Fields("Note")
        Else
            i = .FindRow(RsQry.Fields("HCode"), 1, 7)
            If i > 0 Then .TextMatrix(i, 4) = RsQry.Fields("Note")
        End If
        RsQry.MoveNext
    Loop
End With
SetFinalTot
End Sub

Private Sub SetFinalTot()
Dim i As Double
TxtDTotal.Text = "0.00"
TxtCTotal.Text = "0.00"
With VsfHelp
    For i = 1 To .Rows - 1
        TxtDTotal.Text = Val(TxtDTotal.Text) + Val(.TextMatrix(i, 2))
        TxtCTotal.Text = Val(TxtCTotal.Text) + Val(.TextMatrix(i, 5))
    Next
End With
TxtDTotal.Text = Format(TxtDTotal.Text, "0.00")
TxtCTotal.Text = Format(TxtCTotal.Text, "0.00")
End Sub

Private Sub BSheetPrint()
Dim i As Double
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From TmpBSPrn"
With VsfHelp
    For i = 1 To .Rows - 1
        DbDataDB.Execute "Insert Into TmpBSPrn (LName,LAmt,LNote,RName,RAmt,RNote,SrN) Values ('" & .TextMatrix(i, 0) & "'," & Val(.TextMatrix(i, 2)) & "," & _
        Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 3) & "'," & Val(.TextMatrix(i, 5)) & "," & Val(.TextMatrix(i, 4)) & "," & i & ")"
    Next
End With
DbDataDB.CommitTrans
If Val(TxtCTotal.Text) <> Val(TxtDTotal.Text) Then
    MsgBox "Balance Sheet does not tally. Report cannot be printed.", vbCritical, "Alert"
    Exit Sub
End If
With RepPrint
    .Connect = MSCONNECT
    .ReportFileName = App.Path + "\Report\BSXRpt.Rpt"
    .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
    .Formulas(1) = "mAdd1='" & RsComp.Fields("Add1") & "'"
    .Formulas(2) = "mAdd2='" & RsComp.Fields("Add2") & "'"
    .Formulas(3) = "mCity='" & RsComp.Fields("City") & ", " & RsComp.Fields("State") & ", " & RsComp.Fields("PinCode") & "'"
    .Formulas(4) = "mPhone='" & RsComp.Fields("Phone") & "'"
    .Formulas(5) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
    .Formulas(6) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
    .Formulas(7) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 0) & "'"
    .Formulas(8) = "mSubHead=''"
    .Formulas(9) = "mTitle1='Balance Sheet as on " & RsRep.Fields("RtDt") & "'"
    .Formulas(10) = "mPlace='Place : " & RsRep.Fields("RPlace") & "'"
    .Formulas(11) = "mDate='Date : " & RsRep.Fields("RpDt") & "'"
    .Formulas(12) = "mAuditNm='" & RsRep.Fields("AdtName") & "'"
    .Formulas(13) = "mAuditNo='" & RsRep.Fields("AdtNo") & "'"
    .Formulas(14) = "mTSign='Chief Functionary/Trustee'"
    .Formulas(15) = "mUDIN='" & RsRep.Fields("RUDIN") & "'"
    .Formulas(16) = "mClient='" & LsvClient.TextMatrix(LsvClient.Row, 3) & "'"
    .Formulas(17) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
    .Formulas(18) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
    .Formulas(19) = "mLTitle='LIABILITIES'"
    .Formulas(20) = "mRTitle='ASSETS'"
    .Formulas(21) = "mClient1='" & LsvClient.TextMatrix(LsvClient.Row, 3) & "'"
    .SelectionFormula = ""
    .Action = 1
    For i = 0 To 21
        .Formulas(i) = ""
    Next
End With
End Sub
Private Sub NotePrint()
Dim i As Integer
Dim mAcCodeAll As String
Dim mAcName As String
Dim mGroup As Double
Dim mTotal As Double
Dim RsDataO As New ADODB.Recordset
Dim RsDataD As New ADODB.Recordset
Dim RsQData As New ADODB.Recordset
Dim RsQ As New ADODB.Recordset
Dim mRAmt As Double
Dim mSAmt As Double
Dim mCAmt As Double
Dim mDAmt As Double
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From TmpNotePrn"
Set RsQ = Nothing
RsQ.Open "Select * From AcMst Where AcCode In (" & mAcList & ") And AcName='Foreign Contribution Account'", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsQ.EOF = False Then mAcCodeAll = CStr(RsQ.Fields("AcCode"))
Set RsQ = Nothing
RsQ.Open "Select * From AcMst Where AcCode In (" & mAcList & ") And AcName Not In ('" & LsvClient.TextMatrix(LsvClient.Row, 0) & _
"','Foreign Contribution Account')", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Do While RsQ.EOF = False
    If Len(mAcCodeAll) > 1 Then mAcCodeAll = mAcCodeAll + "," + CStr(RsQ.Fields("AcCode")) Else mAcCodeAll = CStr(RsQ.Fields("AcCode"))
    RsQ.MoveNext
Loop
Dim RsAc As New ADODB.Recordset
RsAc.Open "Select * From AcMst Order By AcCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
With VsfHelp
    Do While Len(mAcList) > 0
        If InStr(1, mAcList, ",") - 1 > 0 Then
            mAcCode = Val(Mid(mAcList, 1, InStr(1, mAcList, ",") - 1))
            mAcList = Mid(mAcList, InStr(1, mAcList, ",") + 1)
        Else
            mAcCode = Val(mAcList)
            mAcList = ""
        End If
        If RsAc.BOF = False Then
            RsAc.MoveFirst
            RsAc.Find "AcCode=" & mAcCode
        End If
        If RsAc.EOF = False Then
            If RsAc.Fields("AcName") = LsvClient.TextMatrix(LsvClient.Row, 0) Then
                mAcName = Space(3) + LsvClient.TextMatrix(LsvClient.Row, 0)
            ElseIf RsAc.Fields("AcName") = "Foreign Contribution Account" Then
                mAcName = Space(1) + RsAc.Fields("AcName")
            Else
                mAcName = RsAc.Fields("AcName") + IIf(Len(RsAc.Fields("City")) <> 0, (", " + RsAc.Fields("City")), "")
            End If
        End If
        If Val(.TextMatrix(2, 1)) <> 0 Then
            i = .FindRow(2, 1, 6)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 6))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.CBal As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("RTotal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & _
                        RsDataO.Fields("LName") & "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
            End If
        End If
        If Val(.TextMatrix(3, 1)) <> 0 Then
            i = .FindRow(3, 1, 6)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 6))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.CBal As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("RTotal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & _
                        .TextMatrix(i, 0) & "','" & RsDataO.Fields("LName") & "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
            End If
        End If
        If Val(.TextMatrix(4, 1)) <> 0 Then
            i = .FindRow(4, 1, 6)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 6))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.CBal As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("RTotal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & _
                        RsDataO.Fields("LName") & "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
            End If
        End If
        If Val(.TextMatrix(1, 4)) <> 0 Then
            i = .FindRow(6, 1, 7)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 7))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.OpDr As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And LedMst.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    Set RsQData = Nothing
                    RsQData.Open "Select IIF(IsNull(Sum(ADr))=True,0,Sum(ADr)) AS RTotal From TmpTrialBal Where AcCode=" & mAcCode & " And LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    mRAmt = 0
                    mRAmt = RsQData.Fields("RTotal")
                    Set RsQData = Nothing
                    RsQData.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) AS RTotal From TmpCtDtl Where AcCode=" & mAcCode & _
                    " And LCode=" & RsDataO.Fields("LCode") & " And EName Like '*Sale*'", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    mSAmt = 0
                    mSAmt = RsQData.Fields("RTotal")
                    Set RsQData = Nothing
                    RsQData.Open "Select IIF(IsNull(Sum(ACr))=True,0,Sum(ACr)) AS RTotal From TmpTrialBal Where AcCode=" & mAcCode & " And LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    mDAmt = 0
                    mDAmt = RsQData.Fields("RTotal") - mSAmt
                    Set RsQData = Nothing
                    RsQData.Open "Select IIF(IsNull(Sum(DBal))=True,0,Sum(DBal)) AS RTotal From TmpTrialBal Where AcCode=" & mAcCode & " And LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    mCAmt = 0
                    mCAmt = RsQData.Fields("RTotal")
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,LAmt,RAmt,LName,SAmt,CAmt,DAmt,HCode,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & _
                    .TextMatrix(i, 3) & "'," & RsDataO.Fields("RTotal") & "," & mRAmt & ",'" & RsDataO.Fields("LName") & "'," & mSAmt & "," & mCAmt & "," & mDAmt & "," & mGroup & ",'" & mAcName & "')"
                    RsDataO.MoveNext
                Loop
            End If
        End If
        If Val(.TextMatrix(2, 4)) <> 0 Then
            i = .FindRow(7, 1, 7)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 7))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.DBal As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("RTotal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "','" & _
                        RsDataO.Fields("LName") & "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
            End If
        End If
        If Val(.TextMatrix(3, 4)) <> 0 Then
            i = .FindRow(8, 1, 7)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 7))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.OpDr As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And LedMst.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    Set RsQData = Nothing
                    RsQData.Open "Select IIF(IsNull(Sum(ADr))=True,0,Sum(ADr)) AS RTotal From TmpTrialBal Where AcCode=" & mAcCode & " And LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    mRAmt = 0
                    mRAmt = RsQData.Fields("RTotal")
                    Set RsQData = Nothing
                    RsQData.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) AS RTotal From TmpCtDtl Where AcCode=" & mAcCode & " And LCode=" & RsDataO.Fields("LCode") & " And EName Like '*Sale*'", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    mSAmt = 0
                    mSAmt = RsQData.Fields("RTotal")
                    Set RsQData = Nothing
                    RsQData.Open "Select IIF(IsNull(Sum(ACr))=True,0,Sum(ACr)) AS RTotal From TmpTrialBal Where AcCode=" & mAcCode & " And LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    mDAmt = 0
                    mDAmt = RsQData.Fields("RTotal") - mSAmt
                    Set RsQData = Nothing
                    RsQData.Open "Select IIF(IsNull(Sum(DBal))=True,0,Sum(DBal)) AS RTotal From TmpTrialBal Where AcCode=" & mAcCode & " And LCode=" & RsDataO.Fields("LCode"), DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
                    mCAmt = 0
                    mCAmt = RsQData.Fields("RTotal")
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,LAmt,RAmt,LName,SAmt,CAmt,DAmt,HCode,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & _
                    .TextMatrix(i, 3) & "'," & RsDataO.Fields("RTotal") & "," & mRAmt & ",'" & RsDataO.Fields("LName") & "'," & mSAmt & "," & mCAmt & "," & mDAmt & "," & mGroup & ",'" & mAcName & "')"
                    RsDataO.MoveNext
                Loop
            End If
        End If
        If Val(.TextMatrix(4, 4)) <> 0 Then
            i = .FindRow(9, 4, 7)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 7))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,LedMst.HCode,TmpTrialBal.DBal As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & _
                " And LedMst.LCode=TmpTrialBal.LCode And TmpTrialBal.HCode<>14 And TmpTrialBal.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("RTotal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "','" & _
                        RsDataO.Fields("LName") & "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
            End If
        End If
        If Val(.TextMatrix(5, 4)) <> 0 Then
            i = .FindRow(10, 4, 7)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 7))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.DBal As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("RTotal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "','" & _
                        RsDataO.Fields("LName") & "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
            End If
        End If
        If Val(.TextMatrix(6, 4)) <> 0 Then
            i = .FindRow(11, 4, 7)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 7))
                Set RsDataO = Nothing
                RsDataO.Open "Select LedMst.LCode,LedMst.LName,TmpTrialBal.DBal As RTotal From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And LedMst.LCode=" & _
                "TmpTrialBal.LCode And TmpTrialBal.HCode=" & mGroup & " Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("RTotal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & "','" & _
                        RsDataO.Fields("LName") & "'," & RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
            End If
        End If
        If Val(.TextMatrix(7, 4)) <> 0 Then
            i = .FindRow(12, 4, 7)
            If i > 0 Then
                mGroup = Val(.TextMatrix(i, 7))
                Set RsDataO = Nothing
                RsDataO.Open "Select TmpTrialBal.DBal,LedMst.LName From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And TmpTrialBal.LCode=LedMst.LCode And " & _
                "TmpTrialBal.LCode In (10000,10001,10315) Order By TmptrialBal.LCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("DBal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & _
                        "',' " & RsDataO.Fields("LName") & "'," & RsDataO.Fields("DBal") & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
                Set RsDataO = Nothing
                RsDataO.Open "Select TmpTrialBal.DBal,LedMst.LName From TmpTrialBal,LedMst Where TmpTrialBal.AcCode=" & mAcCode & " And TmpTrialBal.HCode=" & mGroup & _
                " And TmpTrialBal.LCode=LedMst.LCode And TmpTrialBal.LCode Not In (10000,10001,10315) Order By LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsDataO.EOF = False
                    If RsDataO.Fields("DBal") <> 0 Then
                        DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,CName) Values (" & Val(.TextMatrix(i, 4)) & ",'" & .TextMatrix(i, 3) & _
                        "','" & RsDataO.Fields("LName") & "'," & RsDataO.Fields("DBal") & ",'" & mAcName & "')"
                    End If
                    RsDataO.MoveNext
                Loop
            End If
        End If
    Loop
    If Val(.TextMatrix(1, 1)) <> 0 Then
        If mFinYear <> "19-20" Then
            i = .FindRow(1, 1, 6)
            If i = -1 Then i = .FindRow(13, 1, 6)
        Else
            i = .FindRow(57, 1, 6)
            If i = -1 Then i = .FindRow(1, 1, 6)
        End If
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 6))
            Set RsDataO = Nothing
            If Val(.TextMatrix(i, 6)) <> 1 Then
                RsDataO.Open "Select IIF(IsNull(Sum(OpCr))=True,0,Sum(OpCr)) As RTotal From TmpTrialBal Where HCode=" & mGroup, DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            Else
            End If
            If RsDataO.EOF = False Then
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','  Opening Balance'," & _
                    RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
                End If
            End If
            Set RsDataO = Nothing
            If Val(.TextMatrix(i, 6)) <> 1 Then
                RsDataO.Open "Select IIF(Side=HSide,'Add','Less') As TrnType,EName,Sum(Amt) As Amount From TmpCtDtl Where HCode=" & mGroup & " Group By IIF(Side=HSide,'Add','Less'),EName" & _
                " Order By IIF(Side=HSide,'Add','Less')", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            Else
                RsDataO.Open "Select IIF(Side=HSide,'Add','Less') As TrnType,EName,Sum(Amt) As Amount From TmpCtDtl Where HCode In (1,13) Group By IIF(Side=HSide,'Add','Less'),EName" & _
                " Order By IIF(Side=HSide,'Add','Less')", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            End If
            Do While RsDataO.EOF = False
                If RsDataO.Fields("Amount") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "',' " & _
                    RsDataO.Fields("TrnType") & ":" & Space(1) & RsDataO.Fields("EName") & "'," & IIf(RsDataO.Fields("TrnType") = "Add", RsDataO.Fields("Amount"), RsDataO.Fields("Amount") * -1) & _
                    "," & IIf(RsDataO.Fields("TrnType") = "Add", 0, 1) & ")"
                End If
                RsDataO.MoveNext
            Loop
            Set RsDataO = Nothing
            If Val(.TextMatrix(i, 6)) <> 1 Then
                RsDataO.Open "Select IIF(IsNull(Sum(CBal))=True,0,Sum(CBal)) As RTotal From TmpTrialBal Where HCode=" & mGroup, DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            Else
                RsDataO.Open "Select IIF(IsNull(Sum(CBal))=True,0,Sum(CBal)) As RTotal From TmpTrialBal Where HCode In (1,13)", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            End If
            If RsDataO.EOF = False Then
                If RsDataO.Fields("RTotal") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,LAmt,RAmt,RName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "'," & _
                    RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ",'Closing Balance')"
                End If
            End If
        End If
    End If
    If Val(.TextMatrix(5, 1)) <> 0 Then
        i = .FindRow(5, 1, 6)
        If i = -1 Then i = .FindRow(58, 1, 6)
        If i > 0 Then
            mGroup = Val(.TextMatrix(i, 6))
            Set RsDataO = Nothing
            If mFinYear = "19-20" Then
                RsDataO.Open "Select IIF(IsNull(Sum(OpCr))=True,0,Sum(OpCr)) As RTotal From TmpTrialBal Where HCode In (58,13,5)", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            Else
                RsDataO.Open "Select IIF(IsNull(Sum(OpCr))=True,0,Sum(OpCr)) As RTotal From TmpTrialBal Where HCode In (58,57,5)", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            End If
            If RsDataO.EOF = False Then
                DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','  Opening Balance'," & _
                RsDataO.Fields("RTotal") & "," & IIf(Round(RsDataO.Fields("RTotal"), 2) >= 0, 0, 1) & ")"
            End If
            Set RsDataO = Nothing
            If mFinYear = "19-20" Then
                RsDataO.Open "Select IIF(Side=HSide,'Add','Less') As TrnType,EName,Sum(IIF(Side=HSide,Amt,Amt*-1)) As Amount From TmpCtDtl Where HCode In (5,13,58) Group By " & _
                "IIF(Side=HSide,'Add','Less'),EName Order By IIF(Side=HSide,'Add','Less')", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            Else
                RsDataO.Open "Select IIF(Side=HSide,'Add','Less') As TrnType,EName,Sum(IIF(Side=HSide,Amt,Amt*-1)) As Amount From TmpCtDtl Where HCode In (5,57,58) Group By " & _
                "IIF(Side=HSide,'Add','Less'),EName Order By IIF(Side=HSide,'Add','Less')", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            End If
            Do While RsDataO.EOF = False
                If RsDataO.Fields("Amount") <> 0 Then
                    DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "',' " & _
                    RsDataO.Fields("TrnType") & ":" & Space(1) & RsDataO.Fields("EName") & "'," & RsDataO.Fields("Amount") & "," & IIf(RsDataO.Fields("TrnType") = "Add", 0, 1) & ")"
                End If
                RsDataO.MoveNext
            Loop
            Set RsDataO = Nothing
            mTotal = SetProfitAll(mAcCodeAll)
            DbDataDB.Execute "Insert InTo TmpNotePrn (SrN,GName,RName,LAmt,RAmt,CName) Values (" & Val(.TextMatrix(i, 1)) & ",'" & .TextMatrix(i, 0) & "','" & _
            IIf(Round(mTotal, 2) >= 0, "Add: Surplus brought from Income and Expenditure Account", "Less: Deficit brought from Income and Expenditure Account") & _
            "'," & mTotal & "," & IIf(Round(mTotal, 2) >= 0, 0, 1) & ",'')"
        End If
    End If
End With
DbDataDB.CommitTrans
With RepPrint
    .Connect = MSCONNECT
    .ReportFileName = App.Path + "\Report\NteRpt.Rpt"
    .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
    .Formulas(1) = "mAdd1='" & RsComp.Fields("Add1") & "'"
    .Formulas(2) = "mAdd2='" & RsComp.Fields("Add2") & "'"
    .Formulas(3) = "mCity='" & RsComp.Fields("City") & ", " & RsComp.Fields("State") & ", " & RsComp.Fields("PinCode") & "'"
    .Formulas(4) = "mPhone='" & RsComp.Fields("Phone") & "'"
    .Formulas(5) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
    .Formulas(6) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
    .Formulas(9) = "mTitle1='Notes to Financial Statements'"
    .Formulas(10) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
    .Formulas(11) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
    .Formulas(12) = "mTrack='" & CStr(InStr(1, mAcCodeAll, ",")) & "'"
    .Formulas(13) = "mType='B'"
    .Formulas(14) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 0) & "'"
    .Action = 1
    For i = 0 To 14
        .Formulas(i) = ""
    Next
End With
mAcCode = Val(LsvClient.TextMatrix(LsvClient.Row, 5))
SetParent
End Sub
Private Sub SetData1()
Dim RsQry As New ADODB.Recordset
Dim i As Integer
With VsfHelp
    .Cols = 8
    .Rows = 1
    .TextMatrix(0, 0) = "EXPENDITURE"
    .ColWidth(0) = 5000
    .TextMatrix(0, 1) = "NOTE"
    .ColWidth(1) = 700
    .TextMatrix(0, 2) = "AMOUNT RS"
    .ColWidth(2) = 1700
    .ColAlignment(2) = flexAlignRightCenter
    .TextMatrix(0, 3) = "INCOME"
    .ColWidth(3) = 5000
    .TextMatrix(0, 4) = "NOTE"
    .ColWidth(4) = 700
    .TextMatrix(0, 5) = "AMOUNT RS"
    .ColWidth(5) = 1700
    .ColAlignment(5) = flexAlignRightCenter
    .ColWidth(6) = 0
    .ColWidth(7) = 0
    .Refresh
    .Rows = 2
    .Row = 1
    RsQry.Open "Select EName,Sum(iif(Side=HSide,Amt,Amt*-1)) As Amount From TmpCtDtl Where AcCode In (" & mAcList & ") And HSide='D' And HType=0 And LCode<>107 Group By EName Order By EName", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        .TextMatrix(.Row, 0) = RsQry.Fields("EName")
        .TextMatrix(.Row, 2) = RsQry.Fields("Amount")
        RsQry.MoveNext
        If RsQry.EOF = False Then
            .Rows = .Rows + 1
            .Row = .Rows - 1
        End If
    Loop
    .Rows = .Rows + 1
    .Row = 1
    Set RsQry = Nothing
    RsQry.Open "Select EName,Sum(iif(Side=HSide,Amt,Amt*-1)) As Amount From TmpCtDtl Where AcCode In (" & mAcList & ") And HSide='C' And HType=0 And LCode<>108 Group By EName Order By EName", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        .TextMatrix(.Row, 3) = RsQry.Fields("EName")
        .TextMatrix(.Row, 5) = RsQry.Fields("Amount")
        RsQry.MoveNext
        If RsQry.EOF = False Then
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            Else
                .Rows = .Rows + 1
                .Row = .Rows - 1
            End If
        Else
            .Rows = .Rows + 1
        End If
    Loop
    Dim mTotal As Double
    mTotal = Round(SetProfitAll(mAcList), 2)
    Set RsQry = Nothing
    .Row = .Rows - 1
    If mTotal > 0 Then
        .TextMatrix(.Row, 0) = "Surplus carried over to Balance Sheet"
       .TextMatrix(.Row, 2) = Format(CStr(mTotal), "0.00")
    Else
        .TextMatrix(.Row, 3) = "Deficit carried over to Balance Sheet"
        .TextMatrix(.Row, 5) = Format(CStr(Abs(mTotal)), "0.00")
    End If
    .Refresh
End With
SetFinalTot
End Sub
Private Sub PlAcPrint()
Dim i As Double
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From TmpPLPrn"
DbDataDB.CommitTrans
DbDataDB.BeginTrans
With VsfHelp
    For i = 1 To .Rows - 1
        DbDataDB.Execute "Insert Into TmpPLPrn (LName,LAmt,LNote,RName,RAmt,RNote,SrN) Values ('" & IIf(Len(.TextMatrix(i, 0)) = 0, "", IIf(Mid(.TextMatrix(i, 0), 1, 1) = "(", "", "To ") + .TextMatrix(i, 0)) & "'," & _
        Val(.TextMatrix(i, 2)) & "," & Val(.TextMatrix(i, 1)) & ",'" & IIf(Len(.TextMatrix(i, 3)) = 0, "", "By " + .TextMatrix(i, 3)) & "'," & Val(.TextMatrix(i, 5)) & "," & _
        Val(.TextMatrix(i, 4)) & "," & i & ")"
    Next
End With
DbDataDB.CommitTrans
If Val(TxtCTotal.Text) <> Val(TxtDTotal.Text) Then
    MsgBox "Income & Expenditure Account does not tally. Report cannot be printed.", vbCritical, "Alert"
    Exit Sub
End If
With RepPrint
    .Connect = MSCONNECT
    .ReportFileName = App.Path + "\Report\IEXRpt.Rpt"
    .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
    .Formulas(1) = "mAdd1='" & RsComp.Fields("Add1") & "'"
    .Formulas(2) = "mAdd2='" & RsComp.Fields("Add2") & "'"
    .Formulas(3) = "mCity='" & RsComp.Fields("City") & ", " & RsComp.Fields("State") & ", " & RsComp.Fields("PinCode") & "'"
    .Formulas(4) = "mPhone='" & RsComp.Fields("Phone") & "'"
    .Formulas(5) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
    .Formulas(6) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
    .Formulas(7) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 0) & "'"
    .Formulas(8) = "mSubHead=''"
    .Formulas(9) = "mTitle1='Income and Expenditure Account for the year ended on " & RsRep.Fields("RtDt") & "'"
    .Formulas(10) = "mPlace='Place : " & RsRep.Fields("RPlace") & "'"
    .Formulas(11) = "mDate='Date : " & RsRep.Fields("RpDt") & "'"
    .Formulas(12) = "mAuditNm='" & RsRep.Fields("AdtName") & "'"
    .Formulas(13) = "mAuditNo='" & RsRep.Fields("AdtNo") & "'"
    .Formulas(14) = "mTSign='Chief Functionary/Trustee'"
    .Formulas(15) = "mUDIN='" & RsRep.Fields("RUDIN") & "'"
    .Formulas(16) = "mClient='" & LsvClient.TextMatrix(LsvClient.Row, 3) & "'"
    .Formulas(17) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
    .Formulas(18) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
    .Formulas(19) = "mLTitle='EXPENDITURE'"
    .Formulas(20) = "mRTitle='INCOME'"
    .Formulas(21) = "mClient1='" & LsvClient.TextMatrix(LsvClient.Row, 3) & "'"
    .Action = 1
    For i = 0 To 21
        .Formulas(i) = ""
    Next
End With
End Sub
Private Sub PrintRecPay()
Dim mTrack As Boolean
Dim mAcName As String
Dim RsQ As New ADODB.Recordset
Dim RsQ1 As New ADODB.Recordset
Dim i As Double
Dim m As Double
mTrack = IIf(InStr(1, mAcList, ",") > 0, True, False)
With VsfHelp
    .Cols = 4
    .Rows = 1
    .Rows = 2
    .TextMatrix(0, 0) = "EXPENDITURE"
    .ColWidth(0) = 3000
    .TextMatrix(0, 1) = "UNDER"
    .ColWidth(1) = 3000
    .TextMatrix(0, 2) = "AMOUNT RS"
    .ColWidth(2) = 1500
    .ColFormat(2) = "0.00"
    .ColAlignment(2) = flexAlignRightCenter
    .ColWidth(3) = 0    'CSIDE
    .Col = 1
    .Row = 1
    .Refresh
End With
With VsfHelp1
    .Cols = 4
    .Rows = 1
    .Rows = 2
    .TextMatrix(0, 0) = "INCOME"
    .ColWidth(0) = 3000
    .TextMatrix(0, 1) = "UNDER"
    .ColWidth(1) = 3000
    .TextMatrix(0, 2) = "AMOUNT RS"
    .ColWidth(2) = 1500
    .ColFormat(2) = "0.00"
    .ColAlignment(2) = flexAlignRightCenter
    .ColWidth(3) = 0    'DSIDE
    .Col = 1
    .Row = 1
    .Refresh
End With
With VsfHelp
    .Col = 2
    i = 1
    Set RsQ = Nothing
    RsQ.Open "Select EntMst.EName,Sum(TmpRpDtl.Amt) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.AcCode In (" & mAcList & _
    ") And TmpRpDtl.Side='D' And TmpRpDtl.ECode=EntMst.ECode And TmpRpDtl.LCode Not In (107,108) And TmpRpDtl.HCode<>14 Group By EntMst.EName Order By EntMst.EName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        If i = .Rows Then .Rows = .Rows + 1
        .TextMatrix(i, 0) = RsQ.Fields("EName")
        .TextMatrix(i, 2) = RsQ.Fields("RTotal")
        RsQ.MoveNext
        i = i + 1
        If RsQ.EOF = True Then
            Exit Do
        ElseIf i = .Rows - 1 Then
            .Row = i
            CmdRow_Click
        End If
    Loop
    Set RsQ1 = Nothing
    RsQ1.Open "Select IIF(AcMst.AcType=2,Space(1)+AcMst.AcName,AcMst.AcName) As AcName,IIF(Mid(LedMst.LName,1,4)='Cash',Space(1)+LedMst.LName,LedMst.LName) As LName," & _
    "LedMst.LCode,TmpTrialBal.DBal From TmpTrialBal,LedMst,AcMst Where TmpTrialBal.AcCode In (" & mAcList & ") And TmpTrialBal.HCode=12 And TmpTrialBal.AcCode=AcMst.AcCode" & _
    " And TmpTrialBal.LCode=LedMst.LCode Order By IIF(AcMst.AcType=2,Space(1)+AcMst.AcName,AcMst.AcName),IIF(Mid(LedMst.LName,1,4)='Cash',Space(1)+LedMst.LName,LedMst.LName)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
'    i = i + 1
    .Rows = .Rows + 1
    Do While RsQ1.EOF = False
        .TextMatrix(i, 0) = "Closing Balance"
        If mTrack = True Then
            mAcName = RsQ1.Fields("AcName")
            .TextMatrix(i, 1) = mAcName
            i = i + 1
            .Rows = .Rows + 1
            Do While mAcName = RsQ1.Fields("AcName")
                .TextMatrix(i, 0) = "Closing Balance"
                .TextMatrix(i, 1) = RsQ1.Fields("LName")
                .TextMatrix(i, 2) = RsQ1.Fields("DBal")
                i = i + 1
                RsQ1.MoveNext
                If RsQ1.EOF = False Then .Rows = .Rows + 1 Else Exit Do
            Loop
        Else
            .TextMatrix(i, 1) = RsQ1.Fields("LName")
            .TextMatrix(i, 2) = RsQ1.Fields("DBal")
            i = i + 1
            RsQ1.MoveNext
            If RsQ1.EOF = False Then .Rows = .Rows + 1
        End If
    Loop
End With
With VsfHelp1
    Set RsQ = Nothing
    RsQ.Open "Select IIF(AcMst.AcType=2,Space(1)+AcMst.AcName,AcMst.AcName) As AcName,IIF(Mid(LedMst.LName,1,4)='Cash',Space(1)+LedMst.LName,LedMst.LName) As LName," & _
    "LedMst.LCode,TmpTrialBal.OpDr From TmpTrialBal,LedMst,AcMst Where TmpTrialBal.AcCode In (" & mAcList & ") And TmpTrialBal.HCode=12 And TmpTrialBal.AcCode=AcMst.AcCode" & _
    " And TmpTrialBal.LCode=LedMst.LCode Order By IIF(AcMst.AcType=2,Space(1)+AcMst.AcName,AcMst.AcName),IIF(Mid(LedMst.LName,1,4)='Cash',Space(1)+LedMst.LName,LedMst.LName)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    i = 1
    Do While RsQ.EOF = False
        .TextMatrix(i, 0) = "Opening Balance"
        If mTrack = True Then
            mAcName = RsQ.Fields("AcName")
            .TextMatrix(i, 1) = mAcName
            i = i + 1
            .Rows = .Rows + 1
            Do While mAcName = RsQ.Fields("AcName")
                If RsQ.Fields("OpDr") <> 0 Then
                    .TextMatrix(i, 0) = "Opening Balance"
                    .TextMatrix(i, 1) = RsQ.Fields("LName")
                    .TextMatrix(i, 2) = RsQ.Fields("OpDr")
                    i = i + 1
                End If
                RsQ.MoveNext
                If RsQ.EOF = False Then .Rows = .Rows + 1 Else Exit Do
            Loop
        Else
            If RsQ.Fields("OpDr") <> 0 Then
                .TextMatrix(i, 1) = mAcName
                .TextMatrix(i, 1) = RsQ.Fields("LName")
                .TextMatrix(i, 2) = RsQ.Fields("OpDr")
                i = i + 1
            End If
            RsQ.MoveNext
            If RsQ.EOF = False Then .Rows = .Rows + 1
        End If
    Loop
    Set RsQ = Nothing
    RsQ.Open "Select EntMst.EName,Sum(TmpRpDtl.Amt) As RTotal From TmpRpDtl,EntMst Where TmpRpDtl.AcCode In (" & mAcList & _
    ") And TmpRpDtl.Side='C' And TmpRpDtl.ECode=EntMst.ECode And TmpRpDtl.LCode Not In (107,108) And TmpRpDtl.HCode<>14 Group By EntMst.EName Order By EntMst.EName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        If i = .Rows Then .Rows = .Rows + 1
        .TextMatrix(i, 0) = RsQ.Fields("EName")
        .TextMatrix(i, 2) = RsQ.Fields("RTotal")
        RsQ.MoveNext
        i = i + 1
        If RsQ.EOF = True Then
            Exit Do
        ElseIf i = .Rows - 1 Then
            .Row = i
            CmdRow1_Click
        End If
    Loop
    .Refresh
End With
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From TmpRecPrn"
m = 0
With VsfHelp
    For i = 1 To .Rows - 1
        If .TextMatrix(i, 0) <> "" Then
            If UCase(Mid(.TextMatrix(i, 0), 1, 4)) = "CLOS" Then
                m = m + 1
                DbDataDB.Execute "Insert Into TmpRecPrn (RName,RAmt,SrN) Values ('',0," & m & ")"
                m = m + 1
                DbDataDB.Execute "Insert Into TmpRecPrn (RName,RAmt,SrN) Values ('By Closing Balance',0," & m & ")"
                Set RsQ = Nothing
                RsQ.Open "Select IIF(AcMst.AcType=2,Space(1)+AcMst.AcName,AcMst.AcName) As AcName,IIF(Mid(LedMst.LName,1,4)='Cash',Space(1)+LedMst.LName,LedMst.LName) As LName," & _
                "LedMst.LCode,RpDtl.Amt From RpDtl,LedMst,AcMst Where RpDtl.AcCode In (" & mAcList & ") And RpDtl.SrN>9999 And RpDtl.AcCode=AcMst.AcCode And  RpDtl.ECode=LedMst.LCode " & _
                "Order By IIF(AcMst.AcType=2,Space(1)+AcMst.AcName,AcMst.AcName),IIF(Mid(LedMst.LName,1,4)='Cash',Space(1)+LedMst.LName,LedMst.LName)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                Do While RsQ.EOF = False
                    If mTrack = True Then
                        mAcName = RsQ.Fields("AcName")
                        m = m + 1
                        DbDataDB.Execute "Insert Into TmpRecPrn (RName,RAmt,SrN) Values ('" & mAcName & "',0," & m & ")"
                            Do While mAcName = RsQ.Fields("AcName")
                            If RsQ.Fields("Amt") <> 0 Then
                                m = m + 1
                                DbDataDB.Execute "Insert Into TmpRecPrn (RName,RAmt,SrN) Values ('" & RsQ.Fields("LName") & "'," & RsQ.Fields("Amt") & "," & m & ")"
                            End If
                            RsQ.MoveNext
                            If RsQ.EOF = True Then
                                i = .Rows - 1
                                Exit Do
                            End If
                        Loop
                    Else
                        mAcName = RsQ.Fields("AcName")
'                        m = m + 1
                        Do While mAcName = RsQ.Fields("AcName")
                            If RsQ.Fields("Amt") <> 0 Then
                                m = m + 1
                                DbDataDB.Execute "Insert Into TmpRecPrn (RName,RAmt,SrN) Values ('" & RsQ.Fields("LName") & "'," & RsQ.Fields("Amt") & "," & m & ")"
                            End If
                            RsQ.MoveNext
                            If RsQ.EOF = True Then
                                i = .Rows - 1
                                Exit Do
                            End If
                        Loop
                    End If
                Loop
            Else
                m = m + 1
                DbDataDB.Execute "Insert Into TmpRecPrn (RName,RAmt,SrN) Values ('By " & .TextMatrix(i, 0) & "'," & Val(.TextMatrix(i, 2)) & "," & m & ")"
            End If
        End If
    Next
End With
With VsfHelp1
    i = 1
    m = 1
    Do While i <= .Rows
        If VsfHelp.Rows >= i Then
            If VsfHelp1.Rows = m Then Exit Do
            If .TextMatrix(m, 0) <> "" Then
                If UCase(Mid(.TextMatrix(m, 0), 1, 4)) = "OPEN" Then
                    If i = 1 Then
                        DbDataDB.Execute "Update TmpRecPrn Set LName='To Opening Balance',LAmt=0 Where SrN=" & i
                        i = i + 1
                    End If
                    DbDataDB.Execute "Update TmpRecPrn Set LName='" & .TextMatrix(m, 1) & "',LAmt=" & Val(.TextMatrix(m, 2)) & " Where SrN=" & i
                    i = i + 1
                    m = m + 1
                Else
                    If UCase(Mid(.TextMatrix(m - 1, 0), 1, 4)) = "OPEN" Then
                        DbDataDB.Execute "Update TmpRecPrn Set LName='',LAmt=0 Where SrN=" & i
                        DbDataDB.Execute "Update TmpRecPrn Set LName='To " & .TextMatrix(m, 0) & "',LAmt=" & Val(.TextMatrix(m, 2)) & " Where SrN=" & i + 1
                        m = m + 1
                        i = i + 1
                    Else
                        DbDataDB.Execute "Update TmpRecPrn Set LName='To " & .TextMatrix(m, 0) & "',LAmt=" & Val(.TextMatrix(m, 2)) & " Where SrN=" & i + 1
                        m = m + 1
                        i = i + 1
                    End If
                End If
            Else
                i = i + 1
                m = m + 1
            End If
        Else
            If .TextMatrix(m, 0) <> "" Then
                If UCase(Mid(.TextMatrix(m, 0), 1, 4)) = "OPEN" Then
                    DbDataDB.Execute "Insert Into TmpRecPrn (LName,LAmt,SrN) Values ('To " & .TextMatrix(m, 1) & "'," & Val(.TextMatrix(m, 2)) & "," & m + VsfHelp1.Rows & ")"
                    m = m + 1
                    i = i + 1
                Else
                    If Mid(.TextMatrix(m - 1, 0), 1, 4) = "OPEN" Then
                        DbDataDB.Execute "Insert Into TmpRecPrn (LName,LAmt,SrN) Values ('',0," & m + VsfHelp1.Rows & ")"
                        m = m + 1
                        i = i + 1
                        DbDataDB.Execute "Insert Into TmpRecPrn (LName,LAmt,SrN) Values ('To " & .TextMatrix(m, 0) & "'," & Val(.TextMatrix(m, 2)) & "," & m + VsfHelp1.Rows & ")"
                        m = m + 1
                        i = i + 1
                    Else
                        DbDataDB.Execute "Insert Into TmpRecPrn (LName,LAmt,SrN) Values ('To " & .TextMatrix(m, 0) & "'," & Val(.TextMatrix(m, 2)) & "," & m + VsfHelp1.Rows & ")"
                        m = m + 1
                        i = i + 1
                    End If
                End If
            Else
                m = m + 1
                i = i + 1
            End If
        End If
    Loop
End With
DbDataDB.CommitTrans
If Val(TxtCTotal.Text) <> Val(TxtDTotal.Text) Then
    MsgBox "Receipt Payment Account does not tally. Report cannot be printed.", vbCritical, "Alert"
    Exit Sub
End If
With RepPrint
    .Connect = MSCONNECT
    .ReportFileName = App.Path + "\Report\RpxRpt.Rpt"
    .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
    .Formulas(1) = "mAdd1='" & RsComp.Fields("Add1") & "'"
    .Formulas(2) = "mAdd2='" & RsComp.Fields("Add2") & "'"
    .Formulas(3) = "mCity='" & RsComp.Fields("City") & ", " & RsComp.Fields("State") & ", " & RsComp.Fields("PinCode") & "'"
    .Formulas(4) = "mPhone='" & RsComp.Fields("Phone") & "'"
    .Formulas(5) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
    .Formulas(6) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
    .Formulas(7) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 0) & "'"
    .Formulas(8) = "mSubHead=''"
    .Formulas(9) = "mTitle1='Memorandum of Receipts and Payments for the year ended on " & RsRep.Fields("RtDt") & "'"
    .Formulas(10) = "mPlace='Place: " & RsRep.Fields("RPlace") & "'"
    .Formulas(11) = "mDate='Date: " & RsRep.Fields("RPDt") & "'"
    .Formulas(12) = "mAuditNm='" & RsRep.Fields("AdtName") & "'"
    .Formulas(13) = "mAuditNo='" & RsRep.Fields("AdtNo") & "'"
    .Formulas(14) = "mUDIN='" & RsRep.Fields("RUDIN") & "'"
    .Formulas(15) = "mClient='" & LsvClient.TextMatrix(LsvClient.Row, 3) & "'"
    .Formulas(16) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
    .Formulas(17) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
    .Formulas(18) = "mLTitle='RECEIPTS'"
    .Formulas(19) = "mRTitle='PAYMENTS'"
    .Formulas(20) = "mTSign='Chief Functionary/Trustee'"
    .Formulas(21) = "mClient1='" & LsvClient.TextMatrix(LsvClient.Row, 3) & "'"
    .Action = 1
    For i = 0 To 21
        .Formulas(i) = ""
    Next
End With
End Sub
Private Sub CmdRow_Click()
With VsfHelp
    .Rows = .Rows + 1
    .Row = .Rows - 1
End With
End Sub
Private Sub CmdRow1_Click()
With VsfHelp1
    .Rows = .Rows + 1
    .Row = .Rows - 1
End With
End Sub
