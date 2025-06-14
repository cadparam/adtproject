VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Begin VB.Form FrmPlAc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Income And Expenditure Account"
   ClientHeight    =   9570
   ClientLeft      =   540
   ClientTop       =   435
   ClientWidth     =   15690
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FPlTrn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9570
   ScaleWidth      =   15690
   Begin VB.Frame FraClientHelp 
      Height          =   4212
      Left            =   18000
      TabIndex        =   6
      Top             =   1920
      Width           =   14750
      Begin VSFlex7Ctl.VSFlexGrid LsvClient 
         Height          =   3852
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   14500
         _cx             =   25576
         _cy             =   6794
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
         FormatString    =   $"FPlTrn.frx":0442
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
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         Height          =   4092
         Index           =   1
         Left            =   0
         Top             =   120
         Width           =   14750
      End
   End
   Begin VB.Frame FraMst 
      Height          =   9500
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   15492
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
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Print"
         Top             =   240
         Width           =   855
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
         Left            =   13200
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   9000
         Width           =   1776
      End
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
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   9000
         Width           =   1776
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
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   732
      End
      Begin VB.CommandButton CmdSearch 
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
         Left            =   1440
         Picture         =   "FPlTrn.frx":048B
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Add"
         Top             =   240
         Width           =   372
      End
      Begin VB.TextBox TxtName 
         Height          =   384
         Left            =   1850
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   4896
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfHelp 
         Height          =   8200
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   15255
         _cx             =   26908
         _cy             =   14464
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
         FormatString    =   $"FPlTrn.frx":0AC1
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
      Begin VB.Label LblTotal 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   6960
         TabIndex        =   10
         Top             =   9000
         Width           =   60
      End
      Begin VB.Label LblCompany 
         Caption         =   "Client Name :"
         BeginProperty Font 
            Name            =   "Cambria"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   8
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1236
      End
      Begin VB.Shape ShpMst 
         BorderWidth     =   2
         Height          =   9375
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   15495
      End
   End
   Begin Crystal.CrystalReport RepPrint 
      Left            =   0
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
Attribute VB_Name = "FrmPlAc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DBWorkTmp As New ADODB.Connection
Dim mAcCode As Double
Dim RsLedger As New ADODB.Recordset

Private Sub CmdSearch_Click()
    FraClientHelp.Left = 180
    LsvClient.SetFocus
End Sub

Private Sub Form_Load()
    DBWorkTmp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source='" & App.Path + "\LocalDB.Mdb'"
    DBWorkTmp.Open
    Me.Left = 50
    Me.Top = 50
    SetCombo
End Sub
Private Function SetCombo()
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select AcMst.AcName,AcMst.FileNo,AcMst.City,'' As ParentNm,GrpMst.GName,AcMst.AcCode,'' As RName,AcMst.PACode From AcMst,GrpMst Where AcMst.Active=-1 And AcMst.PaCode=0 And " & _
"AcMst.AcType=GrpMst.GCode Union All Select AcMst.AcName,AcMst.FileNo,AcMst.City,AcMst1.AcName+IIF(Len(AcMst1.City)>0,', '+AcMst1.City,''),GrpMst.GName,AcMst.AcCode," & _
"'',AcMst.PACode From AcMst,GrpMst,AcMst As AcMst1 Where AcMst.Active=-1 And AcMst.PaCode=AcMst1.AcCode And AcMst.AcType=GrpMst.GCode Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Set LsvClient.DataSource = RsQry
With LsvClient
    .TextMatrix(0, 0) = "NAME"
    .ColWidth(0) = 4800
    .TextMatrix(0, 1) = "FILE NO."
    .ColWidth(1) = 1500
    .TextMatrix(0, 2) = "CITY"
    .ColWidth(2) = 2000
    .TextMatrix(0, 3) = "PARENT NAME"
    .ColWidth(3) = 4800
    .TextMatrix(0, 4) = "TYPE"
    .ColWidth(4) = 1000
    .TextMatrix(0, 5) = "ACCODE"
    .ColWidth(5) = 0
    .TextMatrix(0, 6) = "MAIN PARENT"
    .ColWidth(6) = 4000
    .Col = 1
    .ExtendLastCol = True
    If .Rows > 1 Then .Row = 1
    .FontBold = False
    .Refresh
    Set RsQry = Nothing
    RsQry.Open "Select AcMst.AcCode,AcMst2.AcName+', '+AcMst2.City As RName From AcMst,AcMst As AcMst1,AcMst As AcMst2 Where AcMst.PaCode=AcMst1.AcCode And " & _
    "AcMst1.PaCode=AcMst2.AcCode Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        .Row = 1
        .Row = .FindRow(RsQry.Fields("AcCode"), , 5)
        If .Row > 1 Then .TextMatrix(.Row, 6) = RsQry.Fields("RName")
        RsQry.MoveNext
    Loop
End With
End Function
Private Sub ClearText()
    TxtName.Text = ""
    mAcCode = 0
    FraClientHelp.Left = 18000
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MDIWork.PctMdi.Visible = True
    Unload Me
End Sub

Private Sub LsvClient_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtName.Text = LsvClient.TextMatrix(LsvClient.Row, 0)
    mAcCode = Val(LsvClient.TextMatrix(LsvClient.Row, 5))
    FraClientHelp.Left = 18000
    Dim RsQ As New ADODB.Recordset
    DBWorkTmp.BeginTrans
    DBWorkTmp.Execute "Delete From TmpCtDtl"
    DBWorkTmp.CommitTrans
    DBWorkTmp.BeginTrans
    Set RsQ = Nothing
    RsQ.Open "Select * From QCtDtl Where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        DBWorkTmp.Execute "Insert InTo TmpCtDtl (HType,HSide,AcCode,HCode,LCode,ECode,EName,Side,Amt) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HSide") & "'," & _
        RsQ.Fields("AcCode") & "," & RsQ.Fields("HCode") & "," & RsQ.Fields("LCode") & "," & RsQ.Fields("ECode") & ",'" & RsQ.Fields("EName") & "','" & RsQ.Fields("Side") & _
        "'," & RsQ.Fields("Amt") & ")"
        RsQ.MoveNext
    Loop
    DBWorkTmp.CommitTrans
    SetData
    DbDataDB.BeginTrans
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'INEXP_IND','VIEW','" & Date & "','" & Time & "')"
    DbDataDB.CommitTrans
    VsfHelp.SetFocus
End If
End Sub
Private Sub Tlbsav_Click(Index As Integer)
    If TlbSav(Index).ToolTipText = "Exit" Then
        DBWorkTmp.Close
        Unload Me
    ElseIf TlbSav(Index).ToolTipText = "Print" Then
        PrintRec
        DbDataDB.BeginTrans
'            DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'INEXP_IND','PRINT','" & Date & "','" & Time & "')"
        DbDataDB.CommitTrans
    End If
End Sub
Private Sub SetData()
Dim RsQry As New ADODB.Recordset
Dim RsQ As New ADODB.Recordset
RsQry.Open "Select EName,Sum(IIF(HSide=Side,Amt,Amt*-1)) AS TAmt From TmpCtDtl Where AcCode=" & mAcCode & " And HType=0 And HSide='D' Group By EName Order By EName", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
With VsfHelp
    .Cols = 4
    .Rows = 1
    .TextMatrix(0, 0) = "EXPENDITURE"
    .ColWidth(0) = 5000
    .TextMatrix(0, 1) = "AMOUNT RS"
    .ColWidth(1) = 1700
    .ColAlignment(1) = flexAlignRightCenter
    .TextMatrix(0, 2) = "INCOME"
    .ColWidth(2) = 5000
    .TextMatrix(0, 3) = "AMOUNT RS"
    .ColWidth(3) = 1700
    .ColAlignment(3) = flexAlignRightCenter
    .Refresh
    .Rows = 2
    .Row = 1
    Do While RsQry.EOF = False
        .TextMatrix(.Row, 0) = RsQry.Fields("EName")
        .TextMatrix(.Row, 1) = RsQry.Fields("TAmt")
        RsQry.MoveNext
        .Rows = .Rows + 1
        .Row = .Rows - 1
    Loop
    .Rows = .Rows + 1
    .Row = 1
    Set RsQry = Nothing
    RsQry.Open "Select EName,Sum(IIF(HSide=Side,Amt,Amt*-1)) as TAmt From TmpCtDtl Where AcCode=" & mAcCode & " And HType=0 And HSide='C' Group By EName Order By EName", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQry.EOF = False
        .TextMatrix(.Row, 2) = RsQry.Fields("EName")
        .TextMatrix(.Row, 3) = RsQry.Fields("TAmt")
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
    Dim mTotal As Double
    mTotal = SetProfit(mAcCode)
    .Row = .Rows - 1
    If mTotal > 0 Then
        .TextMatrix(.Row, 0) = "Surplus carried over to Balance Sheet"
       .TextMatrix(.Row, 1) = Format(CStr(mTotal), "0.00")
    Else
        .TextMatrix(.Row, 2) = "Deficit carried over to Balance Sheet"
        .TextMatrix(.Row, 3) = Format(CStr(Abs(mTotal)), "0.00")
    End If
End With
SetFinalTot
End Sub

Private Sub SetFinalTot()
Dim i As Double
TxtDTotal.Text = "0.00"
TxtCTotal.Text = "0.00"
With VsfHelp
    For i = 1 To .Rows - 1
        TxtDTotal.Text = Val(TxtDTotal.Text) + Val(.TextMatrix(i, 1))
        TxtCTotal.Text = Val(TxtCTotal.Text) + Val(.TextMatrix(i, 3))
    Next
End With
TxtDTotal.Text = Format(TxtDTotal.Text, "0.00")
TxtCTotal.Text = Format(TxtCTotal.Text, "0.00")
If Val(TxtCTotal.Text) - Val(TxtDTotal.Text) = 0 Then LblTotal.Caption = "" Else LblTotal.Caption = "Total mismatching Of Rs.-->" & Format(CStr(Abs(Val(TxtCTotal.Text) - Val(TxtDTotal.Text))), "0.00")
End Sub

Private Sub PrintRec()
Dim i As Double
Dim mHead As String
Dim mClient1 As String
Dim RsRep As New ADODB.Recordset
If Val(TxtCTotal.Text) <> Val(TxtDTotal.Text) Then
    MsgBox "Income Expenditure Account does not tally. Report cannot be printed.", vbCritical, "Alert"
    Exit Sub
End If
RsRep.Open "Select RepDtl.*,AdtMst.AdtName From RepDtl,AdtMst Where AcCode=" & mAcCode & " And RepDtl.AdtNo=AdtMst.AdtNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsRep.EOF = False Then
    If Len(RsRep.Fields("RUDIN")) = 0 Then
        MsgBox "UDIN not issued. Non-UDIN Report will be printed.", vbInformation, "Information"
    End If
Else
    MsgBox "Signing Information not available. Report cannot be printed.", vbCritical, "Alert"
    Exit Sub
End If
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From TmpPLPrn"
DbDataDB.CommitTrans
DbDataDB.BeginTrans
With VsfHelp
    For i = 1 To .Rows - 1
        DbDataDB.Execute "Insert Into TmpPLPrn (LName,LAmt,RName,RAmt,SrN) Values ('" & IIf(Len(.TextMatrix(i, 0)) = 0, "", IIf(Mid(.TextMatrix(i, 0), 1, 1) = "(", "", "To ") + .TextMatrix(i, 0)) & "'," & _
        Val(.TextMatrix(i, 1)) & ",'" & IIf(Len(.TextMatrix(i, 2)) = 0, "", "By " + .TextMatrix(i, 2)) & "'," & Val(.TextMatrix(i, 3)) & "," & _
        i & ")"
    Next
End With
DbDataDB.CommitTrans
With RepPrint
    .Connect = MSCONNECT
    .ReportFileName = App.Path + "\Report\IeXRpt.Rpt"
    .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
    .Formulas(1) = "mAdd1='" & RsComp.Fields("Add1") & "'"
    .Formulas(2) = "mAdd2='" & RsComp.Fields("Add2") & "'"
    .Formulas(3) = "mCity='" & RsComp.Fields("City") & ", " & RsComp.Fields("State") & ", " & RsComp.Fields("PinCode") & "'"
    .Formulas(4) = "mPhone='" & RsComp.Fields("Phone") & "'"
    .Formulas(5) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
    .Formulas(6) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
    If Len(LsvClient.TextMatrix(LsvClient.Row, 6)) > 0 Then
        .Formulas(7) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 6) & "'"
        .Formulas(8) = "mSubHead='" & LsvClient.TextMatrix(LsvClient.Row, 3) & "'"
        mClient1 = CStr(LsvClient.TextMatrix(LsvClient.Row, 6))
    ElseIf Len(LsvClient.TextMatrix(LsvClient.Row, 3)) > 0 Then
        .Formulas(7) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 3) & "'"
        .Formulas(8) = "mSubHead=''"
        mClient1 = CStr(LsvClient.TextMatrix(LsvClient.Row, 3))
    Else
        .Formulas(7) = "mHead='" & LsvClient.TextMatrix(LsvClient.Row, 0) + ", " + LsvClient.TextMatrix(LsvClient.Row, 2) & "'"
        .Formulas(8) = "mSubHead=''"
        mClient1 = CStr(LsvClient.TextMatrix(LsvClient.Row, 0)) & ", " & CStr(LsvClient.TextMatrix(LsvClient.Row, 2))
    End If
    .Formulas(9) = "mTitle1='Income and Expenditure Account for the year ending on " & RsRep.Fields("RTDt") & "'"
    .Formulas(10) = "mPlace='Place: " & RsRep.Fields("RPlace") & "'"
    .Formulas(11) = "mDate='Date: " & RsRep.Fields("RpDt") & "'"
    .Formulas(12) = "mAuditNm='" & RsRep.Fields("AdtName") & "'"
    .Formulas(13) = "mAuditNo='" & RsRep.Fields("AdtNo") & "'"
    .Formulas(14) = "mUDIN='" & RsRep.Fields("RUDIN") & "'"
    .Formulas(15) = "mClient='" & IIf(Len(LsvClient.TextMatrix(LsvClient.Row, 3)) > 0, TxtName.Text + IIf(Len(LsvClient.TextMatrix(LsvClient.Row, 2)) > 0, ", " + LsvClient.TextMatrix(LsvClient.Row, 2), ""), "") & "'"
    .Formulas(16) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
    .Formulas(17) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
    .Formulas(18) = "mLTitle='EXPENDITURE'"
    .Formulas(19) = "mRTitle='INCOME'"
    If Len(RsRep.Fields("RUDIN")) = 0 Then
        If Len(LsvClient.TextMatrix(LsvClient.Row, 6)) > 0 Then
            mHead = LsvClient.TextMatrix(LsvClient.Row, 6)
        ElseIf Len(LsvClient.TextMatrix(LsvClient.Row, 3)) > 0 Then
            mHead = LsvClient.TextMatrix(LsvClient.Row, 3)
        Else
            mHead = LsvClient.TextMatrix(LsvClient.Row, 0) & ", " & LsvClient.TextMatrix(LsvClient.Row, 2)
        End If
    .Formulas(20) = "mSub='This Income & Expenditure Account is issued for the sole purpose of internal use of " & mHead & " and should not be presented before any third parties/agencies/authorities without our consent.'"
    End If
        If LsvClient.TextMatrix(LsvClient.Row, 4) = "Library" Then
        .Formulas(21) = "mTSign='Trustee/Secretary'"
    ElseIf LsvClient.TextMatrix(LsvClient.Row, 4) = "FC" Then
        .Formulas(21) = "mTSign='Chief Functionary/Trustee'"
    ElseIf LsvClient.TextMatrix(LsvClient.Row, 4) = "School" Then
        .Formulas(21) = "mTSign='Principal/In-charge'"
    Else: .Formulas(21) = "mTSign='Trustee/In-charge'"
    End If
    If LsvClient.TextMatrix(LsvClient.Row, 4) = "School" Then .Formulas(22) = "mClient1='" & TxtName.Text + IIf(Len(LsvClient.TextMatrix(LsvClient.Row, 2)) > 0, ", " + LsvClient.TextMatrix(LsvClient.Row, 2), "") & "'" Else .Formulas(22) = "mClient1='" & CStr(mClient1) & "'"
    .Action = 1
    For i = 0 To 22
        .Formulas(i) = ""
    Next
End With
End Sub
