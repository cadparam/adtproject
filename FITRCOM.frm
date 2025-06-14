VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Begin VB.Form FrmITRCom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ITR Computation"
   ClientHeight    =   8880
   ClientLeft      =   540
   ClientTop       =   435
   ClientWidth     =   13140
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FITRCOM.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   13140
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
         FormatString    =   $"FITRCOM.frx":0442
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
      Height          =   8775
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   12972
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
         Index           =   4
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Print"
         Top             =   240
         Width           =   855
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
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cancel"
         Top             =   240
         Width           =   900
      End
      Begin VB.CommandButton TlbSav 
         Caption         =   "&Save"
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
         Index           =   0
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Save"
         Top             =   240
         Width           =   732
      End
      Begin VB.Frame FraMainExp 
         BackColor       =   &H00F7E0FE&
         Height          =   5412
         Left            =   18000
         TabIndex        =   9
         Top             =   4000
         Width           =   15612
         Begin VB.CommandButton CmpMEClose 
            Caption         =   "Close"
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
            Left            =   6600
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Cancel"
            Top             =   4920
            Width           =   975
         End
         Begin VSFlex7Ctl.VSFlexGrid VsfMainExport 
            Height          =   5052
            Left            =   0
            TabIndex        =   11
            Top             =   240
            Width           =   15480
            _cx             =   27305
            _cy             =   8911
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
            BackColorBkg    =   15007437
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
            FormatString    =   $"FITRCOM.frx":048B
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
            Height          =   5292
            Index           =   5
            Left            =   0
            Top             =   120
            Width           =   15612
         End
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
         Left            =   10485
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   8280
         Width           =   2220
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
         Left            =   9120
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
         Picture         =   "FITRCOM.frx":04D4
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
         Height          =   7500
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   12732
         _cx             =   22458
         _cy             =   13229
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
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
         FormatString    =   $"FITRCOM.frx":0B0A
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
      Begin VB.CommandButton TlbSav 
         Caption         =   "&Delete"
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
         Left            =   11640
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Delete"
         Top             =   240
         Width           =   990
      End
      Begin VSFlex7Ctl.VSFlexGrid VsfHelp1 
         Height          =   255
         Left            =   11640
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   735
         _cx             =   1291
         _cy             =   444
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
         FormatString    =   $"FITRCOM.frx":0B53
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
      Begin VB.Label LblNTi 
         Caption         =   "Net Taxable Income (1 - (2 + 3))"
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
         Index           =   0
         Left            =   7320
         TabIndex        =   16
         Top             =   8280
         Width           =   3000
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
         Height          =   8655
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   12975
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
Attribute VB_Name = "FrmITRCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DBLocalDB As New ADODB.Connection
Dim RsOpDtl As New ADODB.Recordset
Dim RsQ As New ADODB.Recordset
Dim mAcCode As Double
Dim mAcList As String
Dim CForm As Integer
'Dim TaxInc As Double
Dim IntInc As Double
Dim OthExp As Double
Dim TotDon As Double

Private Sub CmdSearch_Click()
    FraClientHelp.Left = 180
    LsvClient.SetFocus
End Sub

Private Sub Form_Load()
    Me.Left = 50
    Me.Top = 50
    DBLocalDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source='" & App.Path + "\LocalDB.Mdb'"
    DBLocalDB.Open
    SetCombo
    SetTool (True)
End Sub
Private Function SetCombo()
Dim RsQry As New ADODB.Recordset
RsQry.Open "Select AcMst.AcName,AcMst.FileNo,AcMst.City,'' As ParentNm,GrpMst.GName,AcMst.AcCode From AcMst,GrpMst Where AcMst.Active=-1 And IsNull(AcMst.TPan)=False And" & _
" AcMst.PaCode=0 And AcMst.AcType=1 And AcMst.AcType=GrpMst.GCode Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
    .Col = 1
    .ExtendLastCol = True
    If .Rows > 1 Then .Row = 1
    .FontBold = False
    .Refresh
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
    Set RsQ = Nothing
        RsQ.Open "Select CForm from ITDtl where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
        If RsQ.EOF = False Then CForm = RsQ.Fields("CForm") Else CForm = 1
    FraClientHelp.Left = 18000
    SetParent
    Set RsOpDtl = Nothing
    RsOpDtl.Open "Select LedMst.HCode, OpDtl.* From OpDtl, LedMst Where OpDtl.AcCode In (" & mAcList & ") And (LedMst.HCode In (2,5) Or LedMst.LCode In (-1,-2,-3,-4))", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Set RsQ = Nothing
    DBLocalDB.BeginTrans
    DBLocalDB.Execute "Delete * From TmpCtDtl"
    DBLocalDB.Execute "Delete * From TmpTrialBal"
    DBLocalDB.CommitTrans
    DBLocalDB.BeginTrans
    Set RsQ = Nothing
    RsQ.Open "Select * From QCtDtl Where AcCode In (" & mAcList & ")", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        DBLocalDB.Execute "Insert InTo TmpCtDtl (HType,HSide,AcCode,HCode,LCode,ECode,EName,Side,Amt) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HSide") & "'," & _
        RsQ.Fields("AcCode") & "," & RsQ.Fields("HCode") & "," & RsQ.Fields("LCode") & "," & RsQ.Fields("ECode") & ",'" & RsQ.Fields("EName") & "','" & RsQ.Fields("Side") & _
        "'," & RsQ.Fields("Amt") & ")"
        RsQ.MoveNext
    Loop
    Set RsQ = Nothing
    RsQ.Open "Select * From QTrialBal Where AcCode In (" & mAcList & ") And HCode In (2,5,7,12,59)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        DBLocalDB.Execute "Insert InTo TmpTrialBal (HType,HSide,AcCode,HCode,LCode,OpDr,OpCr,ADr,ACr,DBal,CBal) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HSide") & "'," & _
        RsQ.Fields("AcCode") & "," & RsQ.Fields("HCode") & "," & RsQ.Fields("LCode") & "," & RsQ.Fields("OpDr") & "," & RsQ.Fields("OpCr") & "," & RsQ.Fields("ADr") & _
        "," & RsQ.Fields("ACr") & "," & RsQ.Fields("DBal") & "," & RsQ.Fields("CBal") & ")"
        RsQ.MoveNext
    Loop
    DBLocalDB.CommitTrans
    SetData
    ShowData
    DbDataDB.BeginTrans
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'ITR_COMP','VIEW','" & Date & "','" & Time & "')"
    DbDataDB.CommitTrans
    VsfHelp.SetFocus
End If
End Sub
Private Sub Tlbsav_Click(Index As Integer)
Select Case TlbSav(Index).ToolTipText
    Case "Delete"
        If MsgBox("Sure To Delete ?", vbInformation + vbYesNo) = vbYes Then
            DbDataDB.BeginTrans
            DbDataDB.Execute "Delete From ITDtl Where AcCode=" & mAcCode
'            DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'ITR_COMP','DELETE','" & Date & "','" & Time & "')"
            DbDataDB.CommitTrans
        End If
        SetData
    Case "Save"
        If TxtName.Text = "" Then
            MsgBox "Sorry! Not Allowed.", vbInformation, "Black Data Error"
            TxtName.SetFocus
        Else
            If MsgBox("Are you sure to save?", vbInformation + vbYesNo, "Confirmation") = vbYes Then SaveData
            DbDataDB.BeginTrans
'                DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'ITR_COMP','SAVE_DATA','" & Date & "','" & Time & "')"
            DbDataDB.CommitTrans
            SetData
            ClearText
            SetCombo
            VsfHelp.SetFocus
        End If
    Case "Cancel"
        SetData
        ClearText
        SetCombo
        VsfHelp.SetFocus
    Case "Print"
        If TxtName.Text <> "" Then
            If TlbSav(0).Enabled = True Then
                If MsgBox("Save data before printing?", vbInformation + vbYesNo, "Confirmation") = vbYes Then SaveData
                DbDataDB.BeginTrans
'                    DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'ITR_COMP','SAVE_DATA','" & Date & "','" & Time & "')"
                DbDataDB.CommitTrans
            End If
        End If
        PrintRec
        DbDataDB.BeginTrans
'            DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'ITR_COMP','PRINT','" & Date & "','" & Time & "')"
        DbDataDB.CommitTrans
        SetData
        ClearText
        SetCombo
    Case "Exit"
        DBLocalDB.Close
        Unload Me
End Select
End Sub
Private Function SetTool(ByVal mVal As Boolean)
    TlbSav(2).Enabled = True
End Function
Private Sub SetData()
With VsfHelp
If CForm = 0 Then
    .Cols = 6
    .ColWidth(5) = 0
    .Rows = 1
    .TextMatrix(0, 0) = "SR."
    .ColWidth(0) = 400
    .TextMatrix(0, 1) = "PARTICULARS"
    .ColWidth(1) = 7500
    .TextMatrix(0, 2) = "TOTAL 1"
    .ColWidth(2) = 1500
    .ColFormat(2) = "0.00"
    .ColAlignment(2) = flexAlignRightCenter
    .TextMatrix(0, 3) = "TOTAL 2"
    .ColWidth(3) = 1500
    .ColFormat(3) = "0.00"
    .TextMatrix(0, 4) = "TOTAL 3"
    .ColWidth(4) = 1500
    .ColFormat(4) = "0.00"
    .Refresh
    .Rows = 2
    .Row = 1
    .TextMatrix(.Row, 0) = "A"
    .TextMatrix(.Row, 1) = "Details of Aggregate Income"
    RowInc
    .TextMatrix(.Row, 0) = "1"
    .TextMatrix(.Row, 1) = "Income as per Income & Expenditure A/c"
    RowInc
    .TextMatrix(.Row, 1) = "i). Rent Income"
    RowInc
    .TextMatrix(.Row, 1) = "ii). Interest Income"
    RowInc
    .TextMatrix(.Row, 1) = "iii). Dividend Income"
    RowInc
    .TextMatrix(.Row, 1) = "iv). Donations in Cash (Other than Corpus)"
    RowInc
    .TextMatrix(.Row, 1) = "a). Local"
    RowInc
    .TextMatrix(.Row, 1) = "b). Foreign (FC)"
    RowInc
    .TextMatrix(.Row, 1) = "v). Donations in Kind (Other than Corpus)"
    RowInc
    .TextMatrix(.Row, 1) = "a). Local"
    RowInc
    .TextMatrix(.Row, 1) = "b). Foreign (FC)"
    RowInc
    .TextMatrix(.Row, 1) = "vi). Grant Income"
    RowInc
    .TextMatrix(.Row, 1) = "a). Government Grants"
    RowInc
    .TextMatrix(.Row, 1) = "b). CSR Grants from Companies"
    RowInc
    .TextMatrix(.Row, 1) = "c). Other Grants"
    RowInc
    .TextMatrix(.Row, 1) = "i). Local"
    RowInc
    .TextMatrix(.Row, 1) = "ii). Foreign (FC)"
    RowInc
    .TextMatrix(.Row, 1) = "vii). Agricultural Income"
    RowInc
    .TextMatrix(.Row, 1) = "viii).  Gain on Transfer of Capital Asset"
    RowInc
    .TextMatrix(.Row, 1) = "ix).    Income from Other Sources"
    RowInc
    .TextMatrix(.Row, 1) = "Total Income as per Income & Expenditure A/c"
    RowInc
    .TextMatrix(.Row, 0) = "2"
    .TextMatrix(.Row, 1) = "Corpus Donations (in Cash or Kind)"
    RowInc
    .TextMatrix(.Row, 1) = "i). Corpus Donation Received"
    RowInc
    .TextMatrix(.Row, 1) = "a). Local"
    RowInc
    .TextMatrix(.Row, 1) = "b). Foreign (FC)"
    RowInc
    .TextMatrix(.Row, 1) = "ii).    Less: Exemption u/s 11(1)(d)"
    RowInc
    .TextMatrix(.Row, 1) = "Net Corpus Donation Income (i - ii)"
    RowInc
    .TextMatrix(.Row, 0) = "3"
    .TextMatrix(.Row, 1) = "Details of Unutilised Accumulation of Previous Years"
    RowInc
    .TextMatrix(.Row, 1) = "i). Unutilised Accumulation u/s 11(2) (5 years Period Completed)"
    RowInc
    .TextMatrix(.Row, 1) = "ii). Unutilised Option under Explanation (2) to 11(1)"
    RowInc
    .TextMatrix(.Row, 1) = "Total Unutilised Accumulation added to Income (i + ii)"
    RowInc
    .TextMatrix(.Row, 1) = "Aggregate Income (1 + 2 + 3)"
    RowInc
    RowInc
    .TextMatrix(.Row, 0) = "B"
    .TextMatrix(.Row, 1) = "Details of Amount applied to Charitable Purposes"
    RowInc
    .TextMatrix(.Row, 0) = "1"
    .TextMatrix(.Row, 1) = "Expenditure on Revenue Account"
    RowInc
    .TextMatrix(.Row, 1) = "i). Total Expenditure as per Income & Expenditure A/c"
    RowInc
    .TextMatrix(.Row, 1) = "ii). Less: Amount transferred to Reserves/Specific Funds"
    RowInc
    .TextMatrix(.Row, 1) = "iii). Less: Depreciation on/Write-off of Assets provided during the year"
    RowInc
    .TextMatrix(.Row, 1) = "Amount applied under Revenue Account (i - (ii + iii))"
    RowInc
    .TextMatrix(.Row, 0) = "2"
    .TextMatrix(.Row, 1) = "Expenditure on Capital Account"
    RowInc
    .TextMatrix(.Row, 1) = "i). Purchase of Capital Assets during the year"
    RowInc
    .TextMatrix(.Row, 1) = "ii). Less: Sale of Capital Assets during the year"
    RowInc
    .TextMatrix(.Row, 1) = "iii). Less: Gain on Transfer of Capital Asset"
    RowInc
    .TextMatrix(.Row, 1) = "Amount applied under Capital Account (i - (ii + iii))"
    RowInc
    .TextMatrix(.Row, 0) = "3"
    .TextMatrix(.Row, 1) = "Amount re-qualifying as Application of Income"
    RowInc
    .TextMatrix(.Row, 1) = "i). Repayment of Borrowed Funds"
    RowInc
    .TextMatrix(.Row, 1) = "ii). Repatriation to Corpus Fund"
    RowInc
    .TextMatrix(.Row, 1) = "Total Amount re-qualifying as Application of Income (i + ii)"
    RowInc
    .TextMatrix(.Row, 0) = "4"
    .TextMatrix(.Row, 1) = "Amount not qualifying as Application of Income"
    RowInc
    .TextMatrix(.Row, 1) = "i). Amount applied out of Borrowed Funds"
    RowInc
    .TextMatrix(.Row, 1) = "ii). Amount applied out of Corpus Fund"
    RowInc
    .TextMatrix(.Row, 1) = "Total Amount not qualifying as Application of Income (i + ii)"
    RowInc
    .TextMatrix(.Row, 0) = "5"
    .TextMatrix(.Row, 1) = "Amount applied out of Accumulated Income"
    RowInc
    .TextMatrix(.Row, 1) = "i). Amount applied out of Option under Explanation (2) to 11(1)"
    RowInc
    .TextMatrix(.Row, 1) = "ii). Amount applied out of Accumulation u/s 11(2)"
    RowInc
    .TextMatrix(.Row, 1) = "Total amount applied out of Accumulated Income (i + ii)"
    RowInc
    .TextMatrix(.Row, 1) = "Total Application of Income for Charitable Purposes ((1 + 2 + 3) - (4 + 5))"
    RowInc
    RowInc
    .TextMatrix(.Row, 0) = "C"
    .TextMatrix(.Row, 1) = "Calculation of Taxable Income"
    RowInc
    .TextMatrix(.Row, 0) = "1"
    .TextMatrix(.Row, 1) = "Surplus/Deficit during the year (A - B)"
    RowInc
    .TextMatrix(.Row, 0) = "2"
    .TextMatrix(.Row, 1) = "Amount set apart out of Income (upto 15%)"
    RowInc
    .TextMatrix(.Row, 0) = "3"
    .TextMatrix(.Row, 1) = "Accumulation u/s 11(2) for 5 years (Form 10 to be Filed)"
    RowInc
    .TextMatrix(.Row, 0) = "4"
    .TextMatrix(.Row, 1) = "Option under Explanation (2) to 11(1) (Form 9A to be Filed)"
    RowInc
    .TextMatrix(.Row, 1) = "Amount deemed to be applied to Charitable Purposes"
    .Cell(flexcpBackColor, 3, 0, 5, .Cols - 1) = RGB(128, 255, 128)
    .Cell(flexcpBackColor, 20, 0, 20, .Cols - 1) = RGB(128, 255, 128)
    .Cell(flexcpBackColor, 29, 0, 30, .Cols - 1) = RGB(128, 255, 128)
    .Cell(flexcpBackColor, 46, 0, 47, .Cols - 1) = RGB(128, 255, 128)
    .Cell(flexcpBackColor, 50, 0, 51, .Cols - 1) = RGB(128, 255, 128)
    .Cell(flexcpBackColor, 55, 0, 55, .Cols - 1) = RGB(128, 255, 128)
    .Cell(flexcpBackColor, 62, 0, 63, .Cols - 1) = RGB(128, 255, 128)
    .Refresh
Else
    .Cols = 4
    .Rows = 1
    .TextMatrix(0, 0) = "PARTICULARS"
    .ColWidth(0) = 7500
    .TextMatrix(0, 1) = "AMOUNT (Rs.)"
    .ColWidth(1) = 1500
    .ColFormat(1) = "0.00"
    .ColAlignment(1) = flexAlignRightCenter
    .TextMatrix(0, 2) = "AMOUNT (Rs.)"
    .ColWidth(2) = 1500
    .ColFormat(2) = "0.00"
    .ColAlignment(2) = flexAlignRightCenter
    .ColWidth(3) = 0
    .Refresh
    .Rows = 2
    .Row = 1
    .TextMatrix(.Row, 0) = "A. Details of Aggregate Income" '1
    .TextMatrix(.Row, 3) = 1
    RowInc
    .TextMatrix(.Row, 0) = Space(5) + "1. Schedule AI" '2
    .TextMatrix(.Row, 3) = 1
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "i. Rent Income" '3
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "ii. Dividend Income" '4
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "iii. Interest Income" '5
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "iv. Agricultural Income" '6
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "v. Net Consideration on Transfer of Capital Asset" '7
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "vi. Any Other Income (Refer Attached Sheet)" '8
    .TextMatrix(.Row, 3) = 3
    RowInc
    .TextMatrix(.Row, 0) = Space(5) + "Income under Schedule AI (i to vi)" '9
    .TextMatrix(.Row, 3) = 1
    RowInc
    RowInc
    .TextMatrix(.Row, 0) = Space(5) + "2. Schedule VC" '11
    .TextMatrix(.Row, 3) = 1
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "i. Non-Corpus Donations (Domestic)" '12
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "a. Government Grants" '13
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "b. CSR Grants from Companies" '14
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "c. Other Grants" '15
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "d. Other Donation (Local)" '16
    .TextMatrix(.Row, 3) = 3
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "ii. Non-Corpus Donations (Foreign)" '17
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "a. Other Donation (Foreign)" '18
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "b. Other Grants (Foreign)" '19
    .TextMatrix(.Row, 3) = 3
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "iii. Corpus Donations (Domestic)" '20
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "Other than received for renovation or repair u/s 80G(2)(b)" '21
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "iii. Corpus Donations (Foreign)" '22
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "Other than received for renovation or repair u/s 80G(2)(b) (Foreign)" '23
    .TextMatrix(.Row, 3) = 3
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "Less: Exemption u/s 11(1)(d)" '24
    .TextMatrix(.Row, 3) = 3
    RowInc
    .TextMatrix(.Row, 0) = Space(5) + "Income under Schedule VC (i + ii + iii)" '25
    .TextMatrix(.Row, 3) = 1
    RowInc
    RowInc
    .TextMatrix(.Row, 0) = Space(5) + "3. Details of Unutilised Accumulation of Previous Years" '27
    .TextMatrix(.Row, 3) = 1
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "i. Unutilised Accumulation u/s 11(2) (5 years Period Completed)" '28
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "ii. Unutilised Option under Explanation (2) to 11(1)" '29
    .TextMatrix(.Row, 3) = 3
    RowInc
    .TextMatrix(.Row, 0) = Space(5) + "Total Unutilised Accumulation added to Income (i + ii)" '30
    .TextMatrix(.Row, 3) = 1
    RowInc
    RowInc
    .TextMatrix(.Row, 0) = "Aggregate Income (1 + 2 + 3)" '32
    .TextMatrix(.Row, 3) = 5
    RowInc
    RowInc
    .TextMatrix(.Row, 0) = "B. Details of Amount applied to Charitable Purposes" '34
    .TextMatrix(.Row, 3) = 1
    RowInc
    .TextMatrix(.Row, 0) = Space(5) + "1. Schedule ER" '35
    .TextMatrix(.Row, 3) = 1
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "A. Administrative Expenses" '36
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "i. Rents" '37
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "ii. Repairs and Maintenance" '38
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "iii. Compensation to Employees" '39
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "iv. Insurance" '40
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "v. Workmen and Staff Welfare Expenses" '41
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "vi. Entertainment and Hospitality" '42
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "vii. Advertisement" '43
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "viii. Professional/Consultancy Fees/Fees for Technical Services" '44
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "ix. Conveyance and Travelling Expenses other than Foreign Travel" '45
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "x. Remuneration to Trustees" '46
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "xi. Rates & Taxes" '47
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "ix. Interest" '48
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "ix. Audit Fees" '49
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "ix. Other Expenses" '50
    .TextMatrix(.Row, 3) = 3
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "B. Expenditure on Objects of the Trust" '51
    .TextMatrix(.Row, 3) = 1
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "i. Donation" '52
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(20) + "a. Corpus" '53
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(20) + "a. Other than Corpus" '54
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "ii. Religious" '55
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "iii. Relief of Poor" '56
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "iv. Educational" '57
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "v. Yoga" '58
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "vi. Medical Relief" '59
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "vii. Preservation of Environment" '60
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "viii. Preservation of Monuments, etc." '61
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "ix. General Public Utility" '62
    .TextMatrix(.Row, 3) = 3
    RowInc
    .TextMatrix(.Row, 0) = Space(5) + "Total Application of Funds on Revenue Account" '63
    .TextMatrix(.Row, 3) = 1
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "C. Source of fund to meet revenue expenditure" '64
    .TextMatrix(.Row, 3) = 1
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "i. Income earned during current year" '65
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "ii. Amount applied out of Accumulation u/s 11(2)" '66
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "iii. Amount applied out of Option under Explanation (2) to 11(1)" '67
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "iv. Amount applied out of Accumulated Surplus (15%)" '68
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "v. Amount applied out of Corpus Fund" '69
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "vi. Amount applied out of Borrowed Funds" '70
    .TextMatrix(.Row, 3) = 3
    RowInc
    .TextMatrix(.Row, 0) = Space(5) + "Total Source of Funds for Application on Revenue Account" '71
    .TextMatrix(.Row, 3) = 4
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "D. Net Application of Funds on Revenue Account" '72
    .TextMatrix(.Row, 3) = 1
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "E. Amount which was not actually paid during the previous year" '73
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "F. Amount paid out of accruals during earlier years" '74
    .TextMatrix(.Row, 3) = 3
    RowInc
    .TextMatrix(.Row, 0) = Space(5) + "G. Allowable Application of Funds on Revenue Account" '75
    .TextMatrix(.Row, 3) = 1
    RowInc
    RowInc
    .TextMatrix(.Row, 0) = Space(5) + "2. Schedule EC" '77
    .TextMatrix(.Row, 3) = 1
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "A. Capital Expenditure" '78
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "i. Addition to capital work-in-progress" '79
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "ii. Acquisition of capital asset" '80
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "iii. Cost of new asset for claim of exemption u/s 11(1A)" '81
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "iv. Other capital expenses" '82
    .TextMatrix(.Row, 3) = 3
    RowInc
    .TextMatrix(.Row, 0) = Space(5) + "Total Application of Funds on Capital Account" '83
    .TextMatrix(.Row, 3) = 1
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "B. Source of fund to meet capital expenditure" '84
    .TextMatrix(.Row, 3) = 1
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "i. Income earned during current year" '85
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "ii. Amount applied out of Accumulation u/s 11(2)" '86
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "iii. Amount applied out of Option under Explanation (2) to 11(1)" '87
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "iv. Amount applied out of Accumulated Surplus (15%)" '88
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "v. Amount applied out of Corpus Fund" '89
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(15) + "vi. Amount applied out of Borrowed Funds" '90
    .TextMatrix(.Row, 3) = 3
    RowInc
    .TextMatrix(.Row, 0) = Space(5) + "Total Source of Funds for Application on Capital Account" '91
    .TextMatrix(.Row, 3) = 4
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "C. Net Application of Funds on Capital Account" '92
    .TextMatrix(.Row, 3) = 1
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "D. Amount which was not actually paid during the previous year" '93
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "E. Amount paid out of accruals during earlier years" '94
    .TextMatrix(.Row, 3) = 3
    RowInc
    .TextMatrix(.Row, 0) = Space(5) + "F. Allowable Application of Funds on Capital Account" '95
    .TextMatrix(.Row, 3) = 1
    RowInc
    RowInc
    .TextMatrix(.Row, 0) = Space(5) + "2. Amount Re-qualifying as Application of Income" '97
    .TextMatrix(.Row, 3) = 1
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "A. Repayment of Borrowed Funds" '98
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "B. Repatriation to Corpus Fund" '99
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(5) + "Total Amount re-qualifying as Application of Income" '100
    .TextMatrix(.Row, 3) = 1
    RowInc
    RowInc
    .TextMatrix(.Row, 0) = "Total Application of Income for Charitable Purposes (1 + 2 + 3)" '102
    .TextMatrix(.Row, 3) = 5
    RowInc
    RowInc
    .TextMatrix(.Row, 0) = "C. Calculation of Taxable Income" '104
    .TextMatrix(.Row, 3) = 1
    RowInc
    .TextMatrix(.Row, 0) = Space(5) + "1. Surplus/Deficit during the year (A - B)" '105
    .TextMatrix(.Row, 3) = 1
    RowInc
    .TextMatrix(.Row, 0) = Space(5) + "2. Amount set apart out of Income (upto 15%)" '106
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "3. Accumulation u/s 11(2) for 5 years (Form 10 to be Filed)" '107
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = Space(10) + "4. Option under Explanation (2) to 11(1) (Form 9A to be Filed)" '108
    .TextMatrix(.Row, 3) = 2
    RowInc
    .TextMatrix(.Row, 0) = "Amount deemed to be applied to Charitable Purposes" '109
    .TextMatrix(.Row, 3) = 1
    .Cell(flexcpBackColor, 3, 0, 5, .Cols - 1) = RGB(128, 255, 128)
    .Cell(flexcpBackColor, 28, 0, 28, .Cols - 1) = RGB(128, 255, 128)
    .Cell(flexcpBackColor, 53, 0, 53, .Cols - 1) = RGB(128, 255, 128)
    .Cell(flexcpBackColor, 58, 0, 58, .Cols - 1) = RGB(128, 255, 128)
    .Cell(flexcpBackColor, 60, 0, 61, .Cols - 1) = RGB(128, 255, 128)
    .Cell(flexcpBackColor, 66, 0, 66, .Cols - 1) = RGB(128, 255, 128)
    .Cell(flexcpBackColor, 69, 0, 70, .Cols - 1) = RGB(128, 255, 128)
    .Cell(flexcpBackColor, 73, 0, 74, .Cols - 1) = RGB(128, 255, 128)
    .Cell(flexcpBackColor, 85, 0, 90, .Cols - 1) = RGB(128, 255, 128)
    .Cell(flexcpBackColor, 93, 0, 94, .Cols - 1) = RGB(128, 255, 128)
    .Cell(flexcpBackColor, 98, 0, 99, .Cols - 1) = RGB(128, 255, 128)
    .Cell(flexcpBackColor, 107, 0, 108, .Cols - 1) = RGB(128, 255, 128)
    .Refresh
End If
End With
End Sub

Private Sub VsfHelp_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim TaxInc As Double
If CForm = 0 Then
    Set RsQ = Nothing
        If Row = 29 Or Row = 55 Then
            If RsOpDtl.BOF = False Then
                RsOpDtl.MoveFirst
                RsOpDtl.Find "LCode=-2"
            End If
            If RsOpDtl.EOF = False Then
                If RsOpDtl.Fields("OpBal") < Val(VsfHelp.TextMatrix(Row, 3)) Then
                    MsgBox "Value cannot exceed Opening Balance.", vbCritical, "Alert"
                    VsfHelp.TextMatrix(Row, 3) = ""
                    VsfHelp.Row = Row
                    VsfHelp.Col = Col
                    VsfHelp.SetFocus
                    Exit Sub
                End If
            Else
                If Val(VsfHelp.TextMatrix(Row, 3)) <> 0 Then
                    MsgBox "Value cannot exceed Opening Balance.", vbCritical, "Alert"
                    VsfHelp.TextMatrix(Row, 3) = ""
                    VsfHelp.Row = Row
                    VsfHelp.Col = Col
                    VsfHelp.SetFocus
                    Exit Sub
                End If
            End If
        ElseIf Row = 30 Then
            If RsOpDtl.BOF = False Then
                RsOpDtl.MoveFirst
                RsOpDtl.Find "LCode=-1"
            End If
            If RsOpDtl.EOF = False Then
                If RsOpDtl.Fields("OpBal") < Val(VsfHelp.TextMatrix(Row, 3)) Then
                    MsgBox "Value cannot exceed Opening Balance.", vbCritical, "Alert"
                    VsfHelp.TextMatrix(Row, 3) = ""
                    VsfHelp.Row = Row
                    VsfHelp.Col = Col
                    VsfHelp.SetFocus
                    Exit Sub
                End If
            Else
                If Val(VsfHelp.TextMatrix(Row, 3)) <> 0 Then
                    MsgBox "Value cannot exceed Opening Balance.", vbCritical, "Alert"
                    VsfHelp.TextMatrix(Row, 3) = ""
                    VsfHelp.Row = Row
                    VsfHelp.Col = Col
                    VsfHelp.SetFocus
                    Exit Sub
                End If
            End If
        ElseIf Row = 46 Then
            If RsOpDtl.BOF = False Then
                RsOpDtl.MoveFirst
                RsOpDtl.Find "LCode=-3"
            End If
            If RsOpDtl.EOF = False Then
                If RsOpDtl.Fields("OpBal") < Val(VsfHelp.TextMatrix(Row, 3)) Then
                    MsgBox "Value cannot exceed Opening Balance.", vbCritical, "Alert"
                    VsfHelp.TextMatrix(Row, 3) = ""
                    VsfHelp.Row = Row
                    VsfHelp.Col = Col
                    VsfHelp.SetFocus
                    Exit Sub
                End If
            Else
                If Val(VsfHelp.TextMatrix(Row, 3)) <> 0 Then
                    MsgBox "Value cannot exceed Opening Balance.", vbCritical, "Alert"
                    VsfHelp.TextMatrix(Row, 3) = ""
                    VsfHelp.Row = Row
                    VsfHelp.Col = Col
                    VsfHelp.SetFocus
                    Exit Sub
                End If
            End If
        ElseIf Row = 47 Then
            If RsOpDtl.BOF = False Then
                RsOpDtl.MoveFirst
                RsOpDtl.Find "LCode=-4"
            End If
            If RsOpDtl.EOF = False Then
                If RsOpDtl.Fields("OpBal") < Val(VsfHelp.TextMatrix(Row, 3)) Then
                    MsgBox "Value cannot exceed Opening Balance.", vbCritical, "Alert"
                    VsfHelp.TextMatrix(Row, 3) = ""
                    VsfHelp.Row = Row
                    VsfHelp.Col = Col
                    VsfHelp.SetFocus
                    Exit Sub
                End If
            Else
                If Val(VsfHelp.TextMatrix(Row, 3)) <> 0 Then
                    MsgBox "Value cannot exceed Opening Balance.", vbCritical, "Alert"
                    VsfHelp.TextMatrix(Row, 3) = ""
                    VsfHelp.Row = Row
                    VsfHelp.Col = Col
                    VsfHelp.SetFocus
                    Exit Sub
                End If
            End If
        ElseIf Row = 62 Then
                If Val(VsfHelp.TextMatrix(Row, 3)) <> 0 Then
                    If Val(VsfHelp.TextMatrix(57, 4)) < 0 Then
                        TxtCTotal.Text = IIf((Val(VsfHelp.TextMatrix(60, 4)) - Val(VsfHelp.TextMatrix(61, 4)) - Val(VsfHelp.TextMatrix(62, 3)) - Val(VsfHelp.TextMatrix(63, 3))) > 0, Round((Val(VsfHelp.TextMatrix(60, 4)) - Val(VsfHelp.TextMatrix(61, 4)) - Val(VsfHelp.TextMatrix(62, 3)) - Val(VsfHelp.TextMatrix(63, 3))), 0), "0.00")
                        If Val(VsfHelp.TextMatrix(Row, 3)) > Val(TxtCTotal.Text) Then
                            MsgBox "Value cannot exceed Net Taxable Income, excluding Unutilised Option.", vbCritical, "Alert"
                            VsfHelp.TextMatrix(Row, 3) = ""
                            VsfHelp.Row = Row
                            VsfHelp.Col = Col
                            VsfHelp.SetFocus
                            SetFinalTot
                            Exit Sub
                        End If
                    Else
                        TxtCTotal.Text = IIf((Val(VsfHelp.TextMatrix(60, 4)) - Val(VsfHelp.TextMatrix(61, 4)) - Val(VsfHelp.TextMatrix(64, 4))) > 0, Round((Val(VsfHelp.TextMatrix(60, 4)) - Val(VsfHelp.TextMatrix(61, 4)) - Val(VsfHelp.TextMatrix(64, 4))), 0), "0.00")
                        If Val(VsfHelp.TextMatrix(Row, 3)) > Val(TxtCTotal.Text) Then
                            MsgBox "Value cannot exceed Net Taxable Income.", vbCritical, "Alert"
                            VsfHelp.TextMatrix(Row, 3) = ""
                            VsfHelp.Row = Row
                            VsfHelp.Col = Col
                            VsfHelp.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
        ElseIf Row = 63 Then
                If Val(VsfHelp.TextMatrix(Row, 3)) <> 0 Then
                    If Val(VsfHelp.TextMatrix(57, 4)) < 0 Then
                        TxtCTotal.Text = IIf((Val(VsfHelp.TextMatrix(60, 4)) - Val(VsfHelp.TextMatrix(61, 4)) - Val(VsfHelp.TextMatrix(62, 3)) - Val(VsfHelp.TextMatrix(63, 3))) > 0, Round((Val(VsfHelp.TextMatrix(60, 4)) - Val(VsfHelp.TextMatrix(61, 4)) - Val(VsfHelp.TextMatrix(62, 3)) - Val(VsfHelp.TextMatrix(63, 3))), 0), "0.00")
                        If Val(TxtCTotal.Text) < Abs(Val(VsfHelp.TextMatrix(57, 4))) Then
                            MsgBox "Net Taxable Income cannot be less than Unutilised Option.", vbCritical, "Alert"
                            VsfHelp.TextMatrix(Row, 3) = ""
                            VsfHelp.Row = Row
                            VsfHelp.Col = Col
                            VsfHelp.SetFocus
                            SetFinalTot
                            Exit Sub
                        End If
                    Else
                        If Val(VsfHelp.TextMatrix(Row, 3)) > Val(TxtCTotal.Text) Then
                            MsgBox "Value cannot exceed Net Taxable Income.", vbCritical, "Alert"
                            VsfHelp.TextMatrix(Row, 3) = ""
                            VsfHelp.Row = Row
                            VsfHelp.Col = Col
                            VsfHelp.SetFocus
                            Exit Sub
                        End If
                    End If
                End If
        End If
Else
    Dim X As Double
    Set RsQ = Nothing
    If Col = 1 Then
        If Row = 3 Or Row = 4 Or Row = 5 Or Row = 28 Or Row = 53 Or Row = 58 Or Row = 60 Or Row = 61 Or Row = 66 Or Row = 69 Or Row = 70 Or Row = 73 Or Row = 74 Or Row = 85 Or Row = 86 Or Row = 87 Or Row = 88 Or Row = 89 Or Row = 90 Or Row = 93 Or Row = 94 Or Row = 98 Or Row = 99 Or Row = 107 Or Row = 108 Then
            If Val(VsfHelp.TextMatrix(Row, 1)) < 0 Then
                MsgBox "Value cannot be Negative.", vbCritical, "Alert"
                VsfHelp.TextMatrix(Row, 1) = ""
                VsfHelp.Row = Row
                VsfHelp.Col = Col
                VsfHelp.SetFocus
                SetFinalTot
                Exit Sub
            End If
        End If
    End If
    If Row = 3 Then 'Rent as per 26AS
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=42", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
        If Val(VsfHelp.TextMatrix(Row, 1)) < Val(RsQ.Fields("RTotal")) Then
                MsgBox "Value cannot be less than rent income as per audited financial statements.", vbCritical, "Alert"
                VsfHelp.TextMatrix(Row, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
                VsfHelp.Row = Row
                VsfHelp.Col = Col
                VsfHelp.SetFocus
                SetFinalTot
                Exit Sub
        End If
    ElseIf Row = 4 Then 'Dividend as per 26AS
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=48", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
        If Val(VsfHelp.TextMatrix(Row, 1)) < Val(RsQ.Fields("RTotal")) Then
                MsgBox "Value cannot be less than dividend income as per audited financial statements.", vbCritical, "Alert"
                VsfHelp.TextMatrix(Row, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
                VsfHelp.Row = Row
                VsfHelp.Col = Col
                VsfHelp.SetFocus
                SetFinalTot
                Exit Sub
        End If
    ElseIf Row = 5 Then 'Interest as per 26AS
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode In (44,45,46,47)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
        If Val(VsfHelp.TextMatrix(Row, 1)) < Val(RsQ.Fields("RTotal")) Then
                MsgBox "Value cannot be less than interest income as per audited financial statements.", vbCritical, "Alert"
                VsfHelp.TextMatrix(Row, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
                VsfHelp.Row = Row
                VsfHelp.Col = Col
                VsfHelp.SetFocus
                SetFinalTot
                Exit Sub
        End If
    ElseIf Row = 28 Or Row = 66 Or Row = 86 Then 'Accumulation u/s 11(2)
        If RsOpDtl.BOF = False Then
            RsOpDtl.MoveFirst
            RsOpDtl.Find "LCode=-2"
        End If
        If RsOpDtl.EOF = False Then
            If RsOpDtl.Fields("OpBal") < (Val(VsfHelp.TextMatrix(28, 1)) + Val(VsfHelp.TextMatrix(66, 1)) + Val(VsfHelp.TextMatrix(86, 1))) Then
                MsgBox "Value cannot exceed Opening Balance.", vbCritical, "Alert"
                VsfHelp.TextMatrix(Row, 1) = ""
                VsfHelp.Row = Row
                VsfHelp.Col = Col
                VsfHelp.SetFocus
                Exit Sub
            Else
                If Row = 66 Then
                VsfHelp.TextMatrix(65, 1) = Val(VsfHelp.TextMatrix(65, 1)) - Val(VsfHelp.TextMatrix(68, 1))
                ElseIf Row = 86 Then
                VsfHelp.TextMatrix(85, 1) = Val(VsfHelp.TextMatrix(85, 1)) - Val(VsfHelp.TextMatrix(86, 1))
                End If
            End If
        Else
            If Val(VsfHelp.TextMatrix(Row, 1)) <> 0 Then
                MsgBox "Value cannot exceed Opening Balance.", vbCritical, "Alert"
                VsfHelp.TextMatrix(Row, 1) = ""
                VsfHelp.Row = Row
                VsfHelp.Col = Col
                VsfHelp.SetFocus
                Exit Sub
            End If
        End If
    ElseIf Row = 53 Then
        If Val(VsfHelp.TextMatrix(Row, 1)) > TotDon Then
            MsgBox "Value cannot exceed Total Donations Given.", vbCritical, "Alert"
            VsfHelp.TextMatrix(Row, 1) = ""
            VsfHelp.TextMatrix(54, 1) = TotDon
            VsfHelp.Row = Row
            VsfHelp.Col = Col
            VsfHelp.SetFocus
            Exit Sub
        Else
            VsfHelp.TextMatrix(54, 1) = TotDon - Val(VsfHelp.TextMatrix(53, 1))
        End If
    ElseIf Row = 58 Or Row = 60 Or Row = 61 Then 'Expenditure on other objects
        If Val(VsfHelp.TextMatrix(Row, 1)) <> 0 Then
            If (Val(VsfHelp.TextMatrix(58, 1)) + Val(VsfHelp.TextMatrix(60, 1)) + Val(VsfHelp.TextMatrix(61, 1)) + Val(VsfHelp.TextMatrix(62, 1))) > OthExp Then
                MsgBox "Value cannot exceed total expenditure on Other Charitable Objects.", vbCritical, "Alert"
                VsfHelp.TextMatrix(Row, 1) = ""
                VsfHelp.Row = Row
                VsfHelp.Col = Col
                VsfHelp.SetFocus
                Exit Sub
            ElseIf Row = 58 Then VsfHelp.TextMatrix(62, 1) = Val(VsfHelp.TextMatrix(62, 1)) - Val(VsfHelp.TextMatrix(58, 1))
            ElseIf Row = 60 Then VsfHelp.TextMatrix(62, 1) = Val(VsfHelp.TextMatrix(62, 1)) - Val(VsfHelp.TextMatrix(60, 1))
            ElseIf Row = 61 Then VsfHelp.TextMatrix(62, 1) = Val(VsfHelp.TextMatrix(62, 1)) - Val(VsfHelp.TextMatrix(61, 1))
            End If
        End If
    ElseIf Row = 69 Or Row = 70 Or Row = 73 Then 'Unpaid Amount Not Allowed as Application - Revenue
        If Val(VsfHelp.TextMatrix(Row, 1)) > Val(VsfHelp.TextMatrix(63, 1)) Then
            MsgBox "Value cannot exceed Total Revenue Expenditure.", vbCritical, "Alert"
            VsfHelp.TextMatrix(Row, 1) = ""
            VsfHelp.Row = Row
            VsfHelp.Col = Col
            VsfHelp.SetFocus
            Exit Sub
        End If
    ElseIf Row = 74 Or Row = 94 Then 'Unpaid Amount Allowed as Application
        If RsOpDtl.BOF = False Then
            RsOpDtl.MoveFirst
            RsOpDtl.Find "LCode=-10"
        End If
        If RsOpDtl.EOF = False Then
            If (RsOpDtl.Fields("OpBal") + Val(VsfHelp.TextMatrix(73, 1)) + Val(VsfHelp.TextMatrix(93, 1))) < (Val(VsfHelp.TextMatrix(74, 1)) + Val(VsfHelp.TextMatrix(94, 1))) Then
                MsgBox "Value cannot exceed Opening Balance.", vbCritical, "Alert"
                VsfHelp.TextMatrix(Row, 1) = ""
                VsfHelp.Row = Row
                VsfHelp.Col = Col
                VsfHelp.SetFocus
                Exit Sub
            End If
        Else
            If Val(VsfHelp.TextMatrix(Row, 1)) <> 0 Then
                MsgBox "Value cannot exceed Opening Balance.", vbCritical, "Alert"
                VsfHelp.TextMatrix(Row, 1) = ""
                VsfHelp.Row = Row
                VsfHelp.Col = Col
                VsfHelp.SetFocus
                Exit Sub
            End If
        End If
    ElseIf Row = 87 Then 'Option under explanation (2) to 11(1)
        If RsOpDtl.BOF = False Then
            RsOpDtl.MoveFirst
            RsOpDtl.Find "LCode=-1"
        End If
        If RsOpDtl.EOF = False Then
            X = Val(VsfHelp.TextMatrix(67, 1))
            VsfHelp.TextMatrix(67, 1) = X - Val(VsfHelp.TextMatrix(87, 1))
            If RsOpDtl.Fields("OpBal") < Val(VsfHelp.TextMatrix(67, 1)) + (Val(VsfHelp.TextMatrix(87, 1))) Then
                MsgBox "Value cannot exceed Opening Balance.", vbCritical, "Alert"
                VsfHelp.TextMatrix(67, 1) = X
                VsfHelp.TextMatrix(Row, 1) = ""
                VsfHelp.Row = Row
                VsfHelp.Col = Col
                VsfHelp.SetFocus
                Exit Sub
            Else
                Set RsQ = Nothing 'Source of Fund - Revenue - PY Option
                RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=67", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
                If RsQ.Fields("RTotal") = 0 Then
                    Set RsQ = Nothing
                    RsQ.Open "Select LCode, IIF(IsNull(OpCr)=True,0,OpCr) As RTotal From TmpTrialBal Where LCode=-1", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
                    If RsQ.BOF = False Then
                        RsQ.MoveFirst
                        RsQ.Find "LCode = -1"
                    End If
                    If RsQ.EOF = False Then
                        X = IIf((RsQ.Fields("RTotal") - Val(VsfHelp.TextMatrix(87, 1))) < Val(VsfHelp.TextMatrix(63, 1)), (RsQ.Fields("RTotal") - Val(VsfHelp.TextMatrix(87, 1))), Val(VsfHelp.TextMatrix(63, 1)))
                        VsfHelp.TextMatrix(67, 1) = CStr(Round(X, 2))
                        VsfHelp.TextMatrix(85, 1) = Val(VsfHelp.TextMatrix(85, 1)) - Val(VsfHelp.TextMatrix(87, 1))
                    End If
                Else
                VsfHelp.TextMatrix(67, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
                End If
            End If
        Else
            If Val(VsfHelp.TextMatrix(Row, 1)) <> 0 Then
                MsgBox "Value cannot exceed Opening Balance.", vbCritical, "Alert"
                VsfHelp.TextMatrix(Row, 1) = ""
                VsfHelp.Row = Row
                VsfHelp.Col = Col
                VsfHelp.SetFocus
                Exit Sub
            End If
        End If
    ElseIf Row = 88 Then 'Application out of 15% accumulation
        If RsOpDtl.BOF = False Then
            RsOpDtl.MoveFirst
            RsOpDtl.Find "HCode=5"
        End If
        If RsOpDtl.EOF = False Then
            Set RsQ = Nothing
            RsQ.Open "Select Sum(IIf(OpCr>0,OpCr,0)) as OpBal from TmpTrialBal where HCode In (2,5)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
            X = RsQ.Fields("OpBal")
            If X > 0 And X < (Val(VsfHelp.TextMatrix(68, 1)) - Val(VsfHelp.TextMatrix(88, 1)) + Val(VsfHelp.TextMatrix(88, 1))) Then
                MsgBox "Value cannot exceed Opening Balance.", vbCritical, "Alert"
                VsfHelp.TextMatrix(Row, 1) = ""
                VsfHelp.Row = Row
                VsfHelp.Col = Col
                VsfHelp.SetFocus
                Exit Sub
            Else
                Set RsQ = Nothing 'Source of Fund - Revenue - 15% Surplus
                RsQ.Open "Select IIF(IsNull(Sum(IIf(OpCr>0,OpCr,0)))=True,0,Sum(IIf(OpCr>0,OpCr,0))) As RTotal From TmpTrialBal Where HCode In (2,5)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
                X = IIf((RsQ.Fields("RTotal") - Val(VsfHelp.TextMatrix(88, 1))) < (Val(VsfHelp.TextMatrix(63, 1)) - Val(VsfHelp.TextMatrix(65, 1)) - Val(VsfHelp.TextMatrix(66, 1)) - Val(VsfHelp.TextMatrix(67, 1)) - Val(VsfHelp.TextMatrix(69, 1)) - Val(VsfHelp.TextMatrix(70, 1))), (RsQ.Fields("RTotal") - Val(VsfHelp.TextMatrix(88, 1))), (Val(VsfHelp.TextMatrix(63, 1)) - Val(VsfHelp.TextMatrix(65, 1)) - Val(VsfHelp.TextMatrix(66, 1)) - Val(VsfHelp.TextMatrix(67, 1)) - Val(VsfHelp.TextMatrix(69, 1)) - Val(VsfHelp.TextMatrix(70, 1))))
                VsfHelp.TextMatrix(68, 1) = CStr(Round(X, 2))
                VsfHelp.TextMatrix(85, 1) = Val(VsfHelp.TextMatrix(85, 1)) - Val(VsfHelp.TextMatrix(88, 1))
                Exit Sub
            End If
        Else
            If Val(VsfHelp.TextMatrix(Row, 1)) <> 0 Then
                MsgBox "Value cannot exceed Opening Balance.", vbCritical, "Alert"
                VsfHelp.TextMatrix(Row, 1) = ""
                VsfHelp.Row = Row
                VsfHelp.Col = Col
                VsfHelp.SetFocus
                Exit Sub
            End If
        End If
    ElseIf Row = 89 Or Row = 90 Or Row = 93 Then 'Unpaid Amount Not Allowed as Application - Capital
        If Val(VsfHelp.TextMatrix(Row, 1)) > Val(VsfHelp.TextMatrix(83, 1)) Then
                MsgBox "Value cannot exceed Total Capital Expenditure.", vbCritical, "Alert"
                VsfHelp.TextMatrix(Row, 1) = ""
                VsfHelp.Row = Row
                VsfHelp.Col = Col
                VsfHelp.SetFocus
                Exit Sub
        End If
    ElseIf Row = 98 Then 'Repayment of Borrowed Funds
        If RsOpDtl.BOF = False Then
            RsOpDtl.MoveFirst
            RsOpDtl.Find "LCode=-3"
        End If
        If RsOpDtl.EOF = False Then
            If (RsOpDtl.Fields("OpBal") + Val(VsfHelp.TextMatrix(70, 1)) + Val(VsfHelp.TextMatrix(90, 1))) < Val(VsfHelp.TextMatrix(Row, 1)) Then
                MsgBox "Value cannot exceed Total Application out of Borrowed Funds.", vbCritical, "Alert"
                VsfHelp.TextMatrix(Row, 1) = ""
                VsfHelp.Row = Row
                VsfHelp.Col = Col
                VsfHelp.SetFocus
                Exit Sub
            End If
        Else
            If Val(VsfHelp.TextMatrix(Row, 1)) <> 0 Then
                MsgBox "Value cannot exceed Opening Balance.", vbCritical, "Alert"
                VsfHelp.TextMatrix(Row, 1) = ""
                VsfHelp.Row = Row
                VsfHelp.Col = Col
                VsfHelp.SetFocus
                Exit Sub
            End If
        End If
    ElseIf Row = 99 Then 'Repartriation to Corpus Fund
        If RsOpDtl.BOF = False Then
            RsOpDtl.MoveFirst
            RsOpDtl.Find "LCode=-4"
        End If
        If RsOpDtl.EOF = False Then
            If (RsOpDtl.Fields("OpBal") + Val(VsfHelp.TextMatrix(69, 1)) + Val(VsfHelp.TextMatrix(89, 1))) < Val(VsfHelp.TextMatrix(Row, 1)) Then
                MsgBox "Value cannot exceed Total Application out of Corpus Fund.", vbCritical, "Alert"
                VsfHelp.TextMatrix(Row, 1) = ""
                VsfHelp.Row = Row
                VsfHelp.Col = Col
                VsfHelp.SetFocus
                Exit Sub
            End If
        Else
            If Val(VsfHelp.TextMatrix(Row, 1)) <> 0 Then
                MsgBox "Value cannot exceed Opening Balance.", vbCritical, "Alert"
                VsfHelp.TextMatrix(Row, 1) = ""
                VsfHelp.Row = Row
                VsfHelp.Col = Col
                VsfHelp.SetFocus
                Exit Sub
            End If
        End If
    ElseIf Row = 107 Or Row = 108 Then 'Fresh option under Explanation (2) to section 11(1) or section 11(2)
            If Val(VsfHelp.TextMatrix(Row, 1)) <> 0 Then
                TxtCTotal.Text = IIf((Val(VsfHelp.TextMatrix(9, 2)) + Val(VsfHelp.TextMatrix(25, 2)) - Val(VsfHelp.TextMatrix(75, 2)) - Val(VsfHelp.TextMatrix(95, 2)) - Val(VsfHelp.TextMatrix(100, 2))) > 0, Round((Val(VsfHelp.TextMatrix(9, 2)) + Val(VsfHelp.TextMatrix(25, 2)) - Val(VsfHelp.TextMatrix(75, 2)) - Val(VsfHelp.TextMatrix(95, 2)) - Val(VsfHelp.TextMatrix(100, 2))), 0), "0.00")
                If Row = 107 Then
                    If Val(VsfHelp.TextMatrix(Row, 1)) > (Val(TxtCTotal.Text) - Val(VsfHelp.TextMatrix(108, 1))) Then
                        MsgBox "Value cannot exceed unutilised income of current year.", vbCritical, "Alert"
                        VsfHelp.TextMatrix(Row, 1) = ""
                        VsfHelp.Row = Row
                        VsfHelp.Col = Col
                        VsfHelp.SetFocus
                        SetFinalTot
                        Exit Sub
                    End If
                Else
                    If Val(VsfHelp.TextMatrix(Row, 1)) > (Val(TxtCTotal.Text) - Val(VsfHelp.TextMatrix(107, 1))) Then
                        MsgBox "Value cannot exceed unutilised income of current year.", vbCritical, "Alert"
                        VsfHelp.TextMatrix(Row, 1) = ""
                        VsfHelp.Row = Row
                        VsfHelp.Col = Col
                        VsfHelp.SetFocus
                        SetFinalTot
                        Exit Sub
                    End If
                End If
            End If
    End If
End If
SetFinalTot
End Sub

Private Sub VsfHelp_EnterCell()
If VsfHelp.Col = 1 Then
    If VsfHelp.Row = 3 Or VsfHelp.Row = 4 Or VsfHelp.Row = 5 Or VsfHelp.Row = 28 Or VsfHelp.Row = 53 Or VsfHelp.Row = 58 Or VsfHelp.Row = 60 Or VsfHelp.Row = 61 Or VsfHelp.Row = 66 Or VsfHelp.Row = 69 Or VsfHelp.Row = 70 Or VsfHelp.Row = 73 Or VsfHelp.Row = 74 Or VsfHelp.Row = 85 Or VsfHelp.Row = 86 Or VsfHelp.Row = 87 Or VsfHelp.Row = 88 Or VsfHelp.Row = 89 Or VsfHelp.Row = 90 Or VsfHelp.Row = 93 Or VsfHelp.Row = 94 Or VsfHelp.Row = 98 Or VsfHelp.Row = 99 Or VsfHelp.Row = 107 Or VsfHelp.Row = 108 Then
        VsfHelp.Editable = flexEDKbd
        VsfHelp.AutoSearch = flexSearchNone
        On Error Resume Next
        SendKeys "{F2}"
    Else
        VsfHelp.Editable = flexEDNone
        VsfHelp.AutoSearch = flexSearchFromCursor
    End If
Else
    VsfHelp.Editable = flexEDNone
    VsfHelp.AutoSearch = flexSearchFromCursor
End If
End Sub

Private Sub SetFinalTot()
Dim i As Double
Dim RsQAcc As New ADODB.Recordset
With VsfHelp
If CForm = 0 Then
    .Row = 8
    .TextMatrix(.Row, 3) = Val(.TextMatrix(7, 2)) + Val(.TextMatrix(8, 2))
    .Row = 11
    .TextMatrix(.Row, 3) = Val(.TextMatrix(10, 2)) + Val(.TextMatrix(11, 2))
    .Row = 17
    .TextMatrix(.Row, 3) = Val(.TextMatrix(13, 2)) + Val(.TextMatrix(14, 2)) + Val(.TextMatrix(16, 2)) + Val(.TextMatrix(17, 2))
    .Row = 21
    .TextMatrix(.Row, 3) = Val(.TextMatrix(3, 3)) + Val(.TextMatrix(4, 3)) + Val(.TextMatrix(5, 3)) + Val(.TextMatrix(8, 3)) + Val(.TextMatrix(11, 3)) + Val(.TextMatrix(17, 3)) + Val(.TextMatrix(18, 3)) + Val(.TextMatrix(19, 3)) + Val(.TextMatrix(20, 3))
    .Row = 25
    .TextMatrix(.Row, 3) = Val(.TextMatrix(24, 2)) + Val(.TextMatrix(25, 2))
    .TextMatrix(26, 3) = Val(.TextMatrix(25, 3)) * -1
    .Row = 27
    .TextMatrix(.Row, 3) = Val(.TextMatrix(25, 3)) + Val(.TextMatrix(26, 3))
    .Row = 31
    .TextMatrix(.Row, 3) = Val(.TextMatrix(29, 3)) + Val(.TextMatrix(30, 3))
    .Row = 32
    .TextMatrix(.Row, 4) = Val(.TextMatrix(21, 3)) + Val(.TextMatrix(27, 3)) + Val(.TextMatrix(31, 3))
    .Row = 39
    .TextMatrix(.Row, 3) = Val(.TextMatrix(36, 3)) + Val(.TextMatrix(37, 3)) + Val(.TextMatrix(38, 3))
    .Row = 43
    If (Val(.TextMatrix(42, 3)) * -1) - Val(.TextMatrix(41, 3)) > 0 Then .TextMatrix(.Row, 3) = ((Val(.TextMatrix(42, 3)) * -1) - Val(.TextMatrix(41, 3))) * -1 Else .TextMatrix(.Row, 3) = "0.00"
    .TextMatrix(19, 3) = Val(.TextMatrix(43, 3)) * -1
    .Row = 44
    .TextMatrix(.Row, 3) = Val(.TextMatrix(41, 3)) + Val(.TextMatrix(42, 3)) + Val(.TextMatrix(43, 3))
    .Row = 48
    .TextMatrix(.Row, 3) = (Val(.TextMatrix(46, 3)) + Val(.TextMatrix(47, 3)))
    .Row = 52
    .TextMatrix(.Row, 3) = (Val(.TextMatrix(50, 3)) + Val(.TextMatrix(51, 3)))
    .Row = 56
    .TextMatrix(.Row, 3) = (Val(.TextMatrix(54, 3)) + Val(.TextMatrix(55, 3)))
    .Row = 57
    .TextMatrix(.Row, 4) = (Val(.TextMatrix(39, 3)) + Val(.TextMatrix(44, 3))) + Val(.TextMatrix(48, 3)) - (Val(.TextMatrix(52, 3)) + Val(.TextMatrix(56, 3)))
    .Row = 60
    .TextMatrix(.Row, 4) = Val(.TextMatrix(32, 4)) - Val(.TextMatrix(57, 4))
    .Row = 61
    If Val(.TextMatrix(60, 4)) > 0 Then
        If (Val(.TextMatrix(32, 4)) * 0.15) < Val(.TextMatrix(60, 4)) Then .TextMatrix(.Row, 4) = CStr(Round((Val(.TextMatrix(32, 4)) * 0.15), 2)) Else .TextMatrix(.Row, 4) = .TextMatrix(60, 4)
    End If
    .Row = 64
    .TextMatrix(.Row, 4) = Val(.TextMatrix(62, 3)) + Val(.TextMatrix(63, 3))
    TxtCTotal.Text = IIf((Val(.TextMatrix(60, 4)) - Val(.TextMatrix(61, 4)) - Val(.TextMatrix(64, 4))) > 0, Round((Val(.TextMatrix(60, 4)) - Val(.TextMatrix(61, 4)) - Val(.TextMatrix(64, 4))), 0), "0.00")
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 2)) = 0 Then .TextMatrix(i, 2) = ""
        If Val(.TextMatrix(i, 3)) = 0 Then .TextMatrix(i, 3) = ""
        If Val(.TextMatrix(i, 4)) = 0 Then .TextMatrix(i, 4) = ""
    Next
    .Refresh
Else
    Dim X As Double
    .TextMatrix(9, 2) = Val(.TextMatrix(3, 1)) + Val(.TextMatrix(4, 1)) + Val(.TextMatrix(5, 1)) + Val(.TextMatrix(6, 1)) + Val(.TextMatrix(7, 1)) + Val(.TextMatrix(8, 1))
    .TextMatrix(24, 1) = (Val(.TextMatrix(21, 1)) + Val(.TextMatrix(23, 1))) * -1
    .TextMatrix(25, 2) = Val(.TextMatrix(13, 1)) + Val(.TextMatrix(14, 1)) + Val(.TextMatrix(15, 1)) + Val(.TextMatrix(16, 1)) + Val(.TextMatrix(18, 1)) + Val(.TextMatrix(19, 1)) + Val(.TextMatrix(21, 1)) + Val(.TextMatrix(23, 1)) + Val(.TextMatrix(24, 1))
    .TextMatrix(30, 2) = Val(.TextMatrix(28, 1)) + Val(.TextMatrix(29, 1))
    .TextMatrix(32, 2) = Val(.TextMatrix(9, 2)) + Val(.TextMatrix(25, 2)) + Val(.TextMatrix(30, 2))
    .TextMatrix(63, 1) = Val(.TextMatrix(37, 1)) + Val(.TextMatrix(38, 1)) + Val(.TextMatrix(39, 1)) + Val(.TextMatrix(40, 1)) + Val(.TextMatrix(41, 1)) + Val(.TextMatrix(42, 1)) + Val(.TextMatrix(43, 1)) + Val(.TextMatrix(44, 1)) + Val(.TextMatrix(45, 1)) + Val(.TextMatrix(46, 1)) + Val(.TextMatrix(47, 1)) + Val(.TextMatrix(48, 1)) + Val(.TextMatrix(49, 1)) + Val(.TextMatrix(50, 1)) + Val(.TextMatrix(54, 1)) + Val(.TextMatrix(55, 1)) + Val(.TextMatrix(56, 1)) + Val(.TextMatrix(57, 1)) + Val(.TextMatrix(58, 1)) + Val(.TextMatrix(59, 1)) + Val(.TextMatrix(60, 1)) + Val(.TextMatrix(61, 1)) + Val(.TextMatrix(62, 1))
    .TextMatrix(72, 1) = IIf((Val(.TextMatrix(65, 1)) > 0), Val(.TextMatrix(65, 1)), 0)
    .TextMatrix(75, 2) = IIf((Val(.TextMatrix(72, 1)) + Val(.TextMatrix(74, 1)) - Val(.TextMatrix(73, 1))) > 0, Val(.TextMatrix(72, 1)) + Val(.TextMatrix(74, 1)) - Val(.TextMatrix(73, 1)), 0)
    .TextMatrix(83, 1) = Val(.TextMatrix(79, 1)) + Val(.TextMatrix(80, 1)) + Val(.TextMatrix(81, 1)) + Val(.TextMatrix(82, 1))
    .TextMatrix(85, 1) = IIf((Val(.TextMatrix(83, 1)) - Val(.TextMatrix(86, 1)) - Val(.TextMatrix(87, 1)) - Val(.TextMatrix(89, 1)) - Val(.TextMatrix(90, 1))) < (Val(.TextMatrix(9, 2)) + Val(.TextMatrix(25, 2))), Val(.TextMatrix(83, 1)) - Val(.TextMatrix(86, 1)) - Val(.TextMatrix(87, 1)) - Val(.TextMatrix(89, 1)) - Val(.TextMatrix(90, 1)), Val(.TextMatrix(9, 2)) + Val(.TextMatrix(25, 2)))
    .TextMatrix(65, 1) = IIf((Val(.TextMatrix(63, 1)) - Val(.TextMatrix(66, 1)) - Val(.TextMatrix(67, 1)) - Val(.TextMatrix(69, 1)) - Val(.TextMatrix(70, 1))) < (Val(.TextMatrix(9, 2)) + Val(.TextMatrix(25, 2)) - Val(.TextMatrix(85, 1))), Val(.TextMatrix(63, 1)) - Val(.TextMatrix(66, 1)) - Val(.TextMatrix(67, 1)) - Val(.TextMatrix(69, 1)) - Val(.TextMatrix(70, 1)), Val(.TextMatrix(9, 2)) + Val(.TextMatrix(25, 2)) - Val(.TextMatrix(85, 1)))
    .TextMatrix(71, 1) = Val(.TextMatrix(65, 1)) + Val(.TextMatrix(66, 1)) + Val(.TextMatrix(67, 1)) + Val(.TextMatrix(68, 1)) + Val(.TextMatrix(69, 1)) + Val(.TextMatrix(70, 1))
    Set RsQ = Nothing 'Unutilised Option under Explanation (2) to 11(1)
    RsQ.Open "Select IIf(IsNull(Sum(OpCr))=True,0,Sum(OpCr)) as RTotal from TmpTrialBal where LCode=-1", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    X = RsQ.Fields("RTotal") - Val(.TextMatrix(67, 1)) - Val(.TextMatrix(87, 1))
    .TextMatrix(29, 1) = CStr(Round(X, 2))
    Set RsQ = Nothing 'Source of Fund - Revenue - 15% Surplus
    RsQ.Open "Select IIF(IsNull(Sum(IIf(OpCr>0,OpCr,0)))=True,0,Sum(IIf(OpCr>0,OpCr,0))) As RTotal From TmpTrialBal Where HCode In (2,5)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    X = IIf((RsQ.Fields("RTotal") - Val(.TextMatrix(88, 1))) < (Val(.TextMatrix(63, 1)) - Val(.TextMatrix(65, 1)) - Val(.TextMatrix(66, 1)) - Val(.TextMatrix(67, 1)) - Val(.TextMatrix(69, 1)) - Val(.TextMatrix(70, 1))), (RsQ.Fields("RTotal") - Val(.TextMatrix(88, 1))), (Val(.TextMatrix(63, 1)) - Val(.TextMatrix(65, 1)) - Val(.TextMatrix(66, 1)) - Val(.TextMatrix(67, 1)) - Val(.TextMatrix(69, 1)) - Val(.TextMatrix(70, 1))))
    .TextMatrix(68, 1) = CStr(Round(X, 2))
    .TextMatrix(91, 1) = Val(.TextMatrix(85, 1)) + Val(.TextMatrix(86, 1)) + Val(.TextMatrix(87, 1)) + Val(.TextMatrix(88, 1)) + Val(.TextMatrix(89, 1)) + Val(.TextMatrix(90, 1))
    .TextMatrix(92, 1) = IIf((Val(.TextMatrix(85, 1)) > 0), Val(.TextMatrix(85, 1)), 0)
    .TextMatrix(95, 2) = IIf((Val(.TextMatrix(92, 1)) + Val(.TextMatrix(94, 1)) - Val(.TextMatrix(93, 1))) > 0, Val(.TextMatrix(92, 1)) + Val(.TextMatrix(94, 1)) - Val(.TextMatrix(93, 1)), 0)
    .TextMatrix(100, 2) = Val(.TextMatrix(98, 1)) + Val(.TextMatrix(99, 1))
    .TextMatrix(102, 2) = Val(.TextMatrix(75, 2)) + Val(.TextMatrix(95, 2)) + Val(.TextMatrix(100, 2))
    .TextMatrix(105, 2) = IIf((Val(.TextMatrix(32, 2)) - Val(.TextMatrix(102, 2))) > 0, Val(.TextMatrix(32, 2)) - Val(.TextMatrix(102, 2)), 0)
    If Val(.TextMatrix(105, 2)) > 0 Then 'CY 15% Surplus
        If ((Val(.TextMatrix(9, 2)) + Val(.TextMatrix(25, 2))) * 0.15) < Val(.TextMatrix(105, 2)) Then .TextMatrix(106, 2) = CStr(Round(((Val(.TextMatrix(9, 2)) + Val(.TextMatrix(25, 2))) * 0.15), 2)) Else .TextMatrix(106, 2) = .TextMatrix(105, 2)
    Else: .TextMatrix(106, 2) = 0
    End If
    .TextMatrix(109, 2) = Val(.TextMatrix(107, 1)) + Val(.TextMatrix(108, 1))
    TxtCTotal.Text = IIf((Val(.TextMatrix(105, 2)) - Val(.TextMatrix(106, 2)) - Val(.TextMatrix(109, 2))) > 0, Round((Val(.TextMatrix(105, 2)) - Val(.TextMatrix(106, 2)) - Val(.TextMatrix(109, 2))), 0), "0.00")
        
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 1)) = 0 Then .TextMatrix(i, 1) = "" Else .TextMatrix(i, 1) = Format(Round(Val(.TextMatrix(i, 1)), 0), "0.00")
        If Val(.TextMatrix(i, 2)) = 0 Then .TextMatrix(i, 2) = "" Else .TextMatrix(i, 2) = Format(Round(Val(.TextMatrix(i, 2)), 0), "0.00")
    Next
    .Refresh
End If
End With
End Sub
Private Sub RowInc()
With VsfHelp
    .Rows = .Rows + 1
    .Row = .Rows - 1
End With
End Sub

Private Sub SetParent()
Set RsQ = Nothing
mAcList = ""
Set RsQ = Nothing
RsQ.Open "Select * From QGroup Where SACode=" & mAcCode & " Order By AcCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Do While RsQ.EOF = False
    If mAcList = "" Then mAcList = CStr(RsQ.Fields("AcCode")) Else mAcList = mAcList + "," + CStr(RsQ.Fields("AcCode"))
    RsQ.MoveNext
Loop
End Sub
Private Function SetGrossInc() As Double
Dim RsQry As New ADODB.Recordset
Dim mTotal As Double
RsQry.Open "Select Sum(CBal) As TotRs From QTrialBal Where AcCode In (" & mAcList & ") And HType=0 And HCode<>54 And LCode<>108", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
mTotal = IIf(IsNull(RsQry.Fields("TotRs")) = True, 0, RsQry.Fields("TotRs"))
SetGrossInc = mTotal
End Function
Private Sub ShowData()
Dim i As Double
With VsfHelp
If CForm = 0 Then
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=3", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.Fields("RTotal") = 0 Then
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=42", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    End If
    .TextMatrix(3, 3) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=4", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.Fields("RTotal") = 0 Then
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode In (44,45,46,47)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    End If
    .TextMatrix(4, 3) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=5", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.Fields("RTotal") = 0 Then
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=48", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    End If
    .TextMatrix(5, 3) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where LCode=70 And ECode<>767", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(7, 2) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where LCode=71 And ECode<>776", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(8, 2) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where ECode=767", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(10, 2) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where ECode=776", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(11, 2) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where LCode In (73,74,83)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(13, 2) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where ECode=781", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(14, 2) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where LCode In (72,106) and ECode<>781", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(16, 2) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where LCode In (75,76)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(17, 2) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where LCode=84", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(18, 3) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=20", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.Fields("RTotal") = 0 Then
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=53 And LCode Not In (84, 108)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    End If
    .TextMatrix(20, 3) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=1", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(24, 2) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=13", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(25, 2) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=29", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(29, 3) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=30", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(30, 3) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode Between 15 And 40 And LCode<>107", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(36, 3) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=34", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(37, 2) = CStr(Round(RsQ.Fields("RTotal"), 2) * -1)
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=33 Or LCode=48", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(38, 2) = CStr(Round(RsQ.Fields("RTotal"), 2) * -1)
    .TextMatrix(38, 3) = CStr(Round(Val(.TextMatrix(37, 2)) + Val(.TextMatrix(38, 2)), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode In (6,8) And Side='D'", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(41, 3) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where ECode In (45,49,52,71,75,79,82,86,87)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(42, 3) = CStr(Round(RsQ.Fields("RTotal"), 2) * -1)
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=46", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(46, 3) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=47", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(47, 3) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=50", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(50, 3) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=51", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(51, 3) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select LCode,OPBal As RTotal From OpDtl Where AcCode=" & mAcCode & " Order By LCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.BOF = False Then
        RsQ.MoveFirst
        RsQ.Find "LCode=-1"
    End If
    If RsQ.EOF = False Then .TextMatrix(54, 3) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=55", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(55, 3) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=62", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(62, 3) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=63", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(63, 3) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing
    RsQ.Open "Select IIf(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl where HCode In (42,44,45,46,47,48,53) And LCode Not In (108,84)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
'    IntInc = RsQ.Fields("RTotal")
'    TaxInc = Val(.TextMatrix(3, 3)) + Val(.TextMatrix(4, 3)) + Val(.TextMatrix(5, 3)) + Val(.TextMatrix(20, 3))
Else
    Dim X As Double
    Set RsQ = Nothing 'Rent Income
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=3", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.Fields("RTotal") = 0 Then
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=42", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(3, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Else
        .TextMatrix(3, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    End If
    Set RsQ = Nothing 'Dividend Income
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=4", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.Fields("RTotal") = 0 Then
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=48", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(4, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Else
        .TextMatrix(4, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    End If
    Set RsQ = Nothing 'Interest Income
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=5", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.Fields("RTotal") = 0 Then
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode In (44,45,46,47)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(5, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Else
        .TextMatrix(5, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    End If
    Set RsQ = Nothing 'Agricultural Income
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where LCode=84", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(6, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Net Consideration on Transfer of Capital Asset
    RsQ.Open "Select IIf(IsNull(Sum(IIF(Side='C',Amt,0)))=True,0,Sum(IIF(Side='C',Amt,0))) As RTotal From TmpCtDtl Where LCode=527 Or ECode In (45,49,52,71,75,79,82,86,87)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(7, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Any other income
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=53 And LCode Not In (84,108,527)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(8, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Government Grants
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where LCode In (73,74,83)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(13, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'CSR Grants
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where ECode=781", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(14, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Other Grants
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where LCode In (72,106) and ECode<>781", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(15, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Local Donations
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where LCode=70", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(16, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Foreign Donations
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where LCode=71", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(18, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Foreign Grants
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where LCode In (75,76)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(19, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Local Corpus Donation
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=1", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(21, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Foreign Corpus Donation
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=13", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(23, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Unutilised Accumulation u/s 11(2)
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=28", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(28, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Unutilised Option under Explanation (2) to 11(1)
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=29", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(29, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Admin Expenses - Repairs & Maintenance
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=17", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(38, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Admin Expenses - Salary
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=18", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(39, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Admin Expenses - Insurance
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=19", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(40, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Admin Expenses - Professional Fees
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=24", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(44, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Admin Expenses - Rates & Taxes
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=16", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(47, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Admin Expenses - Interest
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where ECode=1018", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(48, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Admin Expenses - Audit Fees
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=25", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(49, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Admin Expenses - Other Expenses
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode In (21,26,32) And LCode Not In (107,524) And ECode<>1018 ", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(50, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Religious
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=36 and ECode Not In (341,1037)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(55, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Relief of Poverty
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=39 and ECode<>120", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(56, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Educational
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=37 and ECode Not In (501,1028,1030)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(57, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Yoga
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=58", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(58, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Medical Relief
    RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=38 and ECode<>417", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(59, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Preservation of Environment
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=60", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(60, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Preservation of Monuments, etc.
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=61", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(61, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Source of Fund - Revenue - Accumulation 11(2)
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=66", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(66, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Source of Fund - Revenue - Corpus Fund
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=69", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(69, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Source of Fund - Revenue - Borrowed Fund
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=70", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(70, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Source of Fund - Revenue - Amount Unpaid
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=73", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(73, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Source of Fund - Revenue - Unpaid Amount Reclaimed
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=74", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(74, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Capital WIP
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=79", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(79, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Capital Expenses
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=82", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(82, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Capital Asset Purchase
    RsQ.Open "Select IIF(IsNull(Sum(IIf(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode In (6,8) And Side='D'", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.Fields("RTotal") <> 0 Then
        If RsQ.Fields("RTotal") > Val(.TextMatrix(7, 1)) Then
            .TextMatrix(81, 1) = CStr(Round(Val(.TextMatrix(7, 1)), 2))
            .TextMatrix(80, 1) = CStr(Round(RsQ.Fields("RTotal") - Val(.TextMatrix(79, 1)) - Val(.TextMatrix(81, 1)) - Val(.TextMatrix(82, 1)), 2))
        Else
            X = RsQ.Fields("RTotal")
            Set RsQ = Nothing
            RsQ.Open "Select Sum(IIF((DBal-OpDr)>0,DBal-OpDr,0)) as RTotal from TmpTrialBal where (HCode=7 or LCode=10001) and LCode<>5", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
            If RsQ.EOF = False Then .TextMatrix(81, 1) = CStr(Round(IIf((X + Val(RsQ.Fields("RTotal"))) > Val(.TextMatrix(7, 1)), Val(.TextMatrix(7, 1)), (X + Val(RsQ.Fields("RTotal")))), 2)) Else .TextMatrix(81, 1) = CStr(Round(X, 2))
            .TextMatrix(79, 1) = 0
            .TextMatrix(80, 1) = 0
            .TextMatrix(82, 1) = 0
        End If
        If Val(.TextMatrix(81, 1)) < Val(.TextMatrix(7, 1)) Then
            MsgBox "Purchase of New Capital Asset is less than Sale Consideration. Calculate Capital Gain as per Section 45." & vbCrLf & vbCrLf & "Values of Sale Consideration and Purchase of Capital Asset will be removed now. Please enter the values manually in ITR.", vbCritical, "Alert"
            .TextMatrix(81, 1) = 0
            .TextMatrix(7, 1) = 0
        End If
    End If
    Set RsQ = Nothing 'Source of Fund - Capital - CY Income
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=85", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(85, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Source of Fund - Capital - Accumulation 11(2)
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=86", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(86, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Source of Fund - Capital - PY Option
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=87", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(87, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Source of Fund - Capital - 15% Surplus
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=88", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(88, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Source of Fund - Capital - Corpus Fund
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=89", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(89, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Source of Fund - Capital - Borrowed Funds
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=90", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(90, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Source of Fund - Capital - Amount Unpaid
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=93", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(93, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Source of Fund - Capital - Unpaid Amount Reclaimed
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=94", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(94, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Repayment of Borrowed Funds
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=98", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(98, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Repatriation to Corpus Fund
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=99", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(99, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Accumulation u/s 11(2)
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=107", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(107, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'Option under Explanation (2) to 11(1)
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=108", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    .TextMatrix(108, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    Set RsQ = Nothing 'IntInc and TaxInc
    RsQ.Open "Select IIf(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl where HCode In (42,44,45,46,47,48,53) And LCode Not In (108,84)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
    SetFinalTot
    Set RsQ = Nothing 'Donation Given
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=54", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.Fields("RTotal") = 0 Then
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where ECode In (120,266,341,417,501,1028,1030,1037)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
        TotDon = Round(RsQ.Fields("RTotal"), 2)
        .TextMatrix(54, 1) = CStr(Round(RsQ.Fields("RTotal") - Val(.TextMatrix(53, 1)), 2))
    Else
        .TextMatrix(54, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where ECode In (120,266,341,417,501,1028,1030,1037)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
        TotDon = Round(RsQ.Fields("RTotal"), 2)
    End If
    Set RsQ = Nothing 'General Public Utility
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=62", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.Fields("RTotal") = 0 Then
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=40 and ECode<>266", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
        OthExp = Round(RsQ.Fields("RTotal"), 2)
        X = RsQ.Fields("RTotal") - Val(.TextMatrix(58, 1)) - Val(.TextMatrix(60, 1)) - Val(.TextMatrix(61, 1))
        .TextMatrix(62, 1) = CStr(Round(X, 2))
    Else
        .TextMatrix(62, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(IIF(Side=HSide,Amt,Amt*-1)))=True,0,Sum(IIF(Side=HSide,Amt,Amt*-1))) As RTotal From TmpCtDtl Where HCode=40 and ECode<>266", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
        OthExp = Round(RsQ.Fields("RTotal"), 2)
    End If
    SetFinalTot
    Set RsQ = Nothing 'Source of Fund - Revenue - PY Option
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=67", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.Fields("RTotal") = 0 Then
        Set RsQ = Nothing
        RsQ.Open "Select LCode, IIF(IsNull(OpCr)=True,0,OpCr) As RTotal From TmpTrialBal Where LCode=-1", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
        If RsQ.BOF = False Then
            RsQ.MoveFirst
            RsQ.Find "LCode = -1"
        End If
        If RsQ.EOF = False Then
            X = IIf((RsQ.Fields("RTotal") - Val(.TextMatrix(87, 1))) < Val(.TextMatrix(63, 1)), (RsQ.Fields("RTotal") - Val(.TextMatrix(87, 1))), Val(.TextMatrix(63, 1)))
            .TextMatrix(67, 1) = CStr(Round(X, 2))
        End If
    Else
    .TextMatrix(67, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    End If
    SetFinalTot
    Set RsQ = Nothing 'Source of Fund - Revenue - CY Income
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=65", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.Fields("RTotal") = 0 Then
        X = IIf((Val(.TextMatrix(63, 1)) - Val(.TextMatrix(66, 1)) - Val(.TextMatrix(67, 1)) - Val(.TextMatrix(69, 1)) - Val(.TextMatrix(70, 1))) < (Val(.TextMatrix(9, 2)) + Val(.TextMatrix(25, 2))), Val(.TextMatrix(63, 1)) - Val(.TextMatrix(66, 1)) - Val(.TextMatrix(67, 1)) - Val(.TextMatrix(69, 1)) - Val(.TextMatrix(70, 1)), Val(.TextMatrix(9, 2)) + Val(.TextMatrix(25, 2)))
        .TextMatrix(65, 1) = CStr(Round(X, 2))
    Else
        .TextMatrix(65, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    End If
    SetFinalTot
    Set RsQ = Nothing 'Source of Fund - Capital - CY Income
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=65", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.Fields("RTotal") = 0 Then
        X = IIf((Val(.TextMatrix(83, 1)) - Val(.TextMatrix(86, 1)) - Val(.TextMatrix(87, 1)) - Val(.TextMatrix(89, 1)) - Val(.TextMatrix(90, 1))) < (Val(.TextMatrix(9, 2)) + Val(.TextMatrix(25, 2)) - Val(VsfHelp.TextMatrix(65, 1))), Val(.TextMatrix(83, 1)) - Val(.TextMatrix(86, 1)) - Val(.TextMatrix(87, 1)) - Val(.TextMatrix(89, 1)) - Val(.TextMatrix(90, 1)), Val(.TextMatrix(9, 2)) + Val(.TextMatrix(25, 2)) - Val(VsfHelp.TextMatrix(65, 1)))
        .TextMatrix(85, 1) = CStr(Round(X, 2))
    Else
        .TextMatrix(85, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    End If
    SetFinalTot
    Set RsQ = Nothing 'Source of Fund - Capital - 15% Surplus
    RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As RTotal From ITDtl Where AcCode=" & mAcCode & " And SrN=88", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.Fields("RTotal") = 0 Then
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(IIf(OpCr>0,OpCr,0)))=True,0,Sum(IIf(OpCr>0,OpCr,0))) As RTotal From TmpTrialBal Where HCode In (2,5)", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
        X = IIf((RsQ.Fields("RTotal") - Val(.TextMatrix(85, 1))) < (Val(.TextMatrix(83, 1)) - Val(.TextMatrix(86, 1)) - Val(.TextMatrix(87, 1)) - Val(.TextMatrix(89, 1)) - Val(.TextMatrix(90, 1))), RsQ.Fields("RTotal") - Val(.TextMatrix(85, 1)), Val(.TextMatrix(83, 1)) - Val(.TextMatrix(85, 1)) - Val(.TextMatrix(86, 1)) - Val(.TextMatrix(87, 1)) - Val(.TextMatrix(89, 1)) - Val(.TextMatrix(90, 1)))
        .TextMatrix(88, 1) = CStr(Round(X, 2))
    Else
    .TextMatrix(88, 1) = CStr(Round(RsQ.Fields("RTotal"), 2))
    End If
    SetFinalTot
End If
End With
SetFinalTot
End Sub

Private Sub SaveData()
Dim i As Double
On Error GoTo XErr
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From ITDtl Where AcCode=" & mAcCode
With VsfHelp
If CForm = 0 Then
    i = 3
    If Val(.TextMatrix(i, 3)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-7,'C'," & Val(.TextMatrix(i, 3)) & "," & i & ",0)"
    i = 4
    If Val(.TextMatrix(i, 3)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-5,'C'," & Val(.TextMatrix(i, 3)) & "," & i & ",0)"
    i = 5
    If Val(.TextMatrix(i, 3)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-6,'C'," & Val(.TextMatrix(i, 3)) & "," & i & ",0)"
    i = 20
    If Val(.TextMatrix(i, 3)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-8,'C'," & Val(.TextMatrix(i, 3)) & "," & i & ",0)"
    i = 29
    If Val(.TextMatrix(i, 3)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-2,'D'," & Val(.TextMatrix(i, 3)) & "," & i & ",0)"
    i = 30
    If Val(.TextMatrix(i, 3)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-1,'D'," & Val(.TextMatrix(i, 3)) & "," & i & ",0)"
    i = 46
    If Val(.TextMatrix(i, 3)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-3,'D'," & Val(.TextMatrix(i, 3)) & "," & i & ",0)"
    i = 47
    If Val(.TextMatrix(i, 3)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-4,'D'," & Val(.TextMatrix(i, 3)) & "," & i & ",0)"
    i = 50
    If Val(.TextMatrix(i, 3)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-3,'C'," & Val(.TextMatrix(i, 3)) & "," & i & ",0)"
    i = 51
    If Val(.TextMatrix(i, 3)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-4,'C'," & Val(.TextMatrix(i, 3)) & "," & i & ",0)"
    i = 54
    If Val(.TextMatrix(i, 3)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-1,'D'," & Val(.TextMatrix(i, 3)) & "," & i & ",0)"
    i = 55
    If Val(.TextMatrix(i, 3)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-2,'D'," & Val(.TextMatrix(i, 3)) & "," & i & ",0)"
    i = 62
    If Val(.TextMatrix(i, 3)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-2,'C'," & Val(.TextMatrix(i, 3)) & "," & i & ",0)"
    i = 63
    If Val(.TextMatrix(i, 3)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-1,'C'," & Val(.TextMatrix(i, 3)) & "," & i & ",0)"
Else
    If Val(.TextMatrix(3, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-11,'C'," & Val(.TextMatrix(3, 1)) & ",3,1)"
    If Val(.TextMatrix(4, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-11,'C'," & Val(.TextMatrix(4, 1)) & ",4,1)"
    If Val(.TextMatrix(5, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-11,'C'," & Val(.TextMatrix(5, 1)) & ",5,1)"
    If Val(.TextMatrix(28, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-2,'D'," & Val(.TextMatrix(28, 1)) & ",28,1)"
    If Val(.TextMatrix(29, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-1,'D'," & Val(.TextMatrix(29, 1)) & ",29,1)"
    If Val(.TextMatrix(53, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-11,'D'," & Val(.TextMatrix(53, 1)) & ",53,1)"
    If Val(.TextMatrix(58, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-11,'D'," & Val(.TextMatrix(58, 1)) & ",58,1)"
    If Val(.TextMatrix(60, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-11,'D'," & Val(.TextMatrix(60, 1)) & ",60,1)"
    If Val(.TextMatrix(61, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-11,'D'," & Val(.TextMatrix(61, 1)) & ",61,1)"
    If Val(.TextMatrix(66, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-11,'D'," & Val(.TextMatrix(66, 1)) & ",66,1)"
    If Val(.TextMatrix(69, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-4,'C'," & Val(.TextMatrix(69, 1)) & ",69,1)"
    If Val(.TextMatrix(70, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-3,'C'," & Val(.TextMatrix(70, 1)) & ",70,1)"
    If Val(.TextMatrix(73, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-10,'C'," & Val(.TextMatrix(73, 1)) & ",73,1)"
    If Val(.TextMatrix(74, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-10,'D'," & Val(.TextMatrix(74, 1)) & ",74,1)"
    If Val(.TextMatrix(85, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-11,'D'," & Val(.TextMatrix(85, 1)) & ",85,1)"
    If Val(.TextMatrix(86, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-2,'D'," & Val(.TextMatrix(86, 1)) & ",86,1)"
    If Val(.TextMatrix(87, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-1,'D'," & Val(.TextMatrix(87, 1)) & ",87,1)"
    If Val(.TextMatrix(88, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-11,'D'," & Val(.TextMatrix(88, 1)) & ",88,1)"
    If Val(.TextMatrix(89, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-4,'C'," & Val(.TextMatrix(89, 1)) & ",89,1)"
    If Val(.TextMatrix(90, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-3,'C'," & Val(.TextMatrix(90, 1)) & ",90,1)"
    If Val(.TextMatrix(93, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-10,'C'," & Val(.TextMatrix(93, 1)) & ",93,1)"
    If Val(.TextMatrix(94, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-10,'D'," & Val(.TextMatrix(94, 1)) & ",94,1)"
    If Val(.TextMatrix(98, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-3,'C'," & Val(.TextMatrix(98, 1)) & ",98,1)"
    If Val(.TextMatrix(99, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-4,'C'," & Val(.TextMatrix(99, 1)) & ",99,1)"
    If Val(.TextMatrix(107, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-2,'C'," & Val(.TextMatrix(107, 1)) & ",107,1)"
    If Val(.TextMatrix(108, 1)) <> 0 Then DbDataDB.Execute "Insert InTo ITDtl (AcCode,LCode,Side,Amt,SrN,CForm) Values (" & mAcCode & ",-1,'C'," & Val(.TextMatrix(108, 1)) & ",108,1)"
End If
End With
DbDataDB.CommitTrans
MsgBox "Record Saved.", vbInformation, "Alert"
Exit Sub
XErr:
MsgBox Err.Description
DbDataDB.RollbackTrans
End Sub

Private Sub PrintRec()
Dim RsClient As New ADODB.Recordset
Dim RsRep As New ADODB.Recordset
Dim i As Integer
RsRep.Open "Select RepDtl.*,AdtMst.AdtName From RepDtl,AdtMst Where AcCode=" & mAcCode & " And RepDtl.AdtNo=AdtMst.AdtNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsRep.EOF = False Then
    If Len(RsRep.Fields("RUDIN")) = 0 Then
        MsgBox "UDIN not issued. Report cannot be printed.", vbCritical, "Alert"
        Exit Sub
    End If
Else
    MsgBox "Signing Information not available. Report cannot be printed.", vbCritical, "Alert"
    Exit Sub
End If
RsClient.Open "Select * From AcMst Where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
If RsClient.EOF = False Then
    If CForm = 0 Then
        With RepPrint
            .Connect = MSCONNECT
            .ReportFileName = App.Path + "\Report\ITRRpt.Rpt"
            .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
            .Formulas(1) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
            .Formulas(2) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
            .Formulas(3) = "mAdd1='" & RsComp.Fields("Add1") & "'"
            .Formulas(4) = "mAdd2='" & RsComp.Fields("Add2") & "'"
            .Formulas(5) = "mCity='" & RsComp.Fields("City") & ", " & RsComp.Fields("State") & ", " & RsComp.Fields("PinCode") & "'"
            .Formulas(6) = "mPhone='" & RsComp.Fields("Phone") & "'"
            .Formulas(7) = "mFileNo='" & RsClient.Fields("FileNo") & "'"
            .Formulas(8) = "mITRFY='" & Mid(mYear, 7) & "-" & Mid(mTYear, 7) & "'"
            .Formulas(9) = "mITRAY='" & CStr(Val(Mid(mYear, 7)) + 1) & "-" & CStr(Val(Mid(mTYear, 7)) + 1) & "'"
            .Formulas(10) = "mTName='" & RsClient.Fields("AcName") & ", " & RsClient.Fields("City") & "'"
            .Formulas(11) = "mPAN='" & RsClient.Fields("TPAN") & "'"
            .Formulas(12) = "mRPDate='" & RsRep.Fields("RpDt") & "'"
            .Formulas(13) = "mTitleITR='Draft Computation of Income chargeable to Tax'"
            If RsOpDtl.BOF = False Then
                RsOpDtl.MoveFirst
                RsOpDtl.Find "LCode=-1"
            End If
            If RsOpDtl.EOF = False Then
                .Formulas(14) = "mOpt1112='" & Format(RsOpDtl.Fields("OpBal"), "0.00") & "'"
            Else
                .Formulas(14) = "mOpt1112='0.00'"
            End If
            If RsOpDtl.BOF = False Then
                RsOpDtl.MoveFirst
                RsOpDtl.Find "LCode=-2"
            End If
            If RsOpDtl.EOF = False Then
                .Formulas(15) = "mAcu1125='" & Format(RsOpDtl.Fields("OpBal"), "0.00") & "'"
            Else
                .Formulas(15) = "mAcu1125='0.00'"
            End If
            If RsOpDtl.BOF = False Then
                RsOpDtl.MoveFirst
                RsOpDtl.Find "LCode=-3"
            End If
            If RsOpDtl.EOF = False Then
                .Formulas(16) = "mAppBor='" & Format(RsOpDtl.Fields("OpBal"), "0.00") & "'"
            Else
                .Formulas(16) = "mAppBor='0.00'"
            End If
            If RsOpDtl.BOF = False Then
                RsOpDtl.MoveFirst
                RsOpDtl.Find "LCode=-4"
            End If
            If RsOpDtl.EOF = False Then
                .Formulas(17) = "mAppCorp='" & Format(RsOpDtl.Fields("OpBal"), "0.00") & "'"
            Else
                .Formulas(17) = "mAppCorp='0.00'"
            End If
            .Formulas(18) = "mRnt='" & Format(VsfHelp.TextMatrix(3, 3), "0.00") & "'"
            .Formulas(19) = "mInt='" & Format(VsfHelp.TextMatrix(4, 3), "0.00") & "'"
            .Formulas(20) = "mDiv='" & Format(VsfHelp.TextMatrix(5, 3), "0.00") & "'"
            .Formulas(21) = "mDoCL='" & Format(VsfHelp.TextMatrix(7, 2), "0.00") & "'"
            .Formulas(22) = "mDoCF='" & Format(VsfHelp.TextMatrix(8, 2), "0.00") & "'"
            .Formulas(23) = "mDoC='" & Format(VsfHelp.TextMatrix(8, 3), "0.00") & "'"
            .Formulas(24) = "mDoKL='" & Format(VsfHelp.TextMatrix(10, 2), "0.00") & "'"
            .Formulas(25) = "mDoKF='" & Format(VsfHelp.TextMatrix(11, 2), "0.00") & "'"
            .Formulas(26) = "mDoK='" & Format(VsfHelp.TextMatrix(11, 3), "0.00") & "'"
            .Formulas(27) = "mGrnG='" & Format(VsfHelp.TextMatrix(13, 2), "0.00") & "'"
            .Formulas(28) = "mGrnC='" & Format(VsfHelp.TextMatrix(14, 2), "0.00") & "'"
            .Formulas(29) = "mGrnL='" & Format(VsfHelp.TextMatrix(16, 2), "0.00") & "'"
            .Formulas(30) = "mGrnF='" & Format(VsfHelp.TextMatrix(17, 2), "0.00") & "'"
            .Formulas(31) = "mGrn='" & Format(VsfHelp.TextMatrix(17, 3), "0.00") & "'"
            .Formulas(32) = "mAgr='" & Format(VsfHelp.TextMatrix(18, 3), "0.00") & "'"
            .Formulas(33) = "mCapGn='" & Format(VsfHelp.TextMatrix(19, 3), "0.00") & "'"
            .Formulas(34) = "mOIn='" & Format(VsfHelp.TextMatrix(20, 3), "0.00") & "'"
            .Formulas(35) = "mTInc='" & IIf(VsfHelp.TextMatrix(21, 3) = "", "0.00", Format(VsfHelp.TextMatrix(21, 3), "0.00")) & "'"
            .Formulas(36) = "mCorL='" & Format(VsfHelp.TextMatrix(24, 2), "0.00") & "'"
            .Formulas(37) = "mCorF='" & Format(VsfHelp.TextMatrix(25, 2), "0.00") & "'"
            .Formulas(38) = "mCor='" & Format(VsfHelp.TextMatrix(25, 3), "0.00") & "'"
            .Formulas(39) = "mECor='" & Format(VsfHelp.TextMatrix(26, 3), "0.00") & "'"
            .Formulas(40) = "mCDon='" & IIf(VsfHelp.TextMatrix(27, 3) = "", "0.00", Format(VsfHelp.TextMatrix(27, 3), "0.00")) & "'"
            .Formulas(41) = "mUAcc='" & Format(VsfHelp.TextMatrix(29, 3), "0.00") & "'"
            .Formulas(42) = "mUOpt='" & Format(VsfHelp.TextMatrix(30, 3), "0.00") & "'"
            .Formulas(43) = "mUnAc='" & IIf(VsfHelp.TextMatrix(31, 3) = "", "0.00", Format(VsfHelp.TextMatrix(31, 3), "0.00")) & "'"
            .Formulas(44) = "mTotalInc='" & IIf(VsfHelp.TextMatrix(32, 4) = "", "0.00", Format(VsfHelp.TextMatrix(32, 4), "0.00")) & "'"
            .Formulas(45) = "mTExp='" & Format(VsfHelp.TextMatrix(36, 3), "0.00") & "'"
            .Formulas(46) = "mTrf='" & Format(VsfHelp.TextMatrix(37, 3), "0.00") & "'"
            .Formulas(47) = "mDep='" & Format(VsfHelp.TextMatrix(38, 3), "0.00") & "'"
            .Formulas(48) = "mRevEx='" & IIf(VsfHelp.TextMatrix(39, 3) = "", "0.00", Format(VsfHelp.TextMatrix(39, 3), "0.00")) & "'"
            .Formulas(49) = "mPurFA='" & Format(VsfHelp.TextMatrix(41, 3), "0.00") & "'"
            .Formulas(50) = "mSalFA='" & Format(VsfHelp.TextMatrix(42, 3), "0.00") & "'"
            .Formulas(51) = "mCapFA='" & Format(VsfHelp.TextMatrix(43, 3), "0.00") & "'"
            .Formulas(52) = "mCapEx='" & IIf(VsfHelp.TextMatrix(44, 3) = "", "0.00", Format(VsfHelp.TextMatrix(44, 3), "0.00")) & "'"
            .Formulas(53) = "mLnRep='" & Format(VsfHelp.TextMatrix(46, 3), "0.00") & "'"
            .Formulas(54) = "mCrRep='" & Format(VsfHelp.TextMatrix(47, 3), "0.00") & "'"
            .Formulas(55) = "mReQu='" & IIf(VsfHelp.TextMatrix(48, 3) = "", "0.00", Format(VsfHelp.TextMatrix(48, 3), "0.00")) & "'"
            .Formulas(56) = "mLnUti='" & Format(VsfHelp.TextMatrix(50, 3), "0.00") & "'"
            .Formulas(57) = "mCrUti='" & Format(VsfHelp.TextMatrix(51, 3), "0.00") & "'"
            .Formulas(58) = "mNoQu='" & IIf(VsfHelp.TextMatrix(52, 3) = "", "0.00", Format(VsfHelp.TextMatrix(52, 3), "0.00")) & "'"
            .Formulas(59) = "mAOpt='" & Format(VsfHelp.TextMatrix(54, 3), "0.00") & "'"
            .Formulas(60) = "mAAcc='" & Format(VsfHelp.TextMatrix(55, 3), "0.00") & "'"
            .Formulas(61) = "mAccEx='" & IIf(VsfHelp.TextMatrix(56, 3) = "", "0.00", Format(VsfHelp.TextMatrix(56, 3), "0.00")) & "'"
            .Formulas(62) = "mTotalApp='" & IIf(VsfHelp.TextMatrix(57, 4) = "", "0.00", Format(VsfHelp.TextMatrix(57, 4), "0.00")) & "'"
            .Formulas(63) = "mSurplus='" & IIf(VsfHelp.TextMatrix(60, 4) = "", "0.00", Format(VsfHelp.TextMatrix(60, 4), "0.00")) & "'"
            .Formulas(64) = "mSetInc='" & IIf(VsfHelp.TextMatrix(61, 4) = "", "0.00", Format(VsfHelp.TextMatrix(61, 4), "0.00")) & "'"
            .Formulas(65) = "mAcc='" & Format(VsfHelp.TextMatrix(62, 3), "0.00") & "'"
            .Formulas(66) = "mOpt='" & Format(VsfHelp.TextMatrix(63, 3), "0.00") & "'"
            .Formulas(67) = "mDeem='" & Format(VsfHelp.TextMatrix(64, 4), "0.00") & "'"
            .Formulas(68) = "mNTI='" & IIf(TxtCTotal.Text = "", "0.00", Format(TxtCTotal.Text, "0.00")) & "'"
            If RsOpDtl.BOF = False Then
                RsOpDtl.MoveFirst
                RsOpDtl.Find "LCode=-10"
            End If
            If RsOpDtl.EOF = False Then
                .Formulas(69) = "mPYCrNA='" & Format(RsOpDtl.Fields("OpBal"), "0.00") & "'"
            Else
                .Formulas(69) = "mPYCrNA='0.00'"
            End If
            .Action = 1
            For i = 0 To 69
                .Formulas(i) = ""
            Next
        End With
    Else
        DBLocalDB.BeginTrans
            DBLocalDB.Execute "Delete from TmpRecPrn"
        DBLocalDB.CommitTrans
        DBLocalDB.BeginTrans
            With VsfHelp
                For i = 1 To .Rows - 1
                DBLocalDB.Execute "Insert into TmpRecPrn (SrN,LName,LAmt,RAmt,RName) Values (" & i & ",'" & .TextMatrix(i, 0) & "'," & Val(.TextMatrix(i, 1)) & "," & Val(.TextMatrix(i, 2)) & "," & Val(.TextMatrix(i, 3)) & ")"
                Next
            End With
            i = i + 1
            DBLocalDB.Execute "Insert into TmpRecPrn (SrN,LName,LAmt,RAmt,RName) Values (" & i & ",'',0,0,2)"
            i = i + 1
            DBLocalDB.Execute "Insert into TmpRecPrn (SrN,LName,LAmt,RAmt,RName) Values (" & i & ",'Net Taxable Income (1 - (2 + 3))',0,'" & CStr(Format(TxtCTotal.Text, "0.00")) & "',5)"
            Set RsQ = Nothing
            RsQ.Open "Select EName, Sum(IIf(Side=HSide,Amt,Amt*-1)) As RTotal from TmpCtDtl Where HCode=53 And LCode Not In (84,108) Group By EName Order By EName", DBLocalDB, adOpenDynamic, adLockReadOnly, adCmdText
            If RsQ.EOF = False Then
                i = i + 1
                DBLocalDB.Execute "Insert into TmpRecPrn (SrN,LName,LAmt,RAmt,RName) Values (" & i & ",'',0,0,2)"
                i = i + 1
                DBLocalDB.Execute "Insert Into TmpRecPrn (SrN,LName,LAmt,RAmt,RName) Values (" & i & ",'Details of Other Income (Point A(1)(iv)',0,0,1)"
                Do While RsQ.EOF = False
                    i = i + 1
                    DBLocalDB.Execute "Insert Into TmpRecPrn (SrN,LName,LAmt,RAmt,RName) Values (" & i & ",'" & RsQ.Fields("EName") & "'," & CStr(Format(RsQ.Fields("RTotal"), "0.00")) & ",0,2)"
                    RsQ.MoveNext
                Loop
                i = i + 1
                DBLocalDB.Execute "Insert into TmpRecPrn (SrN,LName,LAmt,RAmt,RName) Values (" & i & ",'',0,0,6)"
            End If
        DBLocalDB.CommitTrans
        With RepPrint
            .Connect = MSCONNECT
            .ReportFileName = App.Path + "\Report\ITRRptX.Rpt"
            .Formulas(0) = "mCmpName='" & RsComp.Fields("FName") & "'"
            .Formulas(1) = "mFRN='(FRN: " & RsComp.Fields("FRN") & ")'"
            .Formulas(2) = "mEmail='E-mail: " & RsComp.Fields("EMail") & "'"
            .Formulas(3) = "mAdd1='" & RsComp.Fields("Add1") & "'"
            .Formulas(4) = "mAdd2='" & RsComp.Fields("Add2") & "'"
            .Formulas(5) = "mCity='" & RsComp.Fields("City") & ", " & RsComp.Fields("State") & ", " & RsComp.Fields("PinCode") & "'"
            .Formulas(6) = "mPhone='" & RsComp.Fields("Phone") & "'"
            .Formulas(7) = "mFileNo='" & RsClient.Fields("FileNo") & "'"
            .Formulas(8) = "mITRFY='" & Mid(mYear, 7) & "-" & Mid(mTYear, 7) & "'"
            .Formulas(9) = "mITRAY='" & CStr(Val(Mid(mYear, 7)) + 1) & "-" & CStr(Val(Mid(mTYear, 7)) + 1) & "'"
            .Formulas(10) = "mTName='" & RsClient.Fields("AcName") & ", " & RsClient.Fields("City") & "'"
            .Formulas(11) = "mPAN='" & RsClient.Fields("TPAN") & "'"
            .Formulas(12) = "mRPDate='" & RsRep.Fields("RpDt") & "'"
            .Formulas(13) = "mTitleITR='Draft Computation of Income chargeable to Tax'"
            If RsOpDtl.BOF = False Then
                RsOpDtl.MoveFirst
                RsOpDtl.Find "LCode=-1"
            End If
            If RsOpDtl.EOF = False Then
                .Formulas(14) = "mOpt1112='" & Format(RsOpDtl.Fields("OpBal"), "0.00") & "'"
            Else
                .Formulas(14) = "mOpt1112='0.00'"
            End If
            If RsOpDtl.BOF = False Then
                RsOpDtl.MoveFirst
                RsOpDtl.Find "LCode=-2"
            End If
            If RsOpDtl.EOF = False Then
                .Formulas(15) = "mAcu1125='" & Format(RsOpDtl.Fields("OpBal"), "0.00") & "'"
            Else
                .Formulas(15) = "mAcu1125='0.00'"
            End If
            If RsOpDtl.BOF = False Then
                RsOpDtl.MoveFirst
                RsOpDtl.Find "LCode=-3"
            End If
            If RsOpDtl.EOF = False Then
                .Formulas(16) = "mAppBor='" & Format(RsOpDtl.Fields("OpBal"), "0.00") & "'"
            Else
                .Formulas(16) = "mAppBor='0.00'"
            End If
            If RsOpDtl.BOF = False Then
                RsOpDtl.MoveFirst
                RsOpDtl.Find "LCode=-4"
            End If
            If RsOpDtl.EOF = False Then
                .Formulas(17) = "mAppCorp='" & Format(RsOpDtl.Fields("OpBal"), "0.00") & "'"
            Else
                .Formulas(17) = "mAppCorp='0.00'"
            End If
            .Action = 1
            .Formulas(18) = "mBranch1='" & RsComp.Fields("Branch1") & "'"
            .Formulas(19) = "mBranch2='" & RsComp.Fields("Branch2") & "'"
            If RsOpDtl.BOF = False Then
                RsOpDtl.MoveFirst
                RsOpDtl.Find "LCode=-10"
            End If
            If RsOpDtl.EOF = False Then
                .Formulas(20) = "mPYCrNA='" & Format(RsOpDtl.Fields("OpBal"), "0.00") & "'"
            Else
                .Formulas(20) = "mPYCrNA='0.00'"
            End If
            For i = 0 To 20
                .Formulas(i) = ""
            Next
        End With
    End If
End If
End Sub
