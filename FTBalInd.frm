VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Begin VB.Form FrmTBalInd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trial Balance (Individual)"
   ClientHeight    =   7755
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
   Icon            =   "FTBalInd.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7755
   ScaleWidth      =   13140
   Begin VB.TextBox TxtOTotal 
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
      Left            =   5550
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   7200
      Width           =   2000
   End
   Begin VB.TextBox TxtATotal 
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
      Left            =   7550
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   7200
      Width           =   1500
   End
   Begin VB.TextBox TxtWTotal 
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
      Left            =   9050
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   7200
      Width           =   1500
   End
   Begin VB.Frame FraClientHelp 
      Height          =   4212
      Left            =   18000
      TabIndex        =   6
      Top             =   1920
      Width           =   12705
      Begin VSFlex7Ctl.VSFlexGrid LsvClient 
         Height          =   3855
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   12465
         _cx             =   21987
         _cy             =   6800
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
         FormatString    =   $"FTBalInd.frx":0442
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
         Height          =   4095
         Index           =   1
         Left            =   0
         Top             =   120
         Width           =   12705
      End
   End
   Begin VB.Frame FraMst 
      Height          =   7695
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   12972
      Begin VB.CommandButton CmdExport 
         Caption         =   "Export Data"
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
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Exit"
         Top             =   240
         Width           =   1450
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
            FormatString    =   $"FTBalInd.frx":048B
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
         Left            =   10450
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   7200
         Width           =   2000
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
         Picture         =   "FTBalInd.frx":04D4
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
         Height          =   6420
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   12735
         _cx             =   22463
         _cy             =   11324
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
         FormatString    =   $"FTBalInd.frx":0B0A
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
         Height          =   255
         Left            =   7440
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   975
         _cx             =   1720
         _cy             =   450
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
         FormatString    =   $"FTBalInd.frx":0B53
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
         Height          =   7575
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   12975
      End
   End
End
Attribute VB_Name = "FrmTBalInd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DBWorkTmp As New ADODB.Connection
Dim mAcCode As Double
Dim mAcList As String

Private Sub CmdExport_Click()
    ExpMainData
    DbDataDB.BeginTrans
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'TRIAL_BALANCE','EXPORT_XLS','" & Date & "','" & Time & "')"
    DbDataDB.CommitTrans
End Sub

Private Sub CmdSearch_Click()
    FraClientHelp.Left = 180
    LsvClient.SetFocus
End Sub

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
RsQry.Open "Select AcMst.AcName,AcMst.FileNo,AcMst.City,'' As ParentNm,GrpMst.GName,AcMst.AcCode,'' As RName From AcMst,GrpMst Where AcMst.Active=-1 And AcMst.PaCode=0 And " & _
"AcMst.AcType=GrpMst.GCode Union All Select AcMst.AcName,AcMst.FileNo,AcMst.City,AcMst1.AcName+IIF(Len(AcMst1.City)>0,', '+AcMst1.City,''),GrpMst.GName,AcMst.AcCode," & _
"'' From AcMst,GrpMst,AcMst As AcMst1 Where AcMst.Active=-1 And AcMst.PaCode=AcMst1.AcCode And AcMst.AcType=GrpMst.GCode Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
    DbDataDB.BeginTrans
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'TRIAL_BALANCE','VIEW','" & Date & "','" & Time & "')"
    DbDataDB.CommitTrans
    DBWorkTmp.BeginTrans
    DBWorkTmp.Execute "Delete From TmpTrialBal"
    DBWorkTmp.Execute "Delete From TmpCtDtl"
    DBWorkTmp.Execute "Delete From TmpBSPrn"
    DBWorkTmp.Execute "Delete From TmpPLPrn"
    DBWorkTmp.Execute "Delete From TmpRecPrn"
    DBWorkTmp.CommitTrans
    DBWorkTmp.BeginTrans
    RsQ.Open "Select * From QTrialBal Where AcCode = " & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        DBWorkTmp.Execute "Insert InTo TmpTrialBal (HType,HSide,AcCode,HCode,LCode,OpDr,OpCr,ADr,ACr,DBal,CBal) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HSide") & _
        "'," & RsQ.Fields("AcCode") & "," & RsQ.Fields("HCode") & "," & RsQ.Fields("LCode") & "," & RsQ.Fields("OpDr") & "," & RsQ.Fields("OpCr") & "," & RsQ.Fields("ADr") & _
        "," & RsQ.Fields("ACr") & "," & RsQ.Fields("DBal") & "," & RsQ.Fields("CBal") & ")"
        RsQ.MoveNext
    Loop
    Set RsQ = Nothing
    RsQ.Open "Select * From QCtDtl Where AcCode = " & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        DBWorkTmp.Execute "Insert InTo TmpCtDtl (HType,HSide,AcCode,HCode,LCode,ECode,EName,Side,Amt) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HSide") & "'," & _
        RsQ.Fields("AcCode") & "," & RsQ.Fields("HCode") & "," & RsQ.Fields("LCode") & "," & RsQ.Fields("ECode") & ",'" & RsQ.Fields("EName") & "','" & RsQ.Fields("Side") & _
        "'," & RsQ.Fields("Amt") & ")"
        RsQ.MoveNext
    Loop
    Set RsQ = Nothing
    RsQ.Open "Select QTrialBal.HType, QTrialBal.HSide, QTrialBal.HCode, HedMst.HName From QTrialBal, HedMst Where QTrialBal.AcCode = " & mAcCode & " And QTrialBal.HCode=HedMst.HCode Group By QTrialBal.HType, QTrialBal.HSide, QTrialBal.HCode, HedMst.HName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        DBWorkTmp.Execute "Insert InTo TmpBSPrn (SrN,LName,LAmt,RName) Values (" & RsQ.Fields("HType") & ",'" & RsQ.Fields("HName") & "'," & RsQ.Fields("HCode") & ",'" & RsQ.Fields("HSide") & "')"
        RsQ.MoveNext
    Loop
    Set RsQ = Nothing
    RsQ.Open "Select QTrialBal.HCode, QTrialBal.LCode, LedMst.LName From QTrialBal, LedMst Where QTrialBal.AcCode = " & mAcCode & " And QTrialBal.LCode=LedMst.LCode Group By QTrialBal.HCode, QTrialBal.LCode, LedMst.LName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        DBWorkTmp.Execute "Insert InTo TmpPLPrn (SrN,LName,LAmt) Values (" & RsQ.Fields("HCode") & ",'" & RsQ.Fields("LName") & "'," & RsQ.Fields("LCode") & ")"
        RsQ.MoveNext
    Loop
    Set RsQ = Nothing
    RsQ.Open "Select QCtDtl.LCode,QCtDtl.ECode,EntMst.EName From QCtDtl, EntMst Where QCtDtl.AcCode = " & mAcCode & " And QCtDtl.ECode=EntMst.ECode Group By QCtDtl.LCode, QCtDtl.ECode, EntMst.EName", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        DBWorkTmp.Execute "Insert InTo TmpRecPrn (SrN,LName,LAmt) Values (" & RsQ.Fields("LCode") & ",'" & RsQ.Fields("EName") & "'," & RsQ.Fields("ECode") & ")"
        RsQ.MoveNext
    Loop
    DBWorkTmp.CommitTrans
    SetData
    ShowData
    VsfHelp.SetFocus
End If
End Sub
Private Sub Tlbsav_Click(Index As Integer)
Select Case TlbSav(Index).ToolTipText
    Case "Exit"
        DBWorkTmp.Close
        Unload Me
End Select
End Sub
Private Function SetTool(ByVal mVal As Boolean)
    TlbSav(2).Enabled = True
End Function
Private Sub SetData()
Dim RsQ As New ADODB.Recordset
Dim RsQL As New ADODB.Recordset
Dim RsQE As New ADODB.Recordset
Dim HeadCd As Integer
Dim LedCd As Integer
With VsfHelp
    .Cols = 10
    .Rows = 1
    .TextMatrix(0, 0) = "SR."
    .ColWidth(0) = 400
    .TextMatrix(0, 1) = "PARTICULARS"
    .ColWidth(1) = 4500
    .TextMatrix(0, 2) = "SIDE"
    .ColWidth(2) = 400
    .TextMatrix(0, 3) = "OPENING BALANCE"
    .ColWidth(3) = 2000
    .ColFormat(3) = "0.00"
    .ColAlignment(3) = flexAlignRightCenter
    .TextMatrix(0, 4) = "DEBIT"
    .ColWidth(4) = 1500
    .ColFormat(4) = "0.00"
    .ColAlignment(4) = flexAlignRightCenter
    .TextMatrix(0, 5) = "CREDIT"
    .ColWidth(5) = 1500
    .ColFormat(5) = "0.00"
    .ColAlignment(5) = flexAlignRightCenter
    .TextMatrix(0, 6) = "CLOSING BALANCE"
    .ColWidth(6) = 2000
    .ColFormat(6) = "0.00"
    .ColAlignment(6) = flexAlignRightCenter
    .ColWidth(7) = 0    'HCode
    .ColWidth(8) = 0    'LCode
    .ColWidth(9) = 0    'ECode
    .Refresh
    .Rows = 2
    .Row = 1
    .TextMatrix(.Row, 1) = "BALANCE SHEET"
    RowInc
    RsQ.Open "Select * From TmpBSPrn Where SrN=1 And LAmt<>14", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        .TextMatrix(.Row, 0) = .Row
        .TextMatrix(.Row, 1) = RsQ.Fields("LName")
        .TextMatrix(.Row, 2) = RsQ.Fields("RName")
        .TextMatrix(.Row, 7) = RsQ.Fields("LAmt")
        HeadCd = RsQ.Fields("LAmt")
        Set RsQL = Nothing
        RsQL.Open "Select * From TmpPLPrn Where SrN=" & HeadCd & " And LAmt<>5", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        If RsQL.EOF = False Then RowInc
        Do While RsQL.EOF = False
            .TextMatrix(.Row, 0) = .Row
            .TextMatrix(.Row, 1) = RsQL.Fields("LName")
            .TextMatrix(.Row, 8) = RsQL.Fields("LAmt")
            RsQL.MoveNext
            If RsQL.EOF = False Then RowInc
        Loop
        RsQ.MoveNext
        If RsQ.EOF = False Then RowInc
    Loop
    RowInc
    RowInc
    .TextMatrix(.Row, 1) = "INCOME & EXPENDITURE"
    RowInc
    Set RsQ = Nothing
    RsQ.Open "Select * From TmpBSPrn Where SrN=0", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        .TextMatrix(.Row, 0) = .Row
        .TextMatrix(.Row, 1) = RsQ.Fields("LName")
        .TextMatrix(.Row, 2) = RsQ.Fields("RName")
        .TextMatrix(.Row, 7) = RsQ.Fields("LAmt")
        HeadCd = RsQ.Fields("LAmt")
        Set RsQL = Nothing
        RsQL.Open "Select * From TmpPLPrn Where SrN=" & HeadCd & " And LAmt Not In (107,108)", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        If RsQL.EOF = False Then RowInc
        Do While RsQL.EOF = False
            .TextMatrix(.Row, 0) = .Row
            .TextMatrix(.Row, 1) = RsQL.Fields("LName")
            .TextMatrix(.Row, 8) = RsQL.Fields("LAmt")
            LedCd = RsQL.Fields("LAmt")
            Set RsQE = Nothing
            RsQE.Open "Select * From TmpRecPrn Where SrN=" & LedCd, DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
            If RsQE.EOF = False Then RowInc
            Do While RsQE.EOF = False
                .TextMatrix(.Row, 0) = .Row
                .TextMatrix(.Row, 1) = RsQE.Fields("LName")
                .TextMatrix(.Row, 9) = RsQE.Fields("LAmt")
                RsQE.MoveNext
                If RsQE.EOF = False Then RowInc
            Loop
            RsQL.MoveNext
            If RsQL.EOF = False Then RowInc
        Loop
        RsQ.MoveNext
        If RsQ.EOF = False Then RowInc
    Loop
    RowInc
    RowInc
    .TextMatrix(.Row, 1) = "CONTRA ENTRIES"
    RowInc
    Set RsQ = Nothing
    Set RsQL = Nothing
    RsQL.Open "Select * From TmpPLPrn Where (SrN=14 Or LAmt In (107,108))", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQL.EOF = False Then RowInc
    Do While RsQL.EOF = False
        .TextMatrix(.Row, 0) = .Row
        .TextMatrix(.Row, 1) = RsQL.Fields("LName")
        .TextMatrix(.Row, 8) = RsQL.Fields("LAmt")
        LedCd = RsQL.Fields("LAmt")
        Set RsQE = Nothing
        RsQE.Open "Select * From TmpRecPrn Where SrN=" & LedCd, DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        If RsQE.EOF = False Then RowInc
        Do While RsQE.EOF = False
            .TextMatrix(.Row, 0) = .Row
            .TextMatrix(.Row, 1) = RsQE.Fields("LName")
            .TextMatrix(.Row, 9) = RsQE.Fields("LAmt")
            RsQE.MoveNext
            If RsQE.EOF = False Then RowInc
        Loop
        RsQL.MoveNext
        If RsQL.EOF = False Then RowInc
    Loop
    .Refresh
End With
End Sub

Private Sub VsfHelp_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    SetFinalTot
End Sub

Private Sub VsfHelp_EnterCell()
If VsfHelp.Col = 3 Or VsfHelp.Col = 4 Then
    VsfHelp.Editable = flexEDKbd
    VsfHelp.AutoSearch = flexSearchNone
    On Error Resume Next
    SendKeys "{F2}"
Else
    VsfHelp.Editable = flexEDNone
    VsfHelp.AutoSearch = flexSearchFromCursor
End If
End Sub

Private Sub SetFinalTot()
Dim i As Double
Dim Profit As Double
Profit = SetProfit(mAcCode)
TxtOTotal.Text = 0
TxtATotal.Text = 0
TxtWTotal.Text = 0
TxtCTotal.Text = 0
With VsfHelp
    For i = 1 To .Rows - 1
        TxtOTotal.Text = Val(TxtOTotal.Text) + Val(.TextMatrix(i, 3))
        TxtATotal.Text = Val(TxtATotal.Text) + Val(.TextMatrix(i, 4))
        TxtWTotal.Text = Val(TxtWTotal.Text) + Val(.TextMatrix(i, 5))
        TxtCTotal.Text = Val(TxtCTotal.Text) + Val(.TextMatrix(i, 6))
    Next
    If Profit > 0 Then TxtWTotal.Text = Val(TxtWTotal.Text) - Profit Else TxtATotal.Text = Val(TxtATotal.Text) - Abs(Profit)
    .Refresh
    TxtOTotal.Text = Format(TxtOTotal.Text, "0.00")
    TxtATotal.Text = Format(TxtATotal.Text, "0.00")
    TxtWTotal.Text = Format(TxtWTotal.Text, "0.00")
    TxtCTotal.Text = Format(TxtCTotal.Text, "0.00")
End With
End Sub
Private Sub RowInc()
With VsfHelp
    .Rows = .Rows + 1
    .Row = .Rows - 1
End With
End Sub

Private Sub SetParent()
Dim RsQ As New ADODB.Recordset
RsQ.Open "Select * From QGroup Where SACode=" & mAcCode & " Order By AcCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
mAcList = CStr(mAcCode)
Do While RsQ.EOF = False
    mAcList = mAcList + "," + CStr(RsQ.Fields("AcCode"))
    RsQ.MoveNext
Loop
End Sub
Private Sub ShowData()
Dim RsQ As New ADODB.Recordset
Dim i As Double
With VsfHelp
    Set RsQ = Nothing
    RsQ.Open "Select * From TmpTrialBal Where HType=1 And HCode Not In (12,14)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        .Row = .FindRow(RsQ.Fields("LCode"), 1, 8)
        If .Row >= 1 Then
            .TextMatrix(.Row, 3) = RsQ.Fields("OpDr") - RsQ.Fields("OpCr")
            .TextMatrix(.Row, 4) = RsQ.Fields("ADr")
            .TextMatrix(.Row, 5) = RsQ.Fields("ACr")
            .TextMatrix(.Row, 6) = RsQ.Fields("DBal") - RsQ.Fields("CBal")
        End If
        RsQ.MoveNext
        If Val(.TextMatrix(.Row, 8)) = 1 Or Val(.TextMatrix(.Row, 8)) = 88 Then
            i = SetProfit(mAcCode)
            If i > 0 Then .TextMatrix(.Row, 5) = Val(.TextMatrix(.Row, 5)) + i Else .TextMatrix(.Row, 4) = Val(.TextMatrix(.Row, 4)) + Abs(i)
            .TextMatrix(.Row, 6) = Val(.TextMatrix(.Row, 3)) + Val(.TextMatrix(.Row, 4)) - Val(.TextMatrix(.Row, 5))
        End If
    Loop
    Set RsQ = Nothing
    RsQ.Open "Select * From TmpTrialBal Where HType=1 And HCode = 12", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        .Row = .FindRow(RsQ.Fields("LCode"), 1, 8)
        If .Row >= 1 Then
            .TextMatrix(.Row, 3) = RsQ.Fields("OpDr") - RsQ.Fields("OpCr")
            .TextMatrix(.Row, 6) = RsQ.Fields("DBal") - RsQ.Fields("CBal")
            If Val(.TextMatrix(.Row, 6)) > Val(.TextMatrix(.Row, 3)) Then
                .TextMatrix(.Row, 4) = Val(.TextMatrix(.Row, 6)) - Val(.TextMatrix(.Row, 3))
            Else
                .TextMatrix(.Row, 5) = Val(.TextMatrix(.Row, 3)) - Val(.TextMatrix(.Row, 6))
            End If
        End If
        RsQ.MoveNext
    Loop
    Set RsQ = Nothing
    RsQ.Open "Select * From TmpCtDtl Where HType=0 And LCode Not In (107,108)", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        .Row = .FindRow(RsQ.Fields("ECode"), 1, 9)
        If .Row >= 1 Then
            If RsQ.Fields("Side") = "D" Then .TextMatrix(.Row, 4) = RsQ.Fields("Amt") Else .TextMatrix(.Row, 5) = RsQ.Fields("Amt")
        End If
        RsQ.MoveNext
    Loop
    Set RsQ = Nothing
    RsQ.Open "Select * From TmpTrialBal Where HType=1 And HCode=14", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        .Row = .FindRow(RsQ.Fields("LCode"), 1, 8)
        If .Row >= 1 Then
            .TextMatrix(.Row, 3) = RsQ.Fields("OpDr") - RsQ.Fields("OpCr")
            .TextMatrix(.Row, 6) = RsQ.Fields("DBal") - RsQ.Fields("CBal")
        End If
        RsQ.MoveNext
    Loop
    Set RsQ = Nothing
    RsQ.Open "Select * From TmpCtDtl Where (HCode=14 Or LCode In (107,108))", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        .Row = .FindRow(RsQ.Fields("ECode"), 1, 9)
        If .Row >= 1 Then
            If RsQ.Fields("Side") = "D" Then .TextMatrix(.Row, 4) = RsQ.Fields("Amt") Else .TextMatrix(.Row, 5) = RsQ.Fields("Amt")
        End If
        RsQ.MoveNext
    Loop
    .Refresh
    SetFinalTot
End With
End Sub
Private Sub VsfHelp_RowColChange()
If VsfHelp.Col >= 3 Then
    VsfHelp.Editable = flexEDKbd
    VsfHelp.AutoSearch = flexSearchNone
    On Error Resume Next
    SendKeys "{F2}"
Else
    VsfHelp.Editable = flexEDNone
    VsfHelp.AutoSearch = flexSearchFromCursor
End If
End Sub
Private Sub ExpMainData()
Dim i As Double
With VsfMainExport
    .Cols = 7
    .Rows = 1
    .TextMatrix(0, 0) = "SR."
    .ColWidth(0) = 400
    .TextMatrix(0, 1) = "PARTICULARS"
    .ColWidth(1) = 4500
    .TextMatrix(0, 2) = "SIDE"
    .ColWidth(2) = 400
    .TextMatrix(0, 3) = "OPENING BALANCE"
    .ColWidth(3) = 2000
    .ColFormat(3) = "0.00"
    .ColAlignment(3) = flexAlignRightCenter
    .TextMatrix(0, 4) = "DEBIT"
    .ColWidth(4) = 1500
    .ColFormat(4) = "0.00"
    .ColAlignment(4) = flexAlignRightCenter
    .TextMatrix(0, 5) = "CREDIT"
    .ColWidth(5) = 1500
    .ColFormat(5) = "0.00"
    .ColAlignment(5) = flexAlignRightCenter
    .TextMatrix(0, 6) = "CLOSING BALANCE"
    .ColWidth(6) = 2000
    .ColFormat(6) = "0.00"
    .ColAlignment(6) = flexAlignRightCenter
    .Refresh
    .Rows = .Rows + 1
    .Row = 1
    For i = 1 To VsfHelp.Rows - 1
        .TextMatrix(i, 0) = VsfHelp.TextMatrix(i, 0)
        .TextMatrix(i, 1) = VsfHelp.TextMatrix(i, 1)
        .TextMatrix(i, 2) = VsfHelp.TextMatrix(i, 2)
        .TextMatrix(i, 3) = VsfHelp.TextMatrix(i, 3)
        .TextMatrix(i, 4) = VsfHelp.TextMatrix(i, 4)
        .TextMatrix(i, 5) = VsfHelp.TextMatrix(i, 5)
        .TextMatrix(i, 6) = VsfHelp.TextMatrix(i, 6)
        .Rows = .Rows + 1
    Next
    .Rows = .Rows + 1
    .Refresh
    i = .Rows - 1
    Do While i > 0
        .TextMatrix(i, 0) = .TextMatrix(i - 1, 0)
        .TextMatrix(i, 1) = .TextMatrix(i - 1, 1)
        .TextMatrix(i, 2) = .TextMatrix(i - 1, 2)
        .TextMatrix(i, 3) = .TextMatrix(i - 1, 3)
        .TextMatrix(i, 4) = .TextMatrix(i - 1, 4)
        .TextMatrix(i, 5) = .TextMatrix(i - 1, 5)
        .TextMatrix(i, 6) = .TextMatrix(i - 1, 6)
        i = i - 1
    Loop
    .Cell(flexcpText, 0, 0, 0, .Cols - 1) = ""
    .TextMatrix(0, 0) = LsvClient.TextMatrix(LsvClient.Row, 0)
    .TextMatrix(0, 1) = LsvClient.TextMatrix(LsvClient.Row, 1)
End With
VsfMainExport.SaveGrid Environ("USERPROFILE") & "\Desktop\TRIALBAL.XLS", flexFileTabText
MsgBox "Successfully Excel File Generated In " + Environ("USERPROFILE") + "\Desktop\TRIALBAL.XLS", vbInformation, "Alert"
End Sub
