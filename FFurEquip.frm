VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Begin VB.Form FrmFurEquip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Furniture And Equipment"
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
   Icon            =   "FFurEquip.frx":0000
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
      Left            =   4900
      Locked          =   -1  'True
      TabIndex        =   18
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
      Left            =   6900
      Locked          =   -1  'True
      TabIndex        =   17
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
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   16
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
         FormatString    =   $"FFurEquip.frx":0442
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
            FormatString    =   $"FFurEquip.frx":048B
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
         Left            =   9900
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
         Picture         =   "FFurEquip.frx":04D4
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
         FormatString    =   $"FFurEquip.frx":0B0A
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
         FormatString    =   $"FFurEquip.frx":0B53
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
Attribute VB_Name = "FrmFurEquip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DBWorkTmp As New ADODB.Connection
Dim mAcCode As Double
Dim mAcList As String
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
RsQry.Open "Select AcMst.AcName,AcMst.FileNo,AcMst.City,GrpMst.GName,AcMst.AcCode From AcMst,GrpMst Where AcMst.Active=-1 And " & _
"AcMst.AcType=3 And AcMst.AcType=GrpMst.GCode Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Set LsvClient.DataSource = RsQry
With LsvClient
    .TextMatrix(0, 0) = "NAME"
    .ColWidth(0) = 4800
    .TextMatrix(0, 1) = "FILE NO."
    .ColWidth(1) = 1200
    .TextMatrix(0, 2) = "CITY"
    .ColWidth(2) = 1200
    .TextMatrix(0, 3) = "TYPE"
    .ColWidth(3) = 600
    .ColWidth(4) = 0    'ACCODE
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
    mAcCode = Val(LsvClient.TextMatrix(LsvClient.Row, 4))
    FraClientHelp.Left = 18000
    SetParent
    Dim RsQ As New ADODB.Recordset
    DBWorkTmp.BeginTrans
    DBWorkTmp.Execute "Delete From TmpTrialBal"
    DBWorkTmp.Execute "Delete From TmpCtDtl"
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
    DBWorkTmp.CommitTrans
    SetData
    ShowData
    CheckState
    VsfHelp.SetFocus
    DbDataDB.BeginTrans
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'SCHOOL_SFE','VIEW_DATA','" & Date & "','" & Time & "')"
    DbDataDB.CommitTrans
End If
End Sub
Private Sub Tlbsav_Click(Index As Integer)
Select Case TlbSav(Index).ToolTipText
    Case "Delete"
        If MsgBox("Sure To Delete ?", vbInformation + vbYesNo) = vbYes Then
            DbDataDB.BeginTrans
            DbDataDB.Execute "Delete From JvDtl Where AcCode=" & mAcCode & " And LCode in (Select LCode from LedMst Where HCode=56)"
'            DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'SCHOOL_SFE','DELETE','" & Date & "','" & Time & "')"
            DbDataDB.CommitTrans
        End If
        SetData
    Case "Save"
        If TxtName.Text = "" Then
            MsgBox "Sorry! Not Allowed.", vbInformation, "Black Data Error"
            TxtName.SetFocus
        Else
            If MsgBox("Are you sure to save?", vbInformation + vbYesNo, "Confirmation") = vbYes Then
                SaveData
                DbDataDB.BeginTrans
'                    DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'SCHOOL_SFE','UPDATE','" & Date & "','" & Time & "')"
                DbDataDB.CommitTrans
                ClearText
                SetTool True
                CmdSearch.SetFocus
            End If
        End If
    Case "Cancel"
        SetData
        VsfHelp.SetFocus
    Case "Exit"
        If TxtName.Text <> "" Then
            If TlbSav(0).Enabled = True Then
                If MsgBox("Save data before exit?", vbInformation + vbYesNo, "Confirmation") = vbYes Then SaveData
                DbDataDB.BeginTrans
'                    DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'SCHOOL_SFE','UPDATE','" & Date & "','" & Time & "')"
                DbDataDB.CommitTrans
            End If
        End If
        DBWorkTmp.Close
        Unload Me
End Select
End Sub
Private Function SetTool(ByVal mVal As Boolean)
    TlbSav(2).Enabled = True
End Function
Private Sub SetData()
Dim RsQ As New ADODB.Recordset
With VsfHelp
    .Cols = 7
    .Rows = 1
    .TextMatrix(0, 0) = "SR."
    .ColWidth(0) = 400
    .TextMatrix(0, 1) = "HEAD"
    .ColWidth(1) = 4500
    .TextMatrix(0, 2) = "OPENING BALANCE"
    .ColWidth(2) = 2000
    .ColFormat(2) = "0.00"
    .ColAlignment(2) = flexAlignRightCenter
    .TextMatrix(0, 3) = "ADDITION"
    .ColWidth(3) = 1500
    .ColFormat(3) = "0.00"
    .ColAlignment(3) = flexAlignRightCenter
    .TextMatrix(0, 4) = "WRITTEN OFF"
    .ColWidth(4) = 1500
    .ColFormat(4) = "0.00"
    .ColAlignment(4) = flexAlignRightCenter
    .TextMatrix(0, 5) = "CLOSING BALANCE"
    .ColWidth(5) = 2000
    .ColFormat(5) = "0.00"
    .ColAlignment(5) = flexAlignRightCenter
    .ColWidth(6) = 0    'LCode
    .Refresh
    .Rows = 2
    .Row = 1
    RsQ.Open "Select * From LedMst Where LCode Between  91 And 103 Order By LCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        .TextMatrix(.Row, 0) = .Row
        .TextMatrix(.Row, 1) = RsQ.Fields("LName")
        .TextMatrix(.Row, 6) = RsQ.Fields("LCode")
        RsQ.MoveNext
        If RsQ.EOF = False Then RowInc
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
TxtATotal.Text = 0
TxtWTotal.Text = 0
TxtCTotal.Text = 0
With VsfHelp
    For i = 1 To .Rows - 1
        If (Val(.TextMatrix(i, 2)) + Val(.TextMatrix(i, 3)) - Val(.TextMatrix(i, 4))) < 0 Then
            MsgBox "Closing Balance cannot be Negative.", vbCritical, "Alert"
            VsfHelp.SetFocus
        Else
            .TextMatrix(i, 5) = Val(.TextMatrix(i, 2)) + Val(.TextMatrix(i, 3)) - Val(.TextMatrix(i, 4))
            TxtATotal.Text = Val(TxtATotal.Text) + Val(.TextMatrix(i, 3))
            TxtWTotal.Text = Val(TxtWTotal.Text) + Val(.TextMatrix(i, 4))
            TxtCTotal.Text = Val(TxtCTotal.Text) + Val(.TextMatrix(i, 5))
            If Val(.TextMatrix(i, 5)) = 0 Then .TextMatrix(i, 5) = ""
        End If
    Next
    .Refresh
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
    RsQ.Open "Select * From OpDtl Where AcCode=" & mAcCode & " And LCode Between 91 And 103 Order By LCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    TxtOTotal.Text = 0
    Do While RsQ.EOF = False
        .Row = .FindRow(RsQ.Fields("LCode"), 1, 6)
        If .Row >= 1 Then .TextMatrix(.Row, 2) = RsQ.Fields("OpBal")
        TxtOTotal.Text = Val(TxtOTotal.Text) + Val(.TextMatrix(.Row, 2))
        RsQ.MoveNext
    Loop
    Set RsQ = Nothing
    RsQ.Open "Select * From ITDtl Where AcCode=" & mAcCode & " And LCode Between 91 And 103 And Side='D' Order By LCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        .Row = .FindRow(RsQ.Fields("LCode"), 1, 6)
        If .Row >= 1 Then .TextMatrix(.Row, 3) = RsQ.Fields("Amt")
        RsQ.MoveNext
    Loop
    Set RsQ = Nothing
    RsQ.Open "Select * From ITDtl Where AcCode=" & mAcCode & " And LCode Between 91 And 103 And Side='C' Order By LCode", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    Do While RsQ.EOF = False
        .Row = .FindRow(RsQ.Fields("LCode"), 1, 6)
        If .Row >= 1 Then .TextMatrix(.Row, 4) = RsQ.Fields("Amt")
        RsQ.MoveNext
    Loop
    .Refresh
    SetFinalTot
End With
End Sub
Private Sub SaveData()
Dim i As Double
Dim RsQ As New ADODB.Recordset
On Error GoTo XErr
RsQ.Open "Select Sum(Amt) As DTotal from QctDtl where AcCode=" & mAcCode & " and Side='D' and HCode=8", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If Val(TxtATotal.Text) <> RsQ.Fields("DTotal") Then
        MsgBox "Total Addition does not tally with Addition as per transactions.", vbCritical, "Alert"
        Exit Sub
    End If
Set RsQ = Nothing
RsQ.Open "Select Sum(Amt) As CTotal from QctDtl where AcCode=" & mAcCode & " and Side='C' and HCode=8", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If Val(TxtWTotal.Text) <> RsQ.Fields("CTotal") Then
        MsgBox "Total Write Off/Sale does not tally with Deletion as per transactions.", vbCritical, "Alert"
        Exit Sub
    End If
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From ItDtl Where AcCode=" & mAcCode
With VsfHelp
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 3)) <> 0 Then DbDataDB.Execute "Insert InTo ItDtl (AcCode,SrN,LCode,Side,Amt) Values (" & mAcCode & "," & i & "," & Val(.TextMatrix(i, 6)) & ",'D','" & Val(.TextMatrix(i, 3)) & "')"
        If Val(.TextMatrix(i, 4)) <> 0 Then DbDataDB.Execute "Insert InTo ItDtl (AcCode,SrN,LCode,Side,Amt) Values (" & mAcCode & "," & i & "," & Val(.TextMatrix(i, 6)) & ",'C','" & Val(.TextMatrix(i, 4)) & "')"
    Next
End With
DbDataDB.CommitTrans
MsgBox "Record Saved.", vbInformation, "Alert"
Exit Sub
XErr:
MsgBox Err.Description
DbDataDB.RollbackTrans
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
Private Sub CheckState()
Dim RsQ As New ADODB.Recordset
Dim mString As String
RsQ.Open "Select * From QGroup Where AcCode=" & mAcCode, DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
mString = CStr(RsQ.Fields("AcCode")) + "," + CStr(RsQ.Fields("PaCode")) + "," + CStr(RsQ.Fields("SaCode"))
Set RsQ = Nothing
RsQ.Open "Select * From RepDtl Where AcCode In (" & mString & ")", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
Do While RsQ.EOF = False
    If RsQ.Fields("RUDIN") <> "" Then
        MsgBox "UDIN generated. Data can not be edited.", vbCritical, "Alert"
        TlbSav(0).Enabled = False
        Exit Sub
    End If
    RsQ.MoveNext
Loop
End Sub
