VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "vsflex7.ocx"
Begin VB.Form FrmSchedule9C 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schedule 9C"
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
   Icon            =   "FSch9CEn.frx":0000
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
         FormatString    =   $"FSch9CEn.frx":0442
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
            FormatString    =   $"FSch9CEn.frx":048B
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
         Left            =   9520
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   8280
         Width           =   1500
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
         Picture         =   "FSch9CEn.frx":04D4
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
         FormatString    =   $"FSch9CEn.frx":0B0A
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
         Left            =   11400
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
         FormatString    =   $"FSch9CEn.frx":0B53
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
         Height          =   8655
         Index           =   0
         Left            =   0
         Top             =   120
         Width           =   12975
      End
   End
End
Attribute VB_Name = "FrmSchedule9C"
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
RsQry.Open "Select AcMst.AcName,AcMst.FileNo,AcMst.City,'' As ParentNm,GrpMst.GName,AcMst.AcCode From AcMst,GrpMst Where AcMst.Active=-1 And AcMst.PaCode=0 And AcMst.AcType=1" & _
" And AcMst.AcType=GrpMst.GCode Order By AcMst.FileNo", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
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
    Set DBWorkTmp = Nothing
    Unload Me
End Sub

Private Sub LsvClient_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxtName.Text = LsvClient.TextMatrix(LsvClient.Row, 0)
    mAcCode = Val(LsvClient.TextMatrix(LsvClient.Row, 5))
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
    DbDataDB.BeginTrans
'        DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'SCHED9C','VIEW','" & Date & "','" & Time & "')"
    DbDataDB.CommitTrans
    VsfHelp.SetFocus
End If
End Sub
Private Sub Tlbsav_Click(Index As Integer)
Select Case TlbSav(Index).ToolTipText
    Case "Delete"
        If MsgBox("Sure To Delete ?", vbInformation + vbYesNo) = vbYes Then
            DbDataDB.BeginTrans
            DbDataDB.Execute "Delete From 9CDtl Where AcCode=" & mAcCode
'            DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'SCHED9C','DELETE','" & Date & "','" & Time & "')"
            DbDataDB.CommitTrans
        End If
        SetData
    Case "Save"
        SaveRec
        DbDataDB.BeginTrans
'            DbDataDB.Execute "Insert Into UsrLog (UPass,FYear,AcCode,Form,Activity,ActDate,ActTime) Values ('" & mPassword & "','" & mFinYear & "'," & mAcCode & ",'SCHED9C','SAVE_DATA','" & Date & "','" & Time & "')"
        DbDataDB.CommitTrans
    Case "Cancel"
        SetData
        VsfHelp.SetFocus
    Case "Exit"
        Unload Me
End Select
End Sub
Private Function SetTool(ByVal mVal As Boolean)
    TlbSav(2).Enabled = True
End Function
Private Sub SetData()
With VsfHelp
    .Cols = 4
    .Rows = 1
    .TextMatrix(0, 0) = "SR."
    .ColWidth(0) = 400
    .TextMatrix(0, 1) = "SCHEDULE 9C ITEM"
    .ColWidth(1) = 7500
    .TextMatrix(0, 2) = "SUB TOTAL"
    .ColWidth(2) = 1500
    .ColFormat(2) = "0.00"
    .ColAlignment(2) = flexAlignRightCenter
    .TextMatrix(0, 3) = "FINAL TOTAL"
    .ColWidth(3) = 1500
    .ColFormat(3) = "0.00"
    .Refresh
    .Rows = 2
    .Row = 1
    .TextMatrix(.Row, 0) = "1"
    .TextMatrix(.Row, 1) = "Gross Annual Income"
    RowInc
    .TextMatrix(.Row, 1) = "Details of income not chargeable to contribution under Section 58 and Rule 32:"
    RowInc
    .TextMatrix(.Row, 1) = "(i) Donations received during the year from any source"
    RowInc
    .TextMatrix(.Row, 1) = "     (a) Corpus"
    RowInc
    .TextMatrix(.Row, 0) = "2"
    .TextMatrix(.Row, 1) = "          (1) From Country"
    RowInc
    .TextMatrix(.Row, 0) = "3"
    .TextMatrix(.Row, 1) = "          (2) From Foreign Country (FC)"
    RowInc
    .TextMatrix(.Row, 1) = "     (b) General"
    RowInc
    .TextMatrix(.Row, 0) = "4"
    .TextMatrix(.Row, 1) = "          (1) From Country"
    RowInc
    .TextMatrix(.Row, 0) = "5"
    .TextMatrix(.Row, 1) = "          (2) From Foreign Country (FC)"
    RowInc
    .TextMatrix(.Row, 1) = "(ii) Grants by Government and Local Authorities"
    RowInc
    .TextMatrix(.Row, 0) = "6"
    .TextMatrix(.Row, 1) = "     (a) Government and Local Authorities"
    RowInc
    .TextMatrix(.Row, 0) = "7"
    .TextMatrix(.Row, 1) = "     (b) From Foreign Country"
    RowInc
    .TextMatrix(.Row, 1) = "     (c) By Funding Agencies"
    RowInc
    .TextMatrix(.Row, 0) = "8"
    .TextMatrix(.Row, 1) = "          (1) From Country"
    RowInc
    .TextMatrix(.Row, 0) = "9"
    .TextMatrix(.Row, 1) = "          (2) From Foreign Country (FC)"
    RowInc
    .TextMatrix(.Row, 0) = "10"
    .TextMatrix(.Row, 1) = "(iii) Amount spent for the purpose of Education"
    RowInc
    .TextMatrix(.Row, 0) = "11"
    .TextMatrix(.Row, 1) = "(iv) Amount spent for the purpose of Medical Relief"
    RowInc
    .TextMatrix(.Row, 1) = "(v) (A) Deductions out of incomes from lands used for Agricultural Purpose"
    RowInc
    .TextMatrix(.Row, 0) = "12"
    .TextMatrix(.Row, 1) = "          (a) Land Revenue and Local Fund Cess"
    RowInc
    .TextMatrix(.Row, 0) = "13"
    .TextMatrix(.Row, 1) = "          (b) Rent payable to Superior Landlord"
    RowInc
    .TextMatrix(.Row, 0) = "14"
    .TextMatrix(.Row, 1) = "          (c) Cost of Production, if lands are cultivated by the Trust"
    RowInc
    .TextMatrix(.Row, 0) = "15"
    .TextMatrix(.Row, 1) = "     (B) Income from lands used for Agricultural Purpose"
    RowInc
    .TextMatrix(.Row, 1) = "(vi) (A) Deductions out of income from lands used for Non-agricultural Purpose"
    RowInc
    .TextMatrix(.Row, 0) = "16"
    .TextMatrix(.Row, 1) = "          (a) Assessment, Cesses and other Municipal Taxes"
    RowInc
    .TextMatrix(.Row, 0) = "17"
    .TextMatrix(.Row, 1) = "          (b) Ground Rent payable to Superior Landlord"
    RowInc
    .TextMatrix(.Row, 0) = "18"
    .TextMatrix(.Row, 1) = "          (c) Insurance Premium"
    RowInc
    .TextMatrix(.Row, 0) = "19"
    .TextMatrix(.Row, 1) = "          (d) Repairs at 8.33% of gross rent of building"
    RowInc
    .TextMatrix(.Row, 0) = "20"
    .TextMatrix(.Row, 1) = "          (e) Cost of collection at 4% of gross rent of buildings let-out"
    RowInc
    .TextMatrix(.Row, 0) = "21"
    .TextMatrix(.Row, 1) = "     (B) Income from lands used for Non-agricultural Purpose"
    RowInc
    .TextMatrix(.Row, 0) = "22"
    .TextMatrix(.Row, 1) = "(vii) Cost of collection of income or receipts from securities, stocks, etc. at 1% of such income"
    RowInc
    .TextMatrix(.Row, 0) = "23"
    .TextMatrix(.Row, 1) = "(viii) Deductions on account of repairs in respect of buildings not rented and not yielding income at 8.33% of the estimated gross annual rent"
    RowInc
    .TextMatrix(.Row, 0) = "24"
    .TextMatrix(.Row, 1) = "Income Liable to Contribution"
End With
End Sub

Private Sub VsfHelp_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    SetFinalTot
End Sub

Private Sub VsfHelp_EnterCell()
If VsfHelp.Col >= 2 Then
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
With VsfHelp
    If VsfHelp.Rows < 31 Then Exit Sub
    .Row = 6
    .TextMatrix(.Row, 3) = Val(.TextMatrix(.Row - 1, 2)) + Val(.TextMatrix(.Row, 2))
    .Row = 9
    .TextMatrix(.Row, 3) = Val(.TextMatrix(.Row - 1, 2)) + Val(.TextMatrix(.Row, 2))
    .Row = 15
    .TextMatrix(.Row, 3) = Val(.TextMatrix(.Row - 1, 2)) + Val(.TextMatrix(.Row, 2))
    .Row = 22
    If Val(.TextMatrix(19, 2)) + Val(.TextMatrix(20, 2)) + Val(.TextMatrix(21, 2)) <> 0 And Val(.TextMatrix(22, 2)) <> 0 Then
        If Val(.TextMatrix(19, 2)) + Val(.TextMatrix(20, 2)) + Val(.TextMatrix(21, 2)) < Val(.TextMatrix(22, 2)) Then
            .TextMatrix(.Row - 1, 3) = Val(.TextMatrix(19, 2)) + Val(.TextMatrix(20, 2)) + Val(.TextMatrix(21, 2))
        Else
            .TextMatrix(.Row - 1, 3) = Val(.TextMatrix(22, 2))
        End If
    Else
        .Cell(flexcpText, 19, 2, 22, 2) = "0.00"
    End If
    .Row = 29
    If Val(.TextMatrix(24, 2)) + Val(.TextMatrix(25, 2)) + Val(.TextMatrix(26, 2)) + Val(.TextMatrix(27, 2)) + Val(.TextMatrix(28, 2)) <> 0 And Val(.TextMatrix(29, 2)) <> 0 Then
        If Val(.TextMatrix(24, 2)) + Val(.TextMatrix(25, 2)) + Val(.TextMatrix(26, 2)) + Val(.TextMatrix(27, 2)) + Val(.TextMatrix(28, 2)) < Val(.TextMatrix(29, 2)) Then
            .TextMatrix(.Row - 1, 3) = Val(.TextMatrix(24, 2)) + Val(.TextMatrix(25, 2)) + Val(.TextMatrix(26, 2)) + Val(.TextMatrix(27, 2)) + Val(.TextMatrix(28, 2))
        Else
            .TextMatrix(.Row - 1, 3) = Val(.TextMatrix(29, 2))
        End If
    Else
        .Cell(flexcpText, 24, 2, 29, 2) = "0.00"
    End If
    .Row = 30
    TxtCTotal.Text = .TextMatrix(1, 3)
    For i = 2 To .Rows - 1
        TxtCTotal.Text = Val(TxtCTotal.Text) - Val(.TextMatrix(i, 3))
        If Val(.TextMatrix(i, 2)) = 0 Then .TextMatrix(i, 2) = ""
        If Val(.TextMatrix(i, 3)) = 0 Then .TextMatrix(i, 3) = ""
    Next
    If Val(TxtCTotal.Text) < 0 Then TxtCTotal.Text = "0.00"
    .Refresh
    TxtCTotal.Text = Format(TxtCTotal.Text, "0.00")
End With
End Sub
Private Sub RowInc()
With VsfHelp
    .Rows = .Rows + 1
    .Row = .Rows - 1
End With
End Sub

Private Sub SaveRec()
    SaveData
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
Private Function SetGrossInc() As Double
Dim RsQry As New ADODB.Recordset
Dim mTotal As Double
RsQry.Open "Select Sum(CBal) As TotRs From TmpTrialBal Where HType=0 And HCode<>54 And LCode<>108", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
mTotal = IIf(IsNull(RsQry.Fields("TotRs")) = True, 0, RsQry.Fields("TotRs"))
SetGrossInc = mTotal
End Function
Private Sub ShowData()
Dim RsQ As New ADODB.Recordset
Dim i As Double
Dim Inc As Double
With VsfHelp
    RsQ.Open "Select * From 9CDtl Where AcCode=" & mAcCode & " Order By SrN", DbDataDB, adOpenDynamic, adLockReadOnly, adCmdText
    If RsQ.EOF = False Then
        .Row = 1
        Do While RsQ.EOF = False
            .TextMatrix(.Row, 2) = RsQ.Fields("Amt1")
            .TextMatrix(.Row, 3) = RsQ.Fields("Amt2")
            RsQ.MoveNext
            If RsQ.EOF = False Then
                If .Row + 1 < .Rows Then .Row = .Row + 1
            End If
        Loop
        TxtCTotal.Text = Val(.TextMatrix(1, 3))
        For i = 2 To .Rows - 1
            TxtCTotal.Text = Val(TxtCTotal.Text) - Val(.TextMatrix(i, 3))
            If Val(.TextMatrix(i, 2)) = 0 Then .TextMatrix(i, 2) = ""
            If Val(.TextMatrix(i, 3)) = 0 Then .TextMatrix(i, 3) = ""
        Next
        If Val(TxtCTotal.Text) < 0 Then TxtCTotal.Text = "0.00"
     Else
        .Row = 1
        .TextMatrix(.Row, 3) = SetGrossInc
        Inc = Val(.TextMatrix(.Row, 3))
        Set RsQ = Nothing
        RsQ.Open "Select IIf(IsNull(Sum(ACr))=True,0,Sum(ACr)) As TotRs From TmpTrialBal Where HCode=42", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        If RsQ.EOF = False Then
            Inc = Inc - RsQ.Fields("TotRs")
        End If
        .Row = 5
        Set RsQ = Nothing
        RsQ.Open "Select IIf(IsNull(Sum(ACr))=True,0,Sum(ACr)) As TotRs From TmpTrialBal Where HCode=1", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        If RsQ.EOF = False Then
            .TextMatrix(.Row, 2) = RsQ.Fields("TotRs")
            .TextMatrix(1, 3) = Val(.TextMatrix(1, 3)) + RsQ.Fields("TotRs")
        End If
        .Row = 6
        Set RsQ = Nothing
        RsQ.Open "Select IIf(IsNull(Sum(ACr))=True,0,Sum(ACr)) As TotRs From TmpTrialBal Where HCode=13", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        If RsQ.EOF = False Then
            .TextMatrix(.Row, 2) = RsQ.Fields("TotRs")
            .TextMatrix(1, 3) = Val(.TextMatrix(1, 3)) + RsQ.Fields("TotRs")
            .TextMatrix(.Row, 3) = Val(.TextMatrix(.Row - 1, 2)) + RsQ.Fields("TotRs")
        Else
            .TextMatrix(.Row, 3) = Val(.TextMatrix(.Row - 1, 2))
        End If
        .Row = 8
        Set RsQ = Nothing
        RsQ.Open "Select IIf(IsNull(Sum(CBal))=True,0,Sum(CBal)) As TotRs From TmpTrialBal Where HCode=50", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(.Row, 2) = RsQ.Fields("TotRs")
        .Row = 9
        Set RsQ = Nothing
        RsQ.Open "Select IIf(IsNull(Sum(CBal))=True,0,Sum(CBal)) As TotRs From TmpTrialBal Where AcCode In (" & mAcList & ") And HCode=51", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(.Row, 2) = RsQ.Fields("TotRs")
        .TextMatrix(.Row, 3) = Val(.TextMatrix(.Row - 1, 2)) + RsQ.Fields("TotRs")
        Inc = Inc - Val(.TextMatrix(.Row, 3))
        .Row = 11
        Set RsQ = Nothing
        RsQ.Open "Select IIf(IsNull(Sum(CBal))=True,0,Sum(CBal)) As TotRs From TmpTrialBal Where AcCode In (" & mAcList & ") And LCode In (73,74,83)", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(.Row, 3) = RsQ.Fields("TotRs")
        Inc = Inc - Val(.TextMatrix(.Row, 3))
        .Row = 12
        Set RsQ = Nothing
        RsQ.Open "Select IIf(IsNull(Sum(CBal))=True,0,Sum(CBal)) As TotRs From TmpTrialBal Where AcCode In (" & mAcList & ") And LCode=75", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(.Row, 3) = RsQ.Fields("TotRs")
        Inc = Inc - Val(.TextMatrix(.Row, 3))
        .Row = 14
        Set RsQ = Nothing
        RsQ.Open "Select IIf(IsNull(Sum(CBal))=True,0,Sum(CBal)) As TotRs From TmpTrialBal Where AcCode In (" & mAcList & ") And LCode In (72,106)", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(.Row, 2) = RsQ.Fields("TotRs")
        .Row = 15
        Set RsQ = Nothing
        RsQ.Open "Select IIf(IsNull(Sum(CBal))=True,0,Sum(CBal)) As TotRs From TmpTrialBal Where AcCode In (" & mAcList & ") And LCode=76", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(.Row, 2) = RsQ.Fields("TotRs")
        .TextMatrix(.Row, 3) = Val(.TextMatrix(.Row - 1, 2)) + RsQ.Fields("TotRs")
        Inc = Inc - Val(.TextMatrix(.Row, 3))
        .Row = 16
        Set RsQ = Nothing
        RsQ.Open "Select IIf(IsNull(Sum(DBal))=True,0,Sum(DBal)) As TotRs From TmpTrialBal Where AcCode In (" & mAcList & ") And HCode=37", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(.Row, 3) = RsQ.Fields("TotRs")
        If Val(.TextMatrix(.Row, 3)) > Inc Then .TextMatrix(.Row, 3) = Val(Inc) Else .TextMatrix(.Row, 3) = Val(.TextMatrix(.Row, 3))
        Inc = Inc - Val(.TextMatrix(.Row, 3))
        .Row = 17
        Set RsQ = Nothing
        RsQ.Open "Select IIf(IsNull(Sum(DBal))=True,0,Sum(DBal)) As TotRs From TmpTrialBal Where AcCode In (" & mAcList & ") And HCode=38", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(.Row, 3) = RsQ.Fields("TotRs")
        If Val(.TextMatrix(.Row, 3)) > Inc Then .TextMatrix(.Row, 3) = Val(Inc) Else .TextMatrix(.Row, 3) = Val(.TextMatrix(.Row, 3))
        .Row = 19
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As TotRs From TmpCtDtl Where AcCode In (" & mAcList & ") And ECode=218", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(.Row, 2) = RsQ.Fields("TotRs")
        .Row = 20
        Set RsQ = Nothing   ' And RpDtl.Side='C'
        RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As TotRs From TmpCtDtl Where AcCode In (" & mAcList & ") And ECode=223", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(.Row, 2) = RsQ.Fields("TotRs")
        .Row = 21
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As TotRs From TmpCtDtl Where ECode<>223 And LCode=46", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(.Row, 2) = RsQ.Fields("TotRs")
        .Row = 22
        Set RsQ = Nothing
        RsQ.Open "Select IIf(IsNull(Sum(CBal))=True,0,Sum(CBal)) As TotRs From TmpTrialBal Where LCode=84", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(.Row, 2) = RsQ.Fields("TotRs")
        If Val(.TextMatrix(19, 2)) + Val(.TextMatrix(20, 2)) + Val(.TextMatrix(21, 2)) <> 0 And Val(.TextMatrix(22, 2)) <> 0 Then
            If Val(.TextMatrix(19, 2)) + Val(.TextMatrix(20, 2)) + Val(.TextMatrix(21, 2)) < Val(.TextMatrix(22, 2)) Then
                .TextMatrix(.Row - 1, 3) = Val(.TextMatrix(19, 2)) + Val(.TextMatrix(20, 2)) + Val(.TextMatrix(21, 2))
            Else
                .TextMatrix(.Row - 1, 3) = Val(.TextMatrix(22, 2))
            End If
        Else
            .Cell(flexcpText, 19, 2, 22, 2) = "0.00"
        End If
        .Row = 24
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As TotRs From TmpCtDtl Where ECode<>218 And LCode=43", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        If RsQ.EOF = False Then .TextMatrix(.Row, 2) = RsQ.Fields("TotRs")
        .Row = 25
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(Amt))=True,0,Sum(Amt)) As TotRs From TmpCtDtl Where ECode In (254,260,333)", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(.Row, 2) = RsQ.Fields("TotRs")
        .Row = 26
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(DBal))=True,0,Sum(DBal)) As TotRs From TmpTrialBal Where HCode=19", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(.Row, 2) = RsQ.Fields("TotRs")
        .Row = 27
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(CBal))=True,0,Sum(CBal)) As TotRs From TmpTrialBal Where HCode=42", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(.Row, 2) = Round(RsQ.Fields("TotRs") * (8.33 / 100), 2)
        .Row = 28
        .TextMatrix(.Row, 2) = Round(RsQ.Fields("TotRs") * (4 / 100), 2)
        .Row = 29
        .TextMatrix(.Row, 2) = RsQ.Fields("TotRs")
        If Val(.TextMatrix(24, 2)) + Val(.TextMatrix(25, 2)) + Val(.TextMatrix(26, 2)) + Val(.TextMatrix(27, 2)) + Val(.TextMatrix(28, 2)) <> 0 And Val(.TextMatrix(29, 2)) <> 0 Then
            If Val(.TextMatrix(24, 2)) + Val(.TextMatrix(25, 2)) + Val(.TextMatrix(26, 2)) + Val(.TextMatrix(27, 2)) + Val(.TextMatrix(28, 2)) < Val(.TextMatrix(29, 2)) Then
                .TextMatrix(.Row - 1, 3) = Val(.TextMatrix(24, 2)) + Val(.TextMatrix(25, 2)) + Val(.TextMatrix(26, 2)) + Val(.TextMatrix(27, 2)) + Val(.TextMatrix(28, 2))
            Else
                .TextMatrix(.Row - 1, 3) = Val(.TextMatrix(29, 2))
            End If
        Else
            .Cell(flexcpText, 24, 2, 29, 2) = "0.00"
        End If
        .Row = 30
        Set RsQ = Nothing
        RsQ.Open "Select IIF(IsNull(Sum(CBal))=True,0,Sum(CBal)) As TotRs From TmpTrialBal Where HCode In (44,48)", DBWorkTmp, adOpenDynamic, adLockReadOnly, adCmdText
        .TextMatrix(.Row, 3) = Round(RsQ.Fields("TotRs") * (1 / 100), 2)
        TxtCTotal.Text = .TextMatrix(1, 3)
        For i = 2 To .Rows - 1
            TxtCTotal.Text = Val(TxtCTotal.Text) - Val(.TextMatrix(i, 3))
            If Val(.TextMatrix(i, 2)) = 0 Then .TextMatrix(i, 2) = ""
            If Val(.TextMatrix(i, 3)) = 0 Then .TextMatrix(i, 3) = ""
        Next
        If Val(TxtCTotal.Text) < 0 Then TxtCTotal.Text = "0.00"
        .Refresh
    End If
    TxtCTotal.Text = Format(TxtCTotal.Text, "0.00")
End With
End Sub
Private Sub SaveData()
Dim i As Double
On Error GoTo XErr
DbDataDB.BeginTrans
DbDataDB.Execute "Delete From 9CDtl Where AcCode=" & mAcCode
With VsfHelp
    For i = 1 To .Rows - 1
        DbDataDB.Execute "Insert InTo 9CDtl (AcCode,SrN,Amt1,Amt2) Values (" & mAcCode & "," & i & "," & Val(.TextMatrix(i, 2)) & _
        "," & Val(.TextMatrix(i, 3)) & ")"
    Next
End With
DbDataDB.CommitTrans
MsgBox "Record Saved.", vbInformation, "Alert"
Exit Sub
XErr:
MsgBox Err.Description
DbDataDB.RollbackTrans
End Sub
